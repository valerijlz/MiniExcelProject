private void pipeExcelToWebView(Uri uri) {
        try {
            Workbook workbook;
            try (InputStream is = getContentResolver().openInputStream(uri)) {
                workbook = WorkbookFactory.create(is);
            }

            Sheet sheet = workbook.getSheetAt(0);
            JSONArray jsonTable = new JSONArray();

            int totalRows = 40;
            try {
                totalRows = sheet.getPhysicalNumberOfRows() > 0 ? sheet.getLastRowNum() : 0;
            } catch (Throwable ignored) {}
            
            int maxCellCount = 0;
            for (int r = 0; r <= totalRows; r++) {
                try {
                    Row row = sheet.getRow(r);
                    if (row != null && row.getLastCellNum() > maxCellCount) {
                        maxCellCount = row.getLastCellNum();
                    }
                } catch (Throwable ignored) {}
            }
            if (maxCellCount < 15) maxCellCount = 15;
            if (totalRows == 0) totalRows = 40;

            // ПОЛУЧАЕМ СТАНДАРТНУЮ ШИРИНУ КОЛОНКИ ЛИСТА (если явная ширина не задана в Excel)
            // В POI getDefaultColumnWidth() возвращает ширину в символах (например, 8). 
            // Переводим её в пиксели (стандарт для Excel: 1 символ ≈ 8-9 пикселей при 96 DPI)
            int defaultColWidthInPx = (int) (sheet.getDefaultColumnWidth() * 8.4 + 5);
            if (defaultColWidthInPx < 40) defaultColWidthInPx = 85;

            // ТОЧНЫЙ АЛГОРИТМ КОНВЕРТАЦИИ ЕДИНИЦ EXCEL В ПИКСЕЛИ
            JSONArray jsonColWidths = new JSONArray();
            for (int c = 0; c < maxCellCount; c++) {
                // sheet.getColumnWidth() возвращает значение в 1/256 ширины символа
                int poiWidth = sheet.getColumnWidth(c);
                
                int widthInPx;
                // Исключаем дефолтное неинициализированное POI-значение (2048)
                if (poiWidth == 2048 && sheet.isColumnHidden(c)) {
                    widthInPx = defaultColWidthInPx;
                } else {
                    double characters = (double) poiWidth / 256.0;
                    // Официальная формула Microsoft для перевода ширины в пиксели при 96 DPI:
                    // Pixels = trunc(((256 * characters + trunc(128 / 7)) / 256) * 7)
                    // Упрощенный стабильный аналог для Android WebView:
                    widthInPx = (int) (characters * 8.4 + 5);
                }
                
                // Предотвращаем схлопывание колонок
                if (widthInPx < 30) widthInPx = defaultColWidthInPx;
                jsonColWidths.put(widthInPx);
            }

            DataFormatter formatter = new DataFormatter();
            JSONArray jsonRowHeights = new JSONArray();

            for (int r = 0; r <= totalRows; r++) {
                Row row = null;
                try { row = sheet.getRow(r); } catch (Throwable ignored) {}
                
                // Стандартная высота строки Excel (15pt ≈ 20px)
                int heightInPx = 20; 
                if (row != null && row.getHeightInPoints() > 0) {
                    // Перевод Points в Pixels (1 pt = 1.33 px при стандартной плотности)
                    heightInPx = (int) (row.getHeightInPoints() * 1.33); 
                }
                jsonRowHeights.put(heightInPx);

                JSONArray jsonRow = new JSONArray();
                for (int c = 0; c < maxCellCount; c++) {
                    JSONObject cellObj = new JSONObject();
                    cellObj.put("v", "");
                    
                    if (row != null) {
                        Cell cell = null;
                        try { cell = row.getCell(c); } catch (Throwable ignored) {}
                        if (cell != null) {
                            try {
                                cellObj.put("v", formatter.formatCellValue(cell));
                            } catch (Exception e) {
                                cellObj.put("v", "");
                            }

                            try {
                                CellStyle style = cell.getCellStyle();
                                if (style != null) {
                                    try {
                                        Color bgColor = style.getFillForegroundColorColor();
                                        if (bgColor != null && style.getFillPattern() != FillPatternType.NO_FILL) {
                                            String hexBg = getHexColor(bgColor);
                                            if (hexBg != null && !hexBg.equals("#000000")) {
                                                cellObj.put("bg", hexBg);
                                            }
                                        }
                                    } catch (Throwable ignored) {}

                                    try {
                                        int fontIdx = style.getFontIndex();
                                        Font font = workbook.getFontAt(fontIdx);
                                        if (font != null) {
                                            if (font.getBold()) cellObj.put("bold", true);
                                            if (font.getItalic()) cellObj.put("italic", true);
                                            
                                            try {
                                                if (font instanceof org.apache.poi.xssf.usermodel.XSSFFont) {
                                                    String fontColor = getHexColor(((org.apache.poi.xssf.usermodel.XSSFFont) font).getXSSFColor());
                                                    if (fontColor != null) cellObj.put("color", fontColor);
                                                } else if (font instanceof org.apache.poi.hssf.usermodel.HSSFFont) {
                                                    short colorIdx = font.getColor();
                                                    HSSFColor hssfColor = HSSFColor.getIndexHash().get((int) colorIdx);
                                                    String fontColor = getHexColor(hssfColor);
                                                    if (fontColor != null) cellObj.put("color", fontColor);
                                                }
                                            } catch (Throwable ignored) {}
                                        }
                                    } catch (Throwable ignored) {}
                                }
                            } catch (Throwable ignored) {}
                        }
                    }
                    jsonRow.put(cellObj);
                }
                jsonTable.put(jsonRow);
            }

            JSONArray jsonMerges = new JSONArray();
            try {
                int numRegions = sheet.getNumMergedRegions();
                for (int i = 0; i < numRegions; i++) {
                    try {
                        CellRangeAddress region = sheet.getMergedRegion(i);
                        if (region != null) {
                            JSONObject mergeObj = new JSONObject();
                            mergeObj.put("sr", region.getFirstRow());
                            mergeObj.put("er", region.getLastRow());
                            mergeObj.put("sc", region.getFirstColumn());
                            mergeObj.put("ec", region.getLastColumn());
                            jsonMerges.put(mergeObj);
                        }
                    } catch (Throwable ignored) {}
                }
            } catch (Throwable ignored) {}

            workbook.close();

            JSONObject payload = new JSONObject();
            payload.put("matrix", jsonTable);
            payload.put("merges", jsonMerges);
            payload.put("widths", jsonColWidths);
            payload.put("heights", jsonRowHeights);

            String jsonString = payload.toString();
            String base64Payload = Base64.encodeToString(jsonString.getBytes(StandardCharsets.UTF_8), Base64.NO_WRAP);

            if (isEngineLoaded) {
                tableWebView.post(() -> tableWebView.evaluateJavascript("loadExcelFromBytes('" + base64Payload + "');", null));
            }
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка чтения: " + e.getLocalizedMessage(), Toast.LENGTH_LONG).show();
        }
    }
