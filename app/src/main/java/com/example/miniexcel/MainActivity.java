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

            // Сверхточный расчет физической ширины колонок из Excel
            JSONArray jsonColWidths = new JSONArray();
            for (int c = 0; c < maxCellCount; c++) {
                int poiWidth = sheet.getColumnWidth(c);
                // Стандартный коэффициент перевода POI в экранные пиксели (символьный шаг Excel)
                int widthInPx = (int) (poiWidth * 0.035); 
                if (widthInPx < 40) widthInPx = 75; // Стандартная ячейка по умолчанию
                jsonColWidths.put(widthInPx);
            }

            DataFormatter formatter = new DataFormatter();
            JSONArray jsonRowHeights = new JSONArray();

            for (int r = 0; r <= totalRows; r++) {
                Row row = null;
                try { row = sheet.getRow(r); } catch (Throwable ignored) {}
                
                int heightInPx = 19; 
                if (row != null && row.getHeightInPoints() > 0) {
                    heightInPx = (int) (row.getHeightInPoints() * 1.3); 
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
