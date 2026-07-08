package com.example.miniexcel;

import android.app.Activity;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.os.ParcelFileDescriptor;
import android.util.Base64;
import android.webkit.JavascriptInterface;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Button;
import android.widget.Toast;
import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.appcompat.app.AppCompatActivity;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;

public class MainActivity extends AppCompatActivity {

    private WebView tableWebView;
    private Button openButton, saveButton;
    private Uri currentFileUri = null;
    private volatile String cachedJsonPayload = "{\"matrix\":[],\"merges\":[],\"widths\":[],\"heights\":[]}";

    private ActivityResultLauncher<Intent> openFileLauncher;
    private ActivityResultLauncher<Intent> saveFileLauncher;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        openButton = findViewById(R.id.openButton);
        saveButton = findViewById(R.id.saveButton);
        tableWebView = findViewById(R.id.tableWebView);

        // Принудительно чистим кэш самого WebView при старте
        tableWebView.clearCache(true);
        tableWebView.clearHistory();
        tableWebView.clearFormData();

        // 1. Инициализируем настройки
        WebSettings webSettings = tableWebView.getSettings();
        webSettings.setJavaScriptEnabled(true);
        
        // Жесткий запрет кэширования на уровне конфигурации
        webSettings.setCacheMode(WebSettings.LOAD_NO_CACHE);
        
        // Фиксация системного масштабирования элементов и текста на 100%
        webSettings.setTextZoom(100); 
        webSettings.setUseWideViewPort(true);
        webSettings.setLoadWithOverviewMode(true);
        webSettings.setSupportZoom(false);
        webSettings.setBuiltInZoomControls(false);
        
        webSettings.setDomStorageEnabled(true);
        webSettings.setAllowFileAccess(true);

        // 2. Регистрация моста данных и клиента ДО загрузки URL
        tableWebView.addJavascriptInterface(new AndroidBridge(), "AndroidBridge");
        tableWebView.setWebViewClient(new WebViewClient());

        // 3. Отладочный клиент для вывода ошибок JS в консоль Logcat (Тег: WebViewJS)
        tableWebView.setWebChromeClient(new android.webkit.WebChromeClient() {
            @Override
            public boolean onConsoleMessage(android.webkit.ConsoleMessage consoleMessage) {
                android.util.Log.d("WebViewJS", consoleMessage.message() + " -- From line "
                        + consoleMessage.lineNumber() + " of "
                        + consoleMessage.sourceId());
                return true;
            }
        });

        // 4. Загружаем сам веб-интерфейс
        tableWebView.loadUrl("file:///android_asset/grid.html");
        
        initFileLaunchers();

        openButton.setOnClickListener(v -> {
            Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
            intent.addCategory(Intent.CATEGORY_OPENABLE);
            intent.setType("*/*");
            String[] mimeTypes = {
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "application/vnd.ms-excel"
            };
            intent.putExtra(Intent.EXTRA_MIME_TYPES, mimeTypes);
            openFileLauncher.launch(intent);
        });

        saveButton.setOnClickListener(v -> {
            tableWebView.post(() -> tableWebView.evaluateJavascript("exportExcelToAndroid();", null));
        });
    }

    private void initFileLaunchers() {
        openFileLauncher = registerForActivityResult(
                new ActivityResultContracts.StartActivityForResult(),
                result -> {
                    if (result.getResultCode() == Activity.RESULT_OK && result.getData() != null) {
                        Uri uri = result.getData().getData();
                        if (uri != null) {
                            currentFileUri = uri;
                            pipeExcelToWebView(uri);
                        }
                    }
                }
        );

        saveFileLauncher = registerForActivityResult(
                new ActivityResultContracts.StartActivityForResult(),
                result -> {
                    if (result.getResultCode() == Activity.RESULT_OK && result.getData() != null) {
                        Uri uri = result.getData().getData();
                        if (uri != null) {
                            currentFileUri = uri;
                            tableWebView.post(() -> tableWebView.evaluateJavascript("exportExcelToAndroid();", null));
                        }
                    }
                }
        );
    }

    private String getHexColor(Color color) {
        if (color == null) return null;
        try {
            if (color instanceof XSSFColor) {
                byte[] rgb = ((XSSFColor) color).getARGB();
                if (rgb != null && rgb.length == 4) {
                    return String.format("#%02x%02x%02x", rgb[1], rgb[2], rgb[3]);
                }
            } else if (color instanceof HSSFColor) {
                short[] triplet = ((HSSFColor) color).getTriplet();
                if (triplet != null && triplet.length == 3) {
                    return String.format("#%02x%02x%02x", triplet[0], triplet[1], triplet[2]);
                }
            }
        } catch (Throwable ignored) {}
        return null;
    }

    private void pipeExcelToWebView(Uri uri) {
        try {
            Workbook workbook;
            try (InputStream is = getContentResolver().openInputStream(uri)) {
                workbook = WorkbookFactory.create(is);
            }

            Sheet sheet = workbook.getSheetAt(0);
            
            int lastRowIdx = sheet.getLastRowNum();
            if (lastRowIdx > 1000) lastRowIdx = 1000;
            if (lastRowIdx < 0) lastRowIdx = 0;

            int maxColsCount = 0;
            for (int r = 0; r <= lastRowIdx; r++) {
                Row row = sheet.getRow(r);
                if (row != null) {
                    // Используем getLastCellNum, но если строка пустая, она не раздует нам таблицу
                    int lastCellNum = row.getLastCellNum();
                    if (lastCellNum > maxColsCount) {
                        maxColsCount = lastCellNum;
                    }
                }
            }
            // ЗАЩИТА: Если POI считает, что колонок слишком много (например, из-за стилей), ограничиваем физическим присутствием контента
            if (maxColsCount > 40) { 
                maxColsCount = 40; // Не даем синей рамке уйти за пределы разумного экспортного листа
            }
            if (maxColsCount < 5) maxColsCount = 5;

            JSONArray jsonColWidths = new JSONArray();
            for (int c = 0; c < maxColsCount; c++) {
                int widthInPx = 64;
                try {
                    int poiWidth = sheet.getColumnWidth(c);
                    if (poiWidth > 0) {
                        double characters = (double) poiWidth / 256.0;
                        widthInPx = (int) Math.round(characters * 7.5);
                    }
                } catch (Throwable ignored) {}
                if (widthInPx < 15) widthInPx = 15;
                jsonColWidths.put(widthInPx);
            }

// --- НАЧАЛО ИСПРАВЛЕННОГО БЛОКА СБОРА ДАННЫХ ---
            DataFormatter formatter = new DataFormatter();
            JSONArray jsonTable = new JSONArray();
            JSONArray jsonRowHeights = new JSONArray();

            // Шаг 1: Сначала собираем базовые данные и тексты ячеек
            for (let r = 0; r <= lastRowIdx; r++) {
                Row row = sheet.getRow(r);
                
                int heightInPx = 18;
                if (row != null && row.getHeightInPoints() > 0) {
                    heightInPx = (int) (row.getHeightInPoints() * 1.33);
                }
                jsonRowHeights.put(heightInPx);

                JSONArray jsonRow = new JSONArray();
                for (int c = 0; c < maxColsCount; c++) {
                    JSONObject cellObj = new JSONObject();
                    cellObj.put("v", "");

                    if (row != null) {
                        Cell cell = row.getCell(c);
                        if (cell != null) {
                            try {
                                cellObj.put("v", formatter.formatCellValue(cell));
                            } catch (Throwable e) {
                                cellObj.put("v", "");
                            }

                            try {
                                CellStyle style = cell.getCellStyle();
                                if (style != null) {
                                    Color bgColor = style.getFillForegroundColorColor();
                                    if (bgColor != null && style.getFillPattern() != FillPatternType.NO_FILL) {
                                        String hexBg = getHexColor(bgColor);
                                        if (hexBg != null && !hexBg.equals("#000000")) {
                                            cellObj.put("bg", hexBg);
                                        }
                                    }

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

                                    // Базовый сбор индивидуальных границ ячейки
                                    if (style.getBorderTop() != BorderStyle.NONE) {
                                        cellObj.put("bt", style.getBorderTop().name());
                                        cellObj.put("btc", getHexColor(HSSFColor.getIndexHash().get((int) style.getTopBorderColor())));
                                    }
                                    if (style.getBorderBottom() != BorderStyle.NONE) {
                                        cellObj.put("bb", style.getBorderBottom().name());
                                        cellObj.put("bbc", getHexColor(HSSFColor.getIndexHash().get((int) style.getBottomBorderColor())));
                                    }
                                    if (style.getBorderLeft() != BorderStyle.NONE) {
                                        cellObj.put("bl", style.getBorderLeft().name());
                                        cellObj.put("blc", getHexColor(HSSFColor.getIndexHash().get((int) style.getLeftBorderColor())));
                                    }
                                    if (style.getBorderRight() != BorderStyle.NONE) {
                                        cellObj.put("br", style.getBorderRight().name());
                                        cellObj.put("brc", getHexColor(HSSFColor.getIndexHash().get((int) style.getRightBorderColor())));
                                    }
                                }
                            } catch (Throwable ignored) {}
                        }
                    }
                    jsonRow.put(cellObj);
                }
                jsonTable.put(jsonRow);
            }

            // Шаг 2: ИСПРАВЛЕНИЕ ВЕРТИКАЛЬНЫХ ГРАНИЦ В ОБЪЕДИНЕНИЯХ
            // Пробегаемся по всем регионам и насильно копируем внешние границы на все внутренние ячейки региона
            JSONArray jsonMerges = new JSONArray();
            try {
                int numRegions = sheet.getNumMergedRegions();
                for (int i = 0; i < numRegions; i++) {
                    CellRangeAddress region = sheet.getMergedRegion(i);
                    if (region != null) {
                        // Ограничиваем сборку рамками нашей рабочей матрицы
                        if (region.getFirstRow() > lastRowIdx || region.getFirstColumn() > maxColsCount) continue;

                        JSONObject mergeObj = new JSONObject();
                        mergeObj.put("sr", region.getFirstRow());
                        mergeObj.put("er", Math.min(region.getLastRow(), lastRowIdx));
                        mergeObj.put("sc", region.getFirstColumn());
                        mergeObj.put("ec", Math.min(region.getLastColumn(), maxColsCount - 1));
                        jsonMerges.put(mergeObj);

                        // Берем главную (топовую) ячейку, чтобы узнать её эталонные границы
                        Row mainRow = sheet.getRow(region.getFirstRow());
                        Cell mainCell = (mainRow != null) ? mainRow.getCell(region.getFirstColumn()) : null;
                        if (mainCell != null && mainCell.getCellStyle() != null) {
                            CellStyle mainStyle = mainCell.getCellStyle();
                            
                            // Проходим по всем ячейкам внутри этого объединения и проставляем им "внешний кант"
                            for (int r = region.getFirstRow(); r <= Math.min(region.getLastRow(), lastRowIdx); r++) {
                                JSONArray rowArray = jsonTable.getJSONArray(r);
                                for (int c = region.getFirstColumn(); c <= Math.min(region.getLastColumn(), maxColsCount - 1); c++) {
                                    JSONObject cellObj = rowArray.getJSONObject(c);

                                    // Если ячейка находится на левом краю объединения -> копируем левую границу главной ячейки
                                    if (c == region.getFirstColumn() && mainStyle.getBorderLeft() != BorderStyle.NONE) {
                                        cellObj.put("bl", mainStyle.getBorderLeft().name());
                                        cellObj.put("blc", getHexColor(HSSFColor.getIndexHash().get((int) mainStyle.getLeftBorderColor())));
                                    }
                                    // Если на правом краю -> копируем правую границу
                                    if (c == region.getLastColumn() && mainStyle.getBorderRight() != BorderStyle.NONE) {
                                        cellObj.put("br", mainStyle.getBorderRight().name());
                                        cellObj.put("brc", getHexColor(HSSFColor.getIndexHash().get((int) mainStyle.getRightBorderColor())));
                                    }
                                    // Если на верхнем краю
                                    if (r == region.getFirstRow() && mainStyle.getBorderTop() != BorderStyle.NONE) {
                                        cellObj.put("bt", mainStyle.getBorderTop().name());
                                        cellObj.put("btc", getHexColor(HSSFColor.getIndexHash().get((int) mainStyle.getTopBorderColor())));
                                    }
                                    // Если на нижнем краю
                                    if (r == region.getLastRow() && mainStyle.getBorderBottom() != BorderStyle.NONE) {
                                        cellObj.put("bb", mainStyle.getBorderBottom().name());
                                        cellObj.put("bbc", getHexColor(HSSFColor.getIndexHash().get((int) mainStyle.getBottomBorderColor())));
                                    }
                                }
                            }
                        }
                    }
                }
            } catch (Throwable ignored) {}

            // Шаг 3: ИСПРАВЛЕНИЕ НАЛОЖЕНИЯ СМЕЖНЫХ ГРАНИЦ (Убираем жирноту)
            // Если у текущей ячейки есть верхняя граница, а у верхней ячейки есть нижняя — убираем верхнюю у текущей, чтобы они не суммировались
            for (int r = 1; r < jsonTable.length(); r++) {
                JSONArray currentRow = jsonTable.getJSONArray(r);
                JSONArray prevRow = jsonTable.getJSONArray(r - 1);
                for (int c = 0; c < currentRow.length(); c++) {
                    JSONObject currentCell = currentRow.getJSONObject(c);
                    JSONObject prevCell = prevRow.getJSONObject(c);
                    if (currentCell.has("bt") && prevCell.has("bb")) {
                        currentCell.remove("bt"); // Движок Canvas отрисует только нижнюю границу предыдущей строки, предотвращая двоение
                    }
                }
            }
            // То же самое для вертикальных смежных границ (лево/право)
            for (int r = 0; r < jsonTable.length(); r++) {
                JSONArray rowArray = jsonTable.getJSONArray(r);
                for (int c = 1; c < rowArray.length(); c++) {
                    JSONObject currentCell = rowArray.getJSONObject(c);
                    JSONObject prevCell = rowArray.getJSONObject(c - 1);
                    if (currentCell.has("bl") && prevCell.has("br")) {
                        currentCell.remove("bl");
                    }
                }
            }
            // --- КОНЕЦ ИСПРАВЛЕННОГО БЛОКА ---

            workbook.close();

            JSONObject payload = new JSONObject();
            payload.put("matrix", jsonTable);
            payload.put("merges", jsonMerges);
            payload.put("widths", jsonColWidths);
            payload.put("heights", jsonRowHeights);

            cachedJsonPayload = payload.toString();

            tableWebView.post(() -> {
                tableWebView.evaluateJavascript("requestDataFromAndroid();", null);
            });

        } catch (Throwable e) {
            e.printStackTrace();
            Toast.makeText(this, "Критический сбой разбора файла: " + e.getLocalizedMessage(), Toast.LENGTH_LONG).show();
        }
    }

    public class AndroidBridge {
        @JavascriptInterface
        public String getExcelData() {
            if (cachedJsonPayload == null || cachedJsonPayload.trim().isEmpty() || cachedJsonPayload.equals("{\"matrix\":[],\"merges\":[],\"widths\":[],\"heights\":[]}")) {
                return "{\"matrix\":[],\"merges\":[],\"widths\":[],\"heights\":[]}";
            }

            runOnUiThread(() -> Toast.makeText(MainActivity.this, "Таблица успешно загружена!", Toast.LENGTH_SHORT).show());
            return cachedJsonPayload;
        }

        @JavascriptInterface
        public void saveFileData(String base64Data) {
            runOnUiThread(() -> {
                if (currentFileUri == null) {
                    Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
                    intent.addCategory(Intent.CATEGORY_OPENABLE);
                    intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    intent.putExtra(Intent.EXTRA_TITLE, "Table.xlsx");
                    saveFileLauncher.launch(intent);
                    return;
                }

                try {
                    byte[] decodedBytes = Base64.decode(base64Data, Base64.DEFAULT);
                    String jsonString = new String(decodedBytes, StandardCharsets.UTF_8);
                    JSONArray jsonTable = new JSONArray(jsonString);

                    Workbook workbook;
                    try (InputStream is = getContentResolver().openInputStream(currentFileUri)) {
                        workbook = WorkbookFactory.create(is);
                    }

                    Sheet sheet = workbook.getSheetAt(0);

                    for (int r = 0; r < jsonTable.length(); r++) {
                        JSONArray jsonRow = jsonTable.getJSONArray(r);
                        Row row = sheet.getRow(r);
                        if (row == null) row = sheet.createRow(r);

                        for (int c = 0; c < jsonRow.length(); c++) {
                            String value = jsonRow.getString(c);
                            Cell cell = row.getCell(c);
                            if (cell == null) cell = row.createCell(c);
                            
                            try {
                                double num = Double.parseDouble(value);
                                cell.setCellValue(num);
                            } catch (NumberFormatException e) {
                                cell.setCellValue(value);
                            }
                        }
                    }

                    ByteArrayOutputStream baos = new ByteArrayOutputStream();
                    workbook.write(baos);
                    byte[] workbookBytes = baos.toByteArray();
                    workbook.close();

                    try (ParcelFileDescriptor pfd = getContentResolver().openFileDescriptor(currentFileUri, "rwt");
                         FileOutputStream fos = new FileOutputStream(pfd.getFileDescriptor())) {
                        fos.write(workbookBytes);
                        fos.flush();
                    }

                    Toast.makeText(MainActivity.this, "Изменения успешно сохранены!", Toast.LENGTH_SHORT).show();
                } catch (Exception e) {
                    e.printStackTrace();
                    Toast.makeText(MainActivity.this, "Ошибка сохранения: " + e.getLocalizedMessage(), Toast.LENGTH_LONG).show();
                }
            });
        }

        @JavascriptInterface
        public void onStatus(String message) {
            runOnUiThread(() -> Toast.makeText(MainActivity.this, message, Toast.LENGTH_SHORT).show());
        }
    }
}
