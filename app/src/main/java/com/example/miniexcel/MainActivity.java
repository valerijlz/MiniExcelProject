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
    private volatile String cachedJsonPayload = "";

    private ActivityResultLauncher<Intent> openFileLauncher;
    private ActivityResultLauncher<Intent> saveFileLauncher;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        openButton = findViewById(R.id.openButton);
        saveButton = findViewById(R.id.saveButton);
        tableWebView = findViewById(R.id.tableWebView);

        tableWebView.clearCache(true);
        tableWebView.clearHistory();
        tableWebView.clearFormData();

        // 1. Сначала настраиваем параметры
        WebSettings webSettings = tableWebView.getSettings();
        webSettings.setJavaScriptEnabled(true);
        // Запрет кэширования внутри WebView
        webSettings.setCacheMode(WebSettings.LOAD_NO_CACHE);
        tableWebView.clearCache(true);
        // -----------------------------------------------
        // --- ДОБАВЬТЕ ЭТИ СТРОКИ ДЛЯ БЛОКИРОВКИ СИСТЕМНОГО РАЗДУВАНИЯ ---
        webSettings.setTextZoom(100); // Жестко фиксирует масштаб текста и элементов на 100%
        webSettings.setUseWideViewPort(true);
        webSettings.setLoadWithOverviewMode(true);
        // ---------------------------------------------------------------
        webSettings.setDomStorageEnabled(true);
        webSettings.setAllowFileAccess(true);
        webSettings.setCacheMode(WebSettings.LOAD_NO_CACHE);
        
        webSettings.setSupportZoom(false);
        webSettings.setBuiltInZoomControls(false);
        webSettings.setUseWideViewPort(true);
        webSettings.setLoadWithOverviewMode(true);
        webSettings.setTextZoom(100);

        // 2. СТРОГО ДО ЗАГРУЗКИ URL регистрируем мост данных и клиент
        tableWebView.addJavascriptInterface(new AndroidBridge(), "AndroidBridge");
        tableWebView.setWebViewClient(new WebViewClient());

        // 3. Только теперь загружаем сам интерфейс
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
            
            // Защита 1: Жестко ограничиваем сканирование строк реальными границами контента
            int lastRowIdx = sheet.getLastRowNum();
            if (lastRowIdx > 1000) lastRowIdx = 1000; // Страховка от пустых отформатированных строк Excel
            if (lastRowIdx < 0) lastRowIdx = 0;

            // Определяем реальное максимальное количество заполненных колонок
            int maxColsCount = 0;
            for (int r = 0; r <= lastRowIdx; r++) {
                Row row = sheet.getRow(r);
                if (row != null) {
                    int lastCellNum = row.getLastCellNum();
                    if (lastCellNum > maxColsCount) {
                        maxColsCount = lastCellNum;
                    }
                }
            }
            if (maxColsCount > 100) maxColsCount = 100; // Ограничение по ширине
            if (maxColsCount < 1) maxColsCount = 1;

            // Сборка точных ширин колонок (уменьшаем коэффициент сжатия до компактного 2.4)
            JSONArray jsonColWidths = new JSONArray();
            for (int c = 0; c < maxColsCount; c++) {
                int widthInPx = 45; 
                try {
                    int poiWidth = sheet.getColumnWidth(c);
                    if (poiWidth > 0 && poiWidth != 2048) {
                        double characters = (double) poiWidth / 256.0;
                        // Коэффициент 2.0 дает идеальное пиксельное соответствие оригиналу
                        widthInPx = (int) (characters * 2.0);
                    }
                } catch (Throwable ignored) {}
                if (widthInPx < 20) widthInPx = 45;
                jsonColWidths.put(widthInPx);
            }

            DataFormatter formatter = new DataFormatter();
            JSONArray jsonTable = new JSONArray();
            JSONArray jsonRowHeights = new JSONArray();

            // Основной цикл безопасного парсинга матрицы ячеек
            for (int r = 0; r <= lastRowIdx; r++) {
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

                            // Безопасный разбор стилей ячейки
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
                                        
                                        if (font instanceof org.apache.poi.xssf.usermodel.XSSFFont) {
                                            String fontColor = getHexColor(((org.apache.poi.xssf.usermodel.XSSFFont) font).getXSSFColor());
                                            if (fontColor != null) cellObj.put("color", fontColor);
                                        } else if (font instanceof org.apache.poi.hssf.usermodel.HSSFFont) {
                                            short colorIdx = font.getColor();
                                            HSSFColor hssfColor = HSSFColor.getIndexHash().get((int) colorIdx);
                                            String fontColor = getHexColor(hssfColor);
                                            if (fontColor != null) cellObj.put("color", fontColor);
                                        }
                                    }
                                }
                            } catch (Throwable ignored) {}
                        }
                    }
                    jsonRow.put(cellObj);
                }
                jsonTable.put(jsonRow);
            }

            // Безопасный сбор объединенных ячеек (Merges)
            JSONArray jsonMerges = new JSONArray();
            try {
                int numRegions = sheet.getNumMergedRegions();
                for (int i = 0; i < numRegions; i++) {
                    CellRangeAddress region = sheet.getMergedRegion(i);
                    if (region != null) {
                        JSONObject mergeObj = new JSONObject();
                        mergeObj.put("sr", region.getFirstRow());
                        mergeObj.put("er", region.getLastRow());
                        mergeObj.put("sc", region.getFirstColumn());
                        mergeObj.put("ec", region.getLastColumn());
                        jsonMerges.put(mergeObj);
                    }
                }
            } catch (Throwable ignored) {}

            workbook.close();

            // Формируем чистый результирующий JSON
            JSONObject payload = new JSONObject();
            payload.put("matrix", jsonTable);
            payload.put("merges", jsonMerges);
            payload.put("widths", jsonColWidths);
            payload.put("heights", jsonRowHeights);

            cachedJsonPayload = payload.toString();

            // Форсируем мгновенное обновление сетки в WebView
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
            // Логируем запрос от WebView в поток приложения
            if (cachedJsonPayload == null || cachedJsonPayload.trim().isEmpty()) {
                runOnUiThread(() -> Toast.makeText(MainActivity.this, "WebView запросил данные, но в кэше Java пусто!", Toast.LENGTH_SHORT).show());
                return "";
            }

            runOnUiThread(() -> Toast.makeText(MainActivity.this, "Java успешно передает JSON в WebView (" + cachedJsonPayload.length() + " символов)", Toast.LENGTH_SHORT).show());
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
