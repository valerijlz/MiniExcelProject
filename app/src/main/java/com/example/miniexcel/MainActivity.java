package com.example.miniexcel;

import android.app.Activity;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.util.Log;
import android.webkit.JavascriptInterface;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Button;
import android.widget.Toast;
import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.appcompat.app.AppCompatActivity;

import org.apache.poi.ss.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.InputStream;

public class MainActivity extends AppCompatActivity {

    private WebView tableWebView;
    private Button openButton, saveButton;
    private Uri currentFileUri = null;
    
    // Держим дефолтный пустой JSON, чтобы WebView не падал при инициализации
    private final String EMPTY_PAYLOAD = "{\"matrix\":[],\"merges\":[],\"widths\":[],\"heights\":[]}";
    private volatile String cachedJsonPayload = EMPTY_PAYLOAD;

    private ActivityResultLauncher<Intent> openFileLauncher;
    private ActivityResultLauncher<Intent> saveFileLauncher;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        openButton = findViewById(R.id.openButton);
        saveButton = findViewById(R.id.saveButton);
        tableWebView = findViewById(R.id.tableWebView);

        setupWebView();
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
            if (tableWebView != null) {
                tableWebView.post(() -> tableWebView.evaluateJavascript("exportExcelToAndroid();", null));
            }
        });
    }

    private void setupWebView() {
        if (tableWebView == null) return;

        tableWebView.clearCache(true);
        tableWebView.clearHistory();

        WebSettings webSettings = tableWebView.getSettings();
        webSettings.setJavaScriptEnabled(true);
        webSettings.setCacheMode(WebSettings.LOAD_NO_CACHE);
        
        webSettings.setTextZoom(100); 
        webSettings.setUseWideViewPort(true);
        webSettings.setLoadWithOverviewMode(true);
        webSettings.setSupportZoom(false);
        webSettings.setBuiltInZoomControls(false);
        
        webSettings.setDomStorageEnabled(true);
        webSettings.setAllowFileAccess(true);

        tableWebView.addJavascriptInterface(new AndroidBridge(), "AndroidBridge");
        tableWebView.setWebViewClient(new WebViewClient());

        tableWebView.setWebChromeClient(new android.webkit.WebChromeClient() {
            @Override
            public boolean onConsoleMessage(android.webkit.ConsoleMessage consoleMessage) {
                android.util.Log.d("WebViewJS", consoleMessage.message() + " -- Line "
                        + consoleMessage.lineNumber() + " of " + consoleMessage.sourceId());
                return true;
            }
        });

        tableWebView.loadUrl("file:///android_asset/grid.html");
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
                            if (tableWebView != null) {
                                tableWebView.post(() -> tableWebView.evaluateJavascript("exportExcelToAndroid();", null));
                            }
                        }
                    }
                }
        );
    }

    private void pipeExcelToWebView(Uri fileUri) {
        // УТЕЧКА ПАМЯТИ И ФРИЗЫ ИСПРАВЛЕНЫ ТУТ: 
        // 1. Полностью сбрасываем кэш в Java-памяти до парсинга.
        cachedJsonPayload = EMPTY_PAYLOAD;
        
        // 2. Перезагружаем страницу grid.html. Это очищает контекст JS-движка V8, 
        // убирая из ОЗУ структуры старого открытого документа.
        tableWebView.post(() -> tableWebView.loadUrl("file:///android_asset/grid.html"));

        new Thread(() -> {
            // Использование try-with-resources гарантирует закрытие InputStream и Workbook (освобождает нативную память POI)
            try (InputStream inputStream = getContentResolver().openInputStream(fileUri);
                 Workbook workbook = WorkbookFactory.create(inputStream)) {
                
                android.util.Log.d("MiniExcelDebug", "Фон: Успешно открыта книга Excel.");
                if (workbook.getNumberOfSheets() > 0) {
                    Sheet sheet = workbook.getSheetAt(0);
                    parseSheetToJson(sheet); 
                } else {
                    runOnUiThread(() -> Toast.makeText(MainActivity.this, "Файл пуст", Toast.LENGTH_SHORT).show());
                }
            } catch (Exception e) {
                android.util.Log.e("MiniExcelDebug", "Ошибка чтения файла: " + e.getMessage());
                runOnUiThread(() -> Toast.makeText(MainActivity.this, "Ошибка чтения файла", Toast.LENGTH_SHORT).show());
            }
            
            // Намекаем системе, что пора собрать отработанные объекты строк/ячеек POI
            System.gc();
        }).start();
    }

    private void parseSheetToJson(org.apache.poi.ss.usermodel.Sheet sheet) {
        JSONArray jsonTable = new JSONArray();
        JSONArray jsonWidths = new JSONArray();
        JSONArray jsonHeights = new JSONArray();
        JSONArray jsonMerges = new JSONArray();

        int lastRowIdx = sheet.getLastRowNum();
        int maxColsCount = 0;

        for (int r = 0; r <= lastRowIdx; r++) {
            org.apache.poi.ss.usermodel.Row row = sheet.getRow(r);
            if (row != null && row.getLastCellNum() > maxColsCount) {
                maxColsCount = row.getLastCellNum();
            }
        }
        if (maxColsCount == 0) maxColsCount = 12;

        try {
            for (int c = 0; c < maxColsCount; c++) {
                int w = sheet.getColumnWidth(c) / 35;
                jsonWidths.put(w > 0 ? w : 64);
            }

            for (int r = 0; r <= lastRowIdx; r++) {
                org.apache.poi.ss.usermodel.Row row = sheet.getRow(r);
                int h = (row != null) ? (int)(row.getHeightInPoints() * 1.33) : 20;
                jsonHeights.put(h > 0 ? h : 20);

                JSONArray rowArray = new JSONArray();
                for (int c = 0; c < maxColsCount; c++) {
                    JSONObject cellObj = new JSONObject();
                    cellObj.put("v", ""); 

                    if (row != null) {
                        org.apache.poi.ss.usermodel.Cell cell = row.getCell(c);
                        if (cell != null) {
                            switch (cell.getCellType()) {
                                case STRING: 
                                    cellObj.put("v", cell.getStringCellValue()); 
                                    break;
                                case NUMERIC: 
                                    if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                                        cellObj.put("v", cell.getDateCellValue().toString());
                                    } else {
                                        cellObj.put("v", cell.getNumericCellValue());
                                    }
                                    break;
                                case FORMULA:
                                    try { cellObj.put("v", cell.getStringCellValue()); } 
                                    catch(Exception e1) {
                                        try { cellObj.put("v", cell.getNumericCellValue()); } 
                                        catch(Exception ignored) {}
                                    }
                                    break;
                                case BOOLEAN: 
                                    cellObj.put("v", cell.getBooleanCellValue()); 
                                    break;
                                default: 
                                    cellObj.put("v", "");
                            }
                        }
                    }
                    rowArray.put(cellObj);
                }
                jsonTable.put(rowArray);
            }

            int numRegions = sheet.getNumMergedRegions();
            for (int i = 0; i < numRegions; i++) {
                org.apache.poi.ss.util.CellRangeAddress region = sheet.getMergedRegion(i);
                if (region != null) {
                    if (region.getFirstRow() > lastRowIdx || region.getFirstColumn() >= maxColsCount) continue;

                    JSONObject mergeObj = new JSONObject();
                    mergeObj.put("sr", region.getFirstRow());
                    mergeObj.put("er", region.getLastRow());
                    mergeObj.put("sc", region.getFirstColumn());
                    mergeObj.put("ec", region.getLastColumn());
                    jsonMerges.put(mergeObj);
                }
            }

            JSONObject rootPayload = new JSONObject();
            rootPayload.put("matrix", jsonTable);
            rootPayload.put("widths", jsonWidths);
            rootPayload.put("heights", jsonHeights);
            rootPayload.put("merges", jsonMerges);

            // Кэшируем новую строку
            cachedJsonPayload = rootPayload.toString();

            // Безопасно пингуем JS-скрипт. Так как чуть выше мы сделали loadUrl,
            // страница гарантированно чистая и готова принять данные заново.
            tableWebView.post(() -> {
                tableWebView.evaluateJavascript("setTimeout(function() { requestDataFromAndroid(); }, 150);", null);
            });

        } catch (Exception e) {
            android.util.Log.e("MiniExcelDebug", "Ошибка парсинга в JSON: " + e.getMessage());
            runOnUiThread(() -> Toast.makeText(MainActivity.this, "Ошибка обработки структуры файла", Toast.LENGTH_SHORT).show());
        }
    }

    public class AndroidBridge {
        @JavascriptInterface
        public String getExcelData() {
            return cachedJsonPayload;
        }

        @JavascriptInterface
        public void onStatus(String message) {
            Log.d("MiniExcelStatus", message);
        }
    }
}
