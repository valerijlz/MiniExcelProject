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

        tableWebView.clearCache(true);
        tableWebView.clearHistory();
        tableWebView.clearFormData();

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
                android.util.Log.d("WebViewJS", consoleMessage.message() + " -- From line "
                        + consoleMessage.lineNumber() + " of "
                        + consoleMessage.sourceId());
                return true;
            }
        });

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

    private void pipeExcelToWebView(Uri fileUri) {
        new Thread(() -> {
            try (InputStream inputStream = getContentResolver().openInputStream(fileUri)) {
                android.util.Log.d("MiniExcelDebug", "Фон: Чтение книги Excel...");
                Workbook workbook = WorkbookFactory.create(inputStream);
                
                if (workbook.getNumberOfSheets() > 0) {
                    Sheet sheet = workbook.getSheetAt(0);
                    parseSheetToJson(sheet); 
                } else {
                    runOnUiThread(() -> Toast.makeText(MainActivity.this, "Файл пуст", Toast.LENGTH_SHORT).show());
                }
            } catch (Exception e) {
                android.util.Log.e("MiniExcelDebug", "Ошибка чтения файла по Uri: " + e.getMessage());
                runOnUiThread(() -> Toast.makeText(MainActivity.this, "Ошибка чтения файла", Toast.LENGTH_SHORT).show());
            }
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
            // Сбор ширин колонок
            for (int c = 0; c < maxColsCount; c++) {
                int w = sheet.getColumnWidth(c) / 35;
                jsonWidths.put(w > 0 ? w : 64);
            }

            // Сбор строк, высот и контента ячеек
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

            // Обработка объединенных регионов
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

// Сборка финального JSON-пакета данных
            JSONObject rootPayload = new JSONObject();
            rootPayload.put("matrix", jsonTable);
            rootPayload.put("widths", jsonWidths);
            rootPayload.put("heights", jsonHeights);
            rootPayload.put("merges", jsonMerges);

            // Сохраняем в кэш-переменную класса
            cachedJsonPayload = rootPayload.toString();

            // ИСПРАВЛЕНО: Безопасное обновление без блокировки UI-потока WebView
            tableWebView.post(() -> {
                // Передаем команду через setTimeout в JS, чтобы разгрузить стек вызовов Android
                tableWebView.evaluateJavascript("setTimeout(function() { requestDataFromAndroid(); }, 50);", null);
            });

        } catch (Exception e) {
            android.util.Log.e("MiniExcelDebug", "Ошибка парсинга в JSON: " + e.getMessage());
            runOnUiThread(() -> Toast.makeText(MainActivity.this, "Ошибка обработки структуры файла", Toast.LENGTH_SHORT).show());
        }
    }

    // Класс моста между Java и JavaScript движком Canvas
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
