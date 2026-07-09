package com.example.miniexcel;

import android.app.Activity;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.os.ParcelFileDescriptor;
import android.util.Base64;
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

        tableWebView.clearCache(true);
        tableWebView.clearHistory();
        tableWebView.clearFormData();

        // УБРАНЫ ОШИБОЧНЫЕ ЛОГИ, КОТОРЫЕ КРАШИЛИ КОМПИЛЯТОР

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

private void pipeExcelToWebView(org.apache.poi.ss.usermodel.Sheet sheet) {
    JSONArray jsonTable = new JSONArray();
    JSONArray jsonWidths = new JSONArray();
    JSONArray jsonHeights = new JSONArray();
    JSONArray jsonMerges = new JSONArray();

    int lastRowIdx = sheet.getLastRowNum();
    int maxColsCount = 0;

    // 1. Считаем максимальное количество колонок
    for (int r = 0; r <= lastRowIdx; r++) {
        org.apache.poi.ss.usermodel.Row row = sheet.getRow(r);
        if (row != null && row.getLastCellNum() > maxColsCount) {
            maxColsCount = row.getLastCellNum();
        }
    }
    if (maxColsCount == 0) maxColsCount = 12;

    try {
        // 2. Сбор ширин колонок
        for (int c = 0; c < maxColsCount; c++) {
            int w = sheet.getColumnWidth(c) / 35;
            jsonWidths.put(w > 0 ? w : 64);
        }

        // 3. Сбор строк, высот и ячеек
        for (int r = 0; r <= lastRowIdx; r++) {
            org.apache.poi.ss.usermodel.Row row = sheet.getRow(r);
            int h = (row != null) ? (int)(row.getHeightInPoints() * 1.33) : 20;
            jsonHeights.put(h > 0 ? h : 20);

            JSONArray rowArray = new JSONArray();
            for (int c = 0; c < maxColsCount; c++) {
                JSONObject cellObj = new JSONObject();
                cellObj.put("v", ""); // Дефолтное значение

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

        // 4. Сбор объединенных регионов и прокидывание их рамок
        int numRegions = sheet.getNumMergedRegions();
        for (int i = 0; i < numRegions; i++) {
            org.apache.poi.ss.util.CellRangeAddress region = sheet.getMergedRegion(i);
            if (region != null) {
                if (region.getFirstRow() > lastRowIdx || region.getFirstColumn() >= maxColsCount) continue;

                int finalEndRow = Math.min(region.getLastRow(), lastRowIdx);
                int finalEndCol = Math.min(region.getLastColumn(), maxColsCount - 1);

                JSONObject mergeObj = new JSONObject();
                mergeObj.put("sr", region.getFirstRow());
                mergeObj.put("er", finalEndRow);
                mergeObj.put("sc", region.getFirstColumn());
                mergeObj.put("ec", finalEndCol);
                jsonMerges.put(mergeObj);

                // Копируем рамки от главной ячейки по периметру объединения
                org.apache.poi.ss.usermodel.Row mainRow = sheet.getRow(region.getFirstRow());
                org.apache.poi.ss.usermodel.Cell mainCell = (mainRow != null) ? mainRow.getCell(region.getFirstColumn()) : null;
                
                if (mainCell != null && mainCell.getCellStyle() != null) {
                    org.apache.poi.ss.usermodel.CellStyle mainStyle = mainCell.getCellStyle();

                    for (int r = region.getFirstRow(); r <= finalEndRow; r++) {
                        JSONArray currentRow = jsonTable.getJSONArray(r);
                        for (int c = region.getFirstColumn(); c <= finalEndCol; c++) {
                            JSONObject cellObj = currentRow.getJSONObject(c);

                            if (c == region.getFirstColumn() && mainStyle.getBorderLeft() != org.apache.poi.ss.usermodel.BorderStyle.NONE) {
                                cellObj.put("bl", mainStyle.getBorderLeft().name());
                            }
                            if (c == finalEndCol && mainStyle.getBorderRight() != org.apache.poi.ss.usermodel.BorderStyle.NONE) {
                                cellObj.put("br", mainStyle.getBorderRight().name());
                            }
                            if (r == region.getFirstRow() && mainStyle.getBorderTop() != org.apache.poi.ss.usermodel.BorderStyle.NONE) {
                                cellObj.put("bt", mainStyle.getBorderTop().name());
                            }
                            if (r == finalEndRow && mainStyle.getBorderBottom() != org.apache.poi.ss.usermodel.BorderStyle.NONE) {
                                cellObj.put("bb", mainStyle.getBorderBottom().name());
                            }
                        }
                    }
                }
            }
        }

        // 5. Финальная сборка payload
        JSONObject payload = new JSONObject();
        payload.put("matrix", jsonTable);
        payload.put("widths", jsonWidths);
        payload.put("heights", jsonHeights);
        payload.put("merges", jsonMerges);

        cachedJsonPayload = payload.toString();
        android.util.Log.d("MiniExcelDebug", "JSON упакован успешно. Объединений: " + jsonMerges.length());

        // 6. Отправка в WebView в UI-потоке
        if (tableWebView != null) {
            tableWebView.post(new Runnable() {
                @Override
                public void run() {
                    tableWebView.evaluateJavascript("requestDataFromAndroid();", null);
                }
            });
        }

    } catch (Exception e) {
        android.util.Log.e("MiniExcelDebug", "Критическая ошибка парсинга: " + e.getMessage());
    }
}
