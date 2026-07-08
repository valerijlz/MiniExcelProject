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

        // 1. Инициализируем настройки ПРАВИЛЬНО и без дубликатов
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

        // 2. РЕГИСТРИРУЕМ мост данных и клиент ДО загрузки URL
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
                    int lastCellNum = row.getLastCellNum();
                    if (lastCellNum > maxColsCount) {
                        maxColsCount = lastCellNum;
                    }
                }
            }
            if (maxColsCount > 100) maxColsCount = 100;
            if (maxColsCount < 1) maxColsCount = 1;

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

            DataFormatter formatter = new DataFormatter();
            JSONArray jsonTable = new JSONArray();
            JSONArray jsonRowHeights = new JSONArray();

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

                                    try {
                                        if (style.getBorderTop() != BorderStyle.NONE) {
                                            cellObj.put("bt", style.getBorderTop().name());
                                            short borderStyleColor = style.getTopBorderColor();
                                            HSSFColor hc = HSSFColor.getIndexHash().get((int) borderStyleColor);
                                            String bc = getHexColor(hc);
                                            cellObj.put("btc", bc != null ? bc : "#000000");
                                        }
                                        if (style.getBorderBottom() != BorderStyle.NONE) {
                                            cellObj.put("bb", style.getBorderBottom().name());
                                            short borderStyleColor = style.getBottomBorderColor();
                                            HSSFColor hc = HSSFColor.getIndexHash().get((int) borderStyleColor);
                                            String bc = getHexColor(hc);
                                            cellObj.put("bbc", bc != null ? bc : "#000000");
                                        }
                                        if (style.getBorderLeft() != BorderStyle.NONE) {
                                            cellObj.put("bl", style.getBorderLeft().name());
                                            short borderStyleColor = style.getLeftBorderColor();
                                            HSSFColor hc = HSSFColor.getIndexHash().get((int) borderStyleColor);
                                            String bc = getHexColor(hc);
                                            cellObj.put("blc", bc != null ? bc : "#000000");
                                        }
                                        if (style.getBorderRight() != BorderStyle.NONE) {
                                            cellObj.put("br", style.getBorderRight().name());
                                            short borderStyleColor = style.getRightBorderColor();
                                            HSSFColor hc = HSSFColor.getIndexHash().get((int) borderStyleColor);
                                            String bc = getHexColor(hc);
                                            cellObj.put("brc", bc != null ? bc : "#000000");
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
                    CellRangeAddress region = sheet.getMergedRegion(i);
