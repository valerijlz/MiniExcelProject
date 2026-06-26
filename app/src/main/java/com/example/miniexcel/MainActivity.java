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
import java.nio.channels.FileChannel;
import java.nio.charset.StandardCharsets;

public class MainActivity extends AppCompatActivity {

    private WebView tableWebView;
    private Button openButton, saveButton;
    private Uri currentFileUri = null;

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
        webSettings.setDomStorageEnabled(true);
        webSettings.setAllowFileAccess(true);
        webSettings.setCacheMode(WebSettings.LOAD_NO_CACHE);
        
        // Отключаем встроенный зум-алгоритм WebView, мешавший fixed-разметке CSS
        webSettings.setSupportZoom(false);
        webSettings.setBuiltInZoomControls(false);
        
        webSettings.setUseWideViewPort(true);
        webSettings.setLoadWithOverviewMode(true);
        webSettings.setTextZoom(100);

        tableWebView.addJavascriptInterface(new AndroidBridge(), "AndroidBridge");
        tableWebView.setWebViewClient(new WebViewClient());

        tableWebView.loadUrl("file:///android_asset/index.html");
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

            int defaultColWidthInPx = 45; 

            JSONArray jsonColWidths = new JSONArray();
            for (int c = 0; c < maxCellCount; c++) {
                int widthInPx = defaultColWidthInPx;
                try {
                    int poiWidth = sheet.getColumnWidth(c);
                    if (poiWidth > 0 && poiWidth != 2048) {
                        double characters = (double) poiWidth / 256.0;
                        widthInPx = (int) (characters * 4.0);
                    }
                } catch (Throwable ignored) {}
                
                if (widthInPx < 20) widthInPx = defaultColWidthInPx;
                jsonColWidths.put(widthInPx);
            }

            DataFormatter formatter = new DataFormatter();
            JSONArray jsonRowHeights = new JSONArray();

            for (int r = 0; r <= totalRows; r++) {
                Row row = null;
                try { row = sheet.getRow(r); } catch (Throwable ignored) {}
                
                int heightInPx = 18; 
                if (row != null && row.getHeightInPoints() > 0) {
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

            tableWebView.post(() -> tableWebView.evaluateJavascript("loadExcelFromBytes('" + base64Payload + "');", null));
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка чтения: " + e.getLocalizedMessage(), Toast.LENGTH_LONG).show();
        }
    }

    public class AndroidBridge {
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
