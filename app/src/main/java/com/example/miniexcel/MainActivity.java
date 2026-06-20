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

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
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
    private boolean isEngineLoaded = false;

    private ActivityResultLauncher<Intent> openFileLauncher;
    private ActivityResultLauncher<Intent> saveFileLauncher;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        openButton = findViewById(R.id.openButton);
        saveButton = findViewById(R.id.saveButton);
        tableWebView = findViewById(R.id.tableWebView);

        WebSettings webSettings = tableWebView.getSettings();
        webSettings.setJavaScriptEnabled(true);
        webSettings.setDomStorageEnabled(true);
        webSettings.setAllowFileAccess(true);

        tableWebView.addJavascriptInterface(new AndroidBridge(), "AndroidBridge");
        tableWebView.setWebViewClient(new WebViewClient() {
            @Override
            public void onPageFinished(WebView view, String url) {
                super.onPageFinished(view, url);
                isEngineLoaded = true;
            }
        });

        tableWebView.loadUrl("file:///android_asset/index.html");
        initFileLaunchers();

        openButton.setOnClickListener(v -> {
            Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
            intent.addCategory(Intent.CATEGORY_OPENABLE);
            // РАЗРЕШАЕМ ОБА ФОРМАТА: и .xlsx, и старый .xls
            intent.setType("*/*");
            String[] mimeTypes = {
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // .xlsx
                    "application/vnd.ms-excel" // .xls
            };
            intent.putExtra(Intent.EXTRA_MIME_TYPES, mimeTypes);
            openFileLauncher.launch(intent);
        });

        saveButton.setOnClickListener(v -> {
            if (!isEngineLoaded) return;
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

    private void pipeExcelToWebView(Uri uri) {
        try (InputStream is = getContentResolver().openInputStream(uri);
             Workbook workbook = WorkbookFactory.create(is)) {
            
            Sheet sheet = workbook.getSheetAt(0);
            
            // 1. Извлекаем текстовую матрицу
            JSONArray jsonTable = new JSONArray();
            int maxCellCount = 0;
            for (Row row : sheet) {
                if (row.getLastCellNum() > maxCellCount) maxCellCount = row.getLastCellNum();
            }
            if (maxCellCount < 15) maxCellCount = 15;

            DataFormatter formatter = new DataFormatter();

            for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                JSONArray jsonRow = new JSONArray();
                for (int c = 0; c < maxCellCount; c++) {
                    if (row == null) {
                        jsonRow.put("");
                    } else {
                        Cell cell = row.getCell(c);
                        jsonRow.put(cell == null ? "" : formatter.formatCellValue(cell));
                    }
                }
                jsonTable.put(jsonRow);
            }

            // 2. Извлекаем координаты объединенных областей, чтобы структура не разъезжалась
            JSONArray jsonMerges = new JSONArray();
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                CellRangeAddress region = sheet.getMergedRegion(i);
                JSONObject mergeObj = new JSONObject();
                mergeObj.put("sr", region.getFirstRow());
                mergeObj.put("er", region.getLastRow());
                mergeObj.put("sc", region.getFirstColumn());
                mergeObj.put("ec", region.getLastColumn());
                jsonMerges.put(mergeObj);
            }

            // Упаковываем всё в один общий объект
            JSONObject payload = new JSONObject();
            payload.put("matrix", jsonTable);
            payload.put("merges", jsonMerges);

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

                    // Перезапись файла в режиме чтения-записи
                    try (ParcelFileDescriptor pfd = getContentResolver().openFileDescriptor(currentFileUri, "rwt");
                         FileOutputStream fos = new FileOutputStream(pfd.getFileDescriptor())) {
                        workbook.write(fos);
                    }
                    workbook.close();

                    Toast.makeText(MainActivity.this, "Изменения сохранены! Оформление защищено.", Toast.LENGTH_SHORT).show();
                } catch (Exception e) {
                    e.printStackTrace();
                    Toast.makeText(MainActivity.this, "Ошибка записи: " + e.getMessage(), Toast.LENGTH_LONG).show();
                }
            });
        }

        @JavascriptInterface
        public void onStatus(String message) {
            runOnUiThread(() -> Toast.makeText(MainActivity.this, message, Toast.LENGTH_SHORT).show());
        }
    }
}
