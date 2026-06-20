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

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

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

        // Глубокая настройка встроенного движка WebView
        WebSettings webSettings = tableWebView.getSettings();
        webSettings.setJavaScriptEnabled(true);
        webSettings.setDomStorageEnabled(true);
        webSettings.setAllowFileAccess(true);
        webSettings.setBuiltInZoomControls(true);
        webSettings.setDisplayZoomControls(false);
        webSettings.setLoadWithOverviewMode(true);
        webSettings.setUseWideViewPort(true);

        // Регистрация защищенного моста связи Java <-> JavaScript
        tableWebView.addJavascriptInterface(new AndroidBridge(), "AndroidBridge");
        
        tableWebView.setWebViewClient(new WebViewClient() {
            @Override
            public void onPageFinished(WebView view, String url) {
                super.onPageFinished(view, url);
                isEngineLoaded = true;
            }
        });

        // Загружаем локальное ядро таблицы
        tableWebView.loadUrl("file:///android_asset/index.html");

        initFileLaunchers();

        openButton.setOnClickListener(v -> {
            Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
            intent.addCategory(Intent.CATEGORY_OPENABLE);
            intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            openFileLauncher.launch(intent);
        });

        saveButton.setOnClickListener(v -> {
            if (!isEngineLoaded) return;
            // Даем команду JS-движку собрать бинарный пакет данных
            tableWebView.post(() -> tableWebView.loadUrl("javascript:exportExcelToAndroid()"));
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
                            pipeExcelFileToWebView(uri);
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
                            // Повторно триггерим сохранение уже в выбранный пользователем новый файл
                            tableWebView.post(() -> tableWebView.loadUrl("javascript:exportExcelToAndroid()"));
                        }
                    }
                }
        );
    }

    // Чтение файла со смартфона и отправка его байтов внутрь JS-движка
    private void pipeExcelFileToWebView(Uri uri) {
        try (InputStream is = getContentResolver().openInputStream(uri);
             ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            if (is == null) return;
            
            byte[] buf = new byte[4096];
            int len;
            while ((len = is.read(buf)) != -1) {
                bos.write(buf, 0, len);
            }
            
            String base64Data = Base64.encodeToString(bos.toByteArray(), Base64.NO_WRAP);
            
            if (isEngineLoaded) {
                tableWebView.post(() -> tableWebView.loadUrl("javascript:loadExcelFromBytes('" + base64Data + "')"));
            }
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка чтения файла: " + e.getLocalizedMessage(), Toast.LENGTH_LONG).show();
        }
    }

    // Внутренний класс-мост, куда JavaScript передает финальный готовый файл Excel
    public class AndroidBridge {
        
        @JavascriptInterface
        public void saveFileData(String base64Data) {
            runOnUiThread(() -> {
                if (currentFileUri == null) {
                    // Если файл новый — открываем системный диалог "Сохранить как..."
                    Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
                    intent.addCategory(Intent.CATEGORY_OPENABLE);
                    intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    intent.putExtra(Intent.EXTRA_TITLE, "Table.xlsx");
                    saveFileLauncher.launch(intent);
                    return;
                }

                // Физическая безопасная запись байт в файл
                try (ParcelFileDescriptor pfd = getContentResolver().openFileDescriptor(currentFileUri, "rwt");
                     FileOutputStream fos = new FileOutputStream(pfd.getFileDescriptor())) {
                    
                    byte[] decodedBytes = Base64.decode(base64Data, Base64.DEFAULT);
                    fos.write(decodedBytes);
                    fos.flush();
                    
                    Toast.makeText(MainActivity.this, "Файл успешно сохранен со всеми стилями!", Toast.LENGTH_SHORT).show();
                } catch (Exception e) {
                    e.printStackTrace();
                    Toast.makeText(MainActivity.this, "Критическая ошибка записи: " + e.getMessage(), Toast.LENGTH_LONG).show();
                }
            });
        }

        @JavascriptInterface
        public void onStatus(String message) {
            runOnUiThread(() -> Toast.makeText(MainActivity.this, message, Toast.LENGTH_SHORT).show());
        }
    }
}
