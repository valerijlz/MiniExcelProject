package com.example.miniexcel

import android.net.Uri
import android.os.Bundle
import android.util.Log
import android.webkit.WebView
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.poi.ss.usermodel.*
import java.io.File
import java.io.FileOutputStream

class MainActivity : AppCompatActivity() {

    private lateinit var webView: WebView
    private var currentWorkbook: Workbook? = null

    private val openFileLauncher = registerForActivityResult(
        ActivityResultContracts.GetContent()
    ) { uri: Uri? ->
        uri?.let { loadExcelFile(it) }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

        // 1. Инициализация системных свойств StAX перед созданием UI и открытием файлов
        initStaxProviders()

        setContentView(R.layout.activity_main)

        webView = findViewById(R.id.webView)
        webView.settings.javaScriptEnabled = true

        // Если вы используете JS-интерфейс:
        // webView.addJavascriptInterface(WebAppInterface(this), "Android")

        webView.loadUrl("file:///android_asset/grid.html")

        openFileLauncher.launch("*/*")
    }

    private fun initStaxProviders() {
        try {
            // Принудительно задаем StAX-парсер Aalto, встроенный в poi-android
            System.setProperty("javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl")
            System.setProperty("javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl")
            System.setProperty("javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl")
        } catch (e: Exception) {
            Log.e("MiniExcel", "Failed to set StAX properties", e)
        }
    }

    private fun loadExcelFile(uri: Uri) {
        lifecycleScope.launch(Dispatchers.IO) {
            val originalClassLoader = Thread.currentThread().contextClassLoader
            try {
                // 2. ВАЖНО: Подменяем ClassLoader потока фоновой корутины на ClassLoader приложения,
                // чтобы ServiceLoader внутри POI нашел зашитые ресурсы парсера в APK
                Thread.currentThread().contextClassLoader = applicationContext.classLoader

                val localFile = File.createTempFile("excel_cache", ".tmp", cacheDir)
                contentResolver.openInputStream(uri)?.use { input ->
                    FileOutputStream(localFile).use { output ->
                        input.copyTo(output)
                    }
                }

                currentWorkbook = WorkbookFactory.create(localFile)
                val sheet = currentWorkbook?.getSheetAt(0) ?: throw Exception("Лист не найден")
                val jsonResult = parseSheetToJSON(sheet)

                withContext(Dispatchers.Main) {
                    webView.evaluateJavascript("window.loadJsonData($jsonResult)", null)
                    Toast.makeText(this@MainActivity, "Файл успешно загружен!", Toast.LENGTH_SHORT).show()
                }
            } catch (t: Throwable) {
                Log.e("MiniExcel", "Error opening Excel file", t)
                withContext(Dispatchers.Main) {
                    Toast.makeText(this@MainActivity, "Ошибка: ${t.localizedMessage}", Toast.LENGTH_LONG).show()
                }
            } finally {
                // Возвращаем фоновому потоку исходный ClassLoader
                Thread.currentThread().contextClassLoader = originalClassLoader
            }
        }
    }

    private fun parseSheetToJSON(sheet: Sheet): String {
        // Здесь находится ваша существующая логика конвертации таблицы в JSON для WebView/JS
        // Для примера минимальный возвращаемый формат:
        return "[]"
    }
}
