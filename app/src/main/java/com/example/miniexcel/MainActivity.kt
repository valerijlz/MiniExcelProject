package com.example.miniexcel

import android.net.Uri
import android.os.Bundle
import android.util.Log
import android.webkit.ConsoleMessage
import android.webkit.ValueCallback
import android.webkit.WebChromeClient
import android.webkit.WebView
import android.widget.Button
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
    private lateinit var btnOpen: Button
    private var currentWorkbook: Workbook? = null
    
    // Переменная для обработки выбора файлов из самого WebView (если в grid.html есть <input type="file">)
    private var filePathCallback: ValueCallback<Array<Uri>>? = null

    // 1. Лаунчер для системного диалога открытия файлов (кнопка "Открыть")
    private val openFileLauncher = registerForActivityResult(
        ActivityResultContracts.GetContent()
    ) { uri: Uri? ->
        uri?.let { loadExcelFile(it) }
    }

    // 2. Лаунчер для обработки выбора файлов из клика внутри WebView
    private val webViewFilePickerLauncher = registerForActivityResult(
        ActivityResultContracts.GetContent()
    ) { uri: Uri? ->
        if (filePathCallback != null) {
            val results = if (uri != null) arrayOf(uri) else null
            filePathCallback?.onReceiveValue(results)
            filePathCallback = null
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

        // Инициализируем StAX-парсер для XLSX до инициализации интерфейса
        initStaxProviders()

        setContentView(R.layout.activity_main)

        // Связываем элементы интерфейса
        webView = findViewById(R.id.webView)
        btnOpen = findViewById(R.id.btnOpen) // Кнопка "Открыть" в разметке

        // Привязываем клик по кнопке "Открыть" к запуск диалога выбора файла
        btnOpen.setOnClickListener {
            openFileLauncher.launch("*/*")
        }

        // Настройка WebView для работы с JS и сеткой
        setupWebView()

        // Загружаем HTML-сетку
        webView.loadUrl("file:///android_asset/grid.html")
    }

    private fun setupWebView() {
        webView.settings.apply {
            javaScriptEnabled = true
            domStorageEnabled = true
            allowFileAccess = true
            allowContentAccess = true
        }

        webView.webChromeClient = object : WebChromeClient() {
            // Перехват логов JavaScript в Logcat для удобной отладки grid.html
            override fun onConsoleMessage(consoleMessage: ConsoleMessage): Boolean {
                Log.d("MiniExcel-JS", "${consoleMessage.message()} -- From line ${consoleMessage.lineNumber()} of ${consoleMessage.sourceId()}")
                return true
            }

            // Поддержка диалога выбора файла, если запуск идет из самого WebView
            override fun onShowFileChooser(
                webView: WebView?,
                filePathCallback: ValueCallback<Array<Uri>>?,
                fileChooserParams: FileChooserParams?
            ): Boolean {
                this@MainActivity.filePathCallback?.onReceiveValue(null)
                this@MainActivity.filePathCallback = filePathCallback
                webViewFilePickerLauncher.launch("*/*")
                return true
            }
        }
    }

    private fun initStaxProviders() {
        try {
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
                // Подменяем ClassLoader потока для корректного чтения XLSX через Aalto
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
                Thread.currentThread().contextClassLoader = originalClassLoader
            }
        }
    }

    private fun parseSheetToJSON(sheet: Sheet): String {
        // Ваши функции обхода ячеек, размеров, текста и стилей таблицы
        return "[]"
    }
}
