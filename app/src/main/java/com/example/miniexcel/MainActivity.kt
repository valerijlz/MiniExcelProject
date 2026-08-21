package com.example.miniexcel

import android.net.Uri
import android.os.Bundle
import android.util.Log
import android.webkit.ConsoleMessage
import android.webkit.ValueCallback
import android.webkit.WebChromeClient
import android.webkit.WebView
import android.webkit.WebViewClient
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
    private var isWebViewLoaded = false

    private var filePathCallback: ValueCallback<Array<Uri>>? = null

    private val openFileLauncher = registerForActivityResult(
        ActivityResultContracts.GetContent()
    ) { uri: Uri? ->
        uri?.let { loadExcelFile(it) }
    }

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

        // 1. Принудительная инициализация фабрик StAX
        setupStaxProperties()

        setContentView(R.layout.activity_main)

        webView = findViewById(R.id.webView)
        btnOpen = findViewById(R.id.btnOpen)

        btnOpen.setOnClickListener {
            openFileLauncher.launch("*/*")
        }

        setupWebView()
        webView.loadUrl("file:///android_asset/grid.html")
    }

    private fun setupWebView() {
        webView.settings.apply {
            javaScriptEnabled = true
            domStorageEnabled = true
            allowFileAccess = true
            allowContentAccess = true
        }

        webView.webViewClient = object : WebViewClient() {
            override fun onPageFinished(view: WebView?, url: String?) {
                super.onPageFinished(view, url)
                isWebViewLoaded = true
                Log.d("MiniExcel", "WebView HTML grid successfully loaded")
            }
        }

        webView.webChromeClient = object : WebChromeClient() {
            override fun onConsoleMessage(consoleMessage: ConsoleMessage): Boolean {
                Log.d("MiniExcel-JS", "${consoleMessage.message()} -- line ${consoleMessage.lineNumber()}")
                return true
            }

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

    private fun setupStaxProperties() {
        try {
            // Указываем явные классы Aalto StAX, вшитые внутрь poi-android
            System.setProperty("javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl")
            System.setProperty("javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl")
            System.setProperty("javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl")
        } catch (e: Exception) {
            Log.e("MiniExcel", "Error setting StAX system properties", e)
        }
    }

    private fun loadExcelFile(uri: Uri) {
        lifecycleScope.launch(Dispatchers.IO) {
            val originalClassLoader = Thread.currentThread().contextClassLoader
            var tempFile: File? = null
            var workbook: Workbook? = null

            try {
                // Подменяем контекстный ClassLoader для корутины
                Thread.currentThread().contextClassLoader = applicationContext.classLoader

                tempFile = File.createTempFile("excel_cache", ".tmp", cacheDir)
                contentResolver.openInputStream(uri)?.use { input ->
                    FileOutputStream(tempFile).use { output ->
                        input.copyTo(output)
                    }
                }

                // Создаем Workbook
                workbook = WorkbookFactory.create(tempFile)
                val sheet = workbook.getSheetAt(0) ?: throw Exception("В файле нет листов")
                
                val jsonResult = parseSheetToJSON(sheet)
                Log.d("MiniExcel", "Parsed JSON length: ${jsonResult.length}")

                withContext(Dispatchers.Main) {
                    sendJsonToWebView(jsonResult)
                }

            } catch (t: Throwable) {
                Log.e("MiniExcel", "Error opening Excel file", t)
                withContext(Dispatchers.Main) {
                    Toast.makeText(this@MainActivity, "Ошибка: ${t.localizedMessage}", Toast.LENGTH_LONG).show()
                }
            } finally {
                try {
                    workbook?.close()
                } catch (e: Exception) {
                    // Игнорируем ошибки закрытия
                }
                tempFile?.delete()
                Thread.currentThread().contextClassLoader = originalClassLoader
            }
        }
    }

    private fun sendJsonToWebView(jsonResult: String) {
        if (!isWebViewLoaded) {
            Toast.makeText(this, "Сетка еще загружается, попробуйте снова через секунду", Toast.LENGTH_SHORT).show()
            return
        }

        // Вызываем функцию JS в WebView
        val script = "if (typeof loadJsonData === 'function') { loadJsonData($jsonResult); } else if (typeof window.loadJsonData === 'function') { window.loadJsonData($jsonResult); } else { console.error('loadJsonData function not found in JS'); }"
        
        webView.evaluateJavascript(script) { result ->
            Log.d("MiniExcel", "JS evaluation result: $result")
            Toast.makeText(this@MainActivity, "Файл прочитан!", Toast.LENGTH_SHORT).show()
        }
    }

    private fun parseSheetToJSON(sheet: Sheet): String {
        val rowsArray = StringBuilder("[")
        val formatter = DataFormatter()

        val lastRow = sheet.lastRowNum
        for (rowIndex in 0..lastRow) {
            val row = sheet.getRow(rowIndex)
            if (rowIndex > 0 && rowsArray.length > 1) rowsArray.append(",")

            rowsArray.append("[")
            if (row != null) {
                val maxCol = row.lastCellNum.toInt()
                for (colIndex in 0 until maxCol) {
                    val cell = row.getCell(colIndex)
                    val cellValue = if (cell != null) formatter.formatCellValue(cell) else ""

                    val escapedValue = cellValue
                        .replace("\\", "\\\\")
                        .replace("\"", "\\\"")
                        .replace("\n", "\\n")
                        .replace("\r", "")

                    if (colIndex > 0) rowsArray.append(",")
                    rowsArray.append("\"$escapedValue\"")
                }
            }
            rowsArray.append("]")
        }

        rowsArray.append("]")
        return rowsArray.toString()
    }
}
