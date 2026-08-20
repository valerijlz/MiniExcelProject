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
import javax.xml.stream.XMLEventFactory
import javax.xml.stream.XMLInputFactory
import javax.xml.stream.XMLOutputFactory

class MainActivity : AppCompatActivity() {

    private lateinit var webView: WebView
    private lateinit var btnOpen: Button
    private var currentWorkbook: Workbook? = null
    
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

        // 1. Прямая инициализация StAX фабрик через явное указание Aalto
        initStaxProvidersDirectly()

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

        webView.webChromeClient = object : WebChromeClient() {
            override fun onConsoleMessage(consoleMessage: ConsoleMessage): Boolean {
                Log.d("MiniExcel-JS", "${consoleMessage.message()} -- From line ${consoleMessage.lineNumber()} of ${consoleMessage.sourceId()}")
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

    private fun initStaxProvidersDirectly() {
        try {
            // Задаем свойства системы
            System.setProperty("javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl")
            System.setProperty("javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl")
            System.setProperty("javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl")

            // ПРИНУДИТЕЛЬНАЯ инициализация классов Aalto в памяти
            // Это решает ошибку Provider com.bea.xml.stream.EventFactory
            Class.forName("com.fasterxml.aalto.stax.InputFactoryImpl")
            Class.forName("com.fasterxml.aalto.stax.OutputFactoryImpl")
            Class.forName("com.fasterxml.aalto.stax.EventFactoryImpl")
        } catch (e: Exception) {
            Log.e("MiniExcel", "Failed to force-init StAX classes", e)
        }
    }

    private fun loadExcelFile(uri: Uri) {
        lifecycleScope.launch(Dispatchers.IO) {
            val originalClassLoader = Thread.currentThread().contextClassLoader
            try {
                // Подменяем ClassLoader для текущего фонового потока
                Thread.currentThread().contextClassLoader = applicationContext.classLoader

                val localFile = File.createTempFile("excel_cache", ".tmp", cacheDir)
                contentResolver.openInputStream(uri)?.use { input ->
                    FileOutputStream(localFile).use { output ->
                        input.copyTo(output)
                    }
                }

                currentWorkbook = WorkbookFactory.create(localFile)
                val sheet = currentWorkbook?.getSheetAt(0) ?: throw Exception("Лист не найден")
                
                // Преобразуем лист в JSON с реальными данными ячеек
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

    /**
     * Исправление для пустых ячеек (.xls / .xlsx)
     * Базовый парсер значений ячеек в JSON для передачи в grid.html
     */
    private fun parseSheetToJSON(sheet: Sheet): String {
        val rowsArray = StringBuilder("[")
        val formatter = DataFormatter()

        for (rowIndex in 0..sheet.lastRowNum) {
            val row = sheet.getRow(rowIndex) ?: continue
            if (rowIndex > 0 && rowsArray.length > 1) rowsArray.append(",")

            rowsArray.append("[")
            val maxCol = row.lastCellNum.toInt()

            for (colIndex in 0 until maxCol) {
                val cell = row.getCell(colIndex)
                val cellValue = if (cell != null) formatter.formatCellValue(cell) else ""

                // Экранируем кавычки и переносы для безопасного JSON
                val escapedValue = cellValue
                    .replace("\\", "\\\\")
                    .replace("\"", "\\\"")
                    .replace("\n", "\\n")

                if (colIndex > 0) rowsArray.append(",")
                rowsArray.append("\"$escapedValue\"")
            }
            rowsArray.append("]")
        }

        rowsArray.append("]")
        return rowsArray.toString()
    }
}
