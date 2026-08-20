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
    private var formulaEvaluator: FormulaEvaluator? = null

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

        // 1. Принудительно внедряем Aalto XML фабрики
        fixStaxProviders()

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

    private fun fixStaxProviders() {
        try {
            // Прописываем системные свойства для парсера Aalto
            System.setProperty("javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl")
            System.setProperty("javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl")
            System.setProperty("javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl")

            // Принудительно загружаем классы Aalto в ClassLoader
            val inputFactory = Class.forName("com.fasterxml.aalto.stax.InputFactoryImpl").newInstance()
            val outputFactory = Class.forName("com.fasterxml.aalto.stax.OutputFactoryImpl").newInstance()
            val eventFactory = Class.forName("com.fasterxml.aalto.stax.EventFactoryImpl").newInstance()

            Log.d("MiniExcel", "StAX providers successfully initialized: $inputFactory, $outputFactory, $eventFactory")
        } catch (e: Throwable) {
            Log.e("MiniExcel", "Failed to force-init StAX providers", e)
        }
    }

    private fun loadExcelFile(uri: Uri) {
        lifecycleScope.launch(Dispatchers.IO) {
            val originalClassLoader = Thread.currentThread().contextClassLoader
            try {
                // Подменяем ClassLoader потока корутины на ClassLoader приложения
                Thread.currentThread().contextClassLoader = applicationContext.classLoader

                val localFile = File.createTempFile("excel_cache", ".tmp", cacheDir)
                contentResolver.openInputStream(uri)?.use { input ->
                    FileOutputStream(localFile).use { output ->
                        input.copyTo(output)
                    }
                }

                // Открываем файл
                val workbook = WorkbookFactory.create(localFile)
                currentWorkbook = workbook
                formulaEvaluator = workbook.creationHelper.createFormulaEvaluator()

                val sheet = workbook.getSheetAt(0) ?: throw Exception("Лист не найден")
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
                    val cellValue = getCellValueAsString(cell, formatter)

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

    private fun getCellValueAsString(cell: Cell?, formatter: DataFormatter): String {
        if (cell == null) return ""
        return try {
            when (cell.cellTypeEnum) {
                CellType.FORMULA -> {
                    val evaluatedCell = formulaEvaluator?.evaluateInCell(cell)
                    if (evaluatedCell != null) {
                        formatter.formatCellValue(evaluatedCell)
                    } else {
                        cell.cellFormula
                    }
                }
                else -> formatter.formatCellValue(cell)
            }
        } catch (e: Exception) {
            try {
                formatter.formatCellValue(cell)
            } catch (ex: Exception) {
                ""
            }
        }
    }
}
