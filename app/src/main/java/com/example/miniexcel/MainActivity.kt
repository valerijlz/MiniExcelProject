package com.example.miniexcel

import android.annotation.SuppressLint
import android.net.Uri
import android.os.Bundle
import android.util.Log
import android.webkit.JavascriptInterface
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
import org.json.JSONArray
import org.json.JSONObject
import java.io.File
import java.io.FileOutputStream

class MainActivity : AppCompatActivity() {

    private lateinit var webView: WebView
    private var currentWorkbook: Workbook? = null
    private var currentUri: Uri? = null
    private var tempFile: File? = null

    private val openFileLauncher = registerForActivityResult(
        ActivityResultContracts.OpenDocument()
    ) { uri: Uri? ->
        uri?.let {
            currentUri = it
            loadExcelFile(it)
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        webView = findViewById(R.id.webView)
        val btnOpen = findViewById<Button>(R.id.btnOpen)
        val btnSave = findViewById<Button>(R.id.btnSave)

        setupWebView()

        btnOpen.setOnClickListener {
            openFileLauncher.launch(
                arrayOf(
                    "application/vnd.ms-excel",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "application/octet-stream"
                )
            )
        }

        btnSave.setOnClickListener {
            if (currentWorkbook == null) {
                Toast.makeText(this, "Сначала откройте файл", Toast.LENGTH_SHORT).show()
            } else {
                webView.evaluateJavascript("window.sendDiffsToKotlin()", null)
            }
        }

        intent?.data?.let { uri ->
            currentUri = uri
            loadExcelFile(uri)
        }
    }

    @SuppressLint("SetJavaScriptEnabled")
    private fun setupWebView() {
        webView.settings.apply {
            javaScriptEnabled = true
            domStorageEnabled = true
            allowFileAccess = true
            useWideViewPort = true
            loadWithOverviewMode = true
        }
        webView.addJavascriptInterface(WebAppInterface(), "AndroidBridge")
        webView.webViewClient = object : WebViewClient() {
            override fun onPageFinished(view: WebView?, url: String?) {
                super.onPageFinished(view, url)
                currentUri?.let { loadExcelFile(it) }
            }
        }
        webView.loadUrl("file:///android_asset/grid.html")
    }

    /**
     * Безопасное чтение XLS / XLSX через локальный кэш-файл
     */
    private fun loadExcelFile(uri: Uri) {
        lifecycleScope.launch(Dispatchers.IO) {
            try {
                // 1. Копируем данные из Uri во временный файл кэша
                val localFile = File.createTempFile("excel_cache", ".tmp", cacheDir)
                contentResolver.openInputStream(uri)?.use { input ->
                    FileOutputStream(localFile).use { output ->
                        input.copyTo(output)
                    }
                }
                tempFile = localFile

                // 2. POI открывает локальный файл гораздо стабильнее, чем прямой Stream
                currentWorkbook = WorkbookFactory.create(localFile)

                val sheet = currentWorkbook?.getSheetAt(0) ?: throw Exception("Лист в Excel не найден")
                val jsonResult = parseSheetToJSON(sheet)

                withContext(Dispatchers.Main) {
                    webView.evaluateJavascript("window.loadJsonData($jsonResult)", null)
                    Toast.makeText(this@MainActivity, "Файл загружен!", Toast.LENGTH_SHORT).show()
                }
            } catch (t: Throwable) {
                // Ловим любые Throwable (включая NoClassDefFoundError)
                Log.e("MiniExcel", "Error opening Excel file", t)
                withContext(Dispatchers.Main) {
                    val errorMsg = t.localizedMessage ?: t.javaClass.simpleName
                    Toast.makeText(this@MainActivity, "Ошибка: $errorMsg", Toast.LENGTH_LONG).show()
                }
            }
        }
    }

    private fun parseSheetToJSON(sheet: Sheet): String {
        val root = JSONObject()
        val rowsArray = JSONArray()
        val evaluator = try {
            currentWorkbook?.creationHelper?.createFormulaEvaluator()
        } catch (e: Exception) {
            null
        }

        val lastRowNum = sheet.lastRowNum
        val maxRows = minOf(lastRowNum, 2000)

        for (r in 0..maxRows) {
            val row = sheet.getRow(r) ?: continue
            val rowObj = JSONObject()
            val cellsArray = JSONArray()
            val lastCellNum = row.lastCellNum.toInt()

            if (lastCellNum < 0) continue

            var hasDataInRow = false
            for (c in 0 until minOf(lastCellNum, 128)) {
                val cell = row.getCell(c)
                val cellObj = JSONObject()
                cellObj.put("r", r)
                cellObj.put("c", c)

                if (cell != null) {
                    val value = getCellValueAsString(cell, evaluator)
                    if (value.isNotEmpty()) {
                        cellObj.put("v", value)
                        hasDataInRow = true
                    }

                    val bgHex = getSafeBackgroundColor(cell)
                    if (!bgHex.isNullOrEmpty()) {
                        cellObj.put("bg", bgHex as String)
                    }
                }
                cellsArray.put(cellObj)
            }

            if (hasDataInRow || cellsArray.length() > 0) {
                rowObj.put("r", r)
                rowObj.put("cells", cellsArray)
                rowsArray.put(rowObj)
            }
        }

        root.put("maxRows", maxRows + 1)
        root.put("maxCols", 26)
        root.put("rows", rowsArray)
        return root.toString()
    }

    private fun getCellValueAsString(cell: Cell, evaluator: FormulaEvaluator?): String {
        return try {
            @Suppress("DEPRECATION")
            val type = cell.cellTypeEnum

            when (type) {
                CellType.STRING -> cell.stringCellValue ?: ""
                CellType.NUMERIC -> {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        cell.dateCellValue?.toString() ?: ""
                    } else {
                        val num = cell.numericCellValue
                        if (num == num.toLong().toDouble()) num.toLong().toString() else num.toString()
                    }
                }
                CellType.BOOLEAN -> cell.booleanCellValue.toString()
                CellType.FORMULA -> {
                    if (evaluator != null) {
                        val evaluated = evaluator.evaluate(cell)
                        @Suppress("DEPRECATION")
                        when (evaluated.cellTypeEnum) {
                            CellType.NUMERIC -> evaluated.numberValue.toString()
                            CellType.STRING -> evaluated.stringValue ?: ""
                            CellType.BOOLEAN -> evaluated.booleanValue.toString()
                            else -> ""
                        }
                    } else {
                        cell.cellFormula ?: ""
                    }
                }
                else -> ""
            }
        } catch (e: Exception) {
            ""
        }
    }

    private fun getSafeBackgroundColor(cell: Cell): String? {
        return try {
            val fill = cell.cellStyle?.fillForegroundColorColor ?: return null
            val hex = fill.toString()
            if (hex.length >= 6) "#${hex.takeLast(6)}" else null
        } catch (t: Throwable) {
            null
        }
    }

    inner class WebAppInterface {

        @JavascriptInterface
        fun saveChanges(diffJsonString: String) {
            lifecycleScope.launch(Dispatchers.IO) {
                try {
                    val wb = currentWorkbook ?: return@launch
                    val uri = currentUri ?: return@launch
                    val sheet = wb.getSheetAt(0)

                    val diffArray = JSONArray(diffJsonString)
                    for (i in 0 until diffArray.length()) {
                        val change = diffArray.getJSONObject(i)
                        val r = change.getInt("r")
                        val c = change.getInt("c")
                        val newValue = change.getString("v")

                        var row = sheet.getRow(r)
                        if (row == null) row = sheet.createRow(r)

                        var cell = row.getCell(c)
                        if (cell == null) cell = row.createCell(c)

                        val doubleValue = newValue.toDoubleOrNull()
                        if (doubleValue != null) {
                            cell.setCellValue(doubleValue)
                        } else {
                            cell.setCellValue(newValue)
                        }
                    }

                    // Сохраняем результат обратно в исходный Uri
                    contentResolver.openOutputStream(uri, "rwt")?.use { os ->
                        wb.write(os)
                        os.flush()
                    }

                    withContext(Dispatchers.Main) {
                        Toast.makeText(this@MainActivity, "Сохранено успешно", Toast.LENGTH_SHORT).show()
                    }
                } catch (t: Throwable) {
                    Log.e("MiniExcel", "Error saving workbook", t)
                    withContext(Dispatchers.Main) {
                        Toast.makeText(this@MainActivity, "Ошибка сохранения: ${t.localizedMessage}", Toast.LENGTH_LONG).show()
                    }
                }
            }
        }
    }
}
