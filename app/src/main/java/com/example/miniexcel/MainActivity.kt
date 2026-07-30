package com.example.miniexcel

import android.annotation.SuppressLint
import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.util.Log
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

class MainActivity : AppCompatActivity() {

    private lateinit var webView: WebView
    private lateinit var btnOpen: Button
    private lateinit var btnSave: Button

    private var currentWorkbook: Workbook? = null
    private var currentUri: Uri? = null

    // Вызов системного проводника для выбора .xls / .xlsx
    private val openFileLauncher = registerForActivityResult(
        ActivityResultContracts.OpenDocument()
    ) { uri ->
        uri?.let {
            currentUri = it
            // Попытка получить постоянные права (безопасно, без краша)
            try {
                contentResolver.takePersistableUriPermission(
                    it, Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_GRANT_WRITE_URI_PERMISSION
                )
            } catch (e: SecurityException) {
                Log.w("MiniExcel", "Persistable permission not granted, proceeding with temporary URI access: ${e.message}")
            }
            loadExcelFile(it)
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        webView = findViewById(R.id.webView)
        btnOpen = findViewById(R.id.btnOpen)
        btnSave = findViewById(R.id.btnSave)

        setupWebView()

        btnOpen.setOnClickListener {
            openFileLauncher.launch(
                arrayOf(
                    "application/vnd.ms-excel",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "application/octet-stream" // Добавлено для совместимости с некоторыми файловыми менеджерами
                )
            )
        }

        btnSave.setOnClickListener {
            saveExcelFile()
        }
    }

    @SuppressLint("SetJavaScriptEnabled")
    private fun setupWebView() {
        webView.settings.apply {
            javaScriptEnabled = true
            domStorageEnabled = true
            allowFileAccess = true
        }
        webView.webViewClient = object : WebViewClient() {
            override fun onPageFinished(view: WebView?, url: String?) {
                super.onPageFinished(view, url)
                currentUri?.let { loadExcelFile(it) }
            }
        }
        webView.loadUrl("file:///android_asset/grid.html")
    }

    private fun loadExcelFile(uri: Uri) {
        lifecycleScope.launch(Dispatchers.IO) {
            try {
                // Закрываем предыдущий Workbook, если он был открыт
                currentWorkbook?.close()

                contentResolver.openInputStream(uri)?.use { inputStream ->
                    // WorkbookFactory автоматически определяет XLS (HSSF) или XLSX (XSSF)
                    currentWorkbook = WorkbookFactory.create(inputStream)
                }

                val sheet = currentWorkbook?.getSheetAt(0)
                if (sheet == null) {
                    withContext(Dispatchers.Main) {
                        Toast.makeText(this@MainActivity, "Файл не содержит листов", Toast.LENGTH_SHORT).show()
                    }
                    return@launch
                }

                val jsonResult = parseSheetToJSON(sheet)

                withContext(Dispatchers.Main) {
                    webView.evaluateJavascript("window.loadJsonData($jsonResult)", null)
                }
            } catch (e: Exception) {
                Log.e("MiniExcel", "Error loading file", e)
                withContext(Dispatchers.Main) {
                    Toast.makeText(this@MainActivity, "Ошибка открытия: ${e.localizedMessage}", Toast.LENGTH_LONG).show()
                }
            }
        }
    }

    private fun parseSheetToJSON(sheet: Sheet): String {
        val root = JSONObject()
        val rowsArray = JSONArray()
        
        // Перехватываем возможные ошибки при вычислении формул в POI
        val evaluator = try {
            currentWorkbook?.creationHelper?.createFormulaEvaluator()
        } catch (e: Exception) {
            null
        }

        val lastRowNum = sheet.lastRowNum
        val maxRows = minOf(lastRowNum, 300)

        for (r in 0..maxRows) {
            val row = sheet.getRow(r) ?: continue
            val rowObj = JSONObject()
            val cellsArray = JSONArray()
            val lastCellNum = row.lastCellNum.toInt()

            if (lastCellNum < 0) continue

            for (c in 0 until minOf(lastCellNum, 50)) {
                val cell = row.getCell(c) ?: continue
                val cellObj = JSONObject()
                cellObj.put("r", r)
                cellObj.put("c", c)

                val value = getCellValueAsString(cell, evaluator)
                if (value.isNotEmpty()) {
                    cellObj.put("v", value)
                }
                cellsArray.put(cellObj)
            }

            rowObj.put("r", r)
            rowObj.put("cells", cellsArray)
            rowsArray.put(rowObj)
        }

        root.put("maxRows", maxOf(maxRows + 1, 50))
        root.put("maxCols", 26)
        root.put("rows", rowsArray)
        return root.toString()
    }

    private fun getCellValueAsString(cell: Cell, evaluator: FormulaEvaluator?): String {
        return try {
            when (cell.cellType) {
                CellType.STRING -> cell.stringCellValue
                CellType.NUMERIC -> {
                    if (DateUtil.isCellDateFormatted(cell)) cell.dateCellValue.toString()
                    else cell.numericCellValue.let {
                        if (it == it.toLong().toDouble()) it.toLong().toString() else it.toString()
                    }
                }
                CellType.BOOLEAN -> cell.booleanCellValue.toString()
                CellType.FORMULA -> {
                    evaluator?.let {
                        val evaluated = it.evaluate(cell)
                        when (evaluated.cellType) {
                            CellType.NUMERIC -> evaluated.numberValue.toString()
                            CellType.STRING -> evaluated.stringValue
                            CellType.BOOLEAN -> evaluated.booleanValue.toString()
                            else -> ""
                        }
                    } ?: cell.cellFormula
                }
                else -> ""
            }
        } catch (e: Exception) {
            ""
        }
    }

    private fun saveExcelFile() {
        val uri = currentUri ?: run {
            Toast.makeText(this, "Нет открытого файла для сохранения", Toast.LENGTH_SHORT).show()
            return
        }

        webView.evaluateJavascript("window.getDiffsJson()") { diffJsonString ->
            if (diffJsonString == null || diffJsonString == "null") return@evaluateJavascript

            val unescapedJson = if (diffJsonString.startsWith("\"") && diffJsonString.endsWith("\"")) {
                diffJsonString.substring(1, diffJsonString.length - 1).replace("\\\"", "\"")
            } else diffJsonString

            lifecycleScope.launch(Dispatchers.IO) {
                try {
                    val wb = currentWorkbook ?: return@launch
                    val sheet = wb.getSheetAt(0)
                    val diffArray = JSONArray(unescapedJson)

                    for (i in 0 until diffArray.length()) {
                        val change = diffArray.getJSONObject(i)
                        val r = change.getInt("r")
                        val c = change.getInt("c")
                        val newValue = change.getString("v")

                        var row = sheet.getRow(r)
                        if (row == null) row = sheet.createRow(r)

                        var cell = row.getCell(c)
                        if (cell == null) cell = row.createCell(c)

                        val numVal = newValue.toDoubleOrNull()
                        if (numVal != null) {
                            cell.setCellValue(numVal)
                        } else {
                            cell.setCellValue(newValue)
                        }
                    }

                    contentResolver.openOutputStream(uri, "rwt")?.use { os ->
                        wb.write(os)
                        os.flush()
                    }

                    withContext(Dispatchers.Main) {
                        Toast.makeText(this@MainActivity, "Файл успешно сохранен!", Toast.LENGTH_SHORT).show()
                    }
                } catch (e: Exception) {
                    Log.e("MiniExcel", "Save error", e)
                    withContext(Dispatchers.Main) {
                        Toast.makeText(this@MainActivity, "Ошибка сохранения: ${e.localizedMessage}", Toast.LENGTH_LONG).show()
                    }
                }
            }
        }
    }
}
