package com.miniexcel.app

import android.annotation.SuppressLint
import android.net.Uri
import android.os.Bundle
import android.util.Log
import android.webkit.JavascriptInterface
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.poi.ss.usermodel.*
import org.json.JSONArray
import org.json.JSONObject
import java.io.InputStream
import java.io.OutputStream

class MainActivity : AppCompatActivity() {

    private lateinit var webView: WebView
    private var currentWorkbook: Workbook? = null
    private var currentUri: Uri? = null

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        webView = WebView(this)
        setContentView(webView)

        setupWebView()

        // Входной Uri (например, из Intent)
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
                // Если файл уже открыт, передаем данные после загрузки HTML
                currentUri?.let { loadExcelFile(it) }
            }
        }
        webView.loadUrl("file:///android_asset/grid.html")
    }

    /**
     * Фича A: Безопасное чтение XLS / XLSX без полного сканирования тяжелых стилей
     */
    private fun loadExcelFile(uri: Uri) {
        lifecycleScope.launch(Dispatchers.IO) {
            try {
                contentResolver.openInputStream(uri)?.use { inputStream ->
                    // WorkbookFactory корректно распознает как HSSF (.xls), так и XSSF (.xlsx)
                    currentWorkbook = WorkbookFactory.create(inputStream)
                }

                val sheet = currentWorkbook?.getSheetAt(0) ?: return@launch
                val jsonResult = parseSheetToJSON(sheet)

                withContext(Dispatchers.Main) {
                    webView.evaluateJavascript("window.loadJsonData($jsonResult)", null)
                }
            } catch (e: Exception) {
                Log.e("MiniExcel", "Error loading workbook", e)
                withContext(Dispatchers.Main) {
                    Toast.makeText(this@MainActivity, "Ошибка открытия файла", Toast.LENGTH_SHORT).show()
                }
            }
        }
    }

    private fun parseSheetToJSON(sheet: Sheet): String {
        val root = JSONObject()
        val rowsArray = JSONArray()
        val evaluator = currentWorkbook?.creationHelper?.createFormulaEvaluator()

        val lastRowNum = sheet.lastRowNum
        // Ограничиваем считывание только активными строками
        val maxRows = minOf(lastRowNum, 5000) 

        for (r in 0..maxRows) {
            val row = sheet.getRow(r) ?: continue
            val rowObj = JSONObject()
            val cellsArray = JSONArray()
            val lastCellNum = row.lastCellNum.toInt()

            if (lastCellNum < 0) continue

            var hasDataInRow = false
            for (c in 0 until minOf(lastCellNum, 256)) { // Ограничение по колонкам (A..Z+)
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
                    
                    // Безопасное извлечение фонового цвета без java.awt
                    val bgHex = getSafeBackgroundColor(cell)
                    if (bgHex != null) {
                        cellObj.put("bg", bgHex)
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
        root.put("maxCols", 26) // По умолчанию сетка A-Z
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

    private fun getSafeBackgroundColor(cell: Cell): String? {
        return try {
            val fill = cell.cellStyle.fillForegroundColorColor ?: return null
            // Извлекаем RGB hex безопасно
            val hex = fill.toString() 
            if (hex.length >= 6) "#${hex.takeLast(6)}" else null
        } catch (t: Throwable) {
            null // Игнорируем ошибки отсутствия java.awt классов
        }
    }

    /**
     * Фича C: Мост синхронизации и сохранение диффов без потери исходного форматирования
     */
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

                        // Обновляем исключительно значение, сохраняя существующий CellStyle
                        if (newValue.toDoubleOrNull() != null) {
                            cell.setCellValue(newValue.toDouble())
                        } else {
                            cell.setCellValue(newValue)
                        }
                    }

                    // Перезапись исходного файла
                    contentResolver.openOutputStream(uri, "rwt")?.use { os ->
                        wb.write(os)
                        os.flush()
                    }

                    withContext(Dispatchers.Main) {
                        Toast.makeText(this@MainActivity, "Сохранено успешно", Toast.LENGTH_SHORT).show()
                    }
                } catch (e: Exception) {
                    Log.e("MiniExcel", "Error saving workbook", e)
                    withContext(Dispatchers.Main) {
                        Toast.makeText(this@MainActivity, "Ошибка сохранения", Toast.LENGTH_SHORT).show()
                    }
                }
            }
        }
    }
}
