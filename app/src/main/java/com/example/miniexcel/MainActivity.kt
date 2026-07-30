package com.example.miniexcel

import android.net.Uri
import android.os.Bundle
import android.util.Log
import android.view.View
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.Button
import android.widget.ProgressBar
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.json.JSONArray
import org.json.JSONObject
import java.io.InputStream
import java.io.OutputStream

class MainActivity : AppCompatActivity() {

    private lateinit var webView: WebView
    private lateinit var progressBar: ProgressBar
    private var cachedJsonPayload: String = "{}"

    private val createFileLauncher = registerForActivityResult(
        ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    ) { uri: Uri? ->
        uri?.let { saveCurrentDataToUri(it) }
    }

    private val openFileLauncher = registerForActivityResult(
        ActivityResultContracts.OpenDocument()
    ) { uri: Uri? ->
        uri?.let { readExcelFromUri(it) }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        webView = findViewById(R.id.webView)
        progressBar = findViewById(R.id.progressBar)

        findViewById<Button>(R.id.btnOpen).setOnClickListener { openExcel() }
        findViewById<Button>(R.id.btnSave).setOnClickListener { exportToXlsx() }

        setupWebView()
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
                if (cachedJsonPayload.isNotEmpty() && cachedJsonPayload != "{}") {
                    sendJsonToWebView(cachedJsonPayload)
                }
            }
        }

        webView.loadUrl("file:///android_asset/grid.html")
    }

    private fun updateGridData(jsonPayload: String) {
        this.cachedJsonPayload = jsonPayload
        sendJsonToWebView(jsonPayload)
    }

    private fun sendJsonToWebView(jsonStr: String) {
        val escapedJson = jsonStr.replace("\\", "\\\\").replace("'", "\\'")
        webView.evaluateJavascript("javascript:window.loadJsonData('$escapedJson');", null)
    }

    fun openExcel() {
        openFileLauncher.launch(
            arrayOf(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-excel",
                "*/*"
            )
        )
    }

    private fun readExcelFromUri(uri: Uri) {
        progressBar.visibility = View.VISIBLE
        lifecycleScope.launch(Dispatchers.IO) {
            try {
                val inputStream: InputStream = contentResolver.openInputStream(uri)
                    ?: throw Exception("Не удалось открыть файл")

                val resultPayload = inputStream.use { stream ->
                    // WorkbookFactory поддерживает и .xls, и .xlsx!
                    val workbook = WorkbookFactory.create(stream)
                    val sheet = workbook.getSheetAt(0) ?: throw Exception("Лист не найден")

                    val matrixArray = JSONArray()
                    val mergesArray = JSONArray()
                    val widthsObj = JSONObject()
                    val heightsObj = JSONObject()

                    var maxCol = 0
                    val maxRowsToRead = Math.min(sheet.lastRowNum + 1, 1000)

                    for (r in 0 until maxRowsToRead) {
                        val row = sheet.getRow(r)
                        val rowArray = JSONArray()
                        if (row != null) {
                            heightsObj.put(r.toString(), (row.heightInPoints * 1.33).toInt())
                            if (row.lastCellNum > maxCol) maxCol = Math.min(row.lastCellNum.toInt(), 50)

                            for (c in 0 until maxCol) {
                                val cell = row.getCell(c)
                                if (cell != null) {
                                    val cellObj = JSONObject()

                                    val valStr = when (cell.cellType) {
                                        CellType.NUMERIC -> cell.numericCellValue.toString().removeSuffix(".0")
                                        CellType.STRING -> cell.stringCellValue
                                        CellType.BOOLEAN -> cell.booleanCellValue.toString()
                                        CellType.FORMULA -> {
                                            try { cell.stringCellValue } catch (e: Exception) { cell.numericCellValue.toString() }
                                        }
                                        else -> ""
                                    }
                                    cellObj.put("v", valStr)

                                    // Оптимизированный разбор стилей без тяжелой работы с XSSFColor
                                    val style = cell.cellStyle
                                    if (style != null) {
                                        val borderObj = JSONObject()
                                        if (style.borderTop != BorderStyle.NONE) borderObj.put("top", JSONObject().put("width", 1))
                                        if (style.borderBottom != BorderStyle.NONE) borderObj.put("bottom", JSONObject().put("width", 1))
                                        if (style.borderLeft != BorderStyle.NONE) borderObj.put("left", JSONObject().put("width", 1))
                                        if (style.borderRight != BorderStyle.NONE) borderObj.put("right", JSONObject().put("width", 1))

                                        if (borderObj.length() > 0) cellObj.put("borders", borderObj)
                                    }

                                    rowArray.put(c, cellObj)
                                } else {
                                    rowArray.put(c, JSONObject.NULL)
                                }
                            }
                        }
                        matrixArray.put(r, rowArray)
                    }

                    for (c in 0 until maxCol) {
                        val colW = sheet.getColumnWidth(c)
                        widthsObj.put(c.toString(), (colW / 256.0 * 7.5).toInt())
                    }

                    for (i in 0 until sheet.numMergedRegions) {
                        val region: CellRangeAddress = sheet.getMergedRegion(i)
                        mergesArray.put(JSONObject().apply {
                            put("sr", region.firstRow)
                            put("sc", region.firstColumn)
                            put("er", region.lastRow)
                            put("ec", region.lastColumn)
                        })
                    }

                    workbook.close()

                    JSONObject().apply {
                        put("matrix", matrixArray)
                        put("merges", mergesArray)
                        put("widths", widthsObj)
                        put("heights", heightsObj)
                    }.toString()
                }

                withContext(Dispatchers.Main) {
                    progressBar.visibility = View.GONE
                    updateGridData(resultPayload)
                    Toast.makeText(this@MainActivity, "Файл успешно открыт", Toast.LENGTH_SHORT).show()
                }

            } catch (e: Exception) {
                Log.e("OpenError", "Ошибка открытия Excel", e)
                withContext(Dispatchers.Main) {
                    progressBar.visibility = View.GONE
                    Toast.makeText(this@MainActivity, "Ошибка открытия: ${e.message}", Toast.LENGTH_LONG).show()
                }
            }
        }
    }

    fun exportToXlsx() {
        webView.evaluateJavascript("javascript:window.getCurrentDataJson();") { json ->
            if (json != null && json != "null") {
                var cleanJson = json
                if (cleanJson.startsWith("\"") && cleanJson.endsWith("\"")) {
                    cleanJson = cleanJson.substring(1, cleanJson.length - 1)
                        .replace("\\\"", "\"")
                        .replace("\\\\", "\\")
                }
                cachedJsonPayload = cleanJson
            }
            createFileLauncher.launch("Exported_Data.xlsx")
        }
    }

    private fun saveCurrentDataToUri(uri: Uri) {
        progressBar.visibility = View.VISIBLE
        lifecycleScope.launch(Dispatchers.IO) {
            try {
                val outputStream: OutputStream = contentResolver.openOutputStream(uri, "rwt")
                    ?: throw Exception("Не удалось открыть поток для записи")

                outputStream.use { stream ->
                    val workbook = XSSFWorkbook()
                    val sheet = workbook.createSheet("Sheet1")

                    val jsonObj = JSONObject(cachedJsonPayload)
                    val matrixArray = jsonObj.optJSONArray("matrix") ?: JSONArray()
                    val mergesArray = jsonObj.optJSONArray("merges") ?: JSONArray()

                    for (r in 0 until matrixArray.length()) {
                        val rowArray = matrixArray.optJSONArray(r) ?: continue
                        val row = sheet.createRow(r)
                        for (c in 0 until rowArray.length()) {
                            val cellData = rowArray.opt(c)
                            var textValue = ""
                            if (cellData is JSONObject) {
                                textValue = cellData.optString("v", cellData.optString("value", ""))
                            } else if (cellData != null && cellData != JSONObject.NULL) {
                                textValue = cellData.toString()
                            }

                            if (textValue.isNotEmpty()) {
                                val cell = row.createCell(c)
                                cell.setCellValue(textValue)
                            }
                        }
                    }

                    for (i in 0 until mergesArray.length()) {
                        val m = mergesArray.optJSONObject(i) ?: continue
                        val sr = m.optInt("sr", m.optInt("startRow", -1))
                        val sc = m.optInt("sc", m.optInt("startCol", -1))
                        val er = m.optInt("er", m.optInt("endRow", -1))
                        val ec = m.optInt("ec", m.optInt("endCol", -1))

                        if (sr >= 0 && sc >= 0 && er >= sr && ec >= sc) {
                            sheet.addMergedRegion(CellRangeAddress(sr, er, sc, ec))
                        }
                    }

                    workbook.write(stream)
                    workbook.close()
                }

                withContext(Dispatchers.Main) {
                    progressBar.visibility = View.GONE
                    Toast.makeText(this@MainActivity, "Файл успешно сохранён", Toast.LENGTH_SHORT).show()
                }
            } catch (e: Exception) {
                Log.e("SaveError", "Ошибка сохранения XLSX", e)
                withContext(Dispatchers.Main) {
                    progressBar.visibility = View.GONE
                    Toast.makeText(this@MainActivity, "Ошибка сохранения: ${e.message}", Toast.LENGTH_LONG).show()
                }
            }
        }
    }
}
