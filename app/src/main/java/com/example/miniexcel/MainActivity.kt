package com.example.miniexcel

import android.net.Uri
import android.os.Bundle
import android.util.Log
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.json.JSONArray
import org.json.JSONObject
import java.io.ByteArrayOutputStream
import java.io.InputStream

class MainActivity : AppCompatActivity() {

    private lateinit var webView: WebView
    private var cachedJsonPayload: String = "{}"

    // Лаунчер для ЭКСПОРТА (Сохранения)
    private val createFileLauncher = registerForActivityResult(
        ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    ) { uri: Uri? ->
        uri?.let { saveCurrentDataToUri(it) }
    }

    // 1. ДОБАВЛЕНО: Лаунчер для ОТКРЫТИЯ файла
    private val openFileLauncher = registerForActivityResult(
        ActivityResultContracts.OpenDocument()
    ) { uri: Uri? ->
        uri?.let { readXlsxFromUri(it) }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        webView = findViewById(R.id.webView)
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

    fun updateGridData(jsonPayload: String) {
        this.cachedJsonPayload = jsonPayload
        sendJsonToWebView(jsonPayload)
    }

    private fun sendJsonToWebView(jsonStr: String) {
        val escapedJson = jsonStr.replace("\\", "\\\\").replace("'", "\\'")
        webView.evaluateJavascript("javascript:window.loadJsonData('$escapedJson');", null)
    }

    // 2. ДОБАВЛЕНО: Вызовите этот метод при нажатии на кнопку "Открыть"
    fun openXlsx() {
        openFileLauncher.launch(arrayOf(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/vnd.ms-excel",
            "*/*"
        ))
    }

    // 3. ДОБАВЛЕНО: Чтение и парсинг XLSX файла
    private fun readXlsxFromUri(uri: Uri) {
        lifecycleScope.launch(Dispatchers.IO) {
            try {
                val inputStream: InputStream? = contentResolver.openInputStream(uri)
                    ?: throw Exception("Не удалось открыть файл")

                val workbook = XSSFWorkbook(inputStream)
                val sheet = workbook.getSheetAt(0) ?: throw Exception("Лист не найден")

                val matrixArray = JSONArray()
                val mergesArray = JSONArray()

                // Чтение строк и ячеек
                for (r in 0..sheet.lastRowNum) {
                    val row = sheet.getRow(r)
                    val rowArray = JSONArray()
                    if (row != null) {
                        for (c in 0 until row.lastCellNum) {
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
                                rowArray.put(c, cellObj)
                            } else {
                                rowArray.put(c, JSONObject.NULL)
                            }
                        }
                    }
                    matrixArray.put(r, rowArray)
                }

                // Чтение объединенных ячеек
                for (i in 0 until sheet.numMergedRegions) {
                    val region: CellRangeAddress = sheet.getMergedRegion(i)
                    val mergeObj = JSONObject().apply {
                        put("sr", region.firstRow)
                        put("sc", region.firstColumn)
                        put("er", region.lastRow)
                        put("ec", region.lastColumn)
                    }
                    mergesArray.put(mergeObj)
                }

                workbook.close()
                inputStream.close()

                val resultJson = JSONObject().apply {
                    put("matrix", matrixArray)
                    put("merges", mergesArray)
                }

                withContext(Dispatchers.Main) {
                    updateGridData(resultJson.toString())
                    Toast.makeText(this@MainActivity, "Файл успешно открыт", Toast.LENGTH_SHORT).show()
                }

            } catch (e: Exception) {
                Log.e("OpenError", "Ошибка открытия XLSX", e)
                withContext(Dispatchers.Main) {
                    Toast.makeText(this@MainActivity, "Ошибка открытия: ${e.message}", Toast.LENGTH_LONG).show()
                }
            }
        }
    }

    fun exportToXlsx() {
        createFileLauncher.launch("Exported_Data.xlsx")
    }

    private fun saveCurrentDataToUri(uri: Uri) {
        lifecycleScope.launch(Dispatchers.IO) {
            try {
                val outputStream = contentResolver.openOutputStream(uri, "rwt")
                    ?: throw Exception("Не удалось открыть поток для записи")

                val workbook = XSSFWorkbook()
                val sheet = workbook.createSheet("Sheet1")

                val jsonObj = JSONObject(cachedJsonPayload)
                val matrixArray = jsonObj.optJSONArray("matrix") ?: JSONArray()
                val mergesArray = jsonObj.optJSONArray("merges") ?: JSONArray()

                // 1. Заполнение ячеек
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

                // 2. Восстановление объединённых областей
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

                // 3. Безопасная запись через буфер байтов
                val bytesOut = ByteArrayOutputStream()
                workbook.write(bytesOut)
                workbook.close()

                outputStream.write(bytesOut.toByteArray())
                outputStream.flush()
                outputStream.close()

                withContext(Dispatchers.Main) {
                    Toast.makeText(this@MainActivity, "Файл успешно сохранён", Toast.LENGTH_SHORT).show()
                }
            } catch (e: Exception) {
                Log.e("SaveError", "Ошибка сохранения XLSX", e)
                withContext(Dispatchers.Main) {
                    Toast.makeText(this@MainActivity, "Ошибка сохранения: ${e.message}", Toast.LENGTH_LONG).show()
                }
            }
        }
    }
}
