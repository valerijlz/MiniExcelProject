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
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.json.JSONArray
import org.json.JSONObject
import java.io.ByteArrayOutputStream

class MainActivity : AppCompatActivity() {

    private lateinit var webView: WebView
    private var cachedJsonPayload: String = "{}"

    private val createFileLauncher = registerForActivityResult(
        ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    ) { uri: Uri? ->
        uri?.let { saveCurrentDataToUri(it) }
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

    fun exportToXlsx() {
        createFileLauncher.launch("Exported_Data.xlsx")
    }

    private fun saveCurrentDataToUri(uri: Uri) {
        lifecycleScope.launch(Dispatchers.IO) {
            try {
                // "rwt" - усечение и очистка файла перед записью (предотвращает битые ZIP-архивы XLSX)
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
