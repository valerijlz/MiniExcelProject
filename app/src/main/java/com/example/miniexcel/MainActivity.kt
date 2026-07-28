package com.example.exceltableView // Замените на ваш пакет

import android.app.Activity
import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.webkit.JavascriptInterface
import android.webkit.WebSettings
import android.webkit.WebView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.json.JSONArray
import org.json.JSONObject
import java.io.File
import java.io.FileOutputStream

class MainActivity : AppCompatActivity() {

    private lateinit var tableWebView: WebView
    private var cachedJsonPayload: String = "{\"matrix\":[],\"widths\":[],\"heights\":[],\"merges\":[]}"

    companion object {
        private const val PICK_FILE_REQUEST_CODE = 100
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        tableWebView = findViewById(R.id.tableWebView) // Убедитесь, что ID совпадает с вашим XML
        setupWebView()
    }

    private fun setupWebView() {
        val webSettings: WebSettings = tableWebView.settings
        webSettings.javaScriptEnabled = true
        webSettings.domStorageEnabled = true
        webSettings.allowFileAccess = true

        // Подключаем мост для связи с JavaScript
        tableWebView.addJavascriptInterface(AndroidBridge(), "AndroidBridge")

        // Загружаем HTML из папки assets
        tableWebView.loadUrl("file:///android_asset/grid.html")
    }

    // Метод для вызова выбора файла (например, по нажатию кнопки)
    fun openFileSelector() {
        val intent = Intent(Intent.ACTION_GET_CONTENT).apply {
            type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            addCategory(Intent.CATEGORY_OPENABLE)
        }
        startActivityForResult(Intent.createChooser(intent, "Выберите Excel файл"), PICK_FILE_REQUEST_CODE)
    }

    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)
        if (requestCode == PICK_FILE_REQUEST_CODE && resultCode == Activity.RESULT_OK) {
            data?.data?.let { uri ->
                processExcelUri(uri)
            }
        }
    }

    private fun processExcelUri(uri: Uri) {
        try {
            val file = getFileFromUri(uri)
            val jsonString = parseExcelFile(file)
            cachedJsonPayload = jsonString

            // Перезагружаем страницу, чтобы инициализировать таблицу новыми данными
            tableWebView.loadUrl("file:///android_asset/grid.html")
        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(this, "Ошибка чтения файла: ${e.localizedMessage}", Toast.LENGTH_LONG).show()
        }
    }

    private fun getFileFromUri(uri: Uri): File {
        val inputStream = contentResolver.openInputStream(uri)
        val tempFile = File(cacheDir, "temp_excel.xlsx")
        val outputStream = FileOutputStream(tempFile)
        inputStream?.use { input ->
            outputStream.use { output ->
                input.copyTo(output)
            }
        }
        return tempFile
    }

    private fun parseExcelFile(file: File): String {
        WorkbookFactory.create(file).use { workbook ->
            val sheet = workbook.getSheetAt(0) ?: return cachedJsonPayload
            val matrix = JSONArray()
            val widths = JSONArray()
            val heights = JSONArray()
            val merges = JSONArray()

            val maxRow = sheet.lastRowNum
            var maxCol = 0

            for (r in 0..maxRow) {
                val row = sheet.getRow(r)
                if (row != null && row.lastCellNum > maxCol) {
                    maxCol = row.lastCellNum.toInt()
                }
            }
            maxCol = Math.max(maxCol, 15)

            for (r in 0..Math.max(maxRow, 29)) {
                val row = sheet.getRow(r)
                val rowArray = JSONArray()
                val h = if (row != null && row.height > 0) (row.height / 20 * 1.33).toInt() else 22
                if (r == 0) {
                    for (c in 0 until maxCol) {
                        widths.put(if (sheet.getColumnWidth(c) > 0) sheet.getColumnWidth(c) / 37 else 70)
                    }
                }
                heights.put(h)

                for (c in 0 until maxCol) {
                    val cell = row?.getCell(c)
                    val cellObj = JSONObject()
                    var cellVal = ""
                    if (cell != null) {
                        cellVal = when (cell.cellType) {
                            org.apache.poi.ss.usermodel.CellType.STRING -> cell.stringCellValue
                            org.apache.poi.ss.usermodel.CellType.NUMERIC -> {
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    cell.dateCellValue.toString()
                                } else {
                                    val num = cell.numericCellValue
                                    if (num == num.toInt().toDouble()) num.toInt().toString() else num.toString()
                                }
                            }
                            org.apache.poi.ss.usermodel.CellType.BOOLEAN -> cell.booleanCellValue.toString()
                            org.apache.poi.ss.usermodel.CellType.FORMULA -> cell.cellFormula
                            else -> ""
                        }
                    }
                    cellObj.put("v", cellVal)
                    rowArray.put(cellObj)
                }
                matrix.put(rowArray)
            }

            for (i in 0 until sheet.numMergedRegions) {
                val region = sheet.getMergedRegion(i)
                val mObj = JSONObject().apply {
                    put("sr", region.firstRow)
                    put("sc", region.firstColumn)
                    put("er", region.lastRow)
                    put("ec", region.lastColumn)
                }
                merges.put(mObj)
            }

            return JSONObject().apply {
                put("matrix", matrix)
                put("widths", widths)
                put("heights", heights)
                put("merges", merges)
            }.toString()
        }
    }

    // Мост для взаимодействия с WebView / JavaScript
    inner class AndroidBridge {
        @JavascriptInterface
        fun getExcelData(): String {
            return cachedJsonPayload
        }
    }
}
