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
import org.apache.poi.ss.util.CellRangeAddress
import org.json.JSONArray
import org.json.JSONObject
import java.io.File
import java.io.FileOutputStream

class MainActivity : AppCompatActivity() {

    private lateinit var webView: WebView
    private var currentWorkbook: Workbook? = null
    private var currentUri: Uri? = null

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

        // Исправление XML-парсеров StAX для Android
        fixXmlStaxProviders()

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

    private fun fixXmlStaxProviders() {
        try {
            System.setProperty("javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl")
            System.setProperty("javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl")
            System.setProperty("javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl")
        } catch (t: Throwable) {
            Log.e("MiniExcel", "Failed to set StAX providers", t)
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

    private fun initPoiEnvironment() {
        try {
            // Исправление ClassLoader для корректной загрузки Marshallers в POI OpenXML
            Thread.currentThread().contextClassLoader = MainActivity::class.java.classLoader
        } catch (e: Exception) {
            Log.e("MiniExcel", "Failed to set contextClassLoader", e)
        }
    }

    private fun loadExcelFile(uri: Uri) {
        lifecycleScope.launch(Dispatchers.IO) {
            try {
                initPoiEnvironment()

                val localFile = File.createTempFile("excel_cache", ".tmp", cacheDir)
                contentResolver.openInputStream(uri)?.use { input ->
                    FileOutputStream(localFile).use { output ->
                        input.copyTo(output)
                    }
                }

                currentWorkbook = WorkbookFactory.create(localFile)
                val sheet = currentWorkbook?.getSheetAt(0) ?: throw Exception("Лист в Excel не найден")
                val jsonResult = parseSheetToJSON(sheet)

                withContext(Dispatchers.Main) {
                    webView.evaluateJavascript("window.loadJsonData($jsonResult)", null)
                    Toast.makeText(this@MainActivity, "Файл загружен!", Toast.LENGTH_SHORT).show()
                }
            } catch (t: Throwable) {
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

        // 1. Объединения ячеек
        val mergedArray = JSONArray()
        for (i in 0 until sheet.numMergedRegions) {
            val range: CellRangeAddress = sheet.getMergedRegion(i)
            val mergedObj = JSONObject()
            mergedObj.put("fromRow", range.firstRow)
            mergedObj.put("toRow", range.lastRow)
            mergedObj.put("fromCol", range.firstColumn)
            mergedObj.put("toCol", range.lastColumn)
            mergedArray.put(mergedObj)
        }
        root.put("merged", mergedArray)

        // 2. Ширины колонок
        val colWidthsObj = JSONObject()
        for (c in 0..30) {
            val w = sheet.getColumnWidth(c)
            if (w > 0) {
                colWidthsObj.put(c.toString(), (w / 256.0 * 8.0).toInt())
            }
        }
        root.put("colWidths", colWidthsObj)

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

                    // Безопасное получение фонового цвета ячейки (без прямых зависимостей от HSSFColor)
                    val bgHex = getSafeBackgroundColor(cell)
                    if (!bgHex.isNullOrEmpty()) {
                        cellObj.put("bg", bgHex)
                    }

                    // Границы ячеек
                    val style = cell.cellStyle
                    if (style != null) {
                        if (style.borderTop.toInt() != 0) cellObj.put("bt", 1)
                        if (style.borderBottom.toInt() != 0) cellObj.put("bb", 1)
                        if (style.borderLeft.toInt() != 0) cellObj.put("bl", 1)
                        if (style.borderRight.toInt() != 0) cellObj.put("br", 1)
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
            val style = cell.cellStyle ?: return null
            val color = style.fillForegroundColorColor ?: return null

            // 1. Попытка получить RGB массив через рефлексию (работает для HSSFColor и XSSFColor)
            try {
                val getTripletMethod = color.javaClass.getMethod("getTriplet")
                val triplet = getTripletMethod.invoke(color) as? ShortArray
                if (triplet != null && triplet.size == 3) {
                    val r = triplet[0].toInt()
                    val g = triplet[1].toInt()
                    val b = triplet[2].toInt()
                    // Игнорируем стандартный чистый белый цвет по умолчанию
                    if (!(r == 255 && g == 255 && b == 255)) {
                        return String.format("#%02X%02X%02X", r, g, b)
                    }
                }
            } catch (_: Exception) { }

            // 2. Попытка получить RGB через getRgb() (для XSSFColor)
            try {
                val getRgbMethod = color.javaClass.getMethod("getRgb")
                val rgb = getRgbMethod.invoke(color) as? ByteArray
                if (rgb != null && rgb.size >= 3) {
                    val r = rgb[rgb.size - 3].toInt() and 0xFF
                    val g = rgb[rgb.size - 2].toInt() and 0xFF
                    val b = rgb[rgb.size - 1].toInt() and 0xFF
                    if (!(r == 255 && g == 255 && b == 255)) {
                        return String.format("#%02X%02X%02X", r, g, b)
                    }
                }
            } catch (_: Exception) { }

            // 3. Fallback: разбор строкового представления
            val hex = color.toString()
            if (hex.length >= 6) {
                val cleanHex = hex.takeLast(6)
                if (cleanHex != "FFFFFF") {
                    return "#$cleanHex"
                }
            }
            null
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

                    // Перезапись файла через ContentResolver (без "rwt" для предотвращения коррупции файлов .xls)
                    contentResolver.openOutputStream(uri, "wt")?.use { os ->
                        wb.write(os)
                        os.flush()
                    }

                    withContext(Dispatchers.Main) {
                        Toast.makeText(this@MainActivity, "Сохранено успешно!", Toast.LENGTH_SHORT).show()
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
