package com.example.miniexcel

import android.os.Build
import android.app.Activity
import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.util.Log
import android.webkit.ConsoleMessage
import android.webkit.JavascriptInterface
import android.webkit.WebChromeClient
import android.webkit.WebSettings
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.Button
import android.widget.Toast
import androidx.activity.result.ActivityResultLauncher
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.poi.openxml4j.util.ZipSecureFile
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.json.JSONArray
import org.json.JSONObject
import java.io.File
import java.io.FileOutputStream

class MainActivity : AppCompatActivity() {

    private lateinit var tableWebView: WebView
    private lateinit var openButton: Button
    private lateinit var saveButton: Button
    
    private var currentFileUri: Uri? = null
    
private val emptyPayload: String
    get() {
        val rowsCount = 30
        val colsCount = 15
        
        val matrix = JSONArray()
        for (r in 0 until rowsCount) {
            val rowArray = JSONArray()
            for (c in 0 until colsCount) {
                val cellObj = JSONObject()
                cellObj.put("v", "")
                rowArray.put(cellObj)
            }
            matrix.put(rowArray)
        }
        
        val widths = JSONArray()
        for (c in 0 until colsCount) {
            widths.put(80) // Дефолтная ширина колонки в пикселях
        }
        
        val heights = JSONArray()
        for (r in 0 until rowsCount) {
            heights.put(25) // Дефолтная высота строки в пикселях
        }

        val root = JSONObject().apply {
            put("matrix", matrix)
            put("widths", widths)
            put("heights", heights)
            put("merges", JSONArray())
        }
        return root.toString()
    }

@Volatile
private var cachedJsonPayload: String = emptyPayload

    private lateinit var openFileLauncher: ActivityResultLauncher<Intent>
    private lateinit var saveFileLauncher: ActivityResultLauncher<Intent>

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        ZipSecureFile.setMinInflateRatio(0.005)

        openButton = findViewById(R.id.openButton)
        saveButton = findViewById(R.id.saveButton)
        tableWebView = findViewById(R.id.tableWebView)

        setupWebView()
        initFileLaunchers()

        openButton.setOnClickListener {
            val intent = Intent(Intent.ACTION_OPEN_DOCUMENT).apply {
                addCategory(Intent.CATEGORY_OPENABLE)
                type = "*/*"
                putExtra(Intent.EXTRA_MIME_TYPES, arrayOf(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "application/vnd.ms-excel"
                ))
            }
            openFileLauncher.launch(intent)
        }

        saveButton.setOnClickListener {
            if (currentFileUri == null) {
                // Если файла еще нет, сначала просим пользователя выбрать, куда его сохранить
                val intent = Intent(Intent.ACTION_CREATE_DOCUMENT).apply {
                    addCategory(Intent.CATEGORY_OPENABLE)
                    type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    putExtra(Intent.EXTRA_TITLE, "edited_sheet.xlsx")
                }
                saveFileLauncher.launch(intent)
            } else {
                // Если файл уже открыт, запрашиваем экспорт данных из WebView
                triggerJSExport()
            }
        }

        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.KITKAT) {
            WebView.setWebContentsDebuggingEnabled(true)
        }
    }

    private fun triggerJSExport() {
        tableWebView.post { 
            tableWebView.evaluateJavascript("exportExcelToAndroid();", null) 
        }
    }

    private fun setupWebView() {
        tableWebView.apply {
            setLayerType(android.view.View.LAYER_TYPE_SOFTWARE, null)
            clearCache(true)
            clearHistory()
            clearFormData()
            
            settings.apply {
                javaScriptEnabled = true
                cacheMode = WebSettings.LOAD_NO_CACHE
                textZoom = 100
                useWideViewPort = true
                loadWithOverviewMode = true
                setSupportZoom(false)
                builtInZoomControls = false
                domStorageEnabled = true
                allowFileAccess = true
            }
            
            addJavascriptInterface(AndroidBridge(), "AndroidBridge")
            webViewClient = WebViewClient()
            webChromeClient = object : WebChromeClient() {
                override fun onConsoleMessage(consoleMessage: ConsoleMessage): Boolean {
                    Log.d("WebViewJS", "${consoleMessage.message()} -- Line ${consoleMessage.lineNumber()} of ${consoleMessage.sourceId()}")
                    return true
                }
            }
            loadUrl("file:///android_asset/grid.html")
        }
    }

    private fun initFileLaunchers() {
        openFileLauncher = registerForActivityResult(ActivityResultContracts.StartActivityForResult()) { result ->
            if (result.resultCode == Activity.RESULT_OK) {
                result.data?.data?.let { uri ->
                    currentFileUri = uri
                    pipeExcelToWebView(uri)
                }
            }
        }

        saveFileLauncher = registerForActivityResult(ActivityResultContracts.StartActivityForResult()) { result ->
            if (result.resultCode == Activity.RESULT_OK) {
                result.data?.data?.let { uri ->
                    currentFileUri = uri
                    // Теперь, когда URI создан, запрашиваем экспорт у JS
                    triggerJSExport()
                }
            }
        }
    }

    private fun pipeExcelToWebView(fileUri: Uri) {
        cachedJsonPayload = emptyPayload
        tableWebView.loadUrl("file:///android_asset/grid.html")

        lifecycleScope.launch {
            var tempFile: File? = null
            try {
                withContext(Dispatchers.IO) {
                    tempFile = File(cacheDir, "current_scanned_sheet.tmp")
                    
                    contentResolver.openInputStream(fileUri)?.use { input ->
                        FileOutputStream(tempFile).use { output ->
                            input.copyTo(output, 16384)
                        }
                    }

                    WorkbookFactory.create(tempFile, null, true).use { workbook ->
                        if (workbook.numberOfSheets > 0) {
                            parseSheetToJson(workbook.getSheetAt(0))
                        } else {
                            withContext(Dispatchers.Main) {
                                Toast.makeText(this@MainActivity, "Файл пуст", Toast.LENGTH_SHORT).show()
                            }
                        }
                    }
                }
            } catch (e: Exception) {
                Log.e("MiniExcelDebug", "Ошибка обработки: ${e.message}")
                Toast.makeText(this@MainActivity, "Ошибка чтения тяжелого файла", Toast.LENGTH_SHORT).show()
            } finally {
                tempFile?.let { if (it.exists()) it.delete() }
                System.gc()
            }
        }
    }

    private fun parseSheetToJson(sheet: Sheet) {
        val jsonTable = JSONArray()
        val jsonWidths = JSONArray()
        val jsonHeights = JSONArray()
        val jsonMerges = JSONArray()

        var lastRowIdx = sheet.lastRowNum
        var maxColsCount = 0

        for (r in 0..lastRowIdx) {
            val row = sheet.getRow(r)
            if (row != null && row.lastCellNum > maxColsCount) {
                maxColsCount = row.lastCellNum.toInt()
            }
        }
        if (maxColsCount == 0) maxColsCount = 12

        if (lastRowIdx > 1500) lastRowIdx = 1500
        if (maxColsCount > 60) maxColsCount = 60

        try {
            for (c in 0 until maxColsCount) {
                val w = sheet.getColumnWidth(c) / 35
                jsonWidths.put(if (w > 0) w else 64)
            }

            for (r in 0..lastRowIdx) {
                val row = sheet.getRow(r)
                val h = if (row != null) (row.heightInPoints * 1.33).toInt() else 20
                jsonHeights.put(if (h > 0) h else 20)

                val rowArray = JSONArray()
                for (c in 0 until maxColsCount) {
                    val cellObj = JSONObject()
                    cellObj.put("v", "")

                    if (row != null) {
                        val cell = row.getCell(c)
                        if (cell != null) {
                            when (cell.cellType) {
                                org.apache.poi.ss.usermodel.CellType.STRING -> 
                                    cellObj.put("v", cell.stringCellValue)
                                org.apache.poi.ss.usermodel.CellType.NUMERIC -> {
                                    if (DateUtil.isCellDateFormatted(cell)) {
                                        cellObj.put("v", cell.dateCellValue.toString())
                                    } else {
                                        cellObj.put("v", cell.numericCellValue)
                                    }
                                }
                                org.apache.poi.ss.usermodel.CellType.FORMULA -> {
                                    try { cellObj.put("v", cell.stringCellValue) } 
                                    catch (e1: Exception) {
                                        try { cellObj.put("v", cell.numericCellValue) } 
                                        catch (ignored: Exception) {}
                                    }
                                }
                                org.apache.poi.ss.usermodel.CellType.BOOLEAN -> 
                                    cellObj.put("v", cell.booleanCellValue)
                                else -> cellObj.put("v", "")
                            }
                        }
                    }
                    rowArray.put(cellObj)
                }
                jsonTable.put(rowArray)
            }

            val numRegions = sheet.numMergedRegions
            for (i in 0 until numRegions) {
                val region = sheet.getMergedRegion(i)
                if (region != null) {
                    if (region.firstRow > lastRowIdx || region.firstColumn >= maxColsCount) continue
                    val mergeObj = JSONObject().apply {
                        put("sr", region.firstRow)
                        put("er", region.lastRow)
                        put("sc", region.firstColumn)
                        put("ec", region.lastColumn)
                    }
                    jsonMerges.put(mergeObj)
                }
            }

            val rootPayload = JSONObject().apply {
                put("matrix", jsonTable)
                put("widths", jsonWidths)
                put("heights", jsonHeights)
                put("merges", jsonMerges)
            }

            cachedJsonPayload = rootPayload.toString()

            tableWebView.post {
                tableWebView.evaluateJavascript("setTimeout(function() { requestDataFromAndroid(); }, 150);", null)
            }

        } catch (e: Exception) {
            Log.e("MiniExcelDebug", "Ошибка сборки JSON: ${e.message}")
        }
    }

    // Сохранение переданных данных из JS обратно в физический XLSX файл на устройстве
    private fun saveJsonToExcelFile(jsonData: String, fileUri: Uri) {
        lifecycleScope.launch {
            var tempFile: File? = null
            try {
                withContext(Dispatchers.IO) {
                    val root = JSONObject(jsonData)
                    val matrix = root.optJSONArray("matrix") ?: JSONArray()
                    
                    val workbook = XSSFWorkbook()
                    val sheet = workbook.createSheet("Sheet1")

                    // Запись структуры матрицы обратно в Excel-ячейки
                    for (r in 0 until matrix.length()) {
                        val rowArray = matrix.optJSONArray(r) ?: continue
                        val row = sheet.createRow(r)
                        for (c in 0 until rowArray.length()) {
                            val cellObj = rowArray.optJSONObject(c) ?: continue
                            val cellValue = cellObj.opt("v") ?: ""
                            val cell = row.createCell(c)
                            
                            // Автоопределение базовых типов для записи
                            when (cellValue) {
                                is Number -> cell.setCellValue(cellValue.toDouble())
                                is Boolean -> cell.setCellValue(cellValue)
                                else -> cell.setCellValue(cellValue.toString())
                            }
                        }
                    }

                    // Сохраняем сначала во временный файл во избежание повреждения данных при сбоях записи
                    tempFile = File(cacheDir, "temp_saving_sheet.xlsx")
                    FileOutputStream(tempFile).use { out ->
                        workbook.write(out)
                    }
                    workbook.close()

                    // Копируем временный файл в целевой Uri через ContentResolver
                    contentResolver.openOutputStream(fileUri, "rwt")?.use { outStream ->
                        tempFile?.inputStream()?.use { inStream ->
                            inStream.copyTo(outStream)
                        }
                    }
                }
                Toast.makeText(this@MainActivity, "Файл успешно сохранен", Toast.LENGTH_SHORT).show()
            } catch (e: Exception) {
                Log.e("MiniExcelDebug", "Ошибка записи: ${e.message}")
                Toast.makeText(this@MainActivity, "Не удалось сохранить файл", Toast.LENGTH_SHORT).show()
            } finally {
                tempFile?.let { if (it.exists()) it.delete() }
                System.gc()
            }
        }
    }

    inner class AndroidBridge {
        @JavascriptInterface
        fun getExcelData(): String = cachedJsonPayload

        @JavascriptInterface
        fun onStatus(message: String) {
            Log.d("MiniExcelStatus", message)
        }

        // JS внутри grid.html должен вызвать этот метод в конце функции exportExcelToAndroid()
        // Пример вызова из JS: window.AndroidBridge.saveExcelData(JSON.stringify(excelData));
        @JavascriptInterface
        fun saveExcelData(jsonData: String) {
            currentFileUri?.let { uri ->
                saveJsonToExcelFile(jsonData, uri)
            } ?: run {
                Log.e("MiniExcelDebug", "Попытка сохранить, но currentFileUri равен null")
            }
        }
    }
}
