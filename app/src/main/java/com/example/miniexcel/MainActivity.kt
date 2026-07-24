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
import org.json.JSONArray
import org.json.JSONObject
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream

class MainActivity : AppCompatActivity() {

    private lateinit var tableWebView: WebView
    private lateinit var openButton: Button
    private lateinit var saveButton: Button
    
    private var currentFileUri: Uri? = null
    private var workingFile: File? = null
    
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
            for (c in 0 until colsCount) widths.put(80)
            
            val heights = JSONArray()
            for (r in 0 until rowsCount) heights.put(25)

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
                val intent = Intent(Intent.ACTION_CREATE_DOCUMENT).apply {
                    addCategory(Intent.CATEGORY_OPENABLE)
                    type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    putExtra(Intent.EXTRA_TITLE, "edited_sheet.xlsx")
                }
                saveFileLauncher.launch(intent)
            } else {
                triggerJSExportAndSave()
            }
        }

        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.KITKAT) {
            WebView.setWebContentsDebuggingEnabled(true)
        }
    }

    private fun triggerJSExportAndSave() {
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
                    createWorkingCopyAndParse(uri)
                }
            }
        }

        saveFileLauncher = registerForActivityResult(ActivityResultContracts.StartActivityForResult()) { result ->
            if (result.resultCode == Activity.RESULT_OK) {
                result.data?.data?.let { uri ->
                    currentFileUri = uri
                    triggerJSExportAndSave()
                }
            }
        }
    }

    private fun createWorkingCopyAndParse(fileUri: Uri) {
        cachedJsonPayload = emptyPayload
        tableWebView.loadUrl("file:///android_asset/grid.html")

        lifecycleScope.launch {
            try {
                withContext(Dispatchers.IO) {
                    workingFile?.let { if (it.exists()) it.delete() }
                    workingFile = File(cacheDir, "working_session_${System.currentTimeMillis()}.tmp")

                    contentResolver.openInputStream(fileUri)?.use { input ->
                        FileOutputStream(workingFile).use { output ->
                            input.copyTo(output, 16384)
                        }
                    }

                    WorkbookFactory.create(workingFile, null, true).use { workbook ->
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
                Log.e("MiniExcelDebug", "Ошибка создания рабочей копии: ${e.message}")
                Toast.makeText(this@MainActivity, "Ошибка чтения файла", Toast.LENGTH_SHORT).show()
            } finally {
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
        if (lastRowIdx <= 0) lastRowIdx = 29 
        if (maxColsCount <= 0) maxColsCount = 15 

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

    // Сохранение с сохранением исходного стиля (Apache POI стили не затираются)
private fun commitChangesAndExportToOriginal(jsonData: String) {
        val currentWorkingFile = workingFile ?: return

        lifecycleScope.launch {
            try {
                withContext(Dispatchers.IO) {
                    val root = JSONObject(jsonData)
                    val matrix = root.optJSONArray("matrix") ?: JSONArray()
                    val isFinalSave = root.optBoolean("isFinalSave", false)

                    // 1. Открываем рабочую копию через POI для сохранения форматирования
                    val fileInputStream = FileInputStream(currentWorkingFile)
                    val workbook = WorkbookFactory.create(fileInputStream)
                    fileInputStream.close()

                    val sheet = workbook.getSheetAt(0)

                    for (r in 0 until matrix.length()) {
                        val rowArray = matrix.optJSONArray(r) ?: continue
                        var poiRow = sheet.getRow(r)
                        if (poiRow == null) {
                            poiRow = sheet.createRow(r)
                        }

                        for (c in 0 until rowArray.length()) {
                            val cellObj = rowArray.optJSONObject(c) ?: continue
                            val cellValue = cellObj.opt("v") ?: ""

                            var poiCell = poiRow.getCell(c)
                            if (poiCell == null) {
                                poiCell = poiRow.createCell(c)
                            }

                            when (cellValue) {
                                is Number -> poiCell.setCellValue(cellValue.toDouble())
                                is Boolean -> poiCell.setCellValue(cellValue)
                                else -> poiCell.setCellValue(cellValue.toString())
                            }
                        }
                    }

                    // 2. Перезаписываем временный рабочий файл
                    FileOutputStream(currentWorkingFile).use { out ->
                        workbook.write(out)
                    }
                    workbook.close()

                    // 3. Если это финальное сохранение (нажата кнопка) — выгружаем в итоговый Uri пользователя
                    if (isFinalSave && currentFileUri != null) {
                        contentResolver.openOutputStream(currentFileUri!!, "w")?.use { outStream ->
                            currentWorkingFile.inputStream().use { inStream ->
                                inStream.copyTo(outStream)
                            }
                        }
                    }
                }
                
                // Тост показываем только при реальном клике на кнопку «Сохранить»
                val rootCheck = JSONObject(jsonData)
                if (rootCheck.optBoolean("isFinalSave", false)) {
                    withContext(Dispatchers.Main) {
                        Toast.makeText(this@MainActivity, "Файл успешно сохранен", Toast.LENGTH_SHORT).show()
                    }
                }
            } catch (e: Exception) {
                Log.e("MiniExcelDebug", "Ошибка записи: ${e.message}", e)
                val rootCheck = JSONObject(jsonData)
                if (rootCheck.optBoolean("isFinalSave", false)) {
                    withContext(Dispatchers.Main) {
                        Toast.makeText(this@MainActivity, "Ошибка сохранения: ${e.localizedMessage}", Toast.LENGTH_LONG).show()
                    }
                }
            }
        }
    }

    override fun onDestroy() {
        super.onDestroy()
        try {
            workingFile?.let { if (it.exists()) it.delete() }
        } catch (e: Exception) {
            Log.e("MiniExcelDebug", "Ошибка очистки временного файла: ${e.message}")
        }
    }

    inner class AndroidBridge {
        @JavascriptInterface
        fun getExcelData(): String = cachedJsonPayload

        @JavascriptInterface
        fun onStatus(message: String) {
            Log.d("MiniExcelStatus", message)
        }

        @JavascriptInterface
        fun saveExcelData(jsonData: String) {
            commitChangesAndExportToOriginal(jsonData)
        }
    }
}
