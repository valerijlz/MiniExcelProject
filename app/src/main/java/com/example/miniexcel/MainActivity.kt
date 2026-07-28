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

    inner class AndroidBridge {
        @JavascriptInterface
        fun getExcelData(): String {
            return cachedJsonPayload
        }

        @JavascriptInterface
        fun saveExcelData(jsonPayload: String) {
            cachedJsonPayload = jsonPayload
            val targetUri = currentFileUri ?: return

            val fileToProcess = workingFile ?: return
            try {
                val root = JSONObject(jsonPayload)
                val matrix = root.optJSONArray("matrix") ?: return
                val isFinalSave = root.optBoolean("isFinalSave", false)

                // 1. Записываем изменения в рабочую копию (.tmp)
                WorkbookFactory.create(fileToProcess, null, false).use { workbook ->
                    val sheet = workbook.getSheetAt(0) ?: workbook.createSheet("Sheet1")

                    for (r in 0 until matrix.length()) {
                        val rowArray = matrix.optJSONArray(r) ?: continue
                        var row = sheet.getRow(r)
                        if (row == null) {
                            row = sheet.createRow(r)
                        }

                        for (c in 0 until rowArray.length()) {
                            val cellObj = rowArray.optJSONObject(c)
                            val cellVal = cellObj?.optString("v", "") ?: ""
                            var cell = row.getCell(c)
                            if (cell == null) {
                                cell = row.createCell(c)
                            }

                            if (cellVal.toDoubleOrNull() != null) {
                                cell.setCellValue(cellVal.toDouble())
                            } else {
                                cell.setCellValue(cellVal)
                            }
                        }
                    }

                    FileOutputStream(fileToProcess).use { fos ->
                        workbook.write(fos)
                        fos.flush()
                    }
                }

                // 2. Если это финальное сохранение — копируем tmp-файл в целевой документ пользователя
                if (isFinalSave) {
                    contentResolver.openOutputStream(targetUri, "w")?.use { outputStream ->
                        fileToProcess.inputStream().use { inputStream ->
                            inputStream.copyTo(outputStream)
                            outputStream.flush()
                        }
                    }
                    lifecycleScope.launch(Dispatchers.Main) {
                        Toast.makeText(this@MainActivity, "Файл успешно сохранен", Toast.LENGTH_SHORT).show()
                    }
                }
            } catch (e: Exception) {
                Log.e("MiniExcelDebug", "Ошибка сохранения: ${e.message}", e)
                lifecycleScope.launch(Dispatchers.Main) {
                    Toast.makeText(this@MainActivity, "Ошибка сохранения файла", Toast.LENGTH_SHORT).show()
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
                                        catch (e2: Exception) { cellObj.put("v", "[formula]") }
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

            for (i in 0 until sheet.numMergedRegions) {
                val region = sheet.getMergedRegion(i)
                val mObj = JSONObject().apply {
                    put("sr", region.firstRow)
                    put("sc", region.firstColumn)
                    put("er", region.lastRow)
                    put("ec", region.lastColumn)
                }
                jsonMerges.put(mObj)
            }

            val root = JSONObject().apply {
                put("matrix", jsonTable)
                put("widths", jsonWidths)
                put("heights", jsonHeights)
                put("merges", jsonMerges)
            }

            cachedJsonPayload = root.toString()

            runOnUiThread {
                tableWebView.evaluateJavascript("requestDataFromAndroid();", null)
            }

        } catch (e: Exception) {
            Log.e("MiniExcelDebug", "Ошибка парсинга листа: ${e.message}")
        }
    }
}
