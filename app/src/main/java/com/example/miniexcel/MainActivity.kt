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

    private fun createWorkingCopyAndParse(uri: Uri) {
        lifecycleScope.launch(Dispatchers.IO) {
            try {
                // Создаем постоянный файл сессии с фиксированным именем, чтобы избежать плождения файлов
                val cacheFile = File(cacheDir, "working_session.tmp")
                contentResolver.openInputStream(uri)?.use { inputStream ->
                    FileOutputStream(cacheFile).use { outputStream ->
                        inputStream.copyTo(outputStream)
                    }
                }
                workingFile = cacheFile

                val payload = parseExcelFile(cacheFile)
                cachedJsonPayload = payload

                withContext(Dispatchers.Main) {
                    tableWebView.loadUrl("file:///android_asset/grid.html")
                }
            } catch (e: Exception) {
                Log.e("MiniExcelDebug", "Ошибка открытия файла: ${e.message}", e)
                withContext(Dispatchers.Main) {
                    Toast.makeText(this@MainActivity, "Ошибка открытия файла: ${e.localizedMessage}", Toast.LENGTH_LONG).show()
                }
            }
        }
    }

    private fun parseExcelFile(file: File): String {
        WorkbookFactory.create(file).use { workbook ->
            val sheet = workbook.getSheetAt(0) ?: return emptyPayload
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

            lifecycleScope.launch(Dispatchers.IO) {
                try {
                    val root = JSONObject(jsonPayload)
                    val matrix = root.optJSONArray("matrix") ?: return@launch
                    val isFinalSave = root.optBoolean("isFinalSave", false)

                    // 1. Записываем изменения в рабочую копию
                    WorkbookFactory.create(fileToProcess).use { workbook ->
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

                    // 2. Если это финальное сохранение — копируем файл в целевой URI
                    if (isFinalSave) {
                        contentResolver.openOutputStream(targetUri, "wt")?.use { outputStream ->
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
                        Toast.makeText(this@MainActivity, "Ошибка сохранения: ${e.localizedMessage}", Toast.LENGTH_LONG).show()
                    }
                }
            }
        }
    }
}
