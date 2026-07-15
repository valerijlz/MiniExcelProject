package com.example.miniexcel

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
    
    private val emptyPayload = "{\"matrix\":[],\"merges\":[],\"widths\":[],\"heights\":[]}"
    @Volatile
    private var cachedJsonPayload: String = emptyPayload

    private lateinit var openFileLauncher: ActivityResultLauncher<Intent>
    private lateinit var saveFileLauncher: ActivityResultLauncher<Intent>

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        // Защита Apache POI от перегрузки памяти
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
            tableWebView.post { 
                tableWebView.evaluateJavascript("exportExcelToAndroid();", null) 
            }
        }
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.KITKAT) {
    WebView.setWebContentsDebuggingEnabled(true);
}
    }

    private fun setupWebView() {
        tableWebView.apply {
            // ИСПРАВЛЕНИЕ: Отключаем аппаратные сбои Canvas в новом SDK Android
            // Переключаем WebView в режим программного рендеринга (Software Layer)
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
                    tableWebView.post { 
                        tableWebView.evaluateJavascript("exportExcelToAndroid();", null) 
                    }
                }
            }
        }
    }

    private fun pipeExcelToWebView(fileUri: Uri) {
        // Мгновенный сброс старой среды
        cachedJsonPayload = emptyPayload
        tableWebView.loadUrl("file:///android_asset/grid.html")

        // Запускаем современную безопасную корутину в контексте Жизненного Цикла Activity
        lifecycleScope.launch {
            var tempFile: File? = null
            try {
                // Переключаемся на фоновый IO-поток для работы с диском и POI
                withContext(Dispatchers.IO) {
                    tempFile = File(cacheDir, "current_scanned_sheet.tmp")
                    
                    // Быстрое потоковое копирование во временный файл
                    contentResolver.openInputStream(fileUri)?.use { input ->
                        FileOutputStream(tempFile).use { output ->
                            input.copyTo(output, 16384)
                        }
                    }

                    // Открытие книги из файла в режиме Read-Only
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
                // Чистим кэш-память накопителя и вызываем сборщик мусора
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

        // Определяем реальные границы
        for (r in 0..lastRowIdx) {
            val row = sheet.getRow(r)
            if (row != null && row.lastCellNum > maxColsCount) {
                maxColsCount = row.lastCellNum.toInt()
            }
        }
        if (maxColsCount == 0) maxColsCount = 12

        // Безопасные лимиты
        if (lastRowIdx > 1500) lastRowIdx = 1500
        if (maxColsCount > 60) maxColsCount = 60

        try {
            // Заполняем ширину колонок
            for (c in 0 until maxColsCount) {
                val w = sheet.getColumnWidth(c) / 35
                jsonWidths.put(if (w > 0) w else 64)
            }

            // РЕШЕНИЕ ПРОБЛЕМЫ С РАМКАМИ: гарантируем генерацию ячеек, даже если они null (пустые в POI)
            for (r in 0..lastRowIdx) {
                val row = sheet.getRow(r)
                val h = if (row != null) (row.heightInPoints * 1.33).toInt() else 20
                jsonHeights.put(if (h > 0) h else 20)

                val rowArray = JSONArray()
                for (c in 0 until maxColsCount) {
                    val cellObj = JSONObject()
                    cellObj.put("v", "") // Каждая ячейка теперь ОБЯЗАТЕЛЬНО получает базовое значение и структуру

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

            // Обработка объединенных регионов
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

            // Передаем данные в WebView, гарантируя выполнение на главном потоке
            tableWebView.post {
                tableWebView.evaluateJavascript("setTimeout(function() { requestDataFromAndroid(); }, 150);", null)
            }

        } catch (e: Exception) {
            Log.e("MiniExcelDebug", "Ошибка сборки JSON: ${e.message}")
        }
    }

    inner class AndroidBridge {
        @JavascriptInterface
        fun getExcelData(): String = cachedJsonPayload

        @JavascriptInterface
        fun onStatus(message: String) {
            Log.d("MiniExcelStatus", message)
        }
    }
}
