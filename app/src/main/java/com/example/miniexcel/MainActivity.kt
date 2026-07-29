package com.example.miniexcel

import android.annotation.SuppressLint
import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.util.Log
import android.view.View
import android.webkit.ConsoleMessage
import android.webkit.JavascriptInterface
import android.webkit.WebChromeClient
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.Button
import android.widget.ProgressBar
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import com.example.miniexcel.R
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.poi.openxml4j.util.ZipSecureFile
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.json.JSONArray
import org.json.JSONObject
import java.io.InputStream

class MainActivity : AppCompatActivity() {

    private lateinit var webView: WebView
    private lateinit var progressBar: ProgressBar
    private lateinit var btnOpen: Button
    private lateinit var btnSave: Button

    @Volatile
    private var cachedJsonPayload: String = "{\"matrix\":[],\"widths\":[],\"heights\":[],\"merges\":[]}"

    // Лаунчер для открытия файла с устройства
    private val openFileLauncher = registerForActivityResult(
        ActivityResultContracts.StartActivityForResult()
    ) { result ->
        if (result.resultCode == RESULT_OK) {
            result.data?.data?.let { uri ->
                loadExcelFromUri(uri)
            }
        }
    }

    // Лаунчер для сохранения/экспорта файла
    private val saveFileLauncher = registerForActivityResult(
        ActivityResultContracts.StartActivityForResult()
    ) { result ->
        if (result.resultCode == RESULT_OK) {
            result.data?.data?.let { uri ->
                Toast.makeText(this, "Файл успешно сохранён", Toast.LENGTH_SHORT).show()
            }
        }
    }

    @SuppressLint("SetJavaScriptEnabled", "JavascriptInterface")
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        webView = findViewById(R.id.webView)
        progressBar = findViewById(R.id.progressBar)
        btnOpen = findViewById(R.id.btnOpen)
        btnSave = findViewById(R.id.btnSave)

        setupButtons()
        setupWebView()

        val fileUri: Uri? = intent.data
        if (fileUri != null) {
            loadExcelFromUri(fileUri)
        } else {
            webView.loadUrl("file:///android_asset/grid.html")
        }
    }

    private fun setupButtons() {
        btnOpen.setOnClickListener {
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

        btnSave.setOnClickListener {
            val intent = Intent(Intent.ACTION_CREATE_DOCUMENT).apply {
                addCategory(Intent.CATEGORY_OPENABLE)
                type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                putExtra(Intent.EXTRA_TITLE, "export.xlsx")
            }
            saveFileLauncher.launch(intent)
        }
    }

    @SuppressLint("SetJavaScriptEnabled")
    private fun setupWebView() {
        webView.settings.apply {
            javaScriptEnabled = true
            domStorageEnabled = true
            useWideViewPort = true
            loadWithOverviewMode = true
            allowFileAccess = true
            allowContentAccess = true
        }

        webView.addJavascriptInterface(AndroidBridge(), "AndroidBridge")

        webView.webChromeClient = object : WebChromeClient() {
            override fun onConsoleMessage(consoleMessage: ConsoleMessage?): Boolean {
                Log.d("WebConsole", "${consoleMessage?.message()} -- Line ${consoleMessage?.lineNumber()} of ${consoleMessage?.sourceId()}")
                return true
            }
        }

        webView.webViewClient = object : WebViewClient() {
            override fun onPageFinished(view: WebView?, url: String?) {
                super.onPageFinished(view, url)
                webView.evaluateJavascript("if(typeof onDataReady === 'function'){ onDataReady(); }", null)
            }
        }
    }

    private fun loadExcelFromUri(uri: Uri) {
        progressBar.visibility = View.VISIBLE

        lifecycleScope.launch(Dispatchers.IO) {
            try {
                ZipSecureFile.setMinInflateRatio(0.0001)

                val inputStream: InputStream? = contentResolver.openInputStream(uri)
                if (inputStream != null) {
                    val workbook = WorkbookFactory.create(inputStream)
                    val jsonResult = parseWorkbookToJson(workbook)
                    workbook.close()
                    inputStream.close()

                    cachedJsonPayload = jsonResult

                    withContext(Dispatchers.Main) {
                        progressBar.visibility = View.GONE
                        webView.loadUrl("file:///android_asset/grid.html")
                    }
                } else {
                    showError("Не удалось открыть файл")
                }
            } catch (e: Exception) {
                Log.e("MainActivity", "Ошибка при обработке Excel: ${e.message}", e)
                showError("Ошибка чтения файла: ${e.localizedMessage}")
            } finally {
                withContext(Dispatchers.Main) {
                    progressBar.visibility = View.GONE
                }
                System.gc()
            }
        }
    }

    private fun parseWorkbookToJson(workbook: Workbook): String {
        val sheet = workbook.getSheetAt(0) ?: return cachedJsonPayload

        val lastRowNum = sheet.lastRowNum
        val matrixArray = JSONArray()

        var maxColsFound = 0
        val maxRowsToRead = Math.min(lastRowNum + 1, 3000)

        for (r in 0 until maxRowsToRead) {
            val row = sheet.getRow(r)
            if (row != null) {
                val lastCellNum = row.lastCellNum.toInt()
                if (lastCellNum > maxColsFound) {
                    maxColsFound = Math.min(lastCellNum, 100)
                }
            }
        }

        for (r in 0 until maxRowsToRead) {
            val row = sheet.getRow(r)
            val rowArray = JSONArray()

            for (c in 0 until maxColsFound) {
                val cell = row?.getCell(c)
                val cellObj = JSONObject()
                cellObj.put("v", getCellValueAsString(cell))
                rowArray.put(cellObj)
            }
            matrixArray.put(rowArray)
        }

        val widthsArray = JSONArray()
        for (c in 0 until maxColsFound) {
            val colWidth = sheet.getColumnWidth(c)
            val pxWidth = Math.max(60, Math.min((colWidth / 256) * 8, 300))
            widthsArray.put(pxWidth)
        }

        val heightsArray = JSONArray()
        for (r in 0 until maxRowsToRead) {
            val row = sheet.getRow(r)
            val rowHeight = row?.heightInPoints?.toInt() ?: 20
            heightsArray.put(Math.max(18, Math.min(rowHeight, 100)))
        }

        val mergesArray = JSONArray()
        val numRegions = sheet.numMergedRegions
        for (i in 0 until numRegions) {
            val region = sheet.getMergedRegion(i)
            if (region.firstRow < maxRowsToRead) {
                val mergeObj = JSONObject().apply {
                    put("sr", region.firstRow)
                    put("sc", region.firstColumn)
                    put("er", Math.min(region.lastRow, maxRowsToRead - 1))
                    put("ec", Math.min(region.lastColumn, maxColsFound - 1))
                }
                mergesArray.put(mergeObj)
            }
        }

        return JSONObject().apply {
            put("matrix", matrixArray)
            put("widths", widthsArray)
            put("heights", heightsArray)
            put("merges", mergesArray)
        }.toString()
    }

    private fun getCellValueAsString(cell: org.apache.poi.ss.usermodel.Cell?): String {
        if (cell == null) return ""
        return try {
            when (cell.cellType) {
                org.apache.poi.ss.usermodel.CellType.STRING -> cell.stringCellValue ?: ""
                org.apache.poi.ss.usermodel.CellType.NUMERIC -> {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        cell.dateCellValue?.toString() ?: ""
                    } else {
                        val num = cell.numericCellValue
                        if (num == Math.floor(num)) {
                            num.toLong().toString()
                        } else {
                            num.toString()
                        }
                    }
                }
                org.apache.poi.ss.usermodel.CellType.BOOLEAN -> cell.booleanCellValue.toString()
                org.apache.poi.ss.usermodel.CellType.FORMULA -> {
                    try {
                        cell.stringCellValue
                    } catch (e: Exception) {
                        try {
                            cell.numericCellValue.toString()
                        } catch (e2: Exception) {
                            ""
                        }
                    }
                }
                else -> ""
            }
        } catch (e: Exception) {
            ""
        }
    }

    private suspend fun showError(message: String) {
        withContext(Dispatchers.Main) {
            progressBar.visibility = View.GONE
            Toast.makeText(this@MainActivity, message, Toast.LENGTH_LONG).show()
        }
    }

    inner class AndroidBridge {
        @JavascriptInterface
        fun getExcelData(): String {
            return cachedJsonPayload
        }
    }
}
