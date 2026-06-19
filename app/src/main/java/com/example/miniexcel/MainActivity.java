package com.example.miniexcel;

import android.app.Activity;
import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.os.Bundle;
import android.os.ParcelFileDescriptor;
import android.provider.OpenableColumns;
import android.view.Gravity;
import android.widget.Button;
import android.widget.EditText;
import android.widget.LinearLayout;
import android.widget.TextView;
import android.widget.Toast;
import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.appcompat.app.AppCompatActivity;
import androidx.recyclerview.widget.LinearLayoutManager;
import androidx.recyclerview.widget.RecyclerView;
import com.example.miniexcel.R;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class MainActivity extends AppCompatActivity {

    private EditText excelEditText;
    private Button saveButton, openButton, btnZoomIn, btnZoomOut;
    private RecyclerView recyclerView;
    private TableAdapter adapter;
    private LinearLayout tableHeaderLayout;
    
    private List<RowData> dataList = new ArrayList<>();
    private int maxColumnsInFile = 5;

    private int selectedRowIndex = -1;
    private int selectedColIndex = -1;

    private Uri currentFileUri = null;
    private boolean isXlsxFormat = true;
    private float currentScale = 1.0f;

    // Журнал правок пользователя: "строка_колонка" -> новый текст
    private Map<String, String> modifiedCellsMap = new HashMap<>();

    private ActivityResultLauncher<Intent> saveFileLauncher;
    private ActivityResultLauncher<Intent> openFileLauncher;

    public static class RowData {
        public int rowIndex;
        public List<String> columns = new ArrayList<>();

        public RowData(int rowIndex, int colCount) {
            this.rowIndex = rowIndex;
            for (int i = 0; i < colCount; i++) {
                columns.add("");
            }
        }
    }

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        excelEditText = findViewById(R.id.excelEditText);
        saveButton = findViewById(R.id.saveButton);
        openButton = findViewById(R.id.openButton);
        btnZoomIn = findViewById(R.id.btnZoomIn);
        btnZoomOut = findViewById(R.id.btnZoomOut);
        recyclerView = findViewById(R.id.recyclerView);
        tableHeaderLayout = findViewById(R.id.tableHeaderLayout);

        generateEmptyTable();

        recyclerView.setLayoutManager(new LinearLayoutManager(this));
        adapter = new TableAdapter(dataList, (rowIndex, colIndex) -> {
            applyCurrentCellChanges();

            selectedRowIndex = rowIndex;
            selectedColIndex = colIndex;
            
            String colLetter = getColumnLetter(colIndex);
            excelEditText.setHint("Ячейка " + colLetter + (rowIndex + 1));
            
            RowData row = dataList.get(rowIndex);
            while (row.columns.size() <= colIndex) {
                row.columns.add("");
            }
            
            String currentText = row.columns.get(colIndex);
            excelEditText.setText(currentText);
            excelEditText.requestFocus();
            if (excelEditText.getText().length() > 0) {
                excelEditText.setSelection(excelEditText.getText().length());
            }
        });
        recyclerView.setAdapter(adapter);

        btnZoomIn.setOnClickListener(v -> {
            if (currentScale < 2.5f) {
                currentScale += 0.15f;
                adapter.setScaleAndColumns(currentScale, maxColumnsInFile);
                rebuildTableHeader();
            }
        });

        btnZoomOut.setOnClickListener(v -> {
            if (currentScale > 0.5f) {
                currentScale -= 0.15f;
                adapter.setScaleAndColumns(currentScale, maxColumnsInFile);
                rebuildTableHeader();
            }
        });

        initFileLaunchers();

        saveButton.setOnClickListener(v -> {
            applyCurrentCellChanges();
            if (currentFileUri != null) {
                saveExcelAbsoluteProtection(currentFileUri);
            } else {
                openSaveAsDialog();
            }
        });

        openButton.setOnClickListener(v -> {
            Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
            intent.addCategory(Intent.CATEGORY_OPENABLE);
            String[] mimeTypes = {"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel"};
            intent.setType("*/*");
            intent.putExtra(Intent.EXTRA_MIME_TYPES, mimeTypes);
            openFileLauncher.launch(intent);
        });
    }

    private void rebuildTableHeader() {
        tableHeaderLayout.removeAllViews();
        float density = getResources().getDisplayMetrics().density;
        
        int rowNumWidth = (int) (50 * density * currentScale);
        int cellWidth = (int) (100 * density * currentScale);
        int cellHeight = (int) (30 * density * currentScale);

        TextView tvEmpty = new TextView(this);
        tvEmpty.setLayoutParams(new LinearLayout.LayoutParams(rowNumWidth, cellHeight));
        tvEmpty.setBackgroundColor(android.graphics.Color.parseColor("#C0C0C0"));
        tableHeaderLayout.addView(tvEmpty);

        for (int i = 0; i < maxColumnsInFile; i++) {
            TextView tvLetter = new TextView(this);
            tvLetter.setLayoutParams(new LinearLayout.LayoutParams(cellWidth, cellHeight));
            tvLetter.setText(getColumnLetter(i));
            tvLetter.setGravity(Gravity.CENTER);
            tvLetter.setTypeface(android.graphics.Typeface.DEFAULT_BOLD);
            tvLetter.setTextColor(android.graphics.Color.BLACK);
            tvLetter.setTextSize(14 * currentScale);
            tvLetter.setBackgroundResource(R.drawable.grid_cell_border);
            tableHeaderLayout.addView(tvLetter);
        }
    }

    private String getColumnLetter(int colIndex) {
        StringBuilder colLetter = new StringBuilder();
        while (colIndex >= 0) {
            colLetter.insert(0, (char) ('A' + (colIndex % 26)));
            colIndex = (colIndex / 26) - 1;
        }
        return colLetter.toString();
    }

    private void generateEmptyTable() {
        dataList.clear();
        modifiedCellsMap.clear();
        maxColumnsInFile = 5;
        for (int i = 0; i < 50; i++) {
            dataList.add(new RowData(i, maxColumnsInFile));
        }
        currentFileUri = null;
        isXlsxFormat = true;
        saveButton.setText("Сохранить как...");
        rebuildTableHeader();
    }

    private void applyCurrentCellChanges() {
        if (selectedRowIndex != -1 && selectedColIndex != -1) {
            String newText = excelEditText.getText().toString();
            RowData row = dataList.get(selectedRowIndex);
            while (row.columns.size() <= selectedColIndex) {
                row.columns.add("");
            }
            
            String oldText = row.columns.get(selectedColIndex);
            if (!newText.equals(oldText)) {
                row.columns.set(selectedColIndex, newText);
                
                String key = selectedRowIndex + "_" + selectedColIndex;
                modifiedCellsMap.put(key, newText);
                
                adapter.notifyItemChanged(selectedRowIndex);
            }
        }
    }

    private void openSaveAsDialog() {
        Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        intent.setType(isXlsxFormat ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "application/vnd.ms-excel");
        intent.putExtra(Intent.EXTRA_TITLE, isXlsxFormat ? "Table.xlsx" : "Table.xls");
        saveFileLauncher.launch(intent);
    }

    private void initFileLaunchers() {
        saveFileLauncher = registerForActivityResult(
                new ActivityResultContracts.StartActivityForResult(),
                result -> {
                    if (result.getResultCode() == Activity.RESULT_OK && result.getData() != null) {
                        Uri uri = result.getData().getData();
                        if (uri != null) {
                            currentFileUri = uri;
                            saveButton.setText("Сохранить");
                            String fileName = getFileNameFromUri(uri);
                            isXlsxFormat = fileName == null || !fileName.endsWith(".xls");
                            
                            createEmptyWorkbookOnUri(uri);
                            saveExcelAbsoluteProtection(uri);
                        }
                    }
                }
        );

        openFileLauncher = registerForActivityResult(
                new ActivityResultContracts.StartActivityForResult(),
                result -> {
                    if (result.getResultCode() == Activity.RESULT_OK && result.getData() != null) {
                        Uri uri = result.getData().getData();
                        if (uri != null) {
                            currentFileUri = uri;
                            saveButton.setText("Сохранить");
                            String fileName = getFileNameFromUri(uri);
                            isXlsxFormat = fileName == null || !fileName.endsWith(".xls");
                            loadExcel(uri);
                        }
                    }
                }
        );
    }

    private void createEmptyWorkbookOnUri(Uri uri) {
        try (ParcelFileDescriptor pfd = getContentResolver().openFileDescriptor(uri, "w");
             FileOutputStream fos = new FileOutputStream(pfd.getFileDescriptor());
             Workbook workbook = isXlsxFormat ? new XSSFWorkbook() : new HSSFWorkbook()) {
            workbook.createSheet("Sheet1");
            workbook.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // МЕТОД АБСОЛЮТНОЙ БЕЗОПАСНОСТИ ДЛЯ СТРУКТУРЫ И СТИЛЕЙ EXCEL
    private void saveExcelAbsoluteProtection(Uri uri) {
        if (!isXlsxFormat) {
            // Для старых файлов .xls используем безопасный поблочный POIFSFileSystem инжект
            saveXlsInPlace(uri);
            return;
        }

        File tempSourceFile = new File(getFilesDir(), "source_hold.xlsx");
        try {
            // 1. Клонируем исходный файл во временную папку
            try (InputStream is = getContentResolver().openInputStream(uri);
                 FileOutputStream fos = new FileOutputStream(tempSourceFile)) {
                byte[] buf = new byte[4096];
                int len;
                while ((len = is.read(buf)) != -1) {
                    fos.write(buf, 0, len);
                }
            }

            // 2. Распаковываем XLSX как ZIP и переносим файлы. Стили и метаданные копируются БАЙТ-В-БАЙТ.
            ByteArrayOutputStream resultZipBos = new ByteArrayOutputStream();
            try (ZipInputStream zin = new ZipInputStream(new FileInputStream(tempSourceFile));
                 ZipOutputStream zout = new ZipOutputStream(resultZipBos)) {
                
                ZipEntry entry;
                while ((entry = zin.getNextEntry()) != null) {
                    zout.putNextEntry(new ZipEntry(entry.getName()));
                    
                    // Если это основной лист с данными, внедряем inline-строки прямо в XML разметку ячеек
                    if (entry.getName().equals("xl/worksheets/sheet1.xml")) {
                        ByteArrayOutputStream xmlBos = new ByteArrayOutputStream();
                        byte[] buf = new byte[4096];
                        int len;
                        while ((len = zin.read(buf)) != -1) {
                            xmlBos.write(buf, 0, len);
                        }
                        
                        String xmlContent = xmlBos.toString("UTF-8");
                        xmlContent = patchXmlSheetWithInlineStrings(xmlContent);
                        zout.write(xmlContent.getBytes("UTF-8"));
                    } else {
                        // КРИТИЧЕСКИ ВАЖНО: Все остальные файлы архива (styles.xml, styles.xml.rels, 
                        // темы, шрифты, границы, слияния) копируются в оригинальном бинарном виде!
                        byte[] buf = new byte[4096];
                        int len;
                        while ((len = zin.read(buf)) != -1) {
                            zout.write(buf, 0, len);
                        }
                    }
                    zin.closeEntry();
                    zout.closeEntry();
                }
            }

            // 3. Перезаписываем целевой файл в Android полученным чистым ZIP-массивом
            try (ParcelFileDescriptor pfd = getContentResolver().openFileDescriptor(uri, "rwt");
                 FileOutputStream fos = new FileOutputStream(pfd.getFileDescriptor())) {
                fos.write(resultZipBos.toByteArray());
                fos.flush();
            }

            modifiedCellsMap.clear();
            if (tempSourceFile.exists()) tempSourceFile.delete();

            Toast.makeText(this, "Документ сохранен! Все стили и разметка защищены.", Toast.LENGTH_SHORT).show();
            clearEditorFocus();

        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка сохранения: " + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }

    // Хирургическая вставка строк внутрь XML-структуры ячеек листа
    private String patchXmlSheetWithInlineStrings(String xml) {
        for (Map.Entry<String, String> entry : modifiedCellsMap.entrySet()) {
            String[] coords = entry.getKey().split("_");
            int rIdx = Integer.parseInt(coords[0]) + 1; // Строки в XLSX начинаются с 1
            int cIdx = Integer.parseInt(coords[1]);
            String cellRef = getColumnLetter(cIdx) + rIdx; // Формат ячейки: "A1", "C12"
            String newValue = entry.getValue();

            // Ищем тег конкретной ячейки по ее координате: r="A1"
            int cellPos = xml.indexOf("r=\"" + cellRef + "\"");
            if (cellPos != -1) {
                int startTagOpen = xml.lastIndexOf("<c ", cellPos);
                int startTagClose = xml.indexOf(">", cellPos) + 1;
                int endTagOpen = xml.indexOf("</c>", startTagClose);

                if (startTagOpen != -1 && startTagClose != -1 && endTagOpen != -1) {
                    String originalOpeningTag = xml.substring(startTagOpen, startTagClose);
                    
                    // Переводим тип ячейки в inlineStr, чтобы обойти деструктивный sharedStrings словарь Excel
                    if (originalOpeningTag.contains("t=\"s\"")) {
                        originalOpeningTag = originalOpeningTag.replace("t=\"s\"", "t=\"inlineStr\"");
                    } else if (!originalOpeningTag.contains("t=")) {
                        originalOpeningTag = originalOpeningTag.substring(0, originalOpeningTag.length() - 1) + " t=\"inlineStr\">";
                    }

                    // Конструируем контент ячейки
                    String newCellBody = "<is><t>" + escapeXml(newValue) + "</t></is>";
                    
                    // Пересобираем строку XML без затрагивания атрибута s="номер_стиля" в открывающем теге ячейки!
                    xml = xml.substring(0, startTagOpen) + originalOpeningTag + newCellBody + xml.substring(endTagOpen);
                }
            } else {
                // Если ячейки не было в исходном XML листа, аккуратно внедряем ее внутрь существующей строки <row r="...">
                String rowTag = "<row r=\"" + rIdx + "\"";
                int rowPos = xml.indexOf(rowTag);
                if (rowPos != -1) {
                    int rowClose = xml.indexOf(">", rowPos) + 1;
                    String newCellXml = "<c r=\"" + cellRef + "\" t=\"inlineStr\"><is><t>" + escapeXml(newValue) + "</t></is></c>";
                    xml = xml.substring(0, rowClose) + newCellXml + xml.substring(rowClose);
                }
            }
        }
        return xml;
    }

    private String escapeXml(String text) {
        if (text == null) return "";
        return text.replace("&", "&amp;")
                   .replace("<", "&lt;")
                   .replace(">", "&gt;")
                   .replace("\"", "&quot;")
                   .replace("'", "&apos;");
    }

    // Безопасное инпут-сохранение для старого бинарного формата .xls
    private void saveXlsInPlace(Uri uri) {
        File localFile = new File(getFilesDir(), "cache_xls.bin");
        Workbook workbook = null;
        POIFSFileSystem poifs = null;
        try {
            try (InputStream is = getContentResolver().openInputStream(uri);
                 FileOutputStream fos = new FileOutputStream(localFile)) {
                byte[] buf = new byte[4096];
                int len;
                while ((len = is.read(buf)) != -1) {
                    fos.write(buf, 0, len);
                }
            }

            poifs = new POIFSFileSystem(localFile, false);
            workbook = new HSSFWorkbook(poifs);
            Sheet sheet = workbook.getNumberOfSheets() > 0 ? workbook.getSheetAt(0) : workbook.createSheet("Sheet1");

            for (Map.Entry<String, String> entry : modifiedCellsMap.entrySet()) {
                String[] coords = entry.getKey().split("_");
                int rIdx = Integer.parseInt(coords[0]);
                int cIdx = Integer.parseInt(coords[1]);
                
                Row row = sheet.getRow(rIdx);
                if (row == null) row = sheet.createRow(rIdx);
                Cell cell = row.getCell(cIdx);
                if (cell == null) cell = row.createCell(cIdx);
                
                cell.setCellValue(entry.getValue());
            }

            try (FileOutputStream fos = new FileOutputStream(localFile)) {
                workbook.write(fos);
            }
            workbook.close();
            poifs.close();

            try (InputStream is = new FileInputStream(localFile);
                 ParcelFileDescriptor pfd = getContentResolver().openFileDescriptor(uri, "rwt");
                 FileOutputStream fos = new FileOutputStream(pfd.getFileDescriptor())) {
                byte[] buf = new byte[4096];
                int len;
                while ((len = is.read(buf)) != -1) {
                    fos.write(buf, 0, len);
                }
                fos.flush();
            }
            modifiedCellsMap.clear();
            if (localFile.exists()) localFile.delete();
            Toast.makeText(this, "Файл .xls успешно обновлен!", Toast.LENGTH_SHORT).show();
            clearEditorFocus();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка .xls: " + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }

    private void loadExcel(Uri uri) {
        Workbook workbook = null;
        try (InputStream is = getContentResolver().openInputStream(uri)) {
            if (is == null) return;
            
            workbook = isXlsxFormat ? new XSSFWorkbook(is) : new HSSFWorkbook(is);
            Sheet sheet = workbook.getNumberOfSheets() > 0 ? workbook.getSheetAt(0) : workbook.createSheet("Sheet1");
            dataList.clear();
            modifiedCellsMap.clear();

            int maxColsFound = 5;
            int maxRows = Math.max(50, sheet.getLastRowNum() + 1);
            
            for (int i = 0; i < maxRows; i++) {
                Row row = sheet.getRow(i);
                if (row != null && row.getLastCellNum() > maxColsFound) {
                    maxColsFound = row.getLastCellNum();
                }
            }
            maxColumnsInFile = maxColsFound;

            for (int i = 0; i < maxRows; i++) {
                Row row = sheet.getRow(i);
                RowData rowData = new RowData(i, maxColumnsInFile);
                if (row != null) {
                    for (int j = 0; j < maxColumnsInFile; j++) {
                        Cell cell = row.getCell(j);
                        if (cell != null) {
                            switch (cell.getCellType()) {
                                case NUMERIC: rowData.columns.set(j, String.valueOf(cell.getNumericCellValue())); break;
                                case FORMULA: rowData.columns.set(j, cell.getCellFormula()); break;
                                default: rowData.columns.set(j, cell.toString()); break;
                            }
                        }
                    }
                }
                dataList.add(rowData);
            }

            rebuildTableHeader();
            adapter.setScaleAndColumns(currentScale, maxColumnsInFile);
            
            Toast.makeText(this, "Файл успешно прочитан!", Toast.LENGTH_SHORT).show();
            clearEditorFocus();

        } catch (Exception e) {
            e.printStackTrace();
            generateEmptyTable();
            Toast.makeText(this, "Ошибка чтения: " + e.getMessage(), Toast.LENGTH_LONG).show();
        } finally {
            if (workbook != null) {
                try { workbook.close(); } catch (Exception ignored) {}
            }
        }
    }

    private void clearEditorFocus() {
        excelEditText.setText("");
        excelEditText.setHint("Редактировать ячейку");
        selectedRowIndex = -1;
        selectedColIndex = -1;
        excelEditText.clearFocus();
    }

    private String getFileNameFromUri(Uri uri) {
        String result = null;
        if (uri.getScheme().equals("content")) {
            try (Cursor cursor = getContentResolver().query(uri, null, null, null, null)) {
                if (cursor != null && cursor.moveToFirst()) {
                    int index = cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME);
                    if (index != -1) result = cursor.getString(index);
                }
            }
        }
        if (result == null) {
            result = uri.getPath();
            int cut = result.lastIndexOf('/');
            if (cut != -1) result = result.substring(cut + 1);
        }
        return result;
    }
}
