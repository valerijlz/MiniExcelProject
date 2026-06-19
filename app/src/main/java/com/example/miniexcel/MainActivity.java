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
import java.io.OutputStream;
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
    private float currentScale = 1.0f;

    // Журнал изменений: хранит точные координаты правок пользователя
    // Ключ: "строка_колонка", Значение: новый текст ячейки
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
                saveExcelInPlaceWithoutPoi(currentFileUri);
            } else {
                openSaveAsDialog();
            }
        });

        openButton.setOnClickListener(v -> {
            Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
            intent.addCategory(Intent.CATEGORY_OPENABLE);
            intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
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
        intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        intent.putExtra(Intent.EXTRA_TITLE, "Table.xlsx");
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
                            createEmptyXlsxOnUri(uri);
                            saveExcelInPlaceWithoutPoi(uri);
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
                            loadExcelExcel(uri);
                        }
                    }
                }
        );
    }

    private void createEmptyXlsxOnUri(Uri uri) {
        try (ParcelFileDescriptor pfd = getContentResolver().openFileDescriptor(uri, "w");
             FileOutputStream fos = new FileOutputStream(pfd.getFileDescriptor());
             Workbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Sheet1");
            workbook.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // АБСОЛЮТНАЯ ЗАЩИТА: Сохранение БЕЗ использования Apache POI write()
    // Изменения внедряются напрямую в XML разметку листа внутри ZIP-структуры XLSX!
    private void saveExcelInPlaceWithoutPoi(Uri uri) {
        File tempFile = new File(getFilesDir(), "temp_update.xlsx");
        
        try {
            // 1. Копируем исходный файл во временное хранилище
            try (InputStream is = getContentResolver().openInputStream(uri);
                 FileOutputStream fos = new FileOutputStream(tempFile)) {
                byte[] buf = new byte[4096];
                int len;
                while ((len = is.read(buf)) != -1) {
                    fos.write(buf, 0, len);
                }
            }

            // Читаем временный файл и создаем измененную копию через ZIP потоки
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ZipInputStream zin = new ZipInputStream(new FileInputStream(tempFile));
                 ZipOutputStream zout = new ZipOutputStream(bos)) {
                
                ZipEntry entry;
                while ((entry = zin.getNextEntry()) != null) {
                    zout.putNextEntry(new ZipEntry(entry.getName()));
                    
                    // Если нашли файл данных первого листа — модифицируем его текст
                    if (entry.getName().equals("xl/worksheets/sheet1.xml")) {
                        ByteArrayOutputStream sheetXmlBos = new ByteArrayOutputStream();
                        byte[] buf = new byte[4096];
                        int len;
                        while ((len = zin.read(buf)) != -1) {
                            sheetXmlBos.write(buf, 0, len);
                        }
                        
                        String sheetXml = sheetXmlBos.toString("UTF-8");
                        sheetXml = injectModifiedCellsIntoXml(sheetXml);
                        
                        zout.write(sheetXml.getBytes("UTF-8"));
                    } else {
                        // Все остальные файлы (стили, темы, шрифты, связи) копируем БАЙТ В БАЙТ без изменений!
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

            // 2. Записываем результирующий ZIP архив обратно в документ Android (Перезапись данных)
            try (ParcelFileDescriptor pfd = getContentResolver().openFileDescriptor(uri, "rwt");
                 FileOutputStream fos = new FileOutputStream(pfd.getFileDescriptor())) {
                fos.write(bos.toByteArray());
                fos.flush();
            }

            modifiedCellsMap.clear(); // Очищаем журнал
            if (tempFile.exists()) tempFile.delete();

            Toast.makeText(this, "Файл успешно модифицирован. Стили сохранены!", Toast.LENGTH_SHORT).show();
            clearEditorFocus();

        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка записи: " + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }

    // Внедрение правок в XML-код листа с сохранением окружающих тегов разметки
    private String injectModifiedCellsIntoXml(String xml) {
        for (Map.Entry<String, String> entry : modifiedCellsMap.entrySet()) {
            String[] coords = entry.getKey().split("_");
            int rIdx = Integer.parseInt(coords[0]) + 1; // В XLSX строки начинаются с 1
            int cIdx = Integer.parseInt(coords[1]);
            String cellRef = getColumnLetter(cIdx) + rIdx; // Например: "A1", "B5"
            String newValue = entry.getValue();

            // Шаблон поиска ячейки: <c r="A1" ...> ... </c>
            String pattern = "<c[^>]*r=\"" + cellRef + "\"[^>]*>";
            int cellPos = xml.indexOf("r=\"" + cellRef + "\"");

            if (cellPos != -1) {
                // Находим границы открывающего тега <c> и закрывающего </c>
                int startTagOpen = xml.lastIndexOf("<c ", cellPos);
                int startTagClose = xml.indexOf(">", cellPos) + 1;
                int endTagOpen = xml.indexOf("</c>", startTagClose);

                if (startTagOpen != -1 && startTagClose != -1 && endTagOpen != -1) {
                    // Извлекаем существующий тег <c ...> для сохранения стиля s="индекс_стиля"
                    String openingTag = xml.substring(startTagOpen, startTagClose);
                    
                    // Чтобы Excel не искал текст в таблице общих строк (shared strings), 
                    // принудительно переводим тип ячейки в inlineStr (прямая строка)
                    if (openingTag.contains("t=\"s\"")) {
                        openingTag = openingTag.replace("t=\"s\"", "t=\"inlineStr\"");
                    } else if (!openingTag.contains("t=")) {
                        openingTag = openingTag.substring(0, openingTag.length() - 1) + " t=\"inlineStr\">";
                    }

                    // Конструируем внутреннее тело ячейки для сохранения обычного текста
                    String newBody = "<is><t>" + escapeXml(newValue) + "</t></is>";
                    
                    // Пересобираем XML строку листа
                    xml = xml.substring(0, startTagOpen) + openingTag + newBody + xml.substring(endTagOpen);
                }
            } else {
                // Если ячейки физически не существовало в XML, внедряем её внутрь тега строки <row r="строка">
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

    private void loadExcelExcel(Uri uri) {
        Workbook workbook = null;
        try (InputStream is = getContentResolver().openInputStream(uri)) {
            if (is == null) return;
            workbook = new XSSFWorkbook(is);

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
            
            Toast.makeText(this, "Файл прочитан. Защита стилей активна!", Toast.LENGTH_SHORT).show();
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
}
