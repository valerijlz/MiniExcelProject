package com.example.miniexcel;

import android.app.Activity;
import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.os.Bundle;
import android.provider.OpenableColumns;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;
import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.appcompat.app.AppCompatActivity;
import androidx.recyclerview.widget.LinearLayoutManager;
import androidx.recyclerview.widget.RecyclerView;
import com.example.miniexcel.R;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class MainActivity extends AppCompatActivity {

    private EditText excelEditText;
    private Button saveButton;
    private Button openButton;
    private RecyclerView recyclerView;
    private TableAdapter adapter;
    private List<RowData> dataList = new ArrayList<>();

    private int selectedRowIndex = -1;
    private int selectedColIndex = -1;

    // Ссылка на Uri текущего открытого или сохраненного файла (для прямой перезаписи)
    private Uri currentFileUri = null;
    // Флаг формата текущего файла: true для .xlsx, false для .xls
    private boolean isXlsxFormat = true;
    // Кэш оригинального файла в байтах, чтобы сохранять стили, цвета и разметку оригинального Excel
    private byte[] originalFileBytes = null;

    private ActivityResultLauncher<Intent> saveFileLauncher;
    private ActivityResultLauncher<Intent> openFileLauncher;

    public static class RowData {
        public int rowIndex;
        public String[] columns = new String[5];

        public RowData(int rowIndex) {
            this.rowIndex = rowIndex;
            for (int i = 0; i < 5; i++) {
                columns[i] = "";
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
        recyclerView = findViewById(R.id.recyclerView);

        generateEmptyTable();

        recyclerView.setLayoutManager(new LinearLayoutManager(this));
        adapter = new TableAdapter(dataList, (rowIndex, colIndex) -> {
            selectedRowIndex = rowIndex;
            selectedColIndex = colIndex;
            char colLetter = (char) ('A' + colIndex);
            excelEditText.setHint("Ячейка " + colLetter + (rowIndex + 1));
            
            String currentText = dataList.get(rowIndex).columns[colIndex];
            excelEditText.setText(currentText);
            excelEditText.requestFocus();
            if (excelEditText.getText().length() > 0) {
                excelEditText.setSelection(excelEditText.getText().length());
            }
        });
        recyclerView.setAdapter(adapter);

        initFileLaunchers();

        // Кнопка «Сохранить». Если файл уже открыт — перезаписывает его. Иначе открывает диалог "Сохранить как"
        saveButton.setOnClickListener(v -> {
            applyCurrentCellChanges();
            
            if (currentFileUri != null) {
                // Файл уже существует на диске — перезаписываем его напрямую без лишних вопросов
                saveExcelToUri(currentFileUri);
            } else {
                // Новая таблица — запрашиваем у системы диалог "Сохранить как..."
                openSaveAsDialog();
            }
        });

        // Кнопка «Открыть файл» поддерживает и .xls, и .xlsx
        openButton.setOnClickListener(v -> {
            Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
            intent.addCategory(Intent.CATEGORY_OPENABLE);
            // Указываем оба поддерживаемых типа MIME для Excel
            String[] mimeTypes = {
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // .xlsx
                    "application/vnd.ms-excel" // .xls
            };
            intent.setType("*/*");
            intent.putExtra(Intent.EXTRA_MIME_TYPES, mimeTypes);
            openFileLauncher.launch(intent);
        });
    }

    private void generateEmptyTable() {
        dataList.clear();
        for (int i = 0; i < 50; i++) {
            dataList.add(new RowData(i));
        }
        originalFileBytes = null;
        currentFileUri = null;
        isXlsxFormat = true;
        saveButton.setText("Сохранить как...");
    }

    private void applyCurrentCellChanges() {
        if (selectedRowIndex != -1 && selectedColIndex != -1) {
            dataList.get(selectedRowIndex).columns[selectedColIndex] = excelEditText.getText().toString();
            adapter.notifyItemChanged(selectedRowIndex);
        }
    }

    private void openSaveAsDialog() {
        Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        if (isXlsxFormat) {
            intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            intent.putExtra(Intent.EXTRA_TITLE, "Table.xlsx");
        } else {
            intent.setType("application/vnd.ms-excel");
            intent.putExtra(Intent.EXTRA_TITLE, "Table.xls");
        }
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
                            saveExcelToUri(uri);
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
                            // Определяем расширение/формат открываемого файла
                            String fileName = getFileNameFromUri(uri);
                            isXlsxFormat = fileName == null || !fileName.endsWith(".xls");
                            loadExcelFromUri(uri);
                        }
                    }
                }
        );
    }

    // БЕЗОПАСНОЕ СОХРАНЕНИЕ: Модифицирует существующий файл, сохраняя все его стили и структуру
    private void saveExcelToUri(Uri uri) {
        Workbook workbook = null;
        try {
            if (originalFileBytes != null) {
                // Если файл был импортирован, загружаем его оригинальную структуру из кэша байт
                java.io.ByteArrayInputStream bais = new java.io.ByteArrayInputStream(originalFileBytes);
                workbook = isXlsxFormat ? new XSSFWorkbook(bais) : new HSSFWorkbook(bais);
            } else {
                // Если это новый документ
                workbook = isXlsxFormat ? new XSSFWorkbook() : new HSSFWorkbook();
            }

            // Ищем или создаем первый лист
            Sheet sheet = workbook.getNumberOfSheets() > 0 ? workbook.getSheetAt(0) : workbook.createSheet("Sheet1");

            // Заменяем исключительно текстовые значения ячеек, стили и разметка строк вокруг остаются нетронутыми!
            for (int i = 0; i < dataList.size(); i++) {
                RowData rowData = dataList.get(i);
                Row row = sheet.getRow(i);
                if (row == null) {
                    row = sheet.createRow(i);
                }
                for (int j = 0; j < 5; j++) {
                    Cell cell = row.getCell(j);
                    if (cell == null) {
                        cell = row.createCell(j);
                    }
                    cell.setCellValue(rowData.columns[j]);
                }
            }

            // Пишем изменения обратно в файл на диске смартфона (Перезапись)
            try (OutputStream os = getContentResolver().openOutputStream(uri, "rwt")) {
                workbook.write(os);
            }

            // Обновляем кэш байт текущей версией файла на случай последующих сохранений
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            originalFileBytes = baos.toByteArray();

            Toast.makeText(this, "Файл успешно сохранен!", Toast.LENGTH_SHORT).show();
            clearEditorFocus();

        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка сохранения: " + e.getMessage(), Toast.LENGTH_LONG).show();
        } finally {
            if (workbook != null) {
                try { workbook.close(); } catch (Exception ignored) {}
            }
        }
    }

    // ИМПОРТ ФАЙЛА: считывает данные и сохраняет весь файл в двоичный кэш для защиты стилей
    private void loadExcelFromUri(Uri uri) {
        Workbook workbook = null;
        try (InputStream is = getContentResolver().openInputStream(uri)) {
            
            // Читаем поток файла в массив байт (кэш оригинального файла со всеми стилями и разметкой)
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            byte[] buffer = new byte[4096];
            int bytesRead;
            while ((bytesRead = is.read(buffer)) != -1) {
                baos.write(buffer, 0, bytesRead);
            }
            originalFileBytes = baos.toByteArray();

            // Создаем воркбук на основе считанных байт в зависимости от формата (.xls или .xlsx)
            java.io.ByteArrayInputStream bais = new java.io.ByteArrayInputStream(originalFileBytes);
            workbook = isXlsxFormat ? new XSSFWorkbook(bais) : new HSSFWorkbook(bais);

            Sheet sheet = workbook.getNumberOfSheets() > 0 ? workbook.getSheetAt(0) : workbook.createSheet("Sheet1");
            dataList.clear();

            // Читаем первые 5 колонок для вывода в сетку интерфейса MiniExcel
            int maxRows = Math.max(50, sheet.getLastRowNum() + 1);
            for (int i = 0; i < maxRows; i++) {
                Row row = sheet.getRow(i);
                RowData rowData = new RowData(i);
                if (row != null) {
                    for (int j = 0; j < 5; j++) {
                        Cell cell = row.getCell(j);
                        if (cell != null) {
                            // Приведение любых типов данных (формулы, числа) к тексту для вывода на экран
                            switch (cell.getCellType()) {
                                case NUMERIC: rowData.columns[j] = String.valueOf(cell.getNumericCellValue()); break;
                                case FORMULA: rowData.columns[j] = cell.getCellFormula(); break;
                                default: rowData.columns[j] = cell.toString(); break;
                            }
                        }
                    }
                }
                dataList.add(rowData);
            }

            adapter.notifyDataSetChanged();
            Toast.makeText(this, "Файл успешно открыт!", Toast.LENGTH_SHORT).show();
            clearEditorFocus();

        } catch (Exception e) {
            e.printStackTrace();
            generateEmptyTable();
            adapter.notifyDataSetChanged();
            Toast.makeText(this, "Ошибка импорта структуры Excel: " + e.getMessage(), Toast.LENGTH_LONG).show();
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
