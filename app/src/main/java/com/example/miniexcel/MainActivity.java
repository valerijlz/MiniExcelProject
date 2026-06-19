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
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
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

    private Uri currentFileUri = null;
    private boolean isXlsxFormat = true;
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

        saveButton.setOnClickListener(v -> {
            applyCurrentCellChanges();
            if (currentFileUri != null) {
                saveExcelToUri(currentFileUri);
            } else {
                openSaveAsDialog();
            }
        });

        openButton.setOnClickListener(v -> {
            Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
            intent.addCategory(Intent.CATEGORY_OPENABLE);
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
                            String fileName = getFileNameFromUri(uri);
                            isXlsxFormat = fileName == null || !fileName.endsWith(".xls");
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
                            String fileName = getFileNameFromUri(uri);
                            isXlsxFormat = fileName == null || !fileName.endsWith(".xls");
                            loadExcelFromUri(uri);
                        }
                    }
                }
        );
    }

    // БУФЕРИЗИРОВАННОЕ АТОМАРНОЕ СОХРАНЕНИЕ (Защищает структуру от разрушения и принудительно пишет текст)
    private void saveExcelToUri(Uri uri) {
        Workbook workbook = null;
        File tempFile = null;
        try {
            // Создаем временный файл в изолированном кэше Android для безопасных манипуляций со структурой POI
            tempFile = File.createTempFile("excel_buffer", isXlsxFormat ? ".xlsx" : ".xls", getCacheDir());

            if (originalFileBytes != null && originalFileBytes.length > 0) {
                // Если мы открыли существующий файл, разворачиваем его структуру во временный файл
                try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                    fos.write(originalFileBytes);
                }
                try (FileInputStream fis = new FileInputStream(tempFile)) {
                    workbook = isXlsxFormat ? new XSSFWorkbook(fis) : new HSSFWorkbook(fis);
                }
            } else {
                // Если это совершенно новый файл
                workbook = isXlsxFormat ? new XSSFWorkbook() : new HSSFWorkbook();
            }

            // Получаем доступ к первому листу
            Sheet sheet = workbook.getNumberOfSheets() > 0 ? workbook.getSheetAt(0) : workbook.createSheet("Sheet1");

            // Заменяем исключительно текстовое наполнение ячеек в рамках MiniExcel. Стили, цвет шрифта и границы остаются нетронутыми!
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

            // Записываем обновленный Workbook во временный файл-буфер
            try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                workbook.write(fos);
            }

            // Читаем измененный буфер обратно в байтовый кэш оперативной памяти приложения
            try (FileInputStream fis = new FileInputStream(tempFile);
                 ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
                byte[] buffer = new byte[4096];
                int read;
                while ((read = fis.read(buffer)) != -1) {
                    baos.write(buffer, 0, read);
                }
                originalFileBytes = baos.toByteArray();
            }

            // Перезаписываем оригинальный файл на диске смартфона атомарным бинарным потоком
            try (OutputStream os = getContentResolver().openOutputStream(uri, "rwt")) {
                if (os != null) {
                    os.write(originalFileBytes);
                    os.flush(); // Принудительно выталкиваем данные из кэша Android на физический накопитель
                }
            }

            Toast.makeText(this, "Файл успешно перезаписан и сохранен!", Toast.LENGTH_SHORT).show();
            clearEditorFocus();

        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка записи файла: " + e.getMessage(), Toast.LENGTH_LONG).show();
        } finally {
            if (workbook != null) {
                try { workbook.close(); } catch (Exception ignored) {}
            }
            if (tempFile != null && tempFile.exists()) {
                tempFile.delete(); // Удаляем временный буфер из кэша системы
            }
        }
    }

    private void loadExcelFromUri(Uri uri) {
        Workbook workbook = null;
        try (InputStream is = getContentResolver().openInputStream(uri)) {
            if (is == null) return;

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            byte[] buffer = new byte[4096];
            int bytesRead;
            while ((bytesRead = is.read(buffer)) != -1) {
                baos.write(buffer, 0, bytesRead);
            }
            originalFileBytes = baos.toByteArray();

            java.io.ByteArrayInputStream bais = new java.io.ByteArrayInputStream(originalFileBytes);
            workbook = isXlsxFormat ? new XSSFWorkbook(bais) : new HSSFWorkbook(bais);

            Sheet sheet = workbook.getNumberOfSheets() > 0 ? workbook.getSheetAt(0) : workbook.createSheet("Sheet1");
            dataList.clear();

            int maxRows = Math.max(50, sheet.getLastRowNum() + 1);
            for (int i = 0; i < maxRows; i++) {
                Row row = sheet.getRow(i);
                RowData rowData = new RowData(i);
                if (row != null) {
                    for (int j = 0; j < 5; j++) {
                        Cell cell = row.getCell(j);
                        if (cell != null) {
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

            if (dataList.isEmpty()) {
                generateEmptyTable();
            }
            
            adapter.notifyDataSetChanged();
            Toast.makeText(this, "Файл успешно открыт!", Toast.LENGTH_SHORT).show();
            clearEditorFocus();

        } catch (Exception e) {
            e.printStackTrace();
            generateEmptyTable();
            adapter.notifyDataSetChanged();
            Toast.makeText(this, "Ошибка импорта Excel: " + e.getMessage(), Toast.LENGTH_LONG).show();
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
