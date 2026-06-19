package com.example.miniexcel;

import android.app.Activity;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.widget.Button;
import android.widget.EditText;
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

    // Лаунчеры для вызова системного проводника Android
    private ActivityResultLauncher<Intent> saveFileLauncher;
    private ActivityResultLauncher<Intent> openFileLauncher;

    public static class RowData {
        public int rowIndex;
        public String[] columns = new String[5];

        public RowData(int rowIndex) {
            this.rowIndex = rowIndex;
            for (int i = 0; i < 5; i++) {
                columns[i] = ""; // По умолчанию ВСЕ ЯЧЕЙКИ АБСОЛЮТНО ПУСТЫЕ
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

        // Создаем абсолютно чистую пустую таблицу при запуске
        generateEmptyTable();

        recyclerView.setLayoutManager(new LinearLayoutManager(this));
        adapter = new TableAdapter(dataList, (rowIndex, colIndex) -> {
            selectedRowIndex = rowIndex;
            selectedColIndex = colIndex;
            
            // Показываем адрес ячейки в подсказке (например: "Редактирование ячейки [строка 3, колонка B]")
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

        // Инициализируем системные диалоги работы с файлами
        initFileLaunchers();

        // Нажатие на кнопку "Сохранить как..." вызывает системный диалог создания файла
        saveButton.setOnClickListener(v -> {
            // Применяем изменения из текстового поля в ячейку, если она выбрана
            if (selectedRowIndex != -1 && selectedColIndex != -1) {
                dataList.get(selectedRowIndex).columns[selectedColIndex] = excelEditText.getText().toString();
                adapter.notifyItemChanged(selectedRowIndex);
            }

            Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
            intent.addCategory(Intent.CATEGORY_OPENABLE);
            intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            intent.putExtra(Intent.EXTRA_TITLE, "NewTable.xlsx"); // Имя по умолчанию
            saveFileLauncher.launch(intent);
        });

        // Нажатие на кнопку "Открыть файл" вызывает системный проводник
        openButton.setOnClickListener(v -> {
            Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
            intent.addCategory(Intent.CATEGORY_OPENABLE);
            intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            openFileLauncher.launch(intent);
        });
    }

    private void generateEmptyTable() {
        dataList.clear();
        for (int i = 0; i < 50; i++) {
            dataList.add(new RowData(i)); // Ячейки создаются полностью пустыми, без текста A1, B2 и т.д.
        }
    }

    private void initFileLaunchers() {
        // Ответ от системы после выбора места сохранения
        saveFileLauncher = registerForActivityResult(
                new ActivityResultContracts.StartActivityForResult(),
                result -> {
                    if (result.getResultCode() == Activity.RESULT_OK && result.getData() != null) {
                        Uri uri = result.getData().getData();
                        if (uri != null) {
                            exportToExcelUri(uri);
                        }
                    }
                }
        );

        // Ответ от системы после выбора файла для открытия
        openFileLauncher = registerForActivityResult(
                new ActivityResultContracts.StartActivityForResult(),
                result -> {
                    if (result.getResultCode() == Activity.RESULT_OK && result.getData() != null) {
                        Uri uri = result.getData().getData();
                        if (uri != null) {
                            importFromExcelUri(uri);
                        }
                    }
                }
        );
    }

    // Запись данных в выбранное пользователем место через Системный Uri
    private void exportToExcelUri(Uri uri) {
        try (Workbook workbook = new XSSFWorkbook();
             OutputStream os = getContentResolver().openOutputStream(uri)) {
            
            Sheet sheet = workbook.createSheet("Sheet1");
            for (int i = 0; i < dataList.size(); i++) {
                Row row = sheet.createRow(i);
                RowData rowData = dataList.get(i);
                for (int j = 0; j < 5; j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(rowData.columns[j]);
                }
            }
            workbook.write(os);
            Toast.makeText(this, "Файл успешно сохранен!", Toast.LENGTH_SHORT).show();
            
            // Сбрасываем фокус редактирования
            excelEditText.setText("");
            excelEditText.setHint("Редактировать ячейку");
            selectedRowIndex = -1;
            selectedColIndex = -1;
            excelEditText.clearFocus();

        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка сохранения: " + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }

    // Чтение данных из выбранного пользователем файла Excel
    private void importFromExcelUri(Uri uri) {
        try (InputStream is = getContentResolver().openInputStream(uri);
             Workbook workbook = new XSSFWorkbook(is)) {

            Sheet sheet = workbook.getSheetAt(0);
            dataList.clear();

            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                RowData rowData = new RowData(i);
                if (row != null) {
                    for (int j = 0; j < 5; j++) {
                        Cell cell = row.getCell(j);
                        if (cell != null) {
                            rowData.columns[j] = cell.toString();
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

        } catch (Exception e) {
            e.printStackTrace();
            generateEmptyTable();
            adapter.notifyDataSetChanged();
            Toast.makeText(this, "Ошибка чтения файла: " + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }
}
