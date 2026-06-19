package com.example.miniexcel;

import android.os.Bundle;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;
import androidx.appcompat.app.AppCompatActivity;
import androidx.recyclerview.widget.LinearLayoutManager;
import androidx.recyclerview.widget.RecyclerView;
import com.example.miniexcel.R;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
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

    // Путь к файлу table.xlsx во внутренней памяти приложения
    private File excelFile;

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

        // Определяем файл во внутреннем хранилище (оно защищено и не требует сложных разрешений)
        excelFile = new File(getExternalFilesDir(null), "table.xlsx");

        // Пытаемся автоматически загрузить сохраненный файл при старте.
        // Если файла нет — создаем дефолтную таблицу (A1...E50)
        if (excelFile.exists()) {
            loadExcelFromFile();
        } else {
            generateDefaultTable();
        }

        recyclerView.setLayoutManager(new LinearLayoutManager(this));
        adapter = new TableAdapter(dataList, (rowIndex, colIndex) -> {
            selectedRowIndex = rowIndex;
            selectedColIndex = colIndex;
            String currentText = dataList.get(rowIndex).columns[colIndex];
            excelEditText.setText(currentText);
            excelEditText.requestFocus();
            if (excelEditText.getText().length() > 0) {
                excelEditText.setSelection(excelEditText.getText().length());
            }
        });
        recyclerView.setAdapter(adapter);

        // КНОПКА «СОХРАНИТЬ» (обновляет ячейку и записывает всё в .xlsx файл)
        saveButton.setOnClickListener(v -> {
            if (selectedRowIndex != -1 && selectedColIndex != -1) {
                String newText = excelEditText.getText().toString();
                dataList.get(selectedRowIndex).columns[selectedColIndex] = newText;
                adapter.notifyItemChanged(selectedRowIndex);
                
                // Сразу сохраняем изменения в реальный Excel файл
                saveExcelToFile();

                excelEditText.setText("");
                selectedRowIndex = -1;
                selectedColIndex = -1;
                excelEditText.clearFocus();
            } else {
                Toast.makeText(MainActivity.this, "Сначала выберите ячейку!", Toast.LENGTH_SHORT).show();
            }
        });

        // КНОПКА «ОТКРЫТЬ ФАЙЛ» (принудительно перечитывает .xlsx файл с диска)
        openButton.setOnClickListener(v -> {
            if (excelFile.exists()) {
                loadExcelFromFile();
                adapter.notifyDataSetChanged();
                Toast.makeText(MainActivity.this, "Файл table.xlsx успешно загружен", Toast.LENGTH_SHORT).show();
            } else {
                Toast.makeText(MainActivity.this, "Сохраненный файл еще не создан", Toast.LENGTH_SHORT).show();
            }
        });
    }

    // Генерация стандартной таблицы
    private void generateDefaultTable() {
        dataList.clear();
        for (int i = 0; i < 50; i++) {
            RowData row = new RowData(i);
            row.columns[0] = "A" + (i + 1);
            row.columns[1] = "B" + (i + 1);
            row.columns[2] = "C" + (i + 1);
            row.columns[3] = "D" + (i + 1);
            row.columns[4] = "E" + (i + 1);
            dataList.add(row);
        }
    }

    // МЕХАНИЗМ ЗАПИСИ ТАБЛИЦЫ В EXCEL (.xlsx)
    private void saveExcelToFile() {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("MiniExcelSheet");

            for (int i = 0; i < dataList.size(); i++) {
                Row row = sheet.createRow(i);
                RowData rowData = dataList.get(i);
                for (int j = 0; j < 5; j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(rowData.columns[j]);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                workbook.write(fos);
            }
            Toast.makeText(this, "Файл Excel сохранен!", Toast.LENGTH_SHORT).show();

        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка записи Excel: " + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }

    // МЕХАНИЗМ ЧТЕНИЯ EXCEL ИЗ ФАЙЛА (.xlsx)
    private void loadExcelFromFile() {
        try (FileInputStream fis = new FileInputStream(excelFile);
             Workbook workbook = new XSSFWorkbook(fis)) {

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
                        } else {
                            rowData.columns[j] = "";
                        }
                    }
                }
                dataList.add(rowData);
            }

            // На случай если файл оказался пустым, подстрахуемся
            if (dataList.isEmpty()) {
                generateDefaultTable();
            }

        } catch (Exception e) {
            e.printStackTrace();
            generateDefaultTable();
            Toast.makeText(this, "Ошибка чтения Excel: " + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }
}
