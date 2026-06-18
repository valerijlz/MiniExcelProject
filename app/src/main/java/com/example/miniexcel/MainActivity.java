package com.example.yourapp;

import android.os.Bundle;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;
import androidx.appcompat.app.AppCompatActivity;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class MainActivity extends AppCompatActivity {

    private EditText excelEditText;
    private Button saveButton;
    private File excelFile;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        excelEditText = findViewById(R.id.excelEditText);
        saveButton = findViewById(findViewById(R.id.saveButton).getId());

        // Файл будет сохраняться во внутреннее хранилище приложения для простоты доступа
        excelFile = new File(getExternalFilesDir(null), "document.xlsx");

        // Загружаем данные при старте, если файл уже существует
        loadExcelData();

        saveButton.setOnClickListener(v -> saveExcelData(excelEditText.getText().toString()));
    }

    private void loadExcelData() {
        if (!excelFile.exists()) return;

        try (FileInputStream fis = new FileInputStream(excelFile);
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(0);
            if (row != null) {
                Cell cell = row.getCell(0);
                if (cell != null) {
                    excelEditText.setText(cell.getStringCellValue());
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка чтения файла", Toast.LENGTH_SHORT).show();
        }
    }

    private void saveExcelData(String text) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Лист1");
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue(text);

            try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                workbook.write(fos);
                Toast.makeText(this, "Сохранено в: " + excelFile.getAbsolutePath(), Toast.LENGTH_LONG).show();
            }
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка сохранения", Toast.LENGTH_SHORT).show();
        }
    }
}