package com.example.miniexcel;

import android.app.Activity;
import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.os.Bundle;
import android.provider.OpenableColumns;
import android.view.ScaleGestureDetector;
import android.view.ViewGroup;
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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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

    // Ссылки на элементы шапки для синхронного изменения размера букв при зуме
    private TextView tvHeaderEmpty, tvHeaderA, tvHeaderB, tvHeaderC, tvHeaderD, tvHeaderE;
    private ScaleGestureDetector scaleGestureDetector;
    private float currentScale = 1.0f;

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

        // Инициализируем элементы шапки
        tvHeaderEmpty = findViewById(R.id.tvHeaderEmpty);
        tvHeaderA = findViewById(R.id.tvHeaderA);
        tvHeaderB = findViewById(R.id.tvHeaderB);
        tvHeaderC = findViewById(R.id.tvHeaderC);
        tvHeaderD = findViewById(R.id.tvHeaderD);
        tvHeaderE = findViewById(R.id.tvHeaderE);

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

        // НАСТРОЙКА КЛАССИЧЕСКОГО ЗУМА EXCEL ЖЕСТАМИ ПАЛЬЦЕВ
        scaleGestureDetector = new ScaleGestureDetector(this, new ScaleGestureDetector.SimpleOnScaleGestureListener() {
            @Override
            public boolean onScale(ScaleGestureDetector detector) {
                currentScale *= detector.getScaleFactor();
                currentScale = Math.max(0.6f, Math.min(currentScale, 2.5f)); // Границы зума
                
                // 1. Отдаем новый масштаб в адаптер строк (перерисует ячейки на лету)
                adapter.setScaleFactor(currentScale);
                
                // 2. Масштабируем элементы верхней шапки (A, B, C...)
                updateHeaderScale();
                return true;
            }
        });

        // Привязываем детектор жестов к контейнеру таблицы
        LinearLayout tableContainer = findViewById(R.id.tableContainer);
        tableContainer.setOnTouchListener((v, event) -> {
            scaleGestureDetector.onTouchEvent(event);
            return false; // Позволяем событиям клика проходить дальше к ячейкам
        });

        initFileLaunchers();

        saveButton.setOnClickListener(v -> {
            applyCurrentCellChanges();
            if (currentFileUri != null) {
                saveExcelWithShadowCopy(currentFileUri);
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

    private void updateHeaderScale() {
        float density = getResources().getDisplayMetrics().density;
        ViewGroup.LayoutParams p = tvHeaderEmpty.getLayoutParams();
        p.width = (int) (40 * density * currentScale);
        tvHeaderEmpty.setLayoutParams(p);

        float textSize = 14 * currentScale;
        tvHeaderA.setTextSize(textSize);
        tvHeaderB.setTextSize(textSize);
        tvHeaderC.setTextSize(textSize);
        tvHeaderD.setTextSize(textSize);
        tvHeaderE.setTextSize(textSize);
    }

    private void generateEmptyTable() {
        dataList.clear();
        for (int i = 0; i < 50; i++) {
            dataList.add(new RowData(i));
        }
        deleteShadowCopy();
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
                            saveExcelWithShadowCopy(uri);
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
                            loadExcelWithShadowCopy(uri);
                        }
                    }
                }
        );
    }

    // СТРАТЕГИЯ ТЕНЕВОГО КЛОНИРОВАНИЯ: Полностью сохраняет структуру оригинального файла Excel
    private void saveExcelWithShadowCopy(Uri uri) {
        File shadowFile = new File(getFilesDir(), "shadow_copy.bin");
        Workbook workbook = null;

        try {
            if (shadowFile.exists() && shadowFile.length() > 0) {
                // Если мы работаем с открытым оригинальным файлом, открываем напрямую локальную копию
                try (FileInputStream fis = new FileInputStream(shadowFile)) {
                    workbook = isXlsxFormat ? new XSSFWorkbook(fis) : new HSSFWorkbook(fis);
                }
            } else {
                // Создание нового файла с нуля
                workbook = isXlsxFormat ? new XSSFWorkbook() : new HSSFWorkbook();
            }

            Sheet sheet = workbook.getNumberOfSheets() > 0 ? workbook.getSheetAt(0) : workbook.createSheet("Sheet1");

            // Заменяем ТОЛЬКО ТЕКСТ в первых 5 колонках. Стили, объединенные ячейки, размеры и другие листы/колонки POI НЕ ТРОГАЕТ
            for (int i = 0; i < dataList.size(); i++) {
                RowData rowData = dataList.get(i);
                Row row = sheet.getRow(i);
                if (row == null) row = sheet.createRow(i);
                
                for (int j = 0; j < 5; j++) {
                    Cell cell = row.getCell(j);
                    if (cell == null) cell = row.createCell(j);
                    cell.setCellValue(rowData.columns[j]);
                }
            }

            // Записываем обновленный воркбук обратно в локальный изолированный теневой файл
            try (FileOutputStream fos = new FileOutputStream(shadowFile)) {
                workbook.write(fos);
            }

            // Копируем этот идеальный локальный файл целиком на место внешнего документа через Android Uri (Атомарная замена)
            try (InputStream is = new FileInputStream(shadowFile);
                 OutputStream os = getContentResolver().openOutputStream(uri, "rwt")) {
                if (os != null) {
                    byte[] buf = new byte[4096];
                    int len;
                    while ((len = is.read(buf)) != -1) {
                        os.write(buf, 0, len);
                    }
                    os.flush();
                }
            }

            Toast.makeText(this, "Файл успешно сохранен!", Toast.LENGTH_SHORT).show();
            clearEditorFocus();

        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка сохранения структуры: " + e.getMessage(), Toast.LENGTH_LONG).show();
        } finally {
            if (workbook != null) {
                try { workbook.close(); } catch (Exception ignored) {}
            }
        }
    }

    private void loadExcelWithShadowCopy(Uri uri) {
        File shadowFile = new File(getFilesDir(), "shadow_copy.bin");
        Workbook workbook = null;

        try {
            // Клонируем файл из внешней системы во внутреннюю безопасную память приложения
            try (InputStream is = getContentResolver().openInputStream(uri);
                 FileOutputStream fos = new FileOutputStream(shadowFile)) {
                if (is == null) return;
                byte[] buf = new byte[4096];
                int len;
                while ((len = is.read(buf)) != -1) {
                    fos.write(buf, 0, len);
                }
                fos.flush();
            }

            // Открываем созданный теневой клон
            try (FileInputStream fis = new FileInputStream(shadowFile)) {
                workbook = isXlsxFormat ? new XSSFWorkbook(fis) : new HSSFWorkbook(fis);
            }

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

            adapter.notifyDataSetChanged();
            Toast.makeText(this, "Файл открыт со всеми стилями!", Toast.LENGTH_SHORT).show();
            clearEditorFocus();

        } catch (Exception e) {
            e.printStackTrace();
            generateEmptyTable();
            adapter.notifyDataSetChanged();
            Toast.makeText(this, "Ошибка чтения: " + e.getMessage(), Toast.LENGTH_LONG).show();
        } finally {
            if (workbook != null) {
                try { workbook.close(); } catch (Exception ignored) {}
            }
        }
    }

    private void deleteShadowCopy() {
        File shadowFile = new File(getFilesDir(), "shadow_copy.bin");
        if (shadowFile.exists()) shadowFile.delete();
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
