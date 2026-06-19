package com.example.miniexcel;

import android.os.Bundle;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;
import androidx.appcompat.app.AppCompatActivity;
import androidx.recyclerview.widget.LinearLayoutManager;
import androidx.recyclerview.widget.RecyclerView;
import com.example.miniexcel.R;
import java.util.ArrayList;
import java.util.List;

public class MainActivity extends AppCompatActivity {

    private EditText excelEditText;
    private Button saveButton;
    private RecyclerView recyclerView;
    private TableAdapter adapter;
    private List<RowData> dataList = new ArrayList<>();

    // Переменные для хранения индексов текущей редактируемой ячейки
    private int selectedRowIndex = -1;
    private int selectedColIndex = -1;

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
        recyclerView = findViewById(R.id.recyclerView);

        // Инициализация тестовых данных (A1...E50)
        for (int i = 0; i < 50; i++) {
            RowData row = new RowData(i);
            row.columns[0] = "A" + (i + 1);
            row.columns[1] = "B" + (i + 1);
            row.columns[2] = "C" + (i + 1);
            row.columns[3] = "D" + (i + 1);
            row.columns[4] = "E" + (i + 1);
            dataList.add(row);
        }

        recyclerView.setLayoutManager(new LinearLayoutManager(this));
        
        // Обработка клика по ячейке таблицы
        adapter = new TableAdapter(dataList, (rowIndex, colIndex) -> {
            // Запоминаем, какую ячейку выбрал пользователь
            selectedRowIndex = rowIndex;
            selectedColIndex = colIndex;
            
            // Выводим её текст в верхнее поле ввода для редактирования
            String currentText = dataList.get(rowIndex).columns[colIndex];
            excelEditText.setText(currentText);
            
            // Фокусируемся на поле ввода, чтобы сразу можно было писать
            excelEditText.requestFocus();
            excelEditText.setSelection(excelEditText.getText().length());
        });
        
        recyclerView.setAdapter(adapter);

        // ОЖИВЛЯЕМ КНОПКУ «СОХРАНИТЬ»
        saveButton.setOnClickListener(v -> {
            // Проверяем, выбрана ли хоть какая-то ячейка
            if (selectedRowIndex != -1 && selectedColIndex != -1) {
                // Берем текст из верхнего поля ввода
                String newText = excelEditText.getText().toString();
                
                // Записываем его в наш массив данных таблицы
                dataList.get(selectedRowIndex).columns[selectedColIndex] = newText;
                
                // Приказываем таблице обновить отображение этой строки
                adapter.notifyItemChanged(selectedRowIndex);
                
                // Очищаем фокус и поле ввода
                excelEditText.setText("");
                selectedRowIndex = -1;
                selectedColIndex = -1;
                excelEditText.clearFocus();
                
                Toast.makeText(MainActivity.this, "Значение ячейки обновлено", Toast.LENGTH_SHORT).show();
            } else {
                Toast.makeText(MainActivity.this, "Сначала выберите ячейку таблицы!", Toast.LENGTH_SHORT).show();
            }
        });
    }
}
