package com.example.miniexcel;

import android.os.Bundle;
import android.widget.Button;
import android.widget.EditText;
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

    // Внутренний класс обязательно должен быть PUBLIC STATIC, чтобы TableAdapter его видел
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

        // Инициализация тестовых данных для сетки таблицы
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
        adapter = new TableAdapter(dataList, (rowIndex, colIndex) -> {
            // Клик по ячейке: выводим текст в поле редактирования
            String currentText = dataList.get(rowIndex).columns[colIndex];
            excelEditText.setText(currentText);
        });
        recyclerView.setAdapter(adapter);
    }
}
