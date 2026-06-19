package com.example.miniexcel;

import android.view.Gravity;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.LinearLayout;
import android.widget.TextView;
import androidx.annotation.NonNull;
import androidx.recyclerview.widget.RecyclerView;
import com.example.miniexcel.R;
import com.example.miniexcel.MainActivity.RowData;
import java.util.List;

public class TableAdapter extends RecyclerView.Adapter<TableAdapter.RowViewHolder> {

    private final List<RowData> dataList;
    private final OnCellClickListener cellClickListener;
    private float scaleFactor = 1.0f;
    private int maxColumnsCount = 5; // По умолчанию отображаем 5 колонок, но расширяем динамически при открытии Excel

    public interface OnCellClickListener {
        void onCellClick(int rowIndex, int colIndex);
    }

    public TableAdapter(List<RowData> dataList, OnCellClickListener cellClickListener) {
        this.dataList = dataList;
        this.cellClickListener = cellClickListener;
    }

    public void setScaleAndColumns(float scaleFactor, int columnsCount) {
        this.scaleFactor = scaleFactor;
        this.maxColumnsCount = Math.max(5, columnsCount);
        notifyDataSetChanged();
    }

    @NonNull
    @Override
    public RowViewHolder onCreateViewHolder(@NonNull ViewGroup parent, int viewType) {
        View view = LayoutInflater.from(parent.getContext()).inflate(R.layout.item_row, parent, false);
        return new RowViewHolder(view);
    }

    @Override
    public void onBindViewHolder(@NonNull RowViewHolder holder, int position) {
        RowData rowData = dataList.get(position);
        float density = holder.itemView.getContext().getResources().getDisplayMetrics().density;

        // Рассчитываем пропорциональные размеры
        int rowNumWidth = (int) (50 * density * scaleFactor);
        int cellWidth = (int) (100 * density * scaleFactor);
        int cellHeight = (int) (40 * density * scaleFactor);

        // Масштабируем боковой номер строки
        holder.tvRowNumber.setText(String.valueOf(rowData.rowIndex + 1));
        holder.tvRowNumber.setTextSize(14 * scaleFactor);
        ViewGroup.LayoutParams rowNumParams = holder.tvRowNumber.getLayoutParams();
        rowNumParams.width = rowNumWidth;
        rowNumParams.height = cellHeight;
        holder.tvRowNumber.setLayoutParams(rowNumParams);

        holder.cellContainer.removeAllViews();
        
        // Рендерим ровно столько колонок, сколько реально есть в документе Excel
        for (int i = 0; i < maxColumnsCount; i++) {
            final int colIdx = i;
            TextView textView = new TextView(holder.itemView.getContext());
            
            // СТРОГИЕ ПРОПОРЦИИ КВАДРАТНОЙ ФОРМЫ ЯЧЕЙКИ С УЧЕТОМ ЗУМА
            LinearLayout.LayoutParams params = new LinearLayout.LayoutParams(cellWidth, cellHeight);
            textView.setLayoutParams(params);
            
            // Проверяем, есть ли данные для текущего индекса столбца
            if (i < rowData.columns.size()) {
                textView.setText(rowData.columns.get(i));
            } else {
                textView.setText("");
            }
            
            textView.setGravity(Gravity.CENTER);
            textView.setTextSize(14 * scaleFactor);
            textView.setBackgroundResource(R.drawable.grid_cell_border);
            textView.setSingleLine(true);

            textView.setOnClickListener(v -> cellClickListener.onCellClick(rowData.rowIndex, colIdx));
            
            holder.cellContainer.addView(textView);
        }
    }

    @Override
    public int getItemCount() {
        return dataList.size();
    }

    public static class RowViewHolder extends RecyclerView.ViewHolder {
        public TextView tvRowNumber;
        public LinearLayout cellContainer;

        public RowViewHolder(@NonNull View itemView) {
            super(itemView);
            tvRowNumber = itemView.findViewById(R.id.tvRowNumber);
            cellContainer = itemView.findViewById(R.id.cellContainer);
        }
    }
}
