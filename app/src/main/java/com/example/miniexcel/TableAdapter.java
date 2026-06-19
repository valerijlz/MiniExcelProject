package com.example.miniexcel;

import android.view.Gravity;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.LinearLayout;
import android.widget.TextView;
import androidx.annotation.NonNull;
import androidx.recyclerview.widget.RecyclerView;
import java.util.List;

public class TableAdapter extends RecyclerView.Adapter<TableAdapter.RowViewHolder> {

    private final List<MainActivity.RowData> dataList;
    private final OnCellClickListener cellClickListener;

    public interface OnCellClickListener {
        void onCellClick(int rowIndex, int colIndex);
    }

    public TableAdapter(List<MainActivity.RowData> dataList, OnCellClickListener cellClickListener) {
        this.dataList = dataList;
        this.cellClickListener = cellClickListener;
    }

    @NonNull
    @Override
    public RowViewHolder onCreateViewHolder(@NonNull ViewGroup parent, int viewType) {
        View view = LayoutInflater.from(parent.getContext()).inflate(R.layout.item_row, parent, false);
        return new RowViewHolder(view);
    }

    @Override
    public void onBindViewHolder(@NonNull RowViewHolder holder, int position) {
        MainActivity.RowData rowData = dataList.get(position);
        holder.tvRowNumber.setText(String.valueOf(rowData.rowIndex + 1));
        
        holder.cellContainer.removeAllViews();
        
        // Отображаем первые 5 колонок для минималистичного интерфейса
        for (int i = 0; i < 5; i++) {
            final int colIdx = i;
            TextView textView = new TextView(holder.itemView.getContext());
            LinearLayout.LayoutParams params = new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 1.0f);
            
            // Стилизуем под ячейку таблицы
            textView.setLayoutParams(params);
            textView.setText(rowData.columns[i]);
            textView.setPadding(12, 16, 12, 16);
            textView.setGravity(Gravity.CENTER);
            textView.setTextSize(14);
            textView.setBackgroundResource(android.R.drawable.edit_text); // рамка ячейки
            textView.setSingleLine(true);

            // Обработка клика по ячейке для редактирования
            textView.setOnClickListener(v -> cellClickListener.onCellClick(rowData.rowIndex, colIdx));
            
            holder.cellContainer.addView(textView);
        }
    }

    @Override
    public int getItemCount() {
        return dataList.size();
    }

    public static class RowViewHolder extends RecyclerView.ViewHolder {
        TextView tvRowNumber;
        LinearLayout cellContainer;

        public RowViewHolder(@NonNull View itemView) {
            super(itemView);
            tvRowNumber = itemView.findViewById(R.id.tvRowNumber);
            cellContainer = itemView.findViewById(R.id.cellContainer);
        }
    }
}
