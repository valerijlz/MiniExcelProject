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

    public interface OnCellClickListener {
        void onCellClick(int rowIndex, int colIndex);
    }

    public TableAdapter(List<RowData> dataList, OnCellClickListener cellClickListener) {
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
        RowData rowData = dataList.get(position);
        holder.tvRowNumber.setText(String.valueOf(rowData.rowIndex + 1));
        
        holder.cellContainer.removeAllViews();
        
        for (int i = 0; i < 5; i++) {
            final int colIdx = i;
            TextView textView = new TextView(holder.itemView.getContext());
            LinearLayout.LayoutParams params = new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.MATCH_PARENT, 1.0f);
            
            textView.setLayoutParams(params);
            textView.setText(rowData.columns[i]); // Ячейка будет пустой, если там нет сохраненного текста
            textView.setPadding(12, 24, 12, 24);
            textView.setGravity(Gravity.CENTER);
            textView.setTextSize(14);
            // Рисуем легкую рамку вокруг ячейки для сходства с таблицей Excel
            textView.setBackgroundResource(android.R.drawable.dialog_holo_light_frame);
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
