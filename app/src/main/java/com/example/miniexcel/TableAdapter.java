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

    public interface OnCellClickListener {
        void onCellClick(int rowIndex, int colIndex);
    }

    public TableAdapter(List<RowData> dataList, OnCellClickListener cellClickListener) {
        this.dataList = dataList;
        this.cellClickListener = cellClickListener;
    }

    public void setScaleFactor(float scaleFactor) {
        this.scaleFactor = scaleFactor;
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
        
        holder.tvRowNumber.setText(String.valueOf(rowData.rowIndex + 1));
        holder.tvRowNumber.setTextSize(14 * scaleFactor);
        ViewGroup.LayoutParams rowNumParams = holder.tvRowNumber.getLayoutParams();
        rowNumParams.width = (int) (40 * holder.itemView.getContext().getResources().getDisplayMetrics().density * scaleFactor);
        holder.tvRowNumber.setLayoutParams(rowNumParams);

        holder.cellContainer.removeAllViews();
        
        for (int i = 0; i < 5; i++) {
            final int colIdx = i;
            TextView textView = new TextView(holder.itemView.getContext());
            LinearLayout.LayoutParams params = new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.MATCH_PARENT, 1.0f);
            textView.setLayoutParams(params);
            
            textView.setText(rowData.columns[i]); 
            
            int paddingVertical = (int) (14 * scaleFactor);
            int paddingHorizontal = (int) (6 * scaleFactor);
            textView.setPadding(paddingHorizontal, paddingVertical, paddingHorizontal, paddingVertical);
            
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
