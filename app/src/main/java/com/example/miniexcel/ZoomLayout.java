package com.example.miniexcel;

import android.content.Context;
import android.util.AttributeSet;
import android.view.ScaleGestureDetector;
import android.view.View;
import android.widget.LinearLayout;

public class ZoomLayout extends LinearLayout {
    private float scaleFactor = 1.0f;
    private ScaleGestureDetector scaleDetector;

    public ZoomLayout(Context context) {
        super(context);
        init(context);
    }

    public ZoomLayout(Context context, AttributeSet attrs) {
        super(context, attrs);
        init(context);
    }

    public ZoomLayout(Context context, AttributeSet attrs, int defStyleAttr) {
        super(context, attrs, defStyleAttr);
        init(context);
    }

    private void init(Context context) {
        scaleDetector = new ScaleGestureDetector(context, new ScaleListener());
    }

    @Override
    public boolean onInterceptTouchEvent(android.view.MotionEvent ev) {
        // Перехватываем двухпальцевые жесты для масштабирования
        scaleDetector.onTouchEvent(ev);
        return super.onInterceptTouchEvent(ev) || ev.getPointerCount() > 1;
    }

    @Override
    public boolean onTouchEvent(android.view.MotionEvent ev) {
        scaleDetector.onTouchEvent(ev);
        return true;
    }

    private class ScaleListener extends ScaleGestureDetector.SimpleOnScaleGestureListener {
        @Override
        public boolean onScale(ScaleGestureDetector detector) {
            scaleFactor *= detector.getScaleFactor();
            // Ограничиваем масштаб: минимум 0.5x, максимум 3.0x
            scaleFactor = Math.max(0.5f, Math.min(scaleFactor, 3.0f));

            // Применяем масштаб ко всем дочерним элементам
            for (int i = 0; i < getChildCount(); i++) {
                View child = getChildAt(i);
                child.setScaleX(scaleFactor);
                child.setScaleY(scaleFactor);
                child.setPivotX(detector.getFocusX() - child.getLeft());
                child.setPivotY(detector.getFocusY() - child.getTop());
                child.invalidate();
            }
            return true;
        }
    }
}
