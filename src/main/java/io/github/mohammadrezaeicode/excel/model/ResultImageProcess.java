package io.github.mohammadrezaeicode.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ResultImageProcess {
    private Setting start;
    private Setting end;

    public void setMargin(Sheet.ImageTypes.Margin margin) {
        if (margin.getAll() != null) {
            this.end.mR = margin.getAll();
            this.end.mB = margin.getAll();
            this.start.mL = margin.getAll();
            this.start.mT = margin.getAll();
        } else {

            if (margin.getRight() != null) {
                this.end.mR = margin.getRight();
            }
            if (margin.getBottom() != null) {
                this.end.mB = margin.getBottom();
            }
            if (margin.getLeft() != null) {
                this.start.mL = margin.getLeft();
            }
            if (margin.getTop() != null) {
                this.start.mT = margin.getTop();
            }
        }
    }

    public void setEndPlusOne(ShapeRC end) {
        this.end.setCol(end.getColNum() + 1);
        this.end.setRow(end.getRowNum() + 1);
    }

    public void setStart(ShapeRC start) {
        this.start.setCol(start.getColNum());
        this.start.setRow(start.getRowNum());
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class Setting {
        private Number col;
        private Number row;
        private Number mL;
        private Number mT;

        private Number mR;
        private Number mB;

    }
}
