package io.github.mohammadrezaeicode.excel.model;

import lombok.Builder;

@Builder
public class ShapeRC {
    private String row;
    private String col;

    public ShapeRC(String row, String col) {
        this.row = row;
        this.col = col;
    }

    public String getRow() {
        return row;
    }

    public void setRow(String row) {
        this.row = row;
    }

    public String getCol() {
        return col;
    }

    public void setCol(String col) {
        this.col = col;
    }

    public Integer getRowNum() {
        try {
            return Integer.parseInt(row);
        } catch (NumberFormatException exception) {
            return -1;
        }
    }

    public Integer getColNum() {
        try {
            return Integer.parseInt(col);
        } catch (
                NumberFormatException exception) {
            return -1;
        }
    }
}
