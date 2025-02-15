package io.github.mohammadrezaeicode.excel.model;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class Position {
    private ShapeRC start;
    private ShapeRC end;

    public ShapeRC getStart() {
        return start;
    }

    public ShapeRC getEnd() {
        return end;
    }

    public Position() {
        start=new ShapeRC("0","0");
        end=new ShapeRC("1","1");
    }

    public Position(ShapeRC start, ShapeRC end) {
        this.start = start;
        this.end = end;
    }

    public void setStart(ShapeRC start) {
        this.start = start;
    }

    public void setEnd(ShapeRC end) {
        this.end = end;
    }
}
