package io.github.mohammadrezaeicode.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@AllArgsConstructor
@Builder
@Data
@NoArgsConstructor
public class SheetRelFlag {
    private Boolean form;
    private Boolean drawing;
    private Boolean vmlDrwing;
    private Boolean comment;
    private Boolean table;
    private Boolean sheetDrawingsPushed;

    public SheetRelFlag(Boolean same) {
        this.form = same;
                this.drawing=same;
        this.vmlDrwing=same;
                this.comment=same;
        this.table=same;
                this.sheetDrawingsPushed=same;
    }
}
