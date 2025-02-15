package io.github.mohammadrezaeicode.excel.model.formula;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@AllArgsConstructor
@NoArgsConstructor
@Builder
@Data
public class FormulaProcessResult {
    private String column;
    private Integer row;
    private Boolean needCalcChain;
    private Boolean isCustom;
    private String cell;
    private String chainCell;
}
