package io.github.mohammadrezaeicode.excel.model.formula;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@AllArgsConstructor
@NoArgsConstructor
@Builder
@Data
public class CustomFormulaSetting implements FormulaMapBody{
    private Boolean isArray;
    private String referenceCells;
    private String formula;
    private String returnType;
    private String styleId;
}
