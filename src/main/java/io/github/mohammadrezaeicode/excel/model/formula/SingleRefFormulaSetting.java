package io.github.mohammadrezaeicode.excel.model.formula;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Builder
@AllArgsConstructor
@NoArgsConstructor
@Data
public class SingleRefFormulaSetting implements FormulaMapBody {
    private SingleRefFormulaType type;
    private String referenceCell;
    private String value;
    private String styleId;
    public enum SingleRefFormulaType {
        LEN("LEN"),
        MODE("MODE"),
        UPPER("UPPER"),
        LOWER("LOWER"),
        PROPER("PROPER"),
        RIGHT("RIGHT"),
        LEFT("LEFT"),
        ABS("ABS"),
        POWER("POWER"),
        MOD("MOD"),
        FLOOR("FLOOR"),
        CEILING("CEILING"),
        ROUND("ROUND"),
        SQRT("SQRT"),
        COS("COS"),
        SIN("SIN"),
        TAN("TAN"),
        COT("COT"),
        COUNTIF("COUNTIF"),
        SUMIF("SUMIF"),
        TRIM("TRIM");
        private final String type;

        SingleRefFormulaType(String type) {
            this.type = type;
        }

        public String getType() {
            return type;
        }
    }
}
