package io.github.mohammadrezaeicode.excel.model.formula;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@AllArgsConstructor
@NoArgsConstructor
@Data
@Builder
public class FormulaSetting implements FormulaMapBody {
    private FormulaType type;
    private String start;
    private String end;
    private String styleId;

    public enum FormulaType {
        AVERAGE("AVERAGE"),
        SUM("SUM"),
        COUNT("COUNT"),
        MAX("MAX"),
        MIN("MIN");
        private final String type;

        FormulaType(String type) {
            this.type = type;
        }

        public String getType() {
            return type;
        }
    }
}
