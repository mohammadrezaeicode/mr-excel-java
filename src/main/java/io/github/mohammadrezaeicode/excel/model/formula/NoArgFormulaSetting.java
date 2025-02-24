package io.github.mohammadrezaeicode.excel.model.formula;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@AllArgsConstructor
@NoArgsConstructor
@Data
@Builder
public class NoArgFormulaSetting implements FormulaMapBody {
    private NoArgFormulaType noArgType;
    private String styleId;

    public enum NoArgFormulaType {
        NOW("NOW"),
        TODAY("TODAY"),
        HOUR("HOUR"),
        NOW_YEAR("NOW_YEAR"),
        NOW_HOUR("NOW_HOUR"),
        NOW_SECOND("NOW_SECOND"),
        NOW_MIN("NOW_MIN"),
        NOW_MONTH("NOW_MONTH"),
        NOW_DAY("NOW_DAY"),
        NOW_WEEKDAY("NOW_WEEKDAY"),
        NOW_MINUTE("NOW_MINUTE");
        private final String type;

        NoArgFormulaType(String type) {
            this.type = type;
        }

        public String getType() {
            return type;
        }
    }
}
