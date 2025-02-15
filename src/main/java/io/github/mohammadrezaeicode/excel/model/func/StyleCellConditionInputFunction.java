package io.github.mohammadrezaeicode.excel.model.func;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Set;

@NoArgsConstructor
@AllArgsConstructor
@Data
@Builder
public class StyleCellConditionInputFunction {
    private Object data;
    private Object header;
    private Integer rowIndex;
    private Integer colIndex;
    private Boolean fromHeader;
    private Set<String> styleKeys;
}
