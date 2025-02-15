package io.github.mohammadrezaeicode.excel.model.func;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.lang.reflect.Method;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class MergeRowDataConditionInputFunction {
    private Object data;
    private Method key;
    private Integer index;
    private Boolean fromHeader;
}
