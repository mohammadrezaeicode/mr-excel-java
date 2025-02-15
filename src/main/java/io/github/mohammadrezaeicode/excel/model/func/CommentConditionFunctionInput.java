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
public class CommentConditionFunctionInput {
    private Object data;
    private Object object;
    private Method headerKey;
    private Integer rowIndex;
    private Integer colIndex;
    private Boolean fromHeader;
}
