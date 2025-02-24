package io.github.mohammadrezaeicode.excel.model;

import lombok.*;

import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
@EqualsAndHashCode(callSuper = true)
public class DataOption extends MergeConfig {
    private Number outlineLevel;
    private Boolean hidden;
    private String rowStyle;
    private Number height;
    private Map<Method, List<MultiStyleValue>> multiStyleValue;
    private Map<Method, Comment> comment;

    public void addComment(Method key, Comment checkCommentCondition) {
        if (comment == null) {
            comment = new HashMap<>();
        }
        comment.put(key, checkCommentCondition);
    }

    public void addMultiStyle(Method key, List<MultiStyleValue> multi) {
        if (multiStyleValue == null) {
            multiStyleValue = new HashMap<>();
        }
        multiStyleValue.put(key, multi);
    }

    public boolean getMultiStyleValueKeyExist(Method key) {
        if (multiStyleValue != null) {
            return multiStyleValue.containsKey(key);
        }
        return false;
    }
}
