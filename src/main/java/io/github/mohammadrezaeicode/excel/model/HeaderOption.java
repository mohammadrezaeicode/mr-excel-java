package io.github.mohammadrezaeicode.excel.model;

import io.github.mohammadrezaeicode.excel.model.formula.FormulaSetting;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class HeaderOption {
    private Comment comment;
    private Integer size;
    private List<MultiStyleValue> multiStyleValue;
    private ConditionalFormatting conditionalFormatting;
    private FormulaSetting formula;
}
