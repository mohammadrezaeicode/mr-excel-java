package io.github.mohammadrezaeicode.excel.model;

import lombok.*;

import java.lang.reflect.Method;

@Builder
@NoArgsConstructor
@AllArgsConstructor
@Data
@EqualsAndHashCode(callSuper = true)
public class Header extends HeaderOption {
    private String label;
    private String text;
    private Method method;

    public void applyHeaderOption(HeaderOption headerOption) {
        if (headerOption.getSize() != null) {
            headerOption.setSize(headerOption.getSize());
        }
        if (headerOption.getComment() != null) {
            headerOption.setComment(headerOption.getComment());
        }
        if (headerOption.getFormula() != null) {
            headerOption.setFormula(headerOption.getFormula());
        }
        if (headerOption.getConditionalFormatting() != null) {
            headerOption.setConditionalFormatting(headerOption.getConditionalFormatting());
        }
        if (headerOption.getMultiStyleValue() != null) {
            headerOption.setMultiStyleValue(headerOption.getMultiStyleValue());
        }

    }
}
