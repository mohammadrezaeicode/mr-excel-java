package io.github.mohammadrezaeicode.excel.model;

import lombok.*;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@EqualsAndHashCode(callSuper = true)
public class ConditionalFormatting extends ConditionalFormattingOption{
    private String start;
    private String end;
}
