package io.github.mohammadrezaeicode.excel.model.converter;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Builder
@AllArgsConstructor
@NoArgsConstructor
@Data
public class RowMap {
    private String startTag;
    private String endTag;
    private String details;
}
