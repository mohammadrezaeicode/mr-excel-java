package io.github.mohammadrezaeicode.excel.model.row;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class SheetDimension {
    private String start;
    private String end;
}
