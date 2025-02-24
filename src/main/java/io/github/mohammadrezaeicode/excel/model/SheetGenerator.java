package io.github.mohammadrezaeicode.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.function.Function;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class SheetGenerator<T> {
    Class headerClass;
    Function<List<Header>, List<Header>> applyHeaderOptionFunction;
    Function<Sheet.SheetBuilder, Sheet> applySheetOptionFunction;
    private List<T> data;

}
