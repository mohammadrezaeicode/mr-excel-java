package io.github.mohammadrezaeicode.excel.model;

import java.util.List;
import java.util.function.Function;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class SheetGenerator<T> {
    private List<T> data;
    Class headerClass;
    Function<List<Header>,List<Header>> applyHeaderOptionFunction;
}
