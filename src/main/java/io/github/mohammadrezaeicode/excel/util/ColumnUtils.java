package io.github.mohammadrezaeicode.excel.util;

import io.github.mohammadrezaeicode.excel.model.ShapeRC;

import java.util.Arrays;
import java.util.List;

public class ColumnUtils {
    public final static List<String> DEFAULT_COLUMN = Arrays.asList("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z");

    public static ShapeRC getColRowBaseOnRefString(
            String refString,
            List<String> cols
    ) {
        String column = refString.replaceAll("[0-9]", "");
        int row = Integer.parseInt(refString.substring(column.length()));

        row = Math.max(0, row - 1);
        int colIndex = cols.indexOf(column);
        if (colIndex < 0) {
            cols = generateColumnName(cols, (int) Math.pow(10, column.length() + 1));
            colIndex = cols.indexOf(column);
            if (colIndex < 0) {
                colIndex = 0;
            }
        }
        return new ShapeRC(String.valueOf(row), String.valueOf(colIndex));
    }

    public static List<String> generateColumnName(
            List<String> cols,
            int num
    ) {
        return generateColumnName(cols, num, "", List.of(), 0);
    }

    public static List<String> generateColumnName(
            List<String> cols,
            int num,
            String startletter,
            List<String> result,
            int nextIndex
    ) {
        int length = cols.size();
        for (int index = 0; index < length; index++) {
            result.add(startletter + cols.get(index));
        }
        if (num < result.size()) {
            return result;
        } else {
            return generateColumnName(
                    cols,
                    num,
                    result.get(nextIndex + 1),
                    result,
                    nextIndex + 1
            );
        }
    }

}
