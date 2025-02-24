package io.github.mohammadrezaeicode.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@NoArgsConstructor
@AllArgsConstructor
@Data
@Builder
public class Title {
    private Integer shiftTop;
    private Integer shiftLeft;
    private Integer consommeRow;
    private Integer consommeCol;
    private Number height;
    private String styleId;
    private String text;
    private List<MultiStyleValue> multiStyleValue;
    private Comment comment;

    public String getStyleId() {
        if (styleId == null) {
            return "titleStyle";
        }
        return styleId;
    }

}
