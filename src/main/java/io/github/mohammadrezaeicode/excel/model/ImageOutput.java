package io.github.mohammadrezaeicode.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ImageOutput {
    private String name;
    private String type;
    private ImageInput image;
    private int ref;
    private Integer index;
    private Sheet.ImageTypes input;

}
