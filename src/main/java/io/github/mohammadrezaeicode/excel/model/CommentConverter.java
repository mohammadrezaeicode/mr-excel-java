package io.github.mohammadrezaeicode.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@AllArgsConstructor
@Data
@Builder
@NoArgsConstructor
public class CommentConverter {
    private Boolean hasAuthor;
    private String author;
    private String commentStyle;
    private String[] commentStr;

}
