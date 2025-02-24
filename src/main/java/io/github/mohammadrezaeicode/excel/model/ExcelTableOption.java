package io.github.mohammadrezaeicode.excel.model;

import io.github.mohammadrezaeicode.excel.model.style.Format;
import io.github.mohammadrezaeicode.excel.model.style.StyleBody;
import io.github.mohammadrezaeicode.excel.model.types.GenerateType;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.HashMap;
import java.util.Map;
import java.util.function.Function;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExcelTableOption {

    private Boolean notSave;
    private String creator;
    private Boolean backend;
    private Boolean activateConditionalFormatting;
    private Function<ImageInput, ImageInput> fetch;
    private String fileName;
    private GenerateType generateType;
    private Boolean addDefaultTitleStyle;
    private String created;
    private String modified;
    private Integer numberOfColumn;
    private String createType;
    private Map<String, StyleBody> styles;
    private Map<String, Format> formatMap;

    public static Builder eBuilder() {
        return new Builder();
    }


    @NoArgsConstructor
    public static class Builder extends ExcelTableOption {

        public Builder notSave(Boolean notSave) {
            super.notSave = notSave;
            return this;
        }

        public Builder creator(String creator) {
            super.creator = creator;
            return this;
        }

        public Builder backend(Boolean backend) {
            super.backend = backend;
            return this;
        }

        public Builder activateConditionalFormatting(Boolean activateConditionalFormatting) {
            super.activateConditionalFormatting = activateConditionalFormatting;
            return this;
        }

        public Builder fetch(Function<ImageInput, ImageInput> fetch) {
            super.fetch = fetch;
            return this;
        }

        public Builder fileName(String fileName) {
            super.fileName = fileName;
            return this;
        }

        public Builder generateType(GenerateType generateType) {
            super.generateType = generateType;
            return this;
        }

        public Builder addDefaultTitleStyle(Boolean addDefaultTitleStyle) {
            super.addDefaultTitleStyle = addDefaultTitleStyle;
            return this;
        }

        public Builder created(String created) {
            super.created = created;
            return this;
        }

        public Builder modified(String modified) {
            super.modified = modified;
            return this;
        }

        public Builder numberOfColumn(Integer numberOfColumn) {
            super.numberOfColumn = numberOfColumn;
            return this;
        }

        public Builder createType(String createType) {
            super.createType = createType;
            return this;
        }

        public Builder styles(Map<String, StyleBody> styles) {
            if (super.styles == null) {

                super.styles = styles;
            } else {
                super.styles.putAll(styles);
            }
            return this;
        }

        public Builder addStyles(String key, StyleBody styleBody) {
            if (super.styles == null) {
                this.setStyles(new HashMap<>());
            }
            super.styles.put(key, styleBody);
            return this;
        }

        public Builder formatMap(Map<String, Format> formatMap) {
            super.formatMap = formatMap;
            return this;
        }

        public ExcelTableOption build() {
            return this;
        }
    }
}
