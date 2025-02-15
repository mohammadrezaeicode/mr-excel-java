package io.github.mohammadrezaeicode.excel.model;

import lombok.*;
import io.github.mohammadrezaeicode.excel.model.style1.StyleBody;

import java.util.List;

@Builder
@AllArgsConstructor
@NoArgsConstructor
@Data
@EqualsAndHashCode(callSuper = true)
public class ExcelTable1 extends ExcelTableOption {

    private List<Sheet1> sheet;

    public void addStyle(String key, StyleBody styleBody) {
        var style = getStyles();
        style.put(key, styleBody);
        this.setStyles(style);
    }

    public void applyExcelTableOption(ExcelTableOption excelTableOption) {
        if (excelTableOption.getBackend() != null) {
            this.setBackend(excelTableOption.getBackend());
        }
        if (excelTableOption.getNotSave() != null) {
            this.setNotSave(excelTableOption.getNotSave());
        }
        if (excelTableOption.getCreator() != null) {
            this.setCreator(excelTableOption.getCreator());
        }
        if (excelTableOption.getActivateConditionalFormatting() != null) {
            this.setActivateConditionalFormatting(excelTableOption.getActivateConditionalFormatting());
        }
        if (excelTableOption.getFetch() != null) {
            this.setFetch(excelTableOption.getFetch());
        }
        if (excelTableOption.getFileName() != null) {
            this.setFileName(excelTableOption.getFileName());
        }
        if (excelTableOption.getGenerateType() != null) {
            this.setGenerateType(excelTableOption.getGenerateType());
        }
        if (excelTableOption.getAddDefaultTitleStyle() != null) {
            this.setAddDefaultTitleStyle(excelTableOption.getAddDefaultTitleStyle());
        }
        if (excelTableOption.getCreated() != null) {
            this.setCreated(excelTableOption.getCreated());
        }
        if (excelTableOption.getModified() != null) {
            this.setModified(excelTableOption.getModified());
        }
        if (excelTableOption.getNumberOfColumn() != null) {
            this.setNumberOfColumn(excelTableOption.getNumberOfColumn());

        }
        if (excelTableOption.getCreateType() != null) {
            this.setCreateType(excelTableOption.getCreateType());
        }
        if (excelTableOption.getStyles() != null) {
            this.setStyles(excelTableOption.getStyles());
        }
        if (excelTableOption.getFormatMap() != null) {
            this.setFormatMap(excelTableOption.getFormatMap());
        }
    }
}
