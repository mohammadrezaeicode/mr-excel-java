package io.github.mohammadrezaeicode.excel.model;

import io.github.mohammadrezaeicode.excel.model.style1.StyleBody;
import io.github.mohammadrezaeicode.excel.model.types.GenerateType;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import io.github.mohammadrezaeicode.excel.model.style1.Format;

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
    private Function fetch;
    private String fileName;
    private GenerateType generateType;
    private Boolean addDefaultTitleStyle;
    private String created;
    private String modified;
    private Integer numberOfColumn;
    private String createType;
    private Map<String, StyleBody> styles;
    private Map<String, Format> formatMap;
    public static Builder eBuilder(){
        return new Builder();
    }
    @NoArgsConstructor
    public static class Builder extends ExcelTableOption{

        public Builder notSave(Boolean notSave){
            super.notSave=notSave;
            return this;
        }
                public Builder creator(String creator){
            super.creator=creator;
            return this;
                }
                public Builder backend(Boolean backend){
            super.backend=backend;
            return this;
                }
                public Builder activateConditionalFormatting(Boolean activateConditionalFormatting){
            super.activateConditionalFormatting=activateConditionalFormatting;
            return this;
                }
                public Builder fetch(Function fetch){
            super.fetch=fetch;
            return this;
                }
                public Builder fileName(String fileName){
            super.fileName=fileName;
            return this;
                }
                public Builder generateType(GenerateType generateType){
            super.generateType=generateType;
            return this;
                }
                public Builder addDefaultTitleStyle(Boolean addDefaultTitleStyle){
            super.addDefaultTitleStyle=addDefaultTitleStyle;
            return this;
                }
                public Builder created(String created){
            super.created=created;
            return this;
                }
                public Builder modified(String modified){
            super.modified=modified;
            return this;
                }
                public Builder numberOfColumn(Integer numberOfColumn){
            super.numberOfColumn=numberOfColumn;
            return this;
                }
                public Builder createType(String createType){
            super.createType=createType;
            return this;
                }
                public Builder styles(Map<String, StyleBody> styles){
            super.styles=styles;
            return this;
                }
                public Builder formatMap(Map<String, Format> formatMap){
            super.formatMap=formatMap;
            return this;
                }
                public ExcelTableOption build(){
            return (ExcelTableOption) this;
                }
    }
}
