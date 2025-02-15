package io.github.mohammadrezaeicode.excel.model;

import lombok.*;
import io.github.mohammadrezaeicode.excel.model.formula.FormulaMapBody;
import io.github.mohammadrezaeicode.excel.model.func.CommentConditionFunctionInput;
import io.github.mohammadrezaeicode.excel.model.func.MergeRowDataConditionInputFunction;
import io.github.mohammadrezaeicode.excel.model.func.MultiStyleConditionInputFunction;
import io.github.mohammadrezaeicode.excel.model.func.StyleCellConditionInputFunction;
import io.github.mohammadrezaeicode.excel.model.style1.SortAndFilter;

import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

@AllArgsConstructor
@NoArgsConstructor
@Data
@Builder
public class Sheet1 {
    private List<Header> headers;
    private List<Object> data;
    private Boolean withoutHeader;
    private MapSheetDataOption mapSheetDataOption;
    private String backgroundImage;
    private List<ConditionalFormatting> conditionalFormatting;
    private Function<MultiStyleConditionInputFunction, List<MultiStyleValue>> multiStyleCondition;
    private Boolean useSplitBaseOnMatch;
    private Boolean convertStringToNumber;
    private List<ImageTypes> images;
    private Map<String, FormulaMapBody> formula;
    private PageOption pageOption;
    private String name;
    private Title title;
    private Integer shiftTop;
    private Integer shiftLeft;
    private Boolean selected;
    private String tabColor;
    private List<String> merges;
    private String headerStyleKey;
    private Function<MergeRowDataConditionInputFunction, Boolean> mergeRowDataCondition;
    private Function<StyleCellConditionInputFunction, String> styleCellCondition;
    private Function<CommentConditionFunctionInput, Comment> commentCondition;
    private SortAndFilter sortAndFilter;
    private State state;
    private Object headerRowOption;
    private Map<ProtectionOptionKey, Boolean> protectionOption;
    private Number headerHeight;
    private List<Checkbox> checkbox;
    private ViewOption viewOption;
    private Boolean rtl;
    private PageBreak pageBreak;
    private AsTableOption asTable;
    private List<DropDown> dropDowns;
    @AllArgsConstructor
    @Data
    @Builder
    @NoArgsConstructor
    public static class MapSheetDataOption{
        private Method outlineLevel;
        private  Method hidden;
        private Method height;
    }
    public String generateProtectedString(){
        if(this.protectionOption!=null){
            var keys=protectionOption.keySet();
            var res="<sheetProtection ";
            for (var cu:keys) {
                var el=protectionOption.get(cu);
                res +=" " +
                            cu +
                            "=\"" +
                            (el?1:0) +
                    "\" ";
            }
        res += "/>";
            return res;
        }
        return "";
    }
//    public enum MapSheetDataOption {
//        outlineLevel("outlineLevel"),
//        hidden("hidden"),
//        height("height");
//        private final String type;
//
//        MapSheetDataOption(String type) {
//            this.type = type;
//        }
//
//        public String getType() {
//            return type;
//        }
//    }
    public enum State {
        HIDDEN("hidden"),
        VISIBLE("visible");
        private final String label;

        State(String label) {
            this.label = label;
        }

        public String getLabel() {
            return label;
        }
    }

    public enum ProtectionOptionKey {
        SHEET("sheet"),
        FORMAT_CELLS("formatCells"),
        FORMAT_COLUMNS("formatColumns"),
        FORMAT_ROWS("formatRows"),
        INSERT_COLUMNS("insertColumns"),
        INSERT_ROWS("insertRows"),
        INSERT_HYPER_LINKS("insertHyperlinks"),
        DELETE_COLUMNS("deleteColumns"),
        DELETE_ROWS("deleteRows"),
        SORT("sort"),
        AUTO_FILTER("autoFilter"),
        PIVOT_TABLES("pivotTables");
        private final String keyName;

        ProtectionOptionKey(String keyName) {
            this.keyName = keyName;
        }

        public String getKeyName() {
            return keyName;
        }
    }

    public static class ImageTypes {
        private String url;
        private String from;
        private String to;
        private Type type;
        private Extend extent;
        private Margin margin;
        public enum Type {
            ONE("one"),
            TWO("two");
            private final String type;

            Type(String type) {
                this.type = type;
            }

            public String getType() {
                return type;
            }
        }

        public class Extend {
            private Number cx;
            private Number cy;
        }

        public class Margin {

            private Number all;
            private Number right;
            private Number left;
            private Number bottom;
            private Number top;
        }
    }

    @NoArgsConstructor
    @AllArgsConstructor
    @Data
    @Builder
    public static class PageBreak {
        private List<Integer> row;
        private List<Integer> column;
    }

    @NoArgsConstructor
    @AllArgsConstructor
    @Data
    @Builder
    public static class AsTableOption {
        private Type type;
        private Number styleNumber;
        private Boolean firstColumn;
        private Boolean lastColumn;
        private Boolean rowStripes;
        private Boolean columnStripes;



        public enum Type {
            LIGHT("Light"),
            MEDIUM("Medium"),
            DARK("Dark");
            private String type;

            public String getType() {
                return type;
            }

            Type(String type) {
                this.type = type;
            }
        }
    }
}
