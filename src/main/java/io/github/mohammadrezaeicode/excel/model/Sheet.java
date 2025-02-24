package io.github.mohammadrezaeicode.excel.model;

import io.github.mohammadrezaeicode.excel.model.formula.FormulaMapBody;
import io.github.mohammadrezaeicode.excel.model.func.CommentConditionFunctionInput;
import io.github.mohammadrezaeicode.excel.model.func.MergeRowDataConditionInputFunction;
import io.github.mohammadrezaeicode.excel.model.func.MultiStyleConditionInputFunction;
import io.github.mohammadrezaeicode.excel.model.func.StyleCellConditionInputFunction;
import io.github.mohammadrezaeicode.excel.model.style.SortAndFilter;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

@AllArgsConstructor
@NoArgsConstructor
@Data
@Builder
public class Sheet {
    private List<Header> headers;
    private List<Object> data;
    private Boolean withoutHeader;
    private MapSheetDataOption mapSheetDataOption;
    private ImageInput backgroundImage;
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

    public String generateProtectedString() {
        if (this.protectionOption != null) {
            var keys = protectionOption.keySet();
            var res = "<sheetProtection ";
            for (var cu : keys) {
                var el = protectionOption.get(cu);
                res += " " +
                        cu.getKeyName() +
                        "=\"" +
                        (el ? 1 : 0) +
                        "\" ";
            }
            res += "/>";
            return res;
        }
        return "";
    }

    public void applySheet(Sheet sheet) {
        if (sheet.getHeaders() != null) {
            if (this.headers != null) {
                this.headers.addAll(sheet.getHeaders());
            } else {
                this.setHeaders(sheet.getHeaders());
            }
        }
        if (sheet.getData() != null) {
            if (this.data != null) {
                this.data.addAll(sheet.getData());
            } else {
                this.setData(sheet.getData());
            }
        }
        if (sheet.getWithoutHeader() != null) {
            this.setWithoutHeader(sheet.getWithoutHeader());
        }
        if (sheet.getMapSheetDataOption() != null) {
            this.setMapSheetDataOption(sheet.getMapSheetDataOption());
        }
        if (sheet.getBackgroundImage() != null) {
            this.setBackgroundImage(sheet.getBackgroundImage());
        }
        if (sheet.getConditionalFormatting() != null) {
            if (this.conditionalFormatting != null) {
                this.conditionalFormatting.addAll(sheet.getConditionalFormatting());
            } else {
                this.setConditionalFormatting(sheet.getConditionalFormatting());
            }
        }
        if (sheet.getMultiStyleCondition() != null) {
            this.setMultiStyleCondition(sheet.getMultiStyleCondition());
        }
        if (sheet.getUseSplitBaseOnMatch() != null) {
            this.setUseSplitBaseOnMatch(sheet.getUseSplitBaseOnMatch());
        }
        if (sheet.getConvertStringToNumber() != null) {
            this.setConvertStringToNumber(sheet.getConvertStringToNumber());
        }
        if (sheet.getImages() != null) {
            if (this.images != null) {
                this.images.addAll(sheet.getImages());
            } else {
                this.setImages(sheet.getImages());
            }
        }
        if (sheet.getFormula() != null) {
            if (this.formula != null) {
                this.formula.putAll(sheet.getFormula());
            } else {
                this.setFormula(sheet.getFormula());
            }
        }
        if (sheet.getPageOption() != null) {
            this.setPageOption(sheet.getPageOption());
        }
        if (sheet.getName() != null) {
            this.setName(sheet.getName());
        }
        if (sheet.getTitle() != null) {
            this.setTitle(sheet.getTitle());
        }
        if (sheet.getShiftTop() != null) {
            this.setShiftTop(sheet.getShiftTop());
        }
        if (sheet.getShiftLeft() != null) {
            this.setShiftLeft(sheet.getShiftLeft());
        }
        if (sheet.getSelected() != null) {
            this.setSelected(sheet.getSelected());
        }
        if (sheet.getTabColor() != null) {
            this.setTabColor(sheet.getTabColor());
        }
        if (sheet.getMerges() != null) {
            if (this.merges != null) {
                this.merges.addAll(sheet.getMerges());
            } else {
                this.setMerges(sheet.getMerges());
            }
        }
        if (sheet.getHeaderStyleKey() != null) {
            this.setHeaderStyleKey(sheet.getHeaderStyleKey());
        }
        if (sheet.getMergeRowDataCondition() != null) {
            this.setMergeRowDataCondition(sheet.getMergeRowDataCondition());
        }
        if (sheet.getStyleCellCondition() != null) {
            this.setStyleCellCondition(sheet.getStyleCellCondition());
        }
        if (sheet.getCommentCondition() != null) {
            this.setCommentCondition(sheet.getCommentCondition());
        }
        if (sheet.getSortAndFilter() != null) {
            this.setSortAndFilter(sheet.getSortAndFilter());
        }
        if (sheet.getState() != null) {
            this.setState(sheet.getState());
        }
        if (sheet.getHeaderRowOption() != null) {
            this.setHeaderRowOption(sheet.getHeaderRowOption());
        }
        if (sheet.getProtectionOption() != null) {
            if (this.protectionOption != null) {
                this.protectionOption.putAll(sheet.getProtectionOption());
            } else {
                this.setProtectionOption(sheet.getProtectionOption());
            }
        }
        if (sheet.getHeaderHeight() != null) {
            this.setHeaderHeight(sheet.getHeaderHeight());
        }
        if (sheet.getCheckbox() != null) {
            if (this.checkbox != null) {
                this.checkbox.addAll(sheet.getCheckbox());
            } else {
                this.setCheckbox(sheet.getCheckbox());
            }
        }
        if (sheet.getViewOption() != null) {
            this.setViewOption(sheet.getViewOption());
        }
        if (sheet.getRtl() != null) {
            this.setRtl(sheet.getRtl());
        }
        if (sheet.getPageBreak() != null) {
            this.setPageBreak(sheet.getPageBreak());
        }
        if (sheet.getAsTable() != null) {
            this.setAsTable(sheet.getAsTable());
        }
        if (sheet.getDropDowns() != null) {
            if (this.dropDowns != null) {
                this.dropDowns.addAll(sheet.getDropDowns());
            } else {
                this.setDropDowns(sheet.getDropDowns());
            }
        }
    }

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

    public static class SheetBuilder {
    }

    @AllArgsConstructor
    @Data
    @Builder
    @NoArgsConstructor
    public static class MapSheetDataOption<T> {
        private Method outlineLevel;
        private Method hidden;
        private Method height;

        public Number invokeHeight(Object obj) {
            T castedObject = (T) obj;
            try {
                return (Number) height.invoke(obj);
            } catch (Exception e) {
                return null;
            }
        }

        public Boolean invokeHidden(Object obj) {
            T castedObject = (T) obj;
            try {
                return (Boolean) hidden.invoke(obj);
            } catch (Exception e) {
                return null;
            }
        }

        public Number invokeOutlineLevel(Object obj) {
            T castedObject = (T) obj;
            try {
                return (Number) outlineLevel.invoke(obj);
            } catch (Exception e) {
                return null;
            }
        }

    }

    @NoArgsConstructor
    @AllArgsConstructor
    @Data
    @Builder
    public static class ImageTypes {
        private ImageInput image;
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

        @AllArgsConstructor
        @Builder
        @Data
        @NoArgsConstructor
        public static class Extend {
            private Number cx;
            private Number cy;
        }

        @NoArgsConstructor
        @AllArgsConstructor
        @Data
        @Builder
        public static class Margin {

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
            private final String type;

            Type(String type) {
                this.type = type;
            }

            public String getType() {
                return type;
            }
        }
    }


}
