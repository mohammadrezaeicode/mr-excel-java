package io.github.mohammadrezaeicode.excel.model.style;

import io.github.mohammadrezaeicode.excel.util.ColorUtils;
import io.github.mohammadrezaeicode.excel.util.FormatUtils;
import io.github.mohammadrezaeicode.excel.util.Utils;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.*;

@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class StyleMapper {
    private StyleDetail conditionalFormatting;
    private Map<String, String> commentSyntax;
    private Map<String, String> headerFooterStyle;
    private Map<String, Number> conditionalFormattingStyle;
    private StyleDetail format;
    private StyleDetail border;
    private StyleDetail fill;
    private StyleDetail font;
    private StyleDetail cell;
    private Boolean addConditionalFormatting;
    public StyleMapper(Map<String, StyleBody> styleBody, Boolean addConditionalFormatting) {
        conditionalFormatting = new StyleDetail(
                addConditionalFormatting ? 1 : 0, "<dxf><font><color rgb=\"FF9C0006\"/></font><fill> <patternFill> <bgColor rgb=\"FFFFC7CE\"/></patternFill></fill></dxf>"
        );
        commentSyntax = new HashMap<>();
        format = new StyleDetail(0, "");
        border = new StyleDetail(1, "");
        fill = new StyleDetail(2, "");
        font = new StyleDetail(2, "");
        cell = new StyleDetail(2, "");
        headerFooterStyle = new HashMap<>();
        conditionalFormattingStyle = new HashMap<>();
        this.addConditionalFormatting = addConditionalFormatting;
        this.fillFields(styleBody);
    }

    private void fillFields(Map<String, StyleBody> styles) {
        Set<String> styleKeys = styles.keySet();
        Iterator<String> styleIterator = styleKeys.iterator();
        var res = this;
        while (styleIterator.hasNext()) {
            String cur = styleIterator.next();
            var styleObject = styles.get(cur);
            var styleTypeCheck = Optional.ofNullable(styleObject.getType()).isPresent();
            if (styleTypeCheck
                    &&
                    (styleObject.getType() == StyleBody.StyleType.HEADER_FOOTER)
            ) {
                String result = "";
                String fontProcessLeft = "-";
                String fontProcessRight = "Regular";
                if (styleObject.getFontFamily() != null) {
                    fontProcessLeft = styleObject.getFontFamily();
                }
                if (Optional.ofNullable(styleObject.getBold()).orElse(false)) {
                    fontProcessRight = "Bold";
                }
                if (Optional.ofNullable(styleObject.getItalic()).orElse(false)) {
                    if (fontProcessRight == "Regular") {
                        fontProcessRight = "";
                    }
                    fontProcessRight += "Italic";
                }
                if (fontProcessLeft != "-" || fontProcessRight != "Regular") {
                    result =
                            "&amp;" + "\"" + fontProcessLeft + "," + fontProcessRight + "\"";
                }
                if (styleObject.getSize() != null) {
                    result += "&amp;" + styleObject.getSize();
                }
                if (Optional.ofNullable(styleObject.getDoubleUnderline()).orElse(false)) {
                    result += "&amp;E";
                } else if (Optional.ofNullable(styleObject.getUnderline()).orElse(false)) {
                    result += "&amp;U";
                }
                if (styleObject.getColor() != null) {
                    var convertedColor = ColorUtils.convertToHex(styleObject.getColor());
                    if (convertedColor.length() > 0) {
                        result += "&amp;K" + convertedColor.toUpperCase();
                    }
                }
                headerFooterStyle.put(cur, result);
                continue;
            }
            if (
                    addConditionalFormatting &&
                            styleTypeCheck &&
                            (styleObject.getType() == StyleBody.StyleType.CONDITIONAL_FORMATTING)
            ) {
                conditionalFormattingStyle.put(cur, this.conditionalFormatting.getCount());
                // TODO data maybe not exists
                var color = ColorUtils.convertToHex(styleObject.getColor());
                var bgColor = ColorUtils.convertToHex(styleObject.getBackgroundColor());
                this.conditionalFormatting.value +=
                        "<dxf><font><color rgb=\"" +
                                color +
                                "\"/></font><fill> <patternFill> <bgColor rgb=\"" +
                                bgColor +
                                "\"/></patternFill></fill></dxf>";
                this.conditionalFormatting.count++;
                continue;
            }
            var fillIndex = 0;
            var fontIndex = 0;
            var borderIndex = 0;
            var formatIndex = 0;
            if (styleObject.getBackgroundColor() != null) {
                var fgConvertor = ColorUtils.convertToHex(styleObject.getBackgroundColor());
                fillIndex = this.fill.count;
                this.fill.count++;
                this.fill.value =
                        this.fill.value +
                                "<fill>" +
                                "<patternFill patternType=\"solid\">" +
                                (fgConvertor.length() > 0
                                        ? "<fgColor rgb=\"" + fgConvertor.replace("#", "") + "\" />"
                                        : "") +
                                "</patternFill>" +
                                "</fill>";
            }
            if (
                    styleObject.colorIsExist() ||
                            styleObject.formatIsExist() ||
                            styleObject.getSize() != null ||
                            styleObject.boldIsExist() ||
                            styleObject.italicIsExist() ||
                            styleObject.underlineIsExist() ||
                            styleObject.doubleUnderlineIsExist()
            ) {
                var colors = ColorUtils.convertToHex(styleObject.getColor());

                fontIndex = res.font.count;
                res.font.count++;
                res.font.value =
                        res.font.value +
                                "<font>" +
                                (styleObject.boldIsExist() ? "<b/>" : "") +
                                (styleObject.italicIsExist() ? "<i />" : "") +
                                (styleObject.underlineIsExist() || styleObject.doubleUnderlineIsExist()
                                        ? "<u " +
                                        (styleObject.doubleUnderlineIsExist() ? " val=\"double\" " : "") +
                                        "/>"
                                        : "") +
                                (styleObject.sizeIsExist() ? "<sz val=\"" + styleObject.getSize() + "\" />" : "") +
                                (colors.length() > 0 ? "<color rgb=\"" + colors.replace("#", "") + "\" />" : "") +
                                (styleObject.fontFamilyIsExist()
                                        ? "<name val=\"" + styleObject.getFontFamily() + "\" />"
                                        : "") +
                                "</font>";
                this.commentSyntax.put(cur,
                        "<rPr>" +
                                (styleObject.boldIsExist() ? "<b/>" : "") +
                                (styleObject.italicIsExist() ? "<i/>" : "") +
                                (styleObject.underlineIsExist() || styleObject.doubleUnderlineIsExist()
                                        ? "<u " +
                                        (styleObject.doubleUnderlineIsExist() ? "val=\"double\" " : "") +
                                        "/>"
                                        : "") +
                                "<sz val=\"" +
                                (styleObject.sizeIsExist() ? styleObject.getSize() : "9") +
                                "\" />" +
                                (colors.length() > 0 ? "<color rgb=\"" + colors.replace("#", "") + "\" />" : "") +
                                "<rFont val=\"" +
                                (styleObject.fontFamilyIsExist() ? styleObject.getFontFamily() : "Tahoma") +
                                "\" />" +
                                "</rPr>");
            }
            var endPart = "/>";
            if (styleObject.getAlignment() != null) {
                var alignment = styleObject.getAlignment();
                if (Utils.booleanCheck(alignment.getRtl())) {
                    alignment.setReadingOrder("2");
                }
                if (Utils.booleanCheck(alignment.getLtr())) {
                    alignment.setReadingOrder("1");
                }
                endPart =
                        " applyAlignment=\"1\">" +
                                "<alignment " +
                                alignment.generateAligmentString() +
                                " />" +
                                "</xf>";
            }
            var borderObj = styleObject.getBorder();
            String borderStr = "";
            if (borderObj != null) {
                borderStr = borderObj.getBorderStyle();
                borderIndex = this.border.count;
                this.border.count++;
                this.border.value +=
                        "<border>" + borderStr + "<diagonal />" + "</border>";
            }
            if (styleObject.formatIsExist()) {
                String formatKey = styleObject.getFormat();

                if (FormatUtils.formatMap.containsKey(formatKey)) {
                    Format format = FormatUtils.formatMap.get(formatKey);

                    formatIndex = format.getKey();
                    if (format.getValue() != null) {
                        this.format.count++;
                        this.format.value += format.getValue();
                    }
                }
            }
            this.cell.value =
                    this.cell.value +
                            "<xf numFmtId=\"" +
                            formatIndex +
                            "\"" +
                            " fontId=\"" +
                            fontIndex +
                            "\" fillId=\"" +
                            fillIndex +
                            "\" borderId=\"" +
                            borderIndex +
                            "\" xfId=\"0\"" +
                            (borderIndex > 0 ? " applyBorder=\"1\" " : "") +
                            (fillIndex > 0 ? " applyFill=\"1\" " : "") +
                            (fontIndex >= 0 ? " applyFont=\"1\" " : "") +
                            (formatIndex > 0 ? " applyNumberFormat=\"1\" " : "") +
                            endPart;
            styleObject.setIndex(this.cell.count);
            this.cell.count++;
        }
    }

    @Builder
    @AllArgsConstructor
    @Data
    @NoArgsConstructor
    public static class StyleDetail {
        private Integer count;
        private String value;
    }
}
