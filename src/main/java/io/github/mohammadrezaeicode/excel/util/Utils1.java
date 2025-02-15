package io.github.mohammadrezaeicode.excel.util;

import io.github.mohammadrezaeicode.excel.model.style1.StyleMapper;

public class Utils1 {
    public static String styleGenerator(StyleMapper styles,Boolean addCF) {
        return (
               "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"+
                        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"" +
                        " xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"" +
                        " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">" +
                        (styles.getFormat().getCount() > 0
                                ? "<numFmts count=\"" +
                                styles.getFormat().getCount() +
                                "\">" +
                                styles.getFormat().getValue() +
                                "</numFmts>"
                                : "") +
                        "<fonts count=\"" +
                        styles.getFont().getCount() +
                        "\">" +
                        "<font>" +
                        "<sz val=\"11\" />" +
                        "<color theme=\"1\" />" +
                        "<name val=\"Calibri\" />" +
                        "<family val=\"2\" />" +
                        "<scheme val=\"minor\" />" +
                        "</font>" +
                        "<font>" +
                        "<sz val=\"11\" />" +
                        "<color rgb=\"FFFF0000\" />" +
                        "<name val=\"Calibri\" />" +
                        "<family val=\"2\" />" +
                        "<scheme val=\"minor\" />" +
                        "</font>" +
                        styles.getFont().getValue() +
                        "</fonts>" +
                        "<fills count=\"" +
                        styles.getFill().getCount() +
                        "\">" +
                        "<fill>" +
                        "<patternFill patternType=\"none\" />" +
                        "</fill>" +
                        "<fill>" +
                        "<patternFill patternType=\"lightGray\" />" +
                        "</fill>" +
                        styles.getFill().getValue() +
                        "</fills>" +
                        "<borders count=\"" +
                        styles.getBorder().getCount() +
                        "\">" +
                        "<border />" +
                        styles.getBorder().getValue() +
                        "</borders>" +
                        "<cellStyleXfs count=\"1\">" +
                        "<xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" applyAlignment=\"1\" applyFont=\"1\" />" +
                        "</cellStyleXfs>" +
                        "<cellXfs count=\"" +
                        styles.getCell().getCount() +
                        "\">" +
                        "<xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\"" +
                        " applyFont=\"1\">" +
                        "<alignment readingOrder=\"0\" shrinkToFit=\"0\" vertical=\"bottom\" wrapText=\"0\" />" +
                        "</xf>" +
                        "<xf borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\"" +
                        " applyFont=\"1\">" +
                        "<alignment readingOrder=\"0\" />" +
                        "</xf>" +
                        styles.getCell().getValue() +
                        "</cellXfs>" +
                        "<cellStyles count=\"1\">" +
                        "<cellStyle xfId=\"0\" name=\"Normal\" builtinId=\"0\" />" +
                        "</cellStyles> " +
                        (addCF
                                ? "<dxfs count=\"" +
                                styles.getConditionalFormatting().getCount() +
                                "\" >" +
                                styles.getConditionalFormatting().getValue() +
                                "</dxfs>"
                                : "<dxfs count=\"0\" />") +
                        "</styleSheet>"
        );
    }
}
