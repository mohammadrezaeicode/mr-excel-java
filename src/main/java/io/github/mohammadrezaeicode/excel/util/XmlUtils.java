package io.github.mohammadrezaeicode.excel.util;

import java.util.HashMap;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

public class XmlUtils {
    public static String appGenerator(Integer sheetLength,String sheetNameApp) {

        return (
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"" +
                        " xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
                        "<Application>Microsoft Excel</Application>" +
                        "<DocSecurity>0</DocSecurity>" +
                        "<ScaleCrop>false</ScaleCrop>" +
                        "<HeadingPairs>" +
                        "<vt:vector size=\"2\" baseType=\"variant\">" +
                        "<vt:variant>" +
                        "<vt:lpstr>Worksheets</vt:lpstr>" +
                        "</vt:variant>" +
                        "<vt:variant>" +
                        "<vt:i4>" +
                        sheetLength +
                        "</vt:i4>" +
                        "</vt:variant>" +
                        "</vt:vector>" +
                        "</HeadingPairs>" +
                        "<TitlesOfParts>" +
                        "<vt:vector size=\"" +
                        sheetLength +
                        "\" baseType=\"lpstr\">" +
                        " " +
                        sheetNameApp +
                        "" +
                        "</vt:vector>" +
                        "</TitlesOfParts>" +
                        "<Company></Company>" +
                        "<LinksUpToDate>false</LinksUpToDate>" +
                        "<SharedDoc>false</SharedDoc>" +
                        "<HyperlinksChanged>false</HyperlinksChanged>" +
                        "<AppVersion>16.0300</AppVersion>" +
                        "</Properties>"
        );
    }
    public static String contentTypeGenerator(
            String sheetContentType,
            List<Integer> commentId,
            List<String> arrTypes,
            List<String> sheetDrawers,
            List<String> checkboxForm,
            Boolean needCalcChain,
            List<String> tableRef
    ) {
       final HashMap<String,Boolean> typeCheck=new HashMap<>();
       final AtomicInteger index=new AtomicInteger(0);
        return (
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                        "<Default Extension=\"rels\"  ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                        "<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\" />" +
                        "<Default Extension=\"xml\" ContentType=\"application/xml\" />" +
                        "<Override" +
                        " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"" +
                        " PartName=\"/xl/workbook.xml\" />" +
                        "<Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"" +
                        " PartName=\"/xl/styles.xml\" />" +
                        "<Override ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"" +
                        " PartName=\"/xl/theme/theme1.xml\" />" +
                        arrTypes.stream().reduce("",(res, curr) -> {
                            curr = curr.toLowerCase();
        if (typeCheck.containsKey(curr)) {
            return res;
        }
        if (curr == "svg") {
            typeCheck.put("png", true);
            typeCheck.put("svg",true);
            return (
                    res +
                            "<Default Extension=\"png\" ContentType=\"image/png\"/>" +
                            "<Default Extension=\"svg\" ContentType=\"image/svg+xml\"/>"
            );
        } else if (curr == "jpeg" || curr == "jpg") {
            typeCheck.put("jpeg", true);
            typeCheck.put("jpg", true);
            return (
                    res + "<Default Extension=\"" + curr + "\" ContentType=\"image/jpeg\"/>"
            );
        } else {
            typeCheck.put(curr, true);
            return (
                    res +
                            "<Default Extension=\"" +
                            curr +
                            "\" ContentType=\"image/" +
                            curr +
                            "\"/>"
            );
        }
    }) +
                commentId.stream().reduce("",(res, curr) -> {
            return (
                    res +
                            "<Override PartName=\"/xl/comments" +
                            curr +
                            ".xml\"" +
                            " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\" />"
            );
        },(v1,v2)->v1+""+v2) +
                sheetContentType +
                (tableRef.size() > 0
                        ? tableRef.stream().reduce("",(res, cur) -> {
        return (
                res +
                        "<Override PartName=\"/xl/tables/" +
                        cur +
                        "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>"
        );
        })
      : "") +
                "<Override" +
                " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"" +
                " PartName=\"/xl/sharedStrings.xml\" />" +
                (needCalcChain
                        ? "<Override PartName=\"/xl/calcChain.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml\"/>"
                        : "") +
                "<Override PartName=\"/docProps/core.xml\" " +
                " ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\" />" +
                sheetDrawers.stream().reduce("",(res, cu) -> {
            return (
                    res +
                            "<Override PartName=\"/xl/drawings/" +
                            cu +
                            "\"" +
                            " ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\" />"
            );
        }) +
                (checkboxForm.size() > 0
                        ? checkboxForm.stream().reduce("",(res, curr) -> {

        return (
                res +
                        "<Override PartName=\"/xl/ctrlProps/ctrlProp" +
                        (index.incrementAndGet()) +
                        ".xml\" ContentType=\"application/vnd.ms-excel.controlproperties+xml\"/>"
        );
        })
      : "") +
                "<Override PartName=\"/docProps/app.xml\" " +
                " ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\" />" +
                "</Types>"
  );
    }

}
