package io.github.mohammadrezaeicode.excel.util;

import io.github.mohammadrezaeicode.excel.model.Comment;
import io.github.mohammadrezaeicode.excel.model.CommentConverter;

import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class CommentUtil {

    public static CommentConverter commentConvertor(
            Comment commentValue,
            Map<String,String> mapStyle,
    String commentStyle
) {
        Boolean hasAuthor = false;
        String[] commentStr;
        String author=commentValue.getAuthor();
            if ( author.length()>0) {
                hasAuthor = true;
            }
            String styleId=commentValue.getStyleId();
            if (styleId.length()>0) {
//                let styleCom = mapStyle[commentValue.styleId];
                if (mapStyle.containsKey(styleId)) {
                    commentStyle = mapStyle.get(styleId);
                }
            }
            String comment=commentValue.getComment();
            commentStr =
                    comment.length()>0
                    ? splitBaseOnbreackline(comment)
                    : new String[]{""};
            List<String> commentStrResult= Arrays.stream(commentStr).collect(Collectors.toList());
        if (hasAuthor) {
            commentStrResult.add(0,author + ":");
        }

        return new CommentConverter(hasAuthor,author,commentStyle,commentStrResult.toArray(new String[0]));
    }
    public static CommentConverter commentConvertor(
            String commentValue,
            Map<String,String> mapStyle,
            String commentStyle
    ) {
        Boolean hasAuthor = false;
        String[] commentStr=commentValue!=null&&commentValue.length()>0 ? splitBaseOnbreackline(commentValue) : new String[]{""};;
        String author=null;

        List<String> commentStrResult= Arrays.stream(commentStr).collect(Collectors.toList());
        if (hasAuthor) {
            commentStrResult.add(0,author + ":");
        }
        return new CommentConverter(hasAuthor,author,commentStyle,commentStrResult.toArray(new String[0]));
    }
    public static String[] splitBaseOnbreackline(String str) {


        // Split the string on \n or \r characters
//        var separateLines = str.split(/\r?\n|\r|\n/g);
        return str.split("\r?\n|\r|\n");
    }

    public static String generateCommentTag(
            String ref,
            String[] comment,
            String style,
            int auIndex
    ) {
        String result =
                "<comment ref=\"" +
                        ref +
                        "\" authorId=\"" +
                        Math.max(0, auIndex - 1) +
                        "\" shapeId=\"0\">" +
                        " <text>";
        ;
        String bac = "";
        int length=comment.length;
        for (int vindex = 0; vindex < length; vindex++) {
            String v=comment[vindex];
            String pr = "";

            if (v.length() == 0) {
                bac += "\n";
                continue;
            }
            if (vindex > 0) {
                pr = "xml:space=\"preserve\"";
                bac += "\n";
            }
            result += "<r>" + style + "<t " + pr + " >" + bac + v + "</t></r>";
            bac = "";
        }
        result += "</text></comment>";
        return result;
    }
    public final static String defaultCellCommentStyle =
            "<rPr>" +
            " <b />" +
            " <sz val=\"9\" />" +
            " <color rgb=\"000000\" />" +
            " <rFont val=\"Tahoma\" />" +
            "</rPr>";
    ;

}
