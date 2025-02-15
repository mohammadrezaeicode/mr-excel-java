package io.github.mohammadrezaeicode.excel.util;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

public class ColorUtils {
    static String removeSpace(String str) {
        return str.replaceAll(" ", "");
    }

    private static String valueToHex(int c) {
        var hex = Integer.toString(c,16);
        return hex.length() == 1 ? "0" + hex : hex;
    }
    public static Optional<String> rgbToHex(String rgb){
        rgb = removeSpace(rgb);
        String[] spResult =
                rgb.indexOf("rgba") >= 0
                        ? rgb.substring(5, rgb.length() - 1).split(",")
                        : rgb.substring(4, rgb.length() - 1).split(",");
        List<Integer> parse=new ArrayList<>();
        boolean validate = true;
        for (String s : spResult) {
            try {
                parse.add(Integer.parseInt(s));
            }catch ( NumberFormatException exception){
                validate=false;
                break;
            }
        }
        if ((spResult.length == 4 && spResult[3] == "0") || !validate) {
            return Optional.ofNullable(null);
        }
        return Optional.of((
                valueToHex(parse.get(0)) +
                        valueToHex(parse.get(1)) +
                        valueToHex(parse.get(2))
        ).toUpperCase());
    }

    public static String convertToHex(
            String fgConvertor
    ) {
        if(fgConvertor==null){
            return "";
        }
        if (fgConvertor.indexOf("rgb") >= 0) {
    Optional<String> rgb = rgbToHex(fgConvertor);
            fgConvertor = rgb.isPresent() ? rgb.get() : "";
        }
        return fgConvertor.replace("#", "");
    }

}
