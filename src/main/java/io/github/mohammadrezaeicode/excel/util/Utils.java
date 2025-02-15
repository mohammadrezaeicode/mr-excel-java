package io.github.mohammadrezaeicode.excel.util;


import java.util.Optional;

public class Utils {
    public static  Boolean booleanCheck(Boolean value){
        return Optional.ofNullable(value).orElse(false);
    }

}
