package io.github.mohammadrezaeicode.excel.util;


import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Optional;

public class Utils {
    public static  Boolean booleanCheck(Boolean value){
        return Optional.ofNullable(value).orElse(false);
    }
    public static void saveZipToFile(byte[] zipData, String filePath) {
        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            fos.write(zipData);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public static boolean isValidDate(String dateStr, String format) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(format);
        try {
            LocalDate date = LocalDate.parse(dateStr, formatter);
            return true;
        } catch (DateTimeParseException e) {
            return false;
        }
    }

}
