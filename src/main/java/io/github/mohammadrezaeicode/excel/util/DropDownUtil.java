package io.github.mohammadrezaeicode.excel.util;

import io.github.mohammadrezaeicode.excel.model.DropDown;

import java.util.List;

public class DropDownUtil {
    public static String generateDropDown(List<DropDown> dropDowns) {
        if (dropDowns == null) {
            return "";
        }
        int dropDownLength = dropDowns.size();
        String result = "<dataValidations>";
        for (int index = 0; index < dropDownLength; index++) {
            var dropDown = dropDowns.get(index);
            var sqref = dropDown.getForCell().stream().reduce("", (resultValue, current) -> resultValue + " " + current);
            var options = String.join(",", dropDown.getOption());
            result +=
                    "<dataValidation type=\"list\" allowBlank=\"1\" showErrorMessage=\"1\" sqref=\"" +
                            sqref.trim() +
                            "\">" +
                            "<formula1>&quot;" +
                            options +
                            "&quot;</formula1></dataValidation>";
        }
        result += "</dataValidations>";
        return result;
    }

}
