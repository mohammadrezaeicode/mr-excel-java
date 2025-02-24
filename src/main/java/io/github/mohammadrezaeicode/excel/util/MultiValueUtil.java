package io.github.mohammadrezaeicode.excel.util;

import io.github.mohammadrezaeicode.excel.model.MultiStyleValue;

import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;

public class MultiValueUtil {
    public static String generateMultiStyleByArray(List<MultiStyleValue> values, Map<String, String> styles, String defStyleId) {
        final AtomicReference<String> result = new AtomicReference<>();
        values.stream().forEach((value) -> {
            value.setValue(specialCharacterConverter(String.valueOf(value.getValue())));

            result.getAndUpdate(curr -> curr +
                    "<r>" +
                    (value.getStyleId() != null && styles.get(value.getStyleId()) != null
                            ? styles.get(value.getStyleId())
                            : styles.get(defStyleId)) +
                    "<t xml:space=\"preserve\">" +
                    value.getValue() +
                    "</t>" +
                    "</r>");
        });
        return "<si>" + result + "</si>";
    }

    public static String specialCharacterConverter(String str) {

        return str.replaceAll("&", "&amp;")
                .replaceAll("<", "&lt;")
                .replaceAll(">", "&gt;");
    }

}
