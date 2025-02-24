package io.github.mohammadrezaeicode.excel.model.style;

public class Format {
    private final int key;
    private final String value;

    public Format(int key, String value) {
        this.key = key;
        this.value = value;
    }

    public int getKey() {
        return key;
    }

    public String getValue() {
        return value;
    }
}
