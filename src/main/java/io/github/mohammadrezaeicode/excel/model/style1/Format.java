package io.github.mohammadrezaeicode.excel.model.style1;

public class Format {
    private int key;
    private String value;

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
