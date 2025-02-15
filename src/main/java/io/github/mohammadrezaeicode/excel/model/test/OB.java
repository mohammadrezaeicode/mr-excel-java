package io.github.mohammadrezaeicode.excel.model.test;

import io.github.mohammadrezaeicode.excel.annotation.ExcludeProperty;
import io.github.mohammadrezaeicode.excel.annotation.GetterMethod;
import io.github.mohammadrezaeicode.excel.annotation.Options;
import io.github.mohammadrezaeicode.excel.model.row.RowOption;

public class OB extends RowOption {
    @Override
    public Integer getOutline() {
        return null;
    }

    @Override
    public Boolean isHidden() {
        return null;
    }

    public OB(String t, String t1, String t2) {
        this.t = t;
        this.t1 = t1;
        this.t2 = t2;
    }

    private String t;

    @Override
    public Integer getHeight() {
        return 20;
    }

    @ExcludeProperty
    private String t1;
    @Options(title="title of column")
    @GetterMethod(method = "getBfT2")
    private String t2;

    public String getT() {
        return t;
    }

    public String getT1() {
        return t1;
    }

    public String getBfT2() {
        return t2;
    }
}
