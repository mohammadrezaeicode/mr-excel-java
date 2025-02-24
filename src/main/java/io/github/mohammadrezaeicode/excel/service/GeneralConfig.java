package io.github.mohammadrezaeicode.excel.service;


import io.github.mohammadrezaeicode.excel.model.ExcelTable;

import java.util.HashMap;
import java.util.Map;

public class GeneralConfig {

    private static final Map<String, ExcelTable> configMap = new HashMap<>();

    public static ExcelTable applyConfig(String key, ExcelTable data) {
        if (configMap.containsKey(key)) {
            var configs = configMap.get(key);
            var styles = data.getStyles();
            data.applyExcelTableOption(configs);
            if (styles != null) {
                data.getStyles().putAll(styles);
            }
            if (data.getSheet() != null && data.getSheet().size() > 0)
                data.getSheet().forEach(sheet -> {
                    sheet.applySheet(data.getSheet().get(0));
                });
        }
        return data;
    }

    public static void addConfig(String key, ExcelTable data) {
        configMap.put(key, data);
    }

}
