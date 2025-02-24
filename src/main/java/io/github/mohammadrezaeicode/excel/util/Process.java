package io.github.mohammadrezaeicode.excel.util;

import io.github.mohammadrezaeicode.excel.annotation.ExcludeProperty;
import io.github.mohammadrezaeicode.excel.annotation.GetterMethod;
import io.github.mohammadrezaeicode.excel.annotation.Options;
import io.github.mohammadrezaeicode.excel.model.Header;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Process {
    public static List<Header> findAnnotation(Class headerClass) throws NoSuchMethodException, SecurityException {
        Class<? extends Object> cls = headerClass;
        List<Field> presentField = new ArrayList<>();
        List<String> textName = new ArrayList<>();
        List<String> methodName = new ArrayList<>();
        Field[] fields = cls.getDeclaredFields();
        Map<Integer, List<Header>> orderMap = new HashMap<>();
        for (Field field : fields) {
            if (!field.isAnnotationPresent(ExcludeProperty.class)) {

                String getterMethodName = "";
                String fieldName = field.getName();
                if (field.isAnnotationPresent(GetterMethod.class)) {
                    GetterMethod getterMethod = field.getAnnotation(GetterMethod.class);
                    getterMethodName = getterMethod.method();
                } else {
                    if (fieldName.length() > 1) {
                        getterMethodName = "get" + (String.valueOf(fieldName.charAt(0))).toUpperCase() + fieldName.substring(1);
                    } else {
                        getterMethodName = "get" + fieldName.toUpperCase();
                    }
                }
                String title = field.getName();
                Integer methodOrder = -1;
                if (field.isAnnotationPresent(Options.class)) {
                    Options options = field.getAnnotation(Options.class);
                    String anTitle = new String(options.title().getBytes(), StandardCharsets.UTF_8);
                    if (anTitle.length() > 0) {
                        title = anTitle;
                    }
                    methodOrder = options.order();
                }
                textName.add(title);
                methodName.add(getterMethodName);
                presentField.add(field);
                Method m = cls.getMethod(getterMethodName);
                Header headerObj = new Header(field.getName(), title, m);
                if (!orderMap.containsKey(methodOrder)) {
                    orderMap.put(methodOrder, new ArrayList<>());
                }
                var ordersArray = orderMap.get(methodOrder);
                ordersArray.add(headerObj);
                orderMap.put(methodOrder, ordersArray);
            }
        }
        var orderKeys = new ArrayList<>(orderMap.keySet());
        orderKeys.sort((a, b) -> a > b ? 1 : -1);
        List<Header> sortHeader = new ArrayList<>();
        for (Integer order : orderKeys) {
            var list = orderMap.get(order);
            sortHeader.addAll(list);
        }
        return sortHeader;
    }

}
