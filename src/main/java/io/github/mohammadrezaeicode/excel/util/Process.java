package io.github.mohammadrezaeicode.excel.util;

import io.github.mohammadrezaeicode.excel.annotation.ExcludeProperty;
import io.github.mohammadrezaeicode.excel.annotation.GetterMethod;
import io.github.mohammadrezaeicode.excel.annotation.Options;
import io.github.mohammadrezaeicode.excel.model.Header;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

public class Process {
    public static List<Header> findAnnotation(Class headerClass) throws NoSuchMethodException, SecurityException {
        Class<? extends Object> cls = headerClass;
        List<Field> presentField = new ArrayList<>();
        List<String> textName = new ArrayList<>();
        List<String> methodName = new ArrayList<>();
        Field[] fields = cls.getDeclaredFields();
//        Map<String,String> fieldVsMethod=new HashMap<>();
        for (Field field : fields) {
            if (!field.isAnnotationPresent(ExcludeProperty.class)) {

                String getterMethodName = "";
                String fieldName = field.getName();
                if (field.isAnnotationPresent(GetterMethod.class)) {
                    GetterMethod getterMethod = field.getAnnotation(GetterMethod.class);
                    getterMethodName = getterMethod.method();
//                    methodName.add(getterMethodName);
//                    fieldVsMethod.put(field.getName(),getterMethodName);
                } else {
                    if (fieldName.length() > 1) {
                        getterMethodName = "get" + (String.valueOf(fieldName.charAt(0))).toUpperCase() + fieldName.substring(1);
                    } else {
                        getterMethodName = "get" + fieldName.toUpperCase();
                    }
                }
                String title = field.getName();
                if (field.isAnnotationPresent(Options.class)) {
                    Options options = field.getAnnotation(Options.class);
                    String anTitle = new String(options.title().getBytes(), StandardCharsets.UTF_8);
                    if (anTitle.length() > 0) {
                        title = anTitle;
                    }
                }
                textName.add(title);
                methodName.add(getterMethodName);
                presentField.add(field);
//                fieldVsMethod.put(getterMethodName,fieldName);
            } 
        }
        List<Header> header = new ArrayList<>();
        int index = 0;
        for (String name : methodName) {

        Method m = cls.getMethod(name);
        header.add(new Header(presentField.get(index).getName(), textName.get(index), m));
//                String result= (String) m.invoke(obj);
//                System.out.println(result);

            index++;
        }
        return header;
    }

}
