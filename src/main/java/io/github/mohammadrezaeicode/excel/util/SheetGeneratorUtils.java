package io.github.mohammadrezaeicode.excel.util;


import io.github.mohammadrezaeicode.excel.model.Header;
import io.github.mohammadrezaeicode.excel.model.Sheet1;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

public class SheetGeneratorUtils {
    public static  <T> Sheet1 generateSheet(List<T> data, Class headerClass, Function<List<Header>,List<Header>> applyHeaderOptionFunction) throws NoSuchMethodException {
        List<Header> headers=new ArrayList<>();
        if(data.size()>0){
            var firstRecord=data.get(0);
            if(headerClass!=null) {
                if (firstRecord.getClass().equals(headerClass)) {
                    System.err.println("Class not match!? data vs headerOption");
                }
            }else{
                System.err.println("headerOption is null");
            }
            headers.addAll(Process.findAnnotation(firstRecord.getClass()));
            if(applyHeaderOptionFunction!=null){
                var funcResult=applyHeaderOptionFunction.apply(headers);
                if(funcResult!=null){
                    headers.clear();
                    headers.addAll(headers);
                }
            }
        }
        return Sheet1.builder().headers(headers).data((List<Object>) data).build();
    }
}
