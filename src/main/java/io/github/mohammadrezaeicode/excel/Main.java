package io.github.mohammadrezaeicode.excel;

import io.github.mohammadrezaeicode.excel.function.GenerateExcel;
import io.github.mohammadrezaeicode.excel.model.ExcelTableOption;
import io.github.mohammadrezaeicode.excel.model.Result;
import io.github.mohammadrezaeicode.excel.model.test.OB;
import io.github.mohammadrezaeicode.excel.model.types.GenerateType;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.util.Arrays;
import java.util.Optional;

public class  Main {
    public static void main(String[] args) throws IOException, InvocationTargetException, IllegalAccessException {

//        ExcelTable1 excelTable1= ExcelTable1.builder()
//                .sheet(Arrays.asList(
//                        Sheet1.builder()
//                                .data(Arrays.asList(new OB("er","gdh","gfdh")))
//                                .headers(
//                                        Process.findAnnotation(OB.class)
//                                )
//                                .build()
//                ))
//                .build();
//        System.out.println(Process.findAnnotation(OB.class));
        try {
//            var map=GenerateExcel.generateExcel(excelTable1,"");
//            System.out.println(map.size()+"***********");
//            FileUtils.writeFile("new-test-x", );
//            System.out.println( );
            Result<InputStream> data= GenerateExcel.generateHeaderAndGenerateExcel(Arrays.asList(
                    new OB("er","gdh","gfdh"),
                    new OB("er1","gdh2","gfd4h"),
                    new OB("er2","gdh3","gfdh5")
            ), ExcelTableOption.eBuilder()
                            .fileName("this is by name")
                    .notSave(false)
                    .generateType(GenerateType.INPUT_STREAM)
                    .build(),null,null);
//            saveZipToFile(ConverterUtil.inputStreamToByteArray(data.getFileData()),"INPUT_STREAM.xlsx");
        }catch (Exception ex){
            ex.printStackTrace();
            System.err.println(ex);
        }
//        FileUtils.base64();
//        Map<String, StyleBody> map=new HashMap<>();
//        map.put("head",new StyleBody(null,null,null,null,null,null,null,null,null,null,null,null,
//                "9A3B3B"));
//        Map<String, StyleBody> map2=new HashMap<>();
//        map2.put("head",new StyleBody(null,null,null,null,null,null,null,null,null,null,null,"FFF2D8",
//                "432143"));
//
//        GeneralConfig.setConfig(GeneralConfig.ConfigType.Root,"test", GeneralConfig.FieldName.styles,map);
//        GeneralConfig.setConfig(GeneralConfig.ConfigType.Sheets,"test", GeneralConfig.FieldName.headerStyleKey,"head");
//        GeneralConfig.setConfig(GeneralConfig.ConfigType.Root,"test", GeneralConfig.FieldName.styles,map2);
//
////        FileUtils.writeFile("");
////        var ob=new OB("er","gdh","gfdh");
////        Process.findAnnotation(ob);
////        Process.processList(Arrays.asList(new OB("test","test1","test23")));
//        ExcelTable table= Process.generateExcelTable(Arrays.asList(new OB("test","test1","test23")));
//        GeneralConfig.applyConfig("test",table);
////        table.setStyles(map);
////        table.getSheet().get(0).setHeaderStyleKey("head");
//        Map<String,String> mapr= GenerateExcel.generateExcel(table);
//        FileUtils.writeFile("aftermissup",mapr);

    }


    // Method to save the ZIP byte array to a file
    private static void saveZipToFile(byte[] zipData, String filePath) {
        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            fos.write(zipData); // Write the byte array to the file
            System.out.println("ZIP file saved successfully to " + filePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}