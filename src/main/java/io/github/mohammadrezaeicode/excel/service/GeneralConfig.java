package io.github.mohammadrezaeicode.excel.service;


import java.util.*;

public class GeneralConfig {
//    public enum ConfigType{
//        Root,Sheets
//    }
//    public enum FieldName{
//        styles(ConfigType.Root),headerStyleKey(ConfigType.Sheets);
//        private final ConfigType level;
//
//        FieldName(ConfigType level) {
//            this.level = level;
//        }
//
//        public ConfigType getLevel() {
//            return level;
//        }
//    }
//    private static class FieldSet{
//        private final ConfigType type;
//        private final FieldName fieldName;
//
//        public FieldSet(ConfigType type, FieldName fieldName) {
//            this.type = type;
//            this.fieldName = fieldName;
//        }
//
//        @Override
//        public boolean equals(Object o) {
//            if (this == o) return true;
//            if (o == null || getClass() != o.getClass()) return false;
//            FieldSet fieldSet = (FieldSet) o;
//            return type == fieldSet.getType() && fieldName == fieldSet.getFieldName();
//        }
//
//        @Override
//        public int hashCode() {
//            return Objects.hash(type, fieldName);
//        }
//
//        public ConfigType getType() {
//            return type;
//        }
//
//        public FieldName getFieldName() {
//            return fieldName;
//        }
//    }
////    private static ExcelTable excelTable=new ExcelTable(Arrays.asList());
//    private static final Map<String,Map<FieldSet,Object>> configMap=new HashMap<>();
//    public static synchronized void setConfig(ConfigType configType,String key,FieldName fieldName,Object value){
//        if(configType==fieldName.getLevel()){
//            checkAndSet(key,fieldName,value);
//        }
//    }
//    public static void checkAndSet(String key,FieldName fieldName,Object value){
//
//        FieldSet set=new FieldSet(fieldName.getLevel(),fieldName);
//        Map<FieldSet,Object> objectMap;
//        if(configMap.containsKey(key)){
//            objectMap=configMap.get(key);
//        }else{
//            objectMap=new HashMap<>();
//        }
//        objectMap.put(set,value);
//        configMap.put(key,objectMap);
//    }
//    public static void applyConfig(String key,ExcelTable excelTable){
//        if(configMap.containsKey(key)){
//            Map<FieldSet,Object> objectMap=configMap.get(key);
//            Set<FieldSet> fieldSets=objectMap.keySet();
//            Iterator<FieldSet> fieldSetIterator=fieldSets.iterator();
//            while (fieldSetIterator.hasNext()){
//                FieldSet fieldSet=fieldSetIterator.next();
//                Object ob=objectMap.get(fieldSet);
//                if(fieldSet.getType()==ConfigType.Root){
//                    setRootData(fieldSet,ob,excelTable);
//                } else if (fieldSet.getType()==ConfigType.Sheets) {
//                    setSheetsData(fieldSet,ob,excelTable);
//                }
//
//            }
//
//        }else{
//
//        }
//    }
//
//    private static void setSheetsData(FieldSet fieldSet, Object ob, ExcelTable excelTable) {
//        List<Sheet> sheets=excelTable.getSheet();
//        sheets.forEach(v->{
//            if(fieldSet.getFieldName()==FieldName.headerStyleKey){
//                v.setHeaderStyleKey((String) ob);
//            }
//        });
//    }
//
//    private static void setRootData(FieldSet fieldSet, Object ob, ExcelTable excelTable) {
//        if(fieldSet.getFieldName()==FieldName.styles){
//            excelTable.setStyles((Map<String, StyleBody>) ob);
//        }
//    }
}
