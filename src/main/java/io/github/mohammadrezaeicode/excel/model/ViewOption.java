package io.github.mohammadrezaeicode.excel.model;

import lombok.*;

@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class ViewOption {


    public enum Type{
        PAGE_LAYOUT("pageLayout") ,
        PAGE_BREAK_PREVIEW("pageBreakPreview");
        private String type;

        Type(String type) {
            this.type = type;
        }

        public String getType() {
            return type;
        }
    }
    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    @Builder
    public static class FrozenOption{

        private Type type;
        private Number index;
        private Number r;
        private Number c;


        public enum Type{
            ROW("ROW"),
            COLUMN("COLUMN"),
            BOTH("BOTH");
            private String option;

            Type(String option) {
                this.option = option;
            }

            public String getOption() {
                return option;
            }
        }
    }
    @Data
    @AllArgsConstructor
    @Builder
    @NoArgsConstructor
    public static class  SplitOption{
        private Type type;
        private ViewStart startAt;
        private Number split;
        private Number x;
        private Number y;
        @Data
        @AllArgsConstructor
        @NoArgsConstructor
        @Builder
        public static class ViewStart {
            private String t;
            private String b;
            private String r;
            private String l;
            private String one;
            private String two;
        }

        public Number getYOrSplit() {
            if(y!=null){
                return y;
            }else{
                return split;
            }
        }
        public Number getXOrSplit() {
            if(x!=null){
                return x;
            }else{
                return split;
            }
        }

        public enum Type{
            VERTICAL("VERTICAL")  ,HORIZONTAL("HORIZONTAL"),
            BOTH("BOTH");;
            private String type;

            Type(String type) {
                this.type = type;
            }

            public String getType() {
                return type;
            }

            @Override
            public String toString() {
                return type;
            }
        }

    }
   private Type type;
    private Boolean hideGrid;
    private Boolean hideHeadlines;
    private Boolean hideRuler;
    private FrozenOption frozenOption;
    private SplitOption splitOption;
}
