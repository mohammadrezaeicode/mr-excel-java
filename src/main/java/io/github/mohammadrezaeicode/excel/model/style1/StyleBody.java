package io.github.mohammadrezaeicode.excel.model.style1;

import io.github.mohammadrezaeicode.excel.util.ColorUtils;
import io.github.mohammadrezaeicode.excel.util.Utils;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Optional;

@AllArgsConstructor
@NoArgsConstructor
@Builder
@Data
public class StyleBody {
    public enum StyleType{
        CONDITIONAL_FORMATTING("conditionalFormatting") ,HEADER_FOOTER("headerFooter") ;
        private String type;

        StyleType(String type) {
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
    public static class BorderOption{
        public static class BorderBody{

            public enum  BorderStyle {
                SLANT_DASH_DOT("slantDashDot"),
                DOTTED("dotted"),
                THICK("thick"),
                HAIR("hair"),
                DASH_DOT("dashDot"),
                DASH_DOT_DOT("dashDotDot"),
                DASHED("dashed"),
                THIN("thin"),
                MEDIUM_DASH_DOT("mediumDashDot"),
                MEDIUM("medium"),
                DOUBLE("double"),
                MEDIUM_DASHED("mediumDashed");
                private String style;

                BorderStyle(String style) {
                    this.style = style;
                }

                public String getStyle() {
                    return style;
                }

                @Override
                public String toString() {
                    return style;
                }
            }
            private String color;
            private BorderStyle style;
        }
        private BorderBody full;
        private BorderBody left;
        private BorderBody right;
        private  BorderBody bottom;
        private BorderBody top;

        private String generateBorderStyle(String key,BorderBody direction){
            return             "<"+key+" style=\"" +
                   direction.style +
                    "\">" +
                    "<color rgb=\"" +
                    ColorUtils.convertToHex(
                            direction.color
            ).replace("#", "") +
                    "\" />" +
                    "</"+key+">";
        }
        public BorderBody getLeftOrFull(){
            return left!=null?left:full;
        }
        public BorderBody getRightOrFull(){
            return right!=null?right:full;
        }
        public BorderBody getTopOrFull(){
            return top!=null?top:full;
        }
        public BorderBody getBottomOrFull(){
            return bottom!=null?bottom:full;
        }
        public String getBorderStyle() {
            String border="";
            BorderBody leftDirection=getLeftOrFull();
            BorderBody rightDirection=getRightOrFull();
            BorderBody topDirection=getTopOrFull();
            BorderBody bottomDirection=getBottomOrFull();
            if(leftDirection!=null){
                border+=generateBorderStyle("left",leftDirection);
            }
            if(rightDirection!=null){
                border+=generateBorderStyle("right",rightDirection);
            }
            if(topDirection!=null){
                border+=generateBorderStyle("top",topDirection);
            }
            if(bottomDirection!=null){
                border+=generateBorderStyle("bottom",bottomDirection);
            }
            return border;
        }
    }
    @AllArgsConstructor
    @NoArgsConstructor
    @Builder
    @Data
    public static class AlignmentOption{
        private String generateAlignmentPropertyValue(String key,String value){
            return " " +
                    key +
                    "=\"" +
                    value +
                    "\" ";
        }
        public String generateAligmentString() {
            String al="";

            if(this.getReadingValidOrder()!=null){
                al+=generateAlignmentPropertyValue("readingOrder",this.getReadingValidOrder());
            }
            if(Utils.booleanCheck(wrapText)){
                al+=generateAlignmentPropertyValue("wrapText","1");
            }
            if(Utils.booleanCheck(shrinkToFit)){
                al+=generateAlignmentPropertyValue("shrinkToFit","1");
            }
            if(indent!=null){
                al+=generateAlignmentPropertyValue("indent",indent+"");
            }
            if(textRotation!=null){
                al+=generateAlignmentPropertyValue("textRotation",textRotation+"");
            }
            if(horizontal!=null){
                al+=generateAlignmentPropertyValue("horizontal",horizontal.getDirection());
            }
            if(vertical!=null){
                al+=generateAlignmentPropertyValue("vertical",vertical.getDirection());
            }

            return al;
        }

        public enum AlignmentHorizontal{
            CENTER("center"),
            LEFT("left"),
            RIGHT("right");
            private String direction;

            AlignmentHorizontal(String direction) {
                this.direction = direction;
            }

            public String getDirection() {
                return direction;
            }
        }
        public enum AlignmentVertical{
            CENTER("center"),
            TOP("top"),
            BOTTOM("bottom");
            private String direction;

            AlignmentVertical(String direction) {
                this.direction = direction;
            }

            public String getDirection() {
                return direction;
            }

        }
        private AlignmentHorizontal horizontal;
        private AlignmentVertical vertical;
        private Boolean wrapText;
        private Boolean shrinkToFit;
        private String readingOrder;
        private Number textRotation;
        private Number indent;
        private Boolean rtl;
        private Boolean ltr;

        public String getReadingValidOrder() {
            if(this.readingOrder.equals("1")||this.readingOrder.equals("2")){
                return this.readingOrder;
            }else{
                return null;
            }
        }
    }
    private String fontFamily;
    private StyleType type;
    private Number size;
    private Number index;
    private AlignmentOption alignment;
    private BorderOption border;
    private String format;
    private Boolean bold;
    private Boolean underline;
    private Boolean italic;
    private Boolean doubleUnderline;
    private String color;
    private String backgroundColor;


    public Boolean boldIsExist(){
        return Optional.ofNullable(bold).orElse(false);
    }
    public Boolean underlineIsExist(){
        return Optional.ofNullable(underline).orElse(false);
    }
    public Boolean italicIsExist(){
        return Optional.ofNullable(italic).orElse(false);
    }
    public Boolean doubleUnderlineIsExist(){
        return Optional.ofNullable(doubleUnderline).orElse(false);
    }
    public Boolean fontFamilyIsExist(){
        return fontFamily!=null;
    }
    public Boolean colorIsExist(){
        return color!=null;
    }
    public Boolean backgroundColorIsExist(){
        return backgroundColor!=null;
    }
    public Boolean sizeIsExist(){
        return size!=null;
    }
    public Boolean formatIsExist(){
        return format!=null;
    }
}
