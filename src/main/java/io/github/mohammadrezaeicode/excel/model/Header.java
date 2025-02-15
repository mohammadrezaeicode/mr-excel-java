package io.github.mohammadrezaeicode.excel.model;

import lombok.*;

import java.lang.reflect.Method;

@Builder
@NoArgsConstructor
@AllArgsConstructor
@Data
@EqualsAndHashCode(callSuper = true)
public class Header extends HeaderOption {
    private String label;
    private String text;
    private Method method;
//    private formula.FormulaSetting formula;
//    private ConditinalFormating conditinalFormating;


//    conditionalFormatting?: ConditionalFormattingOption;
//    formula?: {
//        type: FormulaType;
//        styleId?: string;
//    };
 public void applyHeaderOption(HeaderOption headerOption){
     if(headerOption.getSize()!=null){
         headerOption.setSize(headerOption.getSize());
     }
     if(headerOption.getComment()!=null){
         headerOption.setComment(headerOption.getComment());
     }
     if(headerOption.getFormula()!=null){
         headerOption.setFormula(headerOption.getFormula());
     }
     if(headerOption.getConditionalFormatting()!=null){
         headerOption.setConditionalFormatting(headerOption.getConditionalFormatting());
     }
     if(headerOption.getMultiStyleValue()!=null){
         headerOption.setMultiStyleValue(headerOption.getMultiStyleValue());
     }

 }
//    public formula.FormulaSetting getFormula() {
//        return formula;
//    }

//    public Comment getComment() {
//        return comment;
//    }
//
//    public void setComment(Comment comment) {
//        this.comment = comment;
//    }
//
//    public ConditinalFormating getConditinalFormating() {
//        return conditinalFormating;
//    }
//
////    public void setFormula(formula.FormulaSetting formula) {
////        this.formula = formula;
////    }
//
//    public void setConditinalFormating(ConditinalFormating conditinalFormating) {
//        this.conditinalFormating = conditinalFormating;
//    }
//
//    public void setMethod(Method method) {
//        this.method = method;
//    }
//
//public Header(String label, String text, Method method) {
//    super();
//    this.label = label;
//    this.text = text;
//    this.method = method;
//}
//
//    public Method getMethod() {
//        return method;
//    }
//
//    public Header(String label, String text) {
//        this.label = label;
//        this.text = text;
//    }
//    public Header(String label, String text,Method method) {
//        this.label = label;
//        this.text = text;
//        this.method=method;
//    }
//
//    public String getLabel() {
//        return label;
//    }
//
//    public String getText() {
//        return text;
//    }
//
//    public Integer getSize() {
//        return size;
//    }
//
//    public void setLabel(String label) {
//        this.label = label;
//    }
//
//    public void setText(String text) {
//        this.text = text;
//    }
//
//    public void setSize(Integer size) {
//        this.size = size;
//    }

    //    label: string;
//    text: string;
//    size?: number;
//    multiStyleValue?: MultiStyleValue;
//    comment?: Comment | string;
//    conditinalFormating?: ConditinalFormating;
//    formula?: {
//        type: FormulaType;
//        styleId?: string;
//    };
}
