package io.github.mohammadrezaeicode.excel.util;

import io.github.mohammadrezaeicode.excel.model.formula.*;
import io.github.mohammadrezaeicode.excel.model.style1.StyleBody;

import java.util.Map;

public class FormulaUtil {

    public static FormulaProcessResult generateCellRowCol(
            String string,
            FormulaMapBody formula,
            Integer sheetIndex,
            Map<String, StyleBody> styles
            ) {
        if(formula==null){
            return null;
        }
        string = string.toUpperCase();
        String cell = "";
        if (formula instanceof CustomFormulaSetting) {
            var form = (CustomFormulaSetting)formula;
            if(form.getFormula()==null){
                return null;
            }
            String formulaStr =
                    form.getFormula().indexOf("=") == 0 ?
                            form.getFormula().substring(1) : form.getFormula();
            boolean multiInsertCell = string.indexOf(":") > 0;
            String ref = form.getReferenceCells()!=null ? form.getReferenceCells() : string;
            String startRef = multiInsertCell
                    ? string.substring(0, string.indexOf(":"))
                    : string;
            String column = startRef.replaceAll("[0-9]", "");
            int row = Integer.parseInt(string.substring(column.length()));
            String returnType = form.getReturnType()!=null
                    ? form.getReturnType()
                    : Utils.booleanCheck(form.getIsArray()) || multiInsertCell
                    ? " t=\"str\""
                    : "";
            var style ="";
                  if( form.getStyleId()!=null &&
                          styles!=null &&
                          styles.get(form.getStyleId())!=null) {
                      style = " s=\"" + styles.get(form.getStyleId()).getIndex() + "\"";
                  }
            var arrayStr =
                    Utils.booleanCheck(form.getIsArray()) || multiInsertCell ? " t=\"array\" ref=\"" + ref + "\"" : "";
            cell =
                    "<c r=\"" +
                            startRef +
                            "\"" +
                            style +
                            returnType +
                            "><f" +
                            arrayStr +
                            ">" +
                            formulaStr +
                            "</f></c>";
            return FormulaProcessResult.builder()
                    .column(column)
                    .row(row)
                    .needCalcChain(false)
                    .isCustom(true)
                    .cell(cell)
                    .build();
        }
        String column = string.replaceAll("[0-9]", "");
        Integer row = Integer.parseInt(string.substring(column.length()));
        boolean needCalcChain = false;
        String chainCell = "";
        if (formula instanceof NoArgFormulaSetting) {
        var form = (NoArgFormulaSetting)formula;
        if(form.getNoArgType()==null){
            return null;
        }
            if (form.getNoArgType() == NoArgFormulaSetting.NoArgFormulaType.NOW || form.getNoArgType() == NoArgFormulaSetting.NoArgFormulaType.TODAY) {

      String styleString ="";
                        if(form.getStyleId()!=null&&styles!=null&&
                        styles.get(form.getStyleId())!=null) {
                            styleString = " s=\"" + styles.get(form.getStyleId()).getIndex() + "\""
                            ;
                        }

                cell =
                        "<c r=\"" +
                                string +
                                "\"" +
                                styleString +
                                "><f>" +
                                form.getNoArgType().getType() +
                                "()</f></c>";
            } else {
                String value = "NOW()";
                String styleString ="";
                if(form.getStyleId()!=null&&styles!=null&&
                        styles.get(form.getStyleId())!=null) {
                    styleString = " s=\"" + styles.get(form.getStyleId()).getIndex() + "\""
                    ;
                }
                cell =
                        "<c r=\"" +
                                string +
                                "\"" +
                                styleString +
                                "><f>" +
                                form.getNoArgType().getType().substring(4) +
                                "(" +
                                value +
                                ")</f></c>";
            }
            chainCell = "<c r=\"" + string + "\" i=\"" + sheetIndex + "\"/>";
            needCalcChain = true;
        } else if (formula instanceof SingleRefFormulaSetting) {
    var form = (SingleRefFormulaSetting)formula;
            if(form.getReferenceCell()==null){
                return null;
            }
            String value = "";
            if (form.getValue() != null) {
                value = "," + form.getValue();
            }
            String className = "";
            if (form.getType() == SingleRefFormulaSetting.SingleRefFormulaType.COT) {
                className = "_xlfn.";
            }
            String styleString ="";
            if(form.getStyleId()!=null&&styles!=null&&
                    styles.get(form.getStyleId())!=null) {
                styleString = " s=\"" + styles.get(form.getStyleId()).getIndex() + "\""
                ;
            }
            cell =
                    "<c r=\"" +
                            string +
                            "\"" +
                            styleString +
                            "><f>" +
                            className +
                            form.getType().getType() +
                            "(" +
                            form.getReferenceCell().toUpperCase() +
                            value +
                            ")</f></c>";
            chainCell = "<c r=\"" + string + "\" i=\"" + sheetIndex + "\"/>";
            needCalcChain = true;
        } else {
    var form = (FormulaSetting)formula;
            String styleString ="";
            if(form.getStyleId()!=null&&styles!=null&&
                    styles.get(form.getStyleId())!=null) {
                styleString = " s=\"" + styles.get(form.getStyleId()).getIndex() + "\""
                ;
            }
            cell =
                    "<c r=\"" +
                            string +
                            "\"" +
                            (styleString) +
                    ">" +
                    "<f>" +
                    form.getType().getType() +
                    "(" +
                    form.getStart().toUpperCase() +
                    ":" +
                    form.getEnd().toUpperCase() +
                    ")</f>" +
                    "</c>";
        }

        return FormulaProcessResult.builder()
                .column(column)
                .row(row)
                .cell(cell)
                .needCalcChain(needCalcChain)
                .chainCell(chainCell)
                .build();

    }

}
