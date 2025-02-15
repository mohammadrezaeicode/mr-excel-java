package io.github.mohammadrezaeicode.excel.model;

public class Checkbox {
    private Integer col;
    private Integer row;
    private String text;
    private String link;
    private Boolean checked;
    private Boolean mixed;
    private Boolean threeD;
    private String startStr;
    private String endStr;

    public Checkbox() {
    }

    public Checkbox(Integer col, Integer row, String text, String link, Boolean checked, Boolean mixed, Boolean threeD, String startStr, String endStr) {
        this.col = col;
        this.row = row;
        this.text = text;
        this.link = link;
        this.checked = checked;
        this.mixed = mixed;
        this.threeD = threeD;
        this.startStr = startStr;
        this.endStr = endStr;
    }

    public void setCol(Integer col) {
        this.col = col;
    }

    public void setRow(Integer row) {
        this.row = row;
    }

    public void setText(String text) {
        this.text = text;
    }

    public void setLink(String link) {
        this.link = link;
    }

    public void setChecked(Boolean checked) {
        this.checked = checked;
    }

    public void setMixed(Boolean mixed) {
        this.mixed = mixed;
    }

    public void setThreeD(Boolean threeD) {
        this.threeD = threeD;
    }

    public void setStartStr(String startStr) {
        this.startStr = startStr;
    }

    public void setEndStr(String endStr) {
        this.endStr = endStr;
    }

    public Integer getCol() {
        if(col==null){
            return -1;
        }
        return col;
    }

    public Integer getRow() {
        if(row==null){
            return -1;
        }
        return row;
    }

    public String getText() {
        if(text==null){
            return "";
        }
        return text;
    }

    public String getLink() {
        if(link==null){
            return "";
        }
        return link;
    }

    public Boolean getChecked() {
        if(checked==null){
            return false;
        }
        return checked;
    }

    public Boolean getMixed() {
        if(mixed==null){
            return false;
        }
        return mixed;
    }

    public Boolean getThreeD() {
        if(threeD==null){
            return false;
        }
        return threeD;
    }

    public String getStartStr() {
        if(startStr==null){
            return startStr;
        }
        return startStr;
    }

    public String getEndStr() {
        if(endStr==null){
            return endStr;
        }
        return endStr;
    }
}
