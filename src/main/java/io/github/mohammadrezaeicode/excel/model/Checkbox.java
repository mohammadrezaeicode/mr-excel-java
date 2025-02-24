package io.github.mohammadrezaeicode.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
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

    public Integer getCol() {
        if (col == null) {
            return -1;
        }
        return col;
    }

    public void setCol(Integer col) {
        this.col = col;
    }

    public Integer getRow() {
        if (row == null) {
            return -1;
        }
        return row;
    }

    public void setRow(Integer row) {
        this.row = row;
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public String getLink() {
        return link;
    }

    public void setLink(String link) {
        this.link = link;
    }

    public Boolean getChecked() {
        if (checked == null) {
            return false;
        }
        return checked;
    }

    public void setChecked(Boolean checked) {
        this.checked = checked;
    }

    public Boolean getMixed() {
        if (mixed == null) {
            return false;
        }
        return mixed;
    }

    public void setMixed(Boolean mixed) {
        this.mixed = mixed;
    }

    public Boolean getThreeD() {
        if (threeD == null) {
            return false;
        }
        return threeD;
    }

    public void setThreeD(Boolean threeD) {
        this.threeD = threeD;
    }

    public String getStartStr() {
        if (startStr == null) {
            return startStr;
        }
        return startStr;
    }

    public void setStartStr(String startStr) {
        this.startStr = startStr;
    }

    public String getEndStr() {
        if (endStr == null) {
            return endStr;
        }
        return endStr;
    }

    public void setEndStr(String endStr) {
        this.endStr = endStr;
    }
}
