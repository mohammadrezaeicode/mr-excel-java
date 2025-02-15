package io.github.mohammadrezaeicode.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
@Data
//@Builder
@AllArgsConstructor
@NoArgsConstructor
public class MergeConfig{
    public enum MergeType{
        BOTH("both"),
        COL("col"),
        row("row");
        private String type;

        MergeType(String type) {
            this.type = type;
        }

        public String getType() {
            return type;
        }
    }
    private List<MergeType>mergeType;
    private List<Integer> mergeStart;
    private List<List<Integer>> mergeValue;
}
