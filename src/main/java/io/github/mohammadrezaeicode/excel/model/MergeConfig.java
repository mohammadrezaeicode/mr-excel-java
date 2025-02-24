package io.github.mohammadrezaeicode.excel.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class MergeConfig {
    private List<MergeType> mergeType;
    private List<Integer> mergeStart;
    private List<List<Integer>> mergeValue;

    public enum MergeType {
        BOTH("both"),
        COL("col"),
        row("row");
        private final String type;

        MergeType(String type) {
            this.type = type;
        }

        public String getType() {
            return type;
        }
    }
}
