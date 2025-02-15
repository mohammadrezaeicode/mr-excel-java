package io.github.mohammadrezaeicode.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.HashMap;
import java.util.Map;
@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class MergeRowConditionMap {
    public Condition get(String columnKey) {
        return mergeConditionMap.get(columnKey);
    }

    @Data
    @Builder
    @NoArgsConstructor
    public static class Condition {
        private Boolean inProgress;
        private Integer start;

        public Condition(Boolean inProgress, Integer start) {
            this.inProgress = inProgress;
            this.start = start;
        }

        public void setInProgress(Boolean inProgress) {
            this.inProgress = inProgress;
        }

        public void setStart(Integer start) {
            this.start = start;
        }

        public Boolean getInProgress() {
            return inProgress;
        }

        public Integer getStart() {
            return start;
        }
    }

    private Map<String, Condition> mergeConditionMap;

    public Map<String, Condition> getMergeConditionMap() {
        if (mergeConditionMap == null) {
            mergeConditionMap = new HashMap<>();
        }
        return mergeConditionMap;
    }

    public void put(String s, Condition condition) {
        if (mergeConditionMap == null) {
            mergeConditionMap = new HashMap<>();
        }
        mergeConditionMap.put(s, condition);
    }
}