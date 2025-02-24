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
    private Map<String, Condition> mergeConditionMap;

    public Condition get(String columnKey) {
        if (mergeConditionMap == null) {
            mergeConditionMap = new HashMap<>();
        }
        return mergeConditionMap.get(columnKey);
    }

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

        public Boolean getInProgress() {
            return inProgress;
        }

        public void setInProgress(Boolean inProgress) {
            this.inProgress = inProgress;
        }

        public Integer getStart() {
            return start;
        }

        public void setStart(Integer start) {
            this.start = start;
        }
    }
}