package io.github.mohammadrezaeicode.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ConditionalFormattingOption {
    private Type type;
    private OperatorIml operator;
    private Object value;
    private Integer priority;
    private List<String> colors;
    private Boolean bottom;
    private String styleId;
    private Double percent;

    public Object getValue() {
        Object ob;
        if (value instanceof String) {
            ob = value;
        } else if (value instanceof Number) {
            ob = value;
        } else {
            ob = null;
        }
        return ob;
    }

    public enum ConditionalFormattingTopOperation implements OperatorIml {
        BELOW_AVERAGE("belowAverage"),
        ABOVE_AVERAGE("aboveAverage");
        private final String operationType;

        ConditionalFormattingTopOperation(String operationType) {
            this.operationType = operationType;
        }

        public String getOperationType() {
            return operationType;
        }
    }

    public enum ConditionalFormattingIconSetOperation implements OperatorIml {
        ARROWS_3("3Arrows"),
        ARROWS_4("4Arrows"),
        ARROWS_5("5Arrows"),
        ARROWS_GRAY_5("5ArrowsGray"),
        ARROWS_GRAY_4("4ArrowsGray"),
        ARROWS_GRAY_3("3ArrowsGray");
        private final String operationType;

        ConditionalFormattingIconSetOperation(String operationType) {
            this.operationType = operationType;
        }

        public String getOperationType() {
            return operationType;
        }
    }

    public enum ConditionalFormattingCellsOperation implements OperatorIml {
        LT("lt"),
        GT("gt"),
        BETWEEN("between"),
        EQ("eq"),
        CT("ct");
        private final String operationType;

        ConditionalFormattingCellsOperation(String operationType) {
            this.operationType = operationType;
        }

        public String getOperationType() {
            return operationType;
        }
    }

    public enum Type implements OperatorIml {
        CELLS("cells"), DATABAR("databar"), ICON_SET("iconSet"), COLOR_SCALE("colorScale"), TOP("top");
        private final String type;

        Type(String type) {
            this.type = type;
        }

        public String getType() {
            return type;
        }
    }

    public interface OperatorIml {

    }

    @Data
    @Builder
    @NoArgsConstructor
    public static class Operator<T> implements OperatorIml {
        private T operator;

        public Operator(T operator) {
            this.operator = operator;
        }

        public T getOperator() {
            return operator;
        }
    }
}
