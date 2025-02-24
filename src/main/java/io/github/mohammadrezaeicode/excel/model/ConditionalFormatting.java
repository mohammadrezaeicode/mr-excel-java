package io.github.mohammadrezaeicode.excel.model;

import lombok.*;

import java.util.List;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@EqualsAndHashCode(callSuper = true)
public class ConditionalFormatting extends ConditionalFormattingOption {
    private String start;
    private String end;

    public static EBuilder eBuilder() {
        return new EBuilder();
    }

    public static class EBuilder {
        private final ConditionalFormatting conditionalFormatting;

        public EBuilder() {
            this.conditionalFormatting = new ConditionalFormatting();
        }

        public EBuilder start(String start) {
            conditionalFormatting.start = start;
            return this;
        }

        public EBuilder end(String end) {
            conditionalFormatting.end = end;
            return this;
        }

        public EBuilder type(Type type) {
            conditionalFormatting.setType(type);
            return this;
        }

        public EBuilder operator(OperatorIml operator) {
            conditionalFormatting.setOperator(operator);
            return this;
        }

        public EBuilder value(Object value) {
            conditionalFormatting.setValue(value);
            return this;
        }

        public EBuilder priority(Integer priority) {
            conditionalFormatting.setPriority(priority);
            return this;
        }

        public EBuilder colors(List<String> colors) {
            conditionalFormatting.setColors(colors);
            return this;
        }

        public EBuilder bottom(Boolean bottom) {
            conditionalFormatting.setBottom(bottom);
            return this;
        }

        public EBuilder styleId(String styleId) {
            conditionalFormatting.setStyleId(styleId);
            return this;
        }

        public EBuilder percent(Double percent) {
            conditionalFormatting.setPercent(percent);
            return this;
        }

        public ConditionalFormatting build() {
            return this.conditionalFormatting;
        }
    }
}
