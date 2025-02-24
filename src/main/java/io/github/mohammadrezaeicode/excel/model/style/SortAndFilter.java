package io.github.mohammadrezaeicode.excel.model.style;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@NoArgsConstructor
@AllArgsConstructor
@Data
@Builder
public class SortAndFilter {
    private String ref;
    private Mode mode;

    public enum Mode {
        ALL("all"), REF("ref");
        private final String label;

        Mode(String label) {
            this.label = label;
        }

        public String getLabel() {
            return label;
        }
    }

}
