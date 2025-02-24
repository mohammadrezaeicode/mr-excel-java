package io.github.mohammadrezaeicode.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@AllArgsConstructor
@NoArgsConstructor
@Data
@Builder
public class PageOption {
    private Margin margin;
    private HeaderFooterTypes header;
    private HeaderFooterTypes footer;
    private Boolean isPortrait;

    public HeaderFooterTypes getFooterOrHeader(String type) {
        if (type == "header") {
            return this.header;
        }
        return footer;
    }

    @AllArgsConstructor
    @NoArgsConstructor
    @Data
    @Builder
    public static class Margin {
        private Number left;
        private Number right;
        private Number top;
        private Number bottom;
        private Number header;
        private Number footer;

        public void setFullIfNotNull(Margin margin) {
            if (margin.getLeft() != null) {
                this.left = margin.getLeft();
            }
            if (margin.getRight() != null) {
                this.right = margin.getRight();
            }
            if (margin.getTop() != null) {
                this.top = margin.getTop();
            }
            if (margin.getBottom() != null) {
                this.bottom = margin.getBottom();
            }
            if (margin.getFooter() != null) {
                this.footer = margin.getFooter();
            }
            if (margin.getHeader() != null) {
                this.header = margin.getHeader();
            }
        }
    }

    @AllArgsConstructor
    @NoArgsConstructor
    @Data
    @Builder
    public static class HeaderFooterTypes {
        private HeaderFooterLocationMap odd;
        private HeaderFooterLocationMap even;
        private HeaderFooterLocationMap first;

        public List<String> getKeys() {
            return Arrays.asList("odd", "even", "first");
        }

        public HeaderFooterLocationMap getByKey(String key) {
            if (key.equals("odd")) {
                return odd;
            } else if (key.equals("even")) {
                return even;
            } else if (key.equals("first")) {
                return first;
            }
            return null;
        }

        @AllArgsConstructor
        @NoArgsConstructor
        @Data
        @Builder
        public static class HeaderFooterLocationMap {
            private HeaderFooterOption l;
            private HeaderFooterOption c;
            private HeaderFooterOption r;

            public List<String> getKeys() {
                return Arrays.asList("l", "c", "r");
            }

            public List<HeaderFooterOption> getSortItem() {
                List<HeaderFooterOption> list = new ArrayList<>();
                if (l != null) {
                    list.add(l);
                }
                if (c != null) {
                    list.add(c);
                }
                if (r != null) {
                    list.add(r);
                }
                return list;
            }

            public List<String> getSortItemKey() {
                List<String> list = new ArrayList<>();
                if (l != null) {
                    list.add("l");
                }
                if (c != null) {
                    list.add("c");
                }
                if (r != null) {
                    list.add("r");
                }
                return list;
            }

            public HeaderFooterOption getByKey(String key) {
                if (key.equals("l")) {
                    return l;
                } else if (key.equals("c")) {
                    return c;
                } else if (key.equals("r")) {
                    return r;
                }
                return null;
            }

            @AllArgsConstructor
            @NoArgsConstructor
            @Data
            @Builder
            public static class HeaderFooterOption {
                private String text;
                private String styleId;
            }
        }
    }

}
