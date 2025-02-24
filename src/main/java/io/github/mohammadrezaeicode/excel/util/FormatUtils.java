package io.github.mohammadrezaeicode.excel.util;

import io.github.mohammadrezaeicode.excel.model.style.Format;

import java.util.HashMap;
import java.util.Map;

public class FormatUtils {
    public static Map<String, Format> formatMap;

    static {
        formatMap = new HashMap<>();
        formatMap.put("time", new Format(165, "<numFmt numFmtId=\"165\" formatCode=\"[$-F400]h:mm:ss\\ AM/PM\" />"));
        formatMap.put("date", new Format(187, "<numFmt numFmtId=\"187\" formatCode=\"[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy\" />"));
        formatMap.put("short_date", new Format(14, null));
        formatMap.put("fraction", new Format(13, null));
        formatMap.put("percentage", new Format(9, null));
        formatMap.put("float_1", new Format(180, "<numFmt numFmtId=\"180\" formatCode=\"0.0\" />"));
        formatMap.put("float_2", new Format(181, "<numFmt numFmtId=\"181\" formatCode=\"0.00\" />"));
        formatMap.put("float_3", new Format(164, "<numFmt numFmtId=\"164\" formatCode=\"0.000\" />"));
        formatMap.put("float_4", new Format(182, "<numFmt numFmtId=\"182\" formatCode=\"0.0000\" />"));
        formatMap.put("dollar_rounded", new Format(183, "<numFmt numFmtId=\"183\" formatCode=\"&quot;$&quot;#,##0\" />"));

        formatMap.put("dollar_2", new Format(183, "<numFmt numFmtId=\"183\" formatCode=\"&quot;$&quot;#,##0.00\" />"));
        formatMap.put("num_sep", new Format(184, "<numFmt numFmtId=\"184\" formatCode=\"#,##0\" />"));
        formatMap.put("num_sep_1", new Format(185, "<numFmt numFmtId=\"185\" formatCode=\"#,##0.0\" />"));
        formatMap.put("num_sep_2", new Format(186, "<numFmt numFmtId=\"186\" formatCode=\"#,##0.00\" />"));
        formatMap.put("dollar", new Format(163, "<numFmt numFmtId=\"163\" formatCode=\"_([$$-409]* #,##0.00_);_([$$-409]* \\(#,##0.00\\);_([$$-409]* &quot;-&quot;??_);_(@_)\" />"));
        formatMap.put("$", new Format(163, "<numFmt numFmtId=\"163\" formatCode=\"_([$$-409]* #,##0.00_);_([$$-409]* \\(#,##0.00\\);_([$$-409]* &quot;-&quot;??_);_(@_)\" />"));
        formatMap.put("pound", new Format(162, "<numFmt numFmtId=\"162\" formatCode=\"_-[$£-809]* #,##0.00_-;\\-[$£-809]* #,##0.00_-;_-[$£-809]* &quot;-&quot;??_-;_-@_-\" />"));
        formatMap.put("£", new Format(162, "<numFmt numFmtId=\"162\" formatCode=\"_-[$£-809]* #,##0.00_-;\\-[$£-809]* #,##0.00_-;_-[$£-809]* &quot;-&quot;??_-;_-@_-\" />"));
        formatMap.put("euro", new Format(161, "<numFmt numFmtId=\"161\" formatCode=\"_([$€-2]\\ * #,##0.00_);_([$€-2]\\ * \\(#,##0.00\\);_([$€-2]\\ * &quot;-&quot;??_);_(@_)\" />"));
        formatMap.put("€", new Format(161, "<numFmt numFmtId=\"161\" formatCode=\"_([$€-2]\\ * #,##0.00_);_([$€-2]\\ * \\(#,##0.00\\);_([$€-2]\\ * &quot;-&quot;??_);_(@_)\" />"));
        formatMap.put("yen", new Format(160, "<numFmt numFmtId=\"160\" formatCode=\"_ [$¥-804]* #,##0.00_ ;_ [$¥-804]* \\-#,##0.00_ ;_ [$¥-804]* &quot;-&quot;??_ ;_ @_ \" />"));
        formatMap.put("¥", new Format(160, "<numFmt numFmtId=\"160\" formatCode=\"_ [$¥-804]* #,##0.00_ ;_ [$¥-804]* \\-#,##0.00_ ;_ [$¥-804]* &quot;-&quot;??_ ;_ @_ \" />"));
        formatMap.put("CHF", new Format(179, "<numFmt numFmtId=\"179\" formatCode=\"_-* #,##0.00\\ [$CHF-100C]_-;\\-* #,##0.00\\ [$CHF-100C]_-;_-* &quot;-&quot;??\\ [$CHF-100C]_-;_-@_-\" />"));
        formatMap.put("ruble", new Format(178, "<numFmt numFmtId=\"178\" formatCode=\"_-* #,##0.00\\ [$₽-419]_-;\\-* #,##0.00\\ [$₽-419]_-;_-* &quot;-&quot;??\\ [$₽-419]_-;_-@_-\" />"));
        formatMap.put("₽", new Format(178, "<numFmt numFmtId=\"178\" formatCode=\"_-* #,##0.00\\ [$₽-419]_-;\\-* #,##0.00\\ [$₽-419]_-;_-* &quot;-&quot;??\\ [$₽-419]_-;_-@_-\" />"));
        formatMap.put("֏", new Format(177, "<numFmt numFmtId=\"177\" formatCode=\"_-* #,##0.00\\ [$֏-42B]_-;\\-* #,##0.00\\ [$֏-42B]_-;_-* &quot;-&quot;??\\ [$֏-42B]_-;_-@_-\" />"));
        formatMap.put("manat", new Format(176, "<numFmt numFmtId=\"176\" formatCode=\"_-* #,##0.00\\ [$₼-82C]_-;\\-* #,##0.00\\ [$₼-82C]_-;_-* &quot;-&quot;??\\ [$₼-82C]_-;_-@_-\" />"));
        formatMap.put("₼", new Format(176, "<numFmt numFmtId=\"176\" formatCode=\"_-* #,##0.00\\ [$₼-82C]_-;\\-* #,##0.00\\ [$₼-82C]_-;_-* &quot;-&quot;??\\ [$₼-82C]_-;_-@_-\" />"));
        formatMap.put("₼1", new Format(175, "<numFmt numFmtId=\"175\" formatCode=\"_-* #,##0.00\\ [$₼-42C]_-;\\-* #,##0.00\\ [$₼-42C]_-;_-* &quot;-&quot;??\\ [$₼-42C]_-;_-@_-\" />"));
        formatMap.put("₽1", new Format(174, "<numFmt numFmtId=\"174\" formatCode=\"_-* #,##0.00\\ [$₽-46D]_-;\\-* #,##0.00\\ [$₽-46D]_-;_-* &quot;-&quot;??\\ [$₽-46D]_-;_-@_-\" />"));
        formatMap.put("₽2", new Format(173, "<numFmt numFmtId=\"173\" formatCode=\"_-* #,##0.00\\ [$₽-485]_-;\\-* #,##0.00\\ [$₽-485]_-;_-* &quot;-&quot;??\\ [$₽-485]_-;_-@_-\" />"));
        formatMap.put("₽3", new Format(172, "<numFmt numFmtId=\"172\" formatCode=\"_-* #,##0.00\\ [$₽-444]_-;\\-* #,##0.00\\ [$₽-444]_-;_-* &quot;-&quot;??\\ [$₽-444]_-;_-@_-\" />"));
        formatMap.put("ريال", new Format(171, "<numFmt numFmtId=\"171\" formatCode=\"_ * #,##0.00_-[$ريال-429]_ ;_ * #,##0.00\\-[$ريال-429]_ ;_ * &quot;-&quot;??_-[$ريال-429]_ ;_ @_ \" />"));

    }

    public static String spCh(String str) {
        return str.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
    }

}
