package io.github.mohammadrezaeicode.excel.function;

import io.github.mohammadrezaeicode.excel.model.*;
import io.github.mohammadrezaeicode.excel.model.converter.RowMap;
import io.github.mohammadrezaeicode.excel.model.formula.FormulaMapBody;
import io.github.mohammadrezaeicode.excel.model.formula.FormulaSetting;
import io.github.mohammadrezaeicode.excel.model.func.CommentConditionFunctionInput;
import io.github.mohammadrezaeicode.excel.model.func.MergeRowDataConditionInputFunction;
import io.github.mohammadrezaeicode.excel.model.func.MultiStyleConditionInputFunction;
import io.github.mohammadrezaeicode.excel.model.func.StyleCellConditionInputFunction;
import io.github.mohammadrezaeicode.excel.model.row.SheetDimension;
import io.github.mohammadrezaeicode.excel.model.style.SortAndFilter;
import io.github.mohammadrezaeicode.excel.model.style.StyleBody;
import io.github.mohammadrezaeicode.excel.model.style.StyleMapper;
import io.github.mohammadrezaeicode.excel.model.types.GenerateType;
import io.github.mohammadrezaeicode.excel.service.GeneralConfig;
import io.github.mohammadrezaeicode.excel.util.*;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Function;

import static io.github.mohammadrezaeicode.excel.util.CommentUtil.commentConvertor;

public class GenerateExcel {
    public static <T> Result generateHeaderAndGenerateExcelWithMultiSheet(List<SheetGenerator<T>> sheetData, ExcelTableOption options) throws InvocationTargetException, IllegalAccessException, IOException, NoSuchMethodException {
        List<Sheet> sheetList = new ArrayList<>();
        for (SheetGenerator<T> sheet : sheetData) {
            sheetList.add(SheetGeneratorUtils.generateSheet(sheet.getData(), sheet.getHeaderClass(), sheet.getApplyHeaderOptionFunction(), sheet.getApplySheetOptionFunction()));
        }
        var exOb = ExcelTable.builder()
                .sheet(
                        sheetList
                )
                .build();
        if (options != null) {
            exOb.applyExcelTableOption(options);
        }
        return generateExcel(exOb, "");
    }

    public static <T> Result generateHeaderAndGenerateExcel(List<T> data, ExcelTableOption options, Class headerClass, Function<List<Header>, List<Header>> applyHeaderOptionFunction, Function<Sheet.SheetBuilder, Sheet> applySheetOptionFunction) throws InvocationTargetException, IllegalAccessException, IOException, NoSuchMethodException {
        var exOb = ExcelTable.builder()
                .sheet(
                        Collections.singletonList(SheetGeneratorUtils.generateSheet(data, headerClass, applyHeaderOptionFunction, applySheetOptionFunction))
                )
                .build();
        if (options != null) {
            exOb.applyExcelTableOption(options);
        }
        return generateExcel(exOb, "");
    }

    public static Result generateExcel(ExcelTable data, String styleKey) throws InvocationTargetException, IllegalAccessException, IOException {
        if (data == null) {
            data = new ExcelTable();
        }
        Map<String, String> resultStringMap = new HashMap<>();
        Map<String, ImageOutput> resultAssetsMap = new HashMap<>();

        if (styleKey != null && styleKey.length() > 0) {
            data = GeneralConfig.applyConfig(styleKey, data);
        }
        if (data.getCreator() != null && data.getCreator().trim().length() == 0) {
            throw new Error("length of \"creator\" most be bigger then 0");
        }
        //TODO: next
//        if (
//                typeof data.created == "string" &&
//                new Date(data.created).toString() == "Invalid Date"
//  ) {
//            throw '"created" is not valid date';
//        }
//        if (
//                typeof data.modified == "string" &&
//                new Date(data.modified).toString() == "Invalid Date"
//  ) {
//            throw '"modified" is not valid date';
//        }

        var formatMap = FormatUtils.formatMap;
        if (data.getFormatMap() != null && !data.getFormatMap().isEmpty()) {
            formatMap.putAll(data.getFormatMap());
        }
        var fetchFunc = data.getFetch();
        Map<String, String> operatorMap = new HashMap<>();
        operatorMap.put("lt", "lessThan");
        operatorMap.put("gt", "greaterThan");
        operatorMap.put("between", "between");
        operatorMap.put("ct", "containsText");
        operatorMap.put("eq", "equal");

        List<String> cols = new ArrayList<>(ColumnUtils.DEFAULT_COLUMN);
        if (data.getNumberOfColumn() != null && data.getNumberOfColumn() > 25) {
            List<String> result = ColumnUtils.generateColumnName(cols, data.getNumberOfColumn());
            cols.clear();
            cols.addAll(result);
        }
        if (data.getSheet() == null) {
            data.setSheet(new ArrayList<>(Collections.singletonList(Sheet.builder().headers(new ArrayList<>()).data(new ArrayList<>()).build())));
        }
        int sheetLength = data.getSheet().size();
        String xlFolder = "xl";
        String xl_media_Folder = xlFolder + "/media";
        String xl_drawingsFolder = xlFolder + "/drawings";
        String xl_drawings_relsFolder = xl_drawingsFolder + "/_rels";
        if (data.getStyles() == null) {
            data.setStyles(new HashMap<>());
        }
        if (Utils.booleanCheck(data.getAddDefaultTitleStyle())) {
            data.addStyle("titleStyle", StyleBody.builder().alignment(
                    StyleBody.AlignmentOption.builder().horizontal(
                            StyleBody.AlignmentOption.AlignmentHorizontal.CENTER
                    ).vertical(
                            StyleBody.AlignmentOption.AlignmentVertical.CENTER
                    ).build()
            ).build());
        }
        var styles = data.getStyles();
        var styleKeys = styles.keySet();
        String defaultCommentStyle = CommentUtil.defaultCellCommentStyle;
        boolean addCF = Optional.ofNullable(data.getActivateConditionalFormatting()).orElse(false);
        Map<String, String> headerFooterStyle = new HashMap<>();
        Map<String, Number> cFMapIndex = new HashMap<>();
        var styleMapper = new StyleMapper(styles, addCF);
        if (styleMapper.getConditionalFormattingStyle() != null && addCF) {
            cFMapIndex.putAll(styleMapper.getConditionalFormattingStyle());
        }
        if (styleMapper.getHeaderFooterStyle() != null) {
            headerFooterStyle.putAll(styleMapper.getHeaderFooterStyle());
        }
//
        resultStringMap.put(xlFolder + "/styles.xml", XmlUtils.styleGenerator(styleMapper, addCF));
//
        var sheetContentType =
                "<Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" PartName=\"/xl/worksheets/sheet1.xml\" />";
        var sharedString = "";
        var sharedStringIndex = 0;
        var workbookString = "";
        var workbookRelString = "";
        Map<String, String> sharedStringMap = new HashMap<>();
        Map<String, SheetProcessResult> mapData = new HashMap<>();

        String sheetNameApp = "";
        int indexId = 4;
        boolean selectedAdded = false;
        int activeTabIndex = -1;
        List<String> arrTypes = new ArrayList<>();
        AtomicInteger imageCounter = new AtomicInteger(1);

        var formCtrlMap = new HashMap<String, String>();
        formCtrlMap.put("checkbox",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                        "<formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" objectType=\"CheckBox\" **value** **fmlaLink** lockText=\"1\" noThreeD=\"1\"/>");

        AtomicInteger shapeIdCounter = new AtomicInteger(1024);
        Map<String, String> shapeMap = new HashMap<>();
        shapeMap.put("checkbox", "<v:shape id=\"***id***\" type=\"#_x0000_t201\" style='position:absolute;\n  margin-left:1.5pt;margin-top:1.5pt;width:63pt;height:16.5pt;z-index:1;\n  mso-wrap-style:tight' filled=\"f\" fillcolor=\"window [65]\" stroked=\"f\"\n  strokecolor=\"windowText [64]\" o:insetmode=\"auto\">\n  <v:path shadowok=\"t\" strokeok=\"t\" fillok=\"t\"/>\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\n  <v:textbox style='mso-direction-alt:auto' o:singleclick=\"f\">\n   <div style='text-align:left'><font face=\"Segoe UI\" size=\"160\" color=\"auto\">***text***</font></div>\n  </v:textbox>\n  <x:ClientData ObjectType=\"Checkbox\">\n   <x:SizeWithCells/>\n   <x:Anchor>\n    0, 2, 0, 2, 0, 86, 1, 0</x:Anchor>\n   <x:AutoFill>False</x:AutoFill>\n   <x:AutoLine>False</x:AutoLine>\n   <x:TextVAlign>Center</x:TextVAlign>\n   <x:NoThreeD/>\n  </x:ClientData>\n </v:shape>");
        var shapeTypeMap = new HashMap<>();
        shapeTypeMap.put("checkbox", "<v:shapetype id=\"_x0000_t201\" coordsize=\"21600,21600\" o:spt=\"201\"\n  path=\"m,l,21600r21600,l21600,xe\">\n  <v:stroke joinstyle=\"miter\"/>\n  <v:path shadowok=\"f\" o:extrusionok=\"f\" strokeok=\"f\" fillok=\"f\" o:connecttype=\"rect\"/>\n  <o:lock v:ext=\"edit\" shapetype=\"t\"/>\n </v:shapetype>");
//
        List<String> checkboxForm = new ArrayList<>();
        String calcChainValue = "";
        boolean needCalcChain = false;
        String xl_tableFolder = xlFolder + "/tables";
        for (int index = 0; index < sheetLength; index++) {
            var sheetData = data.getSheet().get(index);
            int sheetDataId = index + 1;
            Map<Integer, RowMap> rowMap = new HashMap<>();
            SheetDimension sheetDimensions = new SheetDimension();
            var asTable = sheetData.getAsTable();
            String sheetDataTableColumns = "";
            int rowCount =
                    sheetData.getShiftTop() != null && sheetData.getShiftTop() >= 0
                            ? sheetData.getShiftTop() + 1
                            : 1;
            var sheetDataString = "";
            var sheetSizeString = "";
            var sheetSortFilter = "";
            var splitOption = "";
            var sheetViewProperties = "";
            var viewType = "";
            var hasCheckbox = false;
            var checkboxDrawingContent = "";
            var checkboxShape = "";
            var formRel = "";
            var checkboxSheetContent = "";
            List<String> mergesCellArray = new ArrayList<>();
            if (sheetData.getMerges() != null) {
                mergesCellArray.addAll(sheetData.getMerges());//TODO should colne
            }
            Map<String, FormulaMapBody> formulaSheetObj = new HashMap<>();
            if (sheetData.getFormula() != null) {
                formulaSheetObj.putAll(sheetData.getFormula());//TODO should colne
            }
            if (formulaSheetObj == null) {
                formulaSheetObj = new HashMap<>();
            }
            List<ConditionalFormatting> conditionalFormatting = new ArrayList<>();
            if (sheetData.getConditionalFormatting() != null && addCF) {
                conditionalFormatting.addAll(sheetData.getConditionalFormatting());//TODO should colne
            }
            boolean hasComment = false;
            List<String> commentAuthor = new ArrayList<>();
            String commentString = "";
            List<ShapeRC> shapeCommentRowCol = new ArrayList<>();
            List<Method> objKey = new ArrayList<>();
            List<Number> headerFormula = new ArrayList<>();
            List<Number> headerConditionalFormatting = new ArrayList<>();
            MergeRowConditionMap mergeRowConditionMap = new MergeRowConditionMap();
            String sheetHeaderFooter = "";
            boolean isPortrait = false;
            String sheetBreakLine = "";
            if (Utils.booleanCheck(sheetData.getRtl())) {
                sheetViewProperties += " rightToLeft=\"1\" ";
            }
            if (sheetData.getPageBreak() != null) {
                var pageBreak = sheetData.getPageBreak();
                if (pageBreak.getRow() != null && pageBreak.getRow().size() > 0) {
                    viewType = "pageBreakPreview";
                    int rowLength = pageBreak.getRow().size();
                    sheetBreakLine +=
                            "<rowBreaks count=\"" +
                                    rowLength +
                                    "\" manualBreakCount=\"" +
                                    rowLength +
                                    "\">" +
                                    pageBreak.getRow().stream()
                                            .map(String::valueOf)
                                            .reduce("",
                                                    (result, current) ->
                                                            result + "<brk id=\"" + current + "\" max=\"16383\" man=\"1\"/>"
                                            ) +
                                    "</rowBreaks>";
                }
                if (pageBreak.getColumn() != null && pageBreak.getColumn().size() > 0) {
                    viewType = "pageBreakPreview";
                    int columnLength = pageBreak.getColumn().size();
                    sheetBreakLine +=
                            "<colBreaks count=\"" +
                                    columnLength +
                                    "\" manualBreakCount=\"" +
                                    columnLength +
                                    "\">" +
                                    pageBreak.getColumn().stream().map(String::valueOf).reduce("",
                                            (result, current) ->
                                                    result + "<brk id=\"" + current + "\" max=\"16383\" man=\"1\"/>"
                                    ) +
                                    "</colBreaks>";
                }
            }
            String sheetMargin = "";
            if (sheetData.getPageOption() != null) {
                var pageOption = sheetData.getPageOption();
                if (Utils.booleanCheck(pageOption.getIsPortrait())) {
                    isPortrait = true;
                }
                if (pageOption.getMargin() != null) {
                    var margin = pageOption.getMargin();
                    var result = PageOption.Margin.builder().left(0.7).right(0.7).top(0.75).bottom(0.75).header(0.3).footer(0.3).build();

                    result.setFullIfNotNull(margin);
                    sheetMargin =
                            "<pageMargins left=\"" +
                                    result.getLeft() +
                                    "\" right=\"" +
                                    result.getRight() +
                                    "\" top=\"" +
                                    result.getTop() +
                                    "\" bottom=\"" +
                                    result.getBottom() +
                                    "\" header=\"" +
                                    result.getHeader() +
                                    "\" footer=\"" +
                                    result.getFooter() +
                                    "\"/>";
                }
                final AtomicReference<String> typeKeeper = new AtomicReference<>("");
                final AtomicReference<String> odd = new AtomicReference<>("");
                final AtomicReference<String> even = new AtomicReference<>("");
                final AtomicReference<String> first = new AtomicReference<>("");
                var keyKey = Arrays.asList("header", "footer");
                keyKey.forEach((keyObj) -> {
                    String endTag = String.valueOf(keyObj.charAt(0)).toUpperCase() + keyObj.substring(1);

                    if (pageOption.getFooterOrHeader(keyObj) != null) {
                        var element = pageOption.getFooterOrHeader(keyObj);
                        if (element != null) {
                            element.getKeys().forEach((typeHF) -> {
                                if (!typeKeeper.get().contains(typeHF)) {
                                    typeKeeper.getAndUpdate(curr -> curr + typeHF);
                                }
                                var typeObj = element.getByKey(typeHF);
                                final AtomicReference<String> node = new AtomicReference<>("");
                                if (typeObj != null) {
                                    typeObj.getSortItemKey()
                                            .forEach((direction) -> {
                                                String currentNodeValue = node.get();
                                                var dirObj = typeObj.getByKey(direction);
                                                currentNodeValue += "&amp;" + direction.toUpperCase();
                                                if (dirObj.getStyleId() != null && headerFooterStyle.get(dirObj.getStyleId()) != null) {
                                                    currentNodeValue += headerFooterStyle.get(dirObj.getStyleId());
                                                }
                                                if (dirObj.getText() != null) {
                                                    currentNodeValue += dirObj.getText();
                                                }
                                                node.set(currentNodeValue);
                                            });
                                }
                                String currentNodeValue = node.get();

                                currentNodeValue =
                                        "<" +
                                                typeHF +
                                                endTag +
                                                ">" +
                                                currentNodeValue +
                                                "</" +
                                                typeHF +
                                                endTag +
                                                ">";

                                node.set(currentNodeValue);
                                switch (typeHF) {
                                    case "odd":
                                        odd.getAndUpdate(curr -> curr + node.get());
                                        break;
                                    case "even":
                                        even.getAndUpdate(curr -> curr + node.get());
                                        break;
                                    case "first":
                                        first.getAndUpdate(curr -> curr + node.get());
                                        break;
                                    default:
                                        throw new Error("type error");
                                }

                            });
                        }
                    }
                });
                sheetHeaderFooter = odd.get() + even.get() + first.get();
                if (sheetHeaderFooter.length() > 0) {
                    isPortrait = true;
                    String typeKeepCurrent = typeKeeper.get();
                    var oddEvenFlag =
                            typeKeepCurrent.length() == "oddeven".length() ||
                                    typeKeepCurrent.length() == "oddevenfirst".length()
                                    ? " differentOddEven=\"1\""
                                    : "";
                    var firstFlag =
                            typeKeepCurrent.contains("first") ? " differentFirst=\"1\"" : "";
                    sheetHeaderFooter =
                            "<headerFooter" +
                                    oddEvenFlag +
                                    firstFlag +
                                    ">" +
                                    sheetHeaderFooter +
                                    "</headerFooter>";
                }
            }
            if (sheetData.getViewOption() != null) {
                String splitState = "";
                var viewOption = sheetData.getViewOption();
                if (viewOption.getType() != null) {
                    viewType = viewOption.getType().getType();
                }
                if (Utils.booleanCheck(viewOption.getHideRuler())) {
                    sheetViewProperties += " showRuler=\"0\" ";
                }
                if (Utils.booleanCheck(viewOption.getHideGrid())) {
                    sheetViewProperties += " showGridLines=\"0\" ";
                }
                if (Utils.booleanCheck(viewOption.getHideHeadlines())) {
                    sheetViewProperties += " showRowColHeaders=\"0\" ";
                }
                var split = viewOption.getSplitOption();
                if (split == null) {
                    isPortrait = false;
                    var frozen = viewOption.getFrozenOption();
                    if (frozen != null) {

                        splitState = " state=\"frozen\" ";
                        if (frozen.getType() == ViewOption.FrozenOption.Type.ROW) {
                            Number fIndex;
                            if (frozen.getR() != null) {
                                fIndex = frozen.getR();
                            } else {
                                fIndex = frozen.getIndex();
                            }
                            fIndex = Optional.ofNullable(fIndex).orElseThrow();
                            split = ViewOption.SplitOption.builder().startAt(
                                            ViewOption.SplitOption.ViewStart.builder()
                                                    .b("A" + (fIndex.intValue() + 1))
                                                    .build()
                                    ).type(ViewOption.SplitOption.Type.HORIZONTAL)
                                    .split(fIndex).build();

                        } else if (frozen.getType() == ViewOption.FrozenOption.Type.COLUMN) {
                            Number fIndex;
                            if (frozen.getC() != null) {
                                fIndex = frozen.getC();
                            } else {
                                fIndex = frozen.getIndex();
                            }
                            fIndex = Optional.ofNullable(fIndex).orElseThrow();
                            if (fIndex.intValue() > cols.size() - 1) {
                                cols = ColumnUtils.generateColumnName(cols, fIndex.intValue());
                            }
                            split = ViewOption.SplitOption.builder()
                                    .type(ViewOption.SplitOption.Type.VERTICAL)
                                    .startAt(
                                            ViewOption.SplitOption.ViewStart.builder()
                                                    .r(
                                                            cols.get(fIndex.intValue()) + 1
                                                    )
                                                    .build()
                                    ).split(fIndex)
                                    .build();

                        } else if (frozen.getType() == ViewOption.FrozenOption.Type.BOTH) {
                            String two;
                            Number splitO = null;
                            var startAtBuilder = ViewOption.SplitOption.builder();
                            if (frozen.getIndex() != null) {
                                splitO = frozen.getIndex();
                                two = cols.get(frozen.getIndex().intValue()) + (frozen.getIndex().intValue() + 1);
                            } else {
                                startAtBuilder.x(frozen.getC());
                                startAtBuilder.y(frozen.getR());
                                two = cols.get(frozen.getC().intValue()) + (frozen.getR().intValue() + 1);
                            }
                            split = startAtBuilder.startAt(ViewOption.SplitOption.ViewStart.builder().two(two).build())
                                    .type(ViewOption.SplitOption.Type.BOTH).split(splitO)
                                    .build();
//                                    {
//                                    startAt: {
//                                two,
//                            },
//                            type: "B",
//                                    split: splitO,
//            };
                        }
                    }
                }
                if (split != null) {
                    if (split.getType() == ViewOption.SplitOption.Type.HORIZONTAL) {
                        String ref = null;
                        if (split.getStartAt() != null) {
                            ref = split.getStartAt().getB();
                            if (split.getStartAt().getT() != null) {
                                sheetViewProperties += " topLeftCell=\"" + split.getStartAt().getT() + "\"";
                            }
                        }
                        if (ref == null) {
                            ref = "A1";
                        }
                        splitOption =
                                "<pane ySplit=\"" +
                                        split.getYOrSplit() +
                                        "\" topLeftCell=\"" +
                                        ref +
                                        "\" activePane=\"bottomLeft\"" +
                                        splitState +
                                        "/>";
                    } else if (split.getType() == ViewOption.SplitOption.Type.VERTICAL) {
                        String ref = null;
                        if (split.getStartAt() != null) {
                            ref = split.getStartAt().getR();
                            if (split.getStartAt().getL() != null) {
                                sheetViewProperties += " topLeftCell=\"" + split.getStartAt().getL() + "\"";
                            }
                        }
                        if (ref == null) {
                            ref = "A1";
                        }
                        splitOption =
                                "<pane xSplit=\"" +
                                        split.getXOrSplit() +
                                        "\" topLeftCell=\"" +
                                        ref +
                                        "\" activePane=\"topLeft\"" +
                                        splitState +
                                        "/>";
                    } else {
                        String ref = null;
                        if (split.getStartAt() != null) {
                            ref = split.getStartAt().getTwo();
                            if (split.getStartAt().getOne() != null) {
                                sheetViewProperties += " topLeftCell=\"" + split.getStartAt().getOne() + "\"";
                            }
                        }
                        if (ref == null) {
                            ref = "A1";
                        }
                        splitOption =
                                "<pane xSplit=\"" +
                                        split.getXOrSplit() +
                                        "\" ySplit=\"" +
                                        split.getYOrSplit() +
                                        "\" topLeftCell=\"" +
                                        ref +
                                        "\" activePane=\"bottomLeft\"" +
                                        splitState +
                                        "/>";
                    }
                }
            }

            if (isPortrait) {
                viewType = "pageLayout";
            }
            if (sheetData.getCheckbox() != null) {
                hasCheckbox = true;
                var strFormDef = formCtrlMap.get("checkbox");
                var shCheckbox = sheetData.getCheckbox();
                for (int i = 0; i < shCheckbox.size(); i++) {
                    var v = shCheckbox.get(i);
                    String formCtlStr = strFormDef;
                    if (v.getLink() != null) {
                        var linkAddress = ColumnUtils.getColRowBaseOnRefString(v.getLink(), cols);
                        formCtlStr = formCtlStr.replace(
                                "**fmlaLink**",
                                "fmlaLink=\"$" +
                                        cols.get(linkAddress.getColNum()) +
                                        "$" +
                                        (linkAddress.getRowNum() + 1) +
                                        "\""
                        );
                    } else {
                        formCtlStr = formCtlStr.replace("**fmlaLink**", "");
                    }
                    if (v.getMixed()) {
                        formCtlStr = formCtlStr.replace("**value**", "checked=\"Mixed\"");
                    } else {
                        if (v.getChecked()) {
                            formCtlStr = formCtlStr.replace("**value**", "checked=\"Checked\"");
                        } else {
                            formCtlStr = formCtlStr.replace("**value**", "");
                        }
                    }
                    if (v.getThreeD()) {
                        formCtlStr.replace("noThreeD=\"1\"", "");
                    }
                    checkboxForm.add(formCtlStr);
                    shapeIdCounter.incrementAndGet();
                    String shapeId = index + String.valueOf(shapeIdCounter.getAndIncrement());
                    String sId = "_x0000_s" + shapeId;
                    checkboxShape += shapeMap.get("checkbox")
                            .replace("***id***", sId)
                            .replace("***text***", v.getText());

                    String from = v.getStartStr();
                    String to = v.getEndStr();
                    var resultVal = new Position();
                    if (v.getCol() != null && v.getRow() != null) {
                        resultVal = new Position(new ShapeRC(String.valueOf(v.getRow() - 1), String.valueOf(v.getCol())),
                                new ShapeRC(String.valueOf(v.getRow()), String.valueOf(v.getCol())));
                    }
                    if (from != null && from.length() >= 2) {
                        var p = ColumnUtils.getColRowBaseOnRefString(from, cols);
                        resultVal.setStart(p);
                        resultVal.setEnd(ShapeRC.builder().col(String.valueOf(p.getColNum() + 1)).row(String.valueOf(p.getRowNum() + 1)).build());
                    }
                    if (to != null && to.length() >= 2) {
                        var p = ColumnUtils.getColRowBaseOnRefString(to, cols);
                        p.setRow(String.valueOf(p.getRowNum() + 1));
                        p.setCol(String.valueOf(p.getColNum() + 1));

                        resultVal.setEnd(p);
                    }
                    checkboxSheetContent +=
                            "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"x14\"><control shapeId=\"" +
                                    shapeId +
                                    "\" r:id=\"rId" +
                                    (7 + i) +
                                    "\" name=\"" +
                                    v.getText() +
                                    "\"><controlPr defaultSize=\"0\" autoFill=\"0\" autoLine=\"0\" autoPict=\"0\"><anchor moveWithCells=\"1\"><from><xdr:col>" +
                                    resultVal.getStart().getCol() +
                                    "</xdr:col><xdr:colOff>19050</xdr:colOff><xdr:row>" +
                                    resultVal.getStart().getRowNum() +
                                    "</xdr:row><xdr:rowOff>19050</xdr:rowOff></from><to><xdr:col>" +
                                    resultVal.getEnd().getCol() +
                                    "</xdr:col><xdr:colOff>819150</xdr:colOff><xdr:row>" +
                                    resultVal.getEnd().getRow() +
                                    "</xdr:row><xdr:rowOff>0</xdr:rowOff></to></anchor></controlPr></control></mc:Choice></mc:AlternateContent>";
                    formRel +=
                            "<Relationship Id=\"rId" +
                                    (7 + i) +
                                    "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp\" " +
                                    "Target=\"../ctrlProps/ctrlProp" +
                                    checkboxForm.size() +
                                    ".xml\" />";
                    checkboxDrawingContent +=
                            "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" Requires=\"a14\"><xdr:twoCellAnchor editAs=\"oneCell\"><xdr:from><xdr:col>" +
                                    resultVal.getStart().getCol() +
                                    "</xdr:col><xdr:colOff>19050</xdr:colOff><xdr:row>" +
                                    resultVal.getStart().getRow() +
                                    "</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:from><xdr:to><xdr:col>" +
                                    resultVal.getEnd().getCol() +
                                    "</xdr:col><xdr:colOff>819150</xdr:colOff><xdr:row>" +
                                    resultVal.getEnd().getRow() +
                                    "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"" +
                                    shapeId +
                                    "\" name=\"" +
                                    v.getText() +
                                    "\" hidden=\"1\"><a:extLst><a:ext uri=\"\"><a14:compatExt spid=\"" +
                                    sId +
                                    "\"/></a:ext><a:ext uri=\"\"><a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"\"/></a:ext>" +
                                    "</a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr bwMode=\"auto\"><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm>" +
                                    "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr><xdr:txBody>" +
                                    "<a:bodyPr vertOverflow=\"clip\" wrap=\"square\" lIns=\"27432\" tIns=\"18288\" rIns=\"0\" bIns=\"18288\" anchor=\"ctr\" upright=\"1\"/>" +
                                    "<a:lstStyle/><a:p><a:pPr algn=\"l\" rtl=\"0\"><a:defRPr sz=\"1000\"/></a:pPr><a:r><a:rPr lang=\"en-US\" sz=\"800\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" baseline=\"0\"><a:solidFill>" +
                                    "<a:srgbClr val=\"000000\"/></a:solidFill><a:latin typeface=\"Segoe UI\"/><a:cs typeface=\"Segoe UI\"/></a:rPr><a:t>" +
                                    v.getText() +
                                    "</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent>";
                }
            }
            ExecutorService executorService = Executors.newFixedThreadPool(10);
            CompletableFuture<ImageOutput> backgroundImagePromise = null;
            if (sheetData.getBackgroundImage() != null) {
                if (xl_media_Folder == null) {
                    xl_media_Folder = xlFolder + "/" + "media";
                }
                var urlImg = sheetData.getBackgroundImage();
                backgroundImagePromise = CompletableFuture.supplyAsync(() -> {
                    ImageInput image;
                    if (fetchFunc != null) {
                        image = fetchFunc.apply(urlImg);
                    } else {
                        image = urlImg;
                    }
                    if (image == null) {
                        throw new Error("image not load");
                    }
                    String type;
                    if (image.getExtension() != null) {
                        type = image.getExtension().getName();
                    } else {
                        type = ImageInput.ImageExtension.PNG.getName();
                    }
                    int ref = imageCounter.getAndIncrement();
                    String name = "image" + ref + "." + type;


                    arrTypes.add(type);
                    return ImageOutput.builder(

                            ).name(name)
                            .type(type)
                            .image(image)
                            .ref(ref).build();
                }, executorService);
            }
            List<CompletableFuture<ImageOutput>> imagePromise = new ArrayList<>();

            if (sheetData.getImages() != null) {
                if (xl_media_Folder == null) {
                    xl_media_Folder = xlFolder + "/media";
                }
                var imgList = sheetData.getImages();
                for (AtomicInteger imgIndex = new AtomicInteger(0); imgIndex.get() < imgList.size(); imgIndex.incrementAndGet()) {
                    var getImage = CompletableFuture.supplyAsync(() -> {
                        var v = imgList.get(imgIndex.get());
                        ImageInput url = v.getImage();
                        ImageInput image;
                        if (fetchFunc != null) {
                            image = fetchFunc.apply(url);
                        } else {
                            image = url;
                        }
                        if (image == null) {
                            throw new Error("image not load");
                        }
                        String type;
                        if (image.getExtension() != null) {
                            type = image.getExtension().getName();
                        } else {
                            type = ImageInput.ImageExtension.PNG.getName();
                        }
                        arrTypes.add(type);
                        String name = "image" + imageCounter.getAndIncrement() + "." + type;

                        return ImageOutput.builder().image(image).type(type).index(imgIndex.get()).input(v).name(name).build();
                    }, executorService);
                    imagePromise.add(getImage);
                    try {
                        getImage.get();
                    } catch (InterruptedException e) {
                        throw new RuntimeException(e);
                    } catch (ExecutionException e) {
                        throw new RuntimeException(e);
                    }

                }
            }
            var sheetHeaderList = sheetData.getHeaders();
            var conditionStyleFunction = sheetData.getStyleCellCondition();
            var mergeDataFunction = sheetData.getMergeRowDataCondition();
            var commentFunction = sheetData.getCommentCondition();
            var multiStyleFunction = sheetData.getMultiStyleCondition();
            if (sheetHeaderList != null) {
                int colsLength = sheetHeaderList.size();
                String titleRow = "";
                if (sheetData.getTitle() != null) {
                    var title = sheetData.getTitle();
                    var commentTitle = title.getComment();
                    int top = title.getShiftTop() != null && title.getShiftTop() >= 0 ? title.getShiftTop() : 0;
                    int sL =
                            sheetData.getShiftLeft() != null && sheetData.getShiftLeft() >= 0
                                    ? sheetData.getShiftLeft()
                                    : 0;
                    int left =
                            title.getShiftLeft() != null && title.getShiftLeft() + sL >= 0
                                    ? title.getShiftLeft() + sL
                                    : sL;
                    int consommeRow = title.getConsommeRow() != null ? title.getConsommeRow() - 1 : 1;
                    int consommeCol = title.getConsommeCol() != null ? title.getConsommeCol() : colsLength;
                    String height =
                            consommeRow == 0 && title.getHeight() != null
                                    ? " ht=\"" + title.getHeight() + "\" customHeight=\"1\" "
                                    : "";
                    String tStyle = title.getStyleId() != null ? title.getStyleId() : "titleStyle";
                    String refString = cols.get(left) + (rowCount + top);
                    mergesCellArray.add(
                            refString +
                                    ":" +
                                    cols.get(left + consommeCol - 1) +
                                    (rowCount + consommeRow + top)
                    );
                    if (commentTitle != null) {
                        hasComment = true;
                        var commentObj = commentConvertor(
                                commentTitle,
                                styleMapper.getCommentSyntax(),
                                defaultCommentStyle
                        );
                        int authorId = commentAuthor.size();
                        if (Utils.booleanCheck(commentObj.getHasAuthor())
                                && commentObj.getAuthor() != null) {
                            String auth = commentObj.getAuthor();
                            int indexAuth = commentAuthor.indexOf(auth);
                            if (indexAuth < 0) {
                                commentAuthor.add(auth);
                            } else {
                                authorId = indexAuth;
                            }
                        }
                        shapeCommentRowCol.add(ShapeRC.builder()
                                .row(String.valueOf(rowCount + top - 1))
                                .col(String.valueOf(left))
                                .build()
                        );
                        commentString += CommentUtil.generateCommentTag(
                                refString,
                                commentObj.getCommentStr(),
                                commentObj.getCommentStyle(),
                                authorId
                        );
                    }
                    if (title.getText() != null) {
                        rowMap.put(rowCount + top, RowMap.builder()
                                .startTag("<row r=\"" +
                                        (rowCount + top) +
                                        "\" " +
                                        height +
                                        " spans=\"1:" +
                                        Math.max((left + consommeCol - 1), 1) +
                                        "\">")
                                .details("<c r=\"" +
                                        refString +
                                        "\" " +
                                        (data.getStyles().get(tStyle) != null
                                                ? " s=\"" + data.getStyles().get(tStyle).getIndex() + "\" "
                                                : "") +
                                        " t=\"s\"><v>" +
                                        sharedStringIndex +
                                        "</v></c>")
                                .endTag("</row>")
                                .build());

                        titleRow +=
                                "<row r=\"" +
                                        (rowCount + top) +
                                        "\" " +
                                        height +
                                        " spans=\"1:" +
                                        Math.max((left + consommeCol - 1), 1) +
                                        "\">";
                        titleRow +=
                                "<c r=\"" +
                                        refString +
                                        "\" " +
                                        (data.getStyles().get(tStyle) != null
                                                ? " s=\"" + data.getStyles().get(tStyle).getIndex() + "\" "
                                                : "") +
                                        " t=\"s\"><v>" +
                                        sharedStringIndex +
                                        "</v></c>";
                        titleRow += "</row>";
                        sharedStringIndex++;
                        sharedStringMap.put(title.getText(), title.getText());
                        if (title.getMultiStyleValue() != null) {
                            sharedString += MultiValueUtil.generateMultiStyleByArray(
                                    title.getMultiStyleValue(),
                                    styleMapper.getCommentSyntax(),
                                    tStyle
                            );
                        } else {
                            sharedString +=
                                    "<si><t>" + MultiValueUtil.specialCharacterConverter(title.getText()) + "</t></si>";
                        }
                    }
                    rowCount += top + consommeRow + 1;
                }
                var headerStyleKey = sheetData.getHeaderStyleKey();
                int shiftCount = 0;
                if (sheetData.getShiftLeft() != null && sheetData.getShiftLeft() >= 0) {
                    shiftCount = sheetData.getShiftLeft();
                }
                if (asTable != null) {
                    sheetDataTableColumns +=
                            "<tableColumns count=\"" + colsLength + "\">";
                }
                sheetDimensions.setStart(cols.get(shiftCount) + rowCount);
                sheetDimensions.setEnd(
                        cols.get(Math.max(shiftCount + colsLength - 1, 0)) +
                                (rowCount + sheetData.getData().size()));
                int innerIndexLoop = 0;
                for (Header v : sheetHeaderList) {
                    int innerIndex = innerIndexLoop;
//                    var v = sheetHeaderList.get(innerIndex);
                    if (asTable != null) {
                        sheetDataTableColumns +=
                                "<tableColumn id=\"" +
                                        (innerIndex + 1) +
                                        "\" name=\"" +
                                        v.getText() +
                                        "\"/>";
                    }
                    if (shiftCount > 0) {
                        innerIndex += shiftCount;
                    }
                    if (v.getFormula() != null) {
                        headerFormula.add(innerIndex);
                    }
                    if (v.getConditionalFormatting() != null && addCF) {
                        headerConditionalFormatting.add(innerIndex);
                    }
                    objKey.add(v.getMethod());
                    if (
                            mergeDataFunction != null
                    ) {
                        var result = mergeDataFunction.apply(
                                new MergeRowDataConditionInputFunction(v,
                                        null,
                                        innerIndex,
                                        true)
                        );
                        if (result) {
                            mergeRowConditionMap.put(cols.get(innerIndex), new MergeRowConditionMap.Condition(true, rowCount));
                        }
                    }
                    if (
                            conditionStyleFunction != null
                    ) {
                        headerStyleKey =
                                Optional.ofNullable(conditionStyleFunction.apply(
                                        new StyleCellConditionInputFunction(
                                                v,
                                                v,
                                                rowCount,
                                                innerIndex,
                                                true,
                                                styleKeys)
                                )).orElse(headerStyleKey);
                    }
                    if (v.getSize() != null && v.getSize() > 0) {
                        sheetSizeString +=
                                "<col min=\"" +
                                        (innerIndex + 1) +
                                        "\" max=\"" +
                                        (innerIndex + 1) +
                                        "\" width=\"" +
                                        v.getSize() +
                                        "\" customWidth=\"1\" />";
                    }
                    if (Utils.booleanCheck(sheetData.getWithoutHeader())) {
                        continue;
                    }
                    String refString = cols.get(innerIndex) + rowCount;
                    if (commentFunction != null) {
                        var checkCommentCondition = commentFunction.apply(
                                new CommentConditionFunctionInput(v,
                                        null,
                                        v.getMethod(),
                                        rowCount,
                                        innerIndex,
                                        true)
                        );
                        if (
                                checkCommentCondition != null
                        ) {
                            v.setComment(checkCommentCondition);
                        }
                    }
                    if (v.getComment() != null) {
                        hasComment = true;
                        var commentObj = commentConvertor(
                                v.getComment(),
                                styleMapper.getCommentSyntax(),
                                defaultCommentStyle
                        );
                        int authorId = commentAuthor.size();
                        if (Utils.booleanCheck(commentObj.getHasAuthor()) && commentObj.getAuthor() != null) {
                            String auth = commentObj.getAuthor();
                            int indexAuth = commentAuthor.indexOf(auth);
                            if (indexAuth < 0) {
                                commentAuthor.add(auth);
                            } else {
                                authorId = indexAuth;
                            }
                        }
                        shapeCommentRowCol.add(ShapeRC.builder()
                                .row(String.valueOf(rowCount - 1))
                                .col(String.valueOf(innerIndex)).build()
                        );
                        commentString += CommentUtil.generateCommentTag(
                                refString,
                                commentObj.getCommentStr(),
                                commentObj.getCommentStyle(),
                                authorId
                        );
                    }

                    var formula = formulaSheetObj.get(refString);
                    if (formula != null) {
                        var f = FormulaUtil.generateCellRowCol(
                                refString,
                                formula,
                                sheetDataId,
                                data.getStyles()
                        );
                        if (Utils.booleanCheck(f.getNeedCalcChain())) {
                            needCalcChain = true;
                            calcChainValue += f.getChainCell();
                        }
                        sheetDataString += f.getCell();
                        formulaSheetObj.remove(refString);
                    } else {
                        sheetDataString +=
                                "<c r=\"" +
                                        cols.get(innerIndex) +
                                        rowCount +
                                        "\" " +
                                        (headerStyleKey != null && data.getStyles() != null && data.getStyles().get(headerStyleKey) != null
                                                ? " s=\"" + data.getStyles().get(headerStyleKey).getIndex() + "\" "
                                                : "") +
                                        " " +
                                        "t=\"s\"><v>" +
                                        sharedStringIndex +
                                        "</v></c>";
                        if (multiStyleFunction != null) {
                            var multi = multiStyleFunction.apply(
                                    new MultiStyleConditionInputFunction(v,
                                            null,
                                            v.getMethod(),
                                            rowCount,
                                            innerIndex,
                                            true)
                            );
                            if (multi != null) {
                                v.setMultiStyleValue(multi);
                            }
                        }
                        if (v.getMultiStyleValue() != null) {
                            sharedString += MultiValueUtil.generateMultiStyleByArray(
                                    v.getMultiStyleValue(),
                                    styleMapper.getCommentSyntax(),
                                    headerStyleKey != null ? headerStyleKey : ""
                            );
                        } else {
                            sharedString +=
                                    "<si><t>" + MultiValueUtil.specialCharacterConverter(v.getText()) + "</t></si>";
                        }
                        sharedStringMap.put(v.getText(), v.getText());

                        sharedStringIndex++;
                    }
                    innerIndexLoop++;
                }
                // sheetData.headers.forEach((v, innerIndex) => );
                if (asTable != null) {
                    sheetDataTableColumns += "</tableColumns>";
                }
                if (!Utils.booleanCheck(sheetData.getWithoutHeader())) {
                    String rowTag =
                            "<row r=\"" +
                                    rowCount +
                                    "\" spans=\"1:" +
                                    Math.max(colsLength, 1) +
                                    "\" " +
                                    (sheetData.getHeaderHeight() != null
                                            ? "ht=\"" + sheetData.getHeaderHeight() + "\" customHeight=\"1\""
                                            : "") +
//                                   //TODO ?? (sheetData.getHeaderRowOption()!=null
//                                            ? Object.keys(sheetData.headerRowOption).reduce((res, curr) => {
//                    return (
//                            res +
//                                    " " +
//                                    curr +
//                                    "=\"" +
//                                    sheetData.headerRowOption![curr as keyof object] +
//                            "\" "
//                );
//              }, "  ")
//            : "")+

                                    ">";
                    rowMap.put(rowCount, RowMap.builder().
                            startTag(rowTag)
                            .endTag("</row>")
                            .details(sheetDataString)
                            .build()
                    );
                    sheetDataString = titleRow + rowTag + sheetDataString + "</row>";
                    rowCount++;
                } else {
                    sheetDataString += titleRow;
                }
                if (sheetData.getData() != null) {
                    Method keyOutline = null;
                    Method keyHidden = null;
                    Method keyHeight = null;
                    if (sheetData.getMapSheetDataOption() != null) {
                        if (sheetData.getMapSheetDataOption().getOutlineLevel() != null) {
                            keyOutline = sheetData.getMapSheetDataOption().getOutlineLevel();
                        }
                        if (sheetData.getMapSheetDataOption().getHidden() != null) {
                            keyHidden = sheetData.getMapSheetDataOption().getHidden();
                        }
                        if (sheetData.getMapSheetDataOption().getHeight() != null) {
                            keyHeight = sheetData.getMapSheetDataOption().getHeight();
                        }
                    }

                    int rowLength = sheetData.getData().size();
                    var dataList = sheetData.getData();
                    for (int innerIndex = 0; innerIndex < rowLength; innerIndex++) {
                        var mData = dataList.get(innerIndex);
                        DataOption mDataOption = null;
                        if (mData instanceof DataOption) {
                            mDataOption = (DataOption) mData;
                        }
                        if (mData instanceof MergeConfig) {
                            MergeConfig asMergeConfig = (MergeConfig) mData;
                            var typeList = asMergeConfig.getMergeType();
                            var startList = asMergeConfig.getMergeStart();
                            var valueList = asMergeConfig.getMergeValue();
                            if (typeList != null) {
                                for (int iIndex = 0; iIndex < typeList.size(); iIndex++) {
                                    var mergeType = typeList.get(iIndex);
                                    var mergeStart = startList.get(iIndex);
                                    var mergeValue = valueList.get(index);
                                    String mergeStr;
                                    if (mergeType == MergeConfig.MergeType.BOTH) {
                                        mergeStr =
                                                cols.get(mergeStart) +
                                                        rowCount +
                                                        ":" +
                                                        cols.get(mergeStart + mergeValue.get(1)) +
                                                        (rowCount + mergeValue.get(0));
                                    } else {
                                        if (mergeType == MergeConfig.MergeType.COL) {
                                            mergeStr =
                                                    cols.get(mergeStart) +
                                                            rowCount +
                                                            ":" +
                                                            cols.get(mergeStart + mergeValue.get(0)) +
                                                            rowCount;
                                        } else {
                                            mergeStr =
                                                    cols.get(mergeStart) +
                                                            rowCount +
                                                            ":" +
                                                            cols.get(mergeStart) +
                                                            (rowCount + mergeValue.get(0));
                                        }
                                    }

                                    mergesCellArray.add(mergeStr);
                                }
                            }
                        }
                        String rowStyle = "";
                        if (mDataOption != null) {
                            rowStyle = mDataOption.getRowStyle();
                        }

                        Number heightVal = null;
                        if (keyHeight != null) {
                            heightVal = (Number) keyHeight.invoke(mData);
                        } else {
                            if (mDataOption != null && mDataOption.getHeight() != null) {
                                heightVal = mDataOption.getHeight();
                            }
                        }
                        Number outlineVal = null;
                        if (keyOutline != null) {
                            outlineVal = (Number) keyOutline.invoke(mData);
                        } else {
                            if (mDataOption != null && mDataOption.getOutlineLevel() != null) {
                                outlineVal = mDataOption.getOutlineLevel();
                            }
                        }
                        Boolean hiddenVal = null;
                        if (keyHidden != null) {
                            hiddenVal = (Boolean) keyHidden.invoke(mData);
                        } else {
                            if (mDataOption != null && mDataOption.getOutlineLevel() != null) {
                                hiddenVal = mDataOption.getHidden();
                            }
                        }
                        String rowTagStart =
                                "<row r=\"" +
                                        rowCount +
                                        "\" spans=\"1:" +
                                        Math.max(colsLength, 1) +
                                        "\" " +
                                        (heightVal != null
                                                ? "ht=\"" + heightVal + "\" customHeight=\"1\""
                                                : "") +
                                        (outlineVal != null
                                                ? " outlineLevel=\"" + outlineVal + "\""
                                                : "") +
                                        (hiddenVal != null ? " hidden=\"" + (hiddenVal ? "1" : "0") + "\"" : "") +
                                        " >";
                        sheetDataString += rowTagStart;
                        String rowDataString = "";

                        for (int keyIndexLoop = 0, objSize = objKey.size(); keyIndexLoop < objSize; keyIndexLoop++) {
                            int keyIndex = keyIndexLoop;
                            Method key = objKey.get(keyIndex);

                            if (shiftCount > 0) {
                                keyIndex += shiftCount;
                            }

                            Object dataEl;
                            try {
                                dataEl = key.invoke(mData);
                            } catch (Exception ex) {
                                dataEl = "";
                            }
                            if (dataEl == null) {
                                dataEl = "null";
                            }
                            if (dataEl instanceof Boolean) {
                                dataEl = dataEl.toString();
                            }
                            Number dataElAsNumber = null;
                            if (Utils.booleanCheck(sheetData.getConvertStringToNumber())) {
                                if (dataEl instanceof Number) {
                                    dataElAsNumber = (Number) dataEl;
                                }
                            }

                            String cellStyle = rowStyle;
                            if (
                                    conditionStyleFunction != null
                            ) {
                                cellStyle =
                                        Optional.ofNullable(
                                                conditionStyleFunction.apply(
                                                        new StyleCellConditionInputFunction(dataEl,
                                                                mData,
                                                                rowCount,
                                                                keyIndex,
                                                                false,
                                                                styleKeys)
                                                )
                                        ).orElse(rowStyle);
                            }
                            if (
                                    mergeDataFunction != null
                            ) {
                                Boolean result = mergeDataFunction.apply(
                                        new MergeRowDataConditionInputFunction(dataEl,
                                                key,
                                                keyIndex,
                                                false)
                                );
                                String columnKey = cols.get(keyIndex);

                                var item = mergeRowConditionMap.get(columnKey);
                                if (Utils.booleanCheck(result)) {
                                    if (item == null || (item != null && !(Utils.booleanCheck(item.getInProgress())))) {
                                        mergeRowConditionMap.put(columnKey, MergeRowConditionMap.Condition.builder()
                                                .inProgress(true).start(rowCount)
                                                .build());
                                    }
                                } else {
                                    if (item != null && Utils.booleanCheck(item.getInProgress())) {
                                        mergesCellArray.add(
                                                columnKey + item.getStart() + ":" + columnKey + (rowCount - 1)
                                        );

                                        mergeRowConditionMap.put(columnKey, MergeRowConditionMap.Condition.builder()
                                                .inProgress(false)
                                                .start(-1)
                                                .build());
                                    }
                                }
                            }

                            String refString = cols.get(keyIndex) + rowCount;
                            if (commentFunction != null) {
                                Comment checkCommentCondition = commentFunction.apply(
                                        new CommentConditionFunctionInput(dataEl,
                                                mData,
                                                key,
                                                rowCount,
                                                keyIndex,
                                                false)
                                );
                                if (
                                        checkCommentCondition != null
                                ) {
                                    if (mDataOption != null)
                                        mDataOption.addComment(key, checkCommentCondition);
                                }
                            }
                            if (mDataOption != null && mDataOption.getComment() != null && mDataOption.getComment().size() > 0) {
                                var cellComment = mDataOption.getComment().get(key);
                                if (cellComment != null) {
                                    hasComment = true;
                                    var commentObj = commentConvertor(
                                            cellComment,
                                            styleMapper.getCommentSyntax(),
                                            defaultCommentStyle
                                    );
                                    if (
                                            Utils.booleanCheck(commentObj.getHasAuthor()) &&
                                                    commentObj.getAuthor() != null
                                    ) {
                                        commentAuthor.add(commentObj.getAuthor());
                                    }
                                    shapeCommentRowCol.add(
                                            ShapeRC.builder()
                                                    .row(String.valueOf(rowCount - 1)).col(String.valueOf(keyIndex))
                                                    .build()
                                    );
                                    var authorId = commentAuthor.size();
                                    if (Utils.booleanCheck(commentObj.getHasAuthor()) &&
                                            commentObj.getAuthor() != null
                                    ) {
                                        var auth = commentObj.getAuthor();
                                        int indexCom = commentAuthor.indexOf(auth);
                                        if (indexCom < 0) {
                                            commentAuthor.add(auth);
                                        } else {
                                            authorId = indexCom;
                                        }
                                    }
                                    commentString += CommentUtil.generateCommentTag(
                                            refString,
                                            commentObj.getCommentStr(),
                                            commentObj.getCommentStyle(),
                                            authorId
                                    );
                                }
                            }
                            var formula = formulaSheetObj.get(refString);
                            if (formula != null) {
                                var f = FormulaUtil.generateCellRowCol(refString, formula, sheetDataId, styles);
                                if (Utils.booleanCheck(f.getNeedCalcChain())) {
                                    needCalcChain = true;
                                    calcChainValue += f.getChainCell();
                                }
                                sheetDataString += f.getCell();
                                rowDataString += f.getCell();
//                    delete formulaSheetObj![refString];
                                formulaSheetObj.remove(refString);
                            } else {
                                if (dataEl instanceof String) {
                                    String dataElString = (String) dataEl;
                                    var hasStyleKey = data.getStyles().containsKey(cellStyle);
                                    String localCell =
                                            "<c r=\"" +
                                                    cols.get(keyIndex) +
                                                    rowCount +
                                                    "\" t=\"s\" " +
                                                    (cellStyle != null && hasStyleKey
                                                            ? "s=\"" + data.getStyles().get(cellStyle).getIndex() + "\""
                                                            : "") +
                                                    "><v>" +
                                                    sharedStringIndex +
                                                    "</v></c>";
                                    rowDataString += localCell;
                                    sheetDataString += localCell;
                                    if (multiStyleFunction != null) {
                                        var multi = multiStyleFunction.apply(
                                                new MultiStyleConditionInputFunction(dataEl,
                                                        mData,
                                                        key,
                                                        rowCount,
                                                        keyIndex,
                                                        false)
                                        );
                                        if (multi != null) {
                                            if (
                                                    mDataOption != null
                                            ) {
                                                mDataOption.addMultiStyle(key, multi);
                                            }
                                        }
                                    }
                                    if (
                                            mDataOption != null &&
                                                    mDataOption.getMultiStyleValueKeyExist(key)

                                    ) {
                                        sharedString += MultiValueUtil.generateMultiStyleByArray(
                                                mDataOption.getMultiStyleValue().get(key),
                                                styleMapper.getCommentSyntax(),
                                                cellStyle != null ? cellStyle : ""
                                        );
                                    } else {
                                        sharedString +=
                                                "<si><t>" + MultiValueUtil.specialCharacterConverter(dataElString) + "</t></si>";
                                    }
                                    sharedStringMap.put(dataElString, dataElString);
                                    sharedStringIndex++;
                                } else {
                                    var hasStyleKey = data.getStyles().containsKey(cellStyle);
                                    String localCell =
                                            "<c r=\"" +
                                                    cols.get(keyIndex) +
                                                    rowCount +
                                                    "\" " +
                                                    (cellStyle != null && hasStyleKey
                                                            ? "s=\"" + data.getStyles().get(cellStyle).getIndex() + "\""
                                                            : "") +
                                                    "><v>" +
                                                    (dataElAsNumber != null ? dataElAsNumber : dataEl) +
                                                    "</v></c>";
                                    sheetDataString += localCell;
                                    rowDataString += localCell;
                                }
                            }
                        }
                        if (rowLength - 1 == innerIndex) {
                            var keySetMerge = mergeRowConditionMap.getMergeConditionMap().keySet();
                            for (String keyOfmerge : keySetMerge) {
                                var elementOfMerge = mergeRowConditionMap.getMergeConditionMap().get(keyOfmerge);
                                if (Utils.booleanCheck(elementOfMerge.getInProgress())) {
                                    mergesCellArray.add(
                                            keyOfmerge + elementOfMerge.getStart() + ":" + keyOfmerge + rowCount
                                    );
                                }
                            }
                        }
                        rowMap.put(rowCount, RowMap.builder()
                                .startTag(rowTagStart)
                                .endTag("</row>")
                                .details(rowDataString)
                                .build());
                        rowCount++;
                        sheetDataString += "</row>";
                    }
                    if (sheetData.getSortAndFilter() != null) {
                        var sortAndFilterOption = sheetData.getSortAndFilter();
                        if (sortAndFilterOption.getMode() == SortAndFilter.Mode.ALL) {
                            sheetSortFilter +=
                                    "<autoFilter ref=\"A1:" +
                                            cols.get(colsLength - 1) +
                                            (rowCount - 1) +
                                            "\" />";
                        } else {
                            if (
                                    sortAndFilterOption.getRef() != null &&
                                            sortAndFilterOption.getRef().length() > 0
                            ) {
                                sheetSortFilter +=
                                        "<autoFilter ref=\"" + sortAndFilterOption.getRef() + "\" />";
                            }
                        }
                    }
                }
                if (headerFormula.size() > 0) {
                    for (var v : headerFormula) {
                        var shiftLeftValue = sheetData.getShiftLeft() != null ? sheetData.getShiftLeft() : 0;
                        var header = sheetData.getHeaders().get(v.intValue() - shiftLeftValue);
                        var columnRef = cols.get(v.intValue());
                        formulaSheetObj.put(columnRef + rowCount,
                                FormulaSetting.builder()
                                        .end(columnRef + (rowCount - 1))
                                        .type(header.getFormula().getType())
                                        .end(columnRef + (rowCount - 1))
                                        .start(Utils.booleanCheck(sheetData.getWithoutHeader()) ? columnRef + "1" : columnRef + "2")
                                        .styleId(header.getFormula().getStyleId())
                                        .build()
                        );
                    }
                }
                if (headerConditionalFormatting.size() > 0 && addCF) {
                    for (var v : headerConditionalFormatting) {
                        var header = sheetData.getHeaders().get(v.intValue());
                        if (header.getConditionalFormatting() == null) {
                            continue;
                        }
                        var ob = header.getConditionalFormatting();
                        ob.setStart(Utils.booleanCheck(sheetData.getWithoutHeader()) ? cols.get(v.intValue()) + "1" : cols.get(v.intValue()) + "2");
                        ob.setEnd(cols.get(v.intValue()) + (rowCount - 1));
                        conditionalFormatting.add(ob);
                    }
                }
                if (formulaSheetObj != null) {
                    List<String> remindFormulaKey = new ArrayList<>(formulaSheetObj.keySet());
                    remindFormulaKey.sort((a, b) -> a.compareTo(b) > 0 ? 1 : -1);

                    if (remindFormulaKey.size() > 0) {
                        HashMap<Number, String> rF = new HashMap<>();
                        for (var v : remindFormulaKey) {
                            var f = FormulaUtil.generateCellRowCol(
                                    v,
                                    formulaSheetObj.get(v),
                                    sheetDataId,
                                    data.getStyles()
                            );
                            if (Utils.booleanCheck(f.getNeedCalcChain())) {
                                needCalcChain = true;
                                calcChainValue += f.getChainCell();
                            }
                            if (!rF.containsKey(f.getRow())) {
                                rF.put(f.getRow(), f.getCell());
                            } else {
                                rF.put(f.getRow(), rF.get(f.getRow()) + f.getCell());
                            }
                        }
                        var rFKeySet = new ArrayList<>(rF.keySet());
                        rFKeySet.sort((a, b) -> a.intValue() > b.intValue() ? 1 : -1);
                        for (var val : rFKeySet) {
                            var l = rF.get(val);
                            var rowDataMap = rowMap.get(val);
                            if (rowDataMap != null) {
                                String body =
                                        rowDataMap.getStartTag() +
                                                rowDataMap.getDetails() +
                                                l +
                                                rowDataMap.getEndTag();

                                sheetDataString = sheetDataString.replaceFirst(rowDataMap.getStartTag() + "[\\n\\s\\S]*?</row>", body);
                            } else {
                                sheetDataString +=
                                        "<row r=\"" +
                                                val +
                                                "\" spans=\"1:" +
                                                Math.max(colsLength, 1) +
                                                "\"  >" +
                                                l +
                                                "</row>";
                                rowMap.put(val.intValue(), RowMap.builder()
                                        .startTag("<row r=\"" + val + "\" spans=\"1:" + Math.max(colsLength, 1) + "\"  >")
                                        .endTag("</row>")
                                        .details(l)
                                        .build()
                                );
                            }
                        }

                    }
                }
            }

            if (index > 0) {
                sheetContentType +=
                        "<Override" +
                                "    ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"" +
                                "    PartName=\"/xl/worksheets/sheet" +
                                (index + 1) +
                                ".xml\" />";
            }
            String shName = sheetData.getName() != null ? sheetData.getName() : "sheet" + (index + 1);
            String shState = sheetData.getState() != null ? sheetData.getState().getLabel() : "visible";
            workbookString +=
                    "<sheet state=\"" +
                            shState +
                            "\" name=\"" +
                            shName +
                            "\" sheetId=\"" +
                            (index + 1) +
                            "\" r:id=\"rId" +
                            (indexId + 1) +
                            "\" />";
            workbookRelString +=
                    "<Relationship Id=\"rId" +
                            (indexId + 1) +
                            "\"" +
                            " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\"" +
                            " Target=\"worksheets/sheet" +
                            (index + 1) +
                            ".xml\" />";
            sheetNameApp += "<vt:lpstr>" + ("sheet" + (index + 1)) + "</vt:lpstr>";
            if (Utils.booleanCheck(sheetData.getSelected())) {
                selectedAdded = true;
                activeTabIndex = index;
            }
            var filterMode = sheetData.getSortAndFilter() != null ? "filterMode=\"1\"" : "";
            var backgroundImageRef = -1;

            if (backgroundImagePromise != null) {
                ImageOutput res = null;
                try {
                    res = backgroundImagePromise.get();
                } catch (InterruptedException e) {
                    throw new RuntimeException(e);
                } catch (ExecutionException e) {
                    throw new RuntimeException(e);
                }
                backgroundImageRef = res.getRef();
                resultAssetsMap.put(xl_media_Folder + "/" + res.getName(), res);

            }
            boolean hasImages = false;
            String drawersContent = "";
            String drawersRels = "";
            if (imagePromise.size() > 0) {
                CompletableFuture<Void> allImage = CompletableFuture.allOf(imagePromise.toArray(new CompletableFuture[imagePromise.size()]));

                try {
                    allImage.get();

                } catch (InterruptedException e) {
                    throw new RuntimeException(e);
                } catch (ExecutionException e) {
                    throw new RuntimeException(e);
                }
                hasImages = true;

                String drawerStr = "";
                int imagIndex = 0;
                for (var resultProcess : imagePromise) {
                    ImageOutput val = null;
                    try {
                        val = resultProcess.get();
                    } catch (InterruptedException e) {
                        throw new RuntimeException(e);
                    } catch (ExecutionException e) {
                        throw new RuntimeException(e);
                    }
// TODO gif problem not moving
                    int processIndex = imagIndex + 1;
                    var name = val.getName();
                    String from = val.getInput().getFrom();
                    String to = val.getInput().getTo();
                    var margin = val.getInput().getMargin();
                    var type = val.getInput().getType();
                    var extent = val.getInput().getExtent();
                    if (extent == null) {
                        extent = Sheet.ImageTypes.Extend.builder()
                                .cx(200000)
                                .cy(200000)
                                .build();
                    }
                    ResultImageProcess result = ResultImageProcess.builder()
                            .start(
                                    ResultImageProcess.Setting.builder()
                                            .col(0).row(0).mL(0).mT(0)
                                            .build()
                            )
                            .end(ResultImageProcess.Setting.builder()
                                    .col(1).row(1).mR(0).mB(0)
                                    .build())
                            .build();
                    if (from != null && from.length() >= 2) {
                        var p = ColumnUtils.getColRowBaseOnRefString(from, cols);
                        result.setStart(p);
                        result.setEndPlusOne(p);
                    }
                    if (to != null && to.length() >= 2) {
                        var p = ColumnUtils.getColRowBaseOnRefString(to, cols);

                        result.setEndPlusOne(p);
                    }
                    if (margin != null) {
                        result.setMargin(margin);
                    }
                    if (type == Sheet.ImageTypes.Type.ONE) {
                        drawersContent +=
                                "<xdr:oneCellAnchor>" +
                                        "<xdr:from>" +
                                        "<xdr:col>" +
                                        result.getStart().getCol() +
                                        "</xdr:col>" +
                                        "<xdr:colOff>" +
                                        result.getStart().getMT() +
                                        "</xdr:colOff>" +
                                        "<xdr:row>" +
                                        result.getStart().getRow() +
                                        "</xdr:row>" +
                                        "<xdr:rowOff>" +
                                        result.getStart().getML() +
                                        "</xdr:rowOff>" +
                                        "</xdr:from>" +
                                        "<xdr:ext cx=\"" +
                                        extent.getCx() +
                                        "\" cy=\"" +
                                        extent.getCy() +
                                        "\"/>" +
                                        "<xdr:pic>" +
                                        "<xdr:nvPicPr>" +
                                        "<xdr:cNvPr id=\"" +
                                        processIndex +
                                        "\" name=\"Picture " +
                                        processIndex +
                                        "\">" +
                                        "</xdr:cNvPr>" +
                                        "<xdr:cNvPicPr preferRelativeResize=\"0\" />" +
                                        "</xdr:nvPicPr>" +
                                        "<xdr:blipFill>" +
                                        "<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId" +
                                        processIndex +
                                        "\">" +
                                        "</a:blip>" +
                                        "<a:stretch>" +
                                        "<a:fillRect />" +
                                        "</a:stretch>" +
                                        "</xdr:blipFill>" +
                                        "<xdr:spPr>" +
                                        "<a:prstGeom prst=\"rect\">" +
                                        "<a:avLst />" +
                                        "</a:prstGeom>" +
                                        "<a:noFill />" +
                                        "</xdr:spPr>" +
                                        "</xdr:pic>" +
                                        "<xdr:clientData />" +
                                        "</xdr:oneCellAnchor>";
                    } else {
                        drawersContent +=
                                "<xdr:twoCellAnchor editAs=\"oneCell\">" +
                                        "<xdr:from>" +
                                        "<xdr:col>" +
                                        result.getStart().getCol() +
                                        "</xdr:col>" +
                                        "<xdr:colOff>" +
                                        result.getStart().getMT() +
                                        "</xdr:colOff>" +
                                        "<xdr:row>" +
                                        result.getStart().getRow() +
                                        "</xdr:row>" +
                                        "<xdr:rowOff>" +
                                        result.getStart().getML() +
                                        "</xdr:rowOff>" +
                                        "</xdr:from>" +
                                        "<xdr:to>" +
                                        "<xdr:col>" +
                                        result.getEnd().getCol() +
                                        "</xdr:col>" +
                                        "<xdr:colOff>" +
                                        result.getEnd().getMB() +
                                        "</xdr:colOff>" +
                                        "<xdr:row>" +
                                        result.getEnd().getRow() +
                                        "</xdr:row>" +
                                        "<xdr:rowOff>" +
                                        result.getEnd().getMR() +
                                        "</xdr:rowOff>" +
                                        "</xdr:to>" +
                                        "<xdr:pic>" +
                                        "<xdr:nvPicPr>" +
                                        "<xdr:cNvPr id=\"" +
                                        processIndex +
                                        "\" name=\"Picture " +
                                        processIndex +
                                        "\">" +
                                        "</xdr:cNvPr>" +
                                        "<xdr:cNvPicPr preferRelativeResize=\"0\" />" +
                                        "</xdr:nvPicPr>" +
                                        "<xdr:blipFill>" +
                                        "<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId" +
                                        processIndex +
                                        "\">" +
                                        "</a:blip>" +
                                        "<a:stretch>" +
                                        "<a:fillRect />" +
                                        "</a:stretch>" +
                                        "</xdr:blipFill>" +
                                        "<xdr:spPr>" +
                                        "<a:prstGeom prst=\"rect\">" +
                                        "<a:avLst />" +
                                        "</a:prstGeom>" +
                                        "<a:noFill />" +
                                        "</xdr:spPr>" +
                                        "</xdr:pic>" +
                                        "<xdr:clientData />" +
                                        "</xdr:twoCellAnchor>";
                    }
                    resultAssetsMap.put(xl_media_Folder + "/" + name, val);
                    drawerStr +=
                            "<Relationship Id=\"rId" +
                                    processIndex +
                                    "\" " +
                                    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" " +
                                    "Target=\"../media/" +
                                    name +
                                    "\" />";

                    imagIndex++;
                }
                drawersRels =
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                drawerStr +
                                "</Relationships>";
            }
            executorService.shutdown();
            Set<String> mergesCellArraySet = new HashSet<>(mergesCellArray);
            String cFDataString = "";
            int priorityCounter = 1;
            if (conditionalFormatting.size() > 0 && addCF) {
                String cf = "";
                for (var cu : conditionalFormatting) {
                    if (cu.getType() == ConditionalFormattingOption.Type.CELLS) {
                        if (cu.getOperator() == ConditionalFormattingOption.ConditionalFormattingCellsOperation.CT) {
                            int styleIndexOfCondition = 0;
                            if (cu.getStyleId() != null) {
                                if (cFMapIndex.containsKey(cu.getStyleId())) {
                                    styleIndexOfCondition = cFMapIndex.get(cu.getStyleId()).intValue();
                                }
                            }
                            cf +=
                                    "<conditionalFormatting sqref=\"" +
                                            cu.getStart() +
                                            ":" +
                                            cu.getEnd() +
                                            "\">" +
                                            "<cfRule type=\"containsText\" dxfId=\"" +
                                            (styleIndexOfCondition) +
                                            "\" priority=\"" +
                                            (cu.getPriority() != null ? cu.getPriority() : priorityCounter++) +
                                            "\"  operator=\"containsText\" text=\"" +
                                            cu.getValue() +
                                            "\"><formula>NOT(ISERROR(SEARCH(\"" +
                                            cu.getValue() +
                                            "\"," +
                                            cu.getStart() +
                                            ")))</formula></cfRule></conditionalFormatting>"
                            ;
                            continue;
                        }
                        String operationType = ((ConditionalFormattingOption.ConditionalFormattingCellsOperation) cu.getOperator()).getOperationType();
                        if (
                                cu.getOperator() == null || !operatorMap.containsKey(operationType)
                        ) {

                            continue;
                        }
                        int styleIndexOfCondition = 0;
                        if (cu.getStyleId() != null) {
                            if (cFMapIndex.containsKey(cu.getStyleId())) {
                                styleIndexOfCondition = cFMapIndex.get(cu.getStyleId()).intValue();
                            }
                        }

                        cf +=
                                "<conditionalFormatting sqref=\"" +
                                        cu.getStart() +
                                        ":" +
                                        cu.getEnd() +
                                        "\"><cfRule type=\"cellIs\" dxfId=\"" +
                                        (styleIndexOfCondition) +
                                        "\" priority=\"" +
                                        (cu.getPriority() != null ? cu.getPriority() : priorityCounter++) +
                                        "\" operator=\"" +
                                        operatorMap.get(operationType) +
                                        "\">" +
                                        (cu.getValue() instanceof List<?>
                                                ? ((List<String>) cu.getValue()).stream().reduce("", (rC, cr) -> rC + "<formula>" + cr + "</formula>")
                                                : "<formula>" + cu.getValue() + "</formula>") +
                                        "</cfRule></conditionalFormatting>"
                        ;
                        continue;
                    } else if (cu.getType() == ConditionalFormattingOption.Type.TOP) {
                        int styleIndexOfCondition = 0;
                        if (cu.getStyleId() != null) {
                            if (cFMapIndex.containsKey(cu.getStyleId())) {
                                styleIndexOfCondition = cFMapIndex.get(cu.getStyleId()).intValue();
                            }
                        }

                        cf +=
                                "<conditionalFormatting sqref=\"" +
                                        cu.getStart() +
                                        ":" +
                                        cu.getEnd() +
                                        "\"><cfRule type=\"" +
                                        (cu.getOperator() != null ? "aboveAverage" : "top10") +
                                        "\" dxfId=\"" +
                                        (styleIndexOfCondition) +
                                        "\" priority=\"" +
                                        (cu.getPriority() != null ? cu.getPriority() : priorityCounter++) +
                                        "\" " +
                                        (Utils.booleanCheck(cu.getBottom()) ? "bottom=\"1\"" : "") +
                                        " " +
                                        (cu.getPercent() != null ? "percent=\"1\"" : "") +
                                        "  rank=\"" +
                                        cu.getValue() +
                                        "\" " +
                                        (cu.getOperator() == ConditionalFormattingOption.ConditionalFormattingTopOperation.BELOW_AVERAGE ? "aboveAverage=\"0\"" : "") +
                                        "/></conditionalFormatting>";
                        continue;
                    } else if (cu.getType() == ConditionalFormattingOption.Type.ICON_SET) {
                        String percentValue = "";
                        if (cu.getOperator() == null || !(cu.getOperator() instanceof ConditionalFormattingOption.ConditionalFormattingIconSetOperation)) {
                            continue;
                        }
                        var opIconSet = (ConditionalFormattingOption.ConditionalFormattingIconSetOperation) cu.getOperator();
                        if (opIconSet.getOperationType().indexOf("5") == 0) {
                            percentValue =
                                    "<cfvo type=\"percent\" val=\"0\"/><cfvo type=\"percent\" val=\"20\"/><cfvo type=\"percent\" val=\"40\"/><cfvo type=\"percent\" val=\"60\"/><cfvo type=\"percent\" val=\"80\"/>";
                        } else if (opIconSet.getOperationType().indexOf("4") == 0) {
                            percentValue =
                                    "<cfvo type=\"percent\" val=\"0\"/><cfvo type=\"percent\" val=\"25\"/><cfvo type=\"percent\" val=\"50\"/><cfvo type=\"percent\" val=\"75\"/>";
                        } else {
                            percentValue =
                                    "<cfvo type=\"percent\" val=\"0\"/><cfvo type=\"percent\" val=\"33\"/><cfvo type=\"percent\" val=\"67\"/>";
                        }

                        cf +=
                                "<conditionalFormatting sqref=\"" +
                                        cu.getStart() +
                                        ":" +
                                        cu.getEnd() +
                                        "\"><cfRule type=\"iconSet\" priority=\"" +
                                        (cu.getPriority() != null ? cu.getPriority() : priorityCounter++) +
                                        "\"><iconSet iconSet=\"" +
                                        opIconSet.getOperationType() +
                                        "\">" +
                                        percentValue +
                                        "</iconSet></cfRule></conditionalFormatting>"
                        ;
                        continue;
                    } else if (cu.getType() == ConditionalFormattingOption.Type.COLOR_SCALE) {
                        cf +=
                                "<conditionalFormatting sqref=\"" +
                                        cu.getStart() +
                                        ":" +
                                        cu.getEnd() +
                                        "\"><cfRule type=\"colorScale\" priority=\"" +
                                        (cu.getPriority() != null ? cu.getPriority() : priorityCounter++) +
                                        "\"><colorScale><cfvo type=\"min\"/>" +
                                        (cu.getOperator() != null && cu.getOperator().equals("percentile")
                                                ? "<cfvo type=\"percentile\" val=\"" + cu.getValue() + "\"/>"
                                                : "") +
                                        "<cfvo type=\"max\"/>" +
                                        (cu.getColors() != null
                                                ? cu.getColors().stream().reduce("", (reColors, colorCu) -> reColors +
                                                "<color rgb=\"" +
                                                ColorUtils.convertToHex(colorCu) +
                                                "\"/>"
                                        )
                                                : "<color rgb=\"FFF8696B\"/><color rgb=\"FFFFEB84\"/><color rgb=\"FF63BE7B\"/>") +
                                        "</colorScale></cfRule></conditionalFormatting>";
                        continue;
                    } else if (cu.getType() == ConditionalFormattingOption.Type.DATABAR) {
                        cf +=
                                "<conditionalFormatting sqref=\"" +
                                        cu.getStart() +
                                        ":" +
                                        cu.getEnd() +
                                        "\"><cfRule type=\"dataBar\" priority=\"" +
                                        (cu.getPriority() != null ? cu.getPriority() : priorityCounter++) +
                                        "\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/>" +
                                        (cu.getColors() != null
                                                ? cu.getColors().stream().reduce("", (reColors, colorCu) -> reColors +
                                                "<color rgb=\"" +
                                                ColorUtils.convertToHex(colorCu) +
                                                "\"/>")
                                                : "<color rgb=\"FF638EC6\"/>") +
                                        "</dataBar></cfRule></conditionalFormatting>";
                    }
                }
                cFDataString = cf;
            }
            if ((hasCheckbox || hasComment || hasImages) && xl_drawingsFolder == null) {
                xl_drawingsFolder = xlFolder + "/drawings";
            }
            if (hasImages && xl_drawings_relsFolder == null) {
                xl_drawings_relsFolder = xl_drawingsFolder + "/_rels";
            }
            mapData.put("sheet" + (index + 1), SheetProcessResult.builder()
                    .indexId(indexId + 1)
                    .key("sheet" + (index + 1))
                    .sheetName(shName)
                    .sheetDataTableColumns(sheetDataTableColumns)
                    .backgroundImageRef(backgroundImageRef)
                    .sheetDimensions(sheetDimensions)
                    .asTable(asTable)

                    .sheetDropDown(DropDownUtil.generateDropDown(sheetData.getDropDowns()))
                    .sheetDataString(sheetDataString)
                    .sheetBreakLine(sheetBreakLine)
                    .viewType(viewType)
                    .hasComment(hasComment)
                    .drawersContent(drawersContent)
                    .cFDataString(cFDataString)
                    .sheetMargin(sheetMargin)
                    .sheetHeaderFooter(sheetHeaderFooter)
                    .isPortrait(isPortrait)
                    .drawersRels(drawersRels)
                    .hasImages(hasImages)
                    .hasCheckbox(hasCheckbox)
                    .formRel(formRel)
                    .checkboxDrawingContent(checkboxDrawingContent)
                    .checkboxForm(checkboxForm)
                    .checkboxSheetContent(checkboxSheetContent)
                    .checkboxShape(checkboxShape)
                    .commentString(commentString)
                    .commentAuthor(commentAuthor)
                    .shapeCommentRowCol(shapeCommentRowCol)
                    .splitOption(splitOption)
                    .sheetViewProperties(sheetViewProperties)
                    .sheetSizeString(
                            sheetSizeString.length() > 0
                                    ? "<cols>" + sheetSizeString + "</cols>"
                                    : "")
                    .protectionOption(sheetData.generateProtectedString())
                    .merges(
                            mergesCellArray.size() > 0
                                    ? mergesCellArray.stream().reduce("<mergeCells count=\"" + mergesCellArray.size() + "\">", (mResult, currRef) -> mResult + " <mergeCell ref=\"" + currRef + "\" />") +
                                    " </mergeCells>"
                                    : "")
                    .selectedView(sheetData.getSelected())
                    .sheetSortFilter(sheetSortFilter)
                    .tabColor(sheetData.getTabColor() != null ? "<sheetPr codeName=\"" +
                            ("Sheet" + (index + 1)) +
                            "\" " +
                            filterMode +
                            " >" +
                            "<tabColor rgb=\"" +
                            sheetData.getTabColor().replace("#", "") +
                            "\" />" +
                            "</sheetPr>"
                            : "<sheetPr " +
                            filterMode +
                            " >" +
                            "<outlinePr summaryBelow=\"0\" summaryRight=\"0\" />" +
                            "</sheetPr>")

                    .build());
            indexId++;
        }
        if (needCalcChain) {
            indexId++;
            workbookRelString +=
                    "<Relationship Id=\"rId" +
                            indexId +
                            "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" Target=\"calcChain.xml\"/>";
            resultStringMap.put(xlFolder + "/calcChain.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                    calcChainValue +
                    "</calcChain>");


        }

        var sheetKeys = mapData.keySet();
        // in _rels
        String relsFolder = "_rels";
        resultStringMap.put(relsFolder + "/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                " <Relationship Id=\"rId3\"" +
                "  Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\"" +
                "  Target=\"docProps/app.xml\" />" +
                " <Relationship Id=\"rId2\"" +
                "  Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\"" +
                "  Target=\"docProps/core.xml\" />" +
                " <Relationship Id=\"rId1\"" +
                "  Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\"" +
                "  Target=\"xl/workbook.xml\" />" +
                "</Relationships>");
        String docPropsFolder = "docProps";
        resultStringMap.put(docPropsFolder + "/core.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                        "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
                        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
                        "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
                        (data.getCreator() != null ? "<dc:creator>" + data.getCreator() + "</dc:creator>" : "") +
                        (data.getCreated() != null
                                ? "<dcterms:created xsi:type=\"dcterms:W3CDTF\">" +
                                data.getCreated() +
                                "</dcterms:created>"
                                : "") +
                        (data.getModified() != null
                                ? "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" +
                                data.getModified() +
                                "</dcterms:modified>"
                                : "") +
                        "</cp:coreProperties>");
        resultStringMap.put(docPropsFolder + "/app.xml", XmlUtils.appGenerator(sheetLength, sheetNameApp));
        resultStringMap.put(xlFolder + "/workbook.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"" +
                " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"" +
                " xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\"" +
                " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"" +
                " xmlns:mv=\"urn:schemas-microsoft-com:mac:vml\"" +
                " xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"" +
                " xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"" +
                " xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"" +
                " xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">" +
                " <workbookPr />" +
                (selectedAdded
                        ? "<bookViews><workbookView xWindow=\"3540\" yWindow=\"1365\" windowWidth=\"21600\" windowHeight=\"11325\" activeTab=\"" +
                        activeTabIndex +
                        "\"/></bookViews>"
                        : "") +
                " <sheets>" +
                "  " +
                workbookString +
                " </sheets>" +
                " <definedNames />" +
                " <calcPr />" +
                "</workbook>");
        resultStringMap.put(xlFolder + "/sharedStrings.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"" +
                (sharedStringIndex - 1) +
                "\" uniqueCount=\"" +
                sharedStringMap.size() +
                "\">" +
                " " +
                sharedString +
                "</sst>");


        String xl__relsFolder = xlFolder + "/_rels";
        resultStringMap.put(xl__relsFolder + "/workbook.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        " <Relationship Id=\"rId1\"" +
                        " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\"" +
                        " Target=\"theme/theme1.xml\" />" +
                        " <Relationship Id=\"rId2\"" +
                        " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\"" +
                        " Target=\"styles.xml\" />" +
                        " <Relationship Id=\"rId3\"" +
                        " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\"" +
                        " Target=\"sharedStrings.xml\" />" +
                        " " +
                        workbookRelString +
                        " " +
                        "</Relationships>");


        //xl/theme
        String xl_themeFolder = xlFolder + "/theme";
        resultStringMap.put(xl_themeFolder + "/theme1.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"  name=\"Office Theme\"><a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"44546A\"/></a:dk2><a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2><a:accent1><a:srgbClr val=\"5B9BD5\"/></a:accent1><a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2><a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3><a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4><a:accent5><a:srgbClr val=\"4472C4\"/></a:accent5><a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6><a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink><a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri Light\" panose=\"020F0302020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"游ゴシック Light\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"等线 Light\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/><a:font script=\"Hebr\" typeface=\"Times New Roman\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"MoolBoran\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Times New Roman\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\" panose=\"020F0502020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"游ゴシック\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"等线\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Arial\"/><a:font script=\"Hebr\" typeface=\"Arial\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"DaunPenh\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Arial\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>");


        // xl/worksheets
        String xl_worksheetsFolder = xlFolder + "/worksheets";
        List<Integer> commentId = new ArrayList<>();

        List<String> tableRef = new ArrayList<>();
        List<String> sheetDrawings = new ArrayList<>();
        int iCo = -1;
        for (var k : sheetKeys) {
            iCo++;
            var sh = mapData.get(k);
            String sheetRelContentStr = "";
            SheetRelFlag sheetRelSeenFlag = new SheetRelFlag(false);
            String sheetDataTableColumn = sh.getSheetDataTableColumns();

            if (sheetDataTableColumn.length() > 0) {
                tableRef.add("table" + (iCo + 1) + ".xml");
                var asTableOption = sh.getAsTable();
                var dimensions = sh.getSheetDimensions();
                resultStringMap.put(xl_tableFolder + "/" + "table" + (iCo + 1) + ".xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                        "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"xr xr3\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\" id=\"" +
                        (iCo + 1) +
                        "\"  name=\"Table" +
                        (iCo + 1) +
                        "\" displayName=\"Table" +
                        (iCo + 1) +
                        "\" ref=\"" +
                        dimensions.getStart() +
                        ":" +
                        dimensions.getEnd() +
                        "\" totalsRowShown=\"0\"><autoFilter ref=\"" +
                        dimensions.getStart() +
                        ":" +
                        dimensions.getEnd() +
                        "\"/>" +
                        sheetDataTableColumn +
                        "<tableStyleInfo name=\"TableStyle" +
                        (asTableOption.getType() != null ? asTableOption.getType().getType() : "Medium") +
                        (asTableOption.getStyleNumber() != null ? asTableOption.getStyleNumber() : 2) +
                        "\" showFirstColumn=\"" +
                        (Utils.booleanCheck(asTableOption.getFirstColumn()) ? asTableOption.getFirstColumn() : "0") +
                        "\" showLastColumn=\"" +
                        (Utils.booleanCheck(asTableOption.getLastColumn()) ? asTableOption.getLastColumn() : "0") +
                        "\" showRowStripes=\"" +
                        (Utils.booleanCheck(asTableOption.getLastColumn()) ? asTableOption.getRowStripes() : "1") +
                        "\" showColumnStripes=\"" +
                        (Utils.booleanCheck(asTableOption.getColumnStripes()) ? asTableOption.getColumnStripes() : "0") +
                        "\"/></table>");

                sheetRelContentStr +=
                        "<Relationship Id=\"rId15\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table" +
                                (iCo + 1) +
                                ".xml\"/>";
            }
            String drawerName = "drawing" + (sheetDrawings.size() + 1) + ".xml";
            if (Utils.booleanCheck(sh.getHasImages())) {
                sheetDrawings.add(drawerName);
                sheetRelSeenFlag.setSheetDrawingsPushed(true);
                resultStringMap.put(xl_drawings_relsFolder + "/" +
                                drawerName + ".rels",
                        sh.getDrawersRels()
                );

                sheetRelSeenFlag.setDrawing(true);
                sheetRelContentStr +=
                        "<Relationship Id=\"rId2\" " +
                                " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" " +
                                " Target=\"../drawings/" +
                                drawerName +
                                "\" />";
            }
            if (Utils.booleanCheck(sh.getHasCheckbox())) {
                if (!sheetRelSeenFlag.getSheetDrawingsPushed()) {
                    sheetDrawings.add(drawerName);
                }

                sheetRelContentStr +=
                        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawing" +
                                (iCo + 1) +
                                ".vml\" />" +
                                (sheetRelSeenFlag.getDrawing()
                                        ? ""
                                        : "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/" +
                                        drawerName +
                                        "\" />");
                sheetRelSeenFlag.setDrawing(true);
                sheetRelSeenFlag.setVmlDrwing(true);
                sheetRelContentStr += sh.getFormRel();
            }
            if (Utils.booleanCheck(sh.getHasCheckbox()) || Utils.booleanCheck(sh.getHasImages())) {
                resultStringMap.put(xl_drawingsFolder + "/" +
                                drawerName,
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                                "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" " +
                                " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                                " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                                " xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" " +
                                " xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" " +
                                " xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" " +
                                " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
                                " xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" " +
                                " xmlns:x3Unk=\"http://schemas.microsoft.com/office/drawing/2010/slicer\" " +
                                " xmlns:sle15=\"http://schemas.microsoft.com/office/drawing/2012/slicer\">" +
                                (sh.getHasImages() ? sh.getDrawersContent() : "") +
                                (sh.getHasCheckbox() ? sh.getCheckboxDrawingContent() : "") +
                                "</xdr:wsDr>"
                );
            }
            if (sh.getHasComment()) {
                commentId.add(iCo + 1);
                var aurt = sh.getCommentAuthor();
                resultStringMap.put(xlFolder + "/" + "comments" + (iCo + 1) + ".xml",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                                "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                                " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
                                " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                "<authors>" +
                                (aurt != null && aurt.size() > 0
                                        ? aurt.stream().reduce("",
                                        (res, currr) -> res + "<author>" + currr + "</author>"

                                )
                                        : "<author></author>") +
                                "</authors>" +
                                "<commentList>" +
                                sh.getCommentString() +
                                "</commentList>" +
                                "</comments>"
                );

                sheetRelContentStr +=
                        "<Relationship Id=\"rId1\" " +
                                "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" " +
                                "Target=\"../comments" +
                                (iCo + 1) +
                                ".xml\" />" +
                                (sheetRelSeenFlag.getVmlDrwing()
                                        ? ""
                                        : "<Relationship Id=\"rId3\" " +
                                        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" " +
                                        "Target=\"../drawings/vmlDrawing" +
                                        (iCo + 1) +
                                        ".vml\" />");
            }

            if (Utils.booleanCheck(sh.getHasCheckbox()) || Utils.booleanCheck(sh.getHasComment())) {
                resultStringMap.put(
                        xl_drawingsFolder + "/" + "vmlDrawing" + (iCo + 1) + ".vml",
                        "<xml xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:oa=\"urn:schemas-microsoft-com:office:activation\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:pvml=\"urn:schemas-microsoft-com:office:powerpoint\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/>" +
                                "</o:shapelayout>" +
                                (Utils.booleanCheck(sh.getHasCheckbox()) ? shapeTypeMap.get("checkbox") + sh.getCheckboxShape() : "") +
                                (Utils.booleanCheck(sh.getHasComment())
                                        ? "  <v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" " +
                                        "    path=\"m,l,21600r21600,l21600,xe\">" +
                                        "    <v:stroke joinstyle=\"miter\"/>" +
                                        "    <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>" +
                                        "  </v:shapetype>" +
                                        sh.getShapeCommentRowCol().stream().reduce("", (res, curr) -> res + "<v:shape id=\"_x0000_s1025\" type=\"#_x0000_t202\" style='position:absolute;" +
                                                "margin-left:77.25pt;margin-top:23.25pt;width:264pt;height:42.75pt;z-index:1;" +
                                                "visibility:hidden' fillcolor=\"#ffffe1\">" +
                                                "  <v:fill color2=\"#ffffe1\"/>" +
                                                "  <v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>" +
                                                "  <v:path o:connecttype=\"none\"/>" +
                                                "  <v:textbox>" +
                                                "   <div style='text-align:left'></div>" +
                                                "  </v:textbox>" +
                                                "  <x:ClientData ObjectType=\"Note\">" +
                                                "   <x:MoveWithCells/>" +
                                                "   <x:SizeWithCells/>" +
                                                "   <x:Anchor>" +
                                                "    1, 15, 1, 10, 5, 15, 4, 4</x:Anchor>" +
                                                "   <x:AutoFill>False</x:AutoFill>" +
                                                "   <x:Row>" +
                                                curr.getRow() +
                                                "</x:Row>" +
                                                "   <x:Column>" +
                                                curr.getCol() +
                                                "</x:Column>" +
                                                "  </x:ClientData>" +
                                                "</v:shape>", (v1, v2) -> v1 + v2)
                                        : "") +
                                "</xml>"
                );
            }
            if (sh.getBackgroundImageRef() > 0) {
                sheetRelContentStr +=
                        "<Relationship Id=\"rId16\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image" +
                                sh.getBackgroundImageRef() +
                                ".png\"/>";
            }
            if (Utils.booleanCheck(sh.getHasCheckbox()) || Utils.booleanCheck(sh.getHasComment())
                    || Utils.booleanCheck(sh.getHasImages()) ||
                    sheetDataTableColumn.length() > 0 ||
                    sh.getBackgroundImageRef() > 0
            ) {
                String xl_worksheets_relsFolder = xl_worksheetsFolder + "/_rels";
                resultStringMap.put(xl_worksheets_relsFolder + "/" +
                                "sheet" + (iCo + 1) + ".xml.rels",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"> " +
                                sheetRelContentStr +
                                "</Relationships>"
                );
            }
            String sheetViews;
            if (Utils.booleanCheck(sh.getSelectedView()) || sh.getSplitOption() != null) {
                sheetViews =
                        "<sheetViews><sheetView tabSelected=\"1\"" +
                                sh.getSheetViewProperties() +
                                ((sh.getViewType()).length() > 0
                                        ? " view=\"" + sh.getViewType() + "\""
                                        : "") +
                                " workbookViewId=\"0\">" +
                                sh.getSplitOption() +
                                (Utils.booleanCheck(sh.getSelectedView()) ? "<selection activeCell=\"A0\" sqref=\"A0\" />" : "") +
                                "</sheetView></sheetViews>";
            } else {
                sheetViews =
                        "<sheetViews><sheetView workbookViewId=\"0\"" +
                                sh.getSheetViewProperties() +
                                ((sh.getViewType()).length() > 0
                                        ? " view=\"" + sh.getViewType() + "\""
                                        : "") +
                                "/></sheetViews>";
            }
            resultStringMap.put(xl_worksheetsFolder + "/" +
                            sh.getKey() + ".xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"" +
                            " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"" +
                            " xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\"" +
                            " xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" " +
                            " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"" +
                            " xmlns:mv=\"urn:schemas-microsoft-com:mac:vml\"" +
                            " xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\"" +
                            " xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\"" +
                            " xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\"" +
                            " xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"" +
                            " xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"" +
                            " xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"" +
                            " xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">" +
                            sh.getTabColor() +
                            sheetViews +
                            "<sheetFormatPr customHeight=\"1\" defaultColWidth=\"12.63\" defaultRowHeight=\"15.75\" />" +
                            sh.getSheetSizeString() +
                            "<sheetData>" +
                            sh.getSheetDataString() +
                            "</sheetData>" +
                            sh.getSheetDropDown() +
                            sh.getProtectionOption() +
                            sh.getSheetSortFilter() +
                            sh.getMerges() +
                            sh.getCFDataString() +
                            (Utils.booleanCheck(sh.getHasImages()) || Utils.booleanCheck(sh.getHasCheckbox()) ? "<drawing r:id=\"rId2\" />" : "") +
                            (Utils.booleanCheck(sh.getHasComment()) || Utils.booleanCheck(sh.getHasCheckbox())
                                    ? "<legacyDrawing r:id=\"rId3\" />"
                                    : "") +
                            (Utils.booleanCheck(sh.getHasCheckbox())
                                    ? "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"x14\"><controls>" +
                                    sh.getCheckboxSheetContent() +
                                    "</controls></mc:Choice></mc:AlternateContent>"
                                    : "") +
                            sh.getSheetMargin() +
                            (Utils.booleanCheck(sh.getIsPortrait()) ||
                                    (sh.getSheetBreakLine()).length() > 0
                                    ? "<pageSetup orientation=\"portrait\" r:id=\"rId" + (iCo + 1) + "\"/>"
                                    : "") +
                            sh.getSheetBreakLine() +
                            sh.getSheetHeaderFooter() +
                            (sh.getBackgroundImageRef() > 0 ? "<picture r:id=\"rId16\"/>" : "") +
                            (sheetDataTableColumn.length() > 0
                                    ? "<tableParts count=\"1\"> <tablePart r:id=\"rId15\"/></tableParts>"
                                    : "") +
                            "</worksheet>"
            );
        }
        if (checkboxForm.size() > 0) {
            String xlCtrlFolder = xlFolder + "/ctrlProps";
            int cIndex = 0;
            for (var v : checkboxForm) {
                resultStringMap.put(xlCtrlFolder + "/ctrlProp" + (cIndex + 1) + ".xml", v);
                cIndex++;
            }
        }
        resultStringMap.put(
                "[Content_Types].xml",
                XmlUtils.contentTypeGenerator(
                        sheetContentType,
                        commentId,
                        new ArrayList<>(new HashSet<>(arrTypes)),
                        sheetDrawings,
                        checkboxForm,
                        needCalcChain,
                        tableRef
                )
        );

        if (Utils.booleanCheck(data.getNotSave())) {
            var gType = data.getGenerateType();
            if (gType == null || gType == GenerateType.BYTE_ARRAY || gType == GenerateType.NODE_BUFFER || gType == GenerateType.ARRAY) {
                return new Result<>(FileUtils.byteArrayResponse(resultStringMap, resultAssetsMap), null);
            } else if (gType == GenerateType.INPUT_STREAM) {
                return new Result<>(FileUtils.inputStreamResponse(resultStringMap, resultAssetsMap), null);
            } else if (gType == GenerateType.BASE64) {
                return new Result<>(FileUtils.base64Response(resultStringMap, resultAssetsMap), null);
            } else if (gType == GenerateType.BYTE_BUFFER) {
                return new Result<>(FileUtils.bufferResponse(resultStringMap, resultAssetsMap), null);
            } else if (gType == GenerateType.BINARY_STRING) {
                return new Result<>(FileUtils.binaryStringResponse(resultStringMap, resultAssetsMap), null);
            }
            return null;
        } else {
            var resultFileName = FileUtils.writeFile((data.getFileName() != null) ? data.getFileName() : "tableRecord", resultStringMap, resultAssetsMap);
            return Result.builder()
                    .fileName(resultFileName)
                    .build();
        }
    }
}
