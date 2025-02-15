package io.github.mohammadrezaeicode.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import io.github.mohammadrezaeicode.excel.model.row.SheetDimension;

import java.util.List;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class SheetProcessResult {
    private Integer indexId;
    private String key;
    private String sheetName;
    private String  sheetDataTableColumns;
   private Integer backgroundImageRef;
    private SheetDimension sheetDimensions;
    private Sheet1.AsTableOption asTable;
    private String sheetDataString;
    private String sheetDropDown;
    private String sheetBreakLine;
    private String viewType;
    private Boolean hasComment;
    private String drawersContent;
    private String cFDataString;
    private String sheetMargin;
    private String sheetHeaderFooter;
    private Boolean isPortrait;
    private String  drawersRels;
    private Boolean  hasImages;
    private Boolean hasCheckbox;
    private String formRel;
    private String checkboxDrawingContent;
    private List<String> checkboxForm;
    private String checkboxSheetContent;
    private String checkboxShape;
    private String commentString;
    private List<String>  commentAuthor;
    private List<ShapeRC> shapeCommentRowCol;
   private String  splitOption;
    private String sheetViewProperties;
    private String sheetSizeString;
    private String  protectionOption;
    private String  merges;
    private Boolean  selectedView;
    private String  sheetSortFilter;
    private String  tabColor;
}
