package io.github.mohammadrezaeicode.excel.util;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.ByteBuffer;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class FileUtils {
    public static void base64(){
        String data = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAApgAAAKYB3X3/OAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAANCSURBVEiJtZZPbBtFFMZ/M7ubXdtdb1xSFyeilBapySVU8h8OoFaooFSqiihIVIpQBKci6KEg9Q6H9kovIHoCIVQJJCKE1ENFjnAgcaSGC6rEnxBwA04Tx43t2FnvDAfjkNibxgHxnWb2e/u992bee7tCa00YFsffekFY+nUzFtjW0LrvjRXrCDIAaPLlW0nHL0SsZtVoaF98mLrx3pdhOqLtYPHChahZcYYO7KvPFxvRl5XPp1sN3adWiD1ZAqD6XYK1b/dvE5IWryTt2udLFedwc1+9kLp+vbbpoDh+6TklxBeAi9TL0taeWpdmZzQDry0AcO+jQ12RyohqqoYoo8RDwJrU+qXkjWtfi8Xxt58BdQuwQs9qC/afLwCw8tnQbqYAPsgxE1S6F3EAIXux2oQFKm0ihMsOF71dHYx+f3NND68ghCu1YIoePPQN1pGRABkJ6Bus96CutRZMydTl+TvuiRW1m3n0eDl0vRPcEysqdXn+jsQPsrHMquGeXEaY4Yk4wxWcY5V/9scqOMOVUFthatyTy8QyqwZ+kDURKoMWxNKr2EeqVKcTNOajqKoBgOE28U4tdQl5p5bwCw7BWquaZSzAPlwjlithJtp3pTImSqQRrb2Z8PHGigD4RZuNX6JYj6wj7O4TFLbCO/Mn/m8R+h6rYSUb3ekokRY6f/YukArN979jcW+V/S8g0eT/N3VN3kTqWbQ428m9/8k0P/1aIhF36PccEl6EhOcAUCrXKZXXWS3XKd2vc/TRBG9O5ELC17MmWubD2nKhUKZa26Ba2+D3P+4/MNCFwg59oWVeYhkzgN/JDR8deKBoD7Y+ljEjGZ0sosXVTvbc6RHirr2reNy1OXd6pJsQ+gqjk8VWFYmHrwBzW/n+uMPFiRwHB2I7ih8ciHFxIkd/3Omk5tCDV1t+2nNu5sxxpDFNx+huNhVT3/zMDz8usXC3ddaHBj1GHj/As08fwTS7Kt1HBTmyN29vdwAw+/wbwLVOJ3uAD1wi/dUH7Qei66PfyuRj4Ik9is+hglfbkbfR3cnZm7chlUWLdwmprtCohX4HUtlOcQjLYCu+fzGJH2QRKvP3UNz8bWk1qMxjGTOMThZ3kvgLI5AzFfo379UAAAAASUVORK5CYII=";
        System.out.println(data.substring(data.indexOf(",")+1));
//        String base64Image = data.split(",")[1];
//        byte[] imageBytes = javax.xml.parsers.DatatypeConverter.parseBase64Binary(base64Image);
        // encode with padding
//        String encoded = Base64.getEncoder().encodeToString(data.getBytes());

// encode without padding
//        String encoded = Base64.getEncoder().withoutPadding().encodeToString();

// decode a String
//        byte [] barr = Base64.getDecoder().decode(encoded);
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream("xx12.zip");
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        ZipOutputStream zipOut = new ZipOutputStream(fos);
        try {
            ZipEntry zipEntry = new ZipEntry("x3.png");
            zipOut.putNextEntry(zipEntry);
//            ByteArrayOutputStream baos = new ByteArrayOutputStream();
//            ImageIO.write(image, "png", baos);
//            byte[] bytes = baos.toByteArray();
            zipOut.write(Base64.getDecoder().decode(data.split(",")[1]));
            zipOut.close();
            fos.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    public static void readFile() throws IOException {
        FileOutputStream fos = new FileOutputStream("xx2.zip");
        ZipOutputStream zipOut = new ZipOutputStream(fos);
        try {
            ZipEntry zipEntry = new ZipEntry("x2.png");
            zipOut.putNextEntry(zipEntry);
            BufferedImage image= ImageIO.read(new File("t.png"));
            ImageIO.write(image,"png",zipOut);
//            ByteArrayOutputStream baos = new ByteArrayOutputStream();
//            ImageIO.write(image, "png", baos);
//            byte[] bytes = baos.toByteArray();
//            zipOut.write(bytes);
            zipOut.close();
            fos.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
//        fis.close();
    }
    public static void readFile2() throws IOException {
        FileOutputStream fos = new FileOutputStream("xx.zip");
        ZipOutputStream zipOut = new ZipOutputStream(fos);
        try {
            ZipEntry zipEntry = new ZipEntry("x2.png");
            zipOut.putNextEntry(zipEntry);
            BufferedImage image= ImageIO.read(new File("t.png"));
//         ImageIO.write(image,"png",zipOut);
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ImageIO.write(image, "png", baos);
            byte[] bytes = baos.toByteArray();
            zipOut.write(bytes);
            zipOut.close();
            fos.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
//        fis.close();
    }
    public static Map<String,String> createExcel(){
        Map<String,String> map=new HashMap<>();
        map.put("_rels/.rels","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                "        <Relationship Id=\"rId3\"\n" +
                "        Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\"\n" +
                "        Target=\"docProps/app.xml\" />\n" +
                "    <Relationship Id=\"rId2\"\n" +
                "        Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\"\n" +
                "        Target=\"docProps/core.xml\" />\n" +
                "    <Relationship Id=\"rId1\"\n" +
                "        Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\"\n" +
                "        Target=\"xl/workbook.xml\" />\n" +
                "</Relationships>");
        map.put("docProps/app.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"\n" +
                "    xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\n" +
                "    <Application>Microsoft Excel</Application>\n" +
                "    <DocSecurity>0</DocSecurity>\n" +
                "    <ScaleCrop>false</ScaleCrop>\n" +
                "    <HeadingPairs>\n" +
                "        <vt:vector size=\"2\" baseType=\"variant\">\n" +
                "            <vt:variant>\n" +
                "                <vt:lpstr>Worksheets</vt:lpstr>\n" +
                "            </vt:variant>\n" +
                "            <vt:variant>\n" +
                "                <vt:i4>1</vt:i4>\n" +
                "            </vt:variant>\n" +
                "        </vt:vector>\n" +
                "    </HeadingPairs>\n" +
                "    <TitlesOfParts>\n" +
                "        <vt:vector size=\"1\" baseType=\"lpstr\">\n" +
                "            \n" +
                "        <vt:lpstr>sheet1</vt:lpstr>\n" +
                "        \n" +
                "        </vt:vector>\n" +
                "    </TitlesOfParts>\n" +
                "    <Company></Company>\n" +
                "    <LinksUpToDate>false</LinksUpToDate>\n" +
                "    <SharedDoc>false</SharedDoc>\n" +
                "    <HyperlinksChanged>false</HyperlinksChanged>\n" +
                "    <AppVersion>16.0300</AppVersion>\n" +
                "</Properties>");
        map.put("docProps/core.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"></cp:coreProperties>");
        map.put("xl/_rels/workbook.xml.rels","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                "    <Relationship Id=\"rId1\"\n" +
                "        Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\"\n" +
                "        Target=\"theme/theme1.xml\" />\n" +
                "    <Relationship Id=\"rId2\"\n" +
                "        Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\"\n" +
                "        Target=\"styles.xml\" />\n" +
                "    <Relationship Id=\"rId3\"\n" +
                "        Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\"\n" +
                "        Target=\"sharedStrings.xml\" />\n" +
                "    <Relationship Id=\"rId5\"\n" +
                "        Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\"\n" +
                "        Target=\"worksheets/sheet1.xml\" />\n" +
                "</Relationships>");
        map.put("xl/theme/theme1.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \n" +
                "    xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"  name=\"Office Theme\"><a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"44546A\"/></a:dk2><a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2><a:accent1><a:srgbClr val=\"5B9BD5\"/></a:accent1><a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2><a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3><a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4><a:accent5><a:srgbClr val=\"4472C4\"/></a:accent5><a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6><a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink><a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri Light\" panose=\"020F0302020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"游ゴシック Light\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"等线 Light\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/><a:font script=\"Hebr\" typeface=\"Times New Roman\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"MoolBoran\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Times New Roman\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\" panose=\"020F0502020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"游ゴシック\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"等线\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Arial\"/><a:font script=\"Hebr\" typeface=\"Arial\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"DaunPenh\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Arial\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri=\"{05A4C25C-085E-4340-85A3-A5531E510DB2}\"><thm15:themeFamily xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\" name=\"Office Theme\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\" vid=\"{4A3C46E8-61CC-4603-A589-7422A47A8E4A}\"/></a:ext></a:extLst></a:theme>");
        map.put("xl/worksheets/sheet1.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"\n" +
                "    xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"\n" +
                "    xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"\n" +
                "    xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\"\n" +
                "    xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"\n" +
                "    xmlns:mv=\"urn:schemas-microsoft-com:mac:vml\"\n" +
                "    xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"\n" +
                "    xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"\n" +
                "    xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"\n" +
                "    xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">\n" +
                "    <sheetPr  >\n" +
                "        <outlinePr summaryBelow=\"0\" summaryRight=\"0\" />\n" +
                "    </sheetPr>\n" +
                "    <sheetViews>\n" +
                "        <sheetView workbookViewId=\"0\" />\n" +
                "    </sheetViews>\n" +
                "    <sheetFormatPr customHeight=\"1\" defaultColWidth=\"12.63\" defaultRowHeight=\"15.75\" />\n" +
                "    \n" +
                "    <sheetData>\n" +
                "       <row r=\"1\" spans=\"1:1\"   >\n" +
                "\n" +
                "<c r=\"A1\"  t=\"s\">\n" +
                "    <v>0</v>\n" +
                "</c>\n" +
                "                        \n" +
                "</row>\n" +
                "                        <row r=\"2\" spans=\"1:1\"    >\n" +
                "<c r=\"A2\" t=\"s\" >\n" +
                "    <v>1</v>\n" +
                "</c>\n" +
                "                        </row>\n" +
                "    </sheetData> \n" +
                "    \n" +
                "    \n" +
                "    \n" +
                "    \n" +
                "    \n" +
                "    \n" +
                "    \n" +
                "</worksheet>");
        map.put("xl/sharedStrings.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">\n" +
                "    \n" +
                "<si>\n" +
                "    <t></t>\n" +
                "</si>\n" +
                "    \n" +
                "<si>\n" +
                "    <t></t>\n" +
                "</si>\n" +
                "    \n" +
                "</sst>");
        map.put("xl/drawings/vmlDrawing1.vml","<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\n" +
                "             xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n" +
                "             xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\n" +
                "             <o:shapelayout v:ext=\"edit\">\n" +
                "              <o:idmap v:ext=\"edit\" data=\"1\"/>\n" +
                "             </o:shapelayout><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"\n" +
                "              path=\"m,l,21600r21600,l21600,xe\">\n" +
                "              <v:stroke joinstyle=\"miter\"/>\n" +
                "              <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\n" +
                "             </v:shapetype><v:shape id=\"_x0000_s1025\" type=\"#_x0000_t202\" style='position:absolute;\n" +
                "              margin-left:77.25pt;margin-top:23.25pt;width:264pt;height:42.75pt;z-index:1;\n" +
                "              visibility:hidden' fillcolor=\"#ffffe1\">\n" +
                "              <v:fill color2=\"#ffffe1\"/>\n" +
                "              <v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>\n" +
                "              <v:path o:connecttype=\"none\"/>\n" +
                "              <v:textbox>\n" +
                "               <div style='text-align:left'></div>\n" +
                "              </v:textbox>\n" +
                "              <x:ClientData ObjectType=\"Note\">\n" +
                "               <x:MoveWithCells/>\n" +
                "               <x:SizeWithCells/>\n" +
                "               <x:Anchor>\n" +
                "                1, 15, 1, 10, 5, 15, 4, 4</x:Anchor>\n" +
                "               <x:AutoFill>False</x:AutoFill>\n" +
                "               <x:Row>0</x:Row>\n" +
                "               <x:Column>0</x:Column>\n" +
                "              </x:ClientData>\n" +
                "             </v:shape></xml>");
        map.put("xl/styles.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"\n" +
                "    xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"\n" +
                "    xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">\n" +
                "    \n" +
                "    <fonts count=\"2\">\n" +
                "         <font>\n" +
                "            <sz val=\"11\" />\n" +
                "            <color theme=\"1\" />\n" +
                "            <name val=\"Calibri\" />\n" +
                "            <family val=\"2\" />\n" +
                "            <scheme val=\"minor\" />\n" +
                "        </font>\n" +
                "        <font>\n" +
                "            <sz val=\"11\" />\n" +
                "            <color rgb=\"FFFF0000\" />\n" +
                "            <name val=\"Calibri\" />\n" +
                "            <family val=\"2\" />\n" +
                "            <scheme val=\"minor\" />\n" +
                "        </font>\n" +
                "        \n" +
                "    </fonts>\n" +
                "    <fills count=\"2\">\n" +
                "        <fill>\n" +
                "            <patternFill patternType=\"none\" />\n" +
                "        </fill>\n" +
                "        <fill>\n" +
                "            <patternFill patternType=\"lightGray\" />\n" +
                "        </fill>\n" +
                "        \n" +
                "    </fills>\n" +
                "    <borders count=\"1\">\n" +
                "        <border />\n" +
                "        \n" +
                "    </borders>\n" +
                "    <cellStyleXfs count=\"1\">\n" +
                "        <xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" applyAlignment=\"1\" applyFont=\"1\" />\n" +
                "    </cellStyleXfs>\n" +
                "    <cellXfs count=\"2\">\n" +
                "        <xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\"\n" +
                "            applyFont=\"1\">\n" +
                "            <alignment readingOrder=\"0\" shrinkToFit=\"0\" vertical=\"bottom\" wrapText=\"0\" />\n" +
                "        </xf>\n" +
                "        <xf borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\" xfId=\"0\" applyAlignment=\"1\"\n" +
                "            applyFont=\"1\">\n" +
                "            <alignment readingOrder=\"0\" />\n" +
                "        </xf>\n" +
                "       \n" +
                "    </cellXfs>\n" +
                "    <cellStyles count=\"1\">\n" +
                "        <cellStyle xfId=\"0\" name=\"Normal\" builtinId=\"0\" />\n" +
                "    </cellStyles>\n" +
                "    <dxfs count=\"0\" />\n" +
                "</styleSheet>");
        map.put("xl/workbook.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"\n" +
                "    xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"\n" +
                "    xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\"\n" +
                "    xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"\n" +
                "    xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"\n" +
                "    xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"\n" +
                "    xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"\n" +
                "    xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">\n" +
                "    <workbookPr />\n" +
                "    <sheets>\n" +
                "        <sheet state=\"visible\" name=\"sheet1\" sheetId=\"1\" r:id=\"rId5\" />\n" +
                "    </sheets>\n" +
                "    <definedNames />\n" +
                "    <calcPr />\n" +
                "</workbook>");
        map.put("[Content_Types].xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n" +
                "    \n" +
                "    <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />\n" +
                "    <Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\" />\n" +
                "    <Default Extension=\"xml\" ContentType=\"application/xml\" />\n" +
                "    <Override\n" +
                "        ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"\n" +
                "        PartName=\"/xl/workbook.xml\" />\n" +
                "    <Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"\n" +
                "        PartName=\"/xl/styles.xml\" />\n" +
                "    <Override ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"\n" +
                "        PartName=\"/xl/theme/theme1.xml\" />\n" +
                "    \n" +
                "    <Override\n" +
                "        ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"\n" +
                "        PartName=\"/xl/worksheets/sheet1.xml\" />\n" +
                "        \n" +
                "    <Override\n" +
                "        ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"\n" +
                "        PartName=\"/xl/sharedStrings.xml\" />\n" +
                "             \n" +
                "            \n" +
                "    <Override PartName=\"/docProps/core.xml\"\n" +
                "        ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\" />\n" +
                "    <Override PartName=\"/docProps/app.xml\"\n" +
                "        ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\" />\n" +
                "</Types>");
        return map;
    }
    public static void writeFile(String fileName) throws IOException {
        String test="Test";
        String test2="Test2";
        FileOutputStream fos = new FileOutputStream("tt.xlsx");
        ZipOutputStream zipOut = new ZipOutputStream(fos);
//        InputStream fis =new ByteArrayInputStream(test.getBytes());
        Map<String,String> items=createExcel();
        Set<String> keys=items.keySet();
        Iterator<String> iterator=keys.iterator();
        while (iterator.hasNext()){
            String key=iterator.next();
            ZipEntry zipEntry = new ZipEntry(key);
            zipOut.putNextEntry(zipEntry);
            zipOut.write(items.get(key).getBytes());
        }
//        ZipEntry zipEntry = new ZipEntry("yet/test.txt");
//        ZipEntry zipEntry2 = new ZipEntry("yet/test2.txt");
//        zipOut.putNextEntry(zipEntry);
//        zipOut.write(test.getBytes());
//        zipOut.putNextEntry(zipEntry2);
//        zipOut.write(test2.getBytes());
        zipOut.close();
//        fis.close();
        fos.close();
    }
    public static String writeFile(String fileName,Map<String,String> items) throws IOException {
        String generatedFileName=fileName+".xlsx";
        FileOutputStream fos = new FileOutputStream(generatedFileName);
        ZipOutputStream zipOut = new ZipOutputStream(fos);
        Set<String> keys=items.keySet();
        for (String key:keys){
            ZipEntry zipEntry = new ZipEntry(key);
            zipOut.putNextEntry(zipEntry);
            zipOut.write(items.get(key).getBytes());
        }
        zipOut.close();
        fos.close();
        return generatedFileName;
    }
    public static ByteBuffer bufferResponse(Map<String,String> items)throws IOException{
        return ByteBuffer.wrap(byteArrayResponse(items));}
    public static String base64Response(Map<String,String> items)throws IOException{
        return Base64.getEncoder().encodeToString(byteArrayResponse(items));
    }
    public static String binaryStringResponse(Map<String,String> items) throws IOException {
        StringBuilder binaryString = new StringBuilder();
        var zipData=byteArrayResponse(items);
        for (byte b : zipData) {
            int value = b & 0xFF;
            String binaryVal = Integer.toBinaryString(value);
            while (binaryVal.length() < 8) {
                binaryVal = "0" + binaryVal;
            }
            binaryString.append(binaryVal);
        }
        return binaryString.toString();
    }
    public static InputStream inputStreamResponse(Map<String,String> items)throws IOException{
        return new ByteArrayInputStream(byteArrayResponse(items));
    }
    public static byte[] byteArrayResponse(Map<String,String> items) throws IOException {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        ZipOutputStream zipOut = new ZipOutputStream(byteArrayOutputStream);
        Set<String> keys=items.keySet();
        for (String key:keys){
            ZipEntry zipEntry = new ZipEntry(key);
            zipOut.putNextEntry(zipEntry);
            zipOut.write(items.get(key).getBytes());
        }
        zipOut.close();
        return byteArrayOutputStream.toByteArray();
    }
}
