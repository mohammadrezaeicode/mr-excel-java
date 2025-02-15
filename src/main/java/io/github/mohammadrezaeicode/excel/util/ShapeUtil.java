package io.github.mohammadrezaeicode.excel.util;

import java.util.HashMap;
import java.util.Map;

public class ShapeUtil {
    public static Map<String,String> formCtrlMap=new HashMap<>();
    public static Map<String,String> shapeMap=new HashMap<>();
    public static Map<String,String> shapeTypeMap=new HashMap<>();
    static {
//        formCtrlMap.put("checkbox", """
//<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
//<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="CheckBox" **value** **fmlaLink** lockText="1" noThreeD="1"/>
//""");
//        shapeMap.put("checkbox", """
//<v:shape id="***id***" type="#_x0000_t201" style='position:absolute;
//  margin-left:1.5pt;margin-top:1.5pt;width:63pt;height:16.5pt;z-index:1;
//  mso-wrap-style:tight' filled="f" fillcolor="window [65]" stroked="f"
//  strokecolor="windowText [64]" o:insetmode="auto">
//  <v:path shadowok="t" strokeok="t" fillok="t"/>
//  <o:lock v:ext="edit" rotation="t"/>
//  <v:textbox style='mso-direction-alt:auto' o:singleclick="f">
//   <div style='text-align:left'><font face="Segoe UI" size="160" color="auto">***text***</font></div>
//  </v:textbox>
//  <x:ClientData ObjectType="Checkbox">
//   <x:SizeWithCells/>
//   <x:Anchor>
//    0, 2, 0, 2, 0, 86, 1, 0</x:Anchor>
//   <x:AutoFill>False</x:AutoFill>
//   <x:AutoLine>False</x:AutoLine>
//   <x:TextVAlign>Center</x:TextVAlign>
//   <x:NoThreeD/>
//  </x:ClientData>
// </v:shape>
//                """);
//        shapeTypeMap.put("checkbox", """
//<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201"
//  path="m,l,21600r21600,l21600,xe">
//  <v:stroke joinstyle="miter"/>
//  <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>
//  <o:lock v:ext="edit" shapetype="t"/>
// </v:shapetype>
//                """);
    }

    public static String getFormCtrlMap(String key) {
        return formCtrlMap.get(key);
    }

    public static String getShapeMap(String key) {
        return shapeMap.get(key);
    }

    public static String getShapeTypeMap(String key) {
        return shapeTypeMap.get(key);
    }
}
