# MR Excel

MR Excel is a tool for generating Excel files. The implementation of this library is similar to ["`MR Excel JavaScript`"](https://github.com/mohammadrezaeicode/mr-excel-repo). You can use the documentation from that repository for reference. [example](https://github.com/mohammadrezaeicode/mr-excel-java-example-gallery) Requirement:**`JDK 11+`**

## Related Projects

The following list includes new repositories related to this project. Documentation and improvements for these projects can be found in the repositories below.

- **`MR Excel Java`**:A similar project using Java is in development. The release version is coming soon; currently, it is available as a snapshot version.["`repository`"](https://github.com/mohammadrezaeicode/mr-excel-java)

- **`MR Excel Editor`**: An editor that utilizes the library is currently under development. At present, it only generates simple results.["`repository`"](https:///github.com/mohammadrezaeicode/mr-excel-editor)["`Demo`"](https://mohammadrezaeicode.github.io/mr-excel-editor/)

## Central repo

[repo](https://central.sonatype.com/artifact/io.github.mohammadrezaeicode/excel/overview)

### Import in `Maven` project

To use the dependency, follow the structure below:

```xml
<dependency>
    <groupId>io.github.mohammadrezaeicode</groupId>
    <artifactId>excel</artifactId>
    <version>1.0.1</version>
</dependency>
```

### Import in `Gradle` project

```gradle
implementation 'io.github.mohammadrezaeicode:excel:1.0.1'
```

## Github

### Import in `Maven` project

To use the dependency, follow the structure below:

1. `Generate a GitHub token` with access to `read:packages` and add the server to `~/.m2/setting.xml`. You can find the exact structure in this [link](https://docs.github.com/en/packages/working-with-a-github-packages-registry/working-with-the-apache-maven-registry#authenticating-with-a-personal-access-token).

```xml
<servers>
    <server>
        <id></id>
        <username></username>
        <password></password>
    </server>
</servers>
```

2. Add the repository to your `pom.xml` file:

```xml
<repositories>
  <repository>
    <id>github</id>
    <url>https://maven.pkg.github.com/mohammadrezaeicode/excel</url>
    <releases>
      <enabled>true</enabled>
    </releases>
    <snapshots>
      <enabled>true</enabled>
    </snapshots>
  </repository>
</repositories>
```

3. Add the library to your `dependencies`:

```xml
<dependency>
  <groupId>io.github.mohammadrezaeicode</groupId>
  <artifactId>excel</artifactId>
  <version>1.0.1</version>
</dependency>
```

### Import in `Gradle` project

To use the dependency, follow the structure below:

1. **Generate a GitHub token** with access to `read:packages` and configure your `~/.gradle/gradle.properties` file:

```properties
gpr.user=YOUR_GITHUB_USERNAME
gpr.token=YOUR_GITHUB_TOKEN
```

2. **Add the repository to your `build.gradle` file**:

```groovy
repositories {
    maven {
        url = uri("https://maven.pkg.github.com/mohammadrezaeicode/excel")
        credentials {
            username = project.findProperty("gpr.user") ?: System.getenv("USERNAME")
            password = project.findProperty("gpr.token") ?: System.getenv("TOKEN")
        }
    }
}
```

3. **Add the library to your dependencies**:

```groovy
dependencies {
    implementation 'io.github.mohammadrezaeicode:excel:0.1-SNAPSHOT'
}
```

## Example

For the most recent updates and examples, please refer to the following [link](https://github.com/mohammadrezaeicode/mr-excel-example-gallery/).

empty excel

```java
GenerateExcel.generateExcel(null,null);
```

empty excel multi sheet

```java
GenerateExcel.generateExcel(ExcelTable.builder()
        .sheet(
                Arrays.asList(
                        Sheet.builder().build(),
                        Sheet.builder().build(),
                        Sheet.builder().build(),
                        Sheet.builder().build()

                )

        )
.build(), null);
```

style, formula

```java
GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Arrays.asList(
        SheetGenerator.builder().data(
                        Arrays.asList(
                                Test.builder()._id(0.3).
                                        test("test1").build(),
                                Test.builder()._id(2).
                                        test("test2").build(),
                                Test.builder()._id(3).
                                        test("test3").build(),
                                Test.builder()._id(4).
                                        test("test4").build(),
                                Test.builder()._id(5).
                                        test("test5").build(),
                                Test.builder()._id(6).
                                        test("test6").build(),
                                Test.builder()._id(7).
                                        test("test7").build(),
                                Test.builder()._id(8).
                                        test("test8").build(),
                                Test.builder()._id(9).
                                        test("test9").build(),
                                Test.builder()._id(10).
                                        test("test10").build(),
                                Test.builder()._id(11).
                                        test("test11").build())
                ).applyHeaderOptionFunction(
                        (headers -> {
                            for (var header : headers) {
                                if (header.getText().equals("ID")) {
                                    header.setFormula(FormulaSetting.builder()
                                            .type(FormulaSetting.FormulaType.MAX)
                                            .styleId("formulaStyle")
                                            .build());
                                }
                            }
                            return headers;
                        })
                ).applySheetOptionFunction(
                        sheetBuilder -> {
                            Map<String, FormulaMapBody> formula = new HashMap<>();
                            formula.put("B16", FormulaSetting.builder()
                                    .type(FormulaSetting.FormulaType.SUM)
                                    .start("B2")
                                    .end("B8")
                                    .styleId("formulaStyle")
                                    .build());
                            sheetBuilder.tabColor("B83B5E").name("Test").formula(formula);
                            return sheetBuilder.build();
                        }
                )
                .build(),
        SheetGenerator.builder().data(
                Arrays.asList(
                        Test.builder()._id(1).
                                test("test1").build(),
                        Test.builder()._id(2).
                                test("test2").build(),
                        Test.builder()._id(3).
                                test("test3").build(),
                        Test.builder()._id(4).
                                test("test4").build(),
                        Test.builder()._id(5).
                                test("test5").build(),
                        Test.builder()._id(6).
                                test("test6").build(),
                        Test.builder()._id(7).
                                test("test7").build(),
                        Test.builder()._id(8).
                                test("test8").build(),
                        Test.builder()._id(9).
                                test("test9").build(),
                        Test.builder()._id(10).
                                test("test10").build(),
                        Test.builder()._id(11).
                                test("test11").build()
                )
        ).build()
), ExcelTableOption
        .eBuilder()
        .created("2023-08-06T07:22:40Z")
        .modified("2023-08-06T07:22:40Z")
        .creator("MR Excel")
        .addStyles("formulaStyle", StyleBody.builder()
                .backgroundColor("B83B5E")
                .border(StyleBody.BorderOption.builder()
                        .full(StyleBody.BorderOption.BorderBody.builder()
                                .style(StyleBody.BorderOption.BorderBody.BorderStyle.MEDIUM)
                                .color("F9ED69")
                                .build())
                        .build())
                .build())
        .fileName("ex1")
        .build());
```

not save, protection option

```java
Result<InputStream> data = GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Arrays.asList(
        SheetGenerator.builder().applySheetOptionFunction(
                        sheetBuilder -> {
                            sheetBuilder.name("family record").tabColor("#a1b4c6").sortAndFilter(SortAndFilter.builder().mode(SortAndFilter.Mode.ALL).build());
                            Map<Sheet.ProtectionOptionKey, Boolean> map = new HashMap<>();
                            map.put(Sheet.ProtectionOptionKey.SHEET, true);
                            map.put(Sheet.ProtectionOptionKey.FORMAT_CELLS, false);
                            map.put(Sheet.ProtectionOptionKey.FORMAT_COLUMNS, false);
                            map.put(Sheet.ProtectionOptionKey.FORMAT_ROWS, false);
                            map.put(Sheet.ProtectionOptionKey.INSERT_COLUMNS, false);
                            map.put(Sheet.ProtectionOptionKey.INSERT_ROWS, false);
                            map.put(Sheet.ProtectionOptionKey.INSERT_HYPER_LINKS, false);
                            map.put(Sheet.ProtectionOptionKey.DELETE_COLUMNS, false);
                            map.put(Sheet.ProtectionOptionKey.DELETE_ROWS, false);
                            map.put(Sheet.ProtectionOptionKey.SORT, false);
                            map.put(Sheet.ProtectionOptionKey.AUTO_FILTER, false);
                            map.put(Sheet.ProtectionOptionKey.PIVOT_TABLES, false);
                            sheetBuilder.protectionOption(map);
                            return sheetBuilder.build();
                        }
                ).data(
                        ExampleData.ex2Sheet1()
                )
                .build(),
        SheetGenerator.builder().applySheetOptionFunction(
                        sheetBuilder -> sheetBuilder.state(Sheet.State.HIDDEN).build())
                .data(
                        ExampleData.ex2Sheet2()
                ).build()
), ExcelTableOption
        .eBuilder()
        .created("2023-08-06T07:22:40Z")
        .modified("2023-08-06T07:22:40Z")
        .notSave(true)
        .generateType(GenerateType.INPUT_STREAM)
        .creator("MR Excel")
        .build());

Utils.saveZipToFile(ConverterUtil.inputStreamToByteArray(data.getFileData()), "ex2.xlsx");
```

style, formula

```java
String font = "Times New Roman";
String c1 = "08D9D6";
String c2 = "252A34";
String c3 = "FF2E63";
String c4 = "EAEAEA";
GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Arrays.asList(
        SheetGenerator.builder()
                .applyHeaderOptionFunction(headers -> {
                    headers.forEach(header -> {
                        if (header.getText().equals("id")) {
                            header.setSize(20);
                            header.setFormula(FormulaSetting.builder()
                                    .type(FormulaSetting.FormulaType.COUNT)
                                    .styleId("formulaStyle")
                                    .build());
                        } else if (header.getLabel().equals("birthDate")) {
                            header.setSize(30);
                        } else {
                            header.setSize(20);
                        }
                    });
                    return headers;
                })
                .applySheetOptionFunction(
                        sheetBuilder -> {
                            sheetBuilder.headerStyleKey("headerStyle").headerHeight(30);
                            return sheetBuilder.build();
                        }
                )
                .data(
                        ExampleData.ex2Sheet1()
                )
                .build(),
        SheetGenerator.builder()
                .applyHeaderOptionFunction(headers -> {
                    headers.forEach(header -> {
                        switch (header.getLabel()) {
                            case "id":
                                header.setFormula(FormulaSetting.builder()
                                        .type(FormulaSetting.FormulaType.COUNT)
                                        .styleId("formulaStyle")
                                        .build());
                            case "surname":
                                header.setSize(70);
                                break;
                            case "name":
                                header.setSize(12);
                                break;
                            default:
                                header.setSize(100);
                                break;

                        }
                    });
                    return headers;
                })
                .applySheetOptionFunction(
                        sheetBuilder -> sheetBuilder
                                .headerStyleKey("headerStyle")
                                .withoutHeader(true)
                                .build())
                .data(
                        ExampleData.ex2Sheet2()
                ).build()
), ExcelTableOption
        .eBuilder()
        .addStyles("headerStyle", StyleBody.builder()
                .backgroundColor(c2)
                .fontFamily(font)
                .color(c4)
                .size(20)
                .build())
        .addStyles("formulaStyle", StyleBody.builder()
                .backgroundColor(c1)
                .fontFamily(font)
                .color(c3)
                .size(15)
                .build())
        .created("2023-08-06T07:22:40Z")
        .modified("2023-08-06T07:22:40Z")
        .fileName("ex3")
        .creator("MR Excel")
        .build());
```

formula

```java
String font = "Times New Roman";
String c1 = "2B2E4A";
String c2 = "E84545";
String c3 = "903749";
String c4 = "53354A";
Result<InputStream> data = GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Arrays.asList(
        SheetGenerator.builder()
                .applyHeaderOptionFunction(headers -> {
                    headers.forEach(header -> {
                        switch (header.getLabel()){
                            case "date", "column5":
                                header.setFormula(FormulaSetting.builder().type(FormulaSetting.FormulaType.COUNT).styleId("formulaStyle").build());
                                break;
                            case "column1":
                                header.setFormula(FormulaSetting.builder().type(FormulaSetting.FormulaType.AVERAGE).styleId("formulaStyle").build());
                                break;
                            case "column2":
                                header.setFormula(FormulaSetting.builder().type(FormulaSetting.FormulaType.SUM).styleId("formulaStyle").build());
                                break;
                            case "column3":
                                header.setFormula(FormulaSetting.builder().type(FormulaSetting.FormulaType.MAX).styleId("formulaStyle").build());
                                break;
                            case "column4":
                                header.setFormula(FormulaSetting.builder().type(FormulaSetting.FormulaType.MIN).styleId("formulaStyle").build());
                                break;
                            default:
                                System.out.println("not found column");
                        }
                    });
                    return headers;
                })
                .applySheetOptionFunction(
                        sheetBuilder -> {
                            Map<String, FormulaMapBody> map=new HashMap<>();
                            map.put("A8",FormulaSetting.builder()
                                    .type(FormulaSetting.FormulaType.SUM)
                                    .start("B2")
                                    .end("D3")
                                    .styleId("customFormulaStyle")
                                    .build());
                            map.put("B8",FormulaSetting.builder()
                                    .type(FormulaSetting.FormulaType.AVERAGE).start("A2").end("F6")
                                    .styleId("customFormulaStyle").build());
                            map.put("C8",FormulaSetting.builder()
                                    .type(FormulaSetting.FormulaType.SUM).start("A2").end("F6")
                                    .styleId("customFormulaStyle").build());
                            map.put("D8",FormulaSetting.builder()
                                    .type(FormulaSetting.FormulaType.MAX).start("A2").end("F6")
                                    .styleId("customFormulaStyle").build());
                            map.put("E8",FormulaSetting.builder()
                                    .type(FormulaSetting.FormulaType.MIN).start("A2").end("F6")
                                    .styleId("customFormulaStyle").build());
                            map.put("F8",FormulaSetting.builder()
                                    .type(FormulaSetting.FormulaType.COUNT).start("A2").end("F6")
                                    .styleId("customFormulaStyle").build());
                            sheetBuilder.formula(map).headerStyleKey("headerStyle");
                            return sheetBuilder.build();
                        }
                )
                .data(
                        ExampleData.ex4()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .addStyles("headerStyle", StyleBody.builder()
                .backgroundColor(c1)
                .fontFamily(font)
                .color(c2)
                .build())
        .addStyles("customFormulaStyle", StyleBody.builder()
                .backgroundColor(c2)
                .fontFamily(font)
                .color(c1)
                .size(15)
                .border(StyleBody.BorderOption.builder()
                        .full(StyleBody.BorderOption.BorderBody.builder()
                                .color(c4)
                                .style(StyleBody.BorderOption.BorderBody.BorderStyle.DASH_DOT)
                                .build())
                        .build())
                .build())
        .addStyles("formulaStyle", StyleBody.builder()
                .backgroundColor(c1)
                .fontFamily(font)
                .color(c2)
                .size(15)
                .border(StyleBody.BorderOption.builder()
                        .full(StyleBody.BorderOption.BorderBody.builder()
                                .color(c3)
                                .style(StyleBody.BorderOption.BorderBody.BorderStyle.MEDIUM)
                                .build())
                        .build())
                .build())
        .created("2023-08-06T07:22:40Z")
        .modified("2023-08-06T07:22:40Z")
        .fileName("ex4")
        .creator("MR Excel")
        .build());
```

format

```java
String font = "Times New Roman";
String c1 = "2C3639";
String c2 = "3F4E4F";
String c3 = "A27B5C";
String c4 = "DCD7C9";
Map<String, StyleBody> styles=new HashMap<>();
StyleBody rowStyle=StyleBody.builder()
        .backgroundColor(c2)
        .fontFamily(font)
        .color(c4)
        .build();
StyleBody headerStyle= StyleBody.builder()
        .backgroundColor(c4)
        .fontFamily(font)
        .color(c2)
        .build();
StyleBody date=(StyleBody) rowStyle.clone();
date.setFormat("short_date");
StyleBody time=(StyleBody) rowStyle.clone();
time.setFormat("time");
StyleBody percentage=(StyleBody) rowStyle.clone();
percentage.setFormat("percentage");
StyleBody float2=(StyleBody) rowStyle.clone();
float2.setFormat("float_2");
float2.setAlignment(StyleBody.AlignmentOption.builder()
                .horizontal(StyleBody.AlignmentOption.AlignmentHorizontal.LEFT)
        .build());
StyleBody fraction=(StyleBody) rowStyle.clone();
fraction.setFormat("fraction");
StyleBody dollarRounded=(StyleBody) rowStyle.clone();
dollarRounded.setFormat("dollar_rounded");
dollarRounded.setAlignment(
        StyleBody.AlignmentOption.builder()
                .indent(3)
                .rtl(true)
                .horizontal(StyleBody.AlignmentOption.AlignmentHorizontal.LEFT)
                .build()
);
StyleBody numSep=(StyleBody) rowStyle.clone();
numSep.setFormat("num_sep");
numSep.setAlignment(StyleBody.AlignmentOption.builder()
                .ltr(true)
                .horizontal(StyleBody.AlignmentOption.AlignmentHorizontal.LEFT)
        .build());
styles.put("Date",date);
styles.put("Time",time);
styles.put("Percentage",percentage);
styles.put("Float",float2);
styles.put("Fraction",fraction);
styles.put("Long Number Column 1",dollarRounded);
styles.put("Long Number Column 2",numSep);
styles.put("headerStyle",headerStyle);
GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Arrays.asList(
        SheetGenerator.builder()
                .applyHeaderOptionFunction(headers -> {
                        headers.forEach(header -> {
                        if (header.getLabel().equals("longNumberColumn1")||header.getLabel().equals("longNumberColumn2")){
                                header.setSize(50);
                        }
                        });
                        return headers;
                })
                .applySheetOptionFunction(
                        sheetBuilder -> {
                                sheetBuilder.styleCellCondition(styleCellConditionInputFunction -> {
                                if(styleCellConditionInputFunction.getFromHeader()){
                                        return "headerStyle";
                                }else{

                                        switch (styleCellConditionInputFunction.getColIndex()){
                                        case 0:
                                                return "Date";
                                        case 1:
                                                return "Time";
                                        case 2:
                                                return "Percentage";
                                        case 3:
                                                return "Float";
                                        case 4:
                                                return "Fraction";
                                        case 5:

                                                return "Long Number Column 1";
                                        default:
                                                return "Long Number Column 2";
                                        }

                                }
                                });
                                return sheetBuilder.build();
                        }
                )
                .data(
                        ExampleData.ex5()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .styles(styles)
        .fileName("ex5")
        .creator("MR Excel")
        .build());

```

format, style, merge cell

```java 
String font = "Times New Roman";
String c1 = "DBE2EF";
String c2 = "112D4E";
String c4 = "F9F7F7";
Map<String, StyleBody> styles=new HashMap<>();
StyleBody rowStyle=StyleBody.builder()
        .backgroundColor(c2)
        .fontFamily(font)
        .color(c4)
        .border(StyleBody.BorderOption.builder()
                .full(
                        StyleBody.BorderOption.BorderBody.builder()
                                .style(StyleBody.BorderOption.BorderBody.BorderStyle.MEDIUM)
                                .color(c1)
                                .build()
                )
                .build())
        .alignment(
                StyleBody.AlignmentOption.builder()
                        .horizontal(
                                StyleBody.AlignmentOption.AlignmentHorizontal.LEFT
                        )
                        .vertical(
                                StyleBody.AlignmentOption.AlignmentVertical.TOP
                        )
                        .build()
        )
        .build();
StyleBody headerStyle= StyleBody.builder()
        .backgroundColor(c4)
        .fontFamily(font)
        .color(c2)
        .build();
StyleBody date=(StyleBody) rowStyle.clone();
date.setFormat("short_date");
StyleBody time=(StyleBody) rowStyle.clone();
time.setFormat("time");
StyleBody percentage=(StyleBody) rowStyle.clone();
percentage.setFormat("percentage");
StyleBody float2=(StyleBody) rowStyle.clone();
float2.setFormat("float_2");
StyleBody fraction=(StyleBody) rowStyle.clone();
fraction.setFormat("fraction");
StyleBody dollarRounded=(StyleBody) rowStyle.clone();
dollarRounded.setFormat("dollar_2");
StyleBody numSep=(StyleBody) rowStyle.clone();
numSep.setFormat("num_sep");
styles.put("Date",date);
styles.put("Time",time);
styles.put("Percentage",percentage);
styles.put("Float",float2);
styles.put("Fraction",fraction);
styles.put("Long Number Column 1",dollarRounded);
styles.put("Long Number Column 2",numSep);
styles.put("headerStyle",headerStyle);
styles.put("headerStyle",headerStyle);
final Map<String,String> merge=new HashMap<>();
final AtomicBoolean mergeStart=new AtomicBoolean(false);

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Arrays.asList(
        SheetGenerator.builder()
                .applyHeaderOptionFunction(headers -> {
                        headers.forEach(header -> {
                        if (header.getLabel().equals("longNumberColumn1")||header.getLabel().equals("longNumberColumn2")){
                                header.setSize(50);
                        }
                        });
                        return headers;
                })
                .applySheetOptionFunction(
                        sheetBuilder -> {
                                sheetBuilder.mergeRowDataCondition(mergeRowDataConditionInputFunction -> {
                                if(mergeRowDataConditionInputFunction.getFromHeader()){
                                        return false;
                                }else{
                                        var keyName=mergeRowDataConditionInputFunction.getKey().getName();
                                        String mergeData=merge.get(keyName);
                                        String data=mergeRowDataConditionInputFunction.getData().toString();

                                        if(mergeStart.get()){
                                        if(mergeData!=null&&mergeData.equals(data)){

                                                return true;
                                        }else{
                                                merge.put(keyName,data);
                                                return false;
                                        }
                                        }else{
                                        mergeStart.set(true);
                                        merge.put(keyName,data);

                                        return true;
                                        }
                                }
                                });
                                sheetBuilder.styleCellCondition(styleCellConditionInputFunction -> {
                                if(styleCellConditionInputFunction.getFromHeader()){
                                        return "headerStyle";
                                }else{
                                        switch (styleCellConditionInputFunction.getColIndex()){
                                        case 0:
                                                return "Date";
                                        case 1:
                                                return "Time";
                                        case 2:
                                                return "Percentage";
                                        case 3:
                                                return "Float";
                                        case 4:
                                                return "Fraction";
                                        case 5:

                                                return "Long Number Column 1";
                                        default:
                                                return "Long Number Column 2";
                                        }

                                }
                                });
                                return sheetBuilder.build();
                        }
                )
                .data(
                        ExampleData.ex6()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .styles(styles)
        .fileName("ex6")
        .creator("MR Excel")
        .build());
```

outlinelevel, hidden

```java
String font = "Times New Roman";
String c2 = "AEE1E1";
String c4 = "FCD1D1";
String c1="112D4E";
Map<String, StyleBody> styles=new HashMap<>();
StyleBody rowStyle=StyleBody.builder()
        .backgroundColor(c2)
        .fontFamily(font)
        .color(c1)
        .alignment(
                StyleBody.AlignmentOption.builder()
                        .horizontal(
                                StyleBody.AlignmentOption.AlignmentHorizontal.LEFT
                        )
                        .vertical(
                                StyleBody.AlignmentOption.AlignmentVertical.TOP
                        )
                        .build()
        )
        .build();
StyleBody headerStyle= StyleBody.builder()
        .backgroundColor(c4)
        .fontFamily(font)
        .color(c1)
        .build();
styles.put("rowStyle",rowStyle);
styles.put("headerStyle",headerStyle);

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Arrays.asList(
        SheetGenerator.builder()
                .applyHeaderOptionFunction(headers -> {
                        headers.forEach(header -> {
                        if (header.getLabel().equals("longNumberColumn1")||header.getLabel().equals("longNumberColumn2")){
                                header.setSize(50);
                        }
                        });
                        return headers;
                })
                .applySheetOptionFunction(
                        sheetBuilder -> {
                                sheetBuilder.styleCellCondition(styleCellConditionInputFunction -> {
                                if(styleCellConditionInputFunction.getFromHeader()){
                                        return "headerStyle";
                                }else{
                                        return "rowStyle";
                                }
                                });
                                return sheetBuilder.build();
                        }
                )
                .data(
                        ExampleData.ex7()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .styles(styles)
        .fileName("ex7")
        .creator("MR Excel")
        .build());
```

change getter method of hidden, outlinelevel, height

```java
String font = "Times New Roman";
String c2 = "F9ED69";
String c4 = "F08A5D";
String c1="6A2C70";
Map<String, StyleBody> styles=new HashMap<>();
StyleBody rowStyle=StyleBody.builder()
        .backgroundColor(c2)
        .fontFamily(font)
        .color(c1)
        .alignment(
                StyleBody.AlignmentOption.builder()
                        .horizontal(
                                StyleBody.AlignmentOption.AlignmentHorizontal.LEFT
                        )
                        .vertical(
                                StyleBody.AlignmentOption.AlignmentVertical.CENTER
                        )
                        .build()
        )
        .build();
StyleBody headerStyle= StyleBody.builder()
        .backgroundColor(c4)
        .fontFamily(font)
        .color(c1)
        .build();
styles.put("rowStyle",rowStyle);
styles.put("headerStyle",headerStyle);
Class<OutlineCustomPropertyTest> outlineCustomPropertyTestClass=OutlineCustomPropertyTest.class;
Method hightMethod=outlineCustomPropertyTestClass.getMethod("getH");
Method hiddenMethod=outlineCustomPropertyTestClass.getMethod("getNotShow");
Method outlineMethod=outlineCustomPropertyTestClass.getMethod("getGroup");
GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Arrays.asList(
        SheetGenerator.builder()
                .applySheetOptionFunction(
                        sheetBuilder -> {
                                sheetBuilder.mapSheetDataOption(
                                        Sheet.MapSheetDataOption.<OutlineCustomPropertyTest>builder()
                                                .height(hightMethod)
                                                .hidden(hiddenMethod)
                                                .outlineLevel(outlineMethod)
                                        .build()
                                );
                                sheetBuilder.styleCellCondition(styleCellConditionInputFunction -> {
                                if(styleCellConditionInputFunction.getFromHeader()){
                                        return "headerStyle";
                                }else{
                                        return "rowStyle";
                                }
                                });
                                return sheetBuilder.build();
                        }
                )
                .data(
                        ExampleData.ex8()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .styles(styles)
        .fileName("ex8")
        .creator("MR Excel")
        .build());
```

styleCellCondition

```java
String font = "Times New Roman";
String c2 = "E84545";
String c4 = "2B2E4A";
String c1="6A2C70";
StyleBody.AlignmentOption rowAlignment= StyleBody.AlignmentOption.builder()
        .horizontal(
                StyleBody.AlignmentOption.AlignmentHorizontal.LEFT
        )
        .vertical(
                StyleBody.AlignmentOption.AlignmentVertical.CENTER
        )
        .build();
Map<String, StyleBody> styles=new HashMap<>();
StyleBody rowStyle=StyleBody.builder()
        .backgroundColor(c2)
        .fontFamily(font)
        .color(c1)
        .alignment(
                rowAlignment
        )
        .build();
StyleBody headerStyle= StyleBody.builder()
        .backgroundColor(c4)
        .fontFamily(font)
        .color("#FFFFFF")
        .build();
StyleBody higherBody= StyleBody.builder()
        .alignment(rowAlignment)
        .backgroundColor("DCD6F7")
        .color("424874")
        .build();
StyleBody male= StyleBody.builder()
        .backgroundColor("95E1D3")
        .color("252A34")
        .alignment(rowAlignment)
        .build();
StyleBody female= StyleBody.builder()
        .backgroundColor("F38181")
        .color("252A34")
        .alignment(rowAlignment)
        .build();
styles.put("rowStyle",rowStyle);
styles.put("headerStyle",headerStyle);
styles.put("c0<0.3",higherBody);
styles.put("male",male);
styles.put("female",female);
GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Arrays.asList(
        SheetGenerator.builder()
                .applySheetOptionFunction(
                        sheetBuilder -> {
                                sheetBuilder.styleCellCondition(styleCellConditionInputFunction -> {
                                if(styleCellConditionInputFunction.getFromHeader()){
                                        return "headerStyle";
                                }else{
                                        Object data=styleCellConditionInputFunction.getData();
                                        Number dataInt=null;
                                        if(data instanceof Number){
                                                dataInt=(Number) data;
                                        }
                                        if (styleCellConditionInputFunction.getColIndex() == 0 &&dataInt!=null &&dataInt.doubleValue() < 0.3) {
                                                return "c0<0.3";
                                        } else if (styleCellConditionInputFunction.getColIndex() == 3) {
                                                if (dataInt!=null && dataInt.intValue()>0) {
                                                        return "male";
                                                } else {
                                                        return "female";
                                                }
                                        } else {
                                                return "rowStyle";
                                        }
                                }
                                });
                                return sheetBuilder.build();
                        }
                )
                .data(
                        ExampleData.ex9()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .styles(styles)
        .fileName("ex9")
        .creator("MR Excel")
        .build());
```

shift sheet

```java
String font = "Times New Roman";
String c2 = "8785A2";
String c4 = "FFC7C7";
String c1="6A2C70";
StyleBody.AlignmentOption rowAlignment= StyleBody.AlignmentOption.builder()
        .horizontal(
                StyleBody.AlignmentOption.AlignmentHorizontal.LEFT
        )
        .vertical(
                StyleBody.AlignmentOption.AlignmentVertical.CENTER
        )
        .build();
Map<String, StyleBody> styles=new HashMap<>();
StyleBody rowStyle=StyleBody.builder()
        .backgroundColor(c2)
        .fontFamily(font)
        .color(c1)
        .alignment(
                rowAlignment
        )
        .build();
StyleBody headerStyle= StyleBody.builder()
        .backgroundColor(c4)
        .fontFamily(font)
        .color("#FFFFFF")
        .build();
StyleBody higherBody= StyleBody.builder()
        .alignment(rowAlignment)
        .backgroundColor("DCD6F7")
        .color("424874")
        .build();
StyleBody male= StyleBody.builder()
        .backgroundColor("95E1D3")
        .color("252A34")
        .alignment(rowAlignment)
        .build();
StyleBody female= StyleBody.builder()
        .backgroundColor("F38181")
        .color("252A34")
        .alignment(rowAlignment)
        .build();
styles.put("rowStyle",rowStyle);
styles.put("headerStyle",headerStyle);
styles.put("c0<0.3",higherBody);
styles.put("male",male);
styles.put("female",female);
GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Arrays.asList(
        SheetGenerator.builder()
                .applySheetOptionFunction(
                        sheetBuilder -> {
                                sheetBuilder.shiftTop(5);
                                sheetBuilder.shiftLeft(5);
                                sheetBuilder.styleCellCondition(styleCellConditionInputFunction -> {
                                if(styleCellConditionInputFunction.getFromHeader()){
                                        return "headerStyle";
                                }else{
                                        Object data=styleCellConditionInputFunction.getData();
                                        Number dataInt=null;
                                        if(data instanceof Number){
                                        dataInt=(Number) data;
                                        }
                                        if (styleCellConditionInputFunction.getColIndex()-5 == 0 &&dataInt!=null &&dataInt.doubleValue() < 0.3) {
                                        return "c0<0.3";
                                        } else if (styleCellConditionInputFunction.getColIndex()-5 == 3) {
                                        if (dataInt!=null && dataInt.intValue()>0) {
                                                return "male";
                                        } else {
                                                return "female";
                                        }
                                        } else {
                                        return "rowStyle";
                                        }
                                }
                                });
                                return sheetBuilder.build();
                        }
                )
                .data(
                        ExampleData.ex9()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .styles(styles)
        .fileName("ex10")
        .creator("MR Excel")
        .build());
```

title, shift sheet

```java
String font = "Times New Roman";
String c2 = "8785A2";
String c4 = "FFC7C7";
String c1 = "6A2C70";
StyleBody.AlignmentOption rowAlignment = StyleBody.AlignmentOption.builder()
        .horizontal(
                StyleBody.AlignmentOption.AlignmentHorizontal.LEFT
        )
        .vertical(
                StyleBody.AlignmentOption.AlignmentVertical.CENTER
        )
        .build();
Map<String, StyleBody> styles = new HashMap<>();
StyleBody rowStyle = StyleBody.builder()
        .backgroundColor(c2)
        .fontFamily(font)
        .color(c1)
        .alignment(
                rowAlignment
        )
        .build();
StyleBody headerStyle = StyleBody.builder()
        .backgroundColor(c4)
        .fontFamily(font)
        .color("#FFFFFF")
        .build();
StyleBody higherBody = StyleBody.builder()
        .alignment(rowAlignment)
        .backgroundColor("DCD6F7")
        .color("424874")
        .build();
StyleBody male = StyleBody.builder()
        .backgroundColor("95E1D3")
        .color("252A34")
        .alignment(rowAlignment)
        .build();
StyleBody female = StyleBody.builder()
        .backgroundColor("F38181")
        .color("252A34")
        .alignment(rowAlignment)
        .build();
styles.put("rowStyle", rowStyle);
styles.put("headerStyle", headerStyle);
styles.put("c0<0.3", higherBody);
styles.put("male", male);
styles.put("female", female);
GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Collections.singletonList(
        SheetGenerator.builder()
                .applySheetOptionFunction(
                        sheetBuilder -> {
                                sheetBuilder.title(
                                        Title.builder()
                                                .text("Title")
                                                .consommeRow(4)
                                                .consommeCol(6)
                                                .shiftLeft(-1)
                                                .shiftTop(1)
                                                .build()
                                );
                                sheetBuilder.shiftTop(1);
                                sheetBuilder.shiftLeft(1);
                                sheetBuilder.styleCellCondition(styleCellConditionInputFunction -> {
                                if (styleCellConditionInputFunction.getFromHeader()) {
                                        return "headerStyle";
                                } else {
                                        Object data = styleCellConditionInputFunction.getData();
                                        Number dataInt = null;
                                        if (data instanceof Number) {
                                        dataInt = (Number) data;
                                        }
                                        if (styleCellConditionInputFunction.getColIndex() - 1 == 0 && dataInt != null && dataInt.doubleValue() < 0.3) {
                                        return "c0<0.3";
                                        } else if (styleCellConditionInputFunction.getColIndex() - 1 == 3) {
                                        if (dataInt != null && dataInt.intValue() > 0) {
                                                return "male";
                                        } else {
                                                return "female";
                                        }
                                        } else {
                                        return "rowStyle";
                                        }
                                }
                                });
                                return sheetBuilder.build();
                        }
                )
                .data(
                        ExampleData.ex9()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .addDefaultTitleStyle(true)
        .styles(styles)
        .fileName("ex11")
        .creator("MR Excel")
        .build());

```

and other

```java

public static void ex14() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String font = "Times New Roman";
String c2 = "FFF6E0";
String c4 = "61677A";
String c1 = "6A2C70";
StyleBody.AlignmentOption rowAlignment = StyleBody.AlignmentOption.builder()
        .horizontal(
                StyleBody.AlignmentOption.AlignmentHorizontal.LEFT
        )
        .vertical(
                StyleBody.AlignmentOption.AlignmentVertical.CENTER
        )
        .build();
Map<String, StyleBody> styles = new HashMap<>();
StyleBody rowStyle = StyleBody.builder()
        .backgroundColor(c2)
        .fontFamily(font)
        .color(c1)
        .alignment(
                rowAlignment
        )
        .build();
StyleBody headerStyle = StyleBody.builder()
        .backgroundColor(c4)
        .fontFamily(font)
        .color("#FFFFFF")
        .build();
StyleBody higherBody = StyleBody.builder()
        .alignment(rowAlignment)
        .backgroundColor("DCD6F7")
        .color("424874")
        .build();
StyleBody male = StyleBody.builder()
        .backgroundColor("95E1D3")
        .color("252A34")
        .alignment(rowAlignment)
        .build();
StyleBody female = StyleBody.builder()
        .backgroundColor("F38181")
        .color("252A34")
        .alignment(rowAlignment)
        .build();
styles.put("rowStyle", rowStyle);
styles.put("headerStyle", headerStyle);
styles.put("c0<0.3", higherBody);
styles.put("male", male);
styles.put("female", female);
GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Collections.singletonList(
        SheetGenerator.builder()
                .applyHeaderOptionFunction(headers -> {
                        headers.get(0).setComment(new Comment("misspell in header", null, "MR Excel"));
                        return headers;
                })
                .applySheetOptionFunction(
                        sheetBuilder -> {
                                sheetBuilder.title(
                                        Title.builder()
                                                .comment(new Comment("This is comment on table"))
                                                .text("Title")
                                                .consommeRow(4)
                                                .consommeCol(6)
                                                .shiftLeft(-1)
                                                .shiftTop(1)
                                                .build()
                                );
                                sheetBuilder.commentCondition(commentConditionFunctionInput -> {
                                if (commentConditionFunctionInput.getFromHeader()) {
                                        Object mainData = commentConditionFunctionInput.getData();

                                        if (!(mainData instanceof Header)) {
                                        return new Comment("Header is invalid");
                                        }
                                        String data = ((Header) mainData).getText();

                                        String textDataC0 = String.valueOf(data.charAt(0));
                                        if (!textDataC0.toUpperCase().equals(textDataC0)) {
                                        return new Comment("Header should start with " + textDataC0.toUpperCase(), null, "System");
                                        }
                                }
                                return null;
                                }).shiftTop(1).shiftLeft(1).styleCellCondition(styleCellConditionInputFunction -> {
                                if (styleCellConditionInputFunction.getFromHeader()) {
                                        return "headerStyle";
                                } else {
                                        Object data = styleCellConditionInputFunction.getData();
                                        Number dataInt = null;
                                        if (data instanceof Number) {
                                        dataInt = (Number) data;
                                        }
                                        if (styleCellConditionInputFunction.getColIndex() - 1 == 0 && dataInt != null && dataInt.doubleValue() < 0.3) {
                                        return "c0<0.3";
                                        } else if (styleCellConditionInputFunction.getColIndex() - 1 == 3) {
                                        if (dataInt != null && dataInt.intValue() > 0) {
                                                return "male";
                                        } else {
                                                return "female";
                                        }
                                        } else {
                                        return "rowStyle";
                                        }
                                }
                                });
                                return sheetBuilder.build();
                        }
                )
                .data(
                        ExampleData.ex14()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .addDefaultTitleStyle(true)
        .styles(styles)
        .fileName("ex14")
        .creator("MR Excel")
        .build());
}

public static void ex15_1() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String font = "Times New Roman";
String c2 = "5B9A8B";
String c4 = "F7E987";
String c1 = "6A2C70";
StyleBody.AlignmentOption rowAlignment = StyleBody.AlignmentOption.builder()
        .horizontal(
                StyleBody.AlignmentOption.AlignmentHorizontal.LEFT
        )
        .vertical(
                StyleBody.AlignmentOption.AlignmentVertical.CENTER
        )
        .build();
Map<String, StyleBody> styles = new HashMap<>();
StyleBody rowStyle = StyleBody.builder()
        .backgroundColor(c2)
        .fontFamily(font)
        .color(c1)
        .alignment(
                rowAlignment
        )
        .build();
StyleBody headerStyle = StyleBody.builder()
        .backgroundColor(c4)
        .fontFamily(font)
        .color("#FFFFFF")
        .build();
StyleBody higherBody = StyleBody.builder()
        .alignment(rowAlignment)
        .backgroundColor("DCD6F7")
        .color("424874")
        .build();
StyleBody male = StyleBody.builder()
        .backgroundColor("95E1D3")
        .color("252A34")
        .alignment(rowAlignment)
        .build();
StyleBody female = StyleBody.builder()
        .backgroundColor("F38181")
        .color("252A34")
        .alignment(rowAlignment)
        .build();
styles.put("rowStyle", rowStyle);
styles.put("headerStyle", headerStyle);
styles.put("c0<0.3", higherBody);
styles.put("male", male);
styles.put("female", female);
styles.put("headC1", StyleBody.builder()
        .size(15).color("16404D")
        .build());
styles.put("headC2", StyleBody.builder()
        .size(10).color("A31D1D")
        .build());
styles.put("bodyC1", StyleBody.builder()
        .size(10).color("FBF5E5")
        .build());
styles.put("bodyC2", StyleBody.builder()
        .size(15).color("FFEB00")
        .build());
styles.put("title", StyleBody.builder()
        .size(48)
        .backgroundColor("E5BA73")
        .alignment(
                StyleBody.AlignmentOption.builder()
                        .vertical(StyleBody.AlignmentOption.AlignmentVertical.CENTER)
                        .horizontal(StyleBody.AlignmentOption.AlignmentHorizontal.CENTER)
                        .build()
        )
        .build());
GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Collections.singletonList(
        SheetGenerator.builder()
                .applyHeaderOptionFunction(headers -> {
                        headers.get(0).setComment(new Comment("misspell in header", null, "MR Excel"));
                        return headers;
                })
                .applySheetOptionFunction(
                        sheetBuilder -> {
                                sheetBuilder.multiStyleCondition(multiStyleConditionInputFunction -> {
                                if (multiStyleConditionInputFunction.getFromHeader()) {
                                        Header data = (Header) multiStyleConditionInputFunction.getData();
                                        var sp = data.getText().split("");
                                        int index = 0;
                                        List<MultiStyleValue> result = new ArrayList<>();
                                        for (String current : sp) {
                                        result.add(new MultiStyleValue(current, index % 2 == 0 ? "headC1" : "headC2"));
                                        index++;
                                        }
                                        return result;
                                } else {
                                        String data = multiStyleConditionInputFunction.getData().toString();
                                        var sp = data.split("");
                                        List<MultiStyleValue> result = new ArrayList<>();
                                        int index = 0;
                                        for (String current : sp) {
                                        result.add(new MultiStyleValue(current, index % 2 == 0 ? "bodyC1" : "bodyC2"));
                                        index++;
                                        }
                                        return result;
                                }
                                });
                                sheetBuilder.title(
                                        Title.builder()
                                                .comment(new Comment("This is comment on table"))
                                                .text("Title")
                                                .styleId("title")
                                                .consommeRow(4)
                                                .consommeCol(6)
                                                .shiftLeft(-1)
                                                .shiftTop(1)
                                                .build()
                                );
                                sheetBuilder.commentCondition(commentConditionFunctionInput -> {
                                if (commentConditionFunctionInput.getFromHeader()) {
                                        Object mainData = commentConditionFunctionInput.getData();

                                        if (!(mainData instanceof Header)) {
                                        return new Comment("Header is invalid");
                                        }
                                        String data = ((Header) mainData).getText();

                                        String textDataC0 = String.valueOf(data.charAt(0));
                                        if (!textDataC0.toUpperCase().equals(textDataC0)) {
                                        return new Comment("Header should start with " + textDataC0.toUpperCase(), null, "System");
                                        }
                                }
                                return null;
                                }).shiftTop(1).shiftLeft(1).styleCellCondition(styleCellConditionInputFunction -> {
                                if (styleCellConditionInputFunction.getFromHeader()) {
                                        return "headerStyle";
                                } else {
                                        Object data = styleCellConditionInputFunction.getData();
                                        Number dataInt = null;
                                        if (data instanceof Number) {
                                        dataInt = (Number) data;
                                        }
                                        if (styleCellConditionInputFunction.getColIndex() - 1 == 0 && dataInt != null && dataInt.doubleValue() < 0.3) {
                                        return "c0<0.3";
                                        } else if (styleCellConditionInputFunction.getColIndex() - 1 == 3) {
                                        if (dataInt != null && dataInt.intValue() > 0) {
                                                return "male";
                                        } else {
                                                return "female";
                                        }
                                        } else {
                                        return "rowStyle";
                                        }
                                }
                                });
                                return sheetBuilder.build();
                        }
                )
                .data(
                        ExampleData.ex14()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .addDefaultTitleStyle(true)
        .styles(styles)
        .fileName("ex15.1")
        .creator("MR Excel")
        .build());
}


public static void ex15() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String font = "Times New Roman";
String c2 = "5B9A8B";
String c4 = "F7E987";
String c1 = "6A2C70";
StyleBody.AlignmentOption rowAlignment = StyleBody.AlignmentOption.builder()
        .horizontal(
                StyleBody.AlignmentOption.AlignmentHorizontal.LEFT
        )
        .vertical(
                StyleBody.AlignmentOption.AlignmentVertical.CENTER
        )
        .build();
Map<String, StyleBody> styles = new HashMap<>();
StyleBody rowStyle = StyleBody.builder()
        .backgroundColor(c2)
        .fontFamily(font)
        .color(c1)
        .alignment(
                rowAlignment
        )
        .build();
StyleBody headerStyle = StyleBody.builder()
        .backgroundColor(c4)
        .fontFamily(font)
        .color("#000000")
        .build();
StyleBody higherBody = StyleBody.builder()
        .alignment(rowAlignment)
        .backgroundColor("DCD6F7")
        .color("424874")
        .build();
StyleBody male = StyleBody.builder()
        .backgroundColor("95E1D3")
        .color("252A34")
        .alignment(rowAlignment)
        .build();
StyleBody female = StyleBody.builder()
        .backgroundColor("F38181")
        .color("252A34")
        .alignment(rowAlignment)
        .build();
styles.put("rowStyle", rowStyle);
styles.put("headerStyle", headerStyle);
styles.put("c0<0.3", higherBody);
styles.put("male", male);
styles.put("female", female);
styles.put("headC1", StyleBody.builder()
        .size(15).color("16404D")
        .build());
styles.put("headC2", StyleBody.builder()
        .size(10).color("A31D1D")
        .build());
styles.put("bodyC1", StyleBody.builder()
        .size(10).color("FBF5E5")
        .build());
styles.put("bodyC2", StyleBody.builder()
        .size(15).color("FFEB00")
        .build());
styles.put("title", StyleBody.builder()
        .size(48)
        .backgroundColor("E5BA73")
        .alignment(
                StyleBody.AlignmentOption.builder()
                        .vertical(StyleBody.AlignmentOption.AlignmentVertical.CENTER)
                        .horizontal(StyleBody.AlignmentOption.AlignmentHorizontal.CENTER)
                        .build()
        )
        .build());
styles.put("t1", StyleBody.builder()
        .color("555555")
        .build());
styles.put("t2", StyleBody.builder()
        .size(40)
        .color("FFFFFF")
        .build());
GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Collections.singletonList(
        SheetGenerator.builder()
                .applyHeaderOptionFunction(headers -> {
                        headers.get(0).setComment(new Comment("misspell in header", null, "MR Excel"));
                        headers.get(2).setMultiStyleValue(Arrays.asList(
                                MultiStyleValue.builder()
                                        .styleId("t1").value("N")
                                        .build(),
                                MultiStyleValue.builder()
                                        .styleId("t2").value("a")
                                        .build()
                        ));
                        return headers;
                })
                .applySheetOptionFunction(
                        sheetBuilder -> {
                                sheetBuilder.title(
                                        Title.builder()
                                                .comment(new Comment("This is comment on table"))
                                                .text("Can't use this text")
                                                .styleId("title")
                                                .multiStyleValue(
                                                        Arrays.asList(
                                                                MultiStyleValue.builder()
                                                                        .styleId("t2")
                                                                        .value("T")
                                                                        .build(),
                                                                MultiStyleValue.builder()
                                                                        .value("i")
                                                                        .build(),

                                                                MultiStyleValue.builder()
                                                                        .styleId("t2")
                                                                        .value("t")
                                                                        .build(),
                                                                MultiStyleValue.builder()
                                                                        .value("l")
                                                                        .build(),
                                                                MultiStyleValue.builder()
                                                                        .styleId("t2")
                                                                        .value("e")
                                                                        .build()
                                                        )
                                                )
                                                .consommeRow(4)
                                                .consommeCol(6)
                                                .shiftLeft(-1)
                                                .shiftTop(1)
                                                .build()
                                );
                                sheetBuilder.commentCondition(commentConditionFunctionInput -> {
                                if (commentConditionFunctionInput.getFromHeader()) {
                                        Object mainData = commentConditionFunctionInput.getData();

                                        if (!(mainData instanceof Header)) {
                                        return new Comment("Header is invalid");
                                        }
                                        String data = ((Header) mainData).getText();

                                        String textDataC0 = String.valueOf(data.charAt(0));
                                        if (!textDataC0.toUpperCase().equals(textDataC0)) {
                                        return new Comment("Header should start with " + textDataC0.toUpperCase(), null, "System");
                                        }
                                }
                                return null;
                                }).shiftTop(1).shiftLeft(1).styleCellCondition(styleCellConditionInputFunction -> {
                                if (styleCellConditionInputFunction.getFromHeader()) {
                                        return "headerStyle";
                                } else {
                                        Object data = styleCellConditionInputFunction.getData();
                                        Number dataInt = null;
                                        if (data instanceof Number) {
                                        dataInt = (Number) data;
                                        }
                                        if (styleCellConditionInputFunction.getColIndex() - 1 == 0 && dataInt != null && dataInt.doubleValue() < 0.3) {
                                        return "c0<0.3";
                                        } else if (styleCellConditionInputFunction.getColIndex() - 1 == 3) {
                                        if (dataInt != null && dataInt.intValue() > 0) {
                                                return "male";
                                        } else {
                                                return "female";
                                        }
                                        } else {
                                        return "rowStyle";
                                        }
                                }
                                });
                                return sheetBuilder.build();
                        }
                )
                .data(
                        ExampleData.ex15()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .addDefaultTitleStyle(true)
        .styles(styles)
        .fileName("ex15")
        .creator("MR Excel")
        .build());
}

public static void ex16() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String font = "Times New Roman";
String c2 = "393E46";
String c4 = "00ADB5";
String c1 = "6A2C70";
StyleBody.AlignmentOption rowAlignment = StyleBody.AlignmentOption.builder()
        .horizontal(
                StyleBody.AlignmentOption.AlignmentHorizontal.LEFT
        )
        .vertical(
                StyleBody.AlignmentOption.AlignmentVertical.CENTER
        )
        .build();
Map<String, StyleBody> styles = new HashMap<>();
StyleBody rowStyle = StyleBody.builder()
        .backgroundColor(c2)
        .fontFamily(font)
        .color(c1)
        .alignment(
                rowAlignment
        )
        .build();
StyleBody headerStyle = StyleBody.builder()
        .backgroundColor(c4)
        .fontFamily(font)
        .color("#000000")
        .build();
StyleBody higherBody = StyleBody.builder()
        .alignment(rowAlignment)
        .backgroundColor("DCD6F7")
        .color("424874")
        .build();
StyleBody male = StyleBody.builder()
        .backgroundColor("95E1D3")
        .color("252A34")
        .alignment(rowAlignment)
        .build();
StyleBody female = StyleBody.builder()
        .backgroundColor("F38181")
        .color("252A34")
        .alignment(rowAlignment)
        .build();
styles.put("rowStyle", rowStyle);
styles.put("headerStyle", headerStyle);
styles.put("c0<0.3", higherBody);
styles.put("male", male);
styles.put("female", female);
styles.put("headC1", StyleBody.builder()
        .size(15).color("16404D")
        .build());
styles.put("headC2", StyleBody.builder()
        .size(10).color("A31D1D")
        .build());
styles.put("bodyC1", StyleBody.builder()
        .size(10).color("FBF5E5")
        .build());
styles.put("bodyC2", StyleBody.builder()
        .size(15).color("FFEB00")
        .build());
styles.put("title", StyleBody.builder()
        .size(48)
        .backgroundColor("F9ED69")
        .color("6A2C70")
        .alignment(
                StyleBody.AlignmentOption.builder()
                        .vertical(StyleBody.AlignmentOption.AlignmentVertical.CENTER)
                        .horizontal(StyleBody.AlignmentOption.AlignmentHorizontal.CENTER)
                        .build()
        )
        .build());
styles.put("col1", StyleBody.builder()
        .color("#F9ED69")
        .build());
styles.put("col2", StyleBody.builder().color("FF2E63").build());
styles.put("t1", StyleBody.builder()
        .color("555555")
        .size(22)
        .build());
styles.put("t2", StyleBody.builder()
        .size(22)
        .color("F08A5D")
        .build());
GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Collections.singletonList(
        SheetGenerator.builder()
                .applyHeaderOptionFunction(headers -> {
                        headers.get(0).setComment(new Comment("misspell in header", null, "MR Excel"));
                        headers.get(2).setMultiStyleValue(Arrays.asList(
                                MultiStyleValue.builder()
                                        .styleId("t1").value("N")
                                        .build(),
                                MultiStyleValue.builder()
                                        .styleId("t2").value("a")
                                        .build()
                        ));
                        return headers;
                })
                .applySheetOptionFunction(
                        sheetBuilder -> {
                                sheetBuilder.title(
                                        Title.builder()
                                                .comment(new Comment("This is comment on table"))
                                                .text("Can't use this text")
                                                .styleId("title")
                                                .multiStyleValue(
                                                        Arrays.asList(
                                                                MultiStyleValue.builder()
                                                                        .styleId("t2")
                                                                        .value("T")
                                                                        .build(),
                                                                MultiStyleValue.builder()
                                                                        .value("i")
                                                                        .styleId("t1")
                                                                        .build(),

                                                                MultiStyleValue.builder()
                                                                        .styleId("t2")
                                                                        .value("t")
                                                                        .build(),
                                                                MultiStyleValue.builder()
                                                                        .styleId("t1")
                                                                        .value("le 1")
                                                                        .build()
                                                        )
                                                )
                                                .consommeRow(4)
                                                .consommeCol(6)
                                                .shiftLeft(-1)
                                                .shiftTop(1)
                                                .build()
                                ).commentCondition(commentConditionFunctionInput -> {
                                if (commentConditionFunctionInput.getFromHeader()) {
                                        Object mainData = commentConditionFunctionInput.getData();

                                        if (!(mainData instanceof Header)) {
                                        return new Comment("Header is invalid");
                                        }
                                        String data = ((Header) mainData).getText();

                                        String textDataC0 = String.valueOf(data.charAt(0));
                                        if (!textDataC0.toUpperCase().equals(textDataC0)) {
                                        return new Comment("Header should start with " + textDataC0.toUpperCase(), null, "System");
                                        }
                                }
                                return null;
                                }).shiftTop(1).shiftLeft(1).styleCellCondition(styleCellConditionInputFunction -> {
                                if (styleCellConditionInputFunction.getFromHeader()) {
                                        return "headerStyle";
                                } else {
                                        Object data = styleCellConditionInputFunction.getData();
                                        Number dataInt = null;
                                        if (data instanceof Number) {
                                        dataInt = (Number) data;
                                        }
                                        if (styleCellConditionInputFunction.getColIndex() - 1 == 0 && dataInt != null && dataInt.doubleValue() < 0.3) {
                                        return "c0<0.3";
                                        } else if (styleCellConditionInputFunction.getColIndex() - 1 == 3) {
                                        if (dataInt != null && dataInt.intValue() > 0) {
                                                return "male";
                                        } else {
                                                return "female";
                                        }
                                        } else {
                                        return "rowStyle";
                                        }
                                }
                                }).multiStyleCondition(multiStyleConditionInputFunction -> {
                                String data = multiStyleConditionInputFunction.getFromHeader() ? ((Header) multiStyleConditionInputFunction.getData()).getText() : (String) multiStyleConditionInputFunction.getData();
                                String[] charList = data.split("");
                                AtomicInteger index = new AtomicInteger(0);
                                return Arrays.stream(charList).reduce(new ArrayList<>(), (result, current) -> {
                                        result.add(MultiStyleValue.builder()
                                                .styleId(
                                                        index.getAndIncrement() % 2 > 0 ? "col1" : "col2"
                                                ).value(current)
                                                .build());
                                        return result;
                                }, (multiStyleValues, multiStyleValues2) -> {
                                        multiStyleValues.addAll(multiStyleValues2);
                                        return multiStyleValues;
                                });
                                });
                                return sheetBuilder.build();
                        }
                )
                .data(
                        ExampleData.ex16()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .addDefaultTitleStyle(true)
        .styles(styles)
        .fileName("ex16")
        .creator("MR Excel")
        .build());
}

public static void ex17() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Collections.singletonList(
        SheetGenerator.builder()
                .applySheetOptionFunction(
                        sheetBuilder -> {
                                sheetBuilder.images(Arrays.asList(
                                        Sheet.ImageTypes.builder()
                                                .image(ImageInput.builder().data(
                                                                "https://mohammadrezaeicode.github.io/mr-excel-page/img/ezgif.com-gif-maker.gif"
                                                        ).build()
                                                ).from("H1").type(Sheet.ImageTypes.Type.ONE)
                                                .build(),
                                        Sheet.ImageTypes.builder()
                                                .image(ImageInput.builder().data(
                                                                "https://mohammadrezaeicode.github.io/mr-excel-page/img/uniqe.jpg"
                                                        ).build()
                                                ).from("H2").type(Sheet.ImageTypes.Type.ONE)
                                                .build(),
                                        Sheet.ImageTypes.builder()
                                                .image(ImageInput.builder().data(
                                                                "https://mohammadrezaeicode.github.io/mr-excel-page/img/ex.PNG"
                                                        ).build()
                                                ).from("H3").type(Sheet.ImageTypes.Type.ONE)
                                                .build(),
                                        Sheet.ImageTypes.builder()
                                                .image(ImageInput.builder().data(
                                                                "https://mohammadrezaeicode.github.io/mr-excel-page/img/ex.PNG"
                                                        ).build()
                                                ).from("H4").type(Sheet.ImageTypes.Type.ONE)
                                                .build()
                                        ,
                                        Sheet.ImageTypes.builder()
                                                .image(ImageInput.builder().data(
                                                                "https://mohammadrezaeicode.github.io/mr-excel-page/img/ezgif.com-gif-maker.gif"
                                                        ).build()
                                                ).from("E1")
                                                .to("F10")
                                                .type(Sheet.ImageTypes.Type.TWO)
                                                .build(),
                                        Sheet.ImageTypes.builder()
                                                .image(ImageInput.builder().data(
                                                                "https://mohammadrezaeicode.github.io/mr-excel-page/img/uniqe.jpg"
                                                        ).build()
                                                ).from("H6").type(Sheet.ImageTypes.Type.TWO)
                                                .build()
                                ));
                                return sheetBuilder.build();
                        }
                )
                .data(
                        ExampleData.ex17()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .fetch(Utils.fetch())
        .fileName("ex17")
        .creator("MR Excel")
        .build());
}

public static void ex18() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Collections.singletonList(
        SheetGenerator.builder()
                .applyHeaderOptionFunction(headers -> {
                        headers.get(0).setConditionalFormatting(ConditionalFormatting.eBuilder()
                                .type(ConditionalFormattingOption.Type.ICON_SET)
                                .operator(ConditionalFormattingOption.ConditionalFormattingIconSetOperation.ARROWS_3)
                                .build());
                        headers.get(1).setConditionalFormatting(ConditionalFormatting.eBuilder()
                                .type(ConditionalFormattingOption.Type.CELLS)
                                .operator(ConditionalFormattingOption.ConditionalFormattingCellsOperation.CT)
                                .value("a")
                                .build());
                        headers.get(2).setConditionalFormatting(ConditionalFormatting.eBuilder()
                                .type(ConditionalFormattingOption.Type.CELLS)
                                .operator(ConditionalFormattingOption.ConditionalFormattingCellsOperation.GT)
                                .value(10)
                                .styleId("ct")
                                .build());
                        AtomicInteger index = new AtomicInteger(1);
                        headers.forEach(v -> {
                        v.setText("Text " + index.getAndIncrement());
                        });
                        return headers;
                })
                .data(
                        ExampleData.ex18()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .activateConditionalFormatting(true)
        .addStyles("ct", StyleBody.builder()
                .type(StyleBody.StyleType.CONDITIONAL_FORMATTING)
                .backgroundColor("222831")
                .color("EEEEEE")
                .build())
        .fileName("ex18")
        .creator("MR Excel")
        .build());
}

public static void ex19() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Collections.singletonList(
        SheetGenerator.builder()
                .applyHeaderOptionFunction(headers -> {
                        AtomicInteger index = new AtomicInteger(1);
                        headers.forEach(v -> {
                        v.setText("Text " + index.getAndIncrement());
                        });
                        return headers;
                })
                .applySheetOptionFunction(sheetBuilder -> {
                        sheetBuilder.conditionalFormatting(
                                Arrays.asList(
                                        ConditionalFormatting.eBuilder()
                                                .type(ConditionalFormattingOption.Type.DATABAR)
                                                .start("A2")
                                                .end("B6")
                                                .build(),
                                        ConditionalFormatting.eBuilder()
                                                .type(ConditionalFormattingOption.Type.COLOR_SCALE)
                                                .start("C2")
                                                .end("C6")
                                                .build(),
                                        ConditionalFormatting.eBuilder()
                                                .type(ConditionalFormattingOption.Type.TOP)
                                                .start("D2")
                                                .end("D6")
                                                .value(2)
                                                .build(),
                                        ConditionalFormatting.eBuilder()
                                                .type(ConditionalFormattingOption.Type.TOP)
                                                .start("E2")
                                                .end("E6")
                                                .styleId("ct")
                                                .value(2)
                                                .bottom(true)
                                                .build(),
                                        ConditionalFormatting.eBuilder()
                                                .type(ConditionalFormattingOption.Type.TOP)
                                                .operator(ConditionalFormattingOption.ConditionalFormattingTopOperation.ABOVE_AVERAGE)
                                                .start("F2")
                                                .end("F6")
                                                .value(1)
                                                .build(),
                                        ConditionalFormatting.eBuilder()
                                                .type(ConditionalFormattingOption.Type.TOP)
                                                .operator(ConditionalFormattingOption.ConditionalFormattingTopOperation.BELOW_AVERAGE)
                                                .start("G2")
                                                .end("G6")
                                                .styleId("ct")
                                                .value(1)
                                                .build()
                                )
                        );
                        return sheetBuilder.build();
                })
                .data(
                        ExampleData.ex19()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .activateConditionalFormatting(true)
        .addStyles("ct", StyleBody.builder()
                .type(StyleBody.StyleType.CONDITIONAL_FORMATTING)
                .backgroundColor("222831")
                .color("EEEEEE")
                .build())
        .fileName("ex19")
        .creator("MR Excel")
        .build());
}

public static void ex20() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Collections.singletonList(
        SheetGenerator.builder()
                .applyHeaderOptionFunction(headers -> {
                        headers.get(0).setConditionalFormatting(ConditionalFormatting.eBuilder()
                                .type(ConditionalFormattingOption.Type.DATABAR)
                                .build());
                        headers.get(1).setConditionalFormatting(ConditionalFormatting.eBuilder()
                                .type(ConditionalFormattingOption.Type.DATABAR)
                                .build());
                        headers.get(2).setConditionalFormatting(ConditionalFormatting.eBuilder()
                                .type(ConditionalFormattingOption.Type.COLOR_SCALE)
                                .build());
                        headers.get(3).setConditionalFormatting(ConditionalFormatting.eBuilder()
                                .type(ConditionalFormattingOption.Type.TOP)
                                .value(2)
                                .build());
                        headers.get(4).setConditionalFormatting(ConditionalFormatting.eBuilder()
                                .type(ConditionalFormattingOption.Type.TOP)
                                .styleId("ct")
                                .bottom(true)
                                .value(2)
                                .build());
                        headers.get(5).setConditionalFormatting(ConditionalFormatting.eBuilder()
                                .type(ConditionalFormattingOption.Type.TOP)
                                .operator(ConditionalFormattingOption.ConditionalFormattingTopOperation.ABOVE_AVERAGE)
                                .value(1)
                                .build());
                        headers.get(6).setConditionalFormatting(ConditionalFormatting.eBuilder()
                                .type(ConditionalFormattingOption.Type.TOP)
                                .operator(ConditionalFormattingOption.ConditionalFormattingTopOperation.BELOW_AVERAGE)
                                .styleId("ct")
                                .value(1)
                                .build());
                        AtomicInteger index = new AtomicInteger(1);
                        headers.forEach(v -> {
                        v.setText("Text " + index.getAndIncrement());
                        });
                        return headers;
                })
                .data(
                        ExampleData.ex19()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .activateConditionalFormatting(true)
        .addStyles("ct", StyleBody.builder()
                .type(StyleBody.StyleType.CONDITIONAL_FORMATTING)
                .backgroundColor("222831")
                .color("EEEEEE")
                .build())
        .fileName("ex20")
        .creator("MR Excel")
        .build());
}

public static void ex21() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
var data = ExcelTable.builder()
        .sheet(
                Arrays.asList(
                        Sheet.builder()
                                .headers(new ArrayList<>())
                                .data(new ArrayList<>())
                                .checkbox(
                                        Arrays.asList(
                                                Checkbox.builder()
                                                        .col(3)
                                                        .row(1)
                                                        .link("G3")
                                                        .text("Item 1")
                                                        .build(),

                                                Checkbox.builder()
                                                        .col(1)
                                                        .row(2)
                                                        .text("Item 2")
                                                        .threeD(true)
                                                        .build(),
                                                Checkbox.builder()
                                                        .col(5)
                                                        .row(1)
                                                        .checked(true)
                                                        .text("Item 3")
                                                        .build(),
                                                Checkbox.builder()
                                                        .col(1)
                                                        .row(12)
                                                        .text("Item 4")
                                                        .mixed(true)
                                                        .build()
                                        )
                                )
                                .build(),
                        Sheet.builder()
                                .headers(new ArrayList<>())
                                .data(new ArrayList<>())
                                .checkbox(
                                        Arrays.asList(
                                                Checkbox.builder()
                                                        .col(0)
                                                        .row(1)
                                                        .mixed(true)
                                                        .text("Item 1")
                                                        .build(),

                                                Checkbox.builder()
                                                        .col(1)
                                                        .row(2)
                                                        .text("Item 2")
                                                        .link("B3")
                                                        .build(),
                                                Checkbox.builder()
                                                        .col(5)
                                                        .row(1)
                                                        .threeD(true)
                                                        .text("Item 3")
                                                        .build(),
                                                Checkbox.builder()
                                                        .col(3)
                                                        .row(3)
                                                        .text("Item 4")
                                                        .checked(true)
                                                        .build()
                                        )
                                )
                                .build()
                )
        )
        .build();
data.applyExcelTableOption(
        ExcelTableOption.eBuilder()
                .fileName("ex21")
                .creator("MR Excel")
                .build()
);
GenerateExcel.generateExcel(data, "");
}

public static void ex22() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String c1 = "2C3639";
String c2 = "3F4E4F";
String c3 = "A27B5C";
String c4 = "DCD7C9";
String font = "Times New Roman";
final Map<String, FormulaMapBody> formulaMapBodyMap = new HashMap<>();
formulaMapBodyMap.put("J7", CustomFormulaSetting.builder()
        .styleId("customFormulaStyle")
        .formula("REPLACE(D3,1,1,\"replced\")")
        .build());
formulaMapBodyMap.put("H8", CustomFormulaSetting.builder()
        .styleId("customFormulaStyle")
        .formula("CONCATENATE(D2, \" \", D5)")
        .build());

formulaMapBodyMap.put("J2:J6", CustomFormulaSetting.builder()
        .formula("YEAR(NOW()-A2:A6)")
        .build());
formulaMapBodyMap.put("K2:K6", CustomFormulaSetting.builder()
        .formula("LOWER(D2:D6)")
        .build());

formulaMapBodyMap.put("I2", CustomFormulaSetting.builder()
        .styleId("customFormulaStyle")
        .formula("COUNT(A1:B8)")
        .build());
formulaMapBodyMap.put("H6", NoArgFormulaSetting.builder()
        .noArgType(NoArgFormulaSetting.NoArgFormulaType.NOW_HOUR)
        .styleId("customFormulaStyle")
        .build());
formulaMapBodyMap.put("H5", NoArgFormulaSetting.builder()
        .noArgType(NoArgFormulaSetting.NoArgFormulaType.NOW)
        .styleId("customFormulaStyle")
        .build());
formulaMapBodyMap.put("H4", SingleRefFormulaSetting.builder()
        .type(SingleRefFormulaSetting.SingleRefFormulaType.TRIM)
        .referenceCell("D3")
        .styleId("customFormulaStyle")
        .build());
formulaMapBodyMap.put("H7", SingleRefFormulaSetting.builder()
        .type(SingleRefFormulaSetting.SingleRefFormulaType.SUMIF)
        .referenceCell("B1:B5")
        .value("\">=5\"")
        .styleId("customFormulaStyle")
        .build());

formulaMapBodyMap.put("H3", SingleRefFormulaSetting.builder()
        .type(SingleRefFormulaSetting.SingleRefFormulaType.COUNTIF)
        .referenceCell("B1:B5")
        .value("\">=5\"")
        .styleId("customFormulaStyle")
        .build());
formulaMapBodyMap.put("H2", SingleRefFormulaSetting.builder()
        .type(SingleRefFormulaSetting.SingleRefFormulaType.ABS)
        .referenceCell("B5")
        .styleId("customFormulaStyle")
        .build());
formulaMapBodyMap.put("I1", SingleRefFormulaSetting.builder()
        .type(SingleRefFormulaSetting.SingleRefFormulaType.PROPER)
        .referenceCell("D1")
        .styleId("customFormulaStyle")
        .build());

formulaMapBodyMap.put("G11", SingleRefFormulaSetting.builder()
        .type(SingleRefFormulaSetting.SingleRefFormulaType.UPPER)
        .referenceCell("D1")
        .styleId("customFormulaStyle")
        .build());

formulaMapBodyMap.put("G10", SingleRefFormulaSetting.builder()
        .type(SingleRefFormulaSetting.SingleRefFormulaType.TAN)
        .referenceCell("B5")
        .styleId("customFormulaStyle")
        .build());

formulaMapBodyMap.put("G8", SingleRefFormulaSetting.builder()
        .type(SingleRefFormulaSetting.SingleRefFormulaType.COS)
        .referenceCell("B5")
        .styleId("customFormulaStyle")
        .build());

formulaMapBodyMap.put("G7", SingleRefFormulaSetting.builder()
        .type(SingleRefFormulaSetting.SingleRefFormulaType.FLOOR)
        .referenceCell("B5")
        .value(5)
        .styleId("customFormulaStyle")
        .build());

formulaMapBodyMap.put("G2", SingleRefFormulaSetting.builder()
        .type(SingleRefFormulaSetting.SingleRefFormulaType.POWER)
        .referenceCell("B2")
        .value(2)
        .styleId("customFormulaStyle")
        .build());

formulaMapBodyMap.put("G3", SingleRefFormulaSetting.builder()
        .type(SingleRefFormulaSetting.SingleRefFormulaType.MOD)
        .referenceCell("B3")
        .value(2)
        .styleId("customFormulaStyle")
        .build());

formulaMapBodyMap.put("G4", SingleRefFormulaSetting.builder()
        .type(SingleRefFormulaSetting.SingleRefFormulaType.SQRT)
        .referenceCell("B4")
        .styleId("customFormulaStyle")
        .build());

formulaMapBodyMap.put("G5", SingleRefFormulaSetting.builder()
        .type(SingleRefFormulaSetting.SingleRefFormulaType.CEILING)
        .referenceCell("B5")
        .value(5)
        .styleId("customFormulaStyle")
        .build());

formulaMapBodyMap.put("G6", SingleRefFormulaSetting.builder()
        .type(SingleRefFormulaSetting.SingleRefFormulaType.ROUND)
        .referenceCell("B5")
        .value(5)
        .styleId("customFormulaStyle")
        .build());

formulaMapBodyMap.put("G1", SingleRefFormulaSetting.builder()
        .type(SingleRefFormulaSetting.SingleRefFormulaType.LEN)
        .referenceCell("A1")
        .styleId("customFormulaStyle")
        .build());

formulaMapBodyMap.put("A8", FormulaSetting.builder()
        .type(FormulaSetting.FormulaType.SUM)
        .start("B2")
        .end("D3")
        .styleId("customFormulaStyle")
        .build());

formulaMapBodyMap.put("B8", FormulaSetting.builder()
        .type(FormulaSetting.FormulaType.AVERAGE)
        .start("A2")
        .end("F6")
        .styleId("customFormulaStyle")
        .build());
formulaMapBodyMap.put("C8", FormulaSetting.builder()
        .type(FormulaSetting.FormulaType.SUM)
        .start("A2")
        .end("F6")
        .styleId("customFormulaStyle")
        .build());
formulaMapBodyMap.put("D8", FormulaSetting.builder()
        .type(FormulaSetting.FormulaType.MAX)
        .start("A2")
        .end("F6")
        .styleId("customFormulaStyle")
        .build());
formulaMapBodyMap.put("E8", FormulaSetting.builder()
        .type(FormulaSetting.FormulaType.MIN)
        .start("A2")
        .end("F6")
        .styleId("customFormulaStyle")
        .build());
formulaMapBodyMap.put("F8", FormulaSetting.builder()
        .type(FormulaSetting.FormulaType.MIN)
        .start("A2")
        .end("F6")
        .styleId("customFormulaStyle")
        .build());
GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Collections.singletonList(
        SheetGenerator.builder()
                .applySheetOptionFunction(sheetBuilder -> sheetBuilder.formula(formulaMapBodyMap).build())
                .applyHeaderOptionFunction(headers -> {

                        AtomicInteger index = new AtomicInteger(1);
                        headers.forEach(v -> {
                        switch (index.get()) {
                                case 1:
                                v.setFormula(FormulaSetting.builder()
                                        .type(FormulaSetting.FormulaType.COUNT)
                                        .styleId("formulaStyle")
                                        .build());
                                break;
                                case 2:
                                v.setFormula(FormulaSetting.builder()
                                        .type(FormulaSetting.FormulaType.AVERAGE)
                                        .styleId("formulaStyle")
                                        .build());
                                break;
                                case 3:
                                v.setFormula(FormulaSetting.builder()
                                        .type(FormulaSetting.FormulaType.SUM)
                                        .styleId("formulaStyle")
                                        .build());
                                break;
                                case 4:
                                v.setFormula(FormulaSetting.builder()
                                        .type(FormulaSetting.FormulaType.MAX)
                                        .styleId("formulaStyle")
                                        .build());
                                break;
                                case 5:
                                v.setFormula(FormulaSetting.builder()
                                        .type(FormulaSetting.FormulaType.MIN)
                                        .styleId("formulaStyle")
                                        .build());
                                break;
                                case 6:
                                v.setFormula(FormulaSetting.builder()
                                        .type(FormulaSetting.FormulaType.COUNT)
                                        .styleId("formulaStyle")
                                        .build());
                                break;
                        }
                        v.setText("Text " + index.getAndIncrement());
                        });
                        return headers;
                })
                .data(
                        ExampleData.ex22()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .activateConditionalFormatting(true)
        .addStyles("headerStyle", StyleBody.builder()
                .backgroundColor("2B2E4A")
                .fontFamily(font)
                .color("E84545")
                .build())
        .addStyles("Date", StyleBody.builder()
                .format("short_date")
                .backgroundColor(c2)
                .fontFamily(font)
                .color(c4)
                .build())
        .addStyles(
                "customFormulaStyle",
                StyleBody.builder()
                        .size(15)
                        .fontFamily(font)
                        .color("2B2E4A")
                        .backgroundColor("E84545")
                        .border(StyleBody.BorderOption.builder()
                                .full(
                                        StyleBody.BorderOption.BorderBody.builder()
                                                .color("53354A")
                                                .style(StyleBody.BorderOption.BorderBody.BorderStyle.DASH_DOT)
                                                .build()
                                ).build())
                        .build()
        )
        .addStyles("formulaStyle", StyleBody.builder()
                .size(15)
                .fontFamily(font)
                .color("E84545")
                .backgroundColor("2B2E4A")
                .border(StyleBody.BorderOption.builder()
                        .full(
                                StyleBody.BorderOption.BorderBody.builder()
                                        .color("903749")
                                        .style(StyleBody.BorderOption.BorderBody.BorderStyle.MEDIUM)
                                        .build()
                        ).build())
                .build())
        .fileName("ex22")
        .creator("MR Excel")
        .build());
}

public static void ex24() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Collections.singletonList(
        SheetGenerator.builder()
                .applySheetOptionFunction(sheetBuilder -> sheetBuilder.asTable(new Sheet.AsTableOption()).build())

                .data(
                        ExampleData.ex2Sheet2()
                )
                .build()
), ExcelTableOption
        .eBuilder()

        .fileName("ex24")
        .creator("MR Excel")
        .build());
}

public static void ex25() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String c1 = "561C24";
String c2 = "6D2932";
String c3 = "C7B7A3";
String c4 = "E8D8C4";
String font = "Times New Roman";

StyleBody first = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c2).color(c1).fontFamily(font)
        .build();
StyleBody second = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c1)
        .color(c2)
        .fontFamily(font)
        .build();

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(
        Collections.singletonList(
                SheetGenerator.builder()
                        .applySheetOptionFunction(sheetBuilder -> {
                                sheetBuilder.pageOption(PageOption.builder()
                                        .margin(PageOption.Margin.builder()
                                                .top(.85)
                                                .left(.85)
                                                .right(.85).bottom(.85)
                                                .build())
                                        .footer(
                                                PageOption.HeaderFooterTypes.builder()
                                                        .even(PageOption.HeaderFooterTypes.HeaderFooterLocationMap.builder()
                                                                .c(
                                                                        PageOption.HeaderFooterTypes.HeaderFooterLocationMap.HeaderFooterOption.builder()
                                                                                .text("Even Footers center")
                                                                                .styleId("second")
                                                                                .build()
                                                                )
                                                                .build())
                                                        .build()
                                        )
                                        .header(PageOption.HeaderFooterTypes.builder()
                                                .first(PageOption.HeaderFooterTypes.HeaderFooterLocationMap.builder()
                                                        .c(
                                                                PageOption.HeaderFooterTypes.HeaderFooterLocationMap.HeaderFooterOption.builder()
                                                                        .text("First Header center")
                                                                        .styleId("first")
                                                                        .build()
                                                        ).l(
                                                                PageOption.HeaderFooterTypes.HeaderFooterLocationMap.HeaderFooterOption.builder()
                                                                        .text("First Header Left")
                                                                        .build()
                                                        )
                                                        .build())
                                                .odd(PageOption.HeaderFooterTypes.HeaderFooterLocationMap.builder()
                                                        .r(PageOption.HeaderFooterTypes.HeaderFooterLocationMap.HeaderFooterOption.builder()
                                                                .text("Odd Headers right")
                                                                .build())
                                                        .build())
                                                .build())

                                        .build());
                                return sheetBuilder.build();
                        })

                        .data(
                                ExampleData.ex2Sheet1()
                        )
                        .build()
        ), ExcelTableOption
                .eBuilder()
                .addStyles("first", first).addStyles(
                        "second", second
                )
                .fileName("ex25")
                .creator("MR Excel")
                .build());
}

public static void ex26() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String c1 = "561C24";
String c2 = "6D2932";
String c3 = "C7B7A3";
String c4 = "E8D8C4";
String font = "Times New Roman";

StyleBody first = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c2).color(c4).fontFamily(font)
        .build();
StyleBody second = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c4)
        .color(c2)
        .fontFamily(font)
        .build();

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(
        Collections.singletonList(
                SheetGenerator.builder()
                        .applyHeaderOptionFunction(headers -> {
                                headers.addAll(headers);
                                return headers;
                        })
                        .applySheetOptionFunction(sheetBuilder -> {
                                sheetBuilder.viewOption(ViewOption.builder()
                                        .type(ViewOption.Type.PAGE_BREAK_PREVIEW).
                                        hideGrid(true).hideHeadlines(true)
                                        .build());
                                return sheetBuilder.build();
                        })

                        .data(
                                ExampleData.ex2Sheet1()
                        )
                        .build()
        ), ExcelTableOption
                .eBuilder()
                .addStyles("first", first)
                .addStyles("second", second)
                .fileName("ex26")
                .creator("MR Excel")
                .build());
}

public static void ex27() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String c1 = "561C24";
String c2 = "6D2932";
String c3 = "C7B7A3";
String c4 = "E8D8C4";
String font = "Times New Roman";

StyleBody first = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c2).color(c4).fontFamily(font)
        .build();
StyleBody second = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c4)
        .color(c2)
        .fontFamily(font)
        .build();

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(
        Collections.singletonList(
                SheetGenerator.builder()
                        .applyHeaderOptionFunction(headers -> {
                                var arr = headers;
                                headers.addAll(arr);
                                return headers;
                        })
                        .applySheetOptionFunction(sheetBuilder -> {
                                sheetBuilder.viewOption(ViewOption.builder()
                                        .type(ViewOption.Type.PAGE_LAYOUT).hideRuler(true)
                                        .build());
                                return sheetBuilder.build();
                        })

                        .data(
                                ExampleData.ex2Sheet1()
                        )
                        .build()
        ), ExcelTableOption
                .eBuilder()
                .addStyles("first", first)
                .addStyles("second", second)
                .fileName("ex27")
                .creator("MR Excel")
                .build());
}

public static void ex28() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String c1 = "561C24";
String c2 = "6D2932";
String c3 = "C7B7A3";
String c4 = "E8D8C4";
String font = "Times New Roman";

StyleBody first = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c2).color(c4).fontFamily(font)
        .build();
StyleBody second = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c4)
        .color(c2)
        .fontFamily(font)
        .build();

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(
        Collections.singletonList(
                SheetGenerator.builder()
                        .applyHeaderOptionFunction(headers -> {
                                var arr = headers;
                                headers.addAll(arr);
                                return headers;
                        })
                        .applySheetOptionFunction(sheetBuilder -> {
                                sheetBuilder.viewOption(
                                        ViewOption.builder()
                                                .frozenOption(
                                                        ViewOption.FrozenOption.builder()
                                                                .type(ViewOption.FrozenOption.Type.ROW)
                                                                .index(1)
                                                                .build()
                                                )
                                                .build()
                                );
                                return sheetBuilder.build();
                        })

                        .data(
                                ExampleData.ex2Sheet1()
                        )
                        .build()
        ), ExcelTableOption
                .eBuilder()
                .addStyles("first", first)
                .addStyles("second", second)
                .fileName("ex28")
                .creator("MR Excel")
                .build());
}

public static void ex29() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String c1 = "561C24";
String c2 = "6D2932";
String c3 = "C7B7A3";
String c4 = "E8D8C4";
String font = "Times New Roman";

StyleBody first = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c2).color(c4).fontFamily(font)
        .build();
StyleBody second = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c4)
        .color(c2)
        .fontFamily(font)
        .build();

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(
        Collections.singletonList(
                SheetGenerator.builder()
                        .applyHeaderOptionFunction(headers -> {
                                var arr = headers;
                                headers.addAll(arr);
                                return headers;
                        })
                        .applySheetOptionFunction(sheetBuilder -> {
                                sheetBuilder.viewOption(
                                        ViewOption.builder()
                                                .frozenOption(
                                                        ViewOption.FrozenOption.builder()
                                                                .type(ViewOption.FrozenOption.Type.COLUMN)
                                                                .index(1)
                                                                .build()
                                                )
                                                .build()
                                );
                                return sheetBuilder.build();
                        })

                        .data(
                                ExampleData.ex2Sheet1()
                        )
                        .build()
        ), ExcelTableOption
                .eBuilder()
                .addStyles("first", first)
                .addStyles("second", second)
                .fileName("ex29")
                .creator("MR Excel")
                .build());
}

public static void ex30() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String c1 = "561C24";
String c2 = "6D2932";
String c3 = "C7B7A3";
String c4 = "E8D8C4";
String font = "Times New Roman";

StyleBody first = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c2).color(c4).fontFamily(font)
        .build();
StyleBody second = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c4)
        .color(c2)
        .fontFamily(font)
        .build();

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(
        Collections.singletonList(
                SheetGenerator.builder()
                        .applyHeaderOptionFunction(headers -> {
                                var arr = headers;
                                headers.addAll(arr);
                                return headers;
                        })
                        .applySheetOptionFunction(sheetBuilder -> {
                                sheetBuilder.viewOption(
                                        ViewOption.builder()
                                                .frozenOption(
                                                        ViewOption.FrozenOption.builder()
                                                                .type(ViewOption.FrozenOption.Type
                                                                        .BOTH)
                                                                .r(2)
                                                                .c(3)
                                                                .build()
                                                )
                                                .build()
                                );
                                return sheetBuilder.build();
                        })

                        .data(
                                ExampleData.ex2Sheet1()
                        )
                        .build()
        ), ExcelTableOption
                .eBuilder()
                .addStyles("first", first)
                .addStyles("second", second)
                .fileName("ex30")
                .creator("MR Excel")
                .build());
}

public static void ex31() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String c1 = "561C24";
String c2 = "6D2932";
String c3 = "C7B7A3";
String c4 = "E8D8C4";
String font = "Times New Roman";

StyleBody first = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c2).color(c4).fontFamily(font)
        .build();
StyleBody second = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c4)
        .color(c2)
        .fontFamily(font)
        .build();

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(
        Collections.singletonList(
                SheetGenerator.builder()
                        .applyHeaderOptionFunction(headers -> {
                                var arr = headers;
                                headers.addAll(arr);
                                return headers;
                        })
                        .applySheetOptionFunction(sheetBuilder -> {
                                sheetBuilder.viewOption(
                                        ViewOption.builder()
                                                .splitOption(
                                                        ViewOption.SplitOption.builder()
                                                                .type(ViewOption.SplitOption.Type.VERTICAL)
                                                                .split(10000)
                                                                .startAt(ViewOption.SplitOption.ViewStart.builder()
                                                                        .r("E1")
                                                                        .l("C1")
                                                                        .build())
                                                                .build()
                                                )
                                                .build()
                                );
                                return sheetBuilder.build();
                        })

                        .data(
                                ExampleData.ex2Sheet1()
                        )
                        .build()
        ), ExcelTableOption
                .eBuilder()
                .addStyles("first", first)
                .addStyles("second", second)
                .fileName("ex31")
                .creator("MR Excel")
                .build());
}

public static void ex32() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String c1 = "561C24";
String c2 = "6D2932";
String c3 = "C7B7A3";
String c4 = "E8D8C4";
String font = "Times New Roman";

StyleBody first = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c2).color(c4).fontFamily(font)
        .build();
StyleBody second = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c4)
        .color(c2)
        .fontFamily(font)
        .build();

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(
        Collections.singletonList(
                SheetGenerator.builder()
                        .applyHeaderOptionFunction(headers -> {
                                var arr = headers;
                                headers.addAll(arr);
                                return headers;
                        })
                        .applySheetOptionFunction(sheetBuilder -> {
                                sheetBuilder.viewOption(
                                        ViewOption.builder()
                                                .splitOption(
                                                        ViewOption.SplitOption.builder()
                                                                .type(ViewOption.SplitOption.Type.HORIZONTAL)
                                                                .split(5000)
                                                                .startAt(ViewOption.SplitOption.ViewStart.builder()
                                                                        .t("A15")
                                                                        .b("A7")
                                                                        .build())
                                                                .build()
                                                )
                                                .build()
                                );
                                return sheetBuilder.build();
                        })

                        .data(
                                ExampleData.ex2Sheet1()
                        )
                        .build()
        ), ExcelTableOption
                .eBuilder()
                .addStyles("first", first)
                .addStyles("second", second)
                .fileName("ex32")
                .creator("MR Excel")
                .build());
}

public static void ex33() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String c1 = "561C24";
String c2 = "6D2932";
String c3 = "C7B7A3";
String c4 = "E8D8C4";
String font = "Times New Roman";

StyleBody first = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c2).color(c4).fontFamily(font)
        .build();
StyleBody second = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c4)
        .color(c2)
        .fontFamily(font)
        .build();

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(
        Collections.singletonList(
                SheetGenerator.builder()
                        .applyHeaderOptionFunction(headers -> {
                                var arr = headers;
                                headers.addAll(arr);
                                headers.addAll(arr);
                                return headers;
                        })
                        .applySheetOptionFunction(sheetBuilder -> {
                                sheetBuilder.viewOption(
                                        ViewOption.builder()
                                                .splitOption(
                                                        ViewOption.SplitOption.builder()
                                                                .type(ViewOption.SplitOption.Type.BOTH)
                                                                .x(10000)
                                                                .y(5000)
                                                                .startAt(ViewOption.SplitOption.ViewStart.builder()
                                                                        .one("A15")
                                                                        .two("A7")
                                                                        .build())
                                                                .build()
                                                )
                                                .build()
                                );
                                return sheetBuilder.build();
                        })

                        .data(
                                ExampleData.ex2Sheet1()
                        )
                        .build()
        ), ExcelTableOption
                .eBuilder()
                .addStyles("first", first)
                .addStyles("second", second)
                .fileName("ex33")
                .creator("MR Excel")
                .build());
}

public static void ex34() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String c1 = "561C24";
String c2 = "6D2932";
String c3 = "C7B7A3";
String c4 = "E8D8C4";
String font = "Times New Roman";

StyleBody first = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c2).color(c4).fontFamily(font)
        .build();
StyleBody second = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c4)
        .color(c2)
        .fontFamily(font)
        .build();

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(
        Collections.singletonList(
                SheetGenerator.builder()
                        .applyHeaderOptionFunction(headers -> {
                                var arr = headers;
                                headers.addAll(arr);
                                headers.addAll(arr);
                                return headers;
                        })
                        .applySheetOptionFunction(sheetBuilder -> sheetBuilder.rtl(true).build())

                        .data(
                                ExampleData.ex2Sheet1()
                        )
                        .build()
        ), ExcelTableOption
                .eBuilder()
                .addStyles("first", first)
                .addStyles("second", second)
                .fileName("ex34")
                .creator("MR Excel")
                .build());
}

public static void ex35() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String c1 = "561C24";
String c2 = "6D2932";
String c3 = "C7B7A3";
String c4 = "E8D8C4";
String font = "Times New Roman";

StyleBody first = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c2).color(c4).fontFamily(font)
        .build();
StyleBody second = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c4)
        .color(c2)
        .fontFamily(font)
        .build();

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(
        Collections.singletonList(
                SheetGenerator.builder()
                        .applyHeaderOptionFunction(headers -> {
                                var arr = headers;
                                headers.addAll(arr);
                                headers.addAll(arr);
                                return headers;
                        })
                        .applySheetOptionFunction(sheetBuilder -> {
                                sheetBuilder.pageBreak(
                                        Sheet.PageBreak.builder()
                                                .column(Arrays.asList(
                                                        4, 6, 15, 17
                                                ))
                                                .row(Arrays.asList(1, 2, 3, 10))
                                                .build()
                                );
                                return sheetBuilder.build();
                        })

                        .data(
                                ExampleData.ex2Sheet1()
                        )
                        .build()
        ), ExcelTableOption
                .eBuilder()
                .addStyles("first", first)
                .addStyles("second", second)
                .fileName("ex35")
                .creator("MR Excel")
                .build());
}

public static void ex36() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
String c1 = "561C24";
String c2 = "6D2932";
String c3 = "C7B7A3";
String c4 = "E8D8C4";
String font = "Times New Roman";

StyleBody first = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c2).color(c4).fontFamily(font)
        .build();
StyleBody second = StyleBody.builder()
        .type(StyleBody.StyleType.HEADER_FOOTER)
        .backgroundColor(c4)
        .color(c2)
        .fontFamily(font)
        .build();

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(
        Collections.singletonList(
                SheetGenerator.builder()
                        .applySheetOptionFunction(sheetBuilder -> sheetBuilder.backgroundImage(ImageInput.builder()
                                .data("https://raw.githubusercontent.com/mohammadrezaeicode/mr-excel-example-gallery/refs/heads/main/CDN/generateExcel/im.png")
                                .type(ImageInput.AcceptType.URL)
                                .build()).build())
                        .data(
                                ExampleData.ex2Sheet1()
                        )
                        .build()
        ), ExcelTableOption
                .eBuilder()
                .fetch(Utils.fetch())
                .addStyles("first", first)
                .addStyles("second", second)
                .fileName("ex36")
                .creator("MR Excel")
                .build());
}

public static void ex39() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {

GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Collections.singletonList(
        SheetGenerator.builder()
                .applySheetOptionFunction(
                        sheetBuilder -> {
                                sheetBuilder.checkbox(Arrays.asList(
                                        Checkbox.builder()
                                                .col(13)
                                                .row(1)
                                                .text("Checkbox 1")
                                                .link("G1")
                                                .build(),
                                        Checkbox.builder()
                                                .col(14)
                                                .row(2)
                                                .text("Checkbox 3")
                                                .link("G2")
                                                .build(),
                                        Checkbox.builder()
                                                .col(15)
                                                .row(3)
                                                .text("Checkbox 4")
                                                .link("G3")
                                                .build()
                                ));
                                sheetBuilder.images(Arrays.asList(
                                        Sheet.ImageTypes.builder()
                                                .image(ImageInput.builder().data(
                                                                "https://mohammadrezaeicode.github.io/mr-excel-page/img/ezgif.com-gif-maker.gif"
                                                        ).build()
                                                ).from("H1").type(Sheet.ImageTypes.Type.ONE)
                                                .build(),
                                        Sheet.ImageTypes.builder()
                                                .image(ImageInput.builder().data(
                                                                "https://mohammadrezaeicode.github.io/mr-excel-page/img/uniqe.jpg"
                                                        ).build()
                                                ).from("H2").type(Sheet.ImageTypes.Type.ONE)
                                                .build(),
                                        Sheet.ImageTypes.builder()
                                                .image(ImageInput.builder().data(
                                                                "https://mohammadrezaeicode.github.io/mr-excel-page/img/ex.PNG"
                                                        ).build()
                                                ).from("H3").type(Sheet.ImageTypes.Type.ONE)
                                                .build(),
                                        Sheet.ImageTypes.builder()
                                                .image(ImageInput.builder().data(
                                                                "https://mohammadrezaeicode.github.io/mr-excel-page/img/ex.PNG"
                                                        ).build()
                                                ).from("H4").type(Sheet.ImageTypes.Type.ONE)
                                                .build()
                                        ,
                                        Sheet.ImageTypes.builder()
                                                .image(ImageInput.builder().data(
                                                                "https://mohammadrezaeicode.github.io/mr-excel-page/img/ezgif.com-gif-maker.gif"
                                                        ).build()
                                                ).from("E1")
                                                .to("F10")
                                                .type(Sheet.ImageTypes.Type.TWO)
                                                .build(),
                                        Sheet.ImageTypes.builder()
                                                .image(ImageInput.builder().data(
                                                                "https://mohammadrezaeicode.github.io/mr-excel-page/img/uniqe.jpg"
                                                        ).build()
                                                ).from("H6").type(Sheet.ImageTypes.Type.TWO)
                                                .build()
                                ));
                                return sheetBuilder.build();
                        }
                )
                .data(
                        ExampleData.ex17()
                )
                .build()
), ExcelTableOption
        .eBuilder()
        .fetch(Utils.fetch())
        .fileName("ex39")
        .creator("MR Excel")
        .build());
}

public static void ex40() throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
GenerateExcel.generateHeaderAndGenerateExcelWithMultiSheet(Collections.singletonList(
        SheetGenerator.builder()
                .applySheetOptionFunction(sheetBuilder -> {
                        return sheetBuilder.dropDowns(
                                Collections.singletonList(
                                        DropDown.builder()
                                                .forCell(ExampleData.dropdownLocation())
                                                .option(ExampleData.dropdownItems())
                                                .build()
                                )
                        ).build();
                })
                .data(ExampleData.ex40())

                .build()
), ExcelTableOption.eBuilder().fileName("ex40").build());

}

```
