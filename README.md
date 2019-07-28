# OOPReports

**为 Visual Foxpro 9.0 开发者提供的对象化报表**

项目管理者： [Doug Hennig](mailto:dhennig@stonefield.com)

翻译：xinjie   2019.07.27

## 简介

虽然 VFP 是一种很好的面向对象的开发语言，但是其报表却从未被对象化。事实上，报表不能在运行时动态更改，也不能程序化的生成。

从 FoxPro 2.0 开始，FRX 文件就是一个表，所以看起来这应该是一件容易的事。 然而，事实是 FRX 结构是丑陋的。尽管对它的结构有描述（您可以在VFP主目录的 Tools \ FileSpec 子目录中运行 90FRX 报表以打印 FRX 文件结构），但是不同的报表对象类型使用表结构中的字段过程中，还是有细微的差别，而且这种差别并没有被详细说明。此外，有很多字段，可以使用一些非常长的 INSERT INTO 语句。 此外，大小和位置值的单位是1/10000英寸（称为FoxPro报表单位，或FRU），并且有一些奇怪的“模糊”因素使你必须应用这些数字，如报表设计器中的带区高度。 所以我创建了一组类来以编程方式处理它。

报表对象类全部在 SFRepObj.vcx 中，并且都基于 SFReportBase 类，它是 Custom 的子类。 每个类都使用 SFRepObj.h 作为其包含文件; 像 cnBAND_HEIGHT 这样的常量比像2083.333这样的硬编码值更容易理解。

## SFReportFile

在使用时，实际上你只需要直接实例化的唯一的一个类就是 SFReportFile。 它表示报表文件，并具有报表设计器界面属性和方法的程序化接口。 它会在必要时（例如，当您向报表添加字段时）从其他类实例化对象。

SFReportFile 可以使用的属性：

| 属性 | 描述                        |
|----------|--------------------------------|
| cDevice  | 要使用的打印机的名称 |
| cFontName	| 报表的默认字体（默认值为“Courier New”） |
| cMemberData	| 报表标题记录的成员数据 |
| cReportFile	| 要创建的报表文件的名称 |
| cUnits	| 计量单位：“C”（字符），“I”（英寸）或“M”（厘米）（默认值为“C”;常量在 SFRepObj.h 中为这些值定义） |
| lAdjustObjectWidths  	| .T. (默认值) 确保没有任何对象比纸张宽度宽 |
| lNoDataEnvironment   	| .T. (默认值) 防止编辑 DataEnvironment |
| lNoPreview	| .T. 防止在报表设计器中预览或打印报表 |
| lNoQuickReport	| .T. (默认值) 阻止访问“快速报表”功能 |
| lPrintColumns	| .T. 从左至右打印, .F. 从上至下打印 |
| lPrivateDataSession	| .T. 表示报表使用私有数据工作期 |
| lSummaryBand	| .T. 表示报表具有概要带区 |
| lTitleBand	| .T. 表示报表具有标题带区 |
| lWholePage	| .T. 表示打印区域为页面, .F. 表示打印区域为整页 |
| nColumns	| 报表的列数(默认值为 1) |
| nColumnSpacing	| 各列之间的间距 |
| nDefaultSource	| 报表默认的纸张来源(默认值为 -1) |
| nDetailBands	| 报表中细节带区的数量 |
nFontSize	| 报表的默认字号(默认值为 10) |
| nFontStyle	| 报表的默认字体样式 |
| nGroups	| 报表中的分组数 |
| nLeftMargin	| 报表的左边距 |
| nMinPaperWidth	| 纸张的最小宽度 |
| nOrientation	| 报表的方向： 0 = 如果报表太宽而无法纵向显示，则自动设置为横向, 1 = 纵向, 2 = 横向 |
| nPaperLength	| 纸张长度 |
| nPaperSize	| 纸张大小（默认值为PRTINFO（2）;请参阅该函数的帮助以获取值列表） |
| nPaperWidth	| 纸张宽度 |
| nRepWidth	| 计算的报表宽度 |
| nRulerScale	| 用于标尺的比例：1 =英寸，2 =公制，3 =像素 |
| nWidth	| 报表宽度 |

除了 cReportFile 和 cUnits 之外，这些属性仅代表您在报表设计器中看到的报表的相同属性选项。 每个属性都有一个 Assign 方法，可以防止存储错误数据类型或范围的值。 cUnits 需要一点解释：有时，用有多少个字符和行比使用报表计量单位（例如对象的水平和垂直位置）更容易。 例如，如果您有一个30个字符的字段，则可以更轻松地将报表上的宽度指定为 30 个字符，而不是计算它应该有多少英寸。 我们将查看的所有对象中的所有位置和大小属性都应以 cUnits 中定义的单位表示。

下面是实例化 SFReportFile 对象的示例代码，指定要创建的报表的名称，指定报告具有摘要带区，并设置报表的默认字体。

```foxpro
loReport = newobject('SFReportFile', 'SFRepObj.vcx')
loReport.cReportFile  = 'CustomerReport.frx'
loReport.lSummaryBand = .T.
loReport.cFontName    = 'Arial'
```

SFReportFile 可用的方法：

| 方法 | 描述                    |
|--------|--------------------------------|
| CreateDetailBand	| 创建一个细节带区 |
| CreateGroupBand	| 创建一个新的分组带区 |
| CreateVariable	| 创建报表变量 |
| GetHFactor	| 计算指定字体的水平因子 |
| GetPaperWidth	| 获取纸张宽度的最大值 |
| GetRelativeVPosition	| 获取指定对象相对于其带区起点的垂直位置 |
| GetReportBand	| 返回指定带区的对象引用 |
| GetVariable	| 返回一个变量对象 |
| GetVFactor	| 计算指定字体的垂直因子 |
| Load	| 将指定的 FRX 加载到报表对象中 |
| LoadFromCursor	| 将当前工作区中的 FRX 加载到报表对象中 |
| Save	| 创建报表文件 |

在探索其他对象时，我们将研究其中的一些方法。

## 报表带区
SFReportFile 通过将 SFReportBand 对象实例化为受保护属性，为报表中的每个带区创建一个对象。 例如，Init方法使用以下代码自动创建页眉，细节和页脚带区对象，因为每个报表至少包含它们三个带区：

```foxpro
with This
    .oPageHeaderBand = newobject('SFReportBand', .ClassLibrary, '', ;
        'Page Header', This)
    .oDetailBand     = newobject('SFReportBand', .ClassLibrary, '', ;
        'Detail', This)
    .oPageFooterBand = newobject('SFReportBand', .ClassLibrary, '', ;
        'Page Footer', This)
endwith
```

SFReportBand 的 Init 方法接受两个参数：带区类型（它存储于 cBandType 属性）和 SFReportFile 对象的对象引用（它存储于 oReport 属性），因此它可以在必要时回调一些 SFReportFile 方法。

除了三个默认带区外，您还可以通过三种方式创建其他带区。 要指定报告具有标题或概要带区，请将 lTitleBand 或 lSummaryBand 属性设置为 .T.。 要创建组标头和组注脚带区，请调用 CreateGroupBand 方法。分组按照定义的顺序自动编号; 可能的增强将允许组重新排序。要创建其他的细节带区，请调用 CreateDetailBand 方法，如果你想要细节页眉和页脚带区，请一并传递 .T.。 与分组一样，细节带区按照它们定义的顺序自动编号。

GetReportBand 方法返回指定带区的对象引用。 对于组标头或组注脚带区，您还可以指定要为其添加的组编号。下面是一些示例代码，用于设置页标头和细节带区的高度。

```foxpro
loPageHeader = loReport.GetReportBand('Page Header')
loPageHeader.nHeight = 8
loDetail = loReport.GetReportBand('Detail')
loDetail.nHeight = 1
```

顺便说一句，你不必像我在这个例子中那样硬编码带区名。 所有带区名称都已定义为 SFRepObj.h 中的常量，它是 SFRepObj.vcx 中所有类的包含文件。 例如，ccBAND_PAGE_HEADER 被定义为 “PAGE HEADER”，因此您可以使用以下代码：

```foxpro
loPageHeader = loReport.GetReportBand(ccBAND_PAGE_HEADER)
```

SFReportBand 可用的属性：

| 属性 | 描述                        |
|----------|--------------------------------|
| cOnEntry	| 在进入带区时需要运行的表达式 |
| cOnExit	| 在退出带区时需要运行的表达式 |
| cTargetAlias	| 细节带区的目标别名表达式 |
| lAdjustBandHeight	| .T. (默认值) 表示应调整带区的高度以适应其中的对象 |
| lConstantHeight	| .T. 标识带区具有恒定的高度 |
| lDeleteObjectsOutsideBand	| .T. (默认值) 表示删除超过带区高度的对象; 仅当 lAdjustBandHeight = .F. 时有效 |
| lPageFooter	| 对于概要带区并且页注脚需要打印，属性值则为 .T. |
| lPageHeader	| 对于概要带区并且页标头需要打印，属性值则为 .T. |
| lStartOnNewPage		| 如果带区在新的一页上打印，属性值则为.T. |
| nCount	| 带区中项目的数量 |
| nHeight	| 带区高度 |
| nNewPageWhenLessThan	| Starts a group on a new page when there is less than this much space left on the current page |

As with SFReportFile, they simply expose band options available in the Report Designer as properties. nHeight is expressed in the units defined in the cUnits property of the SFReportFile object; for example, if cUnits is "C," setting nHeight to 8 defines an 8-line high band. One nice feature: if the objects in a band extend below the defined height of the band, saving the report  automatically adjusts the band height if the lAdjustBandHeight property is .T. (the nHeight property isn’t changed, but the band’s record in the FRX is).

SFReportBand just has three public methods that you'll use: Add, GetReportObjects, and Remove (it has a few other public methods but they're called from methods in SFReportFile). Add is probably the method you'll use the most in any of the classes; it adds an object to a band, so you'll call it once for every object in the report. Pass Add the object type you want to add: "field,"" "text," "image," "line," or "box" (these values are defined in SFRepObj.h, so you can use constants rather than hard-coded values), and it returns a reference to the newly created object. GetReportObjects populates the specified array with object references to the objects added to the band. You can optionally pass it a filter condition to only get certain items; typically, the filter would check for a property of the objects being a certain value. You may not use this method yourself, but SFReportFile uses it to output a band and all of its objects to the FRX file. Remove removes the specified object from the band; it isn’t used often since you likely wouldn’t add an object only to remove it later.

SFReportGroup is a subclass of SFReportBand that's specific for group header and footer bands. Its public properties are:

| Property | Purpose                        |
|----------|--------------------------------|
| cExpression	| The group expression |
| lPrintOnEachPage| 	.T. to print the group header on each page |
| lResetPage	| .T. to reset the page number for each page to 1 |
| lStartInNewColumn	| .T. if each group should start in a new column |
| nNewPageWhenLessThan	| Starts a group on a new page when there is less than this much space left on the current page |

The most important one is obviously cExpression, since this determines what constitutes a group.

Here's some code that creates a group, gets a reference to the group header band, and sets the group expression, height, and printing properties.

```foxpro
loReport.CreateGroupBand()
loGroup = loReport.GetReportBand('Group Header', 1)
loGroup.cExpression          = 'CUSTOMER.COUNTRY'
loGroup.nHeight              = 3
loGroup.lPrintOnEachPage     = .T.
loGroup.nNewPageWhenLessThan = 4
```

## Fields and text
The most common thing you'll add to a report band is a field (actually, an expression which could be a field name but can be any valid FoxPro expression), since the whole purpose of a report is to output data. SFReportField is the class used for fields. However, before we talk about SFReportField, let's look at its parent classes.

SFReportRecord is an ancestor class for every class in SFRepObj.vcx except SFReportFile and SFReportBase. It has a few properties that all objects share:

* Recno: the record number of the object in the report.
* cComment: the comment for the report record.
* cMemberData: the member data for the report record.
* cUniqueID: the unique ID for the report record (normally left blank and auto-assigned).
* cUser: the user-defined property.
* nObjectType: contains the OBJTYPE value for the FRX record of the object (for example, fields have an OBJTYPE of 8).

It also has two methods: CreateRecord and ReadFromFRX. CreateRecord is called from the SFReportFile object when it creates a report. SFReportFile creates an object from a record in the FRX using SCATTER NAME loRecord BLANK MEMO to create an object with one property for each field in the record and then passes that object to the CreateRecord method of the SFReportRecord object, which fills in properties of the report record object from values in its own properties. SFReportRecord is an abstract class; it isn't used directly, but is the parent class for other classes. Its CreateRecord method simply ensures a valid report record object was passed and sets the ObjType property of this object (which is written to the OBJTYPE field in the FRX) to its nObjectType property. ReadFromFRX sort of does the opposite: it assign its properties the values from the current record in an open FRX table.

SFReportObject is a subclass of SFReportRecord that's used for report objects (here, I mean "object" in the "thingy" sense, such as a field, rather than "what you get when you instantiate a class" sense). It has these public properties, which represent the minimum set of options for a report object.

| Property | Purpose                        |
|----------|--------------------------------|
| cAlignment	| The alignment for the object: "left", "center", or "right" (constants are defined for these values in SFRepObj.h) |
| cName	| A name for the object (used by SFReportBand.Item to locate an item by name) |
| cPrintWhen	| The Print When expression |
| lAutoCenter	| .T. (the default) to automatically center this object vertically in a row when using character units for the report |
| lPrintInFirstWholeBand	| .T. (the default) to print in the first whole band of a new page |
| lPrintOnNewPage	| .T. to print when the detail band overflows to a new page |
| lPrintRepeats	| .T. (the default) to print repeated values |
| lRemoveLineIfBlank	| .T. to remove a line if there are no objects on it |
| lStretch	| .T. if the object can stretch |
| lTransparent	| .T. (the default) if the object is transparent, .F. for opaque |
| nBackColor	| The object's background color; use an RGB() value (-1 = default) |
| nFloat	| 0 if the object should float in its band, 1 (the default) if it should be positioned relative to the top of the band, or 2 if it should be relative to the bottom of the band (constants are defined for these values in SFRepObj.h) |
| nForeColor	| The object's foreground color; use an RGB() value (-1 = default) |
| nGroup	| Non-zero if this object is grouped with other objects |
| nHeight	| The height of the object |
| nHPosition	| The horizontal position for the object
| nPrintOnGroupChange	| The group number if this object should print on a group change |
| nVPosition	| The vertical position for the object relative to the top of the band |
| nWidth	| The width of the object |

As with other classes, these properties simply expose options available in the Report Designer as properties.

The CreateRecord method first uses DODEFAULT() to execute the behavior of SFReportRecord, then it has some data conversion to do. For example, object colors are stored in the PENRED, PENGREEN, and PENBLUE fields in the FRX record, but we want to have a single nForeColor property that we set (for example, to red using RGB(255, 0, 0)) like we do with VFP controls. Other properties are similar; for example, the value in nFloat updates the FLOAT, TOP, and BOTTOM fields in the FRX.

Finally, we're back to SFReportField, the subclass of SFReportObject that holds information about fields in a report. This class adds the following properties to those of SFReportObject.

| Property | Purpose                        |
|----------|--------------------------------|
| cCaption	| The design-time caption for the field |
| cDataType	| The data type of the expression: "N" for numeric, "D" for date, and "C" for everything else (only required if you'll edit the report in the Report Designer later) |
| cExpression	| The expression to display |
| cFontName	| The font to use (if blank, which it is by default, SFReportFile.cFontName is used) |
| cPicture	| The picture (format and inputmask) for the field |
| cTotalType	| The total type: "N" for none, "C" for count, "S" for sum, "A" for average, "L" for lowest, "H" for highest, "D" for standard deviation, and "V" for variance (constants are defined for these values in SFRepObj.h) |
| lFontBold	| .T. if the object should be bolded |
| lFontItalic	| .T. if the object should be in italics |
| lFontUnderline	| .T. if the object should be underlined |
| lResetOnPage	| .T. to reset the variable at the end of each page; .F. to reset at the end of the report |
| nDataTrimming	| Specifies how the Trim Mode for Character Expressions is set |
| nFontCharSet	| The font charset to use |
| nFontSize	| The font size to use (if 0, which it is by default, SFReportFile.nFontSize is used) |
| nResetOnDetail	| The detail band number to reset the value on |
| nResetOnGroup	| The group number to reset the value on |

As you can see, you have the same control over the properties of an object in a report as you do in the Report Designer.

As with SFReportObject, SFReportField's CreateRecord method uses DODEFAULT() to get the behavior of SFReportRecord and SFReportObject, then it does some data conversion similar to what SFReportObject does (for example, lFontBold, lFontItalic, and lFontUnderline are combined into a single FONTSTYLE value).
SFReportText is a subclass of SFReportField, since it has the same properties but only slightly different behavior. It automatically adds quotes around the expression since text objects always contain literal strings rather than expressions. It also sets the PICTURE field in the FRX to match the alignment of the data (because that's how alignment is handled for text objects), and sizes the object appropriately for the size of the text (in other words, it acts like setting the AutoSize property of a Label control to .T.).

Here's some code that adds text and field objects to the detail band and sets their properties. This code uses characters as the units, so values are in characters or lines.

```foxpro
loObject = loDetail.Add('Text')
loObject.cExpression = 'Country:'
loObject.nVPosition  = 1
loObject.lFontBold   = .T.
loObject = loDetail.Add('Field')
loObject.cExpression = 'CUSTOMER.COUNTRY'
loObject.nWidth      = fsize('COUNTRY', 'CUSTOMER')
loObject.nVPosition  = 1
loObject.nHPosition  = 10
loObject.lFontBold   = .T.
```

## Lines, boxes, and images
SFReportShape is a subclass of SFReportObject that defines the properties for lines and boxes (it isn;t used directly but is subclassed). nPenPattern is the pen pattern for the object: 0 = none, 1 = dotted, 2 = dashed, 3 = dash-dot, 4 = dash-dot-dot, and 8 = normal. nPenSize is the pen size for the line: 0, 1, 2, 4, or 6.

SFReportLine is a subclass of SFReportShape that's used for line objects. It adds one property, lVertical, that you should set to .T. to create a vertical line or .F. (the default) for a horizontal one. Its CreateRecord method sets the height for a horizontal line or the width for a vertical one to the appropriate value based on the pen size.

The following code adds a heavy blue line on line 4 of the page header band:

```foxpro
loObject = loPageHeader.Add('Line')
loObject.nWidth     = lnWidth
loObject.nVPosition = 4
loObject.nHPosition = 0
loObject.nPenSize   = 6
loObject.nForeColor = rgb(0, 0, 255)
```

Boxes use SFReportBox, which is also a subclass of SFReportShape. It adds nCurvature (the curvature of the box corners; the default is 0, meaning no curvature), lStretchToTallest (.T. to stretch the object relative to the tallest object in the band), and nFillPattern (the fill pattern for the object: 0 = none, 1 = solid, 2 = horizontal lines, 3 = vertical lines, 4 = diagonal lines, leaning left, 5 = diagonal lines, leaning right, 6 = grid, 7 = hatch) properties.

SFReportImage, a subclass of SFReportObject, is used for images. Set the cImageSource property to the name of the image file or General field that's the source of the image and nImageSource to 0 if the image comes from a file, 1 if it comes from a General field, or 2 if it's an expression. nStretch defines how to scale the image: 0 = clip, 1 = isometric, and 2 = stretch (the same values used by the Stretch property of an Image control). Its CreateRecord method automatically puts quotes around the image source if a file is used, and sets the appropriate field in the FRX record if the cAlignment property is set to "Center" (only applicable for General fields).

## Report variables
Report variables are defined using the CreateVariable method of the SFReportFile object. This method returns an object reference to the SFReportVariable object it created so you can set properties of the variable. The public properties for variables are:

| Property | Purpose                        |
|----------|--------------------------------|
| cInitialValue	| The initial value	|
| cName	| The variable name	|
| cTotalType	| The total type: "N" for none, "C" for count, "S" for sum, "A" for average, "L" for lowest, "H" for highest, "D" for standard deviation, and "V" for variance (constants are defined for these values in SFRepObj.h)	|
| cValue	| The value to store	|
| lReleaseAtEnd	| .T. to release the variable at the end of the report	|
| lResetOnPage	| .T. to reset the variable at the end of each page; .F. to reset at the end of the report	|
| nResetOnGroup	| The group number to reset the variable on	|

The following code creates a report variable called lnCount and specifies that it should start at 0 and increment by 1 for each record printed in the report. This variable is then printed in the summary band of the report, showing the number of records printed.

```foxpro
loVariable = loReport.CreateVariable()
loVariable.cName         = 'lnCount'
loVariable.cValue        = 1
loVariable.cInitialValue = 0
loVariable.cTotalType    = 'Sum'
loSummary = loReport.GetReportBand('Summary')
loObject  = loSummary.Add('Field')
loObject.cExpression = 'ltrim(str(lnCount)) + ' + ;
  '" record" + iif(lnCount = 1, "", "s") + " printed"'
loObject.nWidth      = 21
loObject.nVPosition  = 2
loObject.nHPosition  = 0
loObject.lFontBold   = .T.
```

## Examples
Customers.prg and Employees.prg are sample programs that create reports for the Customer and Employee tables in the VFP Testdata database. Customers.prg creates (and previews) CustomerReport.frx, which shows customers grouped by country, with the maximum order amount subtotaled by country and totaled at the end of the report. Employees.prg creates EmployeeReport.frx, which shows the name and photo of each employee. These reports aren't intended to be realistic; they just show off various features of the report classes described in this article, including printing images and lines, setting font sizes and object colors, positioning objects in different bands, use of report variables and group bands, etc.

Craig Boyd's [GridExtras](https://tinyurl.com/ycmqo8fg): sorting, incremental searching, and filtering on each column, column selection, and output to Excel and Print Preview. The Print Preview feature uses SFReportFile to dynamically create a report based on the current grid layout.

Altering existing reports is also much easier to do using these classes. Suppose you have a report that contains some sensitive information. Some staff shouldn't see that information, so you use the Print When expression for those fields to not output them unless the staff have the correct permissions. For example, the sample Employees report shows a full employee listing, but Birth Date and Home Phone should only be visible to Human Resources (HR) staff.
 
The Print When expression for those two fields and their column headers is "plHR." The variable plHR is .T. for HR staff and .F. otherwise. When the report is run by non-HR staf, BirthDate and HomePhone don't appear but leave obvious holes in the report layout. What would be nice is if the other columns could move left to take up the empty space.
 
HackReport.prg shows how to do this. The code isn't complicated: instantiate SFReportFile, load the FRX, remove objects in the Page Header and Detail bands that don't appear because of their Print When expressions, and move the rest of the fields the appropriate distance to the left.

```foxpro
loReport = newobject('SFReportFile', 'SFRepObj.vcx')
loReport.Load('Employees.frx')

* Process objects in the page header and detail bands.

loBand = loReport.GetReportBand('Page Header')
ProcessObjects(loBand)
loBand = loReport.GetReportBand('Detail')
ProcessObjects(loBand)

* Save the updated report and run it.

loReport.cReportFile = addbs(sys(2023)) + 'EmployeeReport.frx'
loReport.Save()
report form (loReport.cReportFile) preview

* Specifically release the report object so proper object cleanup occurs.

loReport.Release()

* Remove any objects that fail their PrintWhen expression and move objects to
* the right of those objects on the same line to the left.

function ProcessObjects(toBand)
local laObjects[1], ;
	lnObjects, ;
	lnAdjust, ;
	lnI, ;
	loObject, ;
	lcExpr, ;
	lnVPos, ;
	llVisible
lnObjects = toBand.GetReportObjects(@laObjects)
lnAdjust  = 0
for lnI = 1 to lnObjects
	loObject = laObjects[lnI]
	lcExpr   = loObject.cPrintWhen
	if lnAdjust <> 0 and loObject.nVPosition = lnVPos
		loObject.nHPosition = loObject.nHPosition - lnAdjust
	endif lnAdjust <> 0 ...
	llVisible = .T.
	if not empty(lcExpr)
		try
			llVisible = evaluate(lcExpr)
		catch
		endtry
	endif not empty(lcExpr)
	if not llVisible
		if lnI < lnObjects
			lnAdjust = laObjects[lnI + 1].nHPosition - ;
				loObject.nHPosition
		endif lnI < lnObjects
		toBand.Remove(loObject.cUniqueID)
		lnVPos = loObject.nVPosition
	endif not llVisible
next lnI
```

## Conclusion
Although you might write a fair bit of code to create an FRX using the report object classes, the code is simple: create some objects and set their properties. It sure beats writing 50 INSERT INTO statement with 75 fields to fill for each. Please report (pun intended) to me any suggestions you have for improvements.
