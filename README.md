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
| nMinPaperWidth	| 纸张的最小宽度(译者注：单位为FRU) |
| nOrientation	| 报表的方向： 0 = 如果报表太宽而无法纵向显示，则自动设置为横向, 1 = 纵向, 2 = 横向 |
| nPaperLength	| 纸张长度 (译者注：单位为FRU) |
| nPaperSize	| 纸张大小（默认值为PRTINFO（2）;请参阅该函数的帮助以获取值列表） |
| nPaperWidth	| 纸张宽度 (译者注：单位为FRU) |
| nRepWidth	| 计算的报表宽度 |
| nRulerScale	| 用于标尺的比例：1 =英寸，2 =公制，3 =像素 |
| nWidth	| 报表宽度 (译者注：单位为FRU) |

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
| nNewPageWhenLessThan	| 当前页面上剩余空间少于此空间时,在新页面上开始新组 |

与 SFReportFile 一样，它们只是将报表设计器中可用的带区选项表示为属性。 nHeight 以 SFReportFile 对象的 cUnits 属性中定义的单位表示; 例如，如果 cUnits为 “C”，则将 nHeight 设置为 8 ，表示带区高度为 8 个字符。一个很好的特性：如果带区中的对象高度大于带区高度，如果 lAdjustBandHeight 属性为 .T，则保存报表时会自动调整带高。（nHeight 属性值没有改变，改变的是带区在 FRX 表中的记录）。

你仅仅可以使用 SFReportBand 中的三个公开的方法：Add，GetReportObjects 和 Remove（其他的公共方法，供 SFReportFile 中的方法调用）。Add 可能是您在任何类中使用最多的方法；它将一个对象添加到一个带区，因此你必须为报表中的每个对象都调用它一次。传递需要添加的对象类型：“field”，“text”，“image”，“line” 或 “box”（这些值在 SFRepObj.h 中定义，因此您可以使用常量而不是硬编码值），它返回对新创建的对象的引用。GetReportObjects 使用对添加到带区的对象的对象引用填充指定的数组。您可以选择将过滤条件作为参数传递以便仅获取某些项;通常，过滤器将检查对象的属性是否为某个值。您可能不会自己使用此方法，但 SFReportFile 使用它将带区及其所有对象输出到 FRX 文件。Remove 从带中删除指定的对象;它不是经常使用，因为你可能不会添加一个对象仅仅是为了删除它。

SFReportGroup 是 SFReportBand 的子类，它用于组标头和组注脚带区。其属性如下：

| 属性 | 描述                        |
|----------|--------------------------------|
| cExpression	| 分组表达式 |
| lPrintOnEachPage| 	.T. 表示在每页都打印组标头 |
| lResetPage	| .T. 表示每页都重置页号为 1 |
| lStartInNewColumn	| .T. 表示每个组都应该在新列中开始 |
| nNewPageWhenLessThan	| 当前页面上剩余空间少于此空间时,在新页面上开始新组 |

最重要的一个属性显然是 cExpression，因为它决定了如何分组。

下面是一些创建组的代码，获取对组标头带区的引用，并设置组表达式，高度和打印属性。

```foxpro
loReport.CreateGroupBand()
loGroup = loReport.GetReportBand('Group Header', 1)
loGroup.cExpression          = 'CUSTOMER.COUNTRY'
loGroup.nHeight              = 3
loGroup.lPrintOnEachPage     = .T.
loGroup.nNewPageWhenLessThan = 4
```

## 字段和标签
您将添加到报表带区的最常见的对象就是字段（实际上，表达式可以是字段名称，也可以是任何有效的 FoxPro 表达式），因为报表的目的就是输出数据。 SFReportField 是用于字段的类。 但是，在我们讨论 SFReportField 之前，让我们看一下它的父类。

SFReportRecord 是 SFRepObj.vcx 中除 SFReportFile 和 SFReportBase 之外的每个类的祖先类。 它有一些用于所有对象的属性：

* Recno: 报表中对象的记录号。
* cComment: 报表记录的描述。
* cMemberData: 报表记录中的成员数据。
* cUniqueID: 报表记录的唯一ID（通常留空并自动分配）。
* cUser: 用户定义的属性。
* nObjectType: 包含对象的 FRX 记录的 OBJTYPE 值（例如，字段的 OBJTYPE 为 8）。

它还具有两个方法：CreateRecord 和 ReadFromFRX。 创建报表时，将从 SFReportFile 对象调用 CreateRecord。SFReportFile 使用 SCATTER NAME loRecord BLANK MEMO 从 FRX 中的记录中创建一个对象，为记录中的每个报表字段创建一个属性，然后将该对象传递给 SFReportRecord 对象的 CreateRecord 方法，该方法从其自己记录对象的属性值中填充 ReportRecord 对象的属性。 SFReportRecord 是一个抽象类; 它并不是让你直接使用的，而是作为其他类的父类。 其CreateRecord 方法只是确保传递了有效的 ReportRecord 对象，并将此对象的 ObjType 属性（被写入 FRX 中的 OBJTYPE 字段）设置为其 nObjectType 属性。 ReadFromFRX 正好相反：它从打开的 FRX 表中读取记录并将记录值分配给对应的属性。

SFReportObject 是 SFReportRecord 的子类，用于报表对象（这里，“对象”是指报表中呈现的对象，例如字段，而不是“当你实例化类时感觉到的”）。它具有如下公共属性，它们代表报表对象的最小选项集。

| 属性 | 描述                        |
|----------|--------------------------------|
| cAlignment	| 对象的对齐方式 "left", "center", or "right" (在 SFRepObj.h 中已经为这些值定义常量) |
| cName	| 对象名 (用于 SFReportBand.Item 通过名字定位 Item) |
| cPrintWhen	| 打印表达式 |
| lAutoCenter	| .T. (默认值) 表示在报表中使用字符个数作为单位时，将此对象垂直自动居中 |
| lPrintInFirstWholeBand	| .T. (默认值) 表示在新页的第一个带区中打印 |
| lPrintOnNewPage	| .T. 表示当细节带区溢出到新页时打印 |
| lPrintRepeats	| .T. (默认值) 表示打印重复值 |
| lRemoveLineIfBlank	| .T. 表示移除空行 |
| lStretch	| .T. 表示如果溢出时伸展 |
| lTransparent	| .T. (默认值)表示背景样式透明，.F. 则表示不透明 |
| nBackColor	| 对象的背景色，它使用 RGB 颜色值（默认值 = -1） |
| nFloat	| 如果对象在其带区内浮动，则为 0；如果它应相对于带区顶部固定，则为 1 （默认值）;如果它相对于带区底部固定，则为2（在 SFRepObj.h 中已经为这些值定义了常量）|
| nForeColor	| 对象的前景色，它使用 RGB 颜色值（默认值 = -1） |
| nGroup	| 如果此对象与其他对象分组,则为非零值 |
| nHeight	| 对象的高度 |
| nHPosition	| 对象的水平位置
| nPrintOnGroupChange	| 如果此对象应在组更改时打印，则为分组编号 |
| nVPosition	| 对象相对于带区顶部的垂直位置 |
| nWidth	| 对象的宽度 |

与其他类一样，这些属性只是将报表设计器中可用的选项公开为属性。

CreateRecord 方法首先使用 DODEFAULT() 来执行 SFReportRecord 的行为，然后进行一些数据转换。 例如，对象颜色存储在 FRX 记录的 PENRED，PENGREEN 和 PENBLUE 字段中，但我们想和其他 VFP 控件一样设置一个 nForeColor 属性（例如，使用RGB（255,0,0）设置为红色）。其他属性与之相似; 例如，nFloat 中的值更新到 FRX 中的 FLOAT，TOP 和 BOTTOM 字段。

最后，我们来看看 SFReportField，它是 SFReportObject 的子类，它包含有关报表中字段的信息。此类将以下属性添加到 SFReportObject 的属性。

| 属性 | 描述                        |
|----------|--------------------------------|
| cCaption	| 设计时刻字段的名字 |
| cDataType	| 表达式结果的数据类型："N" 数值, "D" 日期, "C" 其他 (仅在您稍后在报表设计器中编辑报表时才需要) |
| cExpression	| 显示的表达式 |
| cFontName	| 字体名（如果为空，则默认情况下使用SFReportFile.cFontName） |
| cPicture	| 字段的格式和掩码 |
| cTotalType	| 计算类型："N" 无, "C" 计数, "S" 求和, "A" 平均值, "L" 最小值, "H" 最大值, "D" 标准偏差, "V" 方差(在 SFRepObj.h 中为这些值定义了常量) |
| lFontBold	| .T. 表示对象使用粗体 |
| lFontItalic	| .T. 表示对象使用斜体 |
| lFontUnderline	| .T. 表示对象使用下划线 |
| lResetOnPage	| .T. 表示重置每页末尾的变量; .F. 表示在报表结束时重置 |
| nDataTrimming	| 指定如何设置字符表达式的修剪模式 |
| nFontCharSet	| 要使用的字体charset |
| nFontSize	| 使用的字号 (如果为 0，则默认情况下使用 SFReportFile.nFontSize 的值) |
| nResetOnDetail	| 要重置值的细节带区编号 |
| nResetOnGroup	| 要重置值得分组带区编号 |

如您所见，您可以像在报表设计器中一样控制报表中对象的属性。

与 SFReportObject 一样，SFReportField 的 CreateRecord 方法使用 DODEFAULT() 来获取 SFReportRecord 和 SFReportObject 的行为，然后它执行类似于SFReportObject 的数据转换（例如，lFontBold，lFontItalic 和 lFontUnderline 组合成单个 FONTSTYLE 值）。


SFReportText 是 SFReportField 的子类，因为它具有相同的属性但行为略有不同。 它会自动在表达式周围添加引号，因为文本对象始终包含文字字符串而不是表达式。 它还在 FRX 中设置 PICTURE 字段以匹配数据的对齐（因为这是文本对象的对齐方式），并根据文本的大小适当地调整对象的大小（换句话说，它就像设置标签对象得 AutoSize = .T. 一样）。

这里有一些代码可以将标签(text)和字段对象添加到细节带区并设置其属性。 此代码使用字符作为单位，因此值以字符或行为单位。

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

## 线、形状和图像
SFReportShape 是 SFReportObject 的子类，它定义线和形状的属性（它不是直接被使用）。 nPenPattern 是对象的样式：0 = 无，1 = 点，2 = 虚线，3 = 点划线，4 = 点点划线，8 = 常规。 nPenSize 是对象的粗细：0, 1, 2, 4 或 6 。

SFReportLine 是 SFReportShape 的子类，它用于线对象。 它添加了一个属性 lVertical ，.T. 表示创建垂直线，.F. (默认值)表示创建水平线。其CreateRecord 方法根据线的样式将水平线的高度或垂直线的高度设置为适当的值。

以下代码在页标头的第 4 行添加了一条粗蓝线：

```foxpro
loObject = loPageHeader.Add('Line')
loObject.nWidth     = lnWidth
loObject.nVPosition = 4
loObject.nHPosition = 0
loObject.nPenSize   = 6
loObject.nForeColor = rgb(0, 0, 255)
```

形状使用 SFReportBox ，它也是 SFReportShape 的子类。 它添加了 nCurvature（框角的曲率；默认值为 0，表示没有曲率），lStretchToTallest（.T. 表示相对于带区中的最高对象拉伸）和 nFillPattern（对象的填充图案： 0 = 无，1 = 实线，2 = 水平线，3 = 垂直线，4 = 对角线，左倾，5 = 对角线，右倾，6 = 网格，7 = 舱口）属性。

SFReportImage 是 SFReportObject 的子类，用于表示图像。 将 cImageSource 属性设置为图像文件的名称，或将 General 字段设置为图像的源，如果图像来自文件，则设置为 nImageSource = 0；如果来自 General 字段，则设置为 1;如果是表达式，则设置为 2。 nStretch 定义了如何缩放图像：0 = 剪切，1 = 缩放内容保留形状，2 = 缩放内容拉伸框架（与 Image 控件的 Stretch 属性使用的值相同）。 如果使用文件，其 CreateRecord 方法会自动在图像源周围放置引号，如果cAlignment 属性设置为 “Center”，则会在 FRX 记录中设置相应的字段（仅适用于“General”字段）。

## 报表变量
报表变量是使用SFReportFile对象的CreateVariable方法定义的。 此方法返回对其创建的SFReportVariable对象的对象引用，以便您可以设置变量的属性。 变量的公共属性是：

| 属性 | 描述                        |
|----------|--------------------------------|
| cInitialValue	| 初始值	|
| cName	| 变量名	|
| cTotalType	| 计算类型："N" 无, "C" 计数, "S" 求和, "A" 平均值, "L" 最小值, "H" 最大值, "D" 标准偏差, "V" 方差(在 SFRepObj.h 中为这些值定义了常量)	|
| cValue	| 被存储的值	|
| lReleaseAtEnd	| .T. 表示在报表输出后释放变量	|
| lResetOnPage	| .T. 表示在每页末重置变量；.F. 表示报表输出后重置变量	|
| nResetOnGroup	| 重置变量的组号	|

以下代码创建一个名为 lnCount 的报表变量，并指定它应从 0 开始，并为报表中打印的每个记录递增 1 。 然后在报告的摘要带区中打印此变量，显示打印的记录数。

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

## 示例
Customers.prg 和 Employees.prg 是针对 VFP Testdata 数据库中的 Customer 和 Employee 表创建报表的示例程序。 Customers.prg 创建（和预览）CustomerReport.frx，其显示按国家/地区分组的客户，最大订单金额按国家/地区小计，并在报表结尾处汇总。 Employees.prg 创建 EmployeeReport.frx，其中显示每个员工的姓名和照片。 这些报表并非切合实际; 它们只是展示了本文中描述的报表类的各种功能，包括打印图像和线条，设置字体大小和对象颜色，在不同带区中定位对象，使用报告变量和组带区等。

Craig Boyd 的 [GridExtras](https://tinyurl.com/ycmqo8fg): 对每列进行排序，增量搜索和过滤，列选择以及输出到 Excel 和打印预览。 “打印预览”功能使用 SFReportFile 根据当前表格布局动态创建报表。

使用这些类也可以更轻松地更改现有报表。 假设您有一个包含一些敏感信息的报表。 某些工作人员不应该看到该信息，因此您使用这些字段的打印表达式以便不输出它们，除非工作人员具有正确的权限。 例如，示例员工报表显示完整的员工列表，但出生日期和家庭电话应仅对人力资源（HR）员工可见。
 
这两个字段及其列标题的打印表达式为“plHR”。 变量 plHR 是 .T. 表示为人力资源人员，.F. 则表示除HR之外的人员。当报表由非 HR 人员运行时，BirthDate 和HomePhone 不会出现，但会在报表布局中留下明显的漏洞。 如果其他列可以向左移动以占用空白空间，那会更好。
 
HackReport.prg 显示了如何执行此操作。 代码并不复杂：实例化 SFReportFile，加载 FRX，删除由于其打印表达式而未出现的页眉和细节带区中的对象，并将其余字段向左移动适当的距离。

```foxpro
loReport = newobject('SFReportFile', 'SFRepObj.vcx')
loReport.Load('Employees.frx')

* 处理页眉和细节带区中的对象。

loBand = loReport.GetReportBand('Page Header')
ProcessObjects(loBand)
loBand = loReport.GetReportBand('Detail')
ProcessObjects(loBand)

* 保存更新的报表并运行它。

loReport.cReportFile = addbs(sys(2023)) + 'EmployeeReport.frx'
loReport.Save()
report form (loReport.cReportFile) preview

* 特别释放报表对象，以便进行适当的对象清理。

loReport.Release()

* 删除任何未通过其打印表达式的对象，并将对象移动到左侧同一行上这些对象的右侧。

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

## 总结
虽然您可能会编写相当多的代码来使用报表对象类创建 FRX，但代码其实很简单：创建一些对象并设置其属性。 它确实胜过（向 FRX 表）写入 50 个 INSERT INTO 语句，其中包含75 个字段以填充每个字段。 请向我报告任何有关改进的建议。
