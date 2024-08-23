# 基本

## 简介

这个模块主要利用 golang 的第三方库 `github.com/xuri/excelize/v2`进行基本的读取操作，在此基础上添加了一部分 Excel 数据到 golang 数据转换的方法

基本概念：

- excel 文件
- sheet 名称
- sheet 页数据（由 github.com/xuri/excelize/v2 读取出来的二维数组，我们认为是基础数据。该数据在本模块中没有进行保存）
- sheet 对齐数据。sheet 基础数据 github.com/xuri/excelize/v2 应该是对齐做了数据的优化，会自动识别每行数据某列为最后一列有效列，所以返回的[][]string 数据并不是每行都具有同样长度的。基于此问题，我保存了一份对齐（冗余）的二维数据，主要是为避免访问二维数组时产生越界问题
- sheet 收缩数据。在 sheet 对齐数据的基础上，近一步根据每行**最左侧**保证全部是有效列的原则，对对齐数据进行了适当裁剪。在实际的一些填充功能访问内存数据时，主要数据来源全部采用收缩数据作为基础。收缩数据区别于 github.com/xuri/excelize/v2 读取的格式：收缩数据为了保持 Table 的特性（每行的列数一致），所以会在某些行的结尾保留适当的空数据列，而 github.com/xuri/excelize/v2 作者的数据裁剪是依据每一行进行的，所以数据无法对齐。
- FillTuple。就是通过 sheet 数据填充一个单一形态对象。结构体字段 Tag 标记为 `axis`
- FillList。通过 sheet 数据填充一个列表形态的数据对象。结构体字段 Tag 标记为 `axis_y`

# 方法

## OpenFile

方法：OpenFile(filename string) error
功能：首次运行该方法。调用该方法只会读取 sheet 名称列表，并不会加载 sheet 数据。

参数：

- filename (string): 需要打开的 Excel 路径及文件名。

返回值：
返回一个错误值，如果文件打开失败。

示例：

```go
var xls playXLS.TXlsx
err := xls.OpenFile("example.xlsx")
if err != nil {
    fmt.Println("打开文件失败:", err)
}
```

## OpenFileAndReadAll

方法：OpenFileAndReadAll(filename string) error
功能：首次运行该方法。调用该方法会加载全部 sheet 数据。

参数：

- filename (string): 需要打开的 Excel 路径及文件名。

返回值：
返回一个错误值，如果文件打开失败。

示例：

```go
var xls playXLS.TXlsx
err := xls.OpenFileAndReadAll("example.xlsx")
if err != nil {
    fmt.Println("打开文件失败:", err)
}
```

## GetSheetNames

方法：GetSheetNames() ([]string, error)
功能：获取 sheet 名称列表。

参数：
无。

返回值：
返回一个字符串切片，包含所有 sheet 名称。

示例：

```go
var xls playXLS.TXlsx
err := xls.OpenFile("example.xlsx")
if err != nil {
    fmt.Println("打开文件失败:", err)
}
sheetNames, err := xls.GetSheetNames()
if err != nil {
    fmt.Println("获取 sheet 名称列表失败:", err)
}
fmt.Println(sheetNames)
```

## IsOpened

方法：IsOpened() bool
功能：检查文件是否已打开。

参数：
无。

返回值：
返回一个布尔值，如果文件已打开则返回 true，否则返回 false。

示例：

```go
var xls playXLS.TXlsx
err := xls.OpenFile("example.xlsx")
if err != nil {
    fmt.Println("打开文件失败:", err)
}
if xls.IsOpened() {
    fmt.Println("xls对象打开过文件")
} else {
    fmt.Println("xls对象未打开过文件")
}
```

## GetFilename

方法：GetFilename() string
功能：获取当前打开的 Excel 文件名。

参数：
无。

返回值：
返回一个字符串，表示当前打开的 Excel 文件名。

示例：

```go
var xls playXLS.TXlsx
err := xls.OpenFile("example.xlsx")
if err != nil {
    fmt.Println("打开文件失败:", err)
}
filename := xls.GetFilename()
fmt.Println(filename)
```

## Reset

方法：Reset()
功能：重置 xls 对象，清除文件名，sheet 名称列表及所有相关加载数据。

参数：
无。

返回值：
无。

示例：

```go
var xls playXLS.TXlsx
err := xls.OpenFile("example.xlsx")
if err != nil {
    fmt.Println("打开文件失败:", err)
}
xls.Reset()
```

## SheetNameIsExist

方法：SheetNameIsExist(sheetName string) string
功能：检查 sheet 名称是否存在。不区分大小写。

参数：

- sheetName (string): 需要检查的 sheet 名称。

返回值：
如果 sheet 名称不存在则返回空字符串，否则返回原始 sheet 名称。

示例：

```go
var xls playXLS.TXlsx
err := xls.OpenFile("example.xlsx")
if err != nil {
    fmt.Println("打开文件失败:", err)
}
sheetName := "base"
if name:=xls.SheetNameIsExist(sheetName); name== "" {
    fmt.Println("sheet名称不存在")
} else {
    fmt.Println("sheet存在，原名:",name)
}
```

## GetSheet

方法：GetSheet(sheetName string) ([][]string, error)
功能：获取指定 sheet 的数据。同时该数据（对齐数据、收缩数据）都将被记录在内存对象中，返回的是对齐数据。配合优化，如果内存中存在该部分数据，则直接返回；若不存在则从 Excel 文件中读取，并存入内存中，同时计算出收缩数据，一起保存在内存中。

参数：

- sheetName (string): 需要获取的 sheet 名称。

返回值：
返回该 sheet 的对齐数据，如果 sheetName 无效则返回 nil。

示例：

```go
var xls playXLS.TXlsx
err := xls.OpenFile("example.xlsx")
if err != nil {
    fmt.Println("打开文件失败:", err)
    return
}
sheetName := "base"
data, err := xls.GetSheet(sheetName)
if err != nil {
    fmt.Println("获取 sheet 数据失败:", err)
}
fmt.Printf("sheet: %s \n对齐数据: %v\n",sheetName,data)
```

## GetSheetShrink

方法：GetSheetShrink(sheetName string) ([][]string, error)
功能：获取指定 sheet 的收缩数据。配合动作的优化，这个方法只是直接读取内存中的数据，不直接从 Excel 文件中读取。

参数：

- sheetName (string): 需要获取的 sheet 名称。

返回值：
返回该 sheet 的收缩数据，如果 sheetName 无效则返回 nil。

示例：

```go
var xls playXLS.TXlsx
err := xls.OpenFileAndReadAll("example.xlsx")
if err != nil {
    fmt.Println("打开文件失败:", err)
    return
}
sheetName := "base"
data, err := xls.GetSheetShrink(sheetName)
if err != nil {
    fmt.Println("获取 sheet 数据失败:", err)
}
fmt.Printf("sheet: %s \n收缩数据: %v\n",sheetName,data)

// 或者

var xls playXLS.TXlsx
err := xls.OpenFile("example.xlsx")
if err != nil {
    fmt.Println("打开文件失败:", err)
    return
}
sheetName := "base"
_, err := xls.GetSheet(sheetName)
if err != nil {
    fmt.Println("获取 sheet 数据失败:", err)
}
data, err := xls.GetSheetShrink(sheetName)
if err != nil {
    fmt.Println("获取 sheet 数据失败:", err)
}
fmt.Printf("sheet: %s \n收缩数据: %v\n",sheetName,data)
```

## GetLastConvertErrors

方法：GetLastConvertErrors() [][]string

功能：获取 FillXXX 函数填充过程发生的转换错误信息。每次调用一个 Fill 类函数后跟随调用。每次调用 Fill 类函数，当 Fill 类函数进入到实质性转换前，都会将错误信息列表设置为 nil，所以无法保存历史错误信息。

## FillTuple

方法：FillTuple(fromSheetName string, pObj any) error
功能：根据 sheetName,填充固定对象(结构体),此方法采用 ShrinkData 进行操作，请确保 ShrinkData 已经填充完毕。

参数：

- fromSheetName (string): 需要读取的 sheet 名称。
- pObj (any): 必须为结构体指针，最终方法会将数据从 excel 中反射填充到对象字段内。结构体字段 Tag 应该采用 `axis` 进行标识，如：`axis:"C5"`。表示该字段对应值所在的单元格。

返回值：
返回一个错误值，如果 sheetName 无效或数据填充失败。

示例：

```go
	myXls.OpenFile("/Users/benson/program/go/apsec/sanxia/impdata/test/test.xlsx")
	sheetsName := myXls.GetSheetsName()
	for key, sheetName := range sheetsName {
		fmt.Println(key, sheetName)
		_, err := myXls.GetSheet(sheetName)
		if err != nil {
			fmt.Println("load sheet error: ", sheetName, err)
			continue
		}
	}
	type Test1 struct {
		Id   string `axis:"C5"`
		Name string `axis:"C6"`
		City string `axis:"C7"`
		Area int    `axis:"C8"`
	}
	t1 := &Test1{}
	err := myXls.FillTuple("base", t1)
	if err != nil {
		t.Logf("failed, err: %v", err)
	} else {
		t.Logf("success, %#v", t1)
	}
```

## FillList

方法：FillList(fromSheetName string, rowItem any, opts ...Option) (any, error)
功能：根据 sheetName,返回列表,每行数据结构为 rowItem 结构体定义，此方法采用 ShrinkData 进行操作，请确保 ShrinkData 已经填充完毕

参数：

- fromSheetName (string): 需要读取的 sheet 名称。
- rowItem (any): 结构实例。主要用于确定表格中每行数据的样式。结构体 Tag 应该采用`axis_y`用来标记字段对应的列
- opts (...Option): 配置选项，可以设置行数限制。默认情况从第一行至最后一行。如果需要指定开始行可以使用`WithStartRow(int)`方法。若需要指定结束行可以使用`WithEndRow(int)`方法。`WithIgnoreAllEmptyRow(bool)`方法用于预判是否放弃全空行，默认值 true

返回值：
返回一个结构体切片。该结构体切片由传入的结构体类型组成

示例：

```go
	var myXls TXlsx
	myXls.OpenFile("/Users/benson/program/go/apsec/sanxia/impdata/test/test.xlsx")
	sheetsName := myXls.GetSheetsName()
	for key, sheetName := range sheetsName {
		fmt.Println(key, sheetName)
		_, err := myXls.GetSheet(sheetName)
		if err != nil {
			fmt.Println("load sheet error: ", sheetName, err)
			continue
		}
	}

	type Test2 struct {
		Month  string `axis_y:"B"`
		Value1 string `axis_y:"C"`
		Value2 string `axis_y:"D"`
	}
	ss := Test2{}
	arr, err := myXls.FillList("eb2016", ss, WithStartRow(4))
	if err != nil {
		t.Logf("failed, err: %v", err)
	} else {
		rst := arr.([]Test2)
		for _, v := range rst {
			fmt.Printf("%v\n", v)
		}
	}
```

## FillVerticalList

方法：FillVerticalList(fromSheetName string, item any, opts ...VerticalOption) (any, error)
功能：根据 sheetName,返回列表,每行数据结构为 rowItem 结构体定义，此方法采用 ShrinkData 进行操作，请确保 ShrinkData 已经填充完毕

参数：

- fromSheetName (string): 需要读取的 sheet 名称。
- item (any): 结构实例。主要用于确定表格中每行数据的样式。结构体 Tag 应该采用`axis_y`用来标记字段对应的列。因为是垂直方向列表，所以 tag 值例： 'axis_y:2'
- opts (...VerticalOption): 配置选项，可以设置行数限制。默认情况从第一行至最后一行。如果需要指定开始行可以使用`WithStartColumn(int)`方法。若需要指定结束行可以使用`WithEndColumn(int)`方法。
  `WithIgnoreAllEmpty(bool)`参数用于指明是否忽略全空列，默认忽略，其主要原因考虑到有些元组性数据与垂直列数据混合排版时，可能会因为其他不相关数据列数过多，导致垂直列表位置出现全空列。

返回值：
返回一个结构体切片。该结构体切片由传入的结构体类型组成

示例：

```go
	var myXls TXlsx
	myXls.OpenFile("/Users/benson/program/go/apsec/sanxia/impdata/test/test.xlsx")
	sheetsName := myXls.GetSheetsName()
	for key, sheetName := range sheetsName {
		fmt.Println(key, sheetName)
		_, err := myXls.GetSheet(sheetName)
		if err != nil {
			fmt.Println("load sheet error: ", sheetName, err)
			continue
		}
	}

	type Test2 struct {
		Month  string `axis_y:"3"`
		Value1 string `axis_y:"4"`
		Value2 string `axis_y:"5"`
	}
	ss := Test2{}
	arr, err := myXls.FillList("eb2016", ss, WithStartColumn(4))
	if err != nil {
		t.Logf("failed, err: %v", err)
	} else {
		rst := arr.([]Test2)
		for _, v := range rst {
			fmt.Printf("%v\n", v)
		}
	}
```

## 关于数值转换异常的补充处理

无论哪种填充方式，当结构体字段设置为整形或浮点或 decimal 时，可以在 tag 的内部使用逗号分隔，用于设置当转换失败时，采用的默认值.
另外，如果结构体中某个字段是 string 型,在 Tag 中增加第三参数，则表示这个字段未来会以某种数字形式进行使用，所以可对其进行尝试转换，观察其是否会发生转换错误。第三参数可以设置为：int、float、decimal。
以上内容都的转换错误会被记录到实例内部的日志中，可以调用：GetLastConvertErrors() 获得详细信息

```go
	type Test2 struct {
		Month  string `axis_y:"3"`
		Value1 int `axis_y:"4"`
		Value2 float64 `axis_y:"5"`
	}

	type Test2 struct {
		Month  string `axis_y:"3"`
		Value1 int `axis_y:"4,-1"`
		Value2 float64 `axis_y:"5"`
	}

  	type Test3 struct {
		Month  string `axis_y:"3,,int"`
		Value1 string `axis_y:"4,,float"`
		Value2 string `axis_y:"5,,decimal"`
	}
```
