package base

import (
	"fmt"
	"strings"

	"github.com/tinybear1976/go-playexcel/core"

	"errors"

	"github.com/xuri/excelize/v2"
)

type TXlsx struct {
	// 调用InitFile或InitFileAndReadAll方法后，保存的Excel文件名
	filename string
	// 记录该文件中所有sheet的名称
	sheetsName map[int]string
	// 根据sheet名，保存的对齐数据（每行字段数量一致）
	sheetsAlignedData map[string][][]string
	// 根据sheet名，保存的缩减数据（主要缩减的是空列)
	sheetsShrinkData map[string][][]string
	// 用于临时记录某一次转换过程中的错误信息。当执行某个Fill函数后被改变。原则，如果没有任何错误发生，默认值为nil，如果发生错误才会形成切片
	convertErrors [][]string
}

// 获得最近一次转换填充(Fill类函数)动作中所产生的转换错误.
// 如果没有错误，返回nil，否则返回错误信息切片
// 返回的二维数组的第二维格式:  {"D3","原始字符串","整数|浮点数|定点数"}
func (xls TXlsx) GetLastConvertErrors() [][]string {
	return xls.convertErrors
}

// 通过内部文件名是否存在判断是否打开过文件
func (xls *TXlsx) IsOpened() bool {
	return xls.filename != ""
}

// 获得文件名
func (xls *TXlsx) GetFilename() string {
	return xls.filename
}

// 获得所有sheet名。如果没有打开过文件，返回nil
func (xls *TXlsx) GetSheetsName() map[int]string {
	return xls.sheetsName
}

// 检查某个sheetName是否存在,如果存在返回原始sheetName，否则返回空字符串
func (xls *TXlsx) SheetNameIsExist(sheetName string) string {
	if xls.sheetsName == nil {
		return ""
	}
	for _, name := range xls.sheetsName {
		if strings.EqualFold(name, sheetName) {
			return name
		}
	}
	return ""
}

// 首次运行该方法。调用该方法只会读取sheet名称列表，并不会加载sheet数据
func (xls *TXlsx) OpenFile(filename string) error {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		//fmt.Println(err)
		return err
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// 获取工作表中的所有工作簿
	sheets := f.GetSheetMap()
	xls.filename = filename
	xls.sheetsName = sheets
	// xls.sheetsAlignedData = make(map[string][][]string)
	// xls.sheetsShrinkData = make(map[string][][]string)
	// for idx, name := range sheets {
	// 	fmt.Println(idx, name)
	// }
	return nil
}

// 首次运行该方法，调用该方法，会自动将所有sheet的数据全部加载到内存
func (xls *TXlsx) OpenFileAndReadAll(filename string) error {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		//fmt.Println(err)
		return err
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// 获取工作表中的所有工作簿
	sheets := f.GetSheetMap()
	xls.filename = filename
	xls.sheetsName = sheets
	xls.sheetsAlignedData = make(map[string][][]string)
	xls.sheetsShrinkData = make(map[string][][]string)
	for _, sheetName := range sheets {
		// 获取到当前工作簿的所有合并单元格信息
		mergeCells, err := f.GetMergeCells(sheetName)
		if err != nil {
			return err
		}
		// 获取 Sheet1 上所有单元格
		rows, err := f.GetRows(sheetName)
		if err != nil {
			return err
		}
		fullRows, err := arrangeRows(rows, mergeCells)
		if err == nil {
			xls.sheetsAlignedData[sheetName] = fullRows
			shrinkArr, err := core.ShrinkArray(fullRows)
			if err == nil {
				xls.sheetsShrinkData[sheetName] = shrinkArr
				// fmt.Println("table shrink width: ", core.MaxWidth(shrinkArr))
			}
		}
	}
	return nil
}

// 复位重置对象
func (xls *TXlsx) Reset() {
	xls.filename = ""
	xls.sheetsName = nil        // make(map[int]string)
	xls.sheetsAlignedData = nil // make(map[string][][]string)
	xls.sheetsShrinkData = nil  // make(map[string][][]string)
	xls.convertErrors = nil
}

// 获取sheet对齐数据(所有行)
func (xls *TXlsx) GetSheet(sheetName string) ([][]string, error) {
	if xls.sheetsAlignedData != nil {
		if rowsData, ok := xls.sheetsAlignedData[sheetName]; ok {
			return rowsData, nil
		}
	}

	if !xls.IsOpened() {
		return nil, errors.New("未指定需要打开的文件名")
	}
	f, err := excelize.OpenFile(xls.filename)
	if err != nil {
		//fmt.Println(err)
		return nil, err
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// 获取到当前工作簿的所有合并单元格信息
	mergeCells, err := f.GetMergeCells(sheetName)
	if err != nil {
		return nil, err
	}
	// for _, mc := range mergeCells {
	// 	fmt.Print("StartAxis: ", mc.GetStartAxis(), "\t")
	// 	row, col, _ := core.UnpackAxisName(mc.GetStartAxis())
	// 	fmt.Printf(" [row %d,col %d]", row, col)
	// 	fmt.Print("EndAxis", mc.GetEndAxis(), "\t")
	// 	row, col, _ = core.UnpackAxisName(mc.GetEndAxis())
	// 	fmt.Printf(" [row %d,col %d]", row, col)
	// 	fmt.Println("Value", mc.GetCellValue())

	// }

	// 获取 Sheet1 上所有单元格
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, err
	}
	fullRows, err := arrangeRows(rows, mergeCells)
	if err == nil {
		if xls.sheetsAlignedData == nil {
			xls.sheetsAlignedData = make(map[string][][]string)
		}
		xls.sheetsAlignedData[sheetName] = fullRows
		shrinkArr, err := core.ShrinkArray(fullRows)
		if err == nil {
			if xls.sheetsShrinkData == nil {
				xls.sheetsShrinkData = make(map[string][][]string)
			}
			xls.sheetsShrinkData[sheetName] = shrinkArr
			// fmt.Println("table shrink width: ", core.MaxWidth(shrinkArr))
		}
	}
	return fullRows, err
}

func (xls *TXlsx) GetSheetShrinkRow(sheetName string) [][]string {
	if xls.sheetsShrinkData == nil {
		return nil
	}
	if rowsData, ok := xls.sheetsShrinkData[sheetName]; ok {
		return rowsData
	}
	return nil
}

// 将sheet读取的原始数组整理为全量的规整数组（即全部等长，解除并填充合并的单元格）
func arrangeRows(raw [][]string, mergeCells []excelize.MergeCell) ([][]string, error) {
	_maxWidth := core.MaxWidth(raw)
	rst := make([][]string, 0)
	for rowIndex, row := range raw {
		if len(row) < _maxWidth {
			row = append(row, make([]string, _maxWidth-len(row))...)
		}
		for colIndex := range row {
			ok, text, err := core.IsWithInMergeScope(colIndex+1, rowIndex+1, mergeCells)
			if err != nil {
				return nil, err
			}
			if ok {
				row[colIndex] = text
			}
		}
		rst = append(rst, row)
	}
	return rst, nil
}
