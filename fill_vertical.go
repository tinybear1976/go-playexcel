package base

import (
	"errors"
	"reflect"
	"strconv"

	"github.com/shopspring/decimal"
	"github.com/tinybear1976/go-playexcel/core"
)

type TVerticalParamates struct {
	StartColumn int
	EndColumn   int
}

type VerticalOption func(*TVerticalParamates)

func WithStartColumn(startCol int) VerticalOption {
	if startCol < 1 {
		startCol = 1
	}
	return func(p *TVerticalParamates) {
		p.StartColumn = startCol
	}
}

func WithEndColumn(endCol int) VerticalOption {
	if endCol < 1 {
		endCol = -1
	}
	return func(p *TVerticalParamates) {
		p.EndColumn = endCol
	}
}

// FillVerticalList 填充垂直列表
func (xls *TXlsx) FillVerticalList(fromSheetName string, item any, opts ...VerticalOption) (any, error) {
	paramates := &TVerticalParamates{
		StartColumn: 1,
		EndColumn:   -1,
	}
	for _, opt := range opts {
		opt(paramates)
	}
	if !xls.IsOpened() {
		return nil, errors.New("未指定需要打开的文件名")
	}
	_obj := reflect.ValueOf(item)
	if _obj.Kind() != reflect.Struct {
		return nil, errors.New("传入的参数Obj必须为结构体")
	}
	if data, ok := xls.sheetsShrinkData[fromSheetName]; ok {
		_structType := reflect.TypeOf(item)
		// fmt.Printf("log: %v\n", _structType)
		_structSlicesType := reflect.SliceOf(_structType)
		// fmt.Printf("log: %v\n", _structSlicesType)
		_newSlices := reflect.MakeSlice(_structSlicesType, 0, 0)
		if len(data) <= 0 {
			return _newSlices.Interface(), nil
		}
		if paramates.EndColumn == -1 {
			paramates.EndColumn = len(data[0]) // 获得列数，等价于行式列表中的行数
		}

		for y := paramates.StartColumn; y <= paramates.EndColumn; y++ {
			// 创建结构体的新实例
			_item := reflect.New(_structType).Elem()
			for i := 0; i < _item.NumField(); i++ {
				field := _item.Type().Field(i)
				x := field.Tag.Get(listTag)
				if x != "" {
					// 该x值必须为整数
					_x, err := strconv.Atoi(x)
					if err != nil {
						return nil, errors.New("垂直列表的Tag值必须为整数.例如 axis_y:'2'")
					}
					v, err := core.GetCellValueByVerticalTag(_x, y, data)
					if err == nil {
						switch field.Type.Kind() {
						case reflect.String:
							_item.Field(i).SetString(v)
						case reflect.Int:
							n, _err := strconv.Atoi(v)
							if _err != nil {
								n = 0
							}
							_item.Field(i).SetInt(int64(n))
						case reflect.Float32, reflect.Float64:
							n, _err := strconv.ParseFloat(v, 64)
							if _err != nil {
								n = 0.0
							}
							_item.Field(i).SetFloat(n)
						case reflect.Bool:
							_item.Field(i).SetBool(string2bool(v))
						case reflect.Struct:
							if field.Type.Name() == "Decimal" {
								// Decimal
								var dec decimal.Decimal
								if v == "" {
									dec = decimal.NewFromInt(0)
									_item.Field(i).Set(reflect.ValueOf(dec))
								} else {
									dec, _err := decimal.NewFromString(v)
									if _err != nil {
										dec = decimal.NewFromInt(0)
									}
									_item.Field(i).Set(reflect.ValueOf(dec))
								}
							}
						}
					}
				}
			}
			_newSlices = reflect.Append(_newSlices, _item)
		}
		return _newSlices.Interface(), nil
	}
	return nil, errors.New("未找到sheetName为" + fromSheetName + "的sheet")

}
