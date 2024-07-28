package base

import (
	"errors"
	"reflect"
	"slices"
	"strconv"
	"strings"

	"github.com/shopspring/decimal"
	"github.com/tinybear1976/go-playexcel/core"
)

const (
	tupleTag = "axis"
	listTag  = "axis_y"
)

var (
	trueValues = []string{"true", "1", "yes", "y", "t", "on", "是", "对", "真"}
)

// 根据sheetName,填充固定对象,此方法采用ShrinkData进行操作，请确保ShrinkData已经填充完毕
func (xls *TXlsx) FillTuple(fromSheetName string, pObj any) error {
	if !xls.IsOpened() {
		return errors.New("未指定需要打开的文件名")
	}
	valPtr := reflect.ValueOf(pObj)
	if valPtr.Kind() != reflect.Ptr {
		return errors.New("传入的参数pObj必须为结构体指针")
	}
	if valPtr.Elem().Kind() != reflect.Struct {
		return errors.New("传入的参数pObj必须为结构体指针")
	}

	if data, ok := xls.sheetsShrinkData[fromSheetName]; ok {
		// valPtr.Elem().Set(reflect.ValueOf(data))
		elem := valPtr.Elem()
		for i := 0; i < elem.NumField(); i++ {
			field := elem.Type().Field(i)
			tag := field.Tag.Get(tupleTag)
			if tag != "" {
				v, err := core.GetCellValueByTag(tag, data)
				if err == nil {
					switch field.Type.Kind() {
					case reflect.String:
						elem.Field(i).SetString(v)
					case reflect.Int:
						n, _err := strconv.Atoi(v)
						if _err != nil {
							n = 0
						}
						elem.Field(i).SetInt(int64(n))
					case reflect.Float32, reflect.Float64:
						n, _err := strconv.ParseFloat(v, 64)
						if _err != nil {
							n = 0.0
						}
						elem.Field(i).SetFloat(n)
					case reflect.Bool:
						elem.Field(i).SetBool(string2bool(v))
					case reflect.Struct:
						if field.Type.Name() == "Decimal" {
							// Decimal
							var dec decimal.Decimal
							if v == "" {
								dec = decimal.NewFromInt(0)
								elem.Field(i).Set(reflect.ValueOf(dec))
							} else {
								dec, _err := decimal.NewFromString(v)
								if _err != nil {
									dec = decimal.NewFromInt(0)
								}
								elem.Field(i).Set(reflect.ValueOf(dec))
							}
						}
					}
				}
			}
		}
		return nil
	} else {
		return errors.New("未找到sheetName:" + fromSheetName + "的sheet")
	}
}

func string2bool(v string) bool {
	v = strings.ToLower(v)
	return slices.Contains(trueValues, v)
}

// 行填充  =========================================================================

type TParamates struct {
	StartRow    int
	EndRow      int
	IgnoreEmpty bool
}

type Option func(*TParamates)

func WithStartRow(startRow int) Option {
	if startRow < 1 {
		startRow = 1
	}
	return func(p *TParamates) {
		p.StartRow = startRow
	}
}

func WithEndRow(endRow int) Option {
	if endRow < 1 {
		endRow = -1
	}
	return func(p *TParamates) {
		p.EndRow = endRow
	}
}

func WithIgnoreAllEmptyRow(ignore bool) Option {
	return func(p *TParamates) {
		p.IgnoreEmpty = ignore
	}
}

// 根据sheetName,返回列表,每行数据结构为rowItem结构体定义，此方法采用ShrinkData进行操作，请确保ShrinkData已经填充完毕
func (xls *TXlsx) FillList(fromSheetName string, rowItem any, opts ...Option) (any, error) {
	paramates := &TParamates{
		StartRow:    1,
		EndRow:      -1,
		IgnoreEmpty: true,
	}
	for _, opt := range opts {
		opt(paramates)
	}
	if !xls.IsOpened() {
		return nil, errors.New("未指定需要打开的文件名")
	}
	_obj := reflect.ValueOf(rowItem)
	if _obj.Kind() != reflect.Struct {
		return nil, errors.New("传入的参数Obj必须为结构体")
	}
	// if _slices.Kind() != reflect.Ptr {
	// 	return errors.New("传入的参数listObj必须为切片指针")
	// }
	// if _slices.Elem().Kind() != reflect.Slice {
	// 	return errors.New("传入的参数listObj必须为切片指针")
	// }
	// fmt.Println(reflect.TypeOf(listObj).Elem().Elem())
	// reflect.TypeOf(listObj).Elem().Elem() 结构体

	if data, ok := xls.sheetsShrinkData[fromSheetName]; ok {
		_structType := reflect.TypeOf(rowItem)
		// fmt.Printf("log: %v\n", _structType)
		_structSlicesType := reflect.SliceOf(_structType)
		// fmt.Printf("log: %v\n", _structSlicesType)
		_newSlices := reflect.MakeSlice(_structSlicesType, 0, 0)
		if len(data) <= 0 {
			return _newSlices, nil
		}
		if paramates.EndRow == -1 {
			paramates.EndRow = len(data)
		}
		for x := paramates.StartRow; x <= paramates.EndRow; x++ {
			// 创建结构体的新实例
			_item := reflect.New(_structType).Elem()
			if paramates.IgnoreEmpty {
				allEmpty, err := checkAllEmpty(_item, x, data)
				if err != nil {
					return nil, err
				}
				if allEmpty {
					continue
				}
			}

			for i := 0; i < _item.NumField(); i++ {
				field := _item.Type().Field(i)
				y := field.Tag.Get(listTag)
				if y != "" {
					v, err := core.GetCellValueByTag(y+strconv.Itoa(x), data)
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
	} else {
		return nil, errors.New("未找到sheetName为" + fromSheetName + "的sheet")
	}
}

func checkAllEmpty(obj reflect.Value, x int, data [][]string) (bool, error) {
	excepted := 0
	empty := 0
	for i := 0; i < obj.NumField(); i++ {
		field := obj.Type().Field(i)
		y := field.Tag.Get(listTag)
		if y != "" {
			excepted++
			v, err := core.GetCellValueByTag(y+strconv.Itoa(x), data)
			if err != nil {
				return false, err
			}
			if v == "" {
				empty++
			}
		}
	}
	return empty == excepted, nil
}
