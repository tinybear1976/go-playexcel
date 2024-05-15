package base

import (
	"errors"
	"fmt"
	"os"
	"path"
	"reflect"
	"testing"

	"github.com/shopspring/decimal"
)

func getTestExcelFile() string {
	currentDir, err := os.Getwd()
	if err != nil {
		return ""
	}
	return path.Join(currentDir, "test/test.xlsx")
}

func TestFillFixedStruct(t *testing.T) {
	var myXls TXlsx
	fn := getTestExcelFile()
	if fn == "" {
		t.Errorf("getTestExcelFile failed %s", fn)
	}
	t.Log(fn)
	myXls.OpenFile(fn)
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
		t.Logf("TestFillFixedStruct failed, err: %v", err)
	} else {
		t.Logf("TestFillFixedStruct success, %#v", t1)
	}

}

func TestFillList(t *testing.T) {
	var myXls TXlsx
	fn := getTestExcelFile()
	if fn == "" {
		t.Errorf("getTestExcelFile failed %s", fn)
	}
	t.Log(fn)
	myXls.OpenFile(fn)
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
		Month  int             `axis_y:"B"`
		Value1 decimal.Decimal `axis_y:"C"`
		Value2 decimal.Decimal `axis_y:"D"`
	}
	ss := Test2{}
	arr, err := myXls.FillList("eb2016", ss)
	if err != nil {
		t.Logf("TestFillFixedStruct failed, err: %v", err)
	} else {
		rst := arr.([]Test2)
		for _, v := range rst {
			fmt.Printf("%v\n", v)
		}
		//t.Logf("TestFillFixedStruct success, %#v", arr)
	}
}

func TestFillListStart4(t *testing.T) {
	var myXls TXlsx
	fn := getTestExcelFile()
	if fn == "" {
		t.Errorf("getTestExcelFile failed %s", fn)
	}
	t.Log(fn)
	myXls.OpenFile(fn)
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
		t.Logf("TestFillFixedStruct failed, err: %v", err)
	} else {
		rst := arr.([]Test2)
		for _, v := range rst {
			fmt.Printf("%v\n", v)
		}
		//t.Logf("TestFillFixedStruct success, %#v", arr)
	}
}

func TestFillListStart4End15(t *testing.T) {
	var myXls TXlsx
	fn := getTestExcelFile()
	if fn == "" {
		t.Errorf("getTestExcelFile failed %s", fn)
	}
	t.Log(fn)
	myXls.OpenFile(fn)
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
	arr, err := myXls.FillList("eb2016", ss, WithStartRow(4), WithEndRow(15))
	if err != nil {
		t.Logf("TestFillFixedStruct failed, err: %v", err)
	} else {
		rst := arr.([]Test2)
		for _, v := range rst {
			fmt.Printf("%v\n", v)
		}
		//t.Logf("TestFillFixedStruct success, %#v", arr)
	}
}

func TestDecimalReflectType(t *testing.T) {
	id := &innerDec{}
	err := DecRTyp(id)
	if err != nil {
		t.Logf("TestDecimalReflectType failed, err: %v", err)
	} else {
		t.Logf("TestDecimalReflectType success, %#v", id)
	}
}

type innerDec struct {
	Money decimal.Decimal
}

func DecRTyp(pObj any) error {
	valPtr := reflect.ValueOf(pObj)
	if valPtr.Kind() != reflect.Ptr {
		return errors.New("传入的参数pObj必须为结构体指针")
	}
	if valPtr.Elem().Kind() != reflect.Struct {
		return errors.New("传入的参数pObj必须为结构体指针")
	}

	elem := valPtr.Elem()
	for i := 0; i < elem.NumField(); i++ {
		field := elem.Type().Field(i)
		fmt.Printf("func log: %v\n", field.Type.Kind())
		fmt.Printf("func log: %v\n", field.Type.Name())
		// switch field.Type.Kind() {
		// case reflect.String:
		// 	// elem.Field(i).SetString(v)
		// case reflect.Int:
		// 	// n, _err := strconv.Atoi(v)
		// 	// if _err != nil {
		// 	// 	n = 0
		// 	// }
		// 	// elem.Field(i).SetInt(int64(n))
		// case reflect.Float32, reflect.Float64:
		// 	// n, _err := strconv.ParseFloat(v, 64)
		// 	// if _err != nil {
		// 	// 	n = 0.0
		// 	// }
		// 	// elem.Field(i).SetFloat(n)
		// case reflect.Bool:
		// 	// elem.Field(i).SetBool(string2bool(v))

		// }
	}
	return nil

}
