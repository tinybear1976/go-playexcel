package base

import (
	"fmt"
	"testing"
)

func TestFillFixedStruct(t *testing.T) {
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
