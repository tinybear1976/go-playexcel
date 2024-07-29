package core

import (
	"errors"
	"math"
	"strconv"
	"strings"
	"unicode"
)

// excel轴列以字母方式表达，行以数字方式表达，单元内所有函数行、列都默认基于1开始

// excel轴分析，根据轴名称返回 row，col，错误
func UnpackAxisName(axisName string) (row int, col int, err error) {
	axisName = strings.ToUpper(axisName)
	_colBuilder := strings.Builder{}
	_rowBuilder := strings.Builder{}
	// 开始默认标记为列状态，一旦遇到数字将修改为行状态，并且不允许字母与数字间隔书写。
	_colSequence := true
	for idx, c := range axisName {
		if idx == 0 && !unicode.IsLetter(c) {
			return 0, 0, errors.New("axisName 必须以字母开始")
		}
		// 如果字符是字母，则认为是列
		if unicode.IsLetter(c) {
			if !_colSequence {
				return 0, 0, errors.New("axisName 列字母必须连贯书写, 如A1, AF2")
			}
			_colBuilder.WriteRune(c)
		} else if unicode.IsNumber(c) {
			if _colSequence {
				_colSequence = false
			}
			_rowBuilder.WriteRune(c)
		} else {
			return 0, 0, errors.New("axisName 只能包含字母和数字")
		}
	}
	col = calcColumn(_colBuilder.String())
	if col == 0 {
		return 0, 0, errors.New("axisName 列必须大于0")
	}
	row, err = calcRow(_rowBuilder.String())
	return
}

// 转换为列号（入参应该为字母表达的字符串）
func calcColumn(colString string) int {
	colString = strings.ToUpper(colString)
	_colLen := len(colString)
	var _col float64 = 0
	for idx, c := range colString {
		_sv := float64(c - 'A' + 1)
		_bitWeight := float64(_colLen - idx - 1)
		_col += _sv * math.Pow(26, _bitWeight)
	}
	return int(_col)
}

// 转换为行号（入参应该为数字表达的字符串）
func calcRow(rowString string) (int, error) {
	r, err := strconv.Atoi(rowString)
	if err != nil {
		return 0, err
	}
	if r == 0 {
		return 0, errors.New("axisName 行必须大于0")
	}
	return r, nil
}

// 将数字转为字母列号.
//
//	colIndex  int  列号，必须从1开始
//	@return     string  转换后的字母列号，如果转换不成功返回空字符串
func ConvertColumnToLetter(colIndex int) string {
	if colIndex < 1 {
		return ""
	}
	name := ""
	for colIndex > 0 {
		mod := (colIndex - 1) % 26
		name = string(rune(mod+'A')) + name
		colIndex = (colIndex - 1) / 26
	}
	return name
}
