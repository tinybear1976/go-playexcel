package core

import "fmt"

func GetCellValueByTag(axisName string, alignedArr [][]string) (string, error) {
	r, c, err := UnpackAxisName(axisName)
	if err != nil {
		return "", err
	}
	if r < 1 || r > len(alignedArr) {
		return "", fmt.Errorf("row index out of range")
	}
	cells := alignedArr[r-1]
	if c < 1 || c > len(cells) {
		return "", fmt.Errorf("column index out of range")
	}
	return cells[c-1], nil
}

// 用于获得列式表的数据。tagValue为结构体中标记的值，该值直接用数值表示。
func GetCellValueByVerticalTag(tagValue int, colValue int, alignedArr [][]string) (string, error) {
	r, c := tagValue, colValue
	if r < 1 || r > len(alignedArr) {
		return "", fmt.Errorf("row index out of range")
	}
	cells := alignedArr[r-1]
	if c < 1 || c > len(cells) {
		return "", fmt.Errorf("column index out of range")
	}
	return cells[c-1], nil
}
