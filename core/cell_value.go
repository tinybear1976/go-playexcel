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
