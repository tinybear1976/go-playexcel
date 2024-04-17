package core

import (
	"fmt"
)

// 缩小数组.给定一个已经对齐的数组(第二维长度一致)，从最末列开始，如果整列的值都为空字符串，则该列可以被删除
func ShrinkArray(alignedArray [][]string) ([][]string, error) {
	if len(alignedArray) == 0 {
		return alignedArray, nil
	}
	rowWidth := len(alignedArray[0])
	var minRightEmptyColumnCount int = rowWidth
	for rowIndex, row := range alignedArray {
		if len(row) != rowWidth {
			return nil, fmt.Errorf("数组未对齐,row: %d , len: %d", rowIndex, len(row))
		}
		var _rightEmptyColumnCount int = 0
		for colIndex := rowWidth - 1; colIndex >= 0; colIndex-- {
			if row[colIndex] == "" {
				_rightEmptyColumnCount++
			} else {
				break
			}
		}
		minRightEmptyColumnCount = min(minRightEmptyColumnCount, _rightEmptyColumnCount)
	}
	if minRightEmptyColumnCount == rowWidth {
		return [][]string{}, nil
	}
	if minRightEmptyColumnCount == 0 {
		return alignedArray, nil
	}
	rowWidth -= minRightEmptyColumnCount
	var shrinkedArray [][]string = make([][]string, len(alignedArray))
	for rowIndex, row := range alignedArray {
		cells := make([]string, rowWidth)
		for colIndex := 0; colIndex < rowWidth; colIndex++ {
			cells[colIndex] = row[colIndex]
		}
		shrinkedArray[rowIndex] = cells
	}
	return shrinkedArray, nil
}
