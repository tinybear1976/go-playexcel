package core

import (
	"errors"

	"github.com/xuri/excelize/v2"
)

// 判断单元格是否在合并范围内
func IsWithInMergeScope(col, row int, mergeCells []excelize.MergeCell) (bool, string, error) {
	if col <= 0 || row <= 0 {
		return false, "", errors.New("col or row must be greater than 0")
	}

	// 判断是否在合并范围内
	for _, mc := range mergeCells {
		sRow, sCol, err := UnpackAxisName(mc.GetStartAxis())
		if err != nil {
			return false, "", err
		}
		eRow, eCol, err := UnpackAxisName(mc.GetEndAxis())
		if err != nil {
			return false, "", err
		}
		if col >= sCol && col <= eCol && row >= sRow && row <= eRow {
			return true, mc.GetCellValue(), nil
		}
	}
	return false, "", nil
}
