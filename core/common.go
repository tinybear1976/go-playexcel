package core

// maxWidth 获取二维数组中第二维的最大宽度（即每行上最大的列数）
func MaxWidth(matrix [][]string) int {
	maxWidth := 0
	for _, row := range matrix {
		maxWidth = max(len(row), maxWidth)
	}
	return maxWidth
}
