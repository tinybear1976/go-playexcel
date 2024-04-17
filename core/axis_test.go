package core

import "testing"

func TestUnpackAxisName(t *testing.T) {
	testCases := []struct {
		axisName string
		row      int
		col      int
	}{
		{"A1", 1, 1},
		{"Z30", 30, 26},
		{"AF2", 2, 32},
		{"Z100", 100, 26},
		{"AA1", 1, 27},
		{"AB54", 54, 28},
		{"ZZ1", 1, 702},
		{"A1A", 0, 0},
		{"1A", 0, 0},
		{"A11", 11, 1},
		{"A1.", 0, 0},
		{"A1A1", 0, 0},
		{"A111", 111, 1},
		{"A0", 0, 0},
	}

	for _, tc := range testCases {
		row, col, err := UnpackAxisName(tc.axisName)
		if err != nil {
			t.Logf("unpackAxisName(%v) = %v, %v; want %v, %v .return error: %v", tc.axisName, row, col, tc.row, tc.col, err)
		} else if row != tc.row || col != tc.col {
			t.Errorf("unpackAxisName(%v) = %v, %v; want %v, %v", tc.axisName, row, col, tc.row, tc.col)
		}
	}
}
