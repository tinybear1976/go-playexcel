package core

import (
	"reflect"
	"testing"
)

func TestShrinkArray(t *testing.T) {
	tests := []struct {
		name string
		args [][]string
		want [][]string
	}{
		{"Case1",
			[][]string{{"abc", "", "def", ""}, {"ghi", "jkl", "", ""}, {"", "", "abc", ""}, {"", "", "", ""}},
			[][]string{{"abc", "", "def"}, {"ghi", "jkl", ""}, {"", "", "abc"}, {"", "", ""}},
		},
		{"Case2",
			[][]string{{"", ""}, {"", ""}}, [][]string{}},
		{"Case3",
			[][]string{{"abc", "def"}, {"ghi", "", "jkl"}},
			[][]string{}},
		{"Case4", [][]string{}, [][]string{}},
		{"Case5",
			[][]string{{"abc", "def", "", "", ""}, {"ghi", "", "", "", ""}, {"", "", "", "", ""}, {"", "", "", "", ""}, {"", "lmn", "", "", ""}},
			[][]string{{"abc", "def"}, {"ghi", ""}, {"", ""}, {"", ""}, {"", "lmn"}}},
	}
	for _, tt := range tests {
		got, err := ShrinkArray(tt.args)
		if err != nil {
			t.Logf("%s: error: %v", tt.name, err)
		} else {
			if !reflect.DeepEqual(got, tt.want) {
				t.Errorf("%s: want: %v, got: %v", tt.name, tt.want, got)
			}
		}
	}
}
