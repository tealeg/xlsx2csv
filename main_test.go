package main

import (
	"testing"
)




func TestGenerateCSVFromXLSXFile(t *testing.T) {
	var sheetIndex int
	var excelFileName string
	var csv []string
	var index int
	var testOutputer = func (s string) {
		csv[index] = s
		index++
	}
	index = 0
	csv = make([]string, 2)
	sheetIndex = 0
	excelFileName = "testfile.xlsx"
	error := generateCSVFromXLSXFile(excelFileName, sheetIndex, testOutputer)
	if error != nil {
		t.Error(error.String())
	}
	if len(csv) != 2 {
		t.Error("Expected len(csv) == 2")
	}
	rowString1 := csv[0]
	if rowString1 != "\"Foo\",\"Bar\"\n" {
		t.Error(`Expected rowString1 == "Foo","Bar"\n but got `, rowString1)
	}
	rowString2 := csv[1]
	if rowString2 != "\"Baz \",\"Quuk\"\n" {
		t.Error(`Expected rowString2 == "Baz ","Quuk"\n but got `, rowString2)
	}
}