package main

import (
	"testing"
)

func TestGenerateCSVFromXLSXFile(t *testing.T) {
	var sheetIndex int
	var excelFileName string
	var csv []string
	var index int
	var testOutputer = func(s string) {
		csv[index] = s
		index++
	}
	index = 0
	csv = make([]string, 2)
	sheetIndex = 0
	excelFileName = "testfile.xlsx"
	error := generateCSVFromXLSXFile(excelFileName, sheetIndex, testOutputer)
	if error != nil {
		t.Error(error.Error())
	}
	if len(csv) != 2 {
		t.Error("Expected len(csv) == 2")
	}
	rowString1 := csv[0]
	if rowString1 != "\"Foo\";\"Bar\"\n" {
		t.Error(`Expected rowString1 == "Foo";"Bar"\n but got `, rowString1)
	}
	rowString2 := csv[1]
	if rowString2 != "\"Baz \";\"Quuk\"\n" {
		t.Error(`Expected rowString2 == "Baz ";"Quuk"\n but got `, rowString2)
	}
}

func TestGenerateCSVFromXLSXFileWithEmptyCells(t *testing.T) {
	var sheetIndex int
	var excelFileName string
	var csv []string
	var csvlen int
	var index int
	var testOutputer = func(s string) {
		csv[index] = s
		index++
	}
	index = 0
	csv = make([]string, 3)
	sheetIndex = 0
	excelFileName = "testfile2.xlsx"
	error := generateCSVFromXLSXFile(excelFileName, sheetIndex, testOutputer)
	if error != nil {
		t.Error(error.Error())
	}
	csvlen = len(csv)
	if csvlen != 3 {
		t.Error("Expected len(csv) == 3, but got", csvlen)
	}
	rowString1 := csv[0]
	if rowString1 != "\"Bob\";\"Alice\";\"Sue\"\n" {
		t.Error(`Expected rowString1 == "Bob";"Alice";Sue"\n but got `, rowString1)
	}
	rowString2 := csv[1]
	if rowString2 != "\"Yes\";\"No\";\"Yes\"\n" {
		t.Error(`Expected rowString2 == "Yes";"No";"Yes"\n but got `, rowString2)
	}
	rowString3 := csv[2]
	if rowString3 != "\"No\";\"\";\"Yes\"\n" {
		t.Error(`Expected rowString2 == "No";"";"Yes"\n but got `, rowString3)
	}

}
