package main

import (
	"flag"
	"fmt"
	"os"
	"github.com/tealeg/xlsx"
)

var xlsxPath = flag.String("f", "", "Path to an XLSX file")
var sheetIndex = flag.Int("i", 0, "Index of sheet to convert, zero based")


type Outputer func(s string) 

type XLSX2CSVError struct {
	error string
}

func (e XLSX2CSVError) String() string {
	return e.error
}

func generateCSVFromXLSXFile(excelFileName string, sheetIndex int, outputf Outputer) os.Error {
	var xlFile *xlsx.File
	var error os.Error
	var sheetLen int
	var rowString string

	fmt.Printf("%v\n", excelFileName)
	xlFile, error = xlsx.OpenFile(excelFileName)
	if error != nil {
		return error
	}
	sheetLen = len(xlFile.Sheets)
	switch {
	case sheetLen == 0:
		e := new(XLSX2CSVError)
		e.error = "This XLSX file contains no sheets."
		return (os.Error)(*e)
	case sheetIndex >= sheetLen:
		e := new(XLSX2CSVError)
		e.error = fmt.Sprintf("No sheet %d available, please select a sheet between 0 and %d", sheetIndex, sheetLen - 1)
		return (os.Error)(*e)
	}
	sheet := xlFile.Sheets[sheetIndex]
	for _, row := range sheet.Rows {
		rowString = ""
		for cellIndex, cell := range row.Cells {
			if cellIndex > 0 {
				rowString = fmt.Sprintf("%s,\"%s\"", rowString, cell.String())
			} else {
				rowString = fmt.Sprintf("\"%s\"", cell.String())
			}
		}
		rowString = fmt.Sprintf("%s\n", rowString)
		outputf(rowString)
	}
	return nil
}

func usage() {
	fmt.Printf(`%s: <XLSXFile> <SheetIndex>

Note: SheetIndex should be a number, zero based
`, 
		os.Args[0])
}


func main() {
	flag.Parse()
	var error os.Error
	if len(os.Args) < 3 {
		usage()
		return
	}
	flag.Parse()
	error = generateCSVFromXLSXFile(*xlsxPath, *sheetIndex, func (s string) {fmt.Printf("%s\n", s)})
	if error != nil {
		fmt.Printf(error.String())
		return
	}
}