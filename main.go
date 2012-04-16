package main

import (
	"flag"
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
)

var xlsxPath = flag.String("f", "", "Path to an XLSX file")
var sheetIndex = flag.Int("i", 0, "Index of sheet to convert, zero based")

type Outputer func(s string)

type XLSX2CSVError struct {
	error string
}

func (e XLSX2CSVError) Error() string {
	return e.error
}

func generateCSVFromXLSXFile(excelFileName string, sheetIndex int, outputf Outputer) error {
	var xlFile *xlsx.File
	var error error
	var sheetLen int
	var rowString string

	xlFile, error = xlsx.OpenFile(excelFileName)
	if error != nil {
		return error
	}
	sheetLen = len(xlFile.Sheets)
	switch {
	case sheetLen == 0:
		e := new(XLSX2CSVError)
		e.error = "This XLSX file contains no sheets."
		return (error)(*e)
	case sheetIndex >= sheetLen:
		e := new(XLSX2CSVError)
		e.error = fmt.Sprintf("No sheet %d available, please select a sheet between 0 and %d", sheetIndex, sheetLen-1)
		return (error)(*e)
	}
	sheet := xlFile.Sheets[sheetIndex]
	for _, row := range sheet.Rows {
		rowString = ""
		for cellIndex, cell := range row.Cells {
			if cellIndex > 0 {
				rowString = fmt.Sprintf("%s;\"%s\"", rowString, cell.String())
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
	var error error
	if len(os.Args) < 3 {
		usage()
		return
	}
	flag.Parse()
	error = generateCSVFromXLSXFile(*xlsxPath, *sheetIndex, func(s string) { fmt.Printf("%s", s) })
	if error != nil {
		fmt.Printf(error.Error())
		return
	}
}
