package main

import (
	"errors"
	"flag"
	"fmt"
	"os"
	"strings"

	"github.com/tealeg/xlsx"
)

var xlsxPath = flag.String("f", "", "Path to an XLSX file")
var sheetIndex = flag.Int("i", 0, "Index of sheet to convert, zero based")
var delimiter = flag.String("d", ";", "Delimiter to use between fields")

type Outputer func(s string)

func generateCSVFromXLSXFile(excelFileName string, sheetIndex int, outputf Outputer) error {
	xlFile, error := xlsx.OpenFile(excelFileName)
	if error != nil {
		return error
	}
	sheetLen := len(xlFile.Sheets)
	switch {
	case sheetLen == 0:
		return errors.New("This XLSX file contains no sheets.")
	case sheetIndex >= sheetLen:
		return errors.New(fmt.Sprintf("No sheet %d available, please select a sheet between 0 and %d\n", sheetIndex, sheetLen-1))
	}
	sheet := xlFile.Sheets[sheetIndex]
	for _, row := range sheet.Rows {
		vals := make([]string, 0)
		if row != nil {
			for _, cell := range row.Cells {
				vals = append(vals, fmt.Sprintf("%q", cell.String()))
			}
			outputf(strings.Join(vals, *delimiter) + "\n")
		}
	}
	return nil
}

func usage() {
	fmt.Printf(`%s: -f=<XLSXFile> -i=<SheetIndex> -d=<Delimiter>

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
