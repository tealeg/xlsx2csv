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

type outputer func(s string)

func generateCSVFromXLSXFile(excelFileName string, sheetIndex int, outputf outputer) error {
	xlFile, error := xlsx.OpenFile(excelFileName)
	if error != nil {
		return error
	}
	sheetLen := len(xlFile.Sheets)
	switch {
	case sheetLen == 0:
		return errors.New("This XLSX file contains no sheets.")
	case sheetIndex >= sheetLen:
		return fmt.Errorf("No sheet %d available, please select a sheet between 0 and %d\n", sheetIndex, sheetLen-1)
	}
	sheet := xlFile.Sheets[sheetIndex]
	for _, row := range sheet.Rows {
		var vals []string
		if row != nil {
			for _, cell := range row.Cells {
                                str, err := cell.String()
				if err != nil {
					vals = append(vals, err.Error())
				}		
				vals = append(vals, fmt.Sprintf("%q", str))
			}
			outputf(strings.Join(vals, *delimiter) + "\n")
		}
	}
	return nil
}

func main() {
	flag.Parse()
	if len(os.Args) < 3 {
		flag.PrintDefaults()
		return
	}
	flag.Parse()
	printer := func(s string) { fmt.Printf("%s", s) }
	if err := generateCSVFromXLSXFile(*xlsxPath, *sheetIndex, printer); err != nil {
		fmt.Println(err)
	}
}
