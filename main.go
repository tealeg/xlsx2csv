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
var addMissing = flag.Bool("a", false, "Add blank string as missing tail columns values")

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

	maxRowLen := 0
	// get max columns
	for _, row := range sheet.Rows {
		rowLen := len(row.Cells)
		if rowLen > maxRowLen {
			maxRowLen = rowLen
		}
	}

	for _, row := range sheet.Rows {
		var vals []string
		if row != nil {
			for _, cell := range row.Cells {
				str, err := cell.FormattedValue()
				if err != nil {
					vals = append(vals, err.Error())
				}
				vals = append(vals, fmt.Sprintf("%q", str))
			}

			// fix missing columns
			if *addMissing {
				rowLen := len(row.Cells)
				missingColumns := maxRowLen - rowLen
				for i := 1; i <= missingColumns; i++ {
					vals = append(vals, fmt.Sprintf("%q", ""))
				}
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
