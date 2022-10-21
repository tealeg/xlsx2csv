// Copyright 2011-2015, The xlsx2csv Authors.
// All rights reserved.
// For details, see the LICENSE file.

package main

import (
	"encoding/csv"
	"errors"
	"flag"
	"fmt"
	"io"
	"log"
	"os"

	"github.com/tealeg/xlsx/v3"
)

func generateCSVFromXLSXFile(w io.Writer, excelFileName string, sheetIndex int, cols int, csvOpts csvOptSetter) error {
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		return err
	}
	sheetLen := len(xlFile.Sheets)
	switch {
	case sheetLen == 0:
		return errors.New("This XLSX file contains no sheets.")
	case sheetIndex >= sheetLen:
		return fmt.Errorf("No sheet %d available, please select a sheet between 0 and %d\n", sheetIndex, sheetLen-1)
	}
	cw := csv.NewWriter(w)
	if csvOpts != nil {
		csvOpts(cw)
	}
	sheet := xlFile.Sheets[sheetIndex]
	var vals []string
	isHeader := cols == 0
	err = sheet.ForEachRow(func(row *xlsx.Row) error {
		if row != nil {
			vals = vals[:0]
			col := 0
			err := row.ForEachCell(func(cell *xlsx.Cell) error {
				str, err := cell.FormattedValue()
				if err != nil {
					return err
				}
				if isHeader {
					if len(str) == 0 {
						return nil;
					}
					cols += 1
				} else if col >= cols {
					return nil;
				}
				col += 1
				vals = append(vals, str)
				return nil
			})
			if err != nil {
				return err
			}
		}
		isHeader = false
		if isEmpty(vals) {
			return nil
		}
		cw.Write(vals)
		return nil
	})
	if err != nil {
		return err
	}
	cw.Flush()
	return cw.Error()
}

func isEmpty(vals []string) bool {
	for _, v := range vals {
		if len(v) != 0 {
			return false
		}
	}
	return true
}

type csvOptSetter func(*csv.Writer)

func main() {
	var (
		outFile    = flag.String("o", "-", "filename to output to. -=stdout")
		sheetIndex = flag.Int("i", 0, "Index of sheet to convert, zero based")
		delimiter  = flag.String("d", ";", "Delimiter to use between fields")
		cols  = flag.Int("c", 0, "Number of columns to output. If not specified first row is considered as header")
	)
	flag.Usage = func() {
		fmt.Fprintf(os.Stderr, `%s
	dumps the given xlsx file's chosen sheet as a CSV,
	with the specified delimiter, into the specified output.

Usage:
	%s [flags] <xlsx-to-be-read>
`, os.Args[0], os.Args[0])
		flag.PrintDefaults()
	}

	flag.Parse()
	if flag.NArg() != 1 {
		flag.Usage()
		os.Exit(1)
	}
	out := os.Stdout
	if !(*outFile == "" || *outFile == "-") {
		var err error
		if out, err = os.Create(*outFile); err != nil {
			log.Fatal(err)
		}
	}
	defer func() {
		if closeErr := out.Close(); closeErr != nil {
			log.Fatal(closeErr)
		}
	}()

	if err := generateCSVFromXLSXFile(out, flag.Arg(0), *sheetIndex, *cols,
		func(cw *csv.Writer) { cw.Comma = ([]rune(*delimiter))[0] },
	); err != nil {
		log.Fatal(err)
	}
}
