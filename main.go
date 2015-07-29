// Copyright (c) 2015, Tamás Gulácsi
// Copyright (c) 2011-2013, Geoffrey Teale
//
// All rights reserved.
//
// Redistribution and use in source and binary forms, with or without
// modification, are permitted provided that the following conditions are met:
//
// 1. Redistributions of source code must retain the above copyright notice, this
//    list of conditions and the following disclaimer.
// 2. Redistributions in binary form must reproduce the above copyright notice,
//    this list of conditions and the following disclaimer in the documentation
//    and/or other materials provided with the distribution.
//
// THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
// ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
// WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
// DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR
// ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
// (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
// LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
// ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
// (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
// SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
//

package main

import (
	"encoding/csv"
	"errors"
	"flag"
	"fmt"
	"io"
	"log"
	"os"

	"github.com/tealeg/xlsx"
)

func generateCSVFromXLSXFile(w io.Writer, excelFileName string, sheetIndex int, csvOpts csvOptSetter) error {
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
	cw := csv.NewWriter(w)
	if csvOpts != nil {
		csvOpts(cw)
	}
	sheet := xlFile.Sheets[sheetIndex]
	var vals []string
	for _, row := range sheet.Rows {
		if row != nil {
			vals = vals[:0]
			for _, cell := range row.Cells {
				vals = append(vals, cell.String())
			}
			if err := cw.Write(vals); err != nil {
				return err
			}
		}
	}
	cw.Flush()
	return cw.Error()
}

type csvOptSetter func(*csv.Writer)

func main() {
	var (
		outFile    = flag.String("o", "-", "filename to output to. -=stdout")
		sheetIndex = flag.Int("i", 0, "Index of sheet to convert, zero based")
		delimiter  = flag.String("d", ";", "Delimiter to use between fields")
	)

	flag.Parse()
	if flag.NArg() < 1 {
		flag.PrintDefaults()
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

	if err := generateCSVFromXLSXFile(out, flag.Arg(0), *sheetIndex,
		func(cw *csv.Writer) { cw.Comma = ([]rune(*delimiter))[0] },
	); err != nil {
		log.Fatal(err)
	}
}
