// Copyright 2011-2015, The xlsx2csv Authors.
// All rights reserved.
// For details, see the LICENSE file.

package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"log"
	"os"

	"github.com/dkoston/xlsx2csv/xlsx"
)

type csvOptSetter func(*csv.Writer)

func main() {
	var (
		outFilename = flag.String("o", "-", "filename to output to. -=stdout (default == STDOUT)")
		outFilepath = flag.String("p", "", "path to output to. Current directory if not set")
		sheetIndex  = flag.Int("i", 0, "Index of sheet to convert, zero based (default == 0")
		delimiter   = flag.String("d", ",", "Delimiter to use between fields (default == ,")
		allSheets   = flag.Bool("a", false, "Convert all sheets using the sheet name (lowercased with _ for spaces) as the output file name (default == false)")
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

	csvOpts := func(cw *csv.Writer) { cw.Comma = ([]rune(*delimiter))[0] }

	// Open our XSLX file
	file, err := xlsx.New(flag.Arg(0))
	if err != nil {
		log.Fatal(err)
	}

	if *allSheets {
		// Get the number of sheets from the XSLX file
		err = file.GenerateCSVsFromAllSheets(*outFilepath, csvOpts)
		if err != nil {
			log.Fatal(err)
		}
		os.Exit(0)
	}

	outFile, err := xlsx.GetOutFile(*outFilename, *outFilepath)
	if err != nil {
		log.Fatal(err)
	}

	err = file.GenerateCSVFromSheet(outFile, *sheetIndex, csvOpts)
	if err != nil {
		log.Fatal(err)
	}
	err = outFile.Close()
	if err != nil {
		log.Fatal(err)
	}
}
