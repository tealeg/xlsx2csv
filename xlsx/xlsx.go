package xlsx

import (
	"encoding/csv"
	"fmt"
	x "github.com/tealeg/xlsx/v3"
	"io"
	"os"
	"regexp"
	"strings"
)


// File provides access to the xslx data as well as the path
type File struct {
	Path     string
	XLSXData *x.File
}

type csvOptSetter func(*csv.Writer)

// New provides an open xslx.File
func New(path string) (file *File, err error) {
	xlsxFile, err := x.OpenFile(path)
	if err != nil {
		return file, err
	}

	return &File{
		Path:     path,
		XLSXData: xlsxFile,
	}, nil
}

// SheetCount provides the number of sheets found in the workbook
func (f *File) SheetCount() int {
	return len(f.XLSXData.Sheets)
}

// GetOutFile opens a file for writing at outFilename
func GetOutFile(outFilename string, outFilepath string) (file *os.File, err error) {
	out := os.Stdout
	if !(outFilename == "" || outFilename == "-") {
		if outFilepath != "" {
			outFilepath = outFilepath + "/"
		}
		pathToFile := outFilepath + outFilename
		if out, err = os.Create(pathToFile); err != nil {
			return out, err
		}
	}
	return out, nil
}

// GenerateCSVFromSheet outputs a delimited file based on the data at the sheet index
func (f *File) GenerateCSVFromSheet(w io.Writer, index int, csvOpts csvOptSetter) error {
	if index >= f.SheetCount() {
		return fmt.Errorf(
			"No sheet %d available, please select a sheet between 0 and %d\n",
			index,
			f.SheetCount()-1,
		)
	}

	writer := csv.NewWriter(w)
	if csvOpts != nil {
		csvOpts(writer)
	}

	sheet := f.XLSXData.Sheets[index]
	var vals []string
	err := sheet.ForEachRow(func(row *x.Row) error {
		if row != nil {
			vals = vals[:0]
			err := row.ForEachCell(func(cell *x.Cell) error {
				str, err := cell.FormattedValue()
				if err != nil {
					return err
				}
				vals = append(vals, str)
				return nil
			})
			if err != nil {
				return err
			}
		}
		return writer.Write(vals)
	})

	if err != nil {
		return err
	}

	writer.Flush()
	return writer.Error()
}

// GenerateCSVsFromAllSheets iterates through all sheets and saves them as CSVs using their name
func (f *File) GenerateCSVsFromAllSheets(outFilepath string, csvOpts csvOptSetter) error {
	// Get sheet names
	keys := make([]string, 0, len(f.XLSXData.Sheet))
	for key := range f.XLSXData.Sheet {
		keys = append(keys, key)
	}

	// Keys come out backwards
	for i, j := 0, len(keys)-1; i < j; i, j = i+1, j-1 {
		keys[i], keys[j] = keys[j], keys[i]
	}

	for i := 0; i < f.SheetCount(); i++ {
		outFile, err := GetOutFile(keys[i]+".csv", outFilepath)
		if err != nil {
			return err
		}

		err = f.GenerateCSVFromSheet(outFile, i, csvOpts)
		if err != nil {
			return err
		}

		err = outFile.Close()
		if err != nil {
			return err
		}

	}
	return nil
}

// getSheetFilename takes the sheet name and standardizes the output filename
func getSheetFilename(sheetName string) string {
	name := strings.ToLower(sheetName)

	// Replace spaces with _
	r := regexp.MustCompile(` `)
	name = r.ReplaceAllString(name, "_")

	// Only alpha and (,.) will be left in filename
	reg := regexp.MustCompile("[^a-zA-Z0-9_.]+")
	name = reg.ReplaceAllString(name, "")

	return name + ".csv"
}
