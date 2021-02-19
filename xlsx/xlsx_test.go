package xlsx

import (
	"bytes"
	"encoding/csv"
	"strings"
	"testing"
)

func TestNew(t *testing.T) {
	// Valid file should not error
	_, err := New("../testdata/testfile.xlsx")
	if err != nil {
		t.Errorf("unexpected error: %v", err)
	}

	// Invalid file should error
	_, err = New("../testdata/nonexistentfile.xlsx")
	if err == nil {
		t.Error(err)
	}
}

func TestFile_SheetCount(t *testing.T) {
	// testdata/testfile3.xlsx has 1 sheet
	filename3 := "../testdata/testfile3.xlsx"
	expectedCount3 := 1
	file3, err := New(filename3)
	if err != nil {
		t.Errorf("unexpected error: %v", err)
	}

	count3 := file3.SheetCount()
	if count3 != expectedCount3 {
		t.Errorf("expected %d sheet, got %d. %s", expectedCount3, count3, filename3)
	}


	// testdata/testfile.xlsx has 3 sheets
	filename := "../testdata/testfile.xlsx"
	expectedCount := 3
	file, err := New(filename)
	if err != nil {
		t.Errorf("unexpected error: %v", err)
	}

	count := file.SheetCount()
	if count != expectedCount {
		t.Errorf("expected %d sheet, got %d. %s", expectedCount, count, filename)
	}
}

func Test_GenerateCSVFromSheet(t *testing.T) {
	var testOutput bytes.Buffer

	for i, tc := range []struct {
		excelFileName string
		sheetIndex    int
		await         string
	}{
		{"../testdata/testfile.xlsx", 0, `Foo;Bar
Baz ;Quuk
`},
		{"../testdata/testfile2.xlsx", 0, `Bob;Alice;Sue
Yes;No;Yes
No;;Yes
`},
	} {
		testOutput.Reset()

		file, err := New(tc.excelFileName)
		if err != nil {
			t.Errorf("unexpected error: %v", err)
		}

		if err := file.GenerateCSVFromSheet(
			&testOutput,
			tc.sheetIndex,
			func(cw *csv.Writer) { cw.Comma = ';' },
		); err != nil {
			t.Error(err)
		}
		awaited := strings.Split(tc.await, "\n")
		got := strings.Split(testOutput.String(), "\n")
		if len(got) != len(awaited) {
			t.Errorf("%d. Expected len(csv) == %d, got %d.", i, len(awaited), len(got))
			continue
		}
		for j, aw := range awaited {
			if aw != got[j] {
				t.Errorf(`%d. Expected line %d == %q, but got %q`, i, j, aw, got[j])
			}
		}
	}
}

func Test_getSheetFilename(t *testing.T) {
	type testCase struct {
		sheetName string
		fileName string
	}

	testCases := []testCase{
		{
			sheetName: "[1]Speci@lCh$$rs Time",
			fileName: "1specilchrs_time.csv",
		},
		{
			sheetName: "UG.Industries",
			fileName: "ug.industries.csv",
		},
		{
			sheetName: "Multiple Spaces Here!",
			fileName: "multiple_spaces_here.csv",
		},
	}

	for i := 0; i < len(testCases); i++ {
		name := getSheetFilename(testCases[i].sheetName)
		if name != testCases[i].fileName {
			t.Errorf("expected: %s. got %s. test case: %d", testCases[i].fileName, name, i)
		}
	}
}
