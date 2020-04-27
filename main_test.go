// Copyright 2011-2015, The xlsx2csv Authors.
// All rights reserved.
// For details, see the LICENSE file.

package main

import (
	"bytes"
	"encoding/csv"
	"strings"
	"testing"
)

func TestGenerateCSVFromXLSXFile(t *testing.T) {
	var testOutput bytes.Buffer

	for i, tc := range []struct {
		excelFileName string
		sheetIndex    int
		await         string
	}{
		{"testdata/testfile.xlsx", 0, `Foo;Bar
Baz ;Quuk
`},
		{"testdata/testfile2.xlsx", 0, `Bob;Alice;Sue
Yes;No;Yes
No;;Yes
`},
	} {
		testOutput.Reset()
		if err := generateCSVFromXLSXFile(
			&testOutput,
			tc.excelFileName,
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
