// Copyright (c) 2015, Geoffrey Teale
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
		if err := generateCSVFromXLSXFile(&testOutput,
			tc.excelFileName, tc.sheetIndex,
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
