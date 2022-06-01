package main

import (
	"path/filepath"
	"testing"
)

func TestNormalizePath(t *testing.T) {

}

func TestWindowsPathToSheetName(t *testing.T) {
	data := make(map[string]pathTest, 2)
	data["./Resx/Translations.resx"] = pathTest{
		path:      "./Resx/Translations.resx",
		separator: '/',
		expected:  "Translations",
	}
	data["C:\\Users\\woeit\\Downloads\\Excel2Resource\\Resx\\Translations.resx"] = pathTest{
		path:      "C:\\Users\\woeit\\Downloads\\Excel2Resource\\Resx\\Translations.resx",
		separator: '\\',
		expected:  "Translations",
	}

	for k, v := range data {
		t.Logf("testing %v want %v (slashed %s)", k, v, filepath.ToSlash(k))
		baseDir, baseName, err := getPathInfo(k, v.separator)
		if err != nil {
			t.Fatal(err)
		}
		t.Logf("got basename %s baseDir %s", baseName, baseDir)
		if baseName != v.expected {
			t.Errorf("basename has unexpected result got %v want %v", baseName, v)
		}
	}
}

type pathTest struct {
	path      string
	separator rune
	expected  string
}
