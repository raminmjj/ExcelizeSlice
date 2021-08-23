package main

import "testing"

type columnNameTest struct {
    arg1 int 
	expected string
}

var columnNameTests = []columnNameTest{
    {0, "A"},
    {26, "AA"},
    {30, "AE"},
    {702, "AAA"},
}

func TestGenerateExcelColumnNameByIndex(t *testing.T)  {
	for _, test := range columnNameTests{
        if output := getColumnName(test.arg1); output != test.expected {
            t.Errorf("Output %q not equal to expected %q", output, test.expected)
        }
    }
}