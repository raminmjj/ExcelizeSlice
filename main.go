package main

import (
	"errors"
	"fmt"
	"reflect"
	"time"

	"github.com/xuri/excelize/v2"
)

type sampleStruct struct {
	Id        int
	Feild1    string
	Feild2    float64
	Authors   []string
	Birthdate time.Time
}

func main() {
	timeValue := time.Now()
	structArray := []sampleStruct{
		{1, "string val1", 2.35, []string{"Ramin", "Mahdieh", "Bahar"}, timeValue},
		{2, "string val2", 3, []string{"Rahman", "Arash", "Reza"}, timeValue},
		{3, "string val3", 4.78, []string{"Michael", "John", "Chris"}, timeValue.Add(time.Minute * time.Duration(2))},
		{4, "string val4", 1.56, []string{"Bill", "Steve", "Scott"}, timeValue.Add(time.Minute * time.Duration(4))},
		{5, "string val5", 10.20, []string{"Max", "Bob", "Peter"}, timeValue.Add(time.Minute * time.Duration(6))},
	}
	localize := map[string]string{
		"Id":        "شناسه",
		"Feild1":    "فیلد یک",
		"Feild2":    "فیلد دو",
		"Authors":   "نویسندگان",
		"Birthdate": "تاریخ تولد",
	}
	Export2excel(structArray, localize, "structArray.xlsx")
	Export2excel(nil, nil, "structArray2.xlsx")
}

func Export2excel(data interface{}, localize map[string]string, filename string) error {
	items := reflect.ValueOf(data)
	if items.Kind() != reflect.Slice {
		return errors.New("THE DATA SOURCE MUST BE SLICE")
	}
	if items.Len() == 0 {
		return errors.New("THE DATA SOURCE IS NIL")
	}
	if items.Index(0).Kind() != reflect.Struct {
		return errors.New("THE SLICE ITEM IS NOT A STRUCT")
	}
	fieldsCount := items.Index(0).NumField()
	col := items.Index(0)
	f := excelize.NewFile()
	// Create a new sheet.
	index := f.NewSheet("Sheet1")
	// Set Header of a columns.
	for i := 0; i < fieldsCount; i++ {
		columnName := getColumnName(i)
		columnCaption := col.Type().Field(i).Name
		if localize != nil {
			caption := localize[columnCaption]
			if caption != "" {
				columnCaption = caption
			}
		}
		f.SetCellValue("Sheet1", columnName+"1", columnCaption)
	}
	// Set value of a cell.
	for i := 0; i < items.Len(); i++ {
		item := items.Index(i)
		if item.Kind() == reflect.Struct {
			row := reflect.Indirect(item)
			for j := 0; j < row.NumField(); j++ {
				columnName := getColumnName(j)
				if row.CanInterface() {
					f.SetCellValue("Sheet1", columnName+fmt.Sprint((i+2)), row.Field(j).Interface())
				} else {
					f.SetCellValue("Sheet1", columnName+fmt.Sprint((i+2)), row.Field(j))
				}
			}
		}
	}
	// Set active sheet of the workbook.
	f.SetActiveSheet(index)
	// Save spreadsheet by the given path.
	err := f.SaveAs(filename)
	return err
}

func getColumnName(colIndex int) string {
	ret := ""
	for i := colIndex; i >= 0; i-- {
		ret = string(rune(i%26+65)) + ret
		i = i / 26
	}
	return ret
}
