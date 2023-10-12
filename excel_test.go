package exportToExcel

import (
	"os"
	"testing"
)

func TestNewExcel(t *testing.T) {
	ex := NewExcel("./test", "template.xlsx")
	if err := ex.Save(); err != nil {
		t.Error(err)
	}
	_, err := os.Stat("./test/template.xlsx")
	if err != nil {
		t.Error(err)
	}
}

func TestNewExcelFromTemplate(t *testing.T) {
	ex, err := NewExcelFromTemplate("./test/template.xlsx", "./test", "exprot.xlsx")
	if err != nil {
		t.Error(err)
		return
	}
	if err = ex.Save(); err != nil {
		t.Error(err)
		return
	}
	_, err = os.Stat("./test/template.xlsx")
	if err != nil {
		t.Error(err)
	}
}

type demoData struct {
	Name string `json:"name" ex:"colHeader:名字"`
	Age  string `json:"age" ex:"colHeader:年龄"`
}

func TestExcel_NewSheet(t *testing.T) {
	ex := NewExcel("./test", "template.xlsx")
	defer func() {
		if err := ex.Save(); err != nil {
			t.Error(err)
		}
	}()
	ex.NewSheet("Sheet2", demoData{})
}
