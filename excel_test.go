package exportToExcel

import (
	"os"
	"testing"
)

func TestNewExcel(t *testing.T) {
	ex := NewExcel("./example", "template.xlsx")
	if err := ex.Save(); err != nil {
		t.Error(err)
	}
	_, err := os.Stat("./example/template.xlsx")
	if err != nil {
		t.Error(err)
	}
}

func TestNewExcelFromTemplate(t *testing.T) {
    ex, err := NewExcelFromTemplate("./example/template.xlsx", "./example", "exprot.xlsx")
    if err != nil {
        t.Error(err)
        return
    }
    if err = ex.Save(); err != nil {
        t.Error(err)
        return
    }
    _, err = os.Stat("./example/exprot.xlsx")
    if err != nil {
        t.Error(err)
    }
}

type demoData struct {
	Name string `json:"name" ex:"colHeader:名字"`
	Age  uint   `json:"age" ex:"colHeader:年龄"`
}

func TestExcel_NewSheet(t *testing.T) {
	ex := NewExcel("./example", "template.xlsx")
	defer func() {
		if err := ex.Save(); err != nil {
			t.Error(err)
		}
	}()
	ex.NewSheet("Sheet2", demoData{})
}
