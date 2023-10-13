package exportToExcel

import (
	"testing"
)

func TestNewSheet(t *testing.T) {
	var err error
	ex := NewExcel("./test", "sheet_test.xlsx")
	defer func() {
		if err == nil {
			if err = ex.Save(); err != nil {
				t.Error(err)
			}
		}
	}()
	st := NewSheet(ex.file, "filldata", demoData{})
	data := []demoData{
		{
			Name: "small ming",
			Age:  12,
		},
		{
			Name: "big dams",
			Age:  14,
		},
	}
	err = st.FillData(data)
	if err != nil {
		t.Error(err)
	}
}
