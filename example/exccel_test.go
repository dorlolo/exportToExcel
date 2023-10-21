package main

import (
	"fmt"
	"github.com/dorlolo/exportToExcel"
	"testing"
)

type DemoBaseDataTypeA struct {
	Name   string  `json:"name"`
	Age    int     `json:"age"`
	Height float32 `json:"height"`
}

func TestNewExcelAndFillData(t *testing.T) {
	var err error
	ex := exportToExcel.NewExcel(".", "newfile.xlsx")
	defer func() {
		if err == nil {
			if err = ex.Save(); err != nil {
				fmt.Println("save file failed:", err)
				return
			}
		}
	}()
	// generate sheet-1,fill slice data
	st1 := ex.NewSheet("st1", DemoBaseDataTypeA{})
	err = st1.Title.Gen(
		st1.Title.NewTitleItem(4, "st1-demo", 1, 1).SetFullHorizontalMerge(),
		st1.Title.NewTitleItem(5, "name", 2, 1),
		st1.Title.NewTitleItem(5, "age", 2, 2),
		st1.Title.NewTitleItem(5, "height", 2, 3),
	)
	if err != nil {
		fmt.Println("generate title failed:", err.Error())
		return
	}
	var data1_1 = []DemoBaseDataTypeA{
		{"Mr.Zhang", 16, 180},
		{"Mrs.Li", 18, 220},
	}
	err = st1.FillData(data1_1)
	if err != nil {
		fmt.Println("fill data1_1 err:", err)
		return
	}
	var data1_2 = []*DemoBaseDataTypeA{
		{"Mr.Zhang", 16, 180},
		{"Mrs.Li", 18, 220},
	}
	err = st1.FillData(data1_2)
	if err != nil {
		fmt.Println("fill data1_2 err:", err)
		return
	}

	// generate sheet-2,fill struct data
	st2 := ex.NewSheet("st2", DemoBaseDataTypeA{})
	var data2 = DemoBaseDataTypeA{"Mr.SU", 100, 1800.6}
	err = st2.FillData(data2)
	if err != nil {
		fmt.Println("fill data2 use struct err:", err)
		return
	}
	// use pointer generate again
	err = st2.FillData(&data2)
	if err != nil {
		fmt.Println("fill data2 use pointer err:", err)
		return
	}
}
func TestReadExcelFromTemplateFile(t *testing.T) {
	ex, err := exportToExcel.NewExcelFromTemplate("./template.xlsx", ".", "newfile.xlsx")
	if err != nil {
		t.Error(err)
		return
	}
	st := ex.GetSheetByName("Sheet1")
	if st == nil {
		t.Error("can not file sheet:Sheet1")
		return
	}
	st.SetDataType(DemoBaseDataTypeA{})
	var data1 = []DemoBaseDataTypeA{
		{"Mr.Zhang", 16, 180},
		{"Mrs.Li", 18, 220},
	}
	if err = st.FillData(data1); err != nil {
		t.Error(err)
		return
	}
	// fill data continue
	var data2 = []DemoBaseDataTypeA{
		{"Mr.SU", 100, 1800.6},
		{"Mrs.Xu", 4, 22.0},
	}
	if err = st.FillData(data2); err != nil {
		t.Error(err)
		return
	}

	if err = ex.Save(); err != nil {
		t.Error(err)
		return
	}
}
