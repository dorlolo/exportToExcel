package main

import (
	"github.com/dorlolo/exportToExcel"
	"github.com/xuri/excelize/v2"
	"testing"
)

func TestLockCellExample(t *testing.T) {
	var err error
	var sheetName = "st1"
	ex := exportToExcel.NewExcel(".", "lockCellTest.xlsx")
	st1 := ex.NewSheet(sheetName, DemoBaseDataTypeA{})
	data := DemoBaseDataTypeA{
		Name:   "1212",
		Age:    11,
		Height: 22,
	}
	err = st1.FillData(data)
	if err != nil {
		t.Error(err)
		return
	}

	exf := ex.File()
	lockStyle := exportToExcel.DefaultDataStyle()
	lockStyle.Protection.Locked = true
	lockStyleId, err := exf.NewStyle(lockStyle)
	if err != nil {
		t.Error(err)
		return
	}
	err = exf.SetCellStyle(sheetName, "A1", "C2", lockStyleId)
	if err != nil {
		t.Error(err)
		return
	}
	err = exf.ProtectSheet(sheetName, &excelize.SheetProtectionOptions{
		Password:            "123456",
		AlgorithmName:       "SHA-512",
		SelectLockedCells:   true,
		SelectUnlockedCells: true,
	})
	if err != nil {
		t.Error(err)
		return
	}
    err = ex.Save()
    if err != nil {
        t.Error(err)
    }
}
