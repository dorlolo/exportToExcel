package exportToExcel

import (
    "os"
    "testing"
    "github.com/xuri/excelize/v2"
)

func TestExcel_SheetsAndProps(t *testing.T) {
    ex := NewExcel("./example", "excel_additional.xlsx")
    st1 := ex.NewSheet("A", demoData{})
    _ = ex.NewSheet("B", demoData{})

    // Get by name
    if got := ex.GetSheetByName("A"); got == nil || got != st1 {
        t.Fatalf("GetSheetByName A failed: %v", got)
    }

    // Get by id
    if got := ex.GetSheetById(st1.sheetId); got == nil || got != st1 {
        t.Fatalf("GetSheetById failed: %v", got)
    }

    // Set and get properties
    ex.SetFileName("excel_additional_saved.xlsx").SetFileDir("./example")
    if ex.FileName() != "excel_additional_saved.xlsx" || ex.FileDir() != "./example" {
        t.Fatalf("file props mismatch: %s / %s", ex.FileName(), ex.FileDir())
    }
}

func TestExcel_SaveDeleteSheet1AndCreatesFile(t *testing.T) {
    ex := NewExcel("./example", "excel_save.xlsx")
    _ = ex.NewSheet("OnlySheet", demoData{})
    // Save should delete built-in Sheet1 when not present in e.sheets
    if err := ex.Save(); err != nil {
        t.Fatalf("Save error: %v", err)
    }
    // Ensure file exists
    if _, err := os.Stat("./example/excel_save.xlsx"); err != nil {
        t.Fatalf("saved file not found: %v", err)
    }
    // Ensure Sheet1 not present
    sheets := ex.File().GetSheetList()
    for _, n := range sheets {
        if n == "Sheet1" {
            t.Fatalf("Sheet1 should be deleted, sheets: %v", sheets)
        }
    }
}

func TestExcel_NewStyle(t *testing.T) {
    ex := NewExcel("./example", "excel_style.xlsx")
    style := &excelize.Style{Font: &excelize.Font{Bold: true, Size: 10}}
    id, err := ex.NewStyle(style)
    if err != nil || id == 0 {
        t.Fatalf("NewStyle failed: id=%d err=%v", id, err)
    }
}