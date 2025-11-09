package exportToExcel

import (
    "testing"
)

type widthUser struct {
    Name string `json:"name"`
    Note string `json:"note"`
}

// Verify fixed column width is applied when auto reset is disabled
func TestFixedColWidth_DisableAutoReset(t *testing.T) {
    ex := NewExcel("./example", "fixed_width.xlsx")
    st := ex.NewSheet("Fixed", widthUser{}, OptionEnableStreamWriter(true), OptionAutoResetColWidth(false), OptionSetColWidth(10, 10))

    // write data that would normally expand width significantly
    data := []widthUser{{Name: "short", Note: "this-is-a-very-long-text-to-expand-width"}}
    if err := st.FillData(data); err != nil {
        t.Fatalf("FillData error: %v", err)
    }
    if err := ex.Save(); err != nil {
        t.Fatalf("save error: %v", err)
    }

    // Check col width is near fixed value (min + ComfortColWidth)
    // excelize default width is ~8.43; we expect around 12.
    // Use column A and B
    wA, _ := ex.File().GetColWidth(st.SheetName(), "A")
    wB, _ := ex.File().GetColWidth(st.SheetName(), "B")
    // Expect widths to be equal and larger than default (~8.43)
    if wA < 9.0 {
        t.Fatalf("unexpected small col A width: %v", wA)
    }
    if wB < 9.0 {
        t.Fatalf("unexpected small col B width: %v", wB)
    }
    if wA != wB {
        t.Fatalf("widths should be equal, got A=%v B=%v", wA, wB)
    }
}