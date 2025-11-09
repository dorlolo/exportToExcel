package exportToExcel

import (
    "testing"
)

type wDemo struct {
    Name  string `json:"name"`
    Age   int    `json:"age"`
    Extra string
}

func TestWriter_SupportedKinds(t *testing.T) {
    sw := sliceWriter{}
    if !sw.Supported([]wDemo{{Name: "a", Age: 1}}) {
        t.Fatalf("sliceWriter should support slice of struct")
    }
    if !sw.Supported([]*wDemo{{Name: "a", Age: 1}}) {
        t.Fatalf("sliceWriter should support slice of *struct")
    }
    if sw.Supported([]int{1, 2}) {
        t.Fatalf("sliceWriter should NOT support slice of non-struct")
    }

    stw := structWriter{}
    if !stw.Supported(wDemo{Name: "a", Age: 1}) {
        t.Fatalf("structWriter should support struct")
    }
    if !stw.Supported(&wDemo{Name: "a", Age: 1}) {
        t.Fatalf("structWriter should support *struct")
    }
    if stw.Supported(123) {
        t.Fatalf("structWriter should NOT support non-struct")
    }
}

func TestFillData_SliceValuesAndPointers(t *testing.T) {
    ex := NewExcel("./example", "writer_slice.xlsx")
    st := ex.NewSheet("ws", wDemo{})

    // values
    dataV := []wDemo{
        {Name: "A", Age: 1},
        {Name: "B", Age: 2},
    }
    if err := st.FillData(dataV); err != nil {
        t.Fatalf("FillData values error: %v", err)
    }
    if v, _ := ex.File().GetCellValue("ws", "A1"); v != "A" {
        t.Fatalf("A1=%s, want A", v)
    }
    if v, _ := ex.File().GetCellValue("ws", "B1"); v != "1" {
        t.Fatalf("B1=%s, want 1", v)
    }
    if v, _ := ex.File().GetCellValue("ws", "A2"); v != "B" {
        t.Fatalf("A2=%s, want B", v)
    }
    if v, _ := ex.File().GetCellValue("ws", "B2"); v != "2" {
        t.Fatalf("B2=%s, want 2", v)
    }

    // pointers
    st2 := ex.NewSheet("ws_ptr", wDemo{})
    dataP := []*wDemo{
        {Name: "C", Age: 3},
        {Name: "D", Age: 4},
    }
    if err := st2.FillData(dataP); err != nil {
        t.Fatalf("FillData pointers error: %v", err)
    }
    if v, _ := ex.File().GetCellValue("ws_ptr", "A1"); v != "C" {
        t.Fatalf("A1=%s, want C", v)
    }
    if v, _ := ex.File().GetCellValue("ws_ptr", "B2"); v != "4" {
        t.Fatalf("B2=%s, want 4", v)
    }
}

func TestFillData_EmptySlice(t *testing.T) {
    ex := NewExcel("./example", "writer_empty_slice.xlsx")
    st := ex.NewSheet("ws_empty", wDemo{})
    if err := st.FillData([]wDemo{}); err != nil {
        t.Fatalf("FillData empty slice error: %v", err)
    }
    if v, _ := ex.File().GetCellValue("ws_empty", "A1"); v != "" {
        t.Fatalf("A1=%s, want empty", v)
    }
}

func TestAutoResetCellWidth_LimitAndDisable(t *testing.T) {
    // limit width to 4
    ex := NewExcel("./example", "width_limit.xlsx")
    st := ex.NewSheet("wl", wDemo{}, OptionSetColWidth(2, 4))
    _ = st.FillData([]wDemo{{Name: "123456789", Age: 10}})
    if err := AutoResetCellWidth(st); err != nil {
        t.Fatalf("AutoResetCellWidth limit err: %v", err)
    }

    // disable min, long multi-byte string should be clipped by max limit
    ex2 := NewExcel("./example", "width_disable.xlsx")
    st2 := ex2.NewSheet("wd", wDemo{}, OptionSetColWidth(DisableLimitColWidth, 4))
    _ = st2.FillData([]wDemo{{Name: "你好世界ABCDE", Age: 20}})
    if err := AutoResetCellWidth(st2); err != nil {
        t.Fatalf("AutoResetCellWidth disable err: %v", err)
    }
}