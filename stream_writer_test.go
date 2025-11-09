package exportToExcel

import (
    "testing"
)

type streamUser struct {
    Name string `json:"name"`
    Age  int    `json:"age"`
}

// Test streaming write with slice data
func TestStreamWriter_SliceWrite(t *testing.T) {
    ex := NewExcel("./example", "stream_slice.xlsx")
    st := ex.NewSheet("StreamSlice", streamUser{}, OptionEnableStreamWriter(true))
    ti := st.Title.NewTitleItem(1, "Users", 1, 1)
    ti.SetMergeColTo(2).SetMergeRowTo(1)
    if err := st.Title.Gen(ti); err != nil {
        t.Fatalf("title gen error: %v", err)
    }

    data := []streamUser{{Name: "Alice", Age: 30}, {Name: "Bob", Age: 25}}
    // capture start row before data fill
    start := GetFirstEmptyRowIndex(ex.File(), st.SheetName())
    if err := st.FillData(data); err != nil {
        t.Fatalf("FillData error: %v", err)
    }

    // verify values are written via stream path
    v1, _ := ex.File().GetCellValue(st.SheetName(), GetCellCoord(start+1, 1)) // first data row, col 1
    v2, _ := ex.File().GetCellValue(st.SheetName(), GetCellCoord(start+1, 2)) // first data row, col 2
    if v1 != "Alice" || v2 != "30" {
        t.Fatalf("unexpected first row values: %s, %s", v1, v2)
    }
    v3, _ := ex.File().GetCellValue(st.SheetName(), GetCellCoord(start+2, 1))
    v4, _ := ex.File().GetCellValue(st.SheetName(), GetCellCoord(start+2, 2))
    if v3 != "Bob" || v4 != "25" {
        t.Fatalf("unexpected second row values: %s, %s", v3, v4)
    }

    if err := ex.Save(); err != nil {
        t.Fatalf("save error: %v", err)
    }
}

// Test streaming write with single struct
func TestStreamWriter_StructWrite(t *testing.T) {
    ex := NewExcel("./example", "stream_struct.xlsx")
    st := ex.NewSheet("StreamStruct", streamUser{}, OptionEnableStreamWriter(true))
    ti := st.Title.NewTitleItem(1, "UserOne", 1, 1)
    ti.SetMergeColTo(2).SetMergeRowTo(1)
    if err := st.Title.Gen(ti); err != nil {
        t.Fatalf("title gen error: %v", err)
    }

    u := streamUser{Name: "Carol", Age: 44}
    start := GetFirstEmptyRowIndex(ex.File(), st.SheetName())
    if err := st.FillData(u); err != nil {
        t.Fatalf("FillData error: %v", err)
    }
    v1, _ := ex.File().GetCellValue(st.SheetName(), GetCellCoord(start+1, 1))
    v2, _ := ex.File().GetCellValue(st.SheetName(), GetCellCoord(start+1, 2))
    if v1 != "Carol" || v2 != "44" {
        t.Fatalf("unexpected single row values: %s, %s", v1, v2)
    }
    if err := ex.Save(); err != nil {
        t.Fatalf("save error: %v", err)
    }
}