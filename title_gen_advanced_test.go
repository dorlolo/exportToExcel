package exportToExcel

import (
    "testing"
)

func TestTitle_Gen_MergeCellsAndStyles(t *testing.T) {
    ex := NewExcel("./example", "title_gen.xlsx")
    st := ex.NewSheet("tgen", demoData{})

    // H at A1 merged to B1; S1 at A2; S2 at B2
    if err := st.Title.Gen(
        st.Title.NewTitleItem(4, "H", 1, 1).SetMergeTo(1, 2),
        st.Title.NewTitleItem(5, "S1", 2, 1),
        st.Title.NewTitleItem(5, "S2", 2, 2),
    ); err != nil {
        t.Fatalf("Title.Gen error: %v", err)
    }

    // Verify cell values
    if v, _ := ex.File().GetCellValue("tgen", "A1"); v != "H" {
        t.Fatalf("A1=%s, want H", v)
    }
    if v, _ := ex.File().GetCellValue("tgen", "A2"); v != "S1" {
        t.Fatalf("A2=%s, want S1", v)
    }
    if v, _ := ex.File().GetCellValue("tgen", "B2"); v != "S2" {
        t.Fatalf("B2=%s, want S2", v)
    }

    // Verify merge range exists A1:B1
    merges, err := ex.File().GetMergeCells("tgen")
    if err != nil {
        t.Fatalf("GetMergeCells error: %v", err)
    }
    found := false
    for _, mc := range merges {
        if mc.GetStartAxis() == "A1" && mc.GetEndAxis() == "B1" {
            found = true
            break
        }
    }
    if !found {
        t.Fatalf("expected merge A1:B1, got %+v", merges)
    }

    // Style applied (style id should be non-zero)
    if sid, err := ex.File().GetCellStyle("tgen", "A1"); err != nil || sid == 0 {
        t.Fatalf("GetCellStyle A1 failed: id=%d err=%v", sid, err)
    }
}

func TestTitle_Gen_InsertRowsAndVerticalMerge(t *testing.T) {
    ex := NewExcel("./example", "title_gen_rows.xlsx")
    st := ex.NewSheet("trows", demoData{})

    // Create a vertical merge from A2 to A3
    if err := st.Title.Gen(
        st.Title.NewTitleItem(5, "V", 2, 1).SetMergeRowNum(2),
    ); err != nil {
        t.Fatalf("Title.Gen vertical merge error: %v", err)
    }

    if v, _ := ex.File().GetCellValue("trows", "A2"); v != "V" {
        t.Fatalf("A2=%s, want V", v)
    }

    merges, err := ex.File().GetMergeCells("trows")
    if err != nil {
        t.Fatalf("GetMergeCells error: %v", err)
    }
    found := false
    for _, mc := range merges {
        if mc.GetStartAxis() == "A2" && mc.GetEndAxis() == "A3" {
            found = true
            break
        }
    }
    if !found {
        t.Fatalf("expected vertical merge A2:A3, got %+v", merges)
    }
}