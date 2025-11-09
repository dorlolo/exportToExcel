package exportToExcel

import "testing"

func TestTitle_SetFontSize(t *testing.T) {
    ex := NewExcel("./example", "title_font.xlsx")
    st := ex.NewSheet("t1", demoData{})
    tt := st.Title
    if s := tt.setFontSize(1); s != tt.MaxFontSize {
        t.Fatalf("level1 size = %v, want %v", s, tt.MaxFontSize)
    }
    if s := tt.setFontSize(tt.MaxLevel); s != tt.BaseFontSize {
        t.Fatalf("level%v size = %v, want %v", tt.MaxLevel, s, tt.BaseFontSize)
    }
    if s := tt.setFontSize(tt.MaxLevel+1); s != tt.BaseFontSize {
        t.Fatalf("overflow level size = %v, want %v", s, tt.BaseFontSize)
    }
    if s := tt.setFontSize(0); s <= tt.MaxFontSize {
        t.Fatalf("level0 size = %v, want > %v", s, tt.MaxFontSize)
    }
}

func TestTitleItem_SetMergeRowTo(t *testing.T) {
    ex := NewExcel("./example", "title_merge.xlsx")
    st := ex.NewSheet("t2", demoData{})
    item := st.Title.NewTitleItem(3, "x", 2, 1).SetMergeRowTo(5)
    if item.MergeRowTo != 5 {
        t.Fatalf("MergeRowTo = %d, want 5", item.MergeRowTo)
    }
}