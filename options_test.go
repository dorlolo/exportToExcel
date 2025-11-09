package exportToExcel

import (
    "testing"
    "github.com/xuri/excelize/v2"
)

func TestOptions_StyleAndColWidth(t *testing.T) {
    ex := NewExcel("./example", "options_test.xlsx")
    st := ex.NewSheet(
        "opt1",
        demoData{},
        OptionSetTitleStyle(func() *excelize.Style {
            s := DefaultTitleStyle()
            s.Font.Size = 8
            return s
        }),
        OptionSetDataStyle(func() *excelize.Style {
            s := DefaultDataStyle()
            s.Font.Size = 9
            return s
        }),
        OptionSetColWidth(5, 77),
    )

    ts := st.TitleStyle()()
    if ts.Font.Size != 8 {
        t.Fatalf("title style font size = %v, want 8", ts.Font.Size)
    }
    ds := st.DataStyle()()
    if ds.Font.Size != 9 {
        t.Fatalf("data style font size = %v, want 9", ds.Font.Size)
    }
    if st.MinColWidth() != 5 || st.MaxColWidth() != 77 {
        t.Fatalf("col width min/max = %v/%v, want 5/77", st.MinColWidth(), st.MaxColWidth())
    }
}