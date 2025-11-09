package exportToExcel

import "testing"

func TestSheet_FieldsAutoAndManual(t *testing.T) {
    ex := NewExcel("./example", "sheet_fields.xlsx")
    st := ex.NewSheet("s1", demoData{})
    // auto calculate when fieldSort is nil and base type is set
    st.SetDataType(demoData{})
    auto := st.Fields()
    if len(auto) != 2 || auto[0] != "name" || auto[1] != "age" {
        t.Fatalf("auto fields = %v, want [name age]", auto)
    }

    // manual override
    st.SetFieldSort([]string{"age", "name"})
    manual := st.Fields()
    if len(manual) != 2 || manual[0] != "age" || manual[1] != "name" {
        t.Fatalf("manual fields = %v, want [age name]", manual)
    }

    // force recalculation ignores manual
    recalc := st.Fields(true)
    if len(recalc) != 2 || recalc[0] != "name" || recalc[1] != "age" {
        t.Fatalf("recalc fields = %v, want [name age]", recalc)
    }
}

func TestSheet_Rename(t *testing.T) {
    ex := NewExcel("./example", "sheet_rename.xlsx")
    st := ex.NewSheet("s1", demoData{})
    st.SetSheetName("renamed")
    if st.SheetName() != "renamed" {
        t.Fatalf("sheet name = %s, want renamed", st.SheetName())
    }
}