package exportToExcel

import (
    "reflect"
    "testing"
)

type utilDemo struct {
    Name string `json:"name"`
    Age  int    `json:"age"`
    Skip string
}

func TestGetColumnIndex(t *testing.T) {
    cases := map[int]string{
        1:  "A",
        26: "Z",
        27: "AA",
        52: "AZ",
        53: "BA",
        78: "BZ",
        701: "ZY",
        702: "ZZ",
        703: "AAA",
    }
    for n, want := range cases {
        got := GetColumnIndex(n)
        if got != want {
            t.Fatalf("GetColumnIndex(%d) = %s, want %s", n, got, want)
        }
    }
}

func TestGetCellCoord(t *testing.T) {
    if coord := GetCellCoord(0, 0); coord != "A1" {
        t.Fatalf("GetCellCoord(0,0) = %s, want A1", coord)
    }
    if coord := GetCellCoord(3, 2); coord != "B3" {
        t.Fatalf("GetCellCoord(3,2) = %s, want B3", coord)
    }
}

func TestDataToMapByJsonTag(t *testing.T) {
    d := utilDemo{Name: "n", Age: 18, Skip: "x"}
    m := DataToMapByJsonTag(reflect.ValueOf(d), reflect.TypeOf(d))
    if len(m) != 2 {
        t.Fatalf("map len = %d, want 2", len(m))
    }
    if m["name"].(string) != "n" || m["age"].(int) != 18 {
        t.Fatalf("unexpected values: %+v", m)
    }

    mp := DataToMapByJsonTag(reflect.ValueOf(&d), reflect.TypeOf(&d))
    if len(mp) != 2 {
        t.Fatalf("map len (ptr) = %d, want 2", len(mp))
    }
}

func TestGetJsonFieldList(t *testing.T) {
    list, err := GetJsonFieldList(reflect.TypeOf(utilDemo{}))
    if err != nil {
        t.Fatalf("unexpected err: %v", err)
    }
    if !reflect.DeepEqual(list, []string{"name", "age"}) {
        t.Fatalf("list = %v, want [name age]", list)
    }

    list2, err2 := GetJsonFieldList(reflect.TypeOf(&utilDemo{}))
    if err2 != nil {
        t.Fatalf("unexpected err: %v", err2)
    }
    if !reflect.DeepEqual(list2, []string{"name", "age"}) {
        t.Fatalf("list2 = %v, want [name age]", list2)
    }

    // invalid type: pointer to non-struct
    _, err3 := GetJsonFieldList(reflect.TypeOf(&[]int{}))
    if err3 == nil {
        t.Fatalf("expected error for non-struct pointer, got nil")
    }
}