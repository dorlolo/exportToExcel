package exportToExcel

import (
	"github.com/xuri/excelize/v2"
	"reflect"
	"testing"
)

func TestTitle_Gen_TitleItem(t *testing.T) {
	var err error
	ex := NewExcel("./test", "title_test.xlsx")
	defer func() {
		if err == nil {
			if err = ex.Save(); err != nil {
				t.Error(err)
			}
		}
	}()
	st := ex.NewSheet("title1", demoData{})
	err = st.Title.Gen(
		&TitleItem{Level: 1, Text: "h1-test", InitRow: 1, InitCol: 1, MergeColTo: 6},
		&TitleItem{Level: 2, Text: "h2-test", InitRow: 2, InitCol: 1, MergeColTo: 6},
		&TitleItem{Level: 3, Text: "h3-test", InitRow: 3, InitCol: 1, MergeColTo: 6},
		&TitleItem{Level: 4, Text: "h4-test", InitRow: 4, InitCol: 1, MergeColTo: 6},
		&TitleItem{Level: 4, Text: "h4-test", InitRow: 4, InitCol: 1, MergeColTo: 6},
		&TitleItem{Level: 5, Text: "h5-1-test", InitRow: 5, InitCol: 1, MergeColTo: 2},
		&TitleItem{Level: 5, Text: "h5-2-test", InitRow: 5, InitCol: 3, MergeColTo: 6},
	)
	if err != nil {
		t.Error(err)
	}
}
func TestTitle_NewTitleItem(t *testing.T) {
	ex := NewExcel("./test", "title_test.xlsx")
	st := ex.NewSheet("title1", demoData{})
	type fields struct {
		level            int
		text             string
		initRow, initCol int
		style            *excelize.Style
	}
	tests := []struct {
		name string
		args fields
		want *TitleItem
	}{
		{
			name: "test1",
			args: fields{1, "test1", 1, 1, nil},
			want: &TitleItem{
				id: 0, Level: 1, Text: "test1", InitRow: 1, InitCol: 1, MergeColTo: 1, MergeRowTo: 1, style: nil, sheet: st,
			},
		},
		{
			name: "test2",
			args: fields{0, "test2", 0, 0, nil},
			want: &TitleItem{
				id: 1, Level: 5, Text: "test2", InitRow: 1, InitCol: 1, MergeColTo: 1, MergeRowTo: 1, style: nil, sheet: st,
			},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := st.Title.NewTitleItem(tt.args.level, tt.args.text, tt.args.initRow, tt.args.initCol, tt.args.style); !reflect.DeepEqual(got, tt.want) {
				t.Errorf("NewTitleItem() = %v, want %v", got, tt.want)
			}
		})
	}
}
func TestTitle_Gen_NewTitleItem(t *testing.T) {
	var err error
	ex := NewExcel("./test", "title_test.xlsx")
	defer func() {
		if err == nil {
			if err = ex.Save(); err != nil {
				t.Error(err)
			}
		}
	}()
	style1 := DefaultTitleStyle()
	style1.Font.Size = 6
	st := ex.NewSheet("title1", demoData{})
	tests := []struct {
		name string
		args *TitleItem
	}{
		{
			name: "NewTitleItem",
			args: st.Title.NewTitleItem(5, "NewTitleItem", 1, 1),
		},
		{
			name: "NewTitleItem_SetMergeColNum",
			args: st.Title.NewTitleItem(5, "NewTitleItem_SetMergeColNum", 2, 1).SetMergeColNum(2),
		},
		{
			name: "NewTitleItem_SetMergeRowNum",
			args: st.Title.NewTitleItem(5, "NewTitleItem_SetMergeRowNum", 3, 1).SetMergeRowNum(2),
		},
		{
			name: "NewTitleItem_SetMergeTo",
			args: st.Title.NewTitleItem(5, "NewTitleItem_SetMergeTo", 5, 1).SetMergeTo(6, 2),
		},
		{
			name: "NewTitleItem_SetFullHorizontalMerge",
			args: st.Title.NewTitleItem(5, "NewTitleItem_SetFullHorizontalMerge", 7, 1).SetFullHorizontalMerge(),
		},
		{
			name: "NewTitleItem_SetFullHorizontalMerge_defaultColNum:6",
			args: st.Title.NewTitleItem(5, "NewTitleItem_SetFullHorizontalMerge_defaultColNum:6", 8, 1).SetFullHorizontalMerge(6),
		},
		{
			name: "NewTitleItem_SetFullHorizontalMerge_SetStyle",
			args: st.Title.NewTitleItem(5, "NewTitleItem_SetFullHorizontalMerge_SetStyle:font size = 6", 9, 1).SetStyle(style1),
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if err = st.Title.Gen(tt.args); err != nil {
				t.Error(err)
				return
			}
		})
	}
}
