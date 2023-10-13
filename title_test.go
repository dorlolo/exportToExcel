package exportToExcel

import (
	"github.com/xuri/excelize/v2"
	"reflect"
	"testing"
)

func TestNewTitle(t *testing.T) {
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
	err = st.Title.Add(
		&TitleItem{Level: 1, Text: "h1-test", Row: 1, MergeColFrom: 0, MergeColTo: 6},
		&TitleItem{Level: 2, Text: "h2-test", Row: 2, MergeColFrom: 1, MergeColTo: 6},
		&TitleItem{Level: 3, Text: "h3-test", Row: 3, MergeColFrom: 1, MergeColTo: 6},
		&TitleItem{Level: 4, Text: "h4-test", Row: 4, MergeColFrom: 1, MergeColTo: 6},
		&TitleItem{Level: 4, Text: "h4-test", Row: 4, MergeColFrom: 1, MergeColTo: 6},
		&TitleItem{Level: 5, Text: "h5-1-test", Row: 5, MergeColFrom: 1, MergeColTo: 2},
		&TitleItem{Level: 5, Text: "h5-2-test", Row: 5, MergeColFrom: 3, MergeColTo: 6},
	)
	if err != nil {
		t.Error(err)
	}
}

func TestTitleItem_NeedMergeAllDataCol(t1 *testing.T) {
	type fields struct {
		Level        int
		Text         string
		Row          int
		MergeColFrom int
		MergeColTo   int
		style        *excelize.Style
	}
	type args struct {
		dataColLen int
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   *TitleItem
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t1.Run(tt.name, func(t1 *testing.T) {
			t := &TitleItem{
				Level:        tt.fields.Level,
				Text:         tt.fields.Text,
				Row:          tt.fields.Row,
				MergeColFrom: tt.fields.MergeColFrom,
				MergeColTo:   tt.fields.MergeColTo,
				style:        tt.fields.style,
			}
			if got := t.NeedMergeAllDataCol(tt.args.dataColLen); !reflect.DeepEqual(got, tt.want) {
				t1.Errorf("NeedMergeAllDataCol() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestTitleItem_SetStyle(t1 *testing.T) {
	type fields struct {
		Level        int
		Text         string
		Row          int
		MergeColFrom int
		MergeColTo   int
		style        *excelize.Style
	}
	type args struct {
		style *excelize.Style
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   *TitleItem
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t1.Run(tt.name, func(t1 *testing.T) {
			t := &TitleItem{
				Level:        tt.fields.Level,
				Text:         tt.fields.Text,
				Row:          tt.fields.Row,
				MergeColFrom: tt.fields.MergeColFrom,
				MergeColTo:   tt.fields.MergeColTo,
				style:        tt.fields.style,
			}
			if got := t.SetStyle(tt.args.style); !reflect.DeepEqual(got, tt.want) {
				t1.Errorf("SetStyle() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestTitle_Add(t1 *testing.T) {
	type fields struct {
		sheet           *Sheet
		titles          []*TitleItem
		MaxLevel        int
		FontScaleFactor float64
		BaseFontSize    float64
		MaxFontSize     float64
	}
	type args struct {
		title *TitleItem
	}
	tests := []struct {
		name    string
		fields  fields
		args    args
		wantErr bool
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t1.Run(tt.name, func(t1 *testing.T) {
			t := &Title{
				sheet:           tt.fields.sheet,
				titles:          tt.fields.titles,
				MaxLevel:        tt.fields.MaxLevel,
				FontScaleFactor: tt.fields.FontScaleFactor,
				BaseFontSize:    tt.fields.BaseFontSize,
				MaxFontSize:     tt.fields.MaxFontSize,
			}
			if err := t.Add(tt.args.title); (err != nil) != tt.wantErr {
				t1.Errorf("Add() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

func TestTitle_setFontSize(t1 *testing.T) {
	type fields struct {
		sheet           *Sheet
		titles          []*TitleItem
		MaxLevel        int
		FontScaleFactor float64
		BaseFontSize    float64
		MaxFontSize     float64
	}
	type args struct {
		titleLevel int
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   float64
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t1.Run(tt.name, func(t1 *testing.T) {
			t := &Title{
				sheet:           tt.fields.sheet,
				titles:          tt.fields.titles,
				MaxLevel:        tt.fields.MaxLevel,
				FontScaleFactor: tt.fields.FontScaleFactor,
				BaseFontSize:    tt.fields.BaseFontSize,
				MaxFontSize:     tt.fields.MaxFontSize,
			}
			if got := t.setFontSize(tt.args.titleLevel); got != tt.want {
				t1.Errorf("setFontSize() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestNewTitleItem(t *testing.T) {
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
	type args struct {
		level int
		text  string
		row   int
		style *excelize.Style
	}
	tests := []struct {
		name string
		args args
		want *TitleItem
	}{
		{
			name: "",
			args: args{},
			want: nil,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if err = st.Title.Add(NewTitleItem(tt.args.level, tt.args.text, tt.args.row, tt.args.style)); err != nil {
				t.Error(err)
				return
			}
		})
	}
}
