package exportToExcel

import (
	"github.com/xuri/excelize/v2"
)

func NewTitle(sheet *Sheet) Title {
	return Title{sheet: sheet, FontScaleFactor: 2, BaseFontSize: 12, MaxFontSize: 22, MaxLevel: 5}
}

type Title struct {
	sheet           *Sheet
	titles          []*TitleItem
	MaxLevel        int
	FontScaleFactor float64
	BaseFontSize    float64
	MaxFontSize     float64
}

func (t *Title) Add(title *TitleItem) (err error) {
	defer func() {
		if err == nil {
			t.titles = append(t.titles, title)
		}
	}()
	//insert row
	if len(t.titles) > 0 {
		var sameRowExisted bool
		for _, t := range t.titles {
			if t.Row == title.Row {
				sameRowExisted = true
			}
		}
		if !sameRowExisted {
			_ = t.sheet.file.InsertRows(t.sheet.SheetName(), title.Row, 1)
		}
	} else if t.sheet.rowNum > title.Row {
		_ = t.sheet.file.InsertRows(t.sheet.SheetName(), title.Row, 1)
	}
	if title.MergeColFrom == 0 {
		title.MergeColFrom = 1
	}
	var (
		cellL = GetCellCoord(title.Row, title.MergeColFrom)
		cellR = GetCellCoord(title.Row, title.MergeColTo)
	)
	//Merge cells
	if title.MergeColFrom != 0 && title.MergeColTo != 0 {
		if err = t.sheet.file.MergeCell(t.sheet.SheetName(), cellL, cellR); err != nil {
			return err
		}
	}
	//set text
	if err = t.sheet.file.SetCellStr(t.sheet.SheetName(), cellL, title.Text); err != nil {
		return err
	}
	//set Style
	if title.style == nil {
		title.style = t.sheet.defaultStyle()
		title.style.Font.Size = t.setFontSize(title.Level)
	}
	headerStyleID, errs := t.sheet.file.NewStyle(title.style)
	if errs != nil {
		return
	}
	if err = t.sheet.file.SetCellStyle(t.sheet.SheetName(), cellL, cellR, headerStyleID); err != nil {
		return
	}
	return nil
}
func (t *Title) setFontSize(titleLevel int) float64 {
	if titleLevel > t.MaxLevel || titleLevel < 0 {
		return t.BaseFontSize
	}
	fontSize := t.MaxFontSize - float64(titleLevel-1)*t.FontScaleFactor
	if fontSize < t.BaseFontSize {
		return t.BaseFontSize
	}
	return fontSize
}

type TitleItem struct {
	Level        int
	Text         string
	Row          int //所在行
	MergeColFrom int
	MergeColTo   int
	style        *excelize.Style
}

// SetStyle
// By default, the style is set automatically based on the title level.
func (t *TitleItem) SetStyle(style *excelize.Style) *TitleItem {
	t.style = style
	return t
}
func (t *TitleItem) NeedMergeAllDataCol(dataColLen int) *TitleItem {
	t.MergeColFrom = 1
	t.MergeColTo = dataColLen
	return t
}
