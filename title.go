package exportToExcel

import (
	"github.com/xuri/excelize/v2"
	"reflect"
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
	occupiedRow     []int
}

func (t *Title) Gen(titles ...*TitleItem) (err error) {
	for _, title := range titles {
		//insert row
		if len(t.occupiedRow) > 0 {
			for thisRow := title.InitRow; thisRow < title.MergeRowTo; thisRow++ {
				var sameRowExisted bool
				for _, occRow := range t.occupiedRow {
					if occRow == thisRow {
						sameRowExisted = true
						break
					}
				}
				if !sameRowExisted {
					t.sheet.rowNum++
					_ = t.sheet.file.InsertRows(t.sheet.SheetName(), title.InitRow, 1)
					t.occupiedRow = append(t.occupiedRow, thisRow)
				}
			}
		} else if t.sheet.rowNum > title.InitRow {
			_ = t.sheet.file.InsertRows(t.sheet.SheetName(), title.InitRow, title.MergeRowTo-title.MergeColTo+1)
		}
		if title.InitCol == 0 {
			title.InitCol = 1
		}
		if title.InitRow == 0 {
			title.InitRow = 1
		}
		var (
			cellL = GetCellCoord(title.InitRow, title.InitCol)
			cellR = GetCellCoord(title.MergeRowTo, title.MergeColTo)
		)
		//Merge cells
		if title.MergeRowTo != 0 && title.MergeColTo != 0 {
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
		t.titles = append(t.titles, title)
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

// NewTitleItem
func (t *Title) NewTitleItem(level int, text string, row, column int, style ...*excelize.Style) *TitleItem {
	tm := &TitleItem{
		Level:   level,
		Text:    text,
		InitRow: row,
		InitCol: column,
		sheet:   t.sheet,
	}
	if style != nil {
		tm.style = style[0]
	}
	t.titles = append(t.titles, tm)
	tm.id = len(t.titles) - 1
	return tm
}

type TitleItem struct {
	id         int
	Level      int
	Text       string
	InitRow    int //所在行
	InitCol    int //default is 1
	MergeColTo int
	MergeRowTo int
	style      *excelize.Style
	sheet      *Sheet
}

// GetId
func (t *TitleItem) GetId() int {
	return t.id
}

// SetStyle
// By default, the style is set automatically based on the title level.
func (t *TitleItem) SetStyle(style *excelize.Style) *TitleItem {
	t.style = style
	return t
}

// SetMergeColNum
func (t *TitleItem) SetMergeColNum(num int) *TitleItem {
	t.MergeColTo += num
	return t
}

// SetMergeRowNum
func (t *TitleItem) SetMergeRowNum(num int) *TitleItem {
	t.MergeRowTo += num
	return t
}

// SetMergeColTo
func (t *TitleItem) SetMergeColTo(colTo int) *TitleItem {
	t.MergeColTo = colTo
	return t
}

// SetMergeRowTo
func (t *TitleItem) SetMergeRowTo(colTo int) *TitleItem {
	t.MergeColTo = colTo
	return t
}

// SetMergeRowTo
func (t *TitleItem) SetMergeTo(row, col int) *TitleItem {
	t.MergeRowTo = row
	t.MergeColTo = col
	return t
}

// SetFullMerge
func (t *TitleItem) SetFullMerge(colNum ...int) *TitleItem {
	t.InitCol = 1
	if colNum != nil {
		t.MergeColTo = colNum[0]
	} else {
		baseType := reflect.TypeOf(t.sheet.baseDataType)
		t.MergeColTo = baseType.NumField()
		if colNum != nil && colNum[0] > t.MergeColTo {
			t.MergeColTo = colNum[0]
		}
	}
	return t
}
