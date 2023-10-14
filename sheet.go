package exportToExcel

import (
	"errors"
	"github.com/xuri/excelize/v2"
	"reflect"
)

func newSheet(file *excelize.File, sheetName string, baseDataType any, opts ...Option) *Sheet {
	var a = &Sheet{file: file, baseDataType: baseDataType}
	a.Title = NewTitle(a)
	a.sheetId, _ = file.NewSheet(sheetName)
	a.defaultStyle = DefaultTitleStyle
	for _, opt := range opts {
		opt(a)
	}
	return a
}

type Sheet struct {
	Title         Title
	fieldSort     []string
	file          *excelize.File
	sheetId       int
	rowNum        int
	defaultStyle  func() *excelize.Style
	data          any
	baseDataType  any //the base type of the data, it used to search appropriate writer(IDataWriter.Supported)
	firstEmptyRow int
}

func (s *Sheet) SetSheetName(sheetName string) {
	_ = s.file.SetSheetName(s.file.GetSheetName(s.sheetId), sheetName)
}

func (s *Sheet) FillData(data any) error {
	s.data = data
	if s.file == nil {
		return errors.New("file object is Empty!")
	}
	dataType := reflect.TypeOf(data)
	switch dataType.Kind() {
	case reflect.Slice, reflect.Array:
		value := reflect.ValueOf(data)
		s.rowNum += value.Len()
	default:
		s.rowNum += 1
	}
	s.firstEmptyRow = GetFirstEmptyRowIndex(s.file, s.SheetName())
	return writers.WriteData(s)
}

func (s *Sheet) SetFieldSort(fieldSort []string) {
	s.fieldSort = fieldSort
}

func (s *Sheet) SheetName() string {
	return s.file.GetSheetName(s.sheetId)
}
func (s *Sheet) Data() any {
	return s.data
}
func (s *Sheet) DefaultStyle() func() *excelize.Style {
	return s.defaultStyle
}

func (s *Sheet) SetDefaultStyle(defaultStyle func() *excelize.Style) {
	s.defaultStyle = defaultStyle
}

// 数据将按照这些字段的顺序写入表中
func (s *Sheet) Fields(recalculate ...bool) []string {
	if (s.fieldSort == nil && s.baseDataType != nil) || (len(recalculate) == 1 && recalculate[0] == true) {
		return writers.FieldSort(s)
	}
	return s.fieldSort
}
