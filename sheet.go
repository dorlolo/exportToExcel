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
    a.titleStyle = DefaultTitleStyle
    a.dataStyle = DefaultDataStyle
    a.minColWidth = DefaultColMinWidth
    a.maxColWidth = DefaultColMaxWidth
    a.autoResetColWidth = true
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
    titleStyle    func() *excelize.Style
    dataStyle     func() *excelize.Style
    data          any
    baseDataType  any //the base type of the data, it used to search appropriate writer(IDataWriter.Supported)
    firstEmptyRow int
    minColWidth   float64
    maxColWidth   float64
    autoResetColWidth bool
    // streaming write support
    useStream     bool
    stream        *excelize.StreamWriter
}

func (s *Sheet) SetSheetName(sheetName string) {
	_ = s.file.SetSheetName(s.file.GetSheetName(s.sheetId), sheetName)
}

func (s *Sheet) FillData(data any) error {
	s.data = data
	if s.file == nil {
		return errors.New("file object is Empty!")
	}
	s.firstEmptyRow = GetFirstEmptyRowIndex(s.file, s.SheetName())
	s.rowNum = s.firstEmptyRow
	dataType := reflect.TypeOf(data)
	switch dataType.Kind() {
	case reflect.Slice, reflect.Array:
		value := reflect.ValueOf(data)
		s.rowNum += value.Len()
	default:
		s.rowNum += 1
	}
	return writers.WriteData(s)
}
func (s *Sheet) SetDataType(t any) {
	s.baseDataType = t
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
func (s *Sheet) TitleStyle() func() *excelize.Style {
	return s.titleStyle
}
func (s *Sheet) DataStyle() func() *excelize.Style {
	return s.dataStyle
}
func (s *Sheet) SetTitleStyle(style func() *excelize.Style) {
	s.titleStyle = style
}
func (s *Sheet) SetDataStyle(style func() *excelize.Style) {
    s.dataStyle = style
}

func (s *Sheet) MinColWidth() float64 {
	return s.minColWidth
}

func (s *Sheet) SetMinColWidth(minColWidth float64) {
	s.minColWidth = minColWidth
}

func (s *Sheet) SetMaxColWidth(maxColWidth float64) {
	s.maxColWidth = maxColWidth
}

func (s *Sheet) MaxColWidth() float64 {
    return s.maxColWidth
}

// The data is written to the table in the order of these fields
func (s *Sheet) Fields(recalculate ...bool) []string {
    if (s.fieldSort == nil && s.baseDataType != nil) || (len(recalculate) == 1 && recalculate[0] == true) {
        return writers.FieldSort(s)
    }
    return s.fieldSort
}

// BeginStream initializes a StreamWriter for the current sheet when streaming is enabled.
func (s *Sheet) BeginStream() error {
    if !s.useStream {
        return nil
    }
    if s.stream != nil {
        return nil
    }
    sw, err := s.file.NewStreamWriter(s.SheetName())
    if err != nil {
        return err
    }
    s.stream = sw
    return nil
}

// EndStream flushes and clears the StreamWriter.
func (s *Sheet) EndStream() error {
    if s.stream == nil {
        return nil
    }
    err := s.stream.Flush()
    s.stream = nil
    return err
}
