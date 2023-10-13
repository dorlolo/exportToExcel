package exportToExcel

import (
	"errors"
	"github.com/xuri/excelize/v2"
	"reflect"
)

// 获取excel的列索引
var columnIndices = []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

// 抽象工作簿
//type ISheetBuilder interface {
//	Data() any           //数据，支持struct、[]struct、[]*struct三种类型
//	SheetName() string   //表名称
//	Headers() []Header   //表头名称，没有则跳过插入表头的步骤
//	FieldSort() []string //指定字段的排列顺序，默认按照结构体中的顺序排列(表头与数据没有关联，使用这个指定字段插入顺序)
//	SetData(any) error
//	SetSheetName(string)
//	SetHeaders(...Header)
//	SetFieldSort([]string)
//	FillData(file *excelize.File) error
//}

type DataType interface {
	~struct{} | ~*struct{} | ~[]struct{} | ~[]*struct{}
}

type Option func(s *Sheet)

func NewSheet(file *excelize.File, sheetName string, baseDataType any, opts ...Option) *Sheet {
	var a = &Sheet{file: file, baseDataType: baseDataType}
	a.Title = NewTitle(a)
	a.sheetId, _ = file.NewSheet(sheetName)
	a.defaultStyle = NewDefaultTitleStyle
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
	if (s.fieldSort == nil && s.data != nil) || (len(recalculate) == 1 && recalculate[0] == true) {
		return writers.FieldSort(s)
	}
	return s.fieldSort
}
