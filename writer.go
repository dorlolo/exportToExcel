// Package excel -----------------------------
// @file      : writer.go
// @author    : JJXu
// @contact   : wavingbear@163.com
// @time      : 2023/9/1 14:11
// -------------------------------------------
package exportToExcel

import (
	"errors"
	"fmt"
	"reflect"
)

var writers writerList

type writerList []IDataWriter

func (w writerList) FieldSort(sheetObj *Sheet) []string {
	for _, dw := range w {
		if dw.Supported(sheetObj.data) {
			return dw.FieldSort(sheetObj.baseDataType)
		}
	}
	return nil
}
func (w writerList) WriteData(sheetObj *Sheet) error {
	for _, dw := range w {
		if dw.Supported(sheetObj.data) {
			return dw.WriteData(sheetObj)
		}
	}
	return errors.New("There is no supported data writer")
}

type IDataWriter interface {
	Supported(any) bool //determine whether the data type supports the writer
	//Tag() string
	WriteData(sheetObj *Sheet) error
	FieldSort(baseDataType any) []string //default field sorting method
}

func init() {
	writers = writerList{&sliceWriter{}, &structWriter{}}
}

// The registered Data Writer is preferentially used
func RegisterDataWriter(writer ...IDataWriter) {
	writers = append(writer, writers...)
}

//======================================================================================================================

// ======================================================================================================================
type sliceWriter struct {
}

func (s sliceWriter) Supported(data any) bool {
	dataType := reflect.TypeOf(data)
	if dataType.Kind() == reflect.Slice || dataType.Kind() == reflect.Array {
		elementType := dataType.Elem()
		if elementType.Kind() == reflect.Ptr {
			elementType = elementType.Elem()
		}
		if elementType.Kind() == reflect.Struct {
			return true
		}
	}
	return false
}

// , SheetData reflect.Value, SheetDataType reflect.Type, firstRow int
func (s sliceWriter) WriteData(sheetObj *Sheet) error {
	var cellNameList = sheetObj.Fields()
	var refValue = reflect.ValueOf(sheetObj.Data())
	var refType = reflect.TypeOf(sheetObj.baseDataType)
	var columnLen = refType.NumField()
	for i := 0; i < columnLen; i++ {
		var dataMap = s.dataToMap(refValue.Index(i), refType)
		for column, v := range cellNameList {
			var axis = GetCellCoord(i+sheetObj.firstEmptyRow+1, column+1)
			err := sheetObj.file.SetCellValue(sheetObj.SheetName(), axis, dataMap[v])
			if err != nil {
				return err
			}
		}
	}
	//设置数据格式
	dataStyleID, errs := sheetObj.file.NewStyle(NewDefaultDataStyle())
	if errs != nil {
		return errs
	}
	maxCol := columnLen
	titleLen := len(sheetObj.Title.titles)
	if titleLen > maxCol {
		maxCol = titleLen
	}
	if maxCol == 0 {
		maxCol = 1
	}
	if err := sheetObj.file.SetCellStyle(sheetObj.SheetName(), GetCellCoord(sheetObj.firstEmptyRow+1, 1), GetCellCoord(sheetObj.rowNum, maxCol), dataStyleID); err != nil {
		return err
	}
	//设置默认列宽
	//exc.ex.SetColWidth(sheetObj.SheetName(), GetColumnIndex(1), GetColumnIndex(len(sheetObj.SheetHeaders())), 12.0)
	return AutoResetCellWidth(sheetObj)
}

func (s sliceWriter) FieldSort(baseDataType any) []string {
	refType := reflect.TypeOf(baseDataType)
	list, _ := GetJsonFieldList(refType)
	return list
}

func (s sliceWriter) dataToMap(sheet reflect.Value, sheetType reflect.Type) (dataMap map[string]any) {
	dataMap = make(map[string]any)
	//指针类型结构体拿真实的对象
	if sheetType.Kind() == reflect.Ptr {
		sheetType = sheetType.Elem()
		sheet = sheet.Elem()
	}
	for i := 0; i < sheetType.NumField(); i++ {
		var tag = sheetType.Field(i).Tag.Get("json")
		if tag != "" {
			dataMap[sheetType.Field(i).Tag.Get("json")] = sheet.Field(i).Interface()
		}
	}
	return dataMap
}

type structWriter struct {
}

func (s structWriter) Supported(data any) bool {
	switch data.(type) {
	case struct{}:
		return true
	case *struct{}:
		return true
	default:
		return false
	}
}

func (s structWriter) WriteData(sheetObj *Sheet) error {
	var dataType = reflect.TypeOf(sheetObj.baseDataType)
	var dataValue = reflect.ValueOf(sheetObj.Data())
	var dataMap = DataToMapByJsonTag(dataValue, dataType)
	for column, v := range sheetObj.Fields() {
		var axis = GetCellCoord(sheetObj.firstEmptyRow+1, column+1)
		err := sheetObj.file.SetCellValue(sheetObj.SheetName(), axis, dataMap[v])
		if err != nil {
			fmt.Println(err.Error())
			return err
		}
	}
	//设置数据格式
	dataStyleID, err := sheetObj.file.NewStyle(NewDefaultDataStyle())
	if err != nil {
		return err
	}
	//todo sheetObj.Title.titles 这里应该不对
	if err = sheetObj.file.SetCellStyle(sheetObj.SheetName(), GetCellCoord(sheetObj.firstEmptyRow, 1), GetCellCoord(sheetObj.firstEmptyRow+1, len(sheetObj.Title.titles)), dataStyleID); err != nil {
		return err
	}
	//设置默认列宽
	//exc.ex.SetColWidth(sheetObj.SheetName(), GetColumnIndex(1), GetColumnIndex(len(sheetObj.SheetHeaders())), 12.0)
	return AutoResetCellWidth(sheetObj)
}

func (s structWriter) FieldSort(baseDataType any) []string {
	refType := reflect.TypeOf(baseDataType)
	list, _ := GetJsonFieldList(refType)
	return list
}
