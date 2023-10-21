package exportToExcel

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
)

// 获取excel的列索引
var columnIndices = []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

const minCloWidth float64 = 6

// 自动设置单元格宽度
func AutoResetCellWidth(sheetObj *Sheet, setLimitWidth ...float64) error {
	// 获取最大字符宽度
	maxWidths := make(map[int]float64)
	var sheetData = reflect.ValueOf(sheetObj.Data())
	var rowLen = 1
	var columnLen = len(sheetObj.Fields())
	if sheetData.Kind() == reflect.Slice {
		rowLen = sheetData.Len()
	}
	var limitWidth float64 = 120
	if setLimitWidth != nil {
		limitWidth = setLimitWidth[0]
	}
	for col := 1; col <= columnLen; col++ {
		var maxWidth = minCloWidth
		for row := 0; row < rowLen; row++ {
			value, _ := sheetObj.file.GetCellValue(sheetObj.SheetName(), GetCellCoord(row+1, col))
			width := float64(len(value))
			if width > limitWidth {
				width = limitWidth
			}
			if width > maxWidth {
				maxWidth = width
			}
		}
		maxWidths[col] = maxWidth
	}

	// 设置列宽度
	for col, width := range maxWidths {
		colChar := GetColumnIndex(col)
		err := sheetObj.file.SetColWidth(sheetObj.SheetName(), colChar, colChar, width+2)
		if err != nil {
			return err
		}
	}
	return nil
}

// 行列坐标值转换为excel的坐标。注意row和columnCount的初始值都是1
func GetCellCoord(row int, columnCount int) string {
	if row == 0 {
		row = 1
	}
	if columnCount == 0 {
		columnCount = 1
	}
	var column = GetColumnIndex(columnCount)
	return fmt.Sprintf("%s%d", column, row)
}

func GetColumnIndex(num int) string {
	num--
	var column = columnIndices[num%26]
	for num = num / 26; num > 0; num = num / 26 {
		column = columnIndices[(num-1)%26] + column
		num--
	}
	return column
}

func GetFirstEmptyRowIndex(ex *excelize.File, sheetName string) (index int) {
	rows, err := ex.GetRows(sheetName)
	if err != nil {
		return 1
	}
	return len(rows)
}

func DataToMapByJsonTag(sheet reflect.Value, sheetType reflect.Type) (dataMap map[string]any) {
	dataMap = make(map[string]any)
	if sheet.Kind() == reflect.Ptr {
		sheet = sheet.Elem()
	}
	if sheetType.Kind() == reflect.Ptr {
		sheetType = sheetType.Elem()
	}
	for i := 0; i < sheetType.NumField(); i++ {
		var tag = sheetType.Field(i).Tag.Get("json")
		if tag != "" {
			dataMap[sheetType.Field(i).Tag.Get("json")] = sheet.Field(i).Interface()
		}
	}
	return dataMap
}

func GetJsonFieldList(structObj reflect.Type) (list []string, err error) {
	//根据data字段排序
	if structObj.Kind() == reflect.Ptr {
		structObj = structObj.Elem()
		if structObj.Kind() != reflect.Struct {
			return []string{}, errors.New("GetJsonFieldList: structObj must struct{} or *struct{}!")
		}
	}
	for i := 0; i < structObj.NumField(); i++ {
		var tag = structObj.Field(i).Tag.Get("json")
		if tag != "" {
			list = append(list, tag)
		}
	}
	return
}
