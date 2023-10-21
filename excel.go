package exportToExcel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"path/filepath"
)

type Excel struct {
	fileDir  string
	fileName string
	sheets   []*Sheet
	file     *excelize.File
}

func NewExcel(fileDir, filename string) *Excel {
	return &Excel{fileName: filename, fileDir: fileDir, file: excelize.NewFile()}
}
func NewExcelFromTemplate(templatePath string, saveDir, saveName string) (*Excel, error) {
	f, err := excelize.OpenFile(templatePath)
	if err != nil {
		return nil, err
	}
	ex := &Excel{fileName: saveName, fileDir: saveDir, file: f}
	for _, stName := range f.GetSheetList() {
		id, err := f.GetSheetIndex(stName)
		if err != nil {
			return ex, err
		}
		ex.sheets = append(ex.sheets, &Sheet{
			file:    f,
			sheetId: id,
		})
	}
	return ex, nil
}

func (e *Excel) NewSheet(sheetName string, baseDataType any, opts ...Option) *Sheet {
	st := newSheet(e.file, sheetName, baseDataType, opts...)
	e.sheets = append(e.sheets, st)
	return st
}

func (e *Excel) GetSheetByName(sheetName string) *Sheet {
	for _, v := range e.sheets {
		if v.SheetName() == sheetName {
			return v
		}
	}
	return nil
}

func (e *Excel) GetSheetById(sheetId int) *Sheet {
	for _, v := range e.sheets {
		if v.sheetId == sheetId {
			return v
		}
	}
	return nil
}
func (e *Excel) SetFileName(fileName string) *Excel {
	e.fileName = fileName
	return e
}

func (e *Excel) SetFileDir(fileDir string) *Excel {
	e.fileDir = fileDir
	return e
}

func (e *Excel) FileName() string {
	return e.fileName
}

func (e *Excel) FileDir() string {
	return e.fileDir
}
func (e *Excel) File() *excelize.File {
	return e.file
}

func (e *Excel) Save() (err error) {
	//generate directories
	if err = os.MkdirAll(e.fileDir, os.ModePerm); err != nil {
		return
	}
	//delete Sheet 1
	var hasSheet1 bool
	for _, v := range e.sheets {
		if v.SheetName() == "Sheet1" {
			hasSheet1 = true
			break
		}
	}
	if !hasSheet1 {
		_ = e.file.DeleteSheet("Sheet1")
	}
	//保存
	path := filepath.ToSlash(filepath.Join(e.fileDir, e.fileName))
	if err = e.file.SaveAs(path); err != nil {
		log.Println(fmt.Sprintf("save file error :%v", err))
	}
	return
}
func (e *Excel) NewStyle(newStyle *excelize.Style) (int, error) {
	return e.file.NewStyle(newStyle)
}
