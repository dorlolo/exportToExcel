# exportToExcel
English | [中文版](./README_CN.md)
## Description
This is a module of golang which used to export data,it's based on [excelize](github.com/xuri/excelize/v2),


## How to use?
###  download and import
- download
```cmd
go get github.com/dorlolo/exportToExcel
```
- import
```go
import "github.com/dorlolo/exportToExcel"
```

### excample
```go
package main

import (
	"fmt"
	"github.com/dorlolo/exportToExcel"
)
//
type DemoBaseDataTypeA struct {
	Name   string  `json:"name"`
	Age    int     `json:"age"`
	Height float32 `json:"height"`
}

func main() {
	var err error
	// Generate an excel file object
	ex := exportToExcel.NewExcel(".", "newfile.xlsx")
	defer func() {
		if err == nil {
			if err = ex.Save(); err != nil {
				fmt.Println("save file err:",err)
				return
			}
		}
	}()
	// Generate a sheet，You need to specify the basic data type which will to be stored.
	// it means to support these data type: DemoBaseDataTypeA{} , *DemoBaseDataTypeA{} , []DemoBaseDataTypeA and []*DemoBaseDataTypeA
	st1 := ex.NewSheet("st1", DemoBaseDataTypeA{})
    // Set the title if you needed
	err = st1.Title.Gen(
		st1.Title.NewTitleItem(4, "st1-demo", 1, 1).SetFullHorizontalMerge(),// You can use method like this to make headings span columns or rows
		st1.Title.NewTitleItem(5, "name", 2, 1),
		st1.Title.NewTitleItem(5, "age", 2, 2),
		st1.Title.NewTitleItem(5, "height", 2, 3),
	)
	if err != nil {
		fmt.Println("generate title failed:", err.Error())
		return
	}
	// Fill data to sheet
	var data1 = []DemoBaseDataTypeA{
		{"Mr.Zhang", 16, 180},
		{"Mrs.Li", 18, 220},
	}
	err = st1.FillData(data1)
	if err != nil {
		fmt.Println("fill data1 err:", err)
		return
	}
}
```

### 常用操作

#### 设置样式

#### 从模板读取文件
如果要设置较复杂的表头，在代码中配置会比较复杂。你可以创建一个文件，编辑好表头样式并将其作为模板。
```go
ex, err := exportToExcel.NewExcelFromTemplate("./example/template.xlsx", ".", "newfile.xlsx")
if err != nil {
    t.Error(err)
    return
}
// Get sheet object
st := ex.GetSheetByName("Sheet1")
if st == nil {
    t.Error("can not find sheet:Sheet1")
    return
}
st.SetDataType(DemoBaseDataTypeA{})
var data1 = []DemoBaseDataTypeA{
    {"Mr.Zhang", 16, 180},
    {"Mrs.Li", 18, 220},
}
if err = st.FillData(data1); err != nil {
    t.Error(err)
    return
}
if err = ex.Save(); err != nil {
    fmt.Println("save file err:",err)
    return
}
```