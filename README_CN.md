# exportToExcel
[English](./README.md) | 中文版

## 功能描述
这个Golang模块是专为快速导出数据而设计的。它在 [excelize](github.com/xuri/excelize/v2) 的基础上，封装了一些常用的功能，使数据导出变得更加简单。

## 如何使用?
###  下载和引用
- 下载
```cmd
go get github.com/dorlolo/exportToExcel
```
- 引用
```go
import "github.com/dorlolo/exportToExcel"
```

### 代码示例
```go
package main

import (
	"fmt"
	"github.com/dorlolo/exportToExcel"
)

// 首先需要创建一个工作表的数据模型，并为每个字段指定json标签。
// 默认情况下，将会按照结构体中的字段顺序导出到表中。
type DemoBaseDataTypeA struct {
	Name   string  `json:"name"`
	Age    int     `json:"age"`
	Height float32 `json:"height"`
}

func main() {
	var err error
	// 生成一个Excel文件对象
	ex := exportToExcel.NewExcel(".", "newfile.xlsx")
	defer func() {
		if err == nil {
			if err = ex.Save(); err != nil {
				fmt.Println("save file err:",err)
				return
			}
		}
	}()
	// 生成一个工作表对象
	// 需要传入基础数据类型,最终st1会支持这些数据类型: 
	//    DemoBaseDataTypeA{} , *DemoBaseDataTypeA{} , []DemoBaseDataTypeA{} 和 []*DemoBaseDataTypeA{}
	st1 := ex.NewSheet("st1", DemoBaseDataTypeA{})
    // 设置表头，如果不需要你可以忽略它
	err = st1.Title.Gen(
		st1.Title.NewTitleItem(4, "st1-demo", 1, 1).SetFullHorizontalMerge(),// 你可以使用类似的方法使表头跨列或跨行
		st1.Title.NewTitleItem(5, "name", 2, 1),
		st1.Title.NewTitleItem(5, "age", 2, 2),
		st1.Title.NewTitleItem(5, "height", 2, 3),
	)
	if err != nil {
		fmt.Println("generate title failed:", err.Error())
		return
	}
	// 填充数据
	// 数据默认会按照结构体中的字段顺序填充到表中。
	// st1.Title.Gen和st1.FillData在使用上不存在先后顺序。
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

#### 从模板读取文件
如果表头样式比较复杂，你可以创建一个文件，编辑好表头样式并将其作为模板使用，从而简化代码。
```go
ex, err := exportToExcel.NewExcelFromTemplate("./example/template.xlsx", ".", "newfile.xlsx")
if err != nil {
    t.Error(err)
    return
}
defer func() {
	// 保存
    if err == nil {
        if err = ex.Save(); err != nil {
            fmt.Println("save file err:",err)
            return
        }
    }
}()
// 获取表对象
st := ex.GetSheetByName("Sheet1")
if st == nil {
    t.Error("can not find sheet:Sheet1")
    return
}
// 设置表数据的基本类型，这一步很重要!
st.SetDataType(DemoBaseDataTypeA{})
// 填充数据,数据会自动从空行处添加，无需担心会覆盖表头
var data1 = []DemoBaseDataTypeA{
    {"Mr.Zhang", 16, 180},
    {"Mrs.Li", 18, 220},
}
if err = st.FillData(data1); err != nil {
    t.Error(err)
    return
}
```

#### 指定输出字段和写入顺序
通过以下方法传入json标签，将会按照此顺序输出到表中。
```go
st.SetFieldSort("age","name","height")
```

#### 动态表头

#### 自定义样式

#### 自定义写入器
如果你需要自定义写入器，请使用这个接口[IDataWriter](./writer.go)实现。,然后使用`RegisterDataWriter`方法注册到模块中。填充数据时将优先匹配手动注册进来的写入器。