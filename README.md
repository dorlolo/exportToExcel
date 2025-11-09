# exportToExcel
English | [中文版](./README_CN.md)

## Description

This Golang module is designed for quick data export. It encapsulates some common functions on the basis of [excelize](github.com/xuri/excelize/v2) to simplify data export.

It has the following features:

1.  Support using files as templates;
2.  Support setting dynamic headers, merging cells, header styles;
3.  Support parsing various data types;
4.  Support default cell styles;
5.  Auto fit column width based on data length.

## How To Use?

### Download and Import

-   Download

```cmd
go get github.com/dorlolo/exportToExcel
```

-   Import

```go
import "github.com/dorlolo/exportToExcel"
```

### Code Example

```go
package main

import (
   "fmt"
   "github.com/dorlolo/exportToExcel"
)

// First, create a model for the sheet data, 
// and specify json tags for each field.
// By default, fields will be exported in struct order.
type DemoBaseDataTypeA struct {
   Name   string  `json:"name"`
   Age    int     `json:"age"`
   Height float32 `json:"height"`
}

func main() {
   var err error
   // Create an Excel object
   ex := exportToExcel.NewExcel(".", "newfile.xlsx")
   defer func() {
      if err == nil {
         if err = ex.Save(); err != nil {
            fmt.Println("save file err:",err)
            return
         }
      }
   }()
   // Create a sheet
   // Need to pass the base data type, st1 will eventually support:
   //    DemoBaseDataTypeA{} , *DemoBaseDataTypeA{} , []DemoBaseDataTypeA{} and []*DemoBaseDataTypeA{}
   st1 := ex.NewSheet("sheet1", DemoBaseDataTypeA{})
   // Set header
   err = st1.Title.Gen(
      st1.Title.NewTitleItem(4, "sheet1-demo", 1, 1).SetFullHorizontalMerge(),// You can use similar methods to merge header cells
      st1.Title.NewTitleItem(5, "name", 2, 1),
      st1.Title.NewTitleItem(5, "age", 2, 2),
      st1.Title.NewTitleItem(5, "height", 2, 3),
   )
   if err != nil {
      fmt.Println("generate title failed:", err.Error())
      return
   }
   // Fill data
   // Data will be filled according to struct field order by default;
   // When filling data, rows with data will be skipped automatically, no need to worry about overwriting headers;
   // There is no order requirement between st1.Title.Gen and st1.FillData.  
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

### Common Operations

#### Read From Template

If header styles are complex, you can create a template file to simplify code instead of generating headers in code.


```go
ex, err := exportToExcel.NewExcelFromTemplate("./example/template.xlsx", ".", "newfile.xlsx")
if err != nil {
    t.Error(err)
    return
}
defer func() {
   // Save
   if err == nil {
       if err = ex.Save(); err != nil {
           fmt.Println("save file err:",err)
           return
       }
   }
}()
// Get sheet object
st := ex.GetSheetByName("Sheet1")
if st == nil {
    t.Error("can not find sheet:Sheet1")
    return
}
// Set data type, this is important! 
st.SetDataType(DemoBaseDataTypeA{})
// Fill data, will add from empty rows automatically, no need to worry about overwriting headers
var data1 = []DemoBaseDataTypeA{
    {"Mr.Zhang", 16, 180},
    {"Mrs.Li", 18, 220},
}
if err = st.FillData(data1); err != nil {
    t.Error(err)
    return
}
```

#### Specify Output Fields and Order

Pass in json tags to specify output order, other fields will be ignored.

```go
st.SetFieldSort("age","name","height")
```

#### Custom Styles

Styles are mainly divided into header cell styles and data cell styles. There are two ways to set styles:

1.  Use [Option](./options.go) methods when creating sheet to configure

```go
ex := exportToExcel.NewExcel("./", "aa.xlsx")
st:=ex.NewSheet("sheet1", ExDataTYpe{}, exportToExcel.OptionSetTitleStyle(func() *excelize.Style {
    // Easier to modify based on default style
    newStyle := exportToExcel.DefaultDataStyle()
   newStyle.Border[0].Color = "red"
    return newStyle
}))
```

2.  Use `SetTitleStyle` and `SetDataStyle` methods on Sheet object

Refer to code in [./style.go](./style.go) for style generator functions.

```go
ex := exportToExcel.NewExcel(".", "aa.xlsx")
st:=ex.NewSheet("sheet1", ExDataTYpe{})
// Set header cell style
st.SetTitleStyle(exportToExcel.DefaultTtileStyle())  
// Set data cell style
st.SetDataStyle(exportToExcel.DefaultDataStyle())
```

### Modify AutoFit Column Width Range

Custom column width can be set for each sheet. -1 means no limit.

```go
ex := exportToExcel.NewExcel(".", "aa.xlsx")
st1:=ex.NewSheet("sheet1", ExDataTYpe{},exportToExcel.OptionSetColWidth(exportToExcel.DefaultColMinWidth,500))
st2:=ex.NewSheet("sheet2", ExDataTYpe{},exportToExcel.OptionSetColWidth(-1,500))
```

#### Custom Data Writer

To use a custom writer, implement the [IDataWriter](./writer.go) interface and register with `RegisterDataWriter`. Custom registered writers will be matched first when filling data.

#### Use `*excelize.File` Object for More Features
If the above functions don't meet your needs, you can get the `*excelize.File` object for more features.
```go
ex := exportToExcel.NewExcel(".", "aa.xlsx")
exFile:=ex.File()
// Use exFile for more features  
//...
```
There are some examples below:
-  Add dropdown menu
```go
vd := excelize.NewDataValidation(true)
vd.SetSqref("D2:D100")
_ = vd.SetDropList([]string{"red", "green", "yellow"})
_ = exFile.AddDataValidation("Sheet1", vd)
```
- Lock cells
```go
//Configure locked cell style
lockCellStyle := exportToExcel.DefaultDataStyle()
lockCellStyle.Protection.Locked = true
lockCellStyleId, _ := exf.NewStyle(lockCellStyle)
//Lock cells A1-C2 in Sheet 1
_ = exFile.SetCellStyle("Sheet1", "A1", "C2", lockCellStyleId)
//Set the protection sheet to make the lock function take effect
err = exFile.ProtectSheet("Sheet1", &excelize.SheetProtectionOptions{
	// Set password
    //Password:            "123456",
    //AlgorithmName:       "SHA-512",
    SelectLockedCells:   true,
    SelectUnlockedCells: true,
}
```

### Enable Streaming Write

For large datasets, you can enable streaming write to reduce memory usage. When enabled, rows are appended via `excelize.StreamWriter`.

```go
ex := exportToExcel.NewExcel(".", "stream.xlsx")
st := ex.NewSheet("sheet1", DemoBaseDataTypeA{}, exportToExcel.OptionEnableStreamWriter(true))
// Optional: generate title
_ = st.Title.Gen(st.Title.NewTitleItem(4, "users", 1, 1).SetFullHorizontalMerge())

// Fill data (struct slice or single struct)
data := []DemoBaseDataTypeA{{"Alice", 30, 180}, {"Bob", 25, 175}}
if err := st.FillData(data); err != nil {
    panic(err)
}
if err := ex.Save(); err != nil {
    panic(err)
}
```

Notes:
- API remains the same; only the write path changes to streaming.
- Data styles and auto column width are still applied after writing.

### Disable Auto-Fit Column Width and Use Fixed Width

For large exports, auto-fit column width (scanning cell contents) can be expensive. You can disable auto-fit and apply a fixed width to reduce computation.

```go
ex := exportToExcel.NewExcel(".", "stream.xlsx")
st := ex.NewSheet(
    "sheet1",
    DemoBaseDataTypeA{},
    exportToExcel.OptionEnableStreamWriter(true),  // optional: enable streaming write
    exportToExcel.OptionAutoResetColWidth(false),  // disable auto-fit width
    exportToExcel.OptionSetColWidth(10, 10),       // set fixed width baseline
)

data := []DemoBaseDataTypeA{{"Alice", 30, 180}, {"Bob", 25, 175}}
if err := st.FillData(data); err != nil {
    panic(err)
}
if err := ex.Save(); err != nil {
    panic(err)
}
```

Notes:
- When auto-fit is disabled, columns are set once with a fixed width; `OptionSetColWidth(min, max)` uses `min` as the effective width.
- Adjust `min` as needed to balance visual comfort and export performance.
