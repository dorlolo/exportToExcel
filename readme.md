# excel文件生成器
## excel文件生成器简介
此模块提供了通过`.xlsx`格式的excel模板文件，创建文件并自动添加数据的方法。
此模块通过设计模式，在"github.com/xuri/excelize/v2"的基础上封装了`ExcelCreator`这一泛型的方法。适用于快速开发报表导出功能。

### 基本使用示例
```go
package main
import (
    "fmt"
    "github.com/flipped-aurora/gin-vue-admin/server/utils/excel"
    "github.com/flipped-aurora/gin-vue-admin/server/utils/simpletime"
    "sync"
    "time"
)

//读写锁
var exportEpidemicPreventStaticStaticLock sync.RWMutex

//定义sheet1数据结构
type EpidemicPreventStaticReport struct {
    Xid    int    `json:"xid" form:"xid" db:"column:xid;comment:序号"`
    Time   string `json:"time" form:"time" db:"column:time;comment:日期"`
    Name   string `json:"name" form:"name" db:"column:name;comment:项目名称"`
    Code   string `json:"code" form:"code" db:"column:code;comment:监督备案号"`
    Street string `json:"street" form:"street" db:"column:street;comment:街道"`
    State  string `json:"state" form:"state" db:"column:state;comment:状态：未完成、已完成"`
}

//导出数据
func (m EpidemicPreventStaticReport) WriteToExcel(datas []EpidemicPreventStaticReport) (path string, filename string, err error) {
    defer exportEpidemicPreventStaticStaticLock.Unlock()
	//实例化sheet1,载入数据
    st1 := excel.NewSheet("Sheet1", datas)
    var suffixName = func() string {
        return fmt.Sprintf("%v", simpletime.TimeToString(time.Now(), simpletime.TimeFormat.NoSpacer_YMDhms))
    }
	//实例化excelCreator
    exc, err := excel.NewExcelCreator("防疫日报日完成项目数统计.xlsx", &suffixName, "uploads/template2/file", "uploads/template2/防疫日报日完成项目数统计.xlsx", st1)
    if err != nil {
        fmt.Println(err.error())
    }
    exportEpidemicPreventStaticStaticLock.Lock()
    // 导出数据
    return exc.WriteToExcel()
}

func main(){
    var (
        reportData        []EpidemicPreventStaticReport
        excel             EpidemicPreventStaticReport
    )
    
    //do something
    //..
    //..
    
    path,filename,err:=excel.WriteToExcel(reportData)
    if err!=nil{
        fmt.Println(err.Error())
    }else{
        fmt.Printf("path:%v\nfilename:%v\n",path,filename)
    }
}
```

## 使用说明
此模块的使用流程如下
### 1. 创建excel模板，定义好工作簿名称和表头
请注意，程序默认数据都是一行行连续的。如果两条数据之间间隔了一个或多个空行，导出的数据可能会出现错误。
![img.png](old/doc/img.png)
### 2. 定义工作簿数据结构
注意事项:
1. 结构体字段顺序要与工作簿的表头顺序一一对应、命名规则随意;
2. json标签用于数据的转换，同时也便于将结构体直接作为接口请求参数来使用;没有json标签时，数据将会被忽略;
```go
type Sheet1 struct {
	Xid  int   `json:"xid"`
	Name string `json:"name"`
	Age  int   `json:"age"`
}
```
### 3. 准备数据,实例化sheet对象
数据类型至此结构体和切片<br/>
`NewSheet`方法的第一个参数是工作簿名称，需要与模板文件中的对应,不然导出数据时会报错
```go
//准备数据
var sheet1Data = []Sheet1Define{
{1, "张三", 16},
{2, "黑猫警长", 18},
}
//实例化sheet对象
var sheet1 = excel.NewSheet("Sheet1", &sheet1Data)
```
### 4 生成excel文件
### 4.1 直接生成文件
如果你不需要什么额外操作，只想直接生成excel文件，那么可以直接调用这个方法
```go
//定义后缀名生成器，如果不需要可以传nil
var suffixFunc = func() string { return fmt.Sprintf("%v", time.Now().Unix()) }
//导出文件
path,err:=excel.WriteToExcel("demo.xlsx", &suffixFunc, "./", "./demo.xlsx", sheet_1)
```
#### 生成效果:
![img_1.png](old/doc/img_1.png)

### 4.2 `ExcelCreator`
ExcelCreator可以用来对工作簿和工作簿中的数据进行增删改查，以及生成文件。
#### 4.2.1 实例化`ExcelCreator`
```go

var suffixFunc = func() string { return fmt.Sprintf("%v", time.Now().Unix()) } //文件后缀名生成方法
exCreator,err := excel.NewExcel("demo.xlsx", &suffixFunc, "./", "./demo.xlsx", sheet_1)
if err!=nil{
	fmt.Println(err.Error())
}
//可连续添加多个工作簿或者不传，当然也支持一张工作簿数据的多次传入
//exCreator,err := excel.NewExcel("demo.xlsx", &suffixFunc, "./", "./demo.xlsx", sheet_1,sheet_2,sheet_3)
//exCreator,err := excel.NewExcel("demo.xlsx", &suffixFunc, "./", "./demo.xlsx")
//exCreator,err := excel.NewExcel("demo.xlsx", &suffixFunc, "./", "./demo.xlsx", sheet_1_1,sheet_1_2,sheet_1_3)

//不使用文件后缀名
//exCreator,err := excel.NewExcelCreator("demo.xlsx", nil, "./", "./demo.xlsx", sheet_1)
//exCreator,err := excel.NewExcelCreatorWithoutSuffix("demo.xlsx","./", "./demo.xlsx", sheet_1)
```
#### `FileName`  和 `filesSuffix`参数的说明
当设置了`filesSuffix`参数后，创建excel文件时会自动生成文件后缀名。
如`FileName`设置为"demo.xlsx"、`filesSuffix`返回值为"20200723",那么最终文件名为"demo20200723.xlsx"

### 4.2.2 新增工作簿
如果工作簿已存在，数据会组合而不是覆盖
```go
var err = exCreator.SheetsAdd(sheet_2)
```
### 4.2.3 删除工作簿
```go
var err = exCreator.SheetsDelete(sheet_2.SheetName())
```

### 4.2.5 取出工作簿map和缓存数据
工作簿存储在`ExcelCreator.Sheets`中,存储结构为`map[string][]sheet`。string即工作簿名称。取出后可以进行任意的操作
```go
var sheets = exCreator.Sheets
datas:= sheets["Sheet1"].GetDatas
//do somthing...
```

### 4.2.6 将数据生成到文件
```go
path,fileName, err := exCreator.WriteToExcel()
if err != nil {
    fmt.Println(err.Error())
} else {
    fmt.Println(path)
    fmt.Println(fileName)
}
```

## 相关问题
### 1. 为什么不支持通过代码直接生成表头？
鉴于在部分场景下，表头格式比较复杂，且通过代码生成灵活较度差,单元格格式通过代码设置比较繁琐。相比较下，没有直接创建excel文件模板来调整更为直观方便，所以没做此方面的设计。
