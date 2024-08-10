package exportToExcel

import "github.com/xuri/excelize/v2"

func DefaultTitleStyle() *excelize.Style {
	return &excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 2},
			{Type: "top", Color: "000000", Style: 2},
			{Type: "bottom", Color: "000000", Style: 2},
			{Type: "right", Color: "000000", Style: 2},
		},
		Font: &excelize.Font{Bold: true, Size: 12},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
		Protection: &excelize.Protection{},
	}
}

func DefaultDataStyle() *excelize.Style {
	return &excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
		Font: &excelize.Font{Size: 12},
		Alignment: &excelize.Alignment{
			Horizontal: "left",
		},
		Protection: &excelize.Protection{},
	}
}
