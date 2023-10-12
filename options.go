package exportToExcel

import (
	"github.com/xuri/excelize/v2"
)

type Option func(s *Sheet)

func OptionSetTitleStyle(styleFn func() *excelize.Style) Option {
	return func(s *Sheet) {
		s.titleStyle = styleFn
	}
}
func OptionSetDataStyle(styleFn func() *excelize.Style) Option {
	return func(s *Sheet) {
		s.titleStyle = styleFn
	}
}
func OptionSetColWidth(min, max float64) Option {
	return func(s *Sheet) {
		s.minColWidth = min
		s.maxColWidth = max
	}
}
