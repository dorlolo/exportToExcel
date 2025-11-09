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
        s.dataStyle = styleFn
    }
}
func OptionSetColWidth(min, max float64) Option {
    return func(s *Sheet) {
        s.minColWidth = min
        s.maxColWidth = max
    }
}

// OptionAutoResetColWidth enables or disables auto column width calculation.
// When disabled, fixed width will be applied using the sheet's minColWidth.
func OptionAutoResetColWidth(enable bool) Option {
    return func(s *Sheet) {
        s.autoResetColWidth = enable
    }
}

// OptionEnableStreamWriter enables streaming write via excelize.StreamWriter.
// When enabled, writers will use StreamWriter to append rows with lower memory usage.
func OptionEnableStreamWriter(enable bool) Option {
    return func(s *Sheet) {
        s.useStream = enable
    }
}
