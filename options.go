package exportToExcel

import (
	"github.com/xuri/excelize/v2"
)

type Option func(s *Sheet)

func OptionSetDefaultStyle(styleFn func() *excelize.Style) Option {
	return func(s *Sheet) {
		s.defaultStyle = styleFn
	}
}
