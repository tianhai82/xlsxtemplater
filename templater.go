package xlsxtemplater

import (
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// FillTemplate takes a excel file and a sheet name with a map of key,values and fill up the necessary info
func FillTemplate(values map[string]string, xlsx *excelize.File, sheet string) error {
	rows := xlsx.GetRows(sheet)
	for r, row := range rows {
		for c, cell := range row {
			val := strings.TrimSpace(cell)
			if strings.HasPrefix(val, "{") && strings.HasSuffix(val, "}") {
				key := strings.Trim(val, "{}")
				if replace, ok := values[key]; ok {
					alpha := excelize.ToAlphaString(c)
					alias := alpha + strconv.Itoa(r+1)
					xlsx.SetCellStr(sheet, alias, replace)
				}
			}
		}
	}
	return nil
}
