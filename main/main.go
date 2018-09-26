package main

import (
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/tianhai82/xlsxtemplater"
)

func main() {
	xlsx, err := excelize.OpenFile("ExecCompTemplate.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	values := map[string]string{
		"sectors":   "All",
		"countries": "Singapore, Malaysia",
	}
	xlsxtemplater.FillTemplate(values, xlsx, "Aggregated View")
	xlsx.SaveAs("huat.xlsx")
}
