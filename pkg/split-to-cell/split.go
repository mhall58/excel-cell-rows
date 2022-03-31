package split_to_cell

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strings"
)

func SplitCells(f *excelize.File, columnNotBlank string, splitColumn string, delimiter string) {
	sheets := f.GetSheetList()

	for _, sheetName := range sheets {
		currentRow := 1
		var splitValues []string

		for {

			splitCheckCell := fmt.Sprintf("%s%v", splitColumn, currentRow)
			checkCell := fmt.Sprintf("%s%v", columnNotBlank, currentRow)
			rowHasData := func() bool { v, _ := f.GetCellValue(sheetName, checkCell); return v != "" }()

			if !rowHasData || currentRow >= 1000 {
				break
			}

			splitValues = func() []string { v, _ := f.GetCellValue(sheetName, splitCheckCell); return strings.Split(v, delimiter) }()

			if len(splitValues) == 1 {
				currentRow++
				continue
			}

			for i, sVal := range splitValues {
				if i != 0 {
					f.DuplicateRow(sheetName, currentRow)
					currentRow++
				}
				cell := fmt.Sprintf("%s%v", splitColumn, currentRow)

				if err := f.SetCellValue(sheetName, cell, strings.TrimSpace(sVal)); err != nil {
					return
				}
			}

			currentRow++
		}

	}
}
