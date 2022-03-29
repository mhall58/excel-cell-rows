package cmd

import (
	"fmt"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
	"os"
	"strings"
)

// rootCmd represents the base command when called without any subcommands
var rootCmd = &cobra.Command{
	Use:   "split-by-cell",
	Short: "splits cells into rows by a delimiter",
	Run: func(cmd *cobra.Command, args []string) {

		if len(args) != 1 {
			cmd.Println("invalid number of arguments, only pass the file and/or path to the file")
			os.Exit(1)
		}

		path := args[0]

		f, err := excelize.OpenFile(path)
		if err != nil {
			cmd.PrintErrln(err)
			os.Exit(1)
		}

		defer func() {
			// Close the spreadsheet.
			if err := f.Close(); err != nil {
				cmd.PrintErrln(err)
				os.Exit(1)
			}
		}()

		sheets := f.GetSheetList()
		columnNotBlank, _ := cmd.Flags().GetString("column-not-blank")
		splitColumn, _ := cmd.Flags().GetString("split-column")
		delimiter, _ := cmd.Flags().GetString("delimiter")

		for _, sheetName := range sheets {
			cmd.Println("Scanning sheet: " + sheetName + "...")
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
					cmd.Println(fmt.Sprintf("cell %v has %v values, skipping", splitCheckCell, len(splitValues)))
					currentRow++
					continue
				}
				cmd.Println(fmt.Sprintf("cell %v has %v values", splitCheckCell, len(splitValues)))

				for i, sVal := range splitValues {
					if i != 0 {
						f.DuplicateRow(sheetName, currentRow)
						currentRow++
					}
					cmd.Println(fmt.Sprintf("creating row %v", currentRow))
					cell := fmt.Sprintf("%s%v", splitColumn, currentRow)
					cmd.Println("Setting CELL: " + cell + " TO: " + sVal)

					if err := f.SetCellValue(sheetName, cell, sVal); err != nil {
						cmd.PrintErrln(err)
						os.Exit(1)
					}
				}

				currentRow++
			}

		}

		newPath := strings.ReplaceAll(path, ".xlsx", "-split.xlsx")

		cmd.Println("saving... " + newPath)

		if err := f.SaveAs(newPath); err != nil {
			cmd.PrintErrln(err)
			os.Exit(1)
		}

		cmd.Println("DONE!")

	},
}

// Execute adds all child commands to the root command and sets flags appropriately.
// This is called by main.main(). It only needs to happen once to the rootCmd.
func Execute() {
	err := rootCmd.Execute()
	if err != nil {
		os.Exit(1)
	}
}

func init() {
	rootCmd.Flags().String("delimiter", "\n", "the delimiter to use")
	rootCmd.Flags().String("split-column", "i", "the column to look for the delimiter")
	rootCmd.Flags().String("column-not-blank", "a", "the column to check for data")
}
