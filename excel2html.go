package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/google/uuid"
	"github.com/xuri/excelize/v2"
)

func getCellStyle(f *excelize.File, sheetName, cell string) (string, error) {
	styleID, err := f.GetCellStyle(sheetName, cell)
	if err != nil {
		return "", err
	}
	style, err := f.GetStyle(styleID)
	if err != nil {
		return "", err
	}
	// styleJSON, err := json.Marshal(style.Font)
	// fmt.Println(string(styleJSON))

	var styleStrings []string

	// Background color
	if len(style.Fill.Color) > 0 {
		bgColorHex := fmt.Sprintf("#%s", style.Fill.Color[0])
		styleStrings = append(styleStrings, fmt.Sprintf("background-color:%s;", bgColorHex))
		isLight, err := isColorLight(bgColorHex)
		if err != nil {
			log.Fatalf("Error: %v", err)
		}
		fontColor := "#000000"
		if !isLight {
			fontColor = "#FFFFFF"
		}

		styleStrings = append(styleStrings, fmt.Sprintf("color:%s;", fontColor))
	}

	// Font styles
	if style.Font != nil {
		if style.Font.Bold {
			styleStrings = append(styleStrings, "font-weight:bold;")
		}
		if style.Font.Italic {
			styleStrings = append(styleStrings, "font-style:italic;")
		}
		if style.Font.Underline != "" {
			styleStrings = append(styleStrings, "text-decoration:underline;")
		}
		styleStrings = append(styleStrings, fmt.Sprintf("font-family:%s;", style.Font.Family))
		styleStrings = append(styleStrings, fmt.Sprintf("font-size:%spt;", strconv.FormatFloat(style.Font.Size, 'f', 6, 64)))

	}
	fmt.Printf("%+v\n", style.Border)
	// Borders
	for _, border := range style.Border {
		if border.Style != 0 {
			styleStrings = append(styleStrings, fmt.Sprintf("border-%s:%dpx solid #%s;", border.Type, border.Style, border.Color))
		}
	}
	// if len(style.Border) == 0 {
	// 	styleStrings = append(styleStrings, fmt.Sprintf("border: 1px solid transparent;"))
	// }
	// Padding
	styleStrings = append(styleStrings, fmt.Sprintf("padding: 1pt 1px;"))

	// Alignment
	if style.Alignment != nil {
		if style.Alignment.Horizontal != "" {
			styleStrings = append(styleStrings, fmt.Sprintf("text-align:%s;", style.Alignment.Horizontal))
		}
		if style.Alignment.Vertical != "" {
			styleStrings = append(styleStrings, fmt.Sprintf("vertical-align:%s;", style.Alignment.Vertical))
		}
		if style.Alignment.Indent > 0 {
			if style.Alignment.Horizontal == "right" {
				styleStrings = append(styleStrings, fmt.Sprintf("padding-right:%dpx;", style.Alignment.Indent*12))
			} else if style.Alignment.Horizontal != "center" {
				styleStrings = append(styleStrings, fmt.Sprintf("padding-left:%dpx;", style.Alignment.Indent*12))
			}
		}
	}

	return strings.Join(styleStrings, ""), nil
}
func isMergedCell(f *excelize.File, sheetName, cell string) (bool, int, int) {
	mergedCells, err := f.GetMergeCells(sheetName)
	if err != nil {
		log.Fatalf("Failed to get merged cells: %v", err)
	}

	for _, mc := range mergedCells {
		if mc.GetStartAxis() == cell {
			startCol, startRow, _ := excelize.CellNameToCoordinates(mc.GetStartAxis())
			endCol, endRow, _ := excelize.CellNameToCoordinates(mc.GetEndAxis())
			colSpan := endCol - startCol + 1
			rowSpan := endRow - startRow + 1
			return true, colSpan, rowSpan
		}
	}
	return false, 1, 1
}

func getCellDimensions(f *excelize.File, sheetName string, col, row int) (float64, float64, error) {
	width, err := getColWidth(f, sheetName, col)
	if err != nil {
		return 0, 0, err
	}
	height, err := f.GetRowHeight(sheetName, row)
	if err != nil {
		return 0, 0, fmt.Errorf("Failed to get row height: %v", err)
	}

	return width, height, nil
}

func getColWidth(f *excelize.File, sheetName string, col int) (float64, error) {
	colName, err := excelize.ColumnNumberToName(col)
	if err != nil {
		return 0, fmt.Errorf("Failed to get column Name: %v", err)
	}

	width, err := f.GetColWidth(sheetName, colName)
	pixelWidth := width*7 + 5
	fmt.Println(pixelWidth)
	if err != nil {
		return 0, fmt.Errorf("Failed to get column width: %v", err)
	}
	return pixelWidth, nil
}
func Excel2HTML(fpath string) error {
	mergedCellsMap := make(map[string]bool)
	css := CssClassStyle{}
	maxCol := 0
	// Open the Excel file
	f, err := excelize.OpenFile(fpath)
	if err != nil {
		return fmt.Errorf("failed to open Excel file: %w", err)
	}

	// Get the first sheet's name
	sheetName := f.GetSheetName(0)
	if sheetName == "" {
		return fmt.Errorf("sheet not found")
	}

	// Get all the rows in the first sheet
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("failed to get rows: %w", err)
	}

	// Generate HTML table
	uuid := uuid.New()
	htmlTable := fmt.Sprintf("<body>\n<div id=\"%s\" align=center>\n<table border='0' style='border-collapse: collapse;'>\n", uuid)

	for rowIndex, row := range rows {
		heightAttr := ""
		if len(row) > maxCol {
			maxCol = len(row)
		}

		height, err := f.GetRowHeight(sheetName, rowIndex+1)
		if err != nil {
			log.Fatal(err)
		}
		if height > 0 {
			heightAttr += fmt.Sprintf(" height='%fpt;'", height+4.5*height/15)
		}
		htmlTable += fmt.Sprintf("<tr%s>\n", heightAttr)

		for colIndex, cell := range row {
			var (
				class     string
				style     string
				classAttr string
				widthAttr string
				spanAttr  string
				err       error
			)

			cellRef, _ := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)

			if mergedCellsMap[cellRef] {
				continue
			}

			style, err = getCellStyle(f, sheetName, cellRef)
			if err != nil {
				return fmt.Errorf("failed to get cell style: %w", err)
			}

			if style != "" {
				classId, ok := css[style]
				if !ok {
					classId = fmt.Sprintf("xxx%d-%d", rowIndex, colIndex)
					css[style] = classId
				}
				class += fmt.Sprintf("%s ", classId)
			}

			if class != "" {
				classAttr += fmt.Sprintf("class='%s'", class)
			}
			merged, colSpan, rowSpan := isMergedCell(f, sheetName, cellRef)

			if merged {
				if colSpan > 1 {
					spanAttr += fmt.Sprintf(" colspan='%d'", colSpan)
				}
				if rowSpan > 1 {
					spanAttr += fmt.Sprintf(" rowspan='%d'", rowSpan)
				}
				for i := 0; i < rowSpan; i++ {
					for j := 0; j < colSpan; j++ {
						mergedCellRef, _ := excelize.CoordinatesToCellName(colIndex+1+j, rowIndex+1+i)
						mergedCellsMap[mergedCellRef] = true
					}
				}
			}

			htmlTable += fmt.Sprintf("<td %s%s%s>%s</td>\n", classAttr, widthAttr, spanAttr, cell)
		}
		htmlTable += "</tr>\n"

	}

	if maxCol > 0 {
		htmlTable += "<colgroup>\n"
		for colIndex := 0; colIndex < maxCol; colIndex++ {
			widthAttr := ""
			width, err := getColWidth(f, sheetName, colIndex+1)
			if err != nil {
				log.Fatal(err)
			}
			if width > 0 {
				widthAttr += fmt.Sprintf(" width='%fpx;'", width)
			}
			htmlTable += fmt.Sprintf("<col%s></col>\n", widthAttr)
		}
		htmlTable += "</colgroup>\n"
	}

	htmlTable += "</table>\n</div>\n</body>\n"

	cssStyles := ""
	for key, value := range css {
		cssStyles += fmt.Sprintf(".%s{%s}\n", value, key)
	}

	htmlHead := fmt.Sprintf("<head>\n<style id=\"%s\">\n%s\n</style>\n</head>\n", uuid, cssStyles)
	html := fmt.Sprintf("<html>\n%s%s\n</html>", htmlHead, htmlTable)

	// Optionally, save the HTML to a file
	dir := filepath.Dir(fpath)
	filename := filepath.Base(fpath)
	ext := filepath.Ext(fpath)
	foutput := filepath.Join(dir, fmt.Sprintf("%s.html", filename[0:len(filename)-len(ext)]))
	file, err := os.Create(foutput)
	if err != nil {
		return fmt.Errorf("failed to create HTML file: %w", err)
	}
	defer file.Close()

	_, err = file.WriteString(html)
	if err != nil {
		return fmt.Errorf("failed to write HTML file: %w", err)
	}
	return nil
}

type CssClassStyle map[string]string

// func main() {
// 	fpath := "example.xlsx"
// 	err := Excel2HTML(fpath)
// 	if err != nil {
// 		log.Fatalln(err)
// 	}
// }

// hexToRGB converts a hex color string to its RGB components.
func hexToRGB(hex string) (int, int, int, error) {
	if len(hex) != 7 || hex[0] != '#' {
		return 0, 0, 0, fmt.Errorf("invalid hex color format")
	}

	r, err := strconv.ParseInt(hex[1:3], 16, 0)
	if err != nil {
		return 0, 0, 0, err
	}
	g, err := strconv.ParseInt(hex[3:5], 16, 0)
	if err != nil {
		return 0, 0, 0, err
	}
	b, err := strconv.ParseInt(hex[5:7], 16, 0)
	if err != nil {
		return 0, 0, 0, err
	}

	return int(r), int(g), int(b), nil
}

// isColorLight determines if a color is light based on its hex code.
func isColorLight(hex string) (bool, error) {
	r, g, b, err := hexToRGB(hex)
	if err != nil {
		return false, err
	}

	// Calculate brightness using the formula: (0.299*R + 0.587*G + 0.114*B)
	brightness := (0.299*float64(r) + 0.587*float64(g) + 0.114*float64(b))

	// A threshold of 128 is used to classify light vs dark
	return brightness > 128, nil
}
