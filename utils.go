package main

import (
	"bytes"
	"fmt"
	"os"
	"sort"

	"github.com/xuri/excelize/v2"
)

// writeToExcel 负责将处理好的数据写入Excel文件
func writeToExcel(outputFilePath string, title string, results []fileResult, mapping map[string]string, defaultPrefix string) error {
	if len(results) == 0 {
		return ErrNoData
	}

	f := excelize.NewFile()
	sheetName := "Sheet1"

	// 1. 汇总所有文件中出现过的、独一无二的key，并进行排序
	allKeysSet := make(map[string]struct{})
	for _, result := range results {
		for key := range result.Counts {
			allKeysSet[key] = struct{}{}
		}
	}
	var sortedAllKeys []string
	for k := range allKeysSet {
		sortedAllKeys = append(sortedAllKeys, k)
	}
	sort.Strings(sortedAllKeys)

	// 如果所有文件中都没有数据，也返回错误
	if len(sortedAllKeys) == 0 {
		return ErrNoData
	}

	// --- 2. 新增：写入跨列居中的主标题行 (第1行) ---
	numDataCols := len(sortedAllKeys) + 1 // 数据列数 = key的数量 + 1 (文件名列)
	endCellCol, _ := excelize.ColumnNumberToName(numDataCols)

	// 合并第一行的所有单元格
	f.MergeCell(sheetName, "A1", fmt.Sprintf("%s1", endCellCol))

	// 创建并应用标题样式
	titleStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 14}, // 标题字体可以大一些
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})
	f.SetCellStyle(sheetName, "A1", fmt.Sprintf("%s1", endCellCol), titleStyle)
	f.SetCellValue(sheetName, "A1", title)

	// 3. 写入表头行
	// 设置样式
	headerStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})

	// 用于存储每列所需的最大宽度
	colWidths := make(map[int]float64)

	// 写入 "文件名" 表头 (A2)
	f.SetCellValue(sheetName, "A2", "文件名")
	f.SetCellStyle(sheetName, "A2", "A2", headerStyle)
	colWidths[1] = calculateApproxTextWidth("文件名") // 初始宽度

	// 写入所有 key 的表头 (B2, C2, ...)
	for i, key := range sortedAllKeys {
		colNum := i + 2 // 从第2列 (B) 开始
		colName, _ := excelize.ColumnNumberToName(colNum)

		name, ok := mapping[key]
		if !ok {
			name = defaultPrefix + key
		}
		headerText := fmt.Sprintf("%s (%s)", name, key)

		f.SetCellValue(sheetName, fmt.Sprintf("%s2", colName), headerText)
		f.SetCellStyle(sheetName, fmt.Sprintf("%s2", colName), fmt.Sprintf("%s2", colName), headerStyle)
		colWidths[colNum] = calculateApproxTextWidth(headerText)
	}

	// 4. 逐行写入每个文件的数据
	centeredStyle, _ := f.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{Horizontal: "center"}})
	for i, result := range results {
		rowNum := i + 3 // 从第3行开始

		// 写入文件名 (列 A)
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), result.FileName)
		// 动态更新文件名列的所需宽度
		fileNameWidth := calculateApproxTextWidth(result.FileName)
		if fileNameWidth > colWidths[1] {
			colWidths[1] = fileNameWidth
		}

		// 写入每个 key 对应的计数值
		for j, key := range sortedAllKeys {
			colNum := j + 2
			colName, _ := excelize.ColumnNumberToName(colNum)
			cellName := fmt.Sprintf("%s%d", colName, rowNum)

			// 如果当前文件没有这个key，则填0
			count, ok := result.Counts[key]
			if !ok {
				count = 0
			}
			f.SetCellValue(sheetName, cellName, count)
			f.SetCellStyle(sheetName, cellName, cellName, centeredStyle)
		}
	}

	// 5. 应用计算好的所有列的宽度
	for colNum, width := range colWidths {
		colName, _ := excelize.ColumnNumberToName(colNum)
		f.SetColWidth(sheetName, colName, colName, width)
	}

	// 6. 设置文档属性并保存
	now := nowISO8601()
	_ = f.SetDocProps(&excelize.DocProperties{Created: now, Modified: now, Creator: "deviceParser"})

	buffer := new(bytes.Buffer)
	if err := f.Write(buffer); err != nil {
		return fmt.Errorf("写入内存失败: %w", err)
	}

	if err := os.WriteFile(outputFilePath, buffer.Bytes(), 0644); err != nil {
		return fmt.Errorf("保存到磁盘失败: %w", err)
	}

	return nil
}
