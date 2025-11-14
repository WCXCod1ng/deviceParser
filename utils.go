package main

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"time"

	"github.com/xuri/excelize/v2"
)

// writeToExcel 负责将处理好的数据写入Excel文件
func writeToExcel(outputFilePath string, title string, counts map[string]int, mapping map[string]string, defaultPrefix string) error {
	var records []record
	var sortedKeys []string
	for k := range counts {
		sortedKeys = append(sortedKeys, k)
	}
	sort.Strings(sortedKeys)

	for _, key := range sortedKeys {
		name, ok := mapping[key]
		if !ok {
			name = defaultPrefix + key
		}
		records = append(records, record{name: name, key: key, count: counts[key]})
	}

	f := excelize.NewFile()
	sheetName := "Sheet1"

	numCols := len(records)
	if numCols == 0 {
		return ErrNoData
	}

	endCellCol, _ := excelize.ColumnNumberToName(numCols)
	f.MergeCell(sheetName, "A1", fmt.Sprintf("%s1", endCellCol))
	titleStyle, _ := f.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"}, Font: &excelize.Font{Bold: true, Size: 12}})
	f.SetCellStyle(sheetName, "A1", fmt.Sprintf("%s1", endCellCol), titleStyle)
	f.SetCellValue(sheetName, "A1", title)

	centeredStyle, _ := f.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"}})

	// --- 写入表头 (第二行) 和数据 (第三行) ---
	for i, rec := range records {
		colNum := i + 1
		colName, _ := excelize.ColumnNumberToName(colNum)

		// 准备要写入第二行和第三行的文本
		textForRow2 := fmt.Sprintf("%s (%s)", rec.name, rec.key)
		textForRow3 := fmt.Sprintf("%d", rec.count) // 将数字也转为字符串以便计算宽度

		// 写入第二行
		cell2 := fmt.Sprintf("%s2", colName)
		f.SetCellValue(sheetName, cell2, textForRow2)
		f.SetCellStyle(sheetName, cell2, cell2, centeredStyle)

		// 写入第三行
		cell3 := fmt.Sprintf("%s3", colName)
		f.SetCellValue(sheetName, cell3, rec.count)
		f.SetCellStyle(sheetName, cell3, cell3, centeredStyle)

		// --- 动态计算并设置列宽 ---
		// 1. 分别计算两行文本的估算宽度
		width2 := calculateApproxTextWidth(textForRow2)
		width3 := calculateApproxTextWidth(textForRow3)

		// 2. 取两者中的最大值作为此列的宽度
		maxWidth := width2
		if width3 > maxWidth {
			maxWidth = width3
		}

		// 3. 应用计算出的宽度
		f.SetColWidth(sheetName, colName, colName, maxWidth)
	}

	// --- 时区与 DocProps 设置 ---
	// 使用系统全局 time.Local（在 main 中已设置为 Asia/Shanghai）
	now := nowISO8601()
	// 显式设置文档属性，避免 fallback 时间问题
	_ = f.SetDocProps(&excelize.DocProperties{
		Created:     now,
		Modified:    now,
		Creator:     "deviceParser",
		Description: "RowData Results",
	})

	// --- 将内容写入临时英文路径，然后重命名到目标路径 ---
	buffer := new(bytes.Buffer)
	if _, err := f.WriteTo(buffer); err != nil {
		return fmt.Errorf("将Excel数据写入内存失败: %w", err)
	}

	// 临时文件名放在 os.TempDir() 下，尽量使用 ASCII 名称
	tmpFileName := fmt.Sprintf("tmp_excel_%d.xlsx", time.Now().UnixNano())
	tmpPath := filepath.Join(os.TempDir(), tmpFileName)

	// 写入临时文件
	if err := os.WriteFile(tmpPath, buffer.Bytes(), 0644); err != nil {
		return fmt.Errorf("保存临时Excel文件失败: %w", err)
	}

	// 重命名到目标路径（覆盖同名文件）
	if err := os.Rename(tmpPath, outputFilePath); err != nil {
		// 如果跨设备移动失败（rare），尝试拷贝后删除
		in, rerr := os.ReadFile(tmpPath)
		if rerr != nil {
			_ = os.Remove(tmpPath)
			return fmt.Errorf("重命名临时文件失败: %v，且临时文件读取失败: %w", err, rerr)
		}
		if werr := os.WriteFile(outputFilePath, in, 0644); werr != nil {
			_ = os.Remove(tmpPath)
			return fmt.Errorf("重命名失败: %v；拷贝到目标失败: %w", err, werr)
		}
		_ = os.Remove(tmpPath)
	}

	return nil
}
