package main

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/data/binding"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

// extractKeysFromFile 从文件中提取所有唯一的编号
func extractKeysFromFile(filePath string) ([]string, error) {
	content, err := os.ReadFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("读取文件失败: %w", err)
	}

	lines := strings.Split(string(content), "\n")
	keySet := make(map[string]struct{}) // 使用map来保证key的唯一性

	for _, line := range lines {
		cleanLine := strings.TrimSpace(line)
		if strings.HasPrefix(cleanLine, "RowData:") {
			fields := strings.Fields(strings.TrimPrefix(cleanLine, "RowData:"))
			for _, field := range fields {
				if field != "___" {
					keySet[field] = struct{}{}
				}
			}
		}
	}

	// 将map的key转换为slice并排序
	keys := make([]string, 0, len(keySet))
	for k := range keySet {
		keys = append(keys, k)
	}
	sort.Strings(keys)

	return keys, nil
}

// extractDataFromFile 只负责从单个文件中提取数据
func extractDataFromFile(filePath string) (lot, wafer string, counts map[string]int, err error) {
	content, err := os.ReadFile(filePath)
	if err != nil {
		return "", "", nil, fmt.Errorf("读取文件失败: %w", err)
	}

	isNumericRegex, _ := regexp.Compile(`^\d+$`)
	lines := strings.Split(string(content), "\n")
	counts = make(map[string]int)

	for _, line := range lines {
		cleanLine := strings.TrimSpace(line)
		if strings.HasPrefix(cleanLine, "LOT:") {
			lot = strings.TrimSpace(strings.TrimPrefix(cleanLine, "LOT:"))
		} else if strings.HasPrefix(cleanLine, "WAFER:") {
			wafer = strings.TrimSpace(strings.TrimPrefix(cleanLine, "WAFER:"))
		} else if strings.HasPrefix(cleanLine, "RowData:") {
			fields := strings.Fields(strings.TrimPrefix(cleanLine, "RowData:"))
			for _, field := range fields {
				if isNumericRegex.MatchString(field) {
					counts[field]++
				}
			}
		}
	}

	if len(counts) == 0 {
		return lot, wafer, counts, ErrNoData
	}
	return lot, wafer, counts, nil
}

func configButtonClickHandler(itemListBinding binding.List[string], parentWindow fyne.Window, myApp fyne.App) {
	items, _ := itemListBinding.Get()
	if len(items) == 0 {
		dialog.ShowError(fmt.Errorf("请先添加文件或选择一个文件夹！"), parentWindow)
		return
	}

	inputPath := items[0] // 第一个（也是唯一一个）项是文件或文件夹

	// 检查是文件还是文件夹
	fileInfo, err := os.Stat(inputPath)
	if err != nil {
		dialog.ShowError(fmt.Errorf("无法访问路径: %w", err), parentWindow)
		return
	}

	var filesToScan []string
	if fileInfo.IsDir() {
		// 如果是文件夹, 扫描所有.txt文件
		entries, err := os.ReadDir(inputPath)
		if err != nil {
			dialog.ShowError(fmt.Errorf("读取文件夹失败: %w", err), parentWindow)
			return
		}
		for _, entry := range entries {
			if !entry.IsDir() && strings.HasSuffix(strings.ToLower(entry.Name()), ".txt") {
				filesToScan = append(filesToScan, filepath.Join(inputPath, entry.Name()))
			}
		}
	} else {
		// 如果是文件
		filesToScan = append(filesToScan, inputPath)
	}

	if len(filesToScan) == 0 {
		dialog.ShowInformation("提示", "所选路径下没有找到 .txt 文件。", parentWindow)
		return
	}

	// 汇总所有文件的 keys
	allKeysSet := make(map[string]struct{})
	for _, filePath := range filesToScan {
		keys, err := extractKeysFromFile(filePath)
		if err != nil {
			dialog.ShowError(fmt.Errorf("解析文件 %s 失败: %w", filePath, err), parentWindow)
			return
		}
		for _, k := range keys {
			allKeysSet[k] = struct{}{}
		}
	}

	var sortedAllKeys []string
	for k := range allKeysSet {
		sortedAllKeys = append(sortedAllKeys, k)
	}
	sort.Strings(sortedAllKeys)
	for _, key := range sortedAllKeys {
		dynamicBinNameMapping[key] = defaultPrefix + key
	}
	showMappingEditor(myApp, sortedAllKeys)
}

// showMappingEditor 映射配置窗口
func showMappingEditor(a fyne.App, keys []string) {
	editorWindow := a.NewWindow("配置编号和名称的映射关系")
	editorWindow.Resize(fyne.NewSize(500, 400))

	listData := binding.NewStringList()

	// 刷新列表显示（例如 "001" 或 "001 -> BIN1"）
	refreshList := func() {
		var items []string
		for _, k := range keys {
			if name, ok := dynamicBinNameMapping[k]; ok {
				items = append(items, fmt.Sprintf("%s -> %s", k, name))
			} else {
				items = append(items, k)
			}
		}
		listData.Set(items)
	}

	selectedKey := ""
	selectedKeyLabel := widget.NewLabel("请从左侧列表选择一个编号")
	nameEntry := widget.NewEntry()
	nameEntry.SetPlaceHolder("为此编号指定一个名称...")
	nameEntry.Disable() // 默认禁用，选中后再启用

	list := widget.NewListWithData(listData,
		func() fyne.CanvasObject {
			return widget.NewLabel("template")
		},
		func(i binding.DataItem, o fyne.CanvasObject) {
			o.(*widget.Label).Bind(i.(binding.String))
		},
	)

	list.OnSelected = func(id widget.ListItemID) {
		selectedKey = keys[id] // 直接通过索引获取原始key，更可靠
		selectedKeyLabel.SetText(fmt.Sprintf("为编号 [%s] 设置名称:", selectedKey))

		// 如果已存在映射，则预填入输入框
		if name, ok := dynamicBinNameMapping[selectedKey]; ok {
			nameEntry.SetText(name)
		} else {
			nameEntry.SetText("")
		}
		nameEntry.Enable()
	}
	list.OnUnselected = func(id widget.ListItemID) {
		selectedKey = ""
		selectedKeyLabel.SetText("请从左侧列表选择一个编号")
		nameEntry.SetText("")
		nameEntry.Disable()
	}

	saveButton := widget.NewButton("保存", func() {
		if selectedKey != "" {
			dynamicBinNameMapping[selectedKey] = nameEntry.Text
			refreshList()      // 刷新列表以显示更新后的映射
			list.UnselectAll() // 清除选择状态
		}
	})

	closeButton := widget.NewButton("关闭", func() {
		editorWindow.Close()
	})

	// 右侧的编辑面板
	editorPanel := container.NewVBox(
		selectedKeyLabel,
		nameEntry,
		saveButton,
	)

	content := container.NewHSplit(
		list,
		container.NewBorder(nil, closeButton, nil, nil, editorPanel),
	)
	content.SetOffset(0.4) // 左侧列表占40%宽度

	refreshList() // 初始加载列表
	editorWindow.SetContent(content)
	editorWindow.Show()
}

// processDataLogic 函数 (保持不变)
func processDataLogic(inputFilePath string, outputFilePath string, mapping map[string]string) error {
	// ... 此函数内部逻辑与上一版本完全相同 ...
	content, err := os.ReadFile(inputFilePath)
	if err != nil {
		return fmt.Errorf("读取文件失败: %w", err)
	}
	lines := strings.Split(string(content), "\n")
	var lot, wafer string
	counts := make(map[string]int)
	for _, line := range lines {
		cleanLine := strings.TrimSpace(line)
		if strings.HasPrefix(cleanLine, "LOT:") {
			lot = strings.TrimSpace(strings.TrimPrefix(cleanLine, "LOT:"))
		} else if strings.HasPrefix(cleanLine, "WAFER:") {
			wafer = strings.TrimSpace(strings.TrimPrefix(cleanLine, "WAFER:"))
		} else if strings.HasPrefix(cleanLine, "RowData:") {
			fields := strings.Fields(strings.TrimPrefix(cleanLine, "RowData:"))
			for _, field := range fields {
				if field != "___" && field != "..." {
					counts[field]++
				}
			}
		}
	}
	var records []record
	var sortedKeys []string
	for k := range counts {
		sortedKeys = append(sortedKeys, k)
	}
	sort.Strings(sortedKeys)

	// --- 这里是唯一的、关键的修改点 ---
	for _, key := range sortedKeys {
		// 首先，尝试从用户配置的映射中获取名称
		name, ok := mapping[key]

		// 如果在映射中找不到 (ok == false)，则应用默认命名规则
		if !ok {
			name = defaultPrefix + key
		}

		// 此时，name 要么是用户的自定义名称，要么是 "BIN" + key
		records = append(records, record{name: name, key: key, count: counts[key]})
	}

	f := excelize.NewFile()
	defer f.Close()
	sheetName := "Sheet1"
	title := "未知"
	if lot != "" && wafer != "" {
		title = fmt.Sprintf("扩散批号：%s-%s", lot, wafer)
	}
	if len(records) > 0 {
		endCell, _ := excelize.CoordinatesToCellName(len(records), 1)
		f.MergeCell(sheetName, "A1", endCell)
	}

	// 计算总列数
	numCols := len(records)
	if numCols == 0 {
		return ErrNoData
	}

	// 合并单元格，横跨所有数据列
	// 如果 numCols > 0, endCell 至少是 "A1"
	endCellCol, _ := excelize.ColumnNumberToName(numCols)        // 将列号转换为列字母，例如 1 -> A, 2 -> B
	f.MergeCell(sheetName, "A1", fmt.Sprintf("%s1", endCellCol)) // A1 到最后一列的1行

	// 设置标题内容和样式（居中）
	// 注意：这里的style只针对标题单元格有效，如果需要标题文字居中同时跨多列，需要设置Alignment
	titleStyle, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center", // 水平居中
			Vertical:   "center", // 垂直居中
		},
		Font: &excelize.Font{Bold: true, Size: 12}, // 标题字体加粗，大小可调
	})
	f.SetCellStyle(sheetName, "A1", fmt.Sprintf("%s1", endCellCol), titleStyle)
	f.SetCellValue(sheetName, "A1", title)

	// --- 创建居中样式，用于第二行和第三行 ---
	centeredStyle, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center", // 水平居中
			Vertical:   "center", // 垂直居中
		},
	})

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
