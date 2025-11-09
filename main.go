package main

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"runtime/debug"
	"sort"
	"strings"
	"time"
	_ "time/tzdata"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/data/binding"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/widget"

	"github.com/xuri/excelize/v2"
)

// 全局状态：用于存储用户动态配置的映射
var dynamicBinNameMapping = make(map[string]string)
var defaultPrefix = "BIN"

// record 结构体 (保持不变)
type record struct {
	name  string
	key   string
	count int
}

// 1. 新增函数：从文件中提取所有唯一的编号
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

func nowISO8601() string {
	loc, _ := time.LoadLocation("Asia/Shanghai")
	return time.Now().In(loc).Format("2006-01-02T15:04:05Z07:00")
}

// calculateApproxTextWidth 根据字符串内容估算其在Excel中的显示宽度
func calculateApproxTextWidth(text string) float64 {
	// 这是一个启发式计算：
	// ASCII字符（数字、字母、英文符号）宽度计为 1
	// 其他字符（如中文）宽度计为 2
	// 最后增加一些内边距
	width := 0.0
	for _, r := range text {
		if r <= 127 { // Is ASCII
			width += 1.0
		} else {
			width += 2.0
		}
	}
	return width + 3.0 // 增加 3 个字符的内边距，使显示效果更好
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
		return fmt.Errorf("没有数据可写入Excel")
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

	// 确保目标目录存在
	outDir := filepath.Dir(outputFilePath)
	if err := os.MkdirAll(outDir, 0755); err != nil {
		return fmt.Errorf("创建输出目录失败: %w", err)
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

// 2. 重构映射配置窗口
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

func resizeDialog(d dialog.Dialog, parent fyne.Window) {
	const minWidth float32 = 600
	const minHeight float32 = 400
	const scale float32 = 0.8 // 80%

	parentSize := parent.Canvas().Size()
	targetWidth := parentSize.Width * scale
	targetHeight := parentSize.Height * scale

	if targetWidth < minWidth {
		targetWidth = minWidth
	}
	if targetHeight < minHeight {
		targetHeight = minHeight
	}

	d.Resize(fyne.NewSize(targetWidth, targetHeight))
}

func handleCrash() {
	if r := recover(); r != nil {
		// 记录崩溃信息到文件
		logContent := fmt.Sprintf("FATAL ERROR: %v\n\nSTACK TRACE:\n%s", r, string(debug.Stack()))
		// 将日志文件放在程序同目录下，方便用户找到
		_ = os.WriteFile("crash.log", []byte(logContent), 0644)
	}
}

func main() {
	defer handleCrash()

	myApp := app.NewWithID("com.codingwang.deviceParser.v1")
	mainWindow := myApp.NewWindow("deviceParser")
	mainWindow.Resize(fyne.NewSize(800, 600))

	// --- UI 修改 1: 创建新的输入框 ---
	prefixEntry := widget.NewEntry()
	prefixEntry.SetText("BIN") // 设置默认值为 "BIN"

	inputFileEntry := widget.NewEntry()
	inputFileEntry.SetPlaceHolder("请选择一个 .txt 文件...")
	outputFolderEntry := widget.NewEntry()
	outputFolderEntry.SetPlaceHolder("请选择一个输出文件夹...")
	statusLabel := widget.NewLabel("请先选择输入文件和输出位置")
	statusLabel.Alignment = fyne.TextAlignCenter

	selectFileButton := widget.NewButton("选择输入文件", func() { /* ... */ })
	selectFolderButton := widget.NewButton("选择输出文件夹", func() { /* ... */ })
	prefixChangeButton := widget.NewButton("确定", func() {})

	// 省略了 selectFileButton 和 selectFolderButton 的内部实现代码
	selectFileButton.OnTapped = func() {
		fileDialog := dialog.NewFileOpen(func(reader fyne.URIReadCloser, err error) {
			if err != nil {
				dialog.ShowError(err, mainWindow)
				return
			}
			if reader == nil {
				return
			}
			inputFileEntry.SetText(reader.URI().Path())
			statusLabel.SetText("输入文件已选择")
		}, mainWindow)
		fileDialog.SetFilter(storage.NewExtensionFileFilter([]string{".txt"}))
		//fileDialog.Resize(fyne.NewSize(600, 400))
		resizeDialog(fileDialog, mainWindow)
		fileDialog.Show()
	}
	selectFolderButton.OnTapped = func() {
		folderDialog := dialog.NewFolderOpen(func(uri fyne.ListableURI, err error) {
			if err != nil {
				dialog.ShowError(err, mainWindow)
				return
			}
			if uri == nil {
				return
			}
			outputFolderEntry.SetText(uri.Path())
			statusLabel.SetText("输出文件夹已选择")
		}, mainWindow)
		//folderDialog.Resize(fyne.NewSize(600, 400))
		resizeDialog(folderDialog, mainWindow)
		folderDialog.Show()
	}

	prefixChangeButton.OnTapped = func() {
		// 根据输入的文本作为默认前缀
		defaultPrefix = prefixEntry.Text
	}

	// 3. 修改“配置映射”按钮的逻辑
	// “配置映射”按钮逻辑修改
	configButton := widget.NewButton("配置映射关系", func() {
		inputPath := inputFileEntry.Text
		if inputPath == "" {
			dialog.ShowError(fmt.Errorf("请先选择一个输入文件！"), mainWindow)
			return
		}
		keys, err := extractKeysFromFile(inputPath)
		if err != nil {
			dialog.ShowError(err, mainWindow)
			return
		}
		if len(keys) == 0 {
			dialog.ShowInformation("提示", "在所选文件中没有找到任何可配置的编号。", mainWindow)
			return
		}

		for _, key := range keys {
			// 使用从输入框读取的前缀来设置默认值
			dynamicBinNameMapping[key] = defaultPrefix + key
		}
		showMappingEditor(myApp, keys)
	})

	// “开始处理”按钮的逻辑 (基本不变)
	processButton := widget.NewButton("开始处理", func() {
		inputPath := inputFileEntry.Text
		outputPath := outputFolderEntry.Text
		if inputPath == "" || outputPath == "" {
			dialog.ShowError(fmt.Errorf("输入文件或输出文件夹不能为空"), mainWindow)
			return
		}

		statusLabel.SetText("正在处理中，请稍候...")

		originalFileName := filepath.Base(inputPath)
		fileExt := filepath.Ext(originalFileName)
		fileNameWithoutExt := strings.TrimSuffix(originalFileName, fileExt)
		newOutputFileName := fmt.Sprintf("%s_result.xlsx", fileNameWithoutExt)
		outputFilePath := filepath.Join(outputPath, newOutputFileName)

		err := processDataLogic(inputPath, outputFilePath, dynamicBinNameMapping)
		if err != nil {
			statusLabel.SetText("处理失败！")
			dialog.ShowError(err, mainWindow)
			return
		}

		statusLabel.SetText("处理完成！")
		dialog.ShowInformation("成功", fmt.Sprintf("文件处理成功！\n结果已保存至: %s", outputFilePath), mainWindow)
	})

	content := container.NewVBox(
		container.NewBorder(nil, nil, nil, selectFileButton, inputFileEntry),
		container.NewBorder(nil, nil, nil, selectFolderButton, outputFolderEntry),
		container.NewBorder(nil, nil, widget.NewLabel("默认名称前缀:"), prefixChangeButton, prefixEntry), // 新增的一行
		configButton,
		processButton,
		statusLabel,
	)

	mainWindow.SetContent(content)
	mainWindow.ShowAndRun()
}
