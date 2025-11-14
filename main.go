package main

import (
	"errors"
	"fmt"
	_ "image/png"
	"os"
	"path/filepath"
	"runtime/debug"
	"strings"
	"time"
	_ "time/tzdata"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/data/binding"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/widget"
)

// 全局状态：用于存储用户动态配置的映射

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
	data, err := iconFile.ReadFile("rsc/icon.png")
	if err != nil {
		panic(err)
	}
	icon := fyne.NewStaticResource("icon.png", data)
	myApp.SetIcon(icon)

	mainWindow := myApp.NewWindow("deviceParser")
	mainWindow.Resize(fyne.NewSize(800, 600))
	// --- UI 修改: 仍然使用List来显示待处理项 ---
	// 列表现在可以包含文件路径或一个文件夹路径
	itemListBinding := binding.NewStringList()

	itemListWidget := widget.NewListWithData(itemListBinding,
		func() fyne.CanvasObject { return widget.NewLabel("template item") },
		func(i binding.DataItem, o fyne.CanvasObject) { o.(*widget.Label).Bind(i.(binding.String)) },
	)

	// 添加复选框和汇总文件名输入框
	summarizeCheck := widget.NewCheck("将所有结果汇总到一个Excel文件", nil)
	summaryFileNameEntry := widget.NewEntry()
	summaryFileNameEntry.SetPlaceHolder("请输入汇总文件名 (如: summary_report)")
	summaryFileNameEntry.Hide() // 默认隐藏
	summarizeCheck.OnChanged = func(checked bool) {
		if checked {
			summaryFileNameEntry.Show()
			summaryFileNameEntry.Enable()
		} else {
			summaryFileNameEntry.Hide()
			summaryFileNameEntry.Disable()
		}
	}

	statusLabel := widget.NewLabel("请添加文件或文件夹进行处理")
	statusLabel.Alignment = fyne.TextAlignCenter

	// 新增文件按钮
	addFilesButton := widget.NewButton("添加文件", func() {
		fileDialog := dialog.NewFileOpen(func(reader fyne.URIReadCloser, err error) {
			if err != nil {
				dialog.ShowError(err, mainWindow)
				return
			}
			if reader == nil {
				return
			}

			// 每次选择都清空旧的，确保只处理一种模式
			itemListBinding.Set([]string{})
			itemListBinding.Append(reader.URI().Path())
			statusLabel.SetText("模式: 单文件处理")
		}, mainWindow)
		// ... dialog setup ...
		fileDialog.SetFilter(storage.NewExtensionFileFilter([]string{".txt"}))
		resizeDialog(fileDialog, mainWindow)
		fileDialog.Show()
	})
	addFilesButton.Importance = widget.WarningImportance

	// 新增文件夹按钮
	addFolderButton := widget.NewButton("选择文件夹", func() {
		folderDialog := dialog.NewFolderOpen(func(uri fyne.ListableURI, err error) {
			if err != nil {
				dialog.ShowError(err, mainWindow)
				return
			}
			if uri == nil {
				return
			}

			// 每次选择都清空旧的，确保只处理一种模式
			itemListBinding.Set([]string{uri.Path()})
			statusLabel.SetText("模式: 文件夹批量处理")
		}, mainWindow)
		resizeDialog(folderDialog, mainWindow)
		folderDialog.Show()
	})
	addFolderButton.Importance = widget.WarningImportance

	// 清空文件按钮
	clearAllButton := widget.NewButton("清空", func() {
		itemListBinding.Set([]string{})
		statusLabel.SetText("已清空，请重新添加")
	})
	clearAllButton.Importance = widget.DangerImportance

	// 选择输出目录按钮和输入框
	outputFolderEntry := widget.NewEntry()
	outputFolderEntry.SetPlaceHolder("请选择一个输出根目录...")
	selectOutputFolderButton := widget.NewButton("选择输出位置", func() {
		folderDialog := dialog.NewFolderOpen(func(uri fyne.ListableURI, err error) {
			if err != nil {
				dialog.ShowError(err, mainWindow)
				return
			}
			if uri == nil {
				return
			}
			outputFolderEntry.SetText(uri.Path())
		}, mainWindow)
		resizeDialog(folderDialog, mainWindow)
		folderDialog.Show()
	})
	selectOutputFolderButton.Importance = widget.WarningImportance

	// 修改前缀输入框和按钮
	prefixEntry := widget.NewEntry()
	prefixEntry.SetText("BIN")
	prefixChangeButton := widget.NewButton("修改前缀", func() {
		// 根据输入的文本作为默认前缀
		defaultPrefix = prefixEntry.Text
		dialog.ShowInformation("提示", "前缀修改成功", mainWindow)
	})
	prefixChangeButton.Importance = widget.MediumImportance

	// 配置映射按钮
	configButton := widget.NewButton("配置映射关系", func() {
		configButtonClickHandler(itemListBinding, mainWindow, myApp)
	})
	configButton.Importance = widget.MediumImportance

	// “开始处理”按钮的逻辑 (基本不变)
	processButton := widget.NewButton("开始处理", func() {
		items, _ := itemListBinding.Get()
		outputRootPath := outputFolderEntry.Text
		if len(items) == 0 || outputRootPath == "" {
			dialog.ShowError(fmt.Errorf("输入项或输出根目录不能为空"), mainWindow)
			return
		}

		inputPath := items[0]

		fileInfo, err := os.Stat(inputPath)
		if err != nil {
			dialog.ShowError(fmt.Errorf("无法访问路径: %w", err), mainWindow)
			return
		}

		var filesToProcess []string
		var finalOutputDir string

		// --- 根据输入是文件还是文件夹，决定处理列表和最终输出目录 ---
		if fileInfo.IsDir() {
			// 模式: 文件夹
			entries, err := os.ReadDir(inputPath)
			if err != nil { /* ... error handling ... */
			}
			for _, entry := range entries {
				if !entry.IsDir() && strings.HasSuffix(strings.ToLower(entry.Name()), ".txt") {
					filesToProcess = append(filesToProcess, filepath.Join(inputPath, entry.Name()))
				}
			}
			// 创建新的子文件夹作为输出目录
			newFolderName := fmt.Sprintf("%s_results", filepath.Base(inputPath))
			finalOutputDir = filepath.Join(outputRootPath, newFolderName)
		} else {
			// 模式: 单文件
			filesToProcess = append(filesToProcess, inputPath)
			finalOutputDir = outputRootPath // 直接输出到根目录
		}

		// 创建最终输出目录 (如果不存在)
		if err := os.MkdirAll(finalOutputDir, 0755); err != nil {
			dialog.ShowError(fmt.Errorf("创建输出目录失败: %w", err), mainWindow)
			return
		}

		statusLabel.SetText(fmt.Sprintf("开始处理 %d 个文件...", len(filesToProcess)))

		// --- 数据提取循环 ---
		var results []fileResult
		for i, fPath := range filesToProcess {
			statusLabel.SetText(fmt.Sprintf("正在提取数据: %d/%d", i+1, len(filesToProcess)))
			result, err := extractDataFromFile(fPath)
			if err != nil {
				if errors.Is(err, ErrNoData) {
					continue // 静默跳过没有数据的文件
				}
				dialog.ShowError(fmt.Errorf("意外错误: %w", err), mainWindow)
				return
			}
			results = append(results, result)
		}

		if len(results) == 0 {
			dialog.ShowInformation("提示", "所有文件中都没有提取到有效数据。", mainWindow)
			return
		}

		// 结果标题
		title := "处理结果"

		// 区分是否开启汇总
		if summarizeCheck.Checked {
			// **模式: 汇总**
			statusLabel.SetText("开始汇总处理...")
			summaryFileName := summaryFileNameEntry.Text
			if summaryFileName == "" {
				dialog.ShowError(errors.New("请输入汇总文件的名称！"), mainWindow)
				return
			}
			if !strings.HasSuffix(summaryFileName, ".xls") && !strings.HasSuffix(summaryFileName, ".xlsx") {
				summaryFileName = summaryFileName + ".xlsx"
			}

			outputFilePath := filepath.Join(outputRootPath, fmt.Sprintf("%s", summaryFileName))

			// 写入Excel
			err := writeToExcel(outputFilePath, title, results, dynamicBinNameMapping, defaultPrefix)
			if err != nil {
				dialog.ShowError(fmt.Errorf("写入汇总文件失败: %w", err), mainWindow)
				return
			}

			statusLabel.SetText("汇总处理完成！")
			dialog.ShowInformation("成功", fmt.Sprintf("所有文件已汇总处理完毕！\n结果保存在: %s", outputFilePath), mainWindow)
		} else {
			// ******************************************************
			// *** 模式: 独立文件 (这是您需要补充完整的部分) ***
			// ******************************************************

			// 在这个模式下，我们仍然需要循环，但每次只写入一个文件的结果
			for i, result := range results {
				statusLabel.SetText(fmt.Sprintf("正在处理: %d/%d", i+1, len(results)))

				fileNameNoExt := strings.TrimSuffix(result.FileName, filepath.Ext(result.FileName))
				outFileName := fmt.Sprintf("%s_result.xlsx", fileNameNoExt)
				outFilePath := filepath.Join(finalOutputDir, outFileName)

				// 每次调用写入函数时，只传入包含当前文件结果的切片
				err := writeToExcel(outFilePath, title, []fileResult{result}, dynamicBinNameMapping, defaultPrefix)
				if err != nil {
					if errors.Is(err, ErrNoData) {
						dialog.ShowError(fmt.Errorf("写入文件失败: %w", err), mainWindow)
						continue // 静默跳过没有数据的文件
					}
					dialog.ShowError(fmt.Errorf("意外错误: %w", err), mainWindow)
					return
				}
			}

			statusLabel.SetText(fmt.Sprintf("处理完成！共 %d 个文件。", len(results)))
			dialog.ShowInformation("成功", fmt.Sprintf("所有 %d 个文件已独立处理完毕！\n结果保存在: %s", len(results), finalOutputDir), mainWindow)
		}
	})
	processButton.Importance = widget.HighImportance

	// ******************************************************
	// 布局
	// ******************************************************

	topButtons := container.New(layout.NewGridLayout(3), addFilesButton, addFolderButton, clearAllButton)

	content := container.NewVBox(
		topButtons,
		widget.NewLabel("待处理项 (单个文件或文件夹):"),
		itemListWidget,
		summarizeCheck,
		summaryFileNameEntry,
		container.NewBorder(nil, nil, widget.NewLabel("输出文件夹："), selectOutputFolderButton, outputFolderEntry),
		container.NewBorder(nil, nil, widget.NewLabel("默认前缀:"), prefixChangeButton, prefixEntry),
		container.New(layout.NewGridLayout(2), configButton, processButton),
		statusLabel,
	)

	mainWindow.SetContent(content)
	mainWindow.ShowAndRun()
}
