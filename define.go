package main

import (
	"embed"
	"errors"
)

var (
	//go:embed rsc/icon.png
	iconFile embed.FS

	dynamicBinNameMapping = make(map[string]string)
	defaultPrefix         = "BIN"

	ErrNoData = errors.New("no data found")
)

// fileResult holds the parsed data from a single input file.
type fileResult struct {
	FileName string
	Lot      string
	Wafer    string
	Counts   map[string]int
}

// record 结构体 (保持不变)
type record struct {
	name  string
	key   string
	count int
}
