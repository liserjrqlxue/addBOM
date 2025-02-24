package main

import (
	"flag"
	"log"
	"strings"

	"github.com/xuri/excelize/v2"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
)

// flag
var (
	input = flag.String(
		"i",
		"",
		"input xlsx",
	)
	output = flag.String(
		"o",
		"",
		"output xlsx",
	)
)

// const
const (
	// 生物合成物料清单 Biosynthetic BOM
	bomSheet = "生物合成物料清单"
	// 引物订购单 Primer purchase order
	ppoSheet = "引物订购单"
	// 引物订购单标题行，从0开始
	// ppoTitleRow = 16 - 1
	ppoTitleRow = 1 - 1
	// 引物订购单序列列，从0开始
	// primerSeqCol = 5 - 1
	primerSeqCol = 4 - 1
)

func init() {
	flag.Parse()
	if *input == "" {
		flag.Usage()
		log.Fatal("-i required!")
	}
	if *output == "" {
		*output = strings.TrimSuffix(*input, ".xlsx") + "_BOM.xlsx"
	}
}

func main() {
	// 打开xlsx
	var xlsx = simpleUtil.HandleError(excelize.OpenFile(*input))
	defer simpleUtil.DeferClose(xlsx)

	// 生成BOM，创建BOM表
	var bom = NewBOM(bomSheet)
	bom.Report(xlsx, bomSheet, primerSeqCol, ppoTitleRow)

	// 保存
	simpleUtil.CheckErr(xlsx.SaveAs(*output))
}
