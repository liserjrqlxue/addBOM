package main

import (
	"fmt"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/xuri/excelize/v2"
)

var (
	/* 	ppoTitle = []string{
	   		"编号",
	   		"板号",
	   		"well",
	   		"Primer名称(必填)",
	   		"序列(5'to3')（必填）",
	   		"碱基数",
	   	}
	*/
	ppoTitle = []string{
		"序号",
		"位置",
		"引物名称",
		"序列",
		"长度",
	}
)

// 校验引物订购单表头
func checkPPOTitle(titleRow []string) error {
	for j, v := range ppoTitle {
		if titleRow[j] != v {
			return fmt.Errorf("表头错误!:\t%d[%s]vs.[%s]", j, v, titleRow[j])
		}
	}
	return nil
}

func setBOMStyle(xlsx *excelize.File, sheetName string) {
	xlsx.SetColWidth(sheetName, "A", "A", 10)
	xlsx.SetColWidth(sheetName, "B", "E", 12)

	var titleStyle = simpleUtil.HandleError(xlsx.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	}))
	var baseStyle = simpleUtil.HandleError(xlsx.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "right",
		},
	}))

	xlsx.SetCellStyle(sheetName, "A1", "E7", titleStyle)
	xlsx.SetCellStyle(sheetName, "B2", "E7", baseStyle)
}
