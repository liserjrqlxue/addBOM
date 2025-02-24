package main

import (
	"fmt"
	"log"
	"math"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/xuri/excelize/v2"
)

// const
const (
	// 碱基损耗
	Wastage = 220

	// 初始值
	InitialRA   = 10
	InitialRB   = 10
	InitialDNTP = 1
	InitialRC   = 2
	InitialEF   = 2.5
	InitialH2O  = 100.0
	// 碱基A的偏移
	OffsetA = 4.2
)

type BaseBOM struct {
	Name string

	Base    int
	BaseFix int
	// R-A 10
	RA int
	// R-B 10
	RB int
	// dNTP 1
	DNTP int
	// R-C 2
	RC int
	// E-F A:6.8 CGT:2.6
	EF float64
	// ddH2O 100 - SUM A:70.2 CGT:74.4
	H2O float64
}

// NewBaseBOM creates a new BaseBOM with specified name and offset.
// It initializes the BaseBOM fields with default constants and adjusts
// the EF field using the given offset. The function returns a pointer
// to the created BaseBOM instance.
//
// Parameters:
//   - name: A string representing the name of the BaseBOM.
//   - base: A string representing the base of the BaseBOM.
//
// Returns:
// - *BaseBOM: A pointer to the created BaseBOM instance.

func NewBaseBOM(name, base string) *BaseBOM {
	var offset float64
	if base == "A" {
		offset = 1.0
	}
	return &BaseBOM{
		Name: name,
		Base: Wastage,
		RA:   InitialRA,
		RB:   InitialRB,
		DNTP: InitialDNTP,
		RC:   InitialRC,
		EF:   InitialEF + offset*OffsetA,
		H2O:  InitialH2O,
	}
}

func (bom *BaseBOM) fillBaseBOMColumn(xlsx *excelize.File, sheetName string, col, row int) {
	cellName := simpleUtil.HandleError(excelize.CoordinatesToCellName(col, row))
	xlsx.SetSheetCol(
		sheetName, cellName,
		&[]string{
			fmt.Sprintf("%.2fml", float64(bom.RA)/1000),
			fmt.Sprintf("%.2fml", float64(bom.RB)/1000),
			fmt.Sprintf("%.2fml", float64(bom.DNTP)/1000),
			fmt.Sprintf("%.2fml", float64(bom.RC)/1000),
			fmt.Sprintf("%.2fml", bom.EF/1000),
			fmt.Sprintf("%.2fml", bom.H2O/1000),
		},
	)
}

type BOM struct {
	Name string
	// 碱基BOM
	BaseBOM map[string]*BaseBOM
}

func NewBOM(name string) *BOM {
	return &BOM{
		Name: name,
		BaseBOM: map[string]*BaseBOM{
			"A": NewBaseBOM("A", "A"),
			"T": NewBaseBOM("T", "T"),
			"G": NewBaseBOM("G", "G"),
			"C": NewBaseBOM("C", "C"),
		},
	}
}

func (bom *BOM) addPrimer(seq string) {
	for _, v := range seq {
		switch v {
		case 'A':
			bom.BaseBOM["A"].Base++
		case 'C':
			bom.BaseBOM["C"].Base++
		case 'G':
			bom.BaseBOM["G"].Base++
		case 'T':
			bom.BaseBOM["T"].Base++
		case 'N':
			log.Println("skip base: N")
		default:
			log.Fatalf("unknown base: %c", v)
		}
	}
}

func (bom *BOM) LoadRows(rows [][]string, col, skipRow int) {
	for i := range rows {
		if i <= skipRow {
			continue
		}
		if len(rows[i]) == 0 {
			continue
		}
		primerSeq := rows[i][col]
		// 跳过空行
		if primerSeq == "" {
			continue
		}
		bom.addPrimer(primerSeq)
	}
}

func (bom *BOM) UpdateBaseBOM() {
	for _, v := range bom.BaseBOM {
		// 四舍五入修正10
		v.BaseFix = (v.Base + 5) / 10 * 10
		log.Printf("%s\t%d -> %d\n", v.Name, v.Base, v.BaseFix)

		v.RA *= v.BaseFix
		v.RB *= v.BaseFix
		v.DNTP *= v.BaseFix
		v.RC *= v.BaseFix
		v.EF *= float64(v.BaseFix)
		v.H2O *= float64(v.BaseFix)

		// 向上取整修正500
		v.EF = math.Ceil(v.EF/500) * 500
		v.H2O = v.H2O - (v.EF + float64(v.RA+v.RB+v.DNTP+v.RC))
	}
}

func (bom *BOM) fillBOMTitle(xlsx *excelize.File, sheetName string, col, row int) {
	cellName := simpleUtil.HandleError(excelize.CoordinatesToCellName(col, row))
	// 第一行
	xlsx.SetSheetRow(
		sheetName, cellName,
		&[]string{"BOM", "dATP", "dTTP", "dGTP", "dCTP"},
	)
	// 第一列
	xlsx.SetSheetCol(
		sheetName, cellName,
		&[]string{"BOM", "R-A", "R-B", "dNTP", "R-C", "E-F", "ddH2O"},
	)
}

func (bom *BOM) fillBOMColumns(xlsx *excelize.File, sheetName string, col, row int) {
	for i, n := range []string{"A", "T", "G", "C"} {
		bom.BaseBOM[n].fillBaseBOMColumn(xlsx, sheetName, col+i, row)
	}
}

func (bom *BOM) createBOMTable(xlsx *excelize.File, sheetName string, col, row int) {
	xlsx.NewSheet(sheetName)
	bom.fillBOMTitle(xlsx, sheetName, col, row)
	bom.fillBOMColumns(xlsx, sheetName, col+1, row+1)
}

// Report generates a BOM report based on the data extracted from an Excel sheet.
//
// This function performs the following operations:
// 1. Retrieves all rows from the specified Excel sheet.
// 2. Validates the header row against a predefined template.
// 3. Loads the sequence data from the rows, updating the BaseBOM counts accordingly.
// 4. Calculates the required quantities for each component in the BaseBOM.
// 5. Creates a new Excel sheet and fills it with the BOM table.
// 6. Applies styling to the generated BOM table.
//
// Parameters:
//   - xlsx: A pointer to the Excel file.
//   - sheetName: The name of the sheet where the BOM should be created.
//   - col: The column index where the DNA sequence data is located.
//   - titleRow: The row index of the header row in the input sheet.

func (bom *BOM) Report(xlsx *excelize.File, sheetName string, col, titleRow int) {
	var rows = simpleUtil.HandleError(xlsx.GetRows(ppoSheet))
	// 校验表头
	checkPPOTitle(rows[titleRow])

	bom.LoadRows(rows, col, titleRow)
	bom.UpdateBaseBOM()
	bom.createBOMTable(xlsx, sheetName, 1, 1)
	setBOMStyle(xlsx, sheetName)
}
