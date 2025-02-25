# BOM Generator for XLSX Files

English | [中文](./README_cn.md)

A Go-based tool to generate a Biosynthetic Material Requirements Planning (BOM) from an Excel file containing primer orders.

## Description

This project reads data from an Excel file containing primer purchase order information and generates a new Excel file with a formatted BOM (Biological Material Requirements Planning). The tool processes the input data, calculates the required quantities of biological materials, and outputs a structured report in Excel format.

## Features

- Reads data from an existing XLSX file
- Processes primer sequences and calculates base counts
- Generates a new Excel file with a formatted BOM
- Supports custom styles for better readability
- Command-line interface for easy usage

## Prerequisites

- Go 1.20+ installed
- Required dependencies:
  - `github.com/xuri/excelize/v2`
  - `github.com/liserjrqlxue/goUtil`

## Installation

```bash
go build #  create addBOM
# or 
go install # to $GOPATH/bin/addBOM
```

## Usage

### Command-line Arguments

The tool accepts the following command-line arguments:

- `-i`: Path to the input XLSX file (required)
- `-o`: Path to the output XLSX file (optional)

If no output path is provided, the output will be saved as `<input_filename>_BOM.xlsx`.

### Example

```bash
go run main.go -i primer_orders.xlsx -o bom_output.xlsx
```

This command reads data from `primer_orders.xlsx`, processes it, and saves the BOM to `bom_output.xlsx`.

## Input File Structure

The input Excel file must contain a sheet named **"引物订购单" (Primer Purchase Order)** with the following columns:

| Column Index | Column Name       |
|--------------|-------------------|
| 0            | 序号              |
| 1            | 位置              |
| 2            | 引物名称          |
| 3            | 序列              |
| 4            | 长度              |

### Important Notes

- The input file must follow the exact column structure and naming convention.
- The tool assumes that the header row is at row 1.

## Output File Structure

The output Excel file will contain two sheets:

1. **"生物合成物料清单" (Biosynthetic BOM)**: This sheet contains the generated BOM with calculated quantities of biological materials required for the primers.
2. **"引物订购单" (Primer Purchase Order)**: This sheet is a copy of the input data for reference.

### Styles

The output file includes custom styles to improve readability:
- Highlighted headers
- Formatted numbers
- Alternating row colors

## Project Structure

```
.
├── main.go         # Main entry point
├── tools.go        # Utility functions (header checking, styling)
└── bom.go          # BOM generation logic
```

## Contributing

Contributions are welcome! If you have any suggestions or find any issues, please open an issue or submit a pull request.

## License

This project is MIT licensed. See the [LICENSE](./LICENSE) file for details.

## Contact

For questions or feedback, feel free to contact the maintainers:
- Email: [wangyaoshen@zhonghegene.com](mailto:wangyaoshen@zhonghegene.com)
- GitHub: [liserjrqlxue](https://github.com/liserjrqlxue)

---

**Note**: Ensure that your input Excel file matches the expected format and structure. If you encounter any issues, verify that the column names and indices match those defined in the code.