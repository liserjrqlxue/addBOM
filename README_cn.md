# 用于XLSX文件的物料清单生成器

[English](./README.md)| 中文

这是一个基于Go语言的工具，能够从包含引物订单信息的Excel文件中生成生物合成物料清单（BOM）。

## 项目描述

这是一个基于Go语言的工具，能够读取包含引物订单信息的Excel文件，并生成生物合成物料清单（BOM）。该工具处理输入数据，计算所需的生物材料数量，并输出结构化的报告，保存为Excel格式。

## 功能特点

- 从现有的XLSX文件中读取数据
- 处理引物序列并计算碱基数
- 生成包含格式化BOM的新Excel文件
- 支持自定义样式以提高可读性
- 提供命令行界面，使用便捷

## 预备条件

- 已安装Go 1.20+版本
- 所需依赖：
  - `github.com/xuri/urity/excelize/v2`
  - `github.com/liserjrqlxue/goUtil`

## 安装步骤

### 直接安装

```bash
go install liserjrqlxue/addBOM@latest
```

### 源码编译

```bash
git clone https://github.com/liserjrqlxue/addBOM.git
cd addBOM

go build #  create addBOM
# or 
go install # to $GOPATH/bin/addBOM
```

### 添加右键菜单

- 仅支持 `Windows` 系统
- 依赖 `powershell`
- 运行 `go install -ldflags="-H windowsgui`
  - 参数 `-ldflags="-H windowsgui"` 避免弹出terminal弹窗
- 管理员模式 `powershell` 运行 `./addRightClickMenu.ps1` 将 `addBOM` 添加到 `*.xlsx` 注册表右键菜单

## 使用方法

### 命令行参数

此工具接受以下命令行参数：

- `-i`：输入XLSX文件的路径（必填）
- `-o`：输出XLSX文件的路径（可选）

若未提供输出路径，生成的BOM文件将以 `<input_ilename>_BOM.xlsx` 的形式保存。

### 示例

```bash
go run main.go -i primer_ords.xlsx -o bom_output.xlsx
```

此命令从 `primer_ords.lsx` 文件读取数据，进行处理后，将生成的BOM保存为 `bom_output.xlsx`。

### 右键菜单运行

对输入 excel ``<input_ilename>.xlsx` 文件右键弹出右键菜单，选择addBOM运行，生成 `<input_ilename>_BOM.xlsx`

## 输入文件结构

输入Excel文件必须包含名为 **“引物订购单”** 的工作表，并具有以下列结构：

| 列索引 | 列名称       |
|-------|-------------|
| 0     | 序号        |
| 1     | 位置        |
| 2     | 引物名称    |
| 3     | 序列        |
| 4     | 长度        |

### 注意事项

- 输入文件必须严格遵循规定的列结构和命名规则。
- 工具假设表头位于第1行。

## 输出文件结构

生成的Excel文件包含两个工作表：

1. **“生物合成物料清单”**：此工作表包含计算出的各引物所需生物材料的数量。
2. **“引物订购单”**：此工作表为输入数据的副本，便于参考。

### 样式
输出文件采用自定义样式提升可读性：
- 高亮显示表头
- 数字格式化处理
- 行交替颜色

## 项目结构
```
.
├── main.go                 # 主程序入口
├── tools.go                # 工具函数集合
├── bom.go                  # BOM生成逻辑
└── addRightClickMenu.ps1   # PowerShell脚本用于添加右键菜单
```

## 贡献说明
我们欢迎任何建议和反馈！请通过邮箱或GitHub与我们联系。
```

## 许可证
此项目采用 MIT 协议，具体内容见 [LICENSE](./LICENSE) 文件。

## 联系方式
- 邮箱：[wangyaoshen@zhonghegene.com](mailto:wangyaoshen@zhonghegene.com)
- GitHub：[liserjrqlxue](https://github.com/liserjrqlxue)

---

**注意** ：请确保您的输入Excel文件与预期的格式和结构一致。如果遇到任何问题，请检查列名和索引是否符合代码中定义的内容。