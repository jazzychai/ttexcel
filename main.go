package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"os"
	"strings"
)

func main() {
	if len(os.Args) < 3 {
		fmt.Println("命令错误，正确的命令格式为: ttexcel.exe 文件名 sheet名")
		return
	}
	file := os.Args[1]      //命令行第一个参数是文件路径
	sheetName := os.Args[2] //命令行第二个参数是sheet名

	//打开excel
	f, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println("无法打开文件,", err.Error())
		return
	}
	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println("sheet名错误或未找到有用数据,", err.Error())
		return
	}

	//筛选符合条件的数据
	var newData []map[string]string
	for _, row := range rows {
		if len(row) < 5 {
			continue
		}
		d := row[3]
		if strings.Index(d, "铅") > -1 {
			e := row[4]
			newRow := make(map[string]string)
			newRow["d"] = d
			newRow["e"] = e
			newData = append(newData, newRow)
		}
	}

	//保存到新文件
	if len(newData) > 0 {
		nf := excelize.NewFile()
		nf.SetCellValue("Sheet1", "A1", "检测项目")
		nf.SetCellValue("Sheet1", "B1", "检测标准")
		for i, v := range newData {
			rowNum := i + 2
			nf.SetCellValue("Sheet1", fmt.Sprintf("A%v", rowNum), v["d"])
			nf.SetCellValue("Sheet1", fmt.Sprintf("B%v", rowNum), v["e"])

		}
		nf.SaveAs("转换后_" + file)
		fmt.Println("程序结束，转换完成")
	} else {
		fmt.Println("程序结束，未找到符合条件的数据")
	}
}
