package  main

import (
	"bufio"
	"fmt"
	"os"
	"github.com/tealeg/xlsx"
)
/*
	1. 一行一行读配置
	2. 分析处理
	3. 导出为excel文件
*/

type HeroHorse struct {
	name string
	job string
	date string
	status string
}

func main() {
	file, err := os.Open("bole.txt")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer file.Close()

	var tmp []string
	scan := bufio.NewScanner(file)
	for scan.Scan() {
		lineText := scan.Text()
		if lineText == "撤销" {
			continue
		}
		//fmt.Println(lineText)
		tmp = append(tmp, lineText)
	}

	//pre handle
	fixNum := len(tmp)%4
	//fmt.Println("fixNum: ", fixNum)
	//fmt.Println("tmp len before:", len(tmp))
	for i := 0; i < (4-fixNum); i++ {
		tmp = append(tmp, "")
	}
	//fmt.Println("tmp len after:", len(tmp))
	//fmt.Println(tmp)
	//set
	var heroList []HeroHorse
	for i := 0; i < len(tmp); i=i+4 {
		hero := HeroHorse{
			name: tmp[i],
			job: tmp[i+1],
			date: tmp[i+2],
			status: tmp[i+3],
		}

		heroList = append(heroList, hero)
	}

	for _, hero := range heroList {
		fmt.Println(hero)
	}

	//xlsx
	writingXlsx(heroList)
}

func writingXlsx(heroList []HeroHorse) {
	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	var err error

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		fmt.Printf(err.Error())
		return
	}

	row = sheet.AddRow()
	row.SetHeightCM(0.5)
	cell = row.AddCell()
	cell.Value = "姓名"
	cell = row.AddCell()
	cell.Value = "岗位"
	cell = row.AddCell()
	cell.Value = "时间"
	cell = row.AddCell()
	cell.Value = "状态"

	for _, hero := range heroList {
		var row1 *xlsx.Row
		row1 = sheet.AddRow()
		row1.SetHeightCM(0.5)

		cell = row1.AddCell()
		cell.Value = hero.name
		cell = row1.AddCell()
		cell.Value = hero.job
		cell = row1.AddCell()
		cell.Value = hero.date
		cell = row1.AddCell()
		cell.Value = hero.status
	}

	err = file.Save("job.xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}
}