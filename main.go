package main

import (
	"flag"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

func main() {
	var (
		rowCount      int
		input, output string
	)

	flag.IntVar(&rowCount, "r", 14000, "切分行数")
	flag.StringVar(&input, "i", "", "要切分的文件")
	flag.StringVar(&output, "o", "", "输出目标目录")
	flag.Parse()
	if input == "" || output == "" {
		flag.Usage()
		return
	}

	fmt.Println(input, output)
	xlFile, err := xlsx.OpenFile(input)
	if err != nil {
		fmt.Println("打开XLSX文件出错", err)
		return
	}

	path := fmt.Sprintf("%s/%s_%s", output, strings.Replace(filepath.Base(input), ".", "_", -1), time.Now().Format("20060102150405"))
	_, err = os.Stat(path)
	if os.IsNotExist(err) {
		os.MkdirAll(path, 0777)
	}

	rows := make([]*xlsx.Row, 0, rowCount+1)
	for offset, sheet := range xlFile.Sheets {
		if sheet.MaxRow == 0 {
			continue
		}

		pages := int(math.Ceil(float64(sheet.MaxRow-1) / float64(rowCount)))
		for i := 0; i < pages; i++ {
			rows = rows[:0]
			rows = append(rows, sheet.Rows[0])
			begin := (i * rowCount) + 1
			end := begin + rowCount
			if end > sheet.MaxRow {
				rows = append(rows, sheet.Rows[begin:]...)
			} else {
				rows = append(rows, sheet.Rows[begin:end]...)
			}

			file := xlsx.NewFile()
			ns, err := file.AddSheet(sheet.Name)
			if err != nil {
				fmt.Println("add sheet to new file failed:", err)
				continue
			}
			ns.Cols = sheet.Cols
			ns.Rows = rows
			err = file.Save(fmt.Sprintf("%s/sheet%d_%s_%d.xlsx", path, offset+1, sheet.Name, i+1))
			if err != nil {
				fmt.Println(err)
			}
		}
	}
}
