package exceltojson

import (
	"encoding/json"
	"fmt"


	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"
)

// 获取指定目录下某个后缀的所有文件的列表
// 查找指定的文件列表
type FileFilter struct {
	// file list in directory
	ListFile []string
	// 后缀 eg:.go .xlsx .txt ...
	Suffix string
}

func (this *FileFilter) Listfunc(path string, f os.FileInfo, err error) error {
	var strRet string

	if f == nil {
		return err
	}
	if f.IsDir() {
		return nil
	}

	strRet += path

	//用strings.HasSuffix(src, suffix)//判断src中是否包含 suffix结尾
	ok := strings.HasSuffix(strRet, this.Suffix)
	if ok {
		this.ListFile = append(this.ListFile, strRet) //将目录push到listfile []string中
	}

	return nil
}

func (this *FileFilter) GetFileList(path, suffix string) error {
	this.Suffix = suffix
	//var strRet string
	err := filepath.Walk(path, this.Listfunc)

	if err != nil {
		log.Panicf("filepath.Walk() returned %v\n", err)
		return err
	}

	return nil
}

func ExcelToJson(pwdPath, outPath string) {
	log.Printf("in[%s] out[%s]\n", pwdPath, outPath)
	start(pwdPath, outPath)
}


func start(filepath string, outPath string) {
	var filter FileFilter
	_ = filter.GetFileList(filepath, ".xlsx")

	list := filter.ListFile

	for _, file := range list {
		readExcel(filepath, file, outPath)
	}

	finalMsg := fmt.Sprintf("Excel文件都已转换完成，共计包含%d个失败处理", errCount)

	if errCount != 0 {
		log.Println(finalMsg + "，请查看控制台日志记录")
	} else {
		log.Println(finalMsg)
	}
}


func checkFileIsExist(filename string) bool {
	var exist = true
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		exist = false
	}
	return exist
}

var errCount = 0
//读取excel
func readExcel(basePath string, file string, outPath string) {
	outFile := strings.Replace(file, basePath, outPath, 1)
	outFile = strings.Replace(outFile, ".xlsx", ".json", 1)

	var readErr error
	var wf *os.File
	outPaths, _ := filepath.Split(outFile)
	if checkFileIsExist(outFile) { //如果文件存在
		_ = os.Remove(outFile)
		wf, readErr = os.Create(outFile) //创建文件
	} else {
		_ = os.MkdirAll(outPaths, os.ModePerm)
		wf, readErr = os.Create(outFile) //创建文件
	}
	if readErr != nil {
		errCount++
		fmt.Printf("创建%s文件的写入流失败 %v\n", outFile, readErr)
		return
	}
	defer wf.Close()
	f, err := excelize.OpenFile(file)
	if err != nil {
		errCount++
		fmt.Printf("读取Excel失败：%s, %v\n", file, err)
		return
	}
	firstSheet := f.GetSheetList()[0]
	rows, _ := f.GetRows(firstSheet)
	dataDict := make([]interface{}, 0, 2000)
	keys := make([]string, 0, 50)
	for i, row := range rows {
		if i == 1 {
			for _, colCell := range row {
				keys = append(keys, colCell)
			}
		}
		if i == 2 || len(row) == 0 || i == 0 {
			continue
		}
		cells := make(map[string]string)
		for k, colCell := range row {
			if k >= len(keys) {
				break
			}

			if colCell == "" {
				continue
			}

			cells[keys[k]] = colCell
		}
		//检测字段是否全部为空
		isAppend := false
		for _, v := range cells {
			if v != "" {
				isAppend = true
				break
			}
		}
		if isAppend {
			dataDict = append(dataDict, cells)
		}
	}
	marshal, err := json.MarshalIndent(dataDict, "", "    ")
	if err != nil {
		errCount++
		log.Printf("转换JSON失败：%s, %v\n", file, err)
		return
	}
	_, writeErr := io.WriteString(wf, string(marshal))
	if writeErr != nil {
		errCount++
		log.Printf("写入文件失败失败：%s, %v\n", outFile, writeErr)
	}
}
