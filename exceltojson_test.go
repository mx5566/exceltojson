package exceltojson

import (
	"os"
	"testing"
)

func TestExcelToJson(t *testing.T) {
	pwd, _ := os.Getwd()

	ExcelToJson(pwd, pwd + "/out")
}
