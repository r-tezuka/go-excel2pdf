package main

import (
	"fmt"
	"log"
	"os/exec"
	"path/filepath"
)

func main() {
	fmt.Println("PDFに変換しています")

	inPath, _ := filepath.Abs("./test.xlsx")
	outPath, _ := filepath.Abs("./test.pdf")

	stdout, err := exec.Command(
		// command header
		"powershell", "-NoProfile", "-ExecutionPolicy", "Unrestricted",

		// 変換開始
		"$inputPath = \""+inPath+"\";",
		"$fileitem = Get-Item $inputPath;",
		"$excel = New-Object -ComObject Excel.Application;",
		"$excel.Visible = $false;",
		"$excel.DisplayAlerts = $false;",
		"$wb = $excel.Workbooks.Open($fileitem.FullName);",

		// PDF形式で保存
		"$outputPath = \""+outPath+"\";",
		"$wb.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF, $outputPath);",

		// 後処理
		"$wb.Close();",
		"$excel.Quit();",
		"if ($sheet -ne $null) {[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet)};",
		"if ($wb -ne $null) {[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)};",
		"if ($excel -ne $null) {[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)};",
	).CombinedOutput()

	if err != nil {
		log.Fatal(err)
	}
	fmt.Println(string(stdout))
}
