package main

import (
	"fmt"
	"log"
	"os/exec"
	"path/filepath"
)

func main() {
	fmt.Println("PDFに変換しています")

	inPath, _ := filepath.Abs("./test2.xlsx")

	printPDF(inPath)
	// execCmdByScript(inPath, ".\\excel2pdfbyprinter.ps1")
}

func getFilePathWithoutExt(path string) string {
	return path[:len(path)-len(filepath.Ext(path))]
}

func exportPDF(inPath string) {
	outPath := getFilePathWithoutExt(inPath) + ".pdf"

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

func printPDF(inPath string) {
	outPath := getFilePathWithoutExt(inPath) + ".pdf"

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

		// 印刷設定
		"$Missing = [System.Reflection.Missing]::Value;",
		"$From = $Missing;",
		"$To = $Missing;",
		"$Copies = 1;",
		"$Preview = $false;",
		"$ActivePrinter = \"Microsoft Print to PDF\";",
		"$PrintToFile = $Missing;",
		"$Collate = $Missing;",
		"$OutputFileName = \""+outPath+"\";",
		"$IgnorePrintAreas = $Missing;",

		// 各シートの倍率設定
		"foreach ($sheet in $wb.Worksheets) { $sheet.PageSetup.Zoom = $false };",
		"foreach ($sheet in $wb.Worksheets) { $sheet.PageSetup.FitToPagesWide = 1 };",
		"foreach ($sheet in $wb.Worksheets) { $sheet.PageSetup.FitToPagesTall = 1 };",

		// 印刷
		"$wb.PrintOut.Invoke(@($From, $To, $Copies, $Preview, $ActivePrinter, $PrintToFile, $Collate, $OutputFileName, $IgnorePrintAreas));",

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

func execCmdByScript(path string, script string) {
	fmt.Println(path + " をPDFに変換中...")
	stdout, err := exec.Command("powershell", "-NoProfile", "-ExecutionPolicy", "Unrestricted", script, path).CombinedOutput()
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println(string(stdout))
}
