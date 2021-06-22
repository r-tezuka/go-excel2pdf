param(
    # パラメーターにExcelファイルパスを指定
    [parameter(mandatory)][string]$filepath
)

$fileitem = Get-Item $filepath

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($fileitem.FullName)

    # 保存先PDFファイルパスを生成
    $pdfpath = $fileitem.DirectoryName + "\" + $fileitem.BaseName + ".pdf"

    # PDF形式で保存
    # $wb.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF, $pdfpath)

    # 印刷設定
    $Missing = [System.Reflection.Missing]::Value
    $From = $Missing
    $To = $Missing
    $Copies = 1
    $Preview = $false
    $ActivePrinter = "Microsoft Print to PDF"
    $PrintToFile = $Missing
    $Collate = $Missing
    $OutputFileName = $pdfpath
    $IgnorePrintAreas = $Missing

    # シートの倍率設定
    foreach ($sheet in $wb.Worksheets) {
        $sheet.PageSetup.Zoom = $false
        $sheet.PageSetup.FitToPagesWide = 1 
        $sheet.PageSetup.FitToPagesTall = 1
    }

    # 印刷
    # Ref: https://blog.deltabox.site/post/2019/07/print_excel_only_pages/
    $wb.PrintOut.Invoke(@($From, $To, $Copies, $Preview, $ActivePrinter, $PrintToFile, $Collate, $OutputFileName, $IgnorePrintAreas))
    
    $wb.Close()
    $excel.Quit()
}
finally {
    # オブジェクト解放
    $sheet, $wb, $excel | ForEach-Object {
        if ($_ -ne $null) {
            [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
        }
    }
}