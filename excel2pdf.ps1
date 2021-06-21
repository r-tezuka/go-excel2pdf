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
    $wb.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF, $pdfpath)

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