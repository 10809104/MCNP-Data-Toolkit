# 1. 取得所有包含 Final_Report.csv 的路徑
$csvFiles = Get-ChildItem -Filter "Final_Report.csv" -Recurse

if ($csvFiles.Count -eq 0) {
    Write-Host "找不到任何 Final_Report.csv！" -ForegroundColor Red
    exit
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # 2. 建立一個全新的 Excel 活頁簿
    $mainWorkbook = $excel.Workbooks.Add()
    $sheetCount = 1

    foreach ($file in $csvFiles) {
        $folderName = $file.Directory.Name
        Write-Host "正在整理並美化資料夾: $folderName" -ForegroundColor Cyan

        # 3. 開啟 CSV 檔案
        $tempWorkbook = $excel.Workbooks.Open($file.FullName)
        $tempSheet = $tempWorkbook.Worksheets.Item(1)

        # 4. 確保主活頁簿有足夠的工作表
        if ($sheetCount -gt $mainWorkbook.Worksheets.Count) {
            $mainWorkbook.Worksheets.Add([System.Reflection.Missing]::Value, $mainWorkbook.Worksheets.Item($mainWorkbook.Worksheets.Count)) | Out-Null
        }

        $targetSheet = $mainWorkbook.Worksheets.Item($sheetCount)
        $targetSheet.Name = $folderName

        # 5. 複製數據
        $tempSheet.UsedRange.Copy() | Out-Null
        $targetSheet.Paste($targetSheet.Range("A1")) | Out-Null

        # --- 樣式優化區 ---
        
        # 6. 設定字體為 Times New Roman
        $allCells = $targetSheet.UsedRange
        $allCells.Font.Name = "Times New Roman"
        $allCells.Font.Size = 11

        # 7. 第二列 (Row 2) 有值的地方上黃色底色
        # 遍歷第二列的每一個 Cell，如果不是空的就塗色
        $lastCol = $targetSheet.UsedRange.Columns.Count
        for ($col = 1; $col -le $lastCol; $col++) {
            $cell = $targetSheet.Cells.Item(2, $col)
            if ($cell.Value2 -ne $null -and $cell.Value2 -ne "") {
                $cell.Interior.ColorIndex = 6 # 6 代表黃色 (Standard Yellow)
            }
        }

        # 8. 自動調整欄寬
        $allCells.EntireColumn.AutoFit() | Out-Null

        # 關閉暫存 CSV
        $tempWorkbook.Close($false)
        $sheetCount++
    }

    # 9. 儲存最終成果
    $outputPath = Join-Path $pwd.Path "MRCP_Total_Summary.xlsx"
    $mainWorkbook.SaveAs($outputPath, 51)
    $mainWorkbook.Close()
    $excel.Quit()

    Write-Host "整合與美化完成！產出檔案：MRCP_Total_Summary.xlsx" -ForegroundColor Green
}
catch {
    Write-Host "發生錯誤: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($excel) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
}