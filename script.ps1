$stage = "menu"
$running = $true

function handleInput([string]$message) {
    $in = Read-Host $message

    if ([string]::IsNullOrWhiteSpace($in)) {
        Write-Host "Canceled.`n"
        $global:stage = "menu"
        continue
    }
    return $in
}

$handleError = {
    param([boolean]$con, [string]$message)

    if ($con) {
        Write-Host $message `n
        continue
    }
}

while ($running) {
    if ($stage -eq "menu") {
        Write-Host "Press enter at any input to return back to this menu"
        Write-Host "`n------------------------------------------------------------`n"
        Write-Host "1   Compare a new file"
        if ($diff_map -and $diff_map.Count -gt 0) {
            Write-Host "2   Show difference overview"
            Write-Host "3   List all differences by difference id"
            Write-Host "4   Merge differences by difference id"
        }
        Write-Host "5   Quit"
        Write-Host "`n------------------------------------------------------------`n"

        $option = [int](handleInput "Enter an option")
        & $handleError ($option -lt 1 -or $option -gt 5) "Enter a valid option!"

        switch ($option) {
            "1" { $stage = "file" }
            "5" { $stage = "cleanup" }
        }
        if ($diff_map -and $diff_map.Count -gt 0) {
            switch ($option) {
                "2" { $stage = "overview" }
                "3" { $stage = "list" }
                "4" { $stage = "merge" }
            }
        }
    }

    if ($stage -eq 'file') {
        $file1 = handleInput "Enter path to first Excel file"
        & $handleError (!(Test-Path $file1)) 'File not found! Enter a valid path!'

        $file2 = handleInput "Enter path to second Excel file"
        & $handleError (!(Test-Path $file2)) 'File not found! Enter a valid path!'

        Write-Host "`nAccessing Excel Data..."
                
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.AskToUpdateLinks = $false
        $excel.DisplayAlerts = $false

        try {
            $wb1 = $excel.Workbooks.Open($file1)
            $wb2 = $excel.Workbooks.Open($file2) 
        }
        catch {
            Write-Host "Failed to open Workbook!"
        }

        $stage = "sheet"
    }

    if ($stage -eq "sheet") {
        Write-Host "`nDetected sheets for File1:"
        foreach ($sheet in $wb1.Sheets) {
            Write-Host $sheet.name
        }

        $sheet1 = handleInput "`nEnter worksheet name for File 1"
        $ws1 = $wb1.Worksheets.Item($sheet1)
        & $handleError (!$ws1.Name) "Sheet $sheet1 not found! Enter a valid sheet name!"

        Write-Host "`nDetected sheets for File2:"
        foreach ($sheet in $wb2.Sheets) {
            Write-Host $sheet.name
        }

        $sheet2 = Read-Host "`nEnter worksheet name for File 2 (press Enter to use same)"
        if ([string]::IsNullOrWhiteSpace($sheet2)) { $sheet2 = $sheet1 }
        $ws2 = $wb2.Worksheets.Item($sheet2)
        & $handleError (!$ws2.Name) "Sheet $sheet2 not found! Enter a valid sheet name!"

        $stage = "evaluation"
    }

    if ($stage -eq "evaluation") {
        Write-Host "`nReading Excel Data..."

        $rows = [Math]::Max($ws1.UsedRange.Rows.Count, $ws2.UsedRange.Rows.Count)
        $cols = [Math]::Max($ws1.UsedRange.Columns.Count, $ws2.UsedRange.Columns.Count)

        Write-Host "Comparing $rows rows and $cols columns...`n"

        $diff_map = @{}
        $sha1 = [System.Security.Cryptography.SHA1]::Create()

        $data1 = $ws1.UsedRange.Value2
        $data2 = $ws2.UsedRange.Value2

        for ($c = 1; $c -le $cols; $c++) {
            for ($r = 1; $r -le $rows; $r++) {
                $value1 = $data1[$r, $c]
                $value2 = $data2[$r, $c]

                if ($value1 -ne $value2) {
                    $key = "$c-$value1-$value2"
                    $hashBytes = $sha1.ComputeHash([Text.Encoding]::UTF8.GetBytes($key))

                    $key = [Convert]::ToBase64String($hashBytes[0..7])

                    if ($diff_map.ContainsKey($key)) {
                        $diff_map[$key].Count++
                        $diff_map[$key].Rows += $r
                    }
                    else {
                        $diff_map[$key] = [PSCustomObject]@{
                            ID          = $key
                            Count       = 1
                            ColNumber   = $c
                            Column      = $ws1.Cells.Item(1, $c).Value2
                            Rows        = [int[]]($r)
                            File1_Value = $value1
                            File2_Value = $value2
                        }
                    }
                }
            }
        }

        $stage = "overview"
    }

    if ($stage -eq "overview") {

        if ($diff_map.Count -eq 0) {
            Write-Output @()

            Write-Host "No differences found!`n"
        }
        else {
            Write-Output $diff_map

            $diff_map.Values |
            Sort-Object Count -Descending |
            Select-Object ID, Count, Column, File1_Value, File2_Value, Rows |
            Format-Table -AutoSize
        }

        $stage = "menu"
    }

    if ($stage -eq "list" -or $stage -eq "merge") {
        $id = handleInput "Enter an id (press Enter to cancel)"
        & $handleError (!$diff_map[$id]) 'Enter a valid id!'

        $diff = $diff_map[$id]
        $c = $diff_map[$id].ColNumber

        $data1 = $ws1.UsedRange.Value2
        $data2 = $ws2.UsedRange.Value2

        if ($stage -eq "list") {
            $list = foreach ($r in $diff.Rows) {
                $value1 = $data1[$r, $c]
                $value2 = $data2[$r, $c]

                [PSCustomObject]@{
                    Row         = $r
                    File1_Value = $value1
                    File2_Value = $value2
                }
            }

            Write-Host "`nColumn: $($diff.Column)"
            $list | Format-Table -AutoSize

            $stage = "menu"
        }

        if ($stage -eq "merge") {
            $diff |
            Format-Table -AutoSize
            Select-Object ID, Count, Column, File1_Value, File2_Value, Rows

            Write-Host "1   Accept File1 Value '$($diff.File1_Value)'"
            Write-Host "2   Accept File2 Value '$($diff.File2_Value)'`n"

            $in = handleInput "Enter an option (press Enter to cancel)"
            & $handleError (!($in -eq "1" -or $in -eq "2")) "Invalid input, merging canceled!"

            if ($in -eq "1") {
                $targetWs = $ws2
                $targetWb = $wb2
                $val = $diff.File1_Value
            }
            elseif ($in -eq "2") {
                $targetWs = $ws1
                $targetWb = $wb1
                $val = $diff.File2_Value
            }

            Write-Host "Overwriting differences..."

            foreach ($r in $diff.Rows) {
                $targetWs.Cells.Item($r, $c).Value2 = $val
            }

            $targetWb.Save()

            Write-Host "Changes saved! Reevaluating differences..."
            $stage = "evaluation"
        }
        $id = $null
    }

    if ($stage -eq "cleanup") {
        $diff_map = $null

        if ($excel) {
            Write-Host "`nTerminating Excel sessions..."
            $wb1.Close($false)
            $wb2.Close($false)
            $excel.Quit()

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws1) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws2) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb1) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb2) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

            [GC]::Collect()
            [GC]::WaitForPendingFinalizers() 
        }
        
        if ($sha1) { $sha1.Dispose() }

        $excel = $null
        $running = $false
        Write-Host "Cleanup done. Have a good day!`n"
    }
}
