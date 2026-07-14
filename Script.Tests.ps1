Describe "Excel Comparison" {

    BeforeAll {

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $temp = Join-Path $env:TEMP "ExcelCompareTest"

        if (!(Test-Path $temp)) {
            New-Item $temp -ItemType Directory | Out-Null
        }

        $file1 = Join-Path $temp "file1.xlsx"
        $file2 = Join-Path $temp "file2.xlsx"

        #
        # Workbook 1
        #

        $wb = $excel.Workbooks.Add()

        $ws = $wb.Worksheets.Item(1)

        $ws.Name = "Sheet1"

        $ws.Cells.Item(1, 1).Value2 = "Name"
        $ws.Cells.Item(1, 2).Value2 = "Age"

        $ws.Cells.Item(2, 1).Value2 = "Alice"
        $ws.Cells.Item(2, 2).Value2 = 20

        $ws.Cells.Item(3, 1).Value2 = "Bob"
        $ws.Cells.Item(3, 2).Value2 = 30

        $wb.SaveAs($file1)
        $wb.Close()

        #
        # Workbook 2
        #

        $wb = $excel.Workbooks.Add()

        $ws = $wb.Worksheets.Item(1)

        $ws.Name = "Sheet1"

        $ws.Cells.Item(1, 1).Value2 = "Name"
        $ws.Cells.Item(1, 2).Value2 = "Age"

        $ws.Cells.Item(2, 1).Value2 = "Alice"
        $ws.Cells.Item(2, 2).Value2 = 21

        $ws.Cells.Item(3, 1).Value2 = "Bob"
        $ws.Cells.Item(3, 2).Value2 = 31

        $wb.SaveAs($file2)
        $wb.Close()

        $excel.Quit()

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    It "detects one changed value" {

        $answers = [System.Collections.Queue]::new()

        @(
            "1",
            $file1,
            $file2,
            "Sheet1",
            "Sheet1",
            "5"
        ) | ForEach-Object {
            $answers.Enqueue($_)
        }

        Mock Read-Host {
            if ($answers.Count -eq 0) {
                throw "Read-Host called more times than expected"
            }

            return $answers.Dequeue()
        }

        $result = & "./script.ps1" |
        Where-Object { $_ -is [hashtable] }

        $result | Should -Not -BeNullOrEmpty

        $first = ($result.getEnumerator() | Select-Object -First 1).Value

        $first.Column | Should -Be "Age"
        $first.File1_Value | Should -Be "20"
    }
}