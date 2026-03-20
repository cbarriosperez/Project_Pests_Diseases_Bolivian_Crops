$path = 'D:\OneDrive - CGIAR\Github_Agroclimatic Modeling Team\Project_Pests_Diseases_Bolivian_Crops\Work Plan Componente B5.xlsx'
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Open($path)
    $ws = $wb.Sheets.Item(1)
    $rows = $ws.UsedRange.Rows.Count
    $cols = $ws.UsedRange.Columns.Count

    for ($r = 1; $r -le $rows; $r++) {
        $line = ''
        for ($c = 1; $c -le $cols; $c++) {
            $cell = $ws.Cells.Item($r, $c).Text
            $spec = ""
            if ($cell -ne $null) {
                $spec = $cell.ToString()
            }
            $line += $spec + '|'
        }
        Write-Output $line
    }
    $wb.Close($false)
    $excel.Quit()
} catch {
    Write-Output "Error: $_"
}
