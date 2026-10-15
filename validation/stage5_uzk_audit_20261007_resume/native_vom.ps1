$ErrorActionPreference='Stop'
$inputPath='C:\Users\arara\Desktop\ВЗМК рабочий стол\Работа\Документы\Гараж казахстан\Паспорт\9 от 06.10.2026\ВОМ отгрузочный счет 183  (5).xls'
$excel=$null;$book=$null
$result=[ordered]@{path=$inputPath;before_sha=(Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant();status='STARTED';macro_security=3;sheets=@()}
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3;$excel.EnableEvents=$false;$excel.AskToUpdateLinks=$false
    $book=$excel.Workbooks.Open($inputPath,0,$true)
    if(-not $book.ReadOnly){throw 'STOP: VOM not read-only'}
    $result.readonly=$book.ReadOnly
    $result.format=$book.FileFormat
    $registryData=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'registry_structure.json') -Raw | ConvertFrom-Json
    $requiredMarks=@($registryData.sheets[0].raw_rows | Where-Object {$_.values.A -eq 'Гараж Казахстан'} | ForEach-Object {$_.values.B})
    foreach($sheet in $book.Worksheets){
        $used=$sheet.UsedRange
        $raw=$used.Value2
        $rows=@()
        for($ri=1;$ri -le $used.Rows.Count;$ri++){
            $cells=@()
            for($ci=1;$ci -le $used.Columns.Count;$ci++){
                $value=$raw.GetValue($ri,$ci)
                if($null -ne $value){$cells+=@{column=($used.Column+$ci-1);value=$value}}
            }
            if($cells.Count -gt 0){$rows+=@{row=($used.Row+$ri-1);cells=$cells}}
        }
        $matchedStyles=@()
        foreach($rowItem in $rows){
            $markCell=@($rowItem.cells | Where-Object {$_.column -eq 2})
            if($markCell.Count -eq 1 -and $markCell[0].value -in $requiredMarks){
                $mark=$sheet.Cells.Item($rowItem.row,2);$qty=$sheet.Cells.Item($rowItem.row,4)
                $matchedStyles+=@{row=$rowItem.row;mark=$mark.Value2;quantity=$qty.Value2;row_hidden=$sheet.Rows.Item($rowItem.row).Hidden;mark_strikethrough=$mark.Font.Strikethrough;qty_strikethrough=$qty.Font.Strikethrough;mark_fill=$mark.Interior.Color;qty_fill=$qty.Interior.Color;mark_font_color=$mark.Font.Color;qty_font_color=$qty.Font.Color;quantity_formula=$qty.Formula}
            }
        }
        $hidden=@()
        for($ri=1;$ri -le $used.Rows.Count;$ri++){if($sheet.Rows.Item($ri).Hidden){$hidden+=$ri}}
        $result.sheets+=@{name=$sheet.Name;used=$used.Address($false,$false);rows=$rows;matched_styles=$matchedStyles;hidden_rows=$hidden;autofilter_mode=$sheet.AutoFilterMode;filter_mode=$sheet.FilterMode}
    }
    $result.status='PASS_READ_ONLY_NATIVE_OPEN'
} catch {$result.status='STOP_NATIVE_VOM_ERROR';$result.error=$_.Exception.Message} finally {
    if($null -ne $book){$book.Close($false)}
    if($null -ne $excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    $result.after_sha=(Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $result.hash_unchanged=$result.before_sha -eq $result.after_sha
    $result | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'native_vom.json') -Encoding utf8
    [pscustomobject]$result | Select-Object status,error,readonly,format,hash_unchanged | ConvertTo-Json
}
