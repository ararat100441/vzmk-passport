$ErrorActionPreference='Stop'
$auditRoot=$PSScriptRoot
$devPath=(Resolve-Path -LiteralPath (Join-Path $auditRoot '..\..\DEV\УЗК_DEV.xlsm')).Path
$excel=$null
$book=$null
$result=[ordered]@{input=$devPath;before_sha=(Get-FileHash -LiteralPath $devPath -Algorithm SHA256).Hash.ToLowerInvariant();macro_security=3;events=$false;update_links=0;corrupt_load=0;sheets=@();status='STARTED'}
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false
    $excel.DisplayAlerts=$false
    $excel.AutomationSecurity=3
    $excel.EnableEvents=$false
    $excel.AskToUpdateLinks=$false
    $result.version=$excel.Version
    $missing=[Type]::Missing
    $book=$excel.Workbooks.Open($devPath,0,$true)
    if(-not $book.ReadOnly){throw 'STOP: Excel did not open read-only'}
    $result.readonly=$book.ReadOnly
    $result.file_format=$book.FileFormat
    $result.has_vba_project=$book.HasVBProject
    $result.protect_structure=$book.ProtectStructure
    $result.protect_windows=$book.ProtectWindows
    $result.names=@($book.Names | ForEach-Object { @{name=$_.Name;refers_to=$_.RefersTo;visible=$_.Visible} })
    $result.external_links=@($book.LinkSources(1))
    $result.connections=@($book.Connections | ForEach-Object { @{name=$_.Name;type=$_.Type} })
    foreach($sheet in $book.Worksheets){
        $range=$sheet.UsedRange
        $values=@()
        for($row=1;$row -le $range.Rows.Count;$row++){
            $cells=@()
            for($col=1;$col -le $range.Columns.Count;$col++){
                $cell=$sheet.Cells.Item($row,$col)
                $cells+=@{address=$cell.Address($false,$false);value=$cell.Value2;text=$cell.Text;formula=$cell.Formula;has_formula=$cell.HasFormula;merged=$cell.MergeCells}
            }
            $values+=@{row=$row;cells=$cells}
        }
        $result.sheets+=@{name=$sheet.Name;visible=$sheet.Visible;used=$range.Address($false,$false);rows=$range.Rows.Count;columns=$range.Columns.Count;protect_contents=$sheet.ProtectContents;protect_drawing_objects=$sheet.ProtectDrawingObjects;protect_scenarios=$sheet.ProtectScenarios;auto_filter_mode=$sheet.AutoFilterMode;filter_mode=$sheet.FilterMode;listobjects=$sheet.ListObjects.Count;rows_data=$values}
    }
    try {
        $project=$book.VBProject
        $result.vba=@{accessible=$true;protection=$project.Protection;references=@($project.References | ForEach-Object {@{name=$_.Name;guid=$_.GUID;major=$_.Major;minor=$_.Minor;broken=$_.IsBroken}});components=@($project.VBComponents | ForEach-Object {@{name=$_.Name;type=$_.Type;lines=$_.CodeModule.CountOfLines}})}
        foreach($component in $project.VBComponents){
            $lineCount=$component.CodeModule.CountOfLines
            if($lineCount -gt 0){
                $sourceText=$component.CodeModule.Lines(1,$lineCount)
                $sourceText | Set-Content -LiteralPath (Join-Path $auditRoot ('VBA_'+$component.Name+'.txt')) -Encoding utf8
            }
        }
    } catch {$result.vba=@{accessible=$false;error=$_.Exception.Message;security_settings_changed=$false}}
    $result.status='PASS_READ_ONLY_NATIVE_OPEN'
} catch {
    $result.status='STOP_NATIVE_OPEN_ERROR'
    $result.error=$_.Exception.Message
} finally {
    if($null -ne $book){$book.Close($false)}
    if($null -ne $excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    $result.after_sha=(Get-FileHash -LiteralPath $devPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $result.hash_unchanged=$result.before_sha -eq $result.after_sha
    $result | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath (Join-Path $auditRoot 'native_excel.json') -Encoding utf8
    [pscustomobject]$result | Select-Object status,error,version,readonly,has_vba_project,hash_unchanged | ConvertTo-Json
}
