$ErrorActionPreference = 'Stop'
$taskRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
$taskCache = Join-Path $taskRoot 'DEV\Справочники\source_cache'
$taskResults = Join-Path $PSScriptRoot 'bulk_fetch_results.json'
$taskIndex = Get-Content -LiteralPath (Join-Path $taskCache 'normative_selected_links.html') -Raw
$taskStandards = @('ГОСТ Р 57837-2017','СТО АСЧМ 20-93','ГОСТ 26020-83','ГОСТ 35087-2024',
    'ГОСТ 8240-97','ГОСТ 8509-93','ГОСТ 8510-86','ГОСТ 8639-82','ГОСТ 8645-68',
    'ГОСТ 10704-91','ГОСТ 30245-2003','ГОСТ 2879-2006','ГОСТ 8239-89')
$taskRecords = [System.Collections.Generic.List[object]]::new()
foreach($taskStandard in $taskStandards) {
    if(-not $taskIndex.Contains($taskStandard)) { throw "Standard missing from inspected public index: $taskStandard" }
    Start-Sleep -Seconds 4
    $taskPageUrl = 'https://23met.ru/gosts/'+[uri]::EscapeDataString($taskStandard)
    try {
        $taskResponse = Invoke-WebRequest -Uri $taskPageUrl -UseBasicParsing -TimeoutSec 30
        $taskFrame = [regex]::Match($taskResponse.Content,'<iframe[^>]*src="(https://23met.ru/gost_files/[^"]+)"')
        if(-not $taskFrame.Success) { throw "No public normative PDF frame: $taskStandard" }
        $taskPdfUrl = $taskFrame.Groups[1].Value
        $taskPdfName = [uri]::UnescapeDataString(([uri]$taskPdfUrl).Segments[-1])
        if($taskPdfName -notmatch '^[a-z0-9-]+\.pdf$') { throw 'Unsafe cache name' }
        $taskPdfPath = Join-Path $taskCache $taskPdfName
        if(-not (Test-Path -LiteralPath $taskPdfPath)) {
            Start-Sleep -Seconds 4
            Invoke-WebRequest -Uri $taskPdfUrl -UseBasicParsing -TimeoutSec 45 -OutFile $taskPdfPath
        }
        $taskPdfBytes = [IO.File]::ReadAllBytes($taskPdfPath)
        if($taskPdfBytes.Length -gt 20MB -or [Text.Encoding]::ASCII.GetString($taskPdfBytes,0,5) -ne '%PDF-') { throw 'Invalid/oversized public PDF' }
        $taskRecords.Add([pscustomobject]@{standard=$taskStandard;page_url=$taskPageUrl;url=$taskPdfUrl;
            source_type='PDF';snapshot_path=$taskPdfPath;sha256=(Get-FileHash -LiteralPath $taskPdfPath -Algorithm SHA256).Hash.ToLowerInvariant();
            last_successful_fetch=[DateTimeOffset]::UtcNow.ToString('o');status='FETCHED_NOT_PARSED';bytes=$taskPdfBytes.Length})
        Write-Output "$taskStandard : FETCHED_NOT_PARSED ($($taskPdfBytes.Length) bytes)"
    } catch {
        $taskStatus = if($_.Exception.Response) {[int]$_.Exception.Response.StatusCode} else {$null}
        $taskRecords.Add([pscustomobject]@{standard=$taskStandard;page_url=$taskPageUrl;status='STOP_FETCH';http_status=$taskStatus;error=$_.Exception.Message})
        $taskRecords | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $taskResults -Encoding utf8
        Write-Output "$taskStandard : STOP_FETCH $taskStatus $($_.Exception.Message)"
        break
    }
    $taskRecords | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $taskResults -Encoding utf8
}
