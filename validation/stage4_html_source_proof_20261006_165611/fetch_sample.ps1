# SOURCE PROOF #2 initial gate only. Not a mass downloader.
$ErrorActionPreference = 'Stop'
$htmlProofRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
$htmlProofCache = Join-Path $PSScriptRoot 'cache'
New-Item -ItemType Directory -Path $htmlProofCache -Force | Out-Null
$htmlProofBaseline = Get-Content -LiteralPath (Join-Path $htmlProofRoot 'validation\stage4_20261006_162729\accepted_hashes_after.json') -Raw | ConvertFrom-Json
foreach($htmlProofEntry in $htmlProofBaseline.PSObject.Properties) {
    if((Get-FileHash -LiteralPath $htmlProofEntry.Name -Algorithm SHA256).Hash.ToLowerInvariant() -ne $htmlProofEntry.Value) {
        throw "STOP: protected checkpoint changed: $($htmlProofEntry.Name)"
    }
}
$htmlProofBaseline | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'protected_hashes_before.json') -Encoding utf8
$htmlProofQueue = @(
    [pscustomobject]@{kind='FAMILY_INDEX';family='balka';profile=$null;standard=$null;name='index_balka.html';url='https://23met.ru/spravka/balka'},
    [pscustomobject]@{kind='PROFILE';family='balka';profile='35Б1';standard='ГОСТ Р 57837-2017';name='35B1_gost_r_57837_2017.html';url=('https://23met.ru/spravka/balka/'+[uri]::EscapeDataString('35Б1')+'?g='+[uri]::EscapeDataString('ГОСТ Р 57837-2017'))},
    [pscustomobject]@{kind='PROFILE';family='balka';profile='35Б1';standard='ГОСТ 26020-83';name='35B1_gost_26020_83.html';url=('https://23met.ru/spravka/balka/'+[uri]::EscapeDataString('35Б1')+'?g='+[uri]::EscapeDataString('ГОСТ 26020-83'))}
)
$htmlProofProgress = [System.Collections.Generic.List[object]]::new()
$htmlProofStopped = $false
foreach($htmlProofItem in $htmlProofQueue) {
    if($htmlProofProgress.Count -gt 0) { Start-Sleep -Seconds 10 }
    $htmlProofPath = Join-Path $htmlProofCache $htmlProofItem.name
    if(Test-Path -LiteralPath $htmlProofPath) { throw 'Sample cache exists; offline inspection required, no repeated fetch.' }
    $htmlProofStarted = [DateTimeOffset]::UtcNow.ToString('o')
    try {
        $htmlProofResponse = Invoke-WebRequest -Uri $htmlProofItem.url -UseBasicParsing -TimeoutSec 30 -OutFile $htmlProofPath -PassThru
        if([int]$htmlProofResponse.StatusCode -ne 200) { throw 'Non-200 reference response' }
        $htmlProofBytes = [IO.File]::ReadAllBytes($htmlProofPath)
        if($htmlProofBytes.Length -gt 2MB) { throw 'Oversized reference response: stop rather than cache extra content.' }
        $htmlProofBody = [Text.Encoding]::UTF8.GetString($htmlProofBytes)
        if($htmlProofBody -notmatch '<html\b') { throw 'Response is not a normal HTML page.' }
        $htmlProofProgress.Add([pscustomobject]@{
            kind=$htmlProofItem.kind;family=$htmlProofItem.family;requested_profile=$htmlProofItem.profile;
            requested_standard=$htmlProofItem.standard;url=$htmlProofItem.url;timestamp_utc=$htmlProofStarted;
            http_status=[int]$htmlProofResponse.StatusCode;content_type=($htmlProofResponse.Headers['Content-Type'] -join '; ');
            page_title=([regex]::Match($htmlProofBody,'<title[^>]*>(.*?)</title>','Singleline').Groups[1].Value);
            cache_path=$htmlProofPath;sha256=(Get-FileHash -LiteralPath $htmlProofPath -Algorithm SHA256).Hash.ToLowerInvariant();
            body_bytes=$htmlProofBytes.Length;result='CACHED_NOT_YET_AUDITED'})
        Write-Output "$($htmlProofItem.name): HTTP 200, cached"
    } catch {
        $htmlProofStatus = if($_.Exception.Response) {[int]$_.Exception.Response.StatusCode} else {$null}
        $htmlProofProgress.Add([pscustomobject]@{kind=$htmlProofItem.kind;family=$htmlProofItem.family;
            requested_profile=$htmlProofItem.profile;requested_standard=$htmlProofItem.standard;url=$htmlProofItem.url;
            timestamp_utc=$htmlProofStarted;http_status=$htmlProofStatus;error=$_.Exception.Message;result='STOP_FETCH'})
        $htmlProofStopped = $true
        Write-Output "STOP: HTTP $htmlProofStatus $($_.Exception.Message)"
    }
    [pscustomobject]@{scope='HTML_INITIAL_CRITICAL_GATE_ONLY';profiles_limit=20;delay_seconds=10;
        records=@($htmlProofProgress);stopped=$htmlProofStopped;further_requests='NONE_AFTER_INITIAL_GATE_PENDING_HTML_AUDIT'} |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'fetch_results.json') -Encoding utf8
    if($htmlProofStopped) { break }
}
if($htmlProofStopped) { exit 1 }
