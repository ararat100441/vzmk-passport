$ErrorActionPreference='Stop'
$word=$null
$doc=$null
$inputs=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'word_inputs.json') -Raw | ConvertFrom-Json
$results=@()
try {
    $word=New-Object -ComObject Word.Application
    $word.Visible=$false
    $word.DisplayAlerts=0
    $word.AutomationSecurity=3
    $word.Options.UpdateLinksAtOpen=$false
    foreach($inputItem in $inputs){
        $entry=[ordered]@{path=$inputItem.path;before_sha=(Get-FileHash -LiteralPath $inputItem.path -Algorithm SHA256).Hash.ToLowerInvariant();macro_security=3;open_and_repair=$false;status='STARTED'}
        try {
            $doc=$word.Documents.Open($inputItem.path,$false,$true,$false)
            if(-not $doc.ReadOnly){throw 'STOP: Word did not open read-only'}
            $doc.Repaginate()
            $entry.readonly=$doc.ReadOnly
            $entry.version=$word.Version
            $entry.pages=$doc.ComputeStatistics(2)
            $entry.sections=$doc.Sections.Count
            $entry.paragraphs=$doc.Paragraphs.Count
            $entry.tables=@()
            foreach($table in $doc.Tables){
                $cols=$null
                try {$cols=$table.Columns.Count} catch {$cols='MIXED_WIDTHS'}
                $entry.tables+=@{rows=$table.Rows.Count;columns=$cols;cells=$table.Range.Cells.Count;allow_auto_fit=$table.AllowAutoFit}
            }
            $entry.shapes=$doc.Shapes.Count
            $entry.inline_shapes=$doc.InlineShapes.Count
            $entry.fields=$doc.Fields.Count
            $entry.content_controls=$doc.ContentControls.Count
            $entry.protection_type=$doc.ProtectionType
            $entry.revisions=$doc.Revisions.Count
            $entry.track_revisions=$doc.TrackRevisions
            $entry.status='PASS_READ_ONLY_NATIVE_OPEN'
            # Visual evidence is produced separately by native_render.ps1.

        } catch {
            $entry.status='STOP_NATIVE_WORD_ERROR'
            $entry.error=$_.Exception.Message
        } finally {
            if($null -ne $doc){$doc.Close(0);$doc=$null}
            $entry.after_sha=(Get-FileHash -LiteralPath $inputItem.path -Algorithm SHA256).Hash.ToLowerInvariant()
            $entry.hash_unchanged=$entry.before_sha -eq $entry.after_sha
            $results += [pscustomobject]$entry
        }
        if($entry.status -ne 'PASS_READ_ONLY_NATIVE_OPEN'){break}
    }
} finally {
    if($null -ne $doc){$doc.Close(0)}
    if($null -ne $word){$word.Quit(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)}
    $results | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'native_word.json') -Encoding utf8
    $results | Select-Object path,status,error,pages,hash_unchanged,visual_evidence,visual_error | ConvertTo-Json
}
