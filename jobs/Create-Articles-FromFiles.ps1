$DocConversionTempDir = $tmpDir ?? "c:\docs-tmp"
$sofficePath = Get-LibreMSI -TmpFolder $DocConversionTempDir
$internalCompany = Get-OrSetInternalCompany -internalCompanyName $internalCompanyName
$ArticleMatches = $articleMatches ?? @{}

$filesList = @()
if ($true -ne $useLocalFilesystemFiles){
    $fileItems = @($(get-childitem -path "$ITPDownloads\Documents" -file -recurse -ErrorAction SilentlyContinue)) + @($(get-childitem -path "$ITPDownloads\KBs" -file -recurse -ErrorAction SilentlyContinue))
    foreach ($file in $fileItems) {
        $record = $itportaldata.Documents.CsvData | Where-Object {$_.filename -ieq "$($file.name)"} | Select-Object -First 1
        $itemType = 'Document'
        if (-not $record) {
            $record = $itportaldata.KBs.CsvData | Where-Object {$_.filename -ieq "$($file.name)"} | Select-Object -First 1
            $itemType = 'KB'
        }
        $filesList += [pscustomobject]@{ File = $file; Record = $record; ItemType = $itemType }
    }
} else {
    if (-not (Test-Path -LiteralPath $ITPDownloads)) {
        throw "Local ITPortal filesystem path does not exist: $ITPDownloads"
    }

    $seenFiles = @{}
    foreach ($record in @($itportaldata.Documents.CsvData)) {
        foreach ($file in @(Get-ItPortalLocalRecordFiles -CsvRow $record -LocalRoot $ITPDownloads -ItemType Document)) {
            if ($seenFiles.ContainsKey($file.FullName)) { continue }
            $seenFiles[$file.FullName] = $true
            $filesList += [pscustomobject]@{ File = $file; Record = $record; ItemType = 'Document' }
        }
    }
    foreach ($record in @($itportaldata.KBs.CsvData)) {
        foreach ($file in @(Get-ItPortalLocalRecordFiles -CsvRow $record -LocalRoot $ITPDownloads -ItemType KB)) {
            if ($seenFiles.ContainsKey($file.FullName)) { continue }
            $seenFiles[$file.FullName] = $true
            $filesList += [pscustomobject]@{ File = $file; Record = $record; ItemType = 'KB' }
        }
    }
}

foreach ($item in $filesList) {
    $doc = $item.File
    $uuid = [guid]::NewGuid().ToString()
    $record = $item.Record; $company = $null;
    $dest = Join-Path $DocConversionTempDir ("doc-" + $uuid)
    Get-EnsuredPath -Path $dest | Out-Null

    $copied = Join-Path $dest $doc.Name
    
    Copy-Item -LiteralPath $doc.FullName -Destination $copied -Force
    $articleParams = @{
        ResourceLocation = (Get-Item -LiteralPath $copied)
        IncludeOriginals = $true
        updateOnMatch = $true
    }

    if ($null -ne $record){
        $company = $(Get-HuduCompanies -Name $record.company | select-object -first 1); $company = $company.company ?? $company;
        if ($null -ne $company -and $company.id -ge 1){
            $ArticleParams.CompanyName = $Company.Name
            write-host " Associating document $($doc.name) with company $($company.name)" -ForegroundColor Cyan
        } else {
            write-host " No matching company found for document $($doc.name). adding as global kb" -ForegroundColor Yellow
        }
    }

    $article = New-HuduArticleFromLocalResource @articleParams
    if ($null -ne $article.result){
        $ArticleMatches["DocumentFile_$($doc.name)"] = $article.result
        write-host " Created article from file. $($article.result.name)" -ForegroundColor Green
    }
}        

$ArticleMatches | convertto-json -depth 99 | set-content -path $(join-path $debugDir -childpath "Articles-FromFiles.json") -force
