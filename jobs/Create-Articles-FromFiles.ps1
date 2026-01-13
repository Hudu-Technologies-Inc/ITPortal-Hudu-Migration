$DocConversionTempDir = $tmpDir ?? "c:\docs-tmp"
$sofficePath = Get-LibreMSI -TmpFolder $DocConversionTempDir
$ArticleMatches = $articleMatches ?? @{}
$internalCompany = $internalCompany ?? ($huducompanies | Where-Object { $_.name -ieq $internalCompanyName } | Select-Object -First 1)
$internalCompany = $internalCompany.company ?? $internalCompany

foreach ($doc in $(get-childitem -path "$ITPDownloads\Documents" -file -recurse)) {
    $uuid = [guid]::NewGuid().ToString()
    $record = $null; $company = $null;
    $dest = Join-Path $DocConversionTempDir ("doc-" + $uuid)
    Get-EnsuredPath -Path $dest | Out-Null

    $copied = Join-Path $dest $doc.Name
    
    Copy-Item -LiteralPath $doc.FullName -Destination $copied -Force
    $articleParams = @{
        ResourceLocation = (Get-Item -LiteralPath $copied)
        IncludeOriginals = $true
        updateOnMatch = $true
    }


    $record = $itportaldata.Documents.CsvData | Where-Object {$_.filename -ieq "$($doc.name)"} | Select-Object -First 1
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



