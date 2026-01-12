$DocConversionTempDir = $tmpDir ?? "c:\docs-tmp"
$sofficePath = Get-LibreMSI -TmpFolder $DocConversionTempDir
$ArticleMatches = $articleMatches ?? @{}

# step one - from downloaded files, create hudu articles
foreach ($doc in $itportaldata.Documents.CsvData | where-object {-not ($_.FileDeleted -ieq "True")}){
    $docFile = $null; $article = $null; $company = $null;
    $docFile = $doc.filename

    $docPath = Get-ChildItem -Path "$ITPexports\Documents" -file | Where-Object { $_.Name -ieq $docFile} | Select-Object -First 1
    $docPath = $docPath ?? $(Get-ChildItem -Path "$ITPexports\Documents" -file | Where-Object { $_.Name -ilike "*$($doc.DocumentID)"} | Select-Object -First 1)
    $docPath = $docPath ?? $(Get-ChildItem -Path "$ITPexports\Documents" -file | Where-Object { $_.Name -ilike "*$($doc.Docname)*"} | Select-Object -First 1)
    $company = $(Get-HuduCompanies Name $doc.company); $company = $company.company ?? $company;

    $companyArticles = @()
    if ($null -ne $company){
        $companyArticles = get-huduarticles -CompanyId $company.id
    }

    $uuid = [guid]::NewGuid().ToString()
    $dest = Join-Path $DocConversionTempDir ("doc-$($doc.DocumentID)-" + $uuid)
    Get-EnsuredPath -Path $dest | Out-Null
    if ($docPath -and (Test-Path -LiteralPath $docPath.FullName)){
        Copy-item -Path $docPath.FullName -Destination $dest -Force
    }
    $tempFile = Get-Childitem -Path $dest | Select-Object -First 1


    $articleParams = @{
        ResourceLocation = (Get-Item -LiteralPath $tempFile)
        IncludeOriginals = $true
        CompanyName = $company.name
        companydocs = $companyArticles
        updateOnMatch = $true
    }
    $article = New-HuduArticleFromLocalResource @articleParams
    if ($null -ne $article.result){
        $ArticleMatches["DocumentID_$($doc.DocumentID)"] = $article.result
    }

    write-host " Created article: $($article.result.name)" -ForegroundColor Green



}        

# step 2 - from knowledge base entries, create hudu articles
foreach ($doc in $itportaldata.KBs.CsvData) {
    $MatchedKB = $null; $company = $null; $article = $null;
    $company = $(Get-HuduCompanies Name $doc.company); $company = $company.company ?? $company;
    $name = $(Limit-StringLength $($doc.kbname ?? "KB-$($doc.kbid): $($doc.description)"))

    if ($null -ne $company){
        $matchedKB = get-huduarticles -CompanyId $company.id -name $name | Select-Object -First 1
    }
    $ArticleRequest = @{
        Content = $( if ([string]::isnullorempty($doc.kb)) { "Blank KB Content" } else { $doc.kb }  )
        Name = $name
    }

    if ($company){
        $ArticleRequest.CompanyId = $company.id
    }
    if ($matchedKB){
        $ArticleRequest.Id = $matchedKB.id
        $article = Set-HuduArticle @ArticleRequest
    } else {
        $article = New-HuduArticle @articleRequest
    }
    $article = $article.article ?? $article
    write-host " Created/Updated KB Article: $($article.name)" -ForegroundColor Green
    if ($null -ne $article.result){
        $ArticleMatches["KBID_$($doc.KBID)"] = $article
    }


}