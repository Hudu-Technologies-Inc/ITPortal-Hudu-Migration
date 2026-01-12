$DocConversionTempDir = $tmpDir ?? "c:\docs-tmp"
$ArticleMatches = $articleMatches ?? @{}
$internalCompany = $internalCompany ?? ($huducompanies | Where-Object { $_.name -ieq $internalCompanyName } | Select-Object -First 1)
$internalCompany = $internalCompany.company ?? $internalCompany

# step 2 - from knowledge base entries in CDSV, create hudu articles
foreach ($doc in $itportaldata.KBs.CsvData) {
    $MatchedKB = $null; $company = $null; $article = $null;
    $company = $(Get-HuduCompanies -Name $doc.company); $company = $company.company ?? $company;
    if ($null -eq $company -or $company.id -lt 1) {$company= $internalCompany}

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
    write-host " Created/Updated KB Article from CSV kb contents. $($article.name)" -ForegroundColor Green
    if ($null -ne $article) {
        $ArticleMatches["KBID_$($doc.KBID)"] = $article
        $client = New-ITPHttpClientFromBrowserDump -CookieJson $CookieJSON -ITPhostname $ITPhostname -userId $ITPuserId -PortalOriginUrl "https://$ITPortalBaseUrl.itportal.com/v4/app/kb/$($doc.kbid)?ClientID=0"
        $contents = $article.content
        $contents = Rewrite-ItPortalDownloadNoteFileLinks -Html $contents -BaseUrl "https://$ITPhostname" -Client $client -TempDir 'c:\docs-tmp' -Cache $cache -UploadableId 150
        $article = set-huduarticle -id $article.id -content $contents
        $article = $article.article ?? $article
        $ArticleMatches["KBID_$($doc.KBID)"] = $article
    }


}

# step 3 - docs from CSV doc contents
foreach ($doc in $itportaldata.documents.CsvData | where-object {-not ([string]::IsNullOrWhiteSpace($_.doc))}) {
    $MatchedKB = $null; $company = $null; $article = $null;
    $company = $(Get-HuduCompanies -Name $doc.company); $company = $company.company ?? $company;
    if ($null -eq $company -or $company.id -lt 1) {$company= $internalCompany}
    $name = $(Limit-StringLength $($doc.docname ?? "Doc-$($doc.documentId): $($doc.description)"))

    if ($null -ne $company){
        $matchedKB = get-huduarticles -CompanyId $company.id -name $name | Select-Object -First 1
    }
    $ArticleRequest = @{
        Content = if ([string]::IsNullOrWhiteSpace($doc.doc)) { "Blank Doc Content" } else { $doc.doc }
        Name    = $name
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
    write-host " Created/Updated KB Article from CSV doc contents. $($article.name)" -ForegroundColor Green
    if ($article) { 
        $client = New-ITPHttpClientFromBrowserDump -CookieJson $CookieJSON -ITPhostname $ITPhostname -userId $ITPuserId -PortalOriginUrl "https://$ITPortalBaseUrl.itportal.com/v4/app/documents/$($doc.documentid)?ClientID=0"
        $contents = $article.content
        $contents = Rewrite-ItPortalDownloadNoteFileLinks -Html $contents -BaseUrl "https://$ITPhostname" -Client $client -TempDir 'c:\docs-tmp' -Cache $cache -UploadableId 150
        $article = set-huduarticle -id $article.id -content $contents
        $article = $article.article ?? $article
        $ArticleMatches["DOCCSV_$($doc.documentId)"] = $article 
    
    }
}


