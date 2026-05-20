$DocConversionTempDir = $tmpDir ?? "c:\docs-tmp"


$ArticleMatches = $articleMatches ?? @{}
$internalCompany = $internalCompany ?? $(Get-OrSetInternalCompany -internalCompanyName $internalCompanyName)

if ($true -ne $useLocalFilesystemFiles){
    $cookieJson = Get-ProperCookieJson -project_workdir $project_workdir -neededFor "Fetching/Downloading Images from ITPortal Documents included in csv export"
} elseif (-not (Test-Path -LiteralPath $ITPDownloads)) {
    throw "Local ITPortal filesystem path does not exist: $ITPDownloads"
}
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
        write-host " Found existing article matching doc from CSV: $($matchedKB.name)" -ForegroundColor Yellow
        $ArticleRequest.Id = $matchedKB.id
        $article = Set-HuduArticle @ArticleRequest
    } else {
        $article = New-HuduArticle @articleRequest
    }
    $article = $article.article ?? $article
    write-host " Created/Updated KB Article from CSV kb contents. $($article.name)" -ForegroundColor Green
    if ($article) {
        $ArticleMatches["KBID_$($doc.KBID)"] = $article
        $contents = $ArticleRequest.Content
        if ($true -eq $useLocalFilesystemFiles) {
            $contents = Rewrite-ItPortalLocalNoteFileLinks -Html $contents -LocalRoot $ITPDownloads -Cache $cache -UploadableId $article.id
        } else {
            $client = New-ITPHttpClientFromBrowserDump -CookieJson $CookieJSON -ITPhostname $ITPhostname -userId $ITPuserId -PortalOriginUrl "https://$ITPortalSubdomain.itportal.com/v4/app/kb/$($doc.kbid)?ClientID=0"
            $contents = Rewrite-ItPortalDownloadNoteFileLinks -Html $contents -BaseUrl "https://$ITPhostname" -Client $client -TempDir $DocConversionTempDir -Cache $cache -UploadableId $article.id
        }
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
        write-host " Found existing article matching doc from CSV: $($matchedKB.name)" -ForegroundColor Yellow
        $ArticleRequest.Id = $matchedKB.id
        $article = Set-HuduArticle @ArticleRequest
    } else {
        $article = New-HuduArticle @articleRequest
    }
    $article = $article.article ?? $article
    write-host " Created/Updated KB Article from CSV doc contents. $($article.name)" -ForegroundColor Green
    if ($article) {
        $ArticleMatches["DOCCSV_$($doc.documentId)"] = $article
        $contents = $ArticleRequest.Content
        if ($true -eq $useLocalFilesystemFiles) {
            $contents = Rewrite-ItPortalLocalNoteFileLinks -Html $contents -LocalRoot $ITPDownloads -Cache $cache -UploadableId $article.id
        } else {
            $client = New-ITPHttpClientFromBrowserDump -CookieJson $CookieJSON -ITPhostname $ITPhostname -userId $ITPuserId -PortalOriginUrl "https://$ITPortalSubdomain.itportal.com/v4/app/documents/$($doc.documentid)?ClientID=0"
            $contents = Rewrite-ItPortalDownloadNoteFileLinks -Html $contents -BaseUrl "https://$ITPhostname" -Client $client -TempDir $DocConversionTempDir -Cache $cache -UploadableId $article.id
        }
        $article = set-huduarticle -id $article.id -content $contents
        $article = $article.article ?? $article
        $ArticleMatches["DOCCSV_$($doc.documentId)"] = $article 
    
    }
}


$ArticleMatches | convertto-json -depth 99 | set-content -path $(join-path $debugDir -childpath "Articles-FromRecords.json") -force
