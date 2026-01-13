Read-host "This step requires a fresh set of cookies! Be sure to use 'cookie-editor' and export all cookies, writing or overwriting the file located at $project_workdir\cookiejar.json"
$successReadCookies = $false
while ($false -eq $successReadCookies){
    try {
        $CookieJson = $(get-content -Raw -Path ".\cookiejar.json" | ConvertFrom-Json -depth -99)
        $successReadCookies = $true
        break
    } catch {
        Write-Error "Failed to read cookiejar.json file. Please ensure it exists and is valid JSON. Details: $_"
    }
}


foreach ($doc in $itportaldata.Documents.CsvData){
  # new client for XSSRF protection
  $expiryTimestamp = [math]::Round((([DateTimeOffset]::UtcNow.AddHours(6).UtcDateTime - [datetime]'1970-01-01').TotalSeconds), 6)
  $client = New-ITPHttpClientFromBrowserDump -CookieJson $CookieJSON -ITPhostname $ITPhostname -userId $ITPuserId -PortalOriginUrl "https://$ITPortalSubdomain.itportal.com/v4/app/documents/$($doc.documentId)?ClientID=0"
    try {
        $docname = $doc.filename ?? "document_$($doc.documentId)"
        write-host " Fetching document: $($doc.documentId) as $docname" -ForegroundColor Cyan
        $docname = Get-ItPortalDocument -Client $client -CSVDoc $doc -OutputPath "$ITPDownloads\Documents" -fileName $docname
        if (-not $doc.fileName){
            Write-Host "attempting to find filetype for doc $($doc.documentId)" -ForegroundColor Yellow
            $fileType = Get-FileType -path $docname
            $newFileName = "$($docname).$($fileType.ToLower())"
            Rename-Item -Path $docname -NewName $newFileName -Force
            write-host " Renamed document $($doc.documentId) to $newFileName based on detected file type $fileType" -ForegroundColor Green
        }

    } catch {
        Write-Error "Failed to get document: $_"
    }
    start-sleep 3
}

