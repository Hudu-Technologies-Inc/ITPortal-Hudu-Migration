foreach ($doc in $itportaldata.Documents.CsvData){
  # new client for XSSRF protection
  $expiryTimestamp = [math]::Round((([DateTimeOffset]::UtcNow.AddHours(6).UtcDateTime - [datetime]'1970-01-01').TotalSeconds), 6)
  $client = $client = New-ITPHttpClientFromBrowserDump -CookieJson $CookieJSON -basehost $ITPhostname -userId $ITPuserId -PortalOriginUrl "https://$ITPortalBaseUrl.itportal.com/v4/app/documents/$($doc.id)?ClientID=0"
    try {
        $docname = $doc.filename ?? "document_$($doc.id).bin"
        Get-ItPortalDocument -Client $client -CSVDoc $doc -OutputPath "$ITPexports\Documents" -fileName $docname
    } catch {
        Write-Error "Failed to get document: $_"
    }
    start-sleep 3
}

foreach ($doc in $itportaldata.KBs.CsvData){
  # new client for XSSRF protection
  $expiryTimestamp = [math]::Round((([DateTimeOffset]::UtcNow.AddHours(6).UtcDateTime - [datetime]'1970-01-01').TotalSeconds), 6)
  $client = $client = New-ITPHttpClientFromBrowserDump -CookieJson $CookieJSON -basehost $ITPhostname -userId $ITPuserId -PortalOriginUrl "https://$ITPortalBaseUrl.itportal.com/v4/app/documents/$($doc.id)?ClientID=0"
    try {
        $docname = $doc.filename ?? "document_$($doc.id).bin"
        Get-ItPortalDocument -Client $client -CSVDoc $doc -OutputPath "$ITPexports\Documents" -fileName $docname
    } catch {
        Write-Error "Failed to get document: $_"
    }
    start-sleep 3
}