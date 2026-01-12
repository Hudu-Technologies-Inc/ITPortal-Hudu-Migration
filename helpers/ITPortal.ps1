function Get-ItPortalDocumentWithCookie {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$ITPortalBaseUrl,
    [Parameter(Mandatory)][string]$ITPortalCookie,
    [Parameter(Mandatory)][int]$DocumentId,
    [Parameter(Mandatory)][int]$ClientId,
    [parameter(Mandatory)][string]$DocumentDir,
    [Parameter()][string]$FileName = ("Document-$DocumentId.bin")
  )

  $downloadUrl = "https://$ITPortalBaseUrl.itportal.com/portal3/ajax-updates/?rID=DownloadDoc&DocumentID=$DocumentId&ClientID=$ClientId"
  Invoke-WebRequest -Uri $downloadUrl -Headers @{
    Cookie = $ITPortalCookie
    Accept = '*/*'
  } -OutFile $OutFile -MaximumRedirection  | Out-Null
  return (Resolve-Path $OutFile).Path
}
function New-ITPHttpClientFromBrowserDump {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][pscustomobject[]]$CookieJson,
    [Parameter(Mandatory)][string]$ITPhostname,           # e.g. "redacted.itportal.com"
    [Parameter(Mandatory)][int]$UserId,                # e.g. 588
    [Parameter(Mandatory)][string]$PortalOriginUrl     # e.g. "https://redacted.itportal.com/v4/app/documents/733?ClientID=0"
  )

  $handler = [System.Net.Http.HttpClientHandler]::new()
  $handler.UseCookies = $true
  $handler.CookieContainer = [System.Net.CookieContainer]::new()
  $handler.AutomaticDecompression = `
    [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate

  $baseUri = [Uri]("https://$ITPhostname/")

  foreach ($c in $CookieJson) {
    if (-not $c.name -or -not $c.value) { continue }

    $cookie = [System.Net.Cookie]::new($c.name, $c.value)
    $cookie.Path = if ($c.path) { $c.path } else { '/' }

    # Respect hostOnly cookies
    if ($c.hostOnly -eq $true) {
      $cookie.Domain = $baseUri.Host
    } else {
      $cookie.Domain = ([string]$c.domain).TrimStart('.')
    }

    $handler.CookieContainer.Add($cookie)
  }

  $client = [System.Net.Http.HttpClient]::new($handler)
  $client.BaseAddress = $baseUri

  # Match the browser XHR headers
  $client.DefaultRequestHeaders.Clear()
  $client.DefaultRequestHeaders.TryAddWithoutValidation('Accept', 'application/json, text/plain, */*') | Out-Null
  $client.DefaultRequestHeaders.TryAddWithoutValidation('Accept-Language', 'en-US,en;q=0.5') | Out-Null
  $client.DefaultRequestHeaders.TryAddWithoutValidation('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:146.0) Gecko/20100101 Firefox/146.0') | Out-Null

  # Only advertise encodings we can auto-decompress
  $client.DefaultRequestHeaders.TryAddWithoutValidation('Accept-Encoding', 'gzip, deflate') | Out-Null

  $client.DefaultRequestHeaders.TryAddWithoutValidation('Referer', "https://$ITPhostname/") | Out-Null
  $client.DefaultRequestHeaders.TryAddWithoutValidation('UserID', "$UserId") | Out-Null
  $client.DefaultRequestHeaders.TryAddWithoutValidation('X-Portal-AppVersion', '/v4') | Out-Null
  $client.DefaultRequestHeaders.TryAddWithoutValidation('X-Portal-Origin', $PortalOriginUrl) | Out-Null

  return $client
}
function Get-ItPortalDocument {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][System.Net.Http.HttpClient]$Client,
    [Parameter(Mandatory)][pscustomobject]$CSVDoc,     # expects DocumentID, ClientID
    [Parameter(Mandatory)][string]$OutputPath,
    [Parameter()][string]$FileName
  )

  if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
  }

  $docId = [int]$CSVDoc.DocumentID
  $cid   = [int]$CSVDoc.ClientID

  $outName = if ($FileName) { $FileName } else { "Document-$docId.bin" }
  $outFile = Join-Path $OutputPath $outName
  $client.DefaultRequestHeaders.Remove('X-Portal-Origin') | Out-Null
  $client.DefaultRequestHeaders.TryAddWithoutValidation(
    'X-Portal-Origin',
    "https://redacted.itportal.com/v4/app/documents/$($CSVDoc.DocumentID)?ClientID=$($CSVDoc.ClientID)"
  ) | Out-Null
  $ub = [System.UriBuilder]::new($Client.BaseAddress.AbsoluteUri)
  $ub.Path = '/portal3/ajax-updates/'
  $ub.Query = "rID=DownloadDoc&DocumentID=$docId&ClientID=$cid"
  $url = $ub.Uri.AbsoluteUri

  $resp = $Client.GetAsync($url).GetAwaiter().GetResult()
  $ct = $resp.Content.Headers.ContentType.MediaType

  $bytes = $resp.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()

  if (-not $resp.IsSuccessStatusCode) {
    $head = [Text.Encoding]::UTF8.GetString($bytes, 0, [Math]::Min(500, $bytes.Length))
    throw "HTTP $([int]$resp.StatusCode) $($resp.ReasonPhrase). Content-Type=$ct. Body(head)=$head"
  }

  # If still HTML, dump a readable snippet (after gzip/deflate auto-decompress should already be applied)
  # if ($ct -match 'text/html') {
  #   $head = [Text.Encoding]::UTF8.GetString($bytes, 0, [Math]::Min(800, $bytes.Length))
  #   throw "Got HTML instead of file. Content-Type=$ct. Body(head)=$head"
  # }

  [IO.File]::WriteAllBytes($outFile, $bytes)
  try {
    $completedpath = (Resolve-Path $outFile).Path
    return $completedpath
  } catch {
    write-host $_
    return $outFile
  }
}

function Get-HuduUrlForItPortalUpload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RelativeUrl,
        [Parameter(Mandatory)][string]$BaseUrl,
        [Parameter(Mandatory)][System.Net.Http.HttpClient]$Client,
        [Parameter(Mandatory)][string]$TempDir,
        [Parameter()][hashtable]$Cache,

        [Parameter()][int]$UploadableId,
        [Parameter()][ValidateSet('Article','Asset','Company', IgnoreCase=$true)]
        [string]$UploadableType = 'Article'
    )

    if (-not $Cache) { $Cache = @{} }

    $rel = $RelativeUrl.Trim()
    if (-not $rel.StartsWith('/')) { $rel = "/$rel" }

    $uploadId = [regex]::Match($rel, '(?i)(?:\?|&)UploadID=(\d+)').Groups[1].Value
    if ([string]::IsNullOrWhiteSpace($uploadId)) { return $null }

    # cache key should include uploadable target if that changes destination behavior
    $cacheKey = if ($UploadableId) { "$uploadId|$UploadableType|$UploadableId" } else { "$uploadId" }
    if ($Cache.ContainsKey($cacheKey)) { return $Cache[$cacheKey] }

    $null = New-Item -ItemType Directory -Path $TempDir -Force -ErrorAction SilentlyContinue

    $uri  = [Uri]::new(($BaseUrl.TrimEnd('/') + $rel))
    $resp = $Client.GetAsync($uri).GetAwaiter().GetResult()
    $resp.EnsureSuccessStatusCode() | Out-Null

    $bytes = $resp.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()

    $ct = $resp.Content.Headers.ContentType?.MediaType
    $ext =
        if    ($ct -eq 'image/jpeg') { '.jpg' }
        elseif($ct -eq 'image/png')  { '.png' }
        elseif($ct -eq 'image/gif')  { '.gif' }
        elseif($ct -eq 'image/webp') { '.webp' }
        else { '.bin' }

    $filePath = Join-Path $TempDir ("itp-upload-$uploadId$ext")
    [IO.File]::WriteAllBytes($filePath, $bytes)

    # attach upload to article if requested
    $uploadParams = @{ FilePath = $filePath }
    if ($UploadableId) {
        $uploadParams.UploadableType = $UploadableType
        $uploadParams.UploadableId   = $UploadableId
    }

    $huduUpload = New-HuduUpload @uploadParams
    $huduUpload = $huduUpload.upload ?? $huduUpload
    $url = $huduUpload.url

    $Cache[$cacheKey] = $url
    return $url
}

function Rewrite-ItPortalDownloadNoteFileLinks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Html,
        [Parameter(Mandatory)][string]$BaseUrl,
        [Parameter(Mandatory)][System.Net.Http.HttpClient]$Client,
        [Parameter(Mandatory)][string]$TempDir,
        [Parameter()][hashtable]$Cache,

        [Parameter()][int]$UploadableId,
        [Parameter()][ValidateSet('Article','Asset','Company', IgnoreCase=$true)]
        [string]$UploadableType = 'Article'
    )

    if (-not $Cache) { $Cache = @{} }

    $rx = [regex]'(?i)/portal3/ajax-updates/\?rID=DownloadNoteFile(?:[&]|&amp;)[^"\s>]+'

    $rx.Replace($Html, {
        param($m)
        $relDecoded = $m.Value -replace '&amp;', '&'

        $newUrl = Get-HuduUrlForItPortalUpload `
            -RelativeUrl $relDecoded `
            -BaseUrl $BaseUrl `
            -Client $Client `
            -TempDir $TempDir `
            -Cache $Cache `
            -UploadableId $UploadableId `
            -UploadableType $UploadableType

        if ([string]::IsNullOrWhiteSpace($newUrl)) { $m.Value } else { $newUrl }
    })
}