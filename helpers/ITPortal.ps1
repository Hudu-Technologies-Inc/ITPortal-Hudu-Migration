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
  Get-EnsuredPath -path $DocumentDir
  $OutFile = Join-Path -Path $DocumentDir -ChildPath $FileName

  $downloadUrl = "https://$ITPortalBaseUrl.itportal.com/portal3/ajax-updates/?rID=DownloadDoc&DocumentID=$DocumentId&ClientID=$ClientId"
  Invoke-WebRequest -Uri $downloadUrl -Headers @{
    Cookie = $ITPortalCookie
    Accept = '*/*'
  } -OutFile $OutFile -MaximumRedirection 9 | Out-Null
  return (Resolve-Path $OutFile).Path
}
