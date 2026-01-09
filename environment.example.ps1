$exportLocation = "X:\itp\xpf\xpf"
$tmpDir="X:\itp\tmp"
$hudubaseurl = "yourcompany.huducloud.com"
$huduapikey="hudukeyvaluehere"
$ITPexports = "x:\yourITPortalDocsLocation"

$internalCompanyName = "OurCo Internal"
$ITPortalBaseUrl = "YourCompany"
$IPortalShortname = "YourCo"
$ITPhostname = "$ITPortalBaseUrl.itportal.com"
$ITPInternalCompanyName = "Your Company Inc."
$ITPuserId = 567

#web session info - replace with your own
$ASPSession = @{
  "ASPSESSIONIDAWSBRRSC" = "SECRETASPSESSIONSID" # replace values with your ASP .NET session info
  "ASPSESSIONIDAUQDRSQD" = "SECRETASPSESSIONSID2"
  "ASPSESSIONIDAWRBSTRC" = "SECRETASPSESSIONSID3"
}
$ITPortalJWTTOKEN = "eyJh..."

# when we generate the cookie json for the session, we'll give it a future expiry
$expiryTimestamp = [math]::Round((([DateTimeOffset]::UtcNow.AddHours(6).UtcDateTime - [datetime]'1970-01-01').TotalSeconds), 6)
$CookieJSON =  @(
  @{
    "domain"= "$ITPortalBaseUrl.itportal.com"
    "expirationDate"= $expiryTimestamp
    "hostOnly"= $true
    "httpOnly"= $true
    "name"= "CompanyName"
    "path"= "/"
    "sameSite"= "lax"
    "secure"= $true
    "session"= $false
    "storeId"= "0"
    "value"= "$([System.Uri]::EscapeDataString($ITPInternalCompanyName))"
    "origin"= "https://$ITPortalBaseUrl.itportal.com"
  },
  @{
    "domain"= "$ITPortalBaseUrl.itportal.com"
    "expirationDate"= $expiryTimestamp
    "hostOnly"= $true
    "httpOnly"= $false
    "name"= "login_company"
    "path"= "/"
    "sameSite"= "strict"
    "secure"= $true
    "session"= $false
    "storeId"= "0"
    "value"= "$([System.Uri]::EscapeDataString($ITPInternalCompanyName))"
    "origin"= "https://$ITPortalBaseUrl.itportal.com"
  },
  @{
    "domain"= "$ITPortalBaseUrl.itportal.com"
    "expirationDate"= $expiryTimestamp
    "hostOnly"= $true
    "httpOnly"= $true
    "name"= "LoginCompany"
    "path"= "/"
    "sameSite"= "unspecified"
    "secure"= $false
    "session"= $false
    "storeId"= "0"
    "value"= "$($[System.Uri]::EscapeDataString($ITPInternalCompanyName -replace ' ','+'))"
    "origin"= "https://$ITPortalBaseUrl.itportal.com"
  },
  @{
    "domain"= "$ITPortalBaseUrl.itportal.com"
    "expirationDate"= $expiryTimestamp
    "hostOnly"= $true
    "httpOnly"= $true
    "name"= "HideMenus"
    "path"= "/"
    "sameSite"= "unspecified"
    "secure"= $false
    "session"= $false
    "storeId"= "0"
    "value"= "0%2C0%2C0%2C0%2C0%2C0%2C0%2C1%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C1%2C1%2C1%2C1%2C1%2C1%2C1%2C1%2C1%2C1"
    "origin"= "https://$ITPortalBaseUrl.itportal.com"
  },
  @{
    "domain"= "$ITPortalBaseUrl.itportal.com"
    "hostOnly"= $true
    "httpOnly"= $true
    "name"= "ASPSESSIONIDAWSBRRSC"
    "path"= "/"
    "sameSite"= "unspecified"
    "secure"= $true
    "session"= $true
    "storeId"= "0"
    "value"= "$($ASPSession.ASPSESSIONIDAWSBRRSC)"
    "origin"= "https://$ITPortalBaseUrl.itportal.com"
  },
  @{
    "domain"= "$ITPortalBaseUrl.itportal.com"
    "expirationDate"= $expiryTimestamp
    "hostOnly"= $true
    "httpOnly"= $true
    "name"= "Timeout"
    "path"= "/"
    "sameSite"= "unspecified"
    "secure"= $false
    "session"= $false
    "storeId"= "0"
    "value"= "15"
    "origin"= "https://$ITPortalBaseUrl.itportal.com"
  },
  @{
    "domain"= "$ITPortalBaseUrl.itportal.com"
    "expirationDate"= $expiryTimestamp
    "httpOnly"= $true
    "name"= "PubliccID"
    "path"= "/"
    "sameSite"= "unspecified"
    "secure"= $false
    "session"= $false
    "storeId"= "0"
    "value"= "1805"
    "origin"= "https://$ITPortalBaseUrl.itportal.com"
  },
  @{
    "domain"= "$ITPortalBaseUrl.itportal.com"
    "expirationDate"= $expiryTimestamp
    "hostOnly"= $true
    "httpOnly"= $true
    "name"= "PublicDomainName"
    "path"= "/"
    "sameSite"= "unspecified"
    "secure"= $false
    "session"= $false
    "storeId"= "0"
    "value"= "%2Eie"
    "origin"= "https://$ITPortalBaseUrl.itportal.com"
  },
  @{
    "domain"= "$ITPortalBaseUrl.itportal.com"
    "hostOnly"= $true
    "httpOnly"= $true
    "name"= "SIPortal"
    "path"= "/"
    "sameSite"= "unspecified"
    "secure"= $false
    "session"= $true
    "storeId"= "0"
    "value"= "AgentAdmin=0&Links=%5B%5D&bID=895678"
    "origin"= "https://$ITPortalBaseUrl.itportal.com"
  },
  @{
    "domain"= "$ITPortalBaseUrl.itportal.com"
    "hostOnly"= $true
    "httpOnly"= $true
    "name"= "ASPSESSIONIDAUQDRSQD"
    "path"= "/"
    "sameSite"= "unspecified"
    "secure"= $true
    "session"= $true
    "storeId"= "0"
    "value"= "$($ASPSession.ASPSESSIONIDAUQDRSQD)"
    "origin"= "https://$ITPortalBaseUrl.itportal.com"
  },
  @{
    "domain"= "$ITPortalBaseUrl.itportal.com"
    "expirationDate"= $expiryTimestamp
    "hostOnly"= $true
    "httpOnly"= $true
    "name"= "jwt_token"
    "path"= "/"
    "sameSite"= "lax"
    "secure"= $true
    "session"= $false
    "storeId"= "0"
    "value"= "$($ITPortalJWTTOKEN)"
    "origin"= "https://$ITPortalBaseUrl.itportal.com"
  },
  @{
    "domain"= "$ITPortalBaseUrl.itportal.com"
    "hostOnly"= $true
    "httpOnly"= $true
    "name"= "ASPSESSIONIDAWRBSTRC"
    "path"= "/"
    "sameSite"= "unspecified"
    "secure"= $true
    "session"= $true
    "storeId"= "0"
    "value"= "$($ASPSession.ASPSESSIONIDAWRBSTRC)"
    "origin"= "https://$ITPortalBaseUrl.itportal.com"
  })

. .\Migrate-ITPortal-Hudu.ps1
