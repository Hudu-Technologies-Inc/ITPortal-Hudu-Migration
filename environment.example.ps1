$exportLocation = "X:\yourITPexport"            # this is where you've unzipped your itportal export
$tmpDir="X:\itp\tmp"                            # designated temporary base directory for converting potentially sensitive files
$hudubaseurl = "yourcompany.huducloud.com"      # your hudu instance url
$huduapikey="hudukeyvaluehere"                  # your hudu api key

$ITPDownloads = "x:\yourITPortalDocsLocation"   # this is the path that we'll use for downloading documents and kb images

$internalCompanyName = "OurCo Internal"         # this is the name of your internal company as seen in Hudu

$ITPortalSubdomain = "YourCompany"              # this is the subdomain of your itportal instance
$ITPhostname="$ITPortalSubdomain.itportal.com"  # this is the full hostname of your itportal instance
$useLocalFilesystemFiles = $false # only set to true if you are self-hosted with no web-access. If true you do NOT need to set cookiejar file.

# $CookieJson = $(get-content -Raw -Path ".\cookiejar.json" | ConvertFrom-Json -depth -99)


. .\Migrate-ITPortal-Hudu.ps1
