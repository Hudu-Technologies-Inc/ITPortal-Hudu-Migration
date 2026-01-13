# ITPortal-Hudu-Migration
Migrate from ITPortal to Hudu with ease

## Requirements
- Hudu Instance of 2.39.6 or newer
- Hudu API Key (full permissions scope)
- ITPortal Export (csv)
- Sufficient Free Disk Space to Download/Convert Documents and Images
- Powershell 7.5.1 or later on Windows PC
- Latest Libreoffice*
- Cookie Editor Web Extension (chrome or firefox)

*installer will be launched if you don't have this yet

## getting started

fill out your copy of environment,example your values for these secrets/variables

```
$exportLocation = "X:\yourITPexport"            # this is where you've unzipped your itportal export
$tmpDir="X:\itp\tmp"                            # designated temporary base directory for converting potentially sensitive files
$hudubaseurl = "yourcompany.huducloud.com"      # your hudu instance url
$huduapikey="hudukeyvaluehere"                  # your hudu api key

$ITPDownloads = "x:\yourITPortalDocsLocation"   # this is the path that we'll use for downloading documents and kb images

$internalCompanyName = "OurCo Internal"         # this is the name of your internal company as seen in Hudu

$ITPortalSubdomain = "YourCompany"              # this is the subdomain of your itportal instance
$ITPhostname="$ITPortalSubdomain.itportal.com"  # this is the full hostname of your itportal instance
```

