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
copy-item environment.example.ps1 my-environment.ps1
notepad my-environment.ps1
```

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

And then you can kick it off by opening pwsh7 session as administrator, and dot-sourcing your environment file

```
. .\my-environment.ps1
```

## Refreshing Cookie Jar

This part is easier than it seems
First, you'll need to make sure you have this browser extension

http://chromewebstore.google.com/detail/cookiemanager-cookie-edit/hdhngoamekjhmnpenphenpaiindoinpo?hl=en

Then, log into itportal and move to documents page
Next, you'll click your cookie manager / editor icon and 'select-all'
then, hit 'export'.

when prompted for a save location, you'll save or overwrite the file located in this project folder, named 'cookiejar.json'

it's reccomended to do refresh your cookie jar by repeating the above process when prompted for a second time (as not to miss any images in documents)

