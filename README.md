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
First, you'll need to make sure you have this browser extension. you can remove it afterwards, of course

http://chromewebstore.google.com/detail/cookiemanager-cookie-edit/hdhngoamekjhmnpenphenpaiindoinpo?hl=en
Then, log into itportal and move to documents page

<img width="1596" height="1076" alt="image" src="https://github.com/user-attachments/assets/3151e3aa-adec-4df8-b115-54a4094a0f2d" />

Next, you'll click the <span style="color:#f1c40f; font-weight:bold;">🟨 Cookie Manager icon (#1 in figure)</span>  
Then click <span style="color:#2ecc71; font-weight:bold;">🟩 Select All (#2 in figure)</span>  
Then click <span style="color:#9b59b6; font-weight:bold;">🟪 Export (#3 in figure)</span>

when prompted for a save location, you'll save or overwrite the file located in this project folder, named 'cookiejar.json'
If your browser doesnt ask you where to save this file, it will be in your downloads folder. You can hit CTL+J to get to recent downloads quickly. From there, you'll just need to copy the item to project folder and rename it to 'cookiejar.json', right next to your environment file.

it's reccomended to do refresh your cookie jar by repeating the above process when prompted for a second time (as not to miss any images in documents)

