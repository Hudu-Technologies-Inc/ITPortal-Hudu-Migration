# ITPortal → Hudu Migration
Seamlessly migrate data, documents, and knowledge from ITPortal into Hudu.

## Requirements

- Hudu v2.39.6+
- Hudu API Key
- ITPortal CSV Export
- PowerShell 7.5.1+ on Windows Machine
- LibreOffice
- Cookie Manager Browser Extension

* **installer for newest release of LibreOffice will be launched if you don't have this yet**

## Getting Started
Copy the **environment template**, ```environment.example.ps1``` to a ***new file*** as shown in the example, **below**:

```
Copy-Item environment.example.ps1 my-environment.ps1
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

And then you can kick it off by opening `pwsh7` session as `administrator`, and `dot-sourcing` your `environment file`, as in the below example-

```
. .\my-environment.ps1
```

## What Gets Moved Over?

- `Accounts`
- `ConfigItems`
- `Agreements`
- `Devices`
- `Contacts`
- `Sites` come over as assets *(used as `locations`)*
- `Companies`
- `Kbs/KnowledgeBase Articles` from **CSV export** as `Articles`
- `Documents` from *CSV export* as `Articles`
- `Documents` as *Files downloaded from ITportal* are *converted* to *editable* `Articles`
- `Device-Passwords`, `Account-Passwords`, `LocalPasswords`, and `FieldPasswords` are *all imported* as `AssetPasswords` and *related to objects in Hudu*
- `Relationships` are *carried over*

## What Doesnt get Moved Over?
While all these objects get moved over, including the source images, the images set for companies will need to be re-applied from the images uploaded into hudu (for now)

## information on Jobs/Tasks

### 1. **Read-Data** Task
gets initial information about your Hudu instance, whether or not your internal company exists yet, and what kinds of assets and companies you already have there. Makes sure your internal company exists and reads-in CSV data from your export location.
<img width="596" height="538" alt="image" src="https://github.com/user-attachments/assets/ae6a5a4d-c344-46c7-a67d-92d0360de97d" />

---

### 2. **Assets-And-Layouts** Task
`Accounts`, `ConfigItems`, `Agreements`, `Devices`, `Contacts`, and `Sites` get moved over during the `Assets-And-Layouts` Job as Assets in Hudu. `Companies` are created during this job as well. 
When this job finishes, all created assets can be found in a folder in the project directory, named `Debug`. The assets dump file is named `CreatedAssets.json`

<img width="1882" height="922" alt="image" src="https://github.com/user-attachments/assets/a2d55579-e9bf-427d-aff8-6b90413c6dc1" />

This job is the heart of the migration and does take a while, sometimes several hours for a larger migration.
 
 ---
 
### 3. **Fetch-Docs** Task
Fetch-Docs job requires a fresh `CookieJar.Json` file for downloading the source document files. You only need to refresh this file (see section below) a few times, but it is important to keep this file current when asked, especially for larger migrations.

<img width="1876" height="208" alt="image" src="https://github.com/user-attachments/assets/6c8f76c7-fab1-41f4-99ec-008bbe56e319" />

Sometimes the `KB` or `Document` Record that we reference from CSV doesnt include a `Filename` attribute, so we rely on some basic file parsing to identify filetype by encoding, file header, and magic bytes. *Downloaded files* are recored in `Debug` folder in a file named `Articles-Fetched.json`

<img width="1908" height="554" alt="image" src="https://github.com/user-attachments/assets/ef5bc50f-2545-4b9e-9d2e-acdbba7db674" />

---

### 4. **Create-Articles-FromFiles** Task
`Documents` are downloaded programmatically and converted/added to Hudu as Articles during the `Create-Articles-FromFiles` Job.
When this job finishes, all *Articles from Files* can be found in a folder in the project directory, named `Debug`. The assets dump file is named `Articles-FromFiles.json`.

This part is pretty simple. We use ***LibreOffice*** and the proven 'Articles-Anywhere' client library to convert any files that we were able to grab from ITPortal. This includes images, all manor of text documents, pdf's, presentations, and spreadsheets.

<img width="1712" height="850" alt="image" src="https://github.com/user-attachments/assets/b3a90c9b-04b9-4d84-aefc-0a89dc3ec943" />

There are **very few filetypes** that can't be handled this way, but if we encounter such a condition, we simply upload the file in-place and attach it to a reference article for easy searching and relating.

---

### 5. **Create-Articles-FromRecords** Task
`KBs` And `Documents` get created as `Articles` in Hudu from your CSV export During `Create-Articles-FromRecords` Job. This Job requires that you have a fresh `CookieJar.Json` File in order to download images in these articles.

<img width="1884" height="502" alt="image" src="https://github.com/user-attachments/assets/929efb4c-51b8-4de3-b7cf-126c57fcd293" />

When this job finishes, all *Articles from CSV* can be found in a folder in the project directory, named `Debug`. The assets dump file is named `Articles-FromRecords.json`. These are organized by *whether they started as a `kb` object or an embedded `document` object*.

---

### 6. **Submit-Passwords** Task

`Device-Passwords`, `Account-Passwords`, `LocalPasswords`, and `FieldPasswords` are created during the `Submit-Passwords` Job. During this job, new passwords also related to any `Articles` or `Assets` as-needed.

<img width="1034" height="558" alt="image" src="https://github.com/user-attachments/assets/f1158425-6442-469b-8102-c8b95f8d5f9b" />

When this job finishes, all *passwords from csv export* can be found in a folder in the project directory, named `Debug`. The assets dump file is named `PasswordsCreated.json`. They will be organized by the *type of password* they originally had been.

---

### 7. **Set-Relations** Task
`Relationships` are created during the `Set-Relations` job based on ITportal identifiers and companies they belong to. 
When this job finishes, all *relations from csv export* can be found in a folder in the project directory, named `Debug`. The assets dump file is named `relationsCreated.json`. They will be organized by the *type of relationship* they represent **(to or from an asset or article)**

<img width="1876" height="852" alt="image" src="https://github.com/user-attachments/assets/0cb5cb3c-6195-4bad-b6bd-3d16528255fb" />

---

**There is an optional wrap-up task as well, but that will become a thing of the past as these other tasks are improved upon. The only real clean-up that might need to be handled manually is renaming any fields that you want to be renamed or possibly listify-ing fields (turning them into `listselect`)**

## Cookie Jar
Install the `Cookie Manager` extension and export cookies to `cookiejar.json` in this folder.

http://chromewebstore.google.com/detail/cookiemanager-cookie-edit/hdhngoamekjhmnpenphenpaiindoinpo?hl=en
Then, log into itportal and move to `documents` page

<img width="1596" height="1076" alt="image" src="https://github.com/user-attachments/assets/3151e3aa-adec-4df8-b115-54a4094a0f2d" />

Next, you'll **click** the <span style="color:#f1c40f; font-weight:bold;">🟨 ```Cookie Manager icon``` (#1 in figure)</span>  
Then **click** <span style="color:#2ecc71; font-weight:bold;">🟩 ```Select All``` (#2 in figure)</span>  
Then **click** <span style="color:#9b59b6; font-weight:bold;">🟪 ```Export``` (#3 in figure)</span>

when prompted for a save location, you'll save or overwrite the file located in this project folder, named `cookiejar.json`
If your browser doesnt ask you where to save this file, it will be in your downloads folder. You can hit **CTL+J** to get to `recent downloads` quickly. From there, you'll just need to ***copy the item to project folder*** and ***rename it*** to `cookiejar.json`, right next to your environment file.

it's reccomended to do refresh your `cookiejar` file by repeating the above process when prompted for a second time (as not to miss any images in documents)

## Advanced Use

If you require ***more granular control*** over the `asset layouts` and `fields` in Hudu, you can optionally edit the file, named `fields-config.ps1`

This file contains some key variables that are referenced *when applying changes to assets*.
*More specifically*, if there is a *certain property or column* that you wish to be used for naming a specific object type, you can change ```$NameFields```. This hashtable is refrenced in the format: ```ObjectType = NameField.```

If there are `fields` you wish to *ignore across the board*, you can add them to ```$IgnoreFields``` `array`. It already has some sane defaults, fields that are used internally by ITPortal, but *you can add any other fields you don't want* to be added or read-in to Hudu.

If you want to map certain properties to certain fields for companies in Hudu, you can modify the $companyPropMap hashtable.
