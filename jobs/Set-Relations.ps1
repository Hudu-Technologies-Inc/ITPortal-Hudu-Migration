

# typical system-wide identifiers
# AccountID
# AccountManagerID
# AccountTypeID
# addressID
# AddressID
# AgreementID
# AgreementSubTypeID
# AgreementTypeID
# CabinetFacilityID
# CabinetID
# CabinetSiteID
# ciDescription
# ClientID
# ClientTwoID
# ClientTypeID
# CollectorID
# ConfigObjectID
# ConfigObjectTypeID
# ContactID
# ContactTypeID
# CustomerID
# DeletedByID
# deviceDeviceID
# DeviceID
# DeviceTypeID
# DocumentID
# DocumentTypeID
# DSubcategoryID
# FacilityID
# ForeignClientID
# ForeignConfigObjectID
# ForeignConfigObjectTypeID
# ForeignContactID
# ForeignDeviceID
# ForeignDeviceTypeID
# ForeignSiteID
# gwID
# IPID
# KBCategoryID
# KBID
# KBSubCategoryID
# LogoID
# middleinitial
# MiddleInitial
# objectGUID
# ParentClientID
# RaidController
# ScoreboardID
# SecondaryUserID
# SiteID
# SiteTwoID
# SubnetID
# userid
# UserID
# VLANID

# Contacts                       
# Companies                      
# KBs                            
# Agreements                     
# Passwords-FieldPasswords       
# Sites                          
# ConfigItems                    
# Passwords-LocalPasswords       
# Accounts                       
# Devices                        
# Passwords-Devices              
# Documents                      
# IPNetworks                     
# Passwords-Accounts             

$NameFields = @{
    "Contacts"="FullName"
    "ConfigItems"="ciName"

}

# $allHuduAssets = Get-HuduAssets
# $allhuduLayouts = Get-HuduAssetLayouts
# $allHuduArticles = Get-HuduArticles

# $contactsLayout = $allHuduLayouts | where-object {$_.name -ieq "Contacts"} | select-object -first 1
# $devicesLayout = $allHuduLayouts | where-object {$_.name -ieq "Devices"} | select-object -first 1
# $configurationsLayout = $allHuduLayouts | where-object {$_.name -ieq "ConfigItems"} | select-object -first 1
# $sitesLayout = $allHuduLayouts | where-object {$_.name -ieq "Sites"} | select-object -first 1
# $agreeMentsLayout = $allHuduLayouts | where-object {$_.name -ieq "Agreements"} | select-object -first 1

# $contactsObjects = $allhuduassets | where-object {$_.asset_layout_id -eq $contactsLayout.id}
# $devicesObjects = $allhuduassets | where-object {$_.asset_layout_id -eq $devicesLayout.id}
# $configurationsObjects = $allhuduassets | where-object {$_.asset_layout_id -eq $configurationsLayout.id}
# $sitesObjects = $allhuduassets | where-object {$_.asset_layout_id -eq $sitesLayout.id}
# $agreeMentsObjects = $allhuduassets | where-object {$_.asset_layout_id -eq $agreeMentsLayout.id}
# $allHuduCompanies = Get-HuduCompanies

if (get-command -name Set-HapiErrorsDirectory -ErrorAction SilentlyContinue){try {Set-HapiErrorsDirectory -skipRetry $true} catch {}}

function Normalize-Key([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  return ($s.Trim() -replace '\s+',' ').ToLowerInvariant()
}

# CSV index by ID
$csvBy = @{
  Devices    = @{}
  Sites      = @{}
  Contacts   = @{}
  ConfigItems= @{}
  Agreements = @{}
  KBs        = @{}
  Documents  = @{}
}
foreach ($r in $itportaldata.Devices.csvdata)    { $csvBy.Devices["$($r.DeviceID)"] = $r }
foreach ($r in $itportaldata.Sites.csvdata)      { $csvBy.Sites["$($r.SiteID)"] = $r }
foreach ($r in $itportaldata.Contacts.csvdata)   { $csvBy.Contacts["$($r.ContactID)"] = $r }
foreach ($r in $itportaldata.ConfigItems.csvdata){ $csvBy.ConfigItems["$($r.ConfigObjectID)"] = $r }
foreach ($r in $itportaldata.Agreements.csvdata) { $csvBy.Agreements["$($r.AgreementID)"] = $r }
foreach ($r in $itportaldata.KBs.csvdata)        { $csvBy.KBs["$($r.KBID)"] = $r }
foreach ($r in $itportaldata.Documents.csvdata)  { $csvBy.Documents["$($r.DocumentID)"] = $r }
$internalCompany = Get-OrSetInternalCompany -internalCompanyName $internalCompanyName

$relationsCreated = @{}
foreach ($key in $orderedKeys | where-object {$_ -in @("Devices","Sites","KBs","Documents","ConfigItems","Contacts","Agreements")}) {
    $csvRows = @($ITPortalData[$key].CsvData)
        foreach ($row in $csvRows) {
            $deviceMatched = if ($key -eq 'Devices') { $row.ForeignDeviceID } else { $row.DeviceID ?? $row.ForeignDeviceID }
            $siteMatched   = if ($key -eq 'Sites')   { $row.ForeignSiteID ?? $row.SiteTwoID } else { $row.SiteID ?? $row.SiteTwoID ?? $row.ForeignSiteID }
            $KBMatched = if ($key -eq "KBs") { $null } else { $row.KBID ?? $null }
            $docMatched = if ($key -eq "Documents") {$null} else {$row.documentId ?? $null}
            $configmatched = if ($key -eq "ConfigItems") {$row.ForeignConfigObjectID ?? $null} else {$row.ConfigObjectID ?? $row.ForeignConfigObjectID ?? $null}
            $contactmatched = if ($key -eq "Contacts") {$row.ForeignContactID ?? $null} else { $row.ContactID ?? $row.ForeignContactID ?? $null}
            $agreeMentMatched = if ($key -eq "Agreements") {$null} else {$row.AgreementID ?? $null}

            $relationsToAdd = @{
                Asset=@(); Article=@();
            }
            $company = $null; $company = $allHuduCompanies | where-object {$_.name -ieq $row.company} | Select-Object -First 1; $company = $company.company ?? $company;
            $company = $company ?? $internalCompany
            $sourceObject = $null
            if ($key -eq "Devices"){
                $sourceObject = $devicesObjects | where-object { $_.name -eq $row.Name -and $_.company_id -eq $company.id} | select-object -first 1
            } elseif ($key -eq "Sites"){
                $sourceObject = $sitesObjects | where-object { $_.name -eq $row.Name -and $_.company_id -eq $company.id} | select-object -first 1
            } elseif ($key -eq "KBs"){
                $sourceObject = $allHuduArticles | where-object { $_.name -eq $row.kbname -and $_.company_id -eq $company.id} | select-object -first 1
            } elseif ($key -eq "Documents"){
                $sourceObject = $allHuduArticles | where-object { $_.name -eq $row.docname -and $_.company_id -eq $company.id} | select-object -first 1
            } elseif ($key -eq "ConfigItems"){
                $sourceObject = $configurationsObjects | where-object { $_.name -eq $row.ciName -and $_.company_id -eq $company.id} | select-object -first 1
            } elseif ($key -eq "Contacts"){
                $sourceObject = $contactsObjects | where-object { $_.name -eq $row.FullName -and $_.company_id -eq $company.id} | select-object -first 1
            } elseif ($key -eq "Agreements"){
                $sourceObject = $agreeMentsObjects | where-object { $_.name -eq $row.Name -and $_.company_id -eq $company.id} | select-object -first 1
            }

            $srcName = $row.Name
            if ($NameFields.ContainsKey($key)) { $srcName = $row.($NameFields[$key]) }

            if ($null -eq $sourceObject){
                write-host " Could not find source object for $key with name $($key) = $srcName in company $($company.name). skipping relations." -ForegroundColor Yellow
                continue
            }
            if ($key -in @("KBs","Documents")){
                $fromableType = "Article"
            } else {
                $fromableType = "Asset"
            }

            if ($null -ne $deviceMatched){
                $huduDevice = $null
                $ObjectCSV = $csvBy.Devices["$deviceMatched"]
                write-host "$($($objectcsv | convertto-json -depth 99).ToString())"
                $huduDevice = $devicesObjects | where-object { $_.name -eq $ObjectCSV.Name -and $_.company_id -eq $company.id } | select -first 1
                if ($null -ne $huduDevice){
                    $relationsToAdd["Asset"] += $huduDevice
                }
            }
            if ($null -ne $siteMatched){
                $huduSite = $null
                $ObjectCSV = $csvBy.Sites["$siteMatched"]
                write-host "$($($objectcsv | convertto-json -depth 99).ToString())"
                $huduSite = $sitesObjects | where-object { $_.name -eq $ObjectCSV.Name -and $_.company_id -eq $company.id } | select -first 1
                if ($null -ne $huduSite){
                    $relationsToAdd["Asset"] += $huduSite
                }
            }
            if ($null -ne $KBMatched){
                $huduKB = $null
                $ObjectCSV = $csvBy.KBs["$KBMatched"]
                write-host "$($($objectcsv | convertto-json -depth 99).ToString())"
                $huduKB = $allHuduArticles | where-object { $_.name -eq $ObjectCSV.kbname-and $_.company_id -eq $company.id } | select -first 1
                if ($null -ne $huduKB){
                    $relationsToAdd["Article"] += $huduKB
                }
            }
            if ($null -ne $docMatched){
                $huduKB = $null
                $ObjectCSV = $csvBy.Documents["$docMatched"]
                write-host "$($($objectcsv | convertto-json -depth 99).ToString())"
                $huduKB = $allHuduArticles | where-object { $_.name -eq $ObjectCSV.KBName -and $_.company_id -eq $company.id } | select -first 1
                if ($null -ne $huduKB){
                    $relationsToAdd["Article"] += $huduKB
                }
            }
            if ($null -ne $configmatched){
                $huduConfig = $null
                $ObjectCSV = $csvBy.ConfigItems["$configMatched"]
                write-host "$($($objectcsv | convertto-json -depth 99).ToString())"
                $huduConfig = $configurationsObjects | where-object { $_.name -eq $ObjectCSV.ciName -and $_.company_id -eq $company.id } | select -first 1
                if ($null -ne $huduConfig){
                    $relationsToAdd["Asset"] += $huduConfig
                }
            }
            if ($null -ne $contactmatched){
                $huduContact = $null
                $ObjectCSV = $csvBy.Contacts["$contactmatched"]
                write-host "$($($objectcsv | convertto-json -depth 99).ToString())"
                $huduContact = $contactsObjects | where-object { $_.name -eq $ObjectCSV.Name -and $_.company_id -eq $company.id } | select -first 1
                if ($null -ne $huduContact){
                    $relationsToAdd["Asset"] += $huduContact
                }
            }
            if ($null -ne $agreeMentMatched){
                $huduAgreeMent = $null
                $ObjectCSV = $csvBy.Agreements["$agreeMentMatched"]
                write-host "$($($objectcsv | convertto-json -depth 99).ToString())"
                $huduAgreeMent = $agreeMentsObjects | where-object { $_.name -eq $ObjectCSV.Name -and $_.company_id -eq $company.id } | select -first 1
                if ($null -ne $huduAgreeMent){
                    $relationsToAdd["Asset"] += $huduAgreeMent
                }
            }
        write-host " Adding relations from $fromableType $($sourceObject.name) in company $($company.name): Assets to add: $($relationsToAdd["Asset"].Count), Articles to add: $($relationsToAdd["Article"].Count)" -ForegroundColor Cyan
        foreach ($relationNeeded in $relationsToAdd["Asset"]){
                $r = $null
                $r = new-hudurelation -ToableType "Asset" -fromableType $fromableType -toableId $relationNeeded.id -FromableID $sourceObject.id
                if ($null -ne $r){
                    $relationsCreated["$($fromableType)_$($sourceObject.id)_to_Asset_$($relationNeeded.id)"] = $r
                }
        }
        foreach ($relationNeeded in $relationsToAdd["Article"]){
                $r = $null
                $r = new-hudurelation -ToableType "Article" -fromableType $fromableType -toableId $relationNeeded.id -FromableID $sourceObject.id
                if ($null -ne $r){
                    $relationsCreated["$($fromableType)_$($sourceObject.id)_to_Article_$($relationNeeded.id)"] = $r
                }
       }

    }
}

if (get-command -name Set-HapiErrorsDirectory -ErrorAction SilentlyContinue){try {Set-HapiErrorsDirectory -skipRetry $false} catch {}}
$ArticleMatches | convertto-json -depth 99 | set-content -path $(join-path $debugDir -childpath "relationsCreated.json") -force