

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

$allHuduAssets = Get-HuduAssets
$allhuduLayouts = Get-HuduAssetLayouts
$allHuduArticles = Get-HuduArticles

$contactsLayout = $allHuduLayouts | where-object {$_.name -ieq "Contacts"} | select-object -first 1
$devicesLayout = $allHuduLayouts | where-object {$_.name -ieq "Devices"} | select-object -first 1
$configurationsLayout = $allHuduLayouts | where-object {$_.name -ieq "ConfigItems"} | select-object -first 1
$sitesLayout = $allHuduLayouts | where-object {$_.name -ieq "Sites"} | select-object -first 1
$agreeMentsLayout = $allHuduLayouts | where-object {$_.name -ieq "Agreements"} | select-object -first 1

$contactsObjects = $allhuduassets | where-object {$_.asset_layout_id -eq $contactsLayout.id}
$devicesObjects = $allhuduassets | where-object {$_.asset_layout_id -eq $devicesLayout.id}
$configurationsObjects = $allhuduassets | where-object {$_.asset_layout_id -eq $configurationsLayout.id}
$sitesObjects = $allhuduassets | where-object {$_.asset_layout_id -eq $sitesLayout.id}
$agreeMentsObjects = $allhuduassets | where-object {$_.asset_layout_id -eq $agreeMentsLayout.id}
$allHuduCompanies = Get-HuduCompanies

if (get-command -name Set-HapiErrorsDirectory -ErrorAction SilentlyContinue){try {Set-HapiErrorsDirectory -skipRetry $false} catch {}}


foreach ($key in $orderedKeys | where-object {$_ -in @("Devices","Sites","KBs","Documents","ConfigItems","Contacts","Agreements")}) {
    $csvRows = @($ITPortalData[$key].CsvData)
        foreach ($row in $csvRows) {
            $deviceMatched = $row.DeviceID ?? $row.ForeignDeviceID ?? $null
            $siteMatched = $row.SiteID ?? $row.sitetwoid ?? $row.ForeignSiteID ?? $null
            $KBMatched = $row.KBID ?? $null
            $docMatched = $row.documentId ?? $null
            $configmatched = $row.ConfigObjectID ?? $row.ForeignConfigObjectID ?? $null
            $contactmatched = $row.ContactID ?? $row.ForeignContactID ?? $null
            $agreeMentMatched = $row.AgreementID ?? $null

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
            if ($null -eq $sourceObject){
                write-host " Could not find source object for $key with name $($key) = $($row.Name ?? $row.$($NameFields[$key])) in company $($company.name). skipping relations." -ForegroundColor Yellow
                continue
            }
            if ($key -in @("KBs","Documents")){
                $fromableType = "Article"
            } else {
                $fromableType = "Asset"
            }


            if ($null -ne $deviceMatched){
                $huduDevice = $null
                $ObjectCSV = $itportaldata.Devices.csvdata | where-object { $_.DeviceID -eq $deviceMatched } | select-object -first 1
                $huduDevice = $devicesObjects | where-object { $_.name -eq $ObjectCSV.Name -and $_.company -eq $company.name} | select-object -first 1
                if ($null -ne $huduDevice){
                    $relationsToAdd["Asset"] += $huduDevice
                }
            }
            if ($null -ne $siteMatched){
                $huduSite = $null
                $ObjectCSV = $itportaldata.Sites.csvdata | where-object { $_.SiteID -eq $siteMatched } | select-object -first 1
                $huduSite = $sitesObjects | where-object { $_.name -eq $ObjectCSV.SiteName -and $_.company -eq $company.name} | select-object -first 1
                if ($null -ne $huduSite){
                    $relationsToAdd["Asset"] += $huduSite
                }
            }
            if ($null -ne $KBMatched){
                $huduKB = $null
                $ObjectCSV = $itportaldata.KBs.csvdata | where-object { $_.KBID -eq $KBMatched } | select-object -first 1
                $huduKB = $allHuduArticles | where-object { $_.name -eq $ObjectCSV.kbname -and $_.company -eq $company.name} | select-object -first 1
                if ($null -ne $huduKB){
                    $relationsToAdd["Article"] += $huduKB
                }
            }
            if ($null -ne $docMatched){
                $huduKB = $null
                $ObjectCSV = $itportaldata.docs.csvdata | where-object { $_.documentId -eq $docMatched } | select-object -first 1
                $huduKB = $allHuduArticles | where-object { $_.name -eq $ObjectCSV.docname -and $_.company -eq $company.name} | select-object -first 1
                if ($null -ne $huduKB){
                    $relationsToAdd["Article"] += $huduKB
                }
            }
            if ($null -ne $configmatched){
                $huduConfig = $null
                $ObjectCSV = $itportaldata.ConfigItems.csvdata | where-object { $_.ConfigObjectID -eq $configmatched } | select-object -first 1
                $huduConfig = $configurationsObjects | where-object { $_.name -eq $ObjectCSV.ciName -and $_.company -eq $company.name} | select-object -first 1
                if ($null -ne $huduConfig){
                    $relationsToAdd["Asset"] += $huduConfig
                }
            }
            if ($null -ne $contactmatched){
                $huduContact = $null
                $ObjectCSV = $itportaldata.Contacts.csvdata | where-object { $_.ContactID -eq $contactmatched } | select-object -first 1
                $huduContact = $contactsObjects | where-object { $_.name -eq $ObjectCSV.FullName -and $_.company -eq $company.name} | select-object -first 1
                if ($null -ne $huduContact){
                    $relationsToAdd["Asset"] += $huduContact
                }
            }
            if ($null -ne $agreeMentMatched){
                $huduAgreeMent = $null
                $ObjectCSV = $itportaldata.Agreements.csvdata | where-object { $_.AgreementID -eq $agreeMentMatched } | select-object -first 1
                $huduAgreeMent = $agreeMentsObjects | where-object { $_.name -eq $ObjectCSV.AgreementName -and $_.company -eq $company.name} | select-object -first 1
                if ($null -ne $huduAgreeMent){
                    $relationsToAdd["Asset"] += $huduAgreeMent
                }
            }
        write-host " Adding relations from $fromableType $($sourceObject.name) in company $($company.name): Assets to add: $($relationsToAdd["Asset"].Count), Articles to add: $($relationsToAdd["Article"].Count)" -ForegroundColor Cyan
        foreach ($relationNeeded in $relationsToAdd["Asset"]){
            # $r = new-hudurelation -ToableType "Asset" -fromableType $fromableType -toableId $relationNeeded.id -FromableID $sourceObject.id
        }
        foreach ($relationNeeded in $relationsToAdd["Article"]){
            # $r = new-hudurelation -ToableType "Article" -fromableType $fromableType -toableId $relationNeeded.id -FromableID $sourceObject.id

        }

    }
}