

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

$orderedKeys = $orderedKeys ?? $ITPortalData.Keys | 
                        where-object {$_ -notin $SkipTables -and $_ -notin $PassTypes -and -not $_ -ieq "Companies"} | 
                        Sort-Object {switch ($_) {
                            'Sites'     { 0 }
                            'Devices'   { 2 }
                            'Configurations'   { 3 }
                            'Contacts'   { 9 }
                            default     { 8 }
                        }}, { $_ }

$allHuduAssets = Get-HuduAssets
$allhuduLayouts = Get-HuduAssetLayouts
$allHuduArticles = Get-HuduArticles




foreach ($key in $orderedKeys) {
    $csvRows = @($ITPortalData[$key].CsvData)
        foreach ($row in $csvRows) {
            $clientMatched = $row.ClientID ?? $row.clientTwoId ?? $row.ForeignClientID ?? $null
            $deviceMatched = $row.DeviceID ?? $row.ForeignDeviceID ?? $null
            $siteMatched = $row.SiteID ?? $row.sitetwoid ?? $row.ForeignSiteID ?? $null
            $KBMatched = $row.KBID ?? $row.documentId ?? $null
            $configmatched = $row.ConfigObjectID ?? $row.ForeignConfigObjectID ?? $null
            $contactmatched = $row.ContactID ?? $row.ForeignContactID ?? $null
            $agreeMentMatched = $row.AgreementID ?? $null
            $accountMatched = $row.AccountID ?? $null
            $company = $null; $company = get-huducompanies -name $row.company | Select-Object -First 1; $company = $company.company ?? $company;

            foreach ($matchableObject in @(
                @{ Field = "ClientID";         Value = $clientMatched },
                @{ Field = "DeviceID";         Value = $deviceMatched },
                @{ Field = "SiteID";           Value = $siteMatched },
                @{ Field = "KBID";             Value = $KBMatched },
                @{ Field = "ConfigObjectID";   Value = $configmatched },
                @{ Field = "ContactID";        Value = $contactmatched },
                @{ Field = "AgreementID";      Value = $agreeMentMatched },
                @{ Field = "AccountID";        Value = $accountMatched }
            )) {
                if ($null -ne $matchableObject.Value) {
                    if ($ArticleMatches.ContainsKey($matchKey)) {
                        
                        $HuduArticle = $null
                        $HuduArticle = Get-Huduarticles -name $kbmatched.Name
                        if ($null -ne $company){$huduarticle = $huduarticle | where-object {$_.companyid -eq $company.id}}
                        

                    }
                }
            }






        }
}