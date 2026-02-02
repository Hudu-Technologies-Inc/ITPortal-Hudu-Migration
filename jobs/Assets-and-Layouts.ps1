

$companyMap = @{}; $assetsMap = @{};
$huducompanies = $huducompanies ?? $(get-huducompanies) ?? @()
$locationLayout = $locationLayout ?? $null
$CreatedAssets = $CreatedAssets ?? @{}
. "$project_workdir\fields-config.ps1"


$orderedKeys = $ITPortalData.Keys | Sort-Object {switch ($_) {
                                                'Sites'     { 1 }
                                                'Devices'   { 2 }
                                                'ConfigItems'   { 3 }
                                                'Contacts'   { 9 }
                                                default     { 8 }
                                            }}, { $_ }


foreach ($key in $orderedKeys | where-object {$_ -notin $SkipTables -and $_ -notin $PassTypes}) {

    $csvRows = @($ITPortalData[$key].CsvData)
    if ($specialObjectTypes.keys -contains $key.ToLowerInvariant()){
        $SpecialObjectType = $specialObjectTypes["$($key.ToLowerInvariant())"] ?? $null
        if ($null -eq $SpecialObjectType){write-host "No target for special object type $key"; continue;}
        write-host "Processing objects in $key as $SpecialObjectType"
        if (-not $CreatedAssets.containsKey($key)){$CreatedAssets["$key"] = @()}

        if ($SpecialObjectType -eq "companies"){
            write-host "updating details for $($ITPortalData.Companies.CsvData.Company_name.count) companies"
            foreach ($row in $csvRows) {
            
                $huduCompaniesRef = [ref]$huduCompanies
                $company = Ensure-HuduCompany -Row $row -InternalCompanyName $internalcompanyName -CompanyMap $CompanyMap -HuduCompanies $huduCompaniesRef
                $companyUpdateRequest = @{
                    Id = $company.id
                    Name = $company.name

                }
                foreach ($prop in $companyPropMap){
                    $val = $null
                    $rowVal = [string]$rowVal
                    if ([string]::IsNullOrWhiteSpace($rowVal)) { continue }
                    $huduField = $companyPropMap[$key]
                    if ($hudufield -ieq "NewWebsite"){
                        $url = Normalize-WebURL $val
                        $website = Get-HuduWebsites | where-object {$_.name -ieq $url} | select-object -first 1
                        if ($null -ne $website -and $null -ne $website.id){continue}
                        $website = New-HuduWebsite -companyId $company.id -name $url
                        continue
                    }
                    $companyUpdateRequest["$huduField"]=$rowVal                    
                }
                try {
                    $companyDetails = set-huducompany @companyUpdateRequest
                } catch {
                    write-error "error updating company $_"
                }
            }
        continue
    } elseif ($key -ieq "documents" -or $key -ieq "kbs"){

        # skip for now
        continue
    } else {
        write-host "Processing $key as asset"
    }
    $position = 0
    if (-not $csvRows -or $csvRows.Count -eq 0) {
        Write-Host "Loaded $key with 0 CSV rows; skipping type discernment"
        continue
    }
    $position ++
    $rowLimit = [Math]::Min($discernment, $csvRows.Count)
    Write-Host "Loaded $key with $($csvRows.Count) CSV rows; discerning types with resolution $rowLimit"

    $layoutRequest = @{
         Name   = $key; Fields = @();
         icon="fas fa-person"; color="#6136ff"; icon_color="#ffffff"; include_passwords=$true; include_photos=$true; include_comments=$true; include_files=$true;
    }

    #designate fields
    foreach ($label in $ITPortalData[$key].Properties) {
        # First, we map fields based on fields-config.
        $humanReadableLabel = ApplyUserLabelMappings -Label $label -LabelMappings $LabelMappings

        if ($label -ieq "Company"){continue}
        if ($label -ieq "SiteId" -or $label -ieq "SiteTwoID"){
            if ($null -ne $locationLayoutId -and $null -ne $locationLayout.id){
                Write-Host "`tField '$label' identified as Location Field as linkable: $($locationlayout.id)" -ForegroundColor cyan
                $layoutRequest.Fields += @{label = $humanReadableLabel; field_type = "AssetLink"; required=$false; position = $position; linkable_id=$locationLayout.id;}
                continue
            }
        }
        if ($IgnoreFields -contains $label){
            Write-Host "`tField '$label' is in ignore list; skipping." -ForegroundColor Yellow
            continue
        }
        if (LabelIsSecret -Label $label) {
            Write-Host "`tField '$label' identified as ConfidentialText Field" -ForegroundColor cyan
            $layoutRequest.Fields += @{label = $humanReadableLabel; field_type = "ConfidentialText"; required=$false; position = $position;}
            continue
        } elseif (LabelIsDate -label $label) {
            Write-Host "`tField '$label' identified as Date Field" -ForegroundColor cyan
            $layoutRequest.Fields += @{label = $humanReadableLabel; field_type = "Date"; required=$false; position = $position;}
            continue
        } elseif (LabelIsRichtext -label $label) {
            Write-Host "`tField '$label' identified as RichText Field" -ForegroundColor cyan
            $layoutRequest.Fields += @{label = $humanReadableLabel; field_type = "RichText"; required=$false; position = $position; }
            continue
        } elseif (LabelIsWebsite -label $label) {
            Write-Host "`tField '$label' identified as Website Field" -ForegroundColor cyan
            $layoutRequest.Fields += @{label = $humanReadableLabel; field_type = "Website"; required=$false; position = $position;}
            continue
        }
        
        $sampleValues = New-Object System.Collections.Generic.List[object]

        for ($i = 0; $i -lt $rowLimit; $i++) {
            # read in discernment-level of rows to test data
            $row = $csvRows[$i]
            if ($null -eq $row) { continue }

            # Safe access, works with 'Notes:', 'Some Field Name', etc.
            $prop = $row.PSObject.Properties[$label]
            if (-not $prop) { continue }

            $val = $prop.Value
            if ([string]::IsNullOrWhiteSpace([string]$val)) { continue }

            $sampleValues.Add($val)
        }

        if ($sampleValues.Count -eq 0) {
            $finalType = "Text"
        } elseif (Test-IsListSelect -InputObject $sampleValues -Resolution $rowLimit) {
            Write-Host "`tField '$label' identified as ListSelect Field" -ForegroundColor Cyan
            $list = $null
            $list = New-HuduList -name "$key - $label Options"; $list=$list.list ?? $list;
            $layoutRequest.Fields += @{label = $humanReadableLabel; field_type = "ListSelect"; required=$false; position = $position; list_id=$list.id}
            continue
        } else {
            try {
                $discern = Discern-Type -Objects $sampleValues -Resolution $rowLimit

                if ($discern -and $discern.BestType) {
                    $bestType   = $discern.BestType
                    $confidence = $discern.Confidence
                    $discern.Counts | Format-List
                    $finalType = $(if ($confidence -lt 0.7) {"Text"} else {$bestType})
                } else {$finalType = "Text"}
            } catch {$finalType = "Text"}
            $layoutRequest.Fields += @{label = $humanReadableLabel; field_type = $($finalType ?? "Text"); required=$false; position = $position;}
            Write-Host "`tField '$label' → $finalType" -ForegroundColor Cyan
        }
    }
    $HumanReadableLayoutName = ApplyUserLabelMappings -Label $key -LabelMappings $LabelMappings
    $AL = get-huduassetlayouts | Where-Object {$_.name -ieq $key -or $_.name -ieq $HumanReadableLayoutName} | select-object -first 1; $al = $AL.asset_layout ?? $AL;
    if ($al) {
        Write-Host "Asset Layout '$($al.name)' (ID: $($al.id)) already exists; using existing layout." -ForegroundColor Yellow
    } else {
        $AL = new-huduassetlayout @layoutRequest; $Al = $AL.asset_layout ?? $AL;

        $null = Set-HuduAssetLayout -id $al.id -Active $true
        $AL = get-huduassetlayouts -id $Al.id; $al = $AL.asset_layout ?? $AL;

        write-host "Created/Using Asset Layout '$($al.name)' (ID: $($al.id)) for type '$key'" -ForegroundColor Green
    }
    $ALFields = $al.fields
    if ($key -ieq "sites"){
        $locationLayout = $al
    }

    $ITPortalData[$key] | Add-Member -MemberType NoteProperty -Name "AssetLayout" -Value $AL -Force
    # with layout out of the way, we can loop thru assets

    foreach ($row in $csvRows) {
        $huduCompaniesRef = [ref]$huduCompanies
        $company = Ensure-HuduCompany -Row $row -InternalCompanyName $internalcompanyName -CompanyMap $CompanyMap -HuduCompanies $huduCompaniesRef

        if ($NameFields.ContainsKey($key)){
            $GivenName = $row.$($NameFields[$key]) ?? $row.PSObject.Properties[0].Value ?? "Unnamed $key Asset $($($([Guid]::NewGuid().ToString()) -split "-")[0])"
        } else {
        
            $GivenName = $row.name ?? $row.PSObject.Properties[0].Value ?? "Unnamed $key $($($([Guid]::NewGuid().ToString()) -split "-")[0])"
        }

        if (-not $company -or -not $company.id -or $company.id -le 0){
            $company = $internalCompany
        }
        if (-not $company -or -not $company.id -or $company.id -le 0){
            Write-Error "Uprocessable Asset $GivenName - cannot attribute to company ($($($company | convertto-json -depth 99).ToString())), not even internal!"
        }

        $assetRequest = @{
            AssetLayoutID = $al.id
            CompanyId     = $company.id
            Fields        = @()
            Name          = $GivenName
        }


        $existingAsset = Get-HuduASsets -CompanyId $company.id -name $GivenName -AssetLayoutId $al.id | select-object -first 1
        $existingAsset = $existingAsset.asset ?? $existingAsset
        if ($null -ne $existingAsset) {
            Write-Host "Asset '$GivenName' already exists under company '$($company.name)'; updating instead of creating." -ForegroundColor Yellow
            $assetRequest.Id = $existingAsset.id
        } else {
            Write-Host "Creating asset '$GivenName' under company '$($company.name)'." -ForegroundColor Green
        }




        foreach ($label in $ITPortalData[$key].Properties) {
            $humanReadableLabel = ApplyUserLabelMappings -Label $label -LabelMappings $LabelMappings
            $field = $ALFields | where-object { $_.label -ieq $label -or $_.label -ieq $humanReadableLabel } | select-object -first 1
            if (-not $field) {continue}
            if ($IgnoreFields -contains $label -or $IgnoreFields -contains $humanReadableLabel){
                Write-Host "`tField '$label' is in ignore list; skipping." -ForegroundColor Yellow
                continue
            }
            if ($LayoutSpecificIgnoreFields.ContainsKey($key)){
                if ($LayoutSpecificIgnoreFields[$key] -contains $label){
                    Write-Host "`tField '$label' is in layout-specific ignore list; skipping." -ForegroundColor Yellow
                    continue
                }
            }


            $prop = $row.PSObject.Properties[$label]
            if (-not $prop) { continue }

            $val = $null
            $val = $prop.Value
            if ([string]::IsNullOrWhiteSpace([string]$val)) { continue }
            if ($field.field_type -eq "Date"){
                $val = Get-CoercedDate $val
            } elseif ($field.field_type -eq "Number"){
                $val = Get-CastIfNumeric $val
            } elseif ($field.field_type -eq "Website"){
                $val = Get-NormalizedWebURL $val
            } elseif ($field.field_type -eq "CheckBox"){
                $val = Test-TruthyFalsy $val
                if ($null -eq $val) { continue }
                $val = "$($val)"
            }
            if ([string]::IsNullOrWhiteSpace([string]$val)) { continue }


            $assetRequest.Fields+=@{"$humanReadableLabel" = $val}
        }
        try {
            write-host "Setting/Creating asset '$GivenName'..."
            if ($assetRequest.containsKey("Id") -and $null -ne $assetRequest.Id -and $assetRequest.Id -gt 0) {
                $newAsset = Set-HuduAsset @assetRequest; $updatedAsset = $updatedAsset.asset ?? $updatedAsset;
            } else {
                $newAsset = New-HuduAsset @assetRequest; $newAsset = $newAsset.asset ?? $newAsset;
            }
            $CreatedAssets[$key] = @($CreatedAssets[$key]) + @($newAsset)
        } catch {
            write-error $_
        }
    }

}

$CreatedAssets | convertto-json -depth 99 | set-content -path $(join-path $debugDir -childpath "CreatedAssets.json") -force