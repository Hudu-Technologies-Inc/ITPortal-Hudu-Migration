$exportLocation = $exportLocation ?? (Read-Host "please enter the full path to your export.")
$hudubaseUrl    = $hudubaseUrl    ?? (Read-Host "please enter the hudubase url.")
$huduapikey     = $huduapikey     ?? (Read-Host "please enter your hudu api key.")
$internalCompanyName = $internalCompanyName ?? (Read-Host "please enter the internal company name to use for assets without a company.")

if (-not (Test-Path -Path $exportLocation)) {
    Write-Error "The specified path does not exist: $exportLocation"; exit;
} else {
    write-host "Using export location: $exportLocation"
}
$ITPortalData = $ITPortalData ?? @{}

foreach ($f in Get-ChildItem -Path $exportLocation -Recurse -File -Filter "*.txt") {
    Write-Host "reading $($f.FullName)"
    $contents = Get-Content $f.FullName -Raw
    $ITPortalData["$($f.BaseName)"] = $ITPortalData["$($f.BaseName)"] ??  @{CsvData = @(); TxtData = @()}
    $ITPortalData["$($f.BaseName)"].TxtData+=@{ filename = $f.FullName; content = $contents; name = $f.BaseName}
}

# Import & populate ITPortalData
foreach ($f in Get-ChildItem -Path $exportLocation -Recurse -File -Filter "*.csv") {
    Write-Host "reading $($f.FullName)"

    $contents = Import-Csv $f.FullName

    $key = $f.BaseName

    $ITPortalData[$key] = @{
        Filename   = $f.FullName
        CsvData    = @($contents)
        Properties = (Get-CSVPropertiesSafe$contents)  # assuming this returns header names
    }
}
$huducompanies = get-huducompanies
$companyMap = @{}
$assetsMap = @{}

$discernment = 3092
$position = 0
foreach ($key in $ITPortalData.Keys) {
    $position ++
    $csvRows = @($ITPortalData[$key].CsvData)

    if (-not $csvRows -or $csvRows.Count -eq 0) {
        Write-Host "Loaded $key with 0 CSV rows; skipping type discernment"
        continue
    }

    $rowLimit = [Math]::Min($discernment, $csvRows.Count)
    Write-Host "Loaded $key with $($csvRows.Count) CSV rows; discerning types with resolution $rowLimit"

    $layoutRequest = @{
         Name   = $key; Fields = @();
         icon="fas fa-person"; color="#6136ff"; icon_color="#ffffff"; include_passwords=$true; include_photos=$true; include_comments=$true; include_files=$true;

    }

    foreach ($label in $ITPortalData[$key].Properties) {

        if (LabelIsSecret -Label $label) {
            Write-Host "`tField '$label' identified as ConfidentialText Field" -ForegroundColor cyan
            $layoutRequest.Fields += @{label = $label; field_type = "ConfidentialText"; required=$false; position = $position;}
            continue
        } elseif (LabelIsDate -label $label) {
            Write-Host "`tField '$label' identified as Date Field" -ForegroundColor cyan
            $layoutRequest.Fields += @{label = $label; field_type = "Date"; required=$false; position = $position;}
            continue
        } elseif (LabelIsRichtext -label $label) {
            Write-Host "`tField '$label' identified as RichText Field" -ForegroundColor cyan
            $layoutRequest.Fields += @{label = $label; field_type = "RichText"; required=$false; position = $position; }
            continue
        } elseif (LabelIsWebsite -label $label) {
            Write-Host "`tField '$label' identified as Website Field" -ForegroundColor cyan
            $layoutRequest.Fields += @{label = $label; field_type = "Website"; required=$false; position = $position;}
            continue
        } elseif (LabelIsNumber -label $label) {
            Write-Host "`tField '$label' identified as Number Field" -ForegroundColor cyan
            $layoutRequest.Fields += @{label = $label; field_type = "Number"; required=$false; position = $position;}
            continue
        }
        
        $sampleValues = New-Object System.Collections.Generic.List[object]

        for ($i = 0; $i -lt $rowLimit; $i++) {
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
            $layoutRequest.Fields += @{label = $label; field_type = "ListSelect"; required=$false; position = $position; list_id=$list.id}
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
            $layoutRequest.Fields += @{label = $label; field_type = $($finalType ?? "Text"); required=$false; position = $position;}
            Write-Host "`tField '$label' → $finalType" -ForegroundColor Cyan
        }
    }

    $AL = get-huduassetlayouts -name $key | select-object -first 1; $al = $AL.asset_layout ?? $AL;
    if ($al) {
        Write-Host "Asset Layout '$($al.name)' (ID: $($al.id)) already exists; using existing layout." -ForegroundColor Yellow
    } else {
        $AL = new-huduassetlayout @layoutRequest; $Al = $AL.asset_layout ?? $AL;
        $AL = get-huduassetlayouts -id $Al.id; $al = $AL.asset_layout ?? $AL;
        write-host "Created/Using Asset Layout '$($al.name)' (ID: $($al.id)) for type '$key'" -ForegroundColor Green
    }
    

    $ALFields = $al.fields


    $ITPortalData[$key] | Add-Member -MemberType NoteProperty -Name "AssetLayout" -Value $AL -Force

    foreach ($row in $csvRows) {

        $companyName = $row.PSObject.Properties["Company"]?.Value

        if ([string]::IsNullOrWhiteSpace($companyName)) {
            $companyName = $internalCompanyName
        }
        $company = $null
        if ($companyMap.ContainsKey($companyName)) {
            $company = $companyMap[$companyName]
        } 
        
        $company = $company ?? $(Get-HuduCompanyFromName -CompanyName $companyName -HuduCompanies $huducompanies)
        if (-not $company) {
            $huducompanies = get-huducompanies
            $company = $(Get-HuduCompanyFromName -CompanyName $companyName -HuduCompanies $huducompanies)
        }
        $company = $company.company ?? $company
        if (-not $company) {
            Write-Host "Creating company '$companyName' for asset '$($row.name ?? $row.PSObject.Properties[0].Value)'" -ForegroundColor Yellow
            $companyRequest = @{
                Name = $companyName
            }
            $company = New-HuduCompany @companyRequest; $company = $company.company ?? $company;
            $company = get-huducompanies -id $company.id
        } else {
            Write-Host "Using existing company '$companyName' for asset '$($row.name ?? $row.PSObject.Properties[0].Value)'" -ForegroundColor Green
        }
        if (-not $company) {
            Write-Error "Failed to create or retrieve company '$companyName' for asset '$($row.name ?? $row.PSObject.Properties[0].Value)'; skipping asset creation."
            continue
        }
        if (-not $companyMap.ContainsKey($companyName)) {
            $companyMap["$companyName"] = $company
        }
        $GivenName = $row.name ?? $row.PSObject.Properties[0].Value ?? "Unnamed $key Asset $($($([Guid]::NewGuid().ToString()) -split "-")[0])"

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
            $field = $ALFields | where-object { $_.label -ieq $label } | select-object -first 1
            if (-not $field) {continue}



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


            $assetRequest.Fields+=@{$label = $val}
        }
        try {
            write-host "Setting/Creating asset '$GivenName'..."
            if ($assetRequest.containsKey("Id") -and $null -ne $assetRequest.Id -and $assetRequest.Id -gt 0) {
                $newAsset = Set-HuduAsset @assetRequest; $updatedAsset = $updatedAsset.asset ?? $updatedAsset;
            } else {
                $newAsset = New-HuduAsset @assetRequest; $newAsset = $newAsset.asset ?? $newAsset;
            }
        } catch {
            write-error $_
        }
    }

}
