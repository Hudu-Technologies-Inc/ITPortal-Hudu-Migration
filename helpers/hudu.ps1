function Get-OrSetInternalCompany {
    param ([string]$internalCompanyName)
    $internalCompany = $null
    $internalCompany = get-huducompanies -name $internalCompanyName | select-object -first 1
    $internalCompany = $internalCompany.company ?? $internalCompany
    if ($null -eq $internalCompany -or $internalCompany.id -lt 1){
        New-HuduCompany -name "$internalCompanyName"
        $internalCompany = get-huducompanies -name $internalCompanyName | select-object -first 1
        $internalCompany = $internalCompany.company ?? $internalCompany
    }
    return $internalCompany
}

function New-HuduAddress {
    param([Parameter(Mandatory)][object]$Input)

    # Parse JSON strings; pass through objects
    $o = if ($Input -is [string]) { try { $Input | ConvertFrom-Json } catch { $null } } else { $Input }
    if (-not $o) { return $null }

    # Helper to grab the first present alias
    $first = {
        param($obj, [string[]]$names)
        foreach ($n in $names) { if ($obj.PSObject.Properties.Name -contains $n) { return $obj.$n } }
        return $null
    }

    $addr1 = & $first $o @('address_line_1','address1','address_1','line1','street','street1','address')
    $addr2 = & $first $o @('address_line_2','address2','address_2','line2','street2')
    $city           = & $first $o @('city','town')
    $state          = & $first $o @('state','province','region')
    $zip            = & $first $o @('zip','zipcode','postal','postal_code')
    $cntry   = & $first $o @('country_name','country')
    if ($addr1 -or $addr2 -or $city -or $state -or $zip -or $cntry) {
    $NewAddress = [ordered]@{
        address_line_1 = $addr1
        city           = $city
        state          = $state
        zip            = $zip
        country_name   = $cntry
    }
    if ($addr2) { $NewAddress['address_line_2'] = $addr2 }
    return $NewAddress
    } else {return $null}
}

function Resolve-LocationForCompany {
  param(
    [Parameter(Mandatory)][int]$CompanyId,
    [Parameter(Mandatory)]$Row,
    [Parameter(Mandatory)]$AllHuduLocations,
    [string[]]$Hints
  )

  # If $Hints can be null at call-time, set a safe default here (avoid default param expr that depends on external vars)
  if (-not $Hints) { $Hints = @('location','branch','office','site','building') }

  $candKeys = @()
  foreach ($prop in $Row.PSObject.Properties) {
    $propName = $prop.Name
    if ($Hints.Where({ param($h) (Test-Fuzzy $propName $h) }, 'First')) {
      $candKeys += $propName
    }
  }
  $candKeys = $candKeys | Sort-Object -Unique

  $candVals = @()
  foreach ($k in $candKeys) {
    $v = $Row.$k
    if ($null -ne $v -and "$v".Trim()) { $candVals += "$v" }
  }
  if (-not $candVals) { return $null }

  $companyLocs = $AllHuduLocations | Where-Object { $_.company_id -eq $CompanyId }
  foreach ($cv in $candVals) {
    $hit = $companyLocs | Where-Object { test-equiv -A $_.name -B $cv } | Select-Object -First 1
    if ($hit) { return $hit }
  }
  return $null
}
function Get-ListItemFuzzy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][int]$ListId,
        [double]$MinSimilarity = 0.85   # tweak as needed
    )

    if ([string]::IsNullOrWhiteSpace($Source)) { return $null }

    $list = Get-HuduLists -Id $ListId
    if (-not $list -or -not $list.list_items) { return $null }

    $sNorm = Normalize-Text $Source

    $bestItem  = $null
    $bestScore = -1.0

    foreach ($item in $list.list_items) {
        $name = [string]$item.name
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $nNorm = Normalize-Text $name

        $score = if ($nNorm -eq $sNorm) { 1.0 } else { Get-Similarity $name $Source }
        if ($nNorm.StartsWith($sNorm) -or $sNorm.StartsWith($nNorm)) {
            $score = [Math]::Min(1.0, $score + 0.02)
        }

        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestItem  = $item
        }
    }

    if ($bestScore -lt $MinSimilarity) { return $null }
    return $bestItem
}

function Get-UniqueHuduListName {
  param([Parameter(Mandatory)][string]$BaseName,[bool]$allowReuse=$false)

  $name = $BaseName.Trim()
  $i = 0
  while ($true) {
    $existing = Get-HuduLists -name $name
    if (-not $existing) { return $name }
    if ($existing -and $true -eq $allowReuse) {return $existing}
    $i++
    $name = "{0}-{1}" -f $BaseName.Trim(), $i
  }
}

function Get-HuduLayoutLike {
  param ([array]$LabelSet)

  foreach ($layout in $(get-huduassetlayouts)){
    foreach ($Label in $LabelSet){
      if ($true -eq $(Test-Equiv -A $locationLabel -$layout.name)){
        return $layout
      }
    }
  }
  Write-Host "No location layout found. Ensure your location layout name is in LocationLayoutNames array ($LocationLayoutNames)"
  return $null
}
function Get-HuduCompanyFromName {
    # use index first. Then existing list. Then API call.
    param (
        [Parameter(Mandatory = $true)]
        [string]$CompanyName,
        [array]$HuduCompanies,
        [bool]$includenicknames = $false,
        [array]$existingIndex = $null
    )
    if ([string]::IsNullOrWhiteSpace($CompanyName)) { return $null }
    $matchedCompany = $null
    if ($existingIndex -ne $null -and $existingIndex.count -gt 0){
        $matchedCompany = $matchedCompany ?? $existingIndex | where-object {
            ($_.CompanyName -ieq $CompanyName) -or
            [bool]$(test-equiv -A $_.CompanyName -B $CompanyName) } | Select-Object -First 1
        if ($includenicknames){
            $matchedCompany = $matchedCompany ?? $existingIndex | where-object {
                (-not [string]::IsNullOrWhiteSpace($_.HuduObject.nickname)) -and (
                    ($_.HuduObject.nickname -ieq $CompanyName) -or
                    [bool]$(test-equiv -A $_.HuduObject.nickname -B $CompanyName))
            } | Select-Object -First 1
        }
    }


    $matchedCompany = $matchedCompany ?? $HuduCompanies | where-object {
            ($_.name -ieq $CompanyName) -or
            [bool]$(test-equiv -A $_.name -B $CompanyName)`
        } | Select-Object -First 1
    $matchedCompany = $matchedCompany ?? $(Get-HuduCompanies -Name $CompanyName | select-object -first 1)
    $matchedCompany = $matchedCompany ?? (get-huducompanies | where-object {[bool]$(test-equiv -A $_.name -B $CompanyName)} | select-object -first 1)
    
    if ($includenicknames){
        $matchedCompany = $HuduCompanies | where-object {
                ($_.nickname -ieq $CompanyName) -or
                [bool]$(test-equiv -A $_.nickname -B $CompanyName)`
            } | Select-Object -First 1
        $matchedCompany = $matchedCompany ?? (get-huducompanies | where-object {[bool]$(test-equiv -A $_.name -B $CompanyName)} | select-object -first 1)
    }
    return $matchedCompany
}

function Get-HuduFieldValue {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][object]$Asset,
    [Parameter(Mandatory)][string[]]$Labels
  )
  $labelsLC = $Labels | ForEach-Object { $_.ToLower() }
  $Asset.fields |
    Where-Object { $_.label -and $labelsLC -contains ([string]$_.label).ToLower() } |
    Select-Object -First 1 -ExpandProperty value
}

function Get-HuduAssetFromName {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [int]$AssetLayoutId,
        [array]$Assets
    )
    if ([string]::IsNullOrWhiteSpace($Name) -or -not $AssetLayoutId -or $AssetLayoutId -lt 1) { return $null }
    $matchedAsset = $null
    $matchedAsset = $Assets | where-object {
            ($_.name -ieq $Name) -or
            [bool]$(test-equiv -A $_.name -B $Name)`
        } | Select-Object -First 1
    $matchedAsset = $matchedAsset ?? 
        $(Get-HuduAssets -AssetLayoutId $AssetLayoutId -Name $CompanyName) ?? 
         (get-huduassets -AssetLayoutId $AssetLayoutId | where-object {[bool]$(test-equiv -A $_.name -B $Name)} | select-object -first 1)
    return $matchedCompany
}

function Get-OtpAuthParams {
    param([Parameter(Mandatory)][string]$Uri)

    $qIndex = $Uri.IndexOf('?')
    $result = [ordered]@{ secret = $null; issuer = $null }

    if ($qIndex -lt 0 -or $qIndex -eq $Uri.Length-1) { return [pscustomobject]$result }
    $qs = $Uri.Substring($qIndex + 1)

    foreach ($pair in $qs.Split('&', [System.StringSplitOptions]::RemoveEmptyEntries)) {
        $eq = $pair.IndexOf('=')
        if ($eq -lt 0) { continue }

        $k = $pair.Substring(0, $eq)
        $v = [uri]::UnescapeDataString($pair.Substring($eq + 1))

        if ($k -ieq 'secret') { $result.secret = ("$v" -replace '\s', '') }
        elseif ($k -ieq 'issuer') { $result.issuer = $v }
    }

    [pscustomobject]$result
}
function ConvertTo-ValidatedOtpSecret {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline, Mandatory)]
        [string]$InputObject,

        [int]$MinLen = 16,
        [int]$MaxLen = 80
    )

    process {
        $raw = ($InputObject ?? '').Trim()
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return [pscustomobject]@{
                IsValid  = $false
                Secret   = $null
                Issuer   = $null
                Reason   = 'Empty'
                Raw      = $InputObject
                Source   = 'none'
            }
        }

        # If it's an otpauth URI, extract secret and issuer
        $issuer = $null
        $candidate = $raw
        $source = 'raw'

        if ($raw -like 'otpauth://*') {
            $p = Get-OtpAuthParams -Uri $raw
            $candidate = $p.secret
            $issuer = $p.issuer
            $source = 'otpauth'
        }

        $candidate = ($candidate ?? '').Trim().ToUpperInvariant()

        $candidate = $candidate -replace '[\s-]', ''
        $candidate = $candidate.TrimEnd('=')

        $isValidBase32 = $candidate -match '^[A-Z2-7]+$'
        $lengthOK      = $candidate.Length -ge $MinLen -and $candidate.Length -le $MaxLen

        $isValid = $isValidBase32 -and $lengthOK

        [pscustomobject]@{
            IsValid  = $isValid
            Secret   = if ($isValid) { $candidate } else { $null }
            Issuer   = $issuer
            Reason   = if ($isValid) { $null } else {
                if (-not $isValidBase32 -and -not $lengthOK) { "NotBase32;BadLength($($candidate.Length))" }
                elseif (-not $isValidBase32) { "NotBase32" }
                else { "BadLength($($candidate.Length))" }
            }
            Raw      = $raw
            Source   = $source
        }
    }
}



function ConvertTo-HashtableDeep {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        $InputObject
    )

    process {
        if ($null -eq $InputObject) { return $null }

        if ($InputObject -is [System.Collections.IDictionary]) {
            $h = @{}
            foreach ($k in $InputObject.Keys) {
                $h[$k] = ConvertTo-HashtableDeep -InputObject $InputObject[$k]
            }
            return $h
        }

        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
            return @($InputObject | ForEach-Object { ConvertTo-HashtableDeep -InputObject $_ })
        }

        if ($InputObject -is [psobject]) {
            $h = @{}
            foreach ($p in $InputObject.PSObject.Properties) {
                $h[$p.Name] = ConvertTo-HashtableDeep -InputObject $p.Value
            }
            return $h
        }

        return $InputObject
    }
}
function Omni-Relate {
    
    function _Normalize-AssetName {
        param([string]$Name)
        if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
        $n = $Name.Normalize([Text.NormalizationForm]::FormKC)
        $n = $n -replace '&nbsp;',' ' -replace '\s+',' '
        $n = $n.Trim().ToLowerInvariant()
        return $n
    }

    function _Normalize-WebsiteURL {
        param([string]$Url)
        if ([string]::IsNullOrWhiteSpace($Url)) { return $null }
        $u = $Url.Trim()
        if ($u -match '^(https?://)?(?<host>[^/]+)(?<rest>/.*)?$') {
            $hostname = $matches.host.ToLowerInvariant()
            return $hostname.ToLowerInvariant()
        }
        return $u.ToLowerInvariant()
    }

    if (get-command -name Set-HapiErrorsDirectory -ErrorAction SilentlyContinue){try {Set-HapiErrorsDirectory -skipRetry $true} catch {}}
    write-host "getting companies"; $allcompanies = get-huducompanies;
    write-host "$($allAssets.count) assets (please be patient, this can take some time.)"; $allAssets = get-huduassets -CompanyId $companyID;
    write-host "$($allWebsites.count) websites"; $allWebsites = get-huduwebsites | where-object {$_.company_id -eq $companyID};
    write-host "$($allArticles) articles (please be patient, this can take some time.)"; $allArticles = get-huduarticles -CompanyId $companyID;

    foreach ($c in $allcompanies) { 

        $companyAssets = $allAssets | Where-Object { $_.company_id -eq $c.id }
        $companywebsites = $allWebsites | Where-Object { $_.company_id -eq $c.id }
        $companyArticles = $allArticles | Where-Object { $_.company_id -eq $c.id }

        $companyAssetsByName = $companyAssets | Group-Object { _Normalize-AssetName $_.name } -AsHashTable -AsString

        foreach ($a in $companyAssets) {
            $normalizedAssetName = _Normalize-AssetName $a.name

            $mentionedWebsites = @()
            $mentionedArticles = @()
            $mentionedAssets = @()

            # websites and articles matched by name/description -> name of asset
            if ($companywebsites) {
                $mentionedWebsites = $companywebsites | Where-Object { $_.Notes -and $_.Notes.Contains($normalizedAssetName) }
            }
            if ($companyArticles) {
                $mentionedArticles = $companyArticles | Where-Object { ($_.Name -and $_.Content.Contains($a.name)) -or ($_.Content -and $_.Content.Contains($normalizedAssetName)) }
            }


            # websites where name or url is mentioned in text/richtext fields of the asset
            # articles with content or name mentioned in a website field (either website field or text/richtext fields)
            $a.fields | Where-Object {$_.field_type -eq "Website"} | ForEach-Object {
                $fieldValue = $_.value
                $mentionedWebsites += $companywebsites | Where-Object { "$(_Normalize-WebsiteURL $fieldValue)*" -ilike "$(_Normalize-WebsiteURL $_.name)*" -or $_.name -icontains "$(_Normalize-WebsiteURL $fieldValue)" -or $_.name -icontains $normalizedAssetName }
                $mentionedArticles += $companyArticles | Where-Object { $_.content -and $_.content.Contains("$(_Normalize-WebsiteURL $fieldValue)") -or ($_.Name -and $_.Name.Contains("$(_Normalize-WebsiteURL $fieldValue)")) }
                $mentionedAssets += $companyAssets | Where-Object { $_.name -and $_.name.Contains("$(_Normalize-WebsiteURL $fieldValue)") }
            }
            $a.fields | Where-Object {$_.field_type -eq "RichText"} | ForEach-Object {
                $fieldValue = $_.value
                $mentionedWebsites += $companywebsites | Where-Object { $fieldValue -icontains $normalizedAssetName -or $(_Normalize-AssetName $_.name) -ieq $normalizedAssetName -or $fieldValue -icontains $_.name -or $_.notes -icontains $normalizedAssetName -or $_.notes -icontains $a.name }
                $mentionedArticles += $companyArticles | Where-Object { $_.content -and $_.content.Contains($normalizedAssetName) -or $_.content -icontains $a.name -or $normalizedAssetName -ieq (_Normalize-AssetName $_.name) }
                $mentionedAssets += $companyAssets | Where-Object { $fieldValue -and $fieldValue.Contains($normalizedAssetName) -or $fieldValue.Contains($a.name) }
            }       
            $a.fields | Where-Object {$_.field_type -eq "Text"} | ForEach-Object {
                $fieldValue = $_.value
                $mentionedArticles += $companyArticles | Where-Object { $(_Normalize-AssetName $_.name) -ieq $(_Normalize-AssetName $fieldValue) }
                $mentionedWebsites += $companywebsites | Where-Object { $(_Normalize-AssetName $_.value) -ieq $normalizedAssetName }
            }
    
            # "siblings": other assets with same normalized name but different id
            $siblings = @($companyAssetsByName[$normalizedAssetName] | Where-Object { $_.id -ne $a.id })
            $siblings | ForEach-Object {write-host "Sibling Asset $($a.name)@($($a.asset_layout_id)) -> $($_.name)@($($_.asset_layout_id))"; New-HuduRelation -FromableType "Asset" -ToableType "Asset" -FromableID $a.id -ToableID $_.id}
            $mentionedWebsites | ForEach-Object {Write-Host "[$($c.name)] '$($a.name)' ($($a.id)) mentions website -> '$($_.name)' ($($_.id))"; New-HuduRelation -FromableType "Asset" -ToableType "Website" -FromableID $a.id -ToableID $_.id}
            $mentionedArticles | ForEach-Object {Write-Host "[$($c.name)] '$($a.name)' ($($a.id)) mentions article -> '$($_.name)' ($($_.id))"; New-HuduRelation -FromableType "Asset" -ToableType "Article" -FromableID $a.id -ToableID $_.id}
            $mentionedAssets | ForEach-Object {Write-Host "[$($c.name)] '$($a.name)' ($($a.id)) mentions asset -> '$($_.name)' ($($_.id))"; New-HuduRelation -FromableType "Asset" -ToableType "Asset" -FromableID $a.id -ToableID $_.id}        
            write-host @"
Siblings: $($siblings.count)
Websites Mentioned: $($mentionedWebsites.count)
Articles Mentioned: $($mentionedArticles.count)
Assets Mentioned: $($mentionedAssets.count)
"@
        }
    }
    if (get-command -name Set-HapiErrorsDirectory -ErrorAction SilentlyContinue){try {Set-HapiErrorsDirectory -skipRetry $false} catch {}}

}

function Rename-HuduLayoutFieldsBulk {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [hashtable]$LabelMappings,
        [Parameter(ValueFromPipeline)]
        $Layouts
    )

    begin {
        $map = [System.Collections.Generic.Dictionary[string,string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($k in $LabelMappings.Keys) { $map[$k] = [string]$LabelMappings[$k] }
        $totalRenamed = 0
        $totalLayoutsTouched = 0
    }

    process {
        foreach ($layoutIn in @($Layouts)) {
            $layout = $layoutIn.asset_layout ?? $layoutIn
            if (-not $layout -or -not $layout.fields) { continue }

            $changed = $false
            $renamedHere = 0

            foreach ($f in $layout.fields) {
                $old = [string]($f.label ?? '')
                if ([string]::IsNullOrWhiteSpace($old)) { continue }
                $has = $false
                $new = $null

                $has = $map.TryGetValue($old, [ref]$new)

                if (-not $has) { continue }  # empty field label (shouldnt happen)
                if ([string]::IsNullOrWhiteSpace($new)) { continue }  # empty user-mapping
                if ($old -ceq $new) { continue } # already correct (case-sensitive compare)

                $f.label = $new
                $changed = $true
                $renamedHere++
            }

            if ($true -eq $changed) {
                if ($PSCmdlet.ShouldProcess("LayoutId=$($layout.id) '$($layout.name)'", "Rename $renamedHere field label(s)")) {
                    $fieldsPayload = @($layout.fields | ForEach-Object { ConvertTo-HashtableDeep -InputObject $_ })
                    Set-HuduAssetLayout -Id $layout.id -Fields $fieldsPayload
                }

                Write-Host ("[{0}] {1} field(s) renamed" -f $layout.name, $renamedHere) -ForegroundColor Green
                $totalRenamed += $renamedHere
                $totalLayoutsTouched++
            }
        }
    }

    end {
        Write-Host ("Done. Renamed {0} field(s) across {1} layout(s)." -f $totalRenamed, $totalLayoutsTouched) -ForegroundColor Cyan
    }
}