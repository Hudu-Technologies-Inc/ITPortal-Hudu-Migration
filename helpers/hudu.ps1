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
