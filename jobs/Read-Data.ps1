$ITPortalData = @{}

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
        Properties = (Get-CSVPropertiesSafe $contents)  # assuming this returns header names
    }
}
$huducompanies = get-huducompanies
$internalCompany = $huducompanies | Where-Object { $_.name -ieq $internalCompanyName } | Select-Object -First 1