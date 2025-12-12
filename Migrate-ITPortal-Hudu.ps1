$project_workdir = $PSScriptRoot
$jobs = @(
"read-datas",
"Assets-and-Layouts"
)

$exportLocation = $exportLocation ?? (Read-Host "please enter the full path to your export.")
$hudubaseUrl    = $hudubaseUrl    ?? (Read-Host "please enter the hudubase url.")
$huduapikey     = $huduapikey     ?? (Read-Host "please enter your hudu api key.")
$internalCompanyName = $internalCompanyName ?? (Read-Host "please enter the internal company name to use for assets without a company.")

if (-not (Test-Path -Path $exportLocation)) {
    Write-Error "The specified path does not exist: $exportLocation"; exit;
} else {
    write-host "Using export location: $exportLocation"
}


$ITPortalData =  @{}
$MigrationErrors = @()

foreach ($f in $(Get-ChildItem "$project_workdir\helpers" -Filter *.ps1)) {. $f.FullName}
Get-PSVersionCompatible; Get-HuduModule; Set-HuduInstance; Get-HuduVersionCompatible;

foreach ($job in $jobs){
    write-host "starting $job"
    try {
        . "$project_workdir\jobs\$job.ps1"
    } catch {
        write-error "error during $job $_"
        $migrationerrors += @{job=$job; error = $_;}
    }
    write-host "finished $job"
}
