$project_workdir = $PSScriptRoot
if ($MyInvocation.InvocationName -eq '.') {
    Write-Host "Script was dot-sourced" -ForegroundColor Green
} else {
    Write-Host "Script was executed without dot-sourcing, this is the recommended method of running the script to ensure settings are retained in the session" -ForegroundColor Yellow; write-warning "exiting to prevent issues later on, please dot-source the script by running `. .\yourenvironmentfile.ps1` from powershell 7.5 or later (ideally as Administrator)" -ForegroundColor Red;
    exit 1
}
. "$project_workdir\fields-config.ps1"

$debugDir = Join-Path -path $project_workdir -childpath "Debug"
$ITPortalMigrationStarted = $ITPortalMigrationStarted ?? (Get-Date)

$jobs = @(
"Read-Data",
"Assets-and-Layouts",
"Submit-Passwords",
"Fetch-Docs",
"Create-Articles-FromRecords",
"Create-Articles-FromFiles",
"Set-Relations",
"Wrap-Up"
)

$exportLocation = $exportLocation ?? (Read-Host "please enter the full path to your export.")
$hudubaseUrl    = $hudubaseUrl    ?? (Read-Host "please enter the hudubase url.")
$huduapikey     = $huduapikey     ?? (Read-Host "please enter your hudu api key.")

if (-not (Test-Path -Path $exportLocation)) {
    Write-Error "The specified path does not exist: $exportLocation"; exit;
} else {
    write-host "Using export location: $exportLocation"
}


$ITPortalData =  @{}
$MigrationErrors = @()

foreach ($f in $(Get-ChildItem "$project_workdir\helpers" -Filter *.ps1)) {. $f.FullName}
 Get-PSVersionCompatible; Set-HuduModuleInitialized;
 Get-EnsuredPath $debugDir | Out-Null
 $internalCompany = Get-OrSetInternalCompany -internalCompanyName $internalCompanyName

 foreach ($job in $jobs){
     write-host "starting $job" -foregroundColor cyan
     . "$project_workdir\jobs\$job.ps1"
     write-host "finished $job" -foregroundColor darkcyan
 }

