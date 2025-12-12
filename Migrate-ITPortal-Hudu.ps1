$project_workdir = $PSScriptRoot
$jobs = @("Assets-and-Layouts")

$ITPortalData =  @{}
foreach ($f in $(Get-ChildItem "$project_workdir\helpers" -Filter *.ps1)) {. $f.FullName}
Get-PSVersionCompatible; Get-HuduModule; Set-HuduInstance; Get-HuduVersionCompatible;

foreach ($job in $jobs){
    write-host "starting $job"
    . "$project_workdir\jobs\$job.ps1"
}
