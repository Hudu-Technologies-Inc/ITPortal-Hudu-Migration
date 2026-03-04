# one-time switch for non-idemnopotent actions that we only want to do once
# $listsApplied = $listsApplied ?? $false
# if ($false -eq $listsApplied){
#     write-host "listifying select-fields."
#     foreach ($layout in @{
#     "Devices" = @("os","Make","Model")
#     "Accounts" = @("Account Type")
#     "ConfigItems"  = @("Configuration Type","Vendor")
#     }){
#         $layout = $null; $field = $null;
#         $layout = get-huduassetlayouts -name $layout.key | select-object -first 1; $layout = $layout.asset_layout ?? $layout;
#         $fields = $layout.fields | where-object { $_.label -in $layout.value -and $_.field_type -ne "ListSelect" }
#         foreach ($field in $fields){
#             write-host " Processing listify for asset layout $($layout.name) field $($field.label)" -ForegroundColor Cyan
#             $layout = get-huduassetlayouts -name $layout.key | select-object -first 1; $layout = $layout.asset_layout ?? $layout;
#             Listify-TextField -assetlayout $layout -fieldlabel $field.label
#         }
#     }
#     write-host "Applying listify to applicable fields on asset layouts. This will convert text fields that contain comma-separated values into actual lists that can be easily filtered and reported on. This is a crucial step for ensuring that your data is structured properly in Hudu and can be utilized effectively." -ForegroundColor Green
#     $listsApplied = $true
# }
# one-time switch for long-running actions. though idemnopotent, we only need to do once
$omniRelateApplied = $omniRelateApplied ?? $false
if ($false -eq $omniRelateApplied){
    write-host "picking up and applying any lingering relationships that may not have been directly matched. This can take a while, please be patient."
    Omni-Relate
    $omniRelateApplied = $true
}

# rename fields to something more human-readable
write-host "Label Mappings:`n$($($labelmappings | convertto-Json -depth 99).ToString())`n-[label mappings can be customized in fields-config.ps1]" -ForegroundColor Green
read-host "Press Enter to rename fields to more human-readable labels. Please note, that you should only do this when you're sure you have moved all the assets and items you wanted. You can verify this now and hit ENTER to continue or CTL+C to stop if you have more assets to move still. user-mappings will be re-read after you press ENTER, so you can feel free to edit them now if you'd like."
$ready = $false;
while (-not $ready){
    try {
        . "$project_workdir\fields-config.ps1"
        $ready = $true
    } catch {
        write-host "could not reload fields-config.ps1, make sure it exists and is valid powershell. Error was: $_" -ForegroundColor Red
        read-host "Press Enter to try again after fixing the issue with fields-config.ps1. We'll pick up exactly where we left off."
    }
}
write-host "refreshed final user-defined field name changes now!"; start-sleep -Seconds 3;
Rename-HuduLayoutFieldsBulk -LabelMappings $labelMappings -layouts $(if ($null -ne $ITPortalMigrationStarted) {$(Get-HuduAssetLayouts -updatedafter $ITPortalMigrationStarted.AddMinutes(-45))} else {$(Get-HuduAssetLayouts)})