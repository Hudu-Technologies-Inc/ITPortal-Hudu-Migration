
foreach ($layout in @{
"Devices" = @("os","Make","Model")
"Accounts" = @("Account Type")
"ConfigItems"  = @("Configuration Type","Vendor")

}){
    $layout = $null; $field = $null;
    $layout = get-huduassetlayouts -name $layout.key | select-object -first 1; $layout = $layout.asset_layout ?? $layout;
    $fields = $layout.fields | where-object { $_.label -in $layout.value }
    foreach ($field in $fields){
        write-host " Processing listify for asset layout $($layout.name) field $($field.label)" -ForegroundColor Cyan
        $layout = get-huduassetlayouts -name $layout.key | select-object -first 1; $layout = $layout.asset_layout ?? $layout;
        Listify-TextField -assetlayout $layout -fieldlabel $field.label
    }
}

# rename fields to something more human-readable
write-host "Label Mappings:`n$($($labelmappings | convertto-Json -depth 99).ToString())`n-[label mappings can be customized in fields-config.ps1]" -ForegroundColor Green
read-host "Press Enter to rename fields to more human-readable labels. Please note, that you should only do this when you're sure you have moved all the assets and items you wanted. You can verify this now and hit ENTER to continue or CTL+C to stop and review the data before renaming fields."

$allhudulayouts = get-huduassetlayouts
foreach ($layout in $allhudulayouts){
    $layout = $layout.asset_layout ?? $layout
    foreach ($field in $layout.fields){
        if ($labelMappings.ContainsKey($field.label)){
            $newLabel = $labelMappings[$field.label]
            write-host "Renaming field $($field.label) to $newLabel on layout $($layout.name)" -ForegroundColor Green
            Rename-HuduLayoutField -LayoutID $layout.id -OldLabel $field.label -NewLabel $newLabel | out-null
        }
    }
}