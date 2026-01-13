$MatchedPasswords = @{}
$internalCompany = Get-OrSetInternalCompany -internalCompanyName $internalCompanyName
$PassTypes = @("Passwords-FieldPasswords","Passwords-Devices","Passwords-LocalPasswords","Passwords-Accounts")

foreach ($passType in $PassTypes) {
    Write-Host "Processing passwords of type: $passType"

    foreach ($pass in $itportaldata.$passType.CsvData) {
        if ([string]::IsNullOrwhitespace($pass.Password)) {
            write-host "skipping empty password entry: $($pass.Name)" -ForegroundColor Yellow
            continue
        }
        $otpInput = $pass.'2fasecret'

        $matchedCompany = Get-HuduCompanies -Name $pass.company; $matchedCompany = $matchedCompany.company ?? $matchedCompany;
        if ($null -eq $matchedCompany -or $matchedCompany.id -lt 1) {$matchedCompany= $internalCompany}

        $NewPasswordBody = @{
            Name        = $pass.Name
            URL         = $pass.URL
            Password    = $pass.Password
            Username    = $pass.Username
            CompanyId   = $matchedCompany.id
        }
        $NewPasswordBody.Description = ($NewPasswordBody.Description ?? '')
        if ($NewPasswordBody.Description.Length -gt 2000) {
            $NewPasswordBody.Description = $NewPasswordBody.Description.Substring(0,2000)
        }        

        $matchedPassword = Get-HuduPasswords -CompanyId $matchedCompany.id -Name $NewPasswordBody.Name | Select-Object -First 1
        $matchedPassword = $matchedPassword.asset_password ?? $matchedPassword

        $matchedAsset = Get-HuduAssets -CompanyId $matchedCompany.id |
            Where-Object { $_.name -ilike "*$($pass.Name)*" } |
            Select-Object -First 1
        $matchedAsset = $matchedAsset.asset ?? $matchedAsset

        if ($null -ne $matchedAsset) {
            $NewPasswordBody.PasswordableID   = $matchedAsset.id
            $NewPasswordBody.PasswordableType = "Asset"
        }

        if (-not [string]::IsNullOrWhiteSpace($otpInput)) {
            $otpInfo = ConvertTo-ValidatedOtpSecret $otpInput

            if (-not $otpInfo.IsValid) {
                $raw = $otpInfo.Raw
                $rawShort = if ($raw.Length -gt 200) { $raw.Substring(0,200) + '…' } else { $raw }
                $note = "Unvalidated OTP ($($otpInfo.Source)/$($otpInfo.Reason)): $rawShort"

                $existing = ($NewPasswordBody.Description ?? '').Trim()
                $NewPasswordBody.Description = if ($existing) { "$existing`n$note" } else { $note }
            }
            else {
                $NewPasswordBody.OTPSecret = $otpInfo.Secret
            }
        }

        if ($null -ne $matchedPassword) {
            $NewPasswordBody.ID = $matchedPassword.id
            $Password = Set-HuduPassword @NewPasswordBody
            Write-Host "Updated password: $($NewPasswordBody.Name)" -ForegroundColor Green
        }
        else {
            $Password = New-HuduPassword @NewPasswordBody
            Write-Host "Created password: $($NewPasswordBody.Name)" -ForegroundColor Green
        }
        $password = $Password.asset_password ?? $Password
        $MatchedPasswords["$($passType)_$($pass.ID)"] = $password


    }
}


$MatchedPasswords | convertto-json -depth 99 | set-content -path $(join-path $debugDir -childpath "PasswordsCreated.json") -force