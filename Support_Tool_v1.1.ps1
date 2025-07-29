
# --- Imports & Styling ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global styling
$bgColor  = [System.Drawing.Color]::FromArgb(30,30,30)
$btnColor = [System.Drawing.Color]::FromArgb(45,45,48)
$font     = New-Object System.Drawing.Font("Segoe UI",10)

$Global:LogBasePath = "$env:LOCALAPPDATA\ITTools\MultiServiceTool"
$Global:MgConnected = $false
$Global:TenantUPN   = $null

function Style-Form {
    param($form)
    $form.BackColor = $bgColor
    $form.ForeColor = [System.Drawing.Color]::White
    $form.Font = $font
}

function Style-Button {
    param($btn)
    $btn.FlatStyle = 'Flat'
    $btn.BackColor = $btnColor
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.FlatAppearance.BorderSize = 0
    $btn.Font = $font
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn.Height = 40
}

function Add-BackButton {
    param($form)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = 'Main Screen'
    $btn.Size = New-Object System.Drawing.Size(120, 30)

    # Safe subtraction and fallback default
    $formWidth = if ($form -and $form.ClientSize.Width) { [int]$form.ClientSize.Width } else { 500 }
    $x = $formWidth - 140

    $btn.Location = New-Object System.Drawing.Point($x, 10)
    Style-Button $btn
    $btn.Add_Click({ $form.Close() })
    $form.Controls.Add($btn)
}

function Prompt-InputBox {
    param($Title, $Label)
    $form = New-Object System.Windows.Forms.Form
    Style-Form $form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(400,180)
    $form.StartPosition = 'CenterScreen'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Label
    $lbl.ForeColor = [System.Drawing.Color]::White
    $lbl.Location = New-Object System.Drawing.Point(20,20)
    $lbl.AutoSize = $true
    $form.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Size = New-Object System.Drawing.Size(340,25)
    $txt.Location = New-Object System.Drawing.Point(20,50)
    $form.Controls.Add($txt)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = 'OK'
    $btnOK.Location = New-Object System.Drawing.Point(280,100)
    Style-Button $btnOK
    $btnOK.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txt.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Input cannot be empty.","Input Error")
            return
        }
        $form.Tag = $txt.Text
        $form.DialogResult = 'OK'
        $form.Close()
    })
    $form.Controls.Add($btnOK)
    $form.AcceptButton = $btnOK

    if ($form.ShowDialog() -eq 'OK') { return $form.Tag }
    return $null
}

# --- Logging & Helpers ---
function Start-Logging {
    param($Module)
    $date = Get-Date -Format yyyy-MM-dd
    $folder = Join-Path $Global:LogBasePath ("logs\$date")
    if (-not (Test-Path $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    $time = Get-Date -Format HHmmss
    $script:LogFile = Join-Path $folder ("$Module-$time.log")
}

function Write-Log {
    param($Message)
    if (-not $script:LogFile) { Start-Logging -Module "General" }
    $timestamp = Get-Date -Format HH:mm:ss
    "$timestamp - $Message" | Out-File -FilePath $script:LogFile -Append
}

# --- AD Functions ---
# ------------------------
# Active Directory Functions
# ------------------------

function Invoke-ADLookup {
    Start-Logging -Module "AD_Lookup"
    $sam = Prompt-InputBox 'User Lookup' 'Enter sAMAccountName:'
    if (-not $sam) { return }
    try {
        $user = Get-ADUser -Identity $sam -Properties DisplayName, Mail, Enabled
        $msg = "Name: $($user.DisplayName)`nEmail: $($user.Mail)`nEnabled: $($user.Enabled)"
        Write-Log "AD Lookup for $sam successful"
        [System.Windows.Forms.MessageBox]::Show($msg, 'Lookup Result')
    } catch {
        Write-Log "AD Lookup Error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Lookup Failed')
    }
}

function Invoke-ADUnlock {
    Start-Logging -Module "AD_Unlock"
    $sam = Prompt-InputBox 'Unlock Account' 'Enter sAMAccountName:'
    if (-not $sam) { return }
    try {
        Unlock-ADAccount -Identity $sam
        Write-Log "$sam unlocked"
        [System.Windows.Forms.MessageBox]::Show("Account $sam unlocked", 'Success')
    } catch {
        Write-Log "Unlock Error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Failed')
    }
}

function Invoke-ADCreate {
    Start-Logging -Module "AD_Create"
    $sam = Prompt-InputBox 'Create User' 'Enter sAMAccountName:'
    $name = Prompt-InputBox 'Create User' 'Enter Display Name:'
    if (-not $sam -or -not $name) { return }
    try {
        New-ADUser -Name $name -SamAccountName $sam -AccountPassword (ConvertTo-SecureString 'P@ssw0rd123!' -AsPlainText -Force) -Enabled $true
        Write-Log "$name ($sam) created"
        [System.Windows.Forms.MessageBox]::Show("User created: $name", 'Success')
    } catch {
        Write-Log "Create Error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Failed')
    }
}

function Invoke-ADOffboard {
    Start-Logging -Module "AD_Offboard"
    $sam = Prompt-InputBox 'Offboard User' 'Enter sAMAccountName:'
    if (-not $sam) { return }
    try {
        Disable-ADAccount -Identity $sam
        Write-Log "$sam disabled"
        [System.Windows.Forms.MessageBox]::Show("Account disabled: $sam", 'Success')
    } catch {
        Write-Log "Offboarding Error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Failed')
    }
}

function Invoke-ADAddGroup {
    Start-Logging -Module "AD_AddGroup"
    $sam = Prompt-InputBox 'Add to Group' 'Enter sAMAccountName:'
    $group = Prompt-InputBox 'Add to Group' 'Enter Group Name:'
    if (-not $sam -or -not $group) { return }
    try {
        Add-ADGroupMember -Identity $group -Members $sam
        Write-Log "$sam added to $group"
        [System.Windows.Forms.MessageBox]::Show("Added to group", 'Success')
    } catch {
        Write-Log "Add Group Error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Failed')
    }
}

function Invoke-ADRemoveGroup {
    Start-Logging -Module "AD_RemoveGroup"
    $sam = Prompt-InputBox 'Remove from Group' 'Enter sAMAccountName:'
    $group = Prompt-InputBox 'Remove from Group' 'Enter Group Name:'
    if (-not $sam -or -not $group) { return }
    try {
        Remove-ADGroupMember -Identity $group -Members $sam -Confirm:$false
        Write-Log "$sam removed from $group"
        [System.Windows.Forms.MessageBox]::Show("Removed from group", 'Success')
    } catch {
        Write-Log "Remove Group Error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Failed')
    }
}

function Invoke-ADGroupsReport {
    Start-Logging -Module "AD_GroupsReport"
    $sam = Prompt-InputBox 'Groups Report' 'Enter sAMAccountName:'
    if (-not $sam) { return }
    try {
        $groups = Get-ADUser -Identity $sam -Properties MemberOf | Select-Object -ExpandProperty MemberOf
        $cleaned = $groups | ForEach-Object { ($_ -split ',')[0] -replace '^CN=' }
        $msg = $cleaned -join "`n"
        Write-Log "$sam groups retrieved"
        [System.Windows.Forms.MessageBox]::Show($msg, "$sam Group Memberships")
    } catch {
        Write-Log "Groups Report Error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Failed')
    }
}

# --- M365 Functions ---

# ------------------------
# Microsoft 365 License Mapping
# ------------------------

$Global:SkuMap = @{
    'ENTERPRISEPACK'             = 'Office 365 E3'
    'E5'                         = 'Office 365 E5'
    'BUSINESS_PREMIUM'           = 'Microsoft 365 Business Premium'
    'BUSINESS_BASIC'             = 'Microsoft 365 Business Basic'
    'M365_BUSINESS_STANDARD'     = 'Microsoft 365 Business Standard'
    'M365_F1'                    = 'Microsoft 365 F1'
    'EMS'                        = 'Enterprise Mobility + Security'
    'AAD_PREMIUM'                = 'Azure AD Premium P1'
    'AAD_PREMIUM_P2'             = 'Azure AD Premium P2'
    'PROJECTPROFESSIONAL'        = 'Project Online Professional'
    'VISIOONLINE_PLAN2'          = 'Visio Plan 2'
    'POWER_BI_PRO'               = 'Power BI Pro'
    'SPE_E3'                     = 'Microsoft 365 E3'
    'IDENTITY_THREAT_PROTECTION' = 'Microsoft 365 E5'
}

# ------------------------
# Microsoft Graph Auth
# ------------------------

$Global:MgConnected = $false

function Ensure-MgConnection {
    if (-not $Global:MgConnected) {
        try {
            Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All", "Directory.ReadWrite.All"
            $Global:MgConnected = $true
        } catch {
            $Global:MgConnected = $false
            [System.Windows.Forms.MessageBox]::Show("Authentication failed. Please try again.", "M365 Login Error")
            return $false
        }
    }
    return $true
}

# ------------------------
# M365 Module Functions
# ------------------------

function Invoke-M365Lookup {
    if (-not (Ensure-MgConnection)) { return }
    Start-Logging -Module "M365_Lookup"
    $upn = Prompt-InputBox 'M365 Lookup' 'Enter UPN:'
    if (-not $upn) { return }

    try {
        $user = Get-MgUser -UserId $upn -Property DisplayName, AssignedLicenses
        $licenses = @()
        $skuList = Get-MgSubscribedSku

        foreach ($lic in $user.AssignedLicenses) {
            $sku = $skuList | Where-Object { $_.SkuId -eq $lic.SkuId }
            if ($sku) {
                $skuPart = $sku.SkuPartNumber.Trim().ToUpper()
                if ($Global:SkuMap.ContainsKey($skuPart)) {
                    $licenses += $Global:SkuMap[$skuPart]
                } else {
                    Write-Log "Unknown SKU Part: $skuPart"
                    $licenses += $skuPart
                }
            } else {
                Write-Log "SKU ID not found: $($lic.SkuId)"
                $licenses += "Unknown SKU ID: $($lic.SkuId)"
            }
        }

        if ($licenses.Count -eq 0) {
            $licenses = @('None assigned')
        }

        $msg = "Display Name: $($user.DisplayName)`nLicenses:`n" + ($licenses -join "`n")
        Write-Log "M365 Lookup success for $upn"
        [System.Windows.Forms.MessageBox]::Show($msg, 'License Info')

    } catch {
        Write-Log "M365 Lookup error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error retrieving user license info.`n$($_.Exception.Message)", 'Lookup Failed')
    }
}

function Invoke-M365AssignLicense {
    if (-not (Ensure-MgConnection)) { return }
    Start-Logging -Module "M365_AssignLicense"
    $upn = Prompt-InputBox 'Assign License' 'Enter UPN:'
    if (-not $upn) { return }
    $skuId = Prompt-InputBox 'Assign License' 'Enter SKU ID (GUID):'
    if (-not $skuId) { return }

    try {
        Set-MgUserLicense -UserId $upn -AddLicenses @{SkuId = $skuId} -RemoveLicenses @()
        Write-Log "Assigned license $skuId to $upn"
        [System.Windows.Forms.MessageBox]::Show("License assigned to $upn", 'Success')
    } catch {
        Write-Log "Assign License error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Failed')
    }
}

function Invoke-M365RemoveLicense {
    if (-not (Ensure-MgConnection)) { return }
    Start-Logging -Module "M365_RemoveLicense"
    $upn = Prompt-InputBox 'Remove License' 'Enter UPN:'
    if (-not $upn) { return }
    $skuId = Prompt-InputBox 'Remove License' 'Enter SKU ID (GUID):'
    if (-not $skuId) { return }

    try {
        Set-MgUserLicense -UserId $upn -AddLicenses @() -RemoveLicenses @($skuId)
        Write-Log "Removed license $skuId from $upn"
        [System.Windows.Forms.MessageBox]::Show("License removed from $upn", 'Success')
    } catch {
        Write-Log "Remove License error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Failed')
    }
}

function Invoke-M365EnableUser {
    if (-not (Ensure-MgConnection)) { return }
    Start-Logging -Module "M365_EnableUser"
    $upn = Prompt-InputBox 'Enable User' 'Enter UPN:'
    if (-not $upn) { return }

    try {
        Update-MgUser -UserId $upn -AccountEnabled $true
        Write-Log "$upn enabled"
        [System.Windows.Forms.MessageBox]::Show("$upn enabled", 'Success')
    } catch {
        Write-Log "Enable User error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Failed')
    }
}

function Invoke-M365DisableUser {
    if (-not (Ensure-MgConnection)) { return }
    Start-Logging -Module "M365_DisableUser"
    $upn = Prompt-InputBox 'Disable User' 'Enter UPN:'
    if (-not $upn) { return }

    try {
        Update-MgUser -UserId $upn -AccountEnabled $false
        Write-Log "$upn disabled"
        [System.Windows.Forms.MessageBox]::Show("$upn disabled", 'Success')
    } catch {
        Write-Log "Disable User error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Failed')
    }
}
function Invoke-M365GroupMembership {
    if (-not (Ensure-MgConnection)) { return }
    Start-Logging -Module "M365_GroupMembership"
    $upn = Prompt-InputBox 'Group Membership' 'Enter UPN:'
    if (-not $upn) { return }

    try {
        $user = Get-MgUser -UserId $upn
        $groups = Get-MgUserMemberOf -UserId $user.Id
        $groupList = $groups | ForEach-Object {
            if ($_.AdditionalProperties["displayName"]) {
                $_.AdditionalProperties["displayName"]
            }
        }

        if ($groupList.Count -eq 0) { $groupList = "None" }

        $msg = "Groups for ${upn}:`n" + ($groupList -join "`n")
        Write-Log "Group membership retrieved for $upn"
        [System.Windows.Forms.MessageBox]::Show($msg, 'Group Membership')

    } catch {
        Write-Log "Group membership error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Failed')
    }
}

function Invoke-M365AddToGroup {
    if (-not (Ensure-MgConnection)) { return }
    Start-Logging -Module "M365_AddToGroup"
    $upn = Prompt-InputBox 'Add to Group' 'Enter UPN:'
    $groupId = Prompt-InputBox 'Add to Group' 'Enter Group ID:'
    if (-not $upn -or -not $groupId) { return }

    try {
        $user = Get-MgUser -UserId $upn
        New-MgGroupMember -GroupId $groupId -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)" }
        Write-Log "Added $upn to group $groupId"
        [System.Windows.Forms.MessageBox]::Show("Added $upn to group", 'Success')
    } catch {
        Write-Log "Add to group error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Failed')
    }
}

function Invoke-M365RemoveFromGroup {
    if (-not (Ensure-MgConnection)) { return }
    Start-Logging -Module "M365_RemoveFromGroup"
    $upn = Prompt-InputBox 'Remove from Group' 'Enter UPN:'
    $groupId = Prompt-InputBox 'Remove from Group' 'Enter Group ID:'
    if (-not $upn -or -not $groupId) { return }

    try {
        $user = Get-MgUser -UserId $upn
        Remove-MgGroupMemberByRef -GroupId $groupId -DirectoryObjectId $user.Id -Confirm:$false
        Write-Log "Removed $upn from group $groupId"
        [System.Windows.Forms.MessageBox]::Show("Removed $upn from group", 'Success')
    } catch {
        Write-Log "Remove from group error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Failed')
    }
}
function Export-LogsToCSV {
    Start-Logging -Module "ExportLogs"
    $folder = Join-Path $Global:LogBasePath "logs"
    $output = Join-Path $Global:LogBasePath "LogExport.csv"
    $allLines = @()

    if (-not (Test-Path $folder)) {
        [System.Windows.Forms.MessageBox]::Show("No logs found.","Export Logs")
        return
    }

    Get-ChildItem -Path $folder -Recurse -Filter *.log | ForEach-Object {
        $lines = Get-Content $_.FullName | ForEach-Object {
            [PSCustomObject]@{
                Date     = $_.Substring(0,8)
                Time     = $_.Substring(0,8)
                Message  = $_.Substring(11)
                File     = $_.Name
            }
        }
        $allLines += $lines
    }

    $allLines | Export-Csv -Path $output -NoTypeInformation
    Write-Log "Exported logs to CSV at $output"
    [System.Windows.Forms.MessageBox]::Show("Logs exported to: `n$output", "Export Complete")
}

# --- SECTION 5: M365 GUI Button Handling ---
function Show-M365Screen {
    $form = New-Object System.Windows.Forms.Form
    Style-Form $form
    $form.Text = 'Microsoft 365 Tools'
    $form.Size = New-Object System.Drawing.Size(520, 700)
    $form.StartPosition = 'CenterScreen'

    # Only authenticate if not already connected
    if (-not (Ensure-MgConnection)) {
        $form.Close()
        return
    }

    Add-BackButton $form

    $modules = @(
        @{ Name='License Lookup';         Action={ Invoke-M365Lookup } },
        @{ Name='Assign License';         Action={ Invoke-M365AssignLicense } },
        @{ Name='Remove License';         Action={ Invoke-M365RemoveLicense } },
        @{ Name='Enable User';            Action={ Invoke-M365EnableUser } },
        @{ Name='Disable User';           Action={ Invoke-M365DisableUser } },
        @{ Name='Group Membership';       Action={ Invoke-M365GroupMembership } },
        @{ Name='Add to M365 Group';      Action={ Invoke-M365AddToGroup } },
        @{ Name='Remove from M365 Group'; Action={ Invoke-M365RemoveFromGroup } },
        @{ Name='Tenant Info';            Action={ Show-TenantInfo } },
        @{ Name='Export Logs to CSV';     Action={ Export-LogsToCSV } }
    )

    for ($i = 0; $i -lt $modules.Count; $i++) {
        $row = [math]::Floor($i / 2)
        $col = $i % 2
        $x = 40 + ($col * 220)
        $y = 60 + ($row * 55)

        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $modules[$i].Name
        $btn.Size = New-Object System.Drawing.Size(200,40)
        $btn.Location = New-Object System.Drawing.Point -ArgumentList $x, $y
        Style-Button $btn
        $btn.Add_Click($modules[$i].Action)
        $form.Controls.Add($btn)
    }

    [void]$form.ShowDialog()
}

# --- License Mapping & Tenant Info ---
function Show-TenantInfo {
    if (-not (Ensure-MgConnection)) { return }
    Start-Logging -Module "Tenant_Info"

    try {
        $org = Get-MgOrganization
        $displayName = $org.DisplayName
        $tenantId = $org.Id
        $verifiedDomains = ($org.VerifiedDomains | ForEach-Object { $_.Name }) -join ", "

        $msg = "Display Name: $displayName`nTenant ID: $tenantId`nDomains: $verifiedDomains"
        Write-Log "Tenant info retrieved: $displayName"
        [System.Windows.Forms.MessageBox]::Show($msg, 'Tenant Info')
    } catch {
        Write-Log "Tenant info error: $_"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Failed')
    }
}
$Global:SkuMap = @{
    "ENTERPRISEPACK"             = "Office 365 E3"
    "E5"                         = "Office 365 E5"
    "EMS"                        = "Enterprise Mobility + Security"
    "EMSPREMIUM"                 = "EMS E5"
    "AAD_PREMIUM"                = "Azure AD Premium P1"
    "AAD_PREMIUM_P2"             = "Azure AD Premium P2"
    "M365BASIC"                  = "Microsoft 365 Business Basic"
    "M365BUSINESS"              = "Microsoft 365 Business Standard"
    "M365E3"                     = "Microsoft 365 E3"
    "M365E5"                     = "Microsoft 365 E5"
    "M365F1"                     = "Microsoft 365 F1"
    "VISIOCLIENT"               = "Visio Online Plan"
    "PROJECTPROFESSIONAL"       = "Project Online Professional"
}
function Get-FriendlyLicenseName {
    param($License)
    $skuPartNumber = $License.SkuPartNumber
    if ($Global:SkuMap.ContainsKey($skuPartNumber)) {
        return $Global:SkuMap[$skuPartNumber]
    }
    return $skuPartNumber
}


# --- Main Menu & Launch ---
function Show-MainMenu {
    $form = New-Object System.Windows.Forms.Form
    Style-Form $form
    $form.Text = 'Support Tool - Main Menu'
    $form.Size = New-Object System.Drawing.Size(500, 350)
    $form.StartPosition = 'CenterScreen'

    $header = New-Object System.Windows.Forms.Label
    $header.Text = 'Select a Service'
    $header.Font = New-Object System.Drawing.Font('Segoe UI',16,[System.Drawing.FontStyle]::Bold)
    $header.ForeColor = [System.Drawing.Color]::White
    $header.AutoSize = $true
    $header.Location = New-Object System.Drawing.Point(160, 30)
    $form.Controls.Add($header)

    $btnAD = New-Object System.Windows.Forms.Button
    $btnAD.Text = 'Active Directory'
    $btnAD.Size = New-Object System.Drawing.Size(350,45)
    $btnAD.Location = New-Object System.Drawing.Point(75,100)
    Style-Button $btnAD
    $btnAD.Add_Click({ $form.Hide(); Show-ADScreen; $form.Show() })
    $form.Controls.Add($btnAD)

    $btnM365 = New-Object System.Windows.Forms.Button
    $btnM365.Text = 'Microsoft 365'
    $btnM365.Size = New-Object System.Drawing.Size(350,45)
    $btnM365.Location = New-Object System.Drawing.Point(75,160)
    Style-Button $btnM365
    $btnM365.Add_Click({
    if (Ensure-MgConnection) {
        $form.Hide()
        Show-M365Screen
        $form.Show()
    }
})
    $form.Controls.Add($btnM365)

    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Text = 'Exit'
    $btnExit.Size = New-Object System.Drawing.Size(350,45)
    $btnExit.Location = New-Object System.Drawing.Point(75,220)
    $btnExit.Add_Click({ $form.Close() })
    Style-Button $btnExit
    $form.Controls.Add($btnExit)

    [void]$form.ShowDialog()
}

# Initialize log directory if missing
if (-not (Test-Path $Global:LogBasePath)) {
    New-Item -Path $Global:LogBasePath -ItemType Directory -Force | Out-Null
}

# Launch the tool
Show-MainMenu
