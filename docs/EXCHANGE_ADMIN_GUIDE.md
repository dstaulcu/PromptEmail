# On-Premises Exchange Admin Guide for PromptEmail Add-In

This guide provides instructions for on-premises Exchange administrators to deploy, manage, and remove the PromptEmail add-in using PowerShell cmdlets.

## Prerequisites

- Exchange Management Shell access
- Exchange Server 2016 or later (for New-App and Remove-App cmdlets)
- Administrative permissions for add-in management
- Access to the PromptEmail manifest URL
- Active Directory module (for group-based deployments)
- Cross-domain trust relationships configured (if deploying across domains)

## Cross-Domain Active Directory Support

The PromptEmail add-in deployment scripts support cross-domain Active Directory scenarios where users or groups may exist in trusted domains. This is particularly useful in enterprise environments with multiple domain controllers or forest trusts.

### Configuration Options

All AD-related deployment examples support three configuration modes:

1. **Local Domain** (default): Uses current domain context
2. **Trusted Domain with Domain Controller**: Specify `-Server` parameter for domain controller
3. **Trusted Domain with Alternate Credentials**: Use both `-Server` and `-Credential` parameters

### Cross-Domain Configuration Variables

Add these variables to any group-based deployment script:

```powershell
# Cross-domain configuration (uncomment and modify as needed)
# $domainController = "dc.trusted-domain.com"  # For groups/users in trusted domains
# $credential = Get-Credential -Message "Enter credentials for trusted domain"  # If different credentials needed
```

### Supported Scenarios

- **Single Domain**: Standard deployment within your primary domain
- **Multi-Domain Forest**: Deploy to groups in child domains using domain controller specification
- **Forest Trusts**: Deploy across forest boundaries with appropriate credentials
- **External Trusts**: Deploy to specific trusted domains with explicit authentication

### Domain Controller Discovery

```powershell
# Find available domain controllers for a trusted domain
Get-ADDomainController -DomainName "trusted-domain.com" -Discover

# Test domain controller connectivity
Test-NetConnection "dc.trusted-domain.com" -Port 389  # LDAP
Test-NetConnection "dc.trusted-domain.com" -Port 636  # LDAPS

# Verify trust relationships
Get-ADTrust -Filter * | Select-Object Name, Direction, TrustType
```

### Troubleshooting Cross-Domain Issues

Common issues and solutions:

1. **Domain Controller Unreachable**
   ```powershell
   Test-NetConnection $domainController -Port 389
   Resolve-DnsName $domainController
   ```

2. **Authentication Failures**
   - Verify credentials have appropriate permissions in target domain
   - Check if delegation is required for cross-domain operations
   - Ensure Kerberos is properly configured

3. **Trust Relationship Issues**
   ```powershell
   Get-ADTrust -Filter * | Format-Table Name, Direction, TrustType
   Test-ComputerSecureChannel -Verbose
   ```

4. **Permission Errors**
   - Ensure account has read permissions on target AD objects
   - Verify Exchange permissions for target user mailboxes
   - Check for delegation requirements in multi-domain scenarios

## Quick Reference

### Install Add-In for a User
```powershell
# Preview what will happen
New-App -OrganizationApp -Url "https://your-domain.com/path/to/manifest.xml" -DefaultStateForUser Enabled -UserList "user@domain.com" -WhatIf

# Execute after confirming
New-App -OrganizationApp -Url "https://your-domain.com/path/to/manifest.xml" -DefaultStateForUser Enabled -UserList "user@domain.com"
```

### Remove Add-In for a User
```powershell
# Preview removal
Remove-App -Identity "PromptEmail" -Mailbox "user@domain.com" -WhatIf

# Execute removal
Remove-App -Identity "PromptEmail" -Mailbox "user@domain.com"
```

### Install Add-In for All Users
```powershell
# DANGER: This affects ALL users - always preview first
New-App -OrganizationApp -Url "https://your-domain.com/path/to/manifest.xml" -DefaultStateForUser Enabled -WhatIf

# Only execute after careful consideration
$confirmation = Read-Host "This will install for ALL users in the organization. Type 'INSTALL-ALL' to confirm"
if ($confirmation -eq 'INSTALL-ALL') {
    New-App -OrganizationApp -Url "https://your-domain.com/path/to/manifest.xml" -DefaultStateForUser Enabled
} else {
    Write-Host "Installation cancelled for safety."
}
```

## Detailed Instructions

### 1. Installing the Add-In

#### For Individual Users
```powershell
# Install for a single user (with safety checks)
$targetUser = "john.doe@company.com"
$manifestUrl = "https://your-domain.com/path/to/manifest.xml"

# Verify user exists first
try {
    Get-Mailbox $targetUser -ErrorAction Stop | Out-Null
    Write-Host "✓ User $targetUser found"
} catch {
    Write-Error "User $targetUser not found. Please verify the email address."
    exit
}

# Preview the installation
Write-Host "Preview: Installing add-in for $targetUser"
New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Enabled -UserList $targetUser -WhatIf

# Confirm before proceeding
$proceed = Read-Host "Proceed with installation? (y/N)"
if ($proceed -eq 'y' -or $proceed -eq 'Y') {
    New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Enabled -UserList $targetUser
    Write-Host "✓ Add-in installed for $targetUser"
} else {
    Write-Host "Installation cancelled."
}

# Install for multiple specific users (with validation)
$targetUsers = @("user1@company.com", "user2@company.com", "user3@company.com")
$validUsers = @()

# Validate all users exist
foreach ($user in $targetUsers) {
    try {
        Get-Mailbox $user -ErrorAction Stop | Out-Null
        $validUsers += $user
        Write-Host "✓ User $user found"
    } catch {
        Write-Warning "✗ User $user not found - excluding from installation"
    }
}

if ($validUsers.Count -gt 0) {
    Write-Host "`nPreview: Installing add-in for $($validUsers.Count) users"
    New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Enabled -UserList $validUsers -WhatIf
    
    $proceed = Read-Host "`nProceed with installation for these $($validUsers.Count) users? (y/N)"
    if ($proceed -eq 'y' -or $proceed -eq 'Y') {
        New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Enabled -UserList $validUsers
        Write-Host "✓ Add-in installed for $($validUsers.Count) users"
    }
} else {
    Write-Error "No valid users found. Installation aborted."
}
```

#### For All Users in Organization
```powershell
# CRITICAL: Organization-wide installation with multiple safety checks
$manifestUrl = "https://your-domain.com/path/to/manifest.xml"

Write-Host "⚠️  WARNING: This will install the add-in for ALL users in your organization!" -ForegroundColor Red
Write-Host "This action affects every mailbox in your Exchange environment." -ForegroundColor Yellow

# Get total user count for impact assessment
$totalUsers = (Get-Mailbox -ResultSize Unlimited).Count
Write-Host "`nTotal users that will be affected: $totalUsers" -ForegroundColor Yellow

# Preview the installation
Write-Host "`nPreviewing organization-wide installation:"
New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Enabled -WhatIf

# Multiple confirmation steps
Write-Host "`n" + "="*60
Write-Host "ORGANIZATION-WIDE INSTALLATION CONFIRMATION" -ForegroundColor Red
Write-Host "="*60

$step1 = Read-Host "Step 1/3: Have you tested this add-in with a pilot group? (yes/no)"
if ($step1 -ne 'yes') {
    Write-Host "❌ Please test with a pilot group first. Installation cancelled." -ForegroundColor Red
    exit
}

$step2 = Read-Host "Step 2/3: Have you notified users about this new add-in? (yes/no)"
if ($step2 -ne 'yes') {
    Write-Host "❌ Please notify users first. Installation cancelled." -ForegroundColor Red
    exit
}

Write-Host "Step 3/3: Type 'INSTALL-FOR-ALL-USERS' to confirm organization-wide installation:"
$finalConfirm = Read-Host
if ($finalConfirm -eq 'INSTALL-FOR-ALL-USERS') {
    Write-Host "`n🚀 Installing add-in for all $totalUsers users..." -ForegroundColor Green
    New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Enabled
    Write-Host "✅ Organization-wide installation completed!" -ForegroundColor Green
} else {
    Write-Host "❌ Installation cancelled for safety." -ForegroundColor Red
}

# Alternative: Install organization-wide but disabled by default
Write-Host "`n--- ALTERNATIVE: Install but let users enable themselves ---"
$installDisabled = Read-Host "Install organization-wide but DISABLED by default? Users can enable themselves. (y/N)"
if ($installDisabled -eq 'y' -or $installDisabled -eq 'Y') {
    New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Disabled -WhatIf
    $confirm = Read-Host "Proceed with disabled-by-default installation? (y/N)"
    if ($confirm -eq 'y' -or $confirm -eq 'Y') {
        New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Disabled
        Write-Host "✅ Add-in installed organization-wide but disabled. Users can enable it themselves."
    }
}
```

### 2. Managing Existing Installations

#### Check Current Add-In Status
```powershell
# List all organization apps
Get-App -OrganizationApp

# Check specific user's add-ins
Get-App -Mailbox "user@company.com"

# Get detailed information about PromptEmail add-in
Get-App -Identity "PromptEmail" -OrganizationApp
```

#### Enable/Disable Add-In for Users
```powershell
# Enable for specific user (with validation)
$targetUser = "user@company.com"
$addInName = "PromptEmail"

# Check current status first
try {
    $currentStatus = Get-App -Identity $addInName -Mailbox $targetUser -ErrorAction Stop
    Write-Host "Current status for $targetUser`: $($currentStatus.Enabled)"
} catch {
    Write-Error "Add-in '$addInName' not found for user $targetUser or user doesn't exist"
    exit
}

# Preview the change
Write-Host "Preview: Enabling add-in for $targetUser"
Set-App -Identity $addInName -Mailbox $targetUser -Enabled $true -WhatIf

# Confirm and execute
$proceed = Read-Host "Enable add-in for $targetUser? (y/N)"
if ($proceed -eq 'y' -or $proceed -eq 'Y') {
    Set-App -Identity $addInName -Mailbox $targetUser -Enabled $true
    Write-Host "✅ Add-in enabled for $targetUser"
}

# Disable for specific user (with confirmation)
Write-Host "`n--- Disable Add-In ---"
$disable = Read-Host "Disable add-in for $targetUser? (y/N)"
if ($disable -eq 'y' -or $disable -eq 'Y') {
    Set-App -Identity $addInName -Mailbox $targetUser -Enabled $false -WhatIf
    $confirmDisable = Read-Host "Confirm disable? (y/N)"
    if ($confirmDisable -eq 'y' -or $confirmDisable -eq 'Y') {
        Set-App -Identity $addInName -Mailbox $targetUser -Enabled $false
        Write-Host "✅ Add-in disabled for $targetUser"
    }
}

# Enable for multiple users (with validation and progress)
$targetUsers = @("user1@company.com", "user2@company.com")
$addInName = "PromptEmail"

Write-Host "Enabling add-in for multiple users..."
Write-Host "Target users: $($targetUsers -join ', ')"

# Validate all users first
$validUsers = @()
foreach ($user in $targetUsers) {
    try {
        Get-App -Identity $addInName -Mailbox $user -ErrorAction Stop | Out-Null
        $validUsers += $user
        Write-Host "✓ $user - add-in found"
    } catch {
        Write-Warning "✗ $user - add-in not found or user doesn't exist"
    }
}

if ($validUsers.Count -gt 0) {
    # Preview changes
    Write-Host "`nPreview: Enabling add-in for $($validUsers.Count) users"
    $validUsers | ForEach-Object { 
        Write-Host "  - $_"
        Set-App -Identity $addInName -Mailbox $_ -Enabled $true -WhatIf
    }
    
    $proceed = Read-Host "`nEnable add-in for these $($validUsers.Count) users? (y/N)"
    if ($proceed -eq 'y' -or $proceed -eq 'Y') {
        $validUsers | ForEach-Object { 
            Set-App -Identity $addInName -Mailbox $_ -Enabled $true
            Write-Host "✅ Enabled for $_"
        }
        Write-Host "✅ Add-in enabled for $($validUsers.Count) users"
    }
} else {
    Write-Error "No valid users found for add-in enablement"
}
```

### 3. Removing the Add-In

#### Remove for Individual Users
```powershell
# Remove from specific user (with safety checks)
$targetUser = "user@company.com"
$addInName = "PromptEmail"

# Check if add-in exists for user first
try {
    $addInInfo = Get-App -Identity $addInName -Mailbox $targetUser -ErrorAction Stop
    Write-Host "✓ Add-in '$addInName' found for user $targetUser"
    Write-Host "  Display Name: $($addInInfo.DisplayName)"
    Write-Host "  Enabled: $($addInInfo.Enabled)"
} catch {
    Write-Warning "Add-in '$addInName' not found for user $targetUser"
    exit
}

# Preview removal
Write-Host "`nPreview: Removing add-in from $targetUser"
Remove-App -Identity $addInName -Mailbox $targetUser -WhatIf

# Confirm removal
Write-Host "`n⚠️  This will permanently remove the add-in from $targetUser" -ForegroundColor Yellow
$proceed = Read-Host "Proceed with removal? (y/N)"
if ($proceed -eq 'y' -or $proceed -eq 'Y') {
    Remove-App -Identity $addInName -Mailbox $targetUser
    Write-Host "✅ Add-in removed from $targetUser"
} else {
    Write-Host "Removal cancelled."
}

# Remove from multiple users (with validation and progress)
$targetUsers = @("user1@company.com", "user2@company.com")
$addInName = "PromptEmail"

Write-Host "`nRemoving add-in from multiple users..."
Write-Host "Target users: $($targetUsers -join ', ')"

# Validate which users have the add-in
$usersWithAddIn = @()
foreach ($user in $targetUsers) {
    try {
        Get-App -Identity $addInName -Mailbox $user -ErrorAction Stop | Out-Null
        $usersWithAddIn += $user
        Write-Host "✓ $user - add-in found"
    } catch {
        Write-Warning "✗ $user - add-in not found"
    }
}

if ($usersWithAddIn.Count -gt 0) {
    # Preview removals
    Write-Host "`nPreview: Removing add-in from $($usersWithAddIn.Count) users"
    $usersWithAddIn | ForEach-Object { 
        Write-Host "  - $_"
        Remove-App -Identity $addInName -Mailbox $_ -WhatIf
    }
    
    Write-Host "`n⚠️  This will permanently remove the add-in from $($usersWithAddIn.Count) users" -ForegroundColor Yellow
    $proceed = Read-Host "Proceed with removal? (y/N)"
    if ($proceed -eq 'y' -or $proceed -eq 'Y') {
        $usersWithAddIn | ForEach-Object { 
            Remove-App -Identity $addInName -Mailbox $_
            Write-Host "✅ Removed from $_"
        }
        Write-Host "✅ Add-in removed from $($usersWithAddIn.Count) users"
    } else {
        Write-Host "Removal cancelled."
    }
} else {
    Write-Warning "No users found with the add-in installed"
}
```

#### Remove Organization-Wide
```powershell
# CRITICAL: Organization-wide removal with safety checks
$addInName = "PromptEmail"

Write-Host "⚠️  WARNING: This will remove the add-in from ALL users in your organization!" -ForegroundColor Red

# Check if the add-in exists organization-wide
try {
    $orgApp = Get-App -Identity $addInName -OrganizationApp -ErrorAction Stop
    Write-Host "✓ Organization app '$addInName' found"
    Write-Host "  Display Name: $($orgApp.DisplayName)"
    Write-Host "  Manifest URL: $($orgApp.ManifestUrl)"
    
    # Get count of affected users
    $affectedUsers = Get-App -Identity $addInName -OrganizationApp | Select-Object -ExpandProperty UserList
    if ($affectedUsers) {
        Write-Host "  Affected Users: $($affectedUsers.Count)" -ForegroundColor Yellow
    } else {
        Write-Host "  Affected: All organization users" -ForegroundColor Yellow
    }
} catch {
    Write-Warning "Organization app '$addInName' not found"
    exit
}

# Preview removal
Write-Host "`nPreview: Removing organization app"
Remove-App -Identity $addInName -OrganizationApp -WhatIf

# Multiple confirmation steps for organization-wide removal
Write-Host "`n" + "="*60 -ForegroundColor Red
Write-Host "ORGANIZATION-WIDE REMOVAL CONFIRMATION" -ForegroundColor Red
Write-Host "="*60 -ForegroundColor Red

$step1 = Read-Host "Step 1/3: Have you notified users about this removal? (yes/no)"
if ($step1 -ne 'yes') {
    Write-Host "❌ Please notify users first. Removal cancelled." -ForegroundColor Red
    exit
}

$step2 = Read-Host "Step 2/3: Do you have a backup of the current configuration? (yes/no)"
if ($step2 -ne 'yes') {
    Write-Host "❌ Please backup configuration first. Removal cancelled." -ForegroundColor Red
    Write-Host "Run this to backup: Get-App -Identity '$addInName' -OrganizationApp | Export-Clixml 'PromptEmail-Backup.xml'"
    exit
}

Write-Host "Step 3/3: Type 'REMOVE-FROM-ALL-USERS' to confirm organization-wide removal:"
$finalConfirm = Read-Host
if ($finalConfirm -eq 'REMOVE-FROM-ALL-USERS') {
    Write-Host "`n🗑️  Removing add-in from all users..." -ForegroundColor Yellow
    Remove-App -Identity $addInName -OrganizationApp -Confirm:$false
    Write-Host "✅ Organization-wide removal completed!" -ForegroundColor Green
} else {
    Write-Host "❌ Removal cancelled for safety." -ForegroundColor Red
}

# Alternative: Confirm with built-in confirmation prompt
Write-Host "`n--- ALTERNATIVE: Use built-in confirmation ---"
$useBuiltIn = Read-Host "Use Exchange's built-in confirmation prompt instead? (y/N)"
if ($useBuiltIn -eq 'y' -or $useBuiltIn -eq 'Y') {
    Remove-App -Identity $addInName -OrganizationApp
    # This will prompt: "Are you sure you want to perform this action?"
}
```

### 4. Bulk Operations

#### Install for Department/Group
```powershell
# Enhanced AD group installation with comprehensive safety checks and cross-domain support
$groupName = "Marketing Department"
$manifestUrl = "https://your-domain.com/path/to/manifest.xml"

# Cross-domain configuration (uncomment and modify as needed)
# $domainController = "dc.trusted-domain.com"  # For groups in trusted domains
# $credential = Get-Credential -Message "Enter credentials for trusted domain"  # If different credentials needed

Write-Host "Installing add-in for AD group: $groupName" -ForegroundColor Green
Write-Host "Manifest URL: $manifestUrl" -ForegroundColor Green

# Step 1: Validate AD group exists and get members (with cross-domain support)
try {
    Write-Host "`n📋 Step 1: Validating AD group..." -ForegroundColor Cyan
    
    # Standard domain (current domain)
    if (-not $domainController) {
        $adGroup = Get-ADGroup $groupName -ErrorAction Stop
        $groupMembers = Get-ADGroupMember $groupName -ErrorAction Stop
    } 
    # Cross-domain with domain controller
    elseif ($domainController -and -not $credential) {
        Write-Host "   Using domain controller: $domainController" -ForegroundColor Yellow
        $adGroup = Get-ADGroup $groupName -Server $domainController -ErrorAction Stop
        $groupMembers = Get-ADGroupMember $groupName -Server $domainController -ErrorAction Stop
    }
    # Cross-domain with credentials
    else {
        Write-Host "   Using domain controller: $domainController with alternate credentials" -ForegroundColor Yellow
        $adGroup = Get-ADGroup $groupName -Server $domainController -Credential $credential -ErrorAction Stop
        $groupMembers = Get-ADGroupMember $groupName -Server $domainController -Credential $credential -ErrorAction Stop
    }
    
    Write-Host "✅ AD group '$groupName' found" -ForegroundColor Green
    Write-Host "   Distinguished Name: $($adGroup.DistinguishedName)"
    Write-Host "   Domain: $($adGroup.DistinguishedName.Split(',') | Where-Object {$_ -like 'DC=*'} | Join-String -Separator '.')" -ForegroundColor Cyan
    Write-Host "✅ Group has $($groupMembers.Count) members" -ForegroundColor Green
    
} catch {
    Write-Error "❌ AD group '$groupName' not found or inaccessible"
    Write-Host "`nTroubleshooting cross-domain issues:" -ForegroundColor Yellow
    Write-Host "1. Verify domain controller is accessible: Test-NetConnection $domainController -Port 389" -ForegroundColor Cyan
    Write-Host "2. Check DNS resolution: Resolve-DnsName $domainController" -ForegroundColor Cyan
    Write-Host "3. Verify trust relationship: Get-ADTrust -Filter *" -ForegroundColor Cyan
    
    Write-Host "`nAvailable AD groups containing 'Marketing':" -ForegroundColor Yellow
    if ($domainController) {
        if ($credential) {
            Get-ADGroup -Filter "Name -like '*Marketing*'" -Server $domainController -Credential $credential | Select-Object Name, DistinguishedName | Format-Table
        } else {
            Get-ADGroup -Filter "Name -like '*Marketing*'" -Server $domainController | Select-Object Name, DistinguishedName | Format-Table
        }
    } else {
        Get-ADGroup -Filter "Name -like '*Marketing*'" | Select-Object Name, DistinguishedName | Format-Table
    }
    exit
}

# Step 2: Get user accounts and validate they have mailboxes (with cross-domain support)
Write-Host "`n📋 Step 2: Getting user accounts and validating mailboxes..." -ForegroundColor Cyan
$validUsers = @()
$invalidUsers = @()

foreach ($member in $groupMembers) {
    try {
        # Get full user object with cross-domain support
        if (-not $domainController) {
            $adUser = Get-ADUser $member.SamAccountName -Properties UserPrincipalName -ErrorAction Stop
        } elseif ($domainController -and -not $credential) {
            $adUser = Get-ADUser $member.SamAccountName -Properties UserPrincipalName -Server $domainController -ErrorAction Stop
        } else {
            $adUser = Get-ADUser $member.SamAccountName -Properties UserPrincipalName -Server $domainController -Credential $credential -ErrorAction Stop
        }
        
        $upn = $adUser.UserPrincipalName
        if ($upn) {
            try {
                # Validate mailbox exists in Exchange
                Get-Mailbox $upn -ErrorAction Stop | Out-Null
                $validUsers += $upn
                Write-Host "✅ $upn - mailbox found" -ForegroundColor Green
            } catch {
                $invalidUsers += $upn
                Write-Warning "⚠️  $upn - no mailbox found"
            }
        } else {
            Write-Warning "⚠️  $($member.Name) - no UserPrincipalName"
        }
    } catch {
        Write-Warning "⚠️  $($member.Name) - could not retrieve AD user details"
    }
}

# Step 3: Display summary and get confirmation
Write-Host "`n📋 Step 3: Installation Summary" -ForegroundColor Cyan
Write-Host "AD Group: $groupName" -ForegroundColor White
Write-Host "Total AD members: $($groupMembers.Count)" -ForegroundColor White
Write-Host "Valid mailbox users: $($validUsers.Count)" -ForegroundColor Green
if ($invalidUsers.Count -gt 0) {
    Write-Host "Users without mailboxes: $($invalidUsers.Count)" -ForegroundColor Yellow
    Write-Host "Users excluded: $($invalidUsers -join ', ')" -ForegroundColor Yellow
}

if ($validUsers.Count -eq 0) {
    Write-Error "❌ No valid users found with mailboxes. Installation aborted."
    exit
}

# Step 4: Preview installation
Write-Host "`n📋 Step 4: Preview installation" -ForegroundColor Cyan
Write-Host "The following $($validUsers.Count) users will have the add-in installed:" -ForegroundColor Yellow
$validUsers | ForEach-Object { Write-Host "  • $_" }

Write-Host "`nPreviewing installation..." -ForegroundColor Yellow
New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Enabled -UserList $validUsers -WhatIf

# Step 5: Final confirmation and execution
Write-Host "`n📋 Step 5: Final confirmation" -ForegroundColor Cyan
Write-Host "⚠️  This will install the PromptEmail add-in for $($validUsers.Count) users" -ForegroundColor Yellow

$proceed = Read-Host "Type 'INSTALL' to proceed with installation for AD group '$groupName'"
if ($proceed -eq 'INSTALL') {
    Write-Host "`n🚀 Installing add-in for $($validUsers.Count) users..." -ForegroundColor Green
    
    try {
        New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Enabled -UserList $validUsers
        Write-Host "✅ Add-in installation completed for AD group '$groupName'!" -ForegroundColor Green
        Write-Host "✅ Successfully installed for $($validUsers.Count) users" -ForegroundColor Green
        
        # Log the installation
        $logEntry = @{
            Timestamp = Get-Date
            Action = "Install"
            Group = $groupName
            UsersCount = $validUsers.Count
            Users = $validUsers
            ManifestUrl = $manifestUrl
        }
        $logEntry | ConvertTo-Json | Out-File "PromptEmail-Installation-Log.json" -Append
        Write-Host "📝 Installation logged to PromptEmail-Installation-Log.json" -ForegroundColor Cyan
        
    } catch {
        Write-Error "❌ Installation failed: $($_.Exception.Message)"
        exit
    }
} else {
    Write-Host "❌ Installation cancelled for safety." -ForegroundColor Red
}

# Step 6: Verification (optional)
$verify = Read-Host "`nVerify installation? This will check if the add-in is properly installed. (y/N)"
if ($verify -eq 'y' -or $verify -eq 'Y') {
    Write-Host "`n📋 Verifying installation..." -ForegroundColor Cyan
    $successCount = 0
    $failCount = 0
    
    foreach ($user in $validUsers) {
        try {
            $userApp = Get-App -Identity "PromptEmail" -Mailbox $user -ErrorAction Stop
            Write-Host "✅ $user - add-in installed and $($userApp.Enabled ? 'enabled' : 'disabled')" -ForegroundColor Green
            $successCount++
        } catch {
            Write-Warning "⚠️  $user - verification failed"
            $failCount++
        }
    }
    
    Write-Host "`n📊 Verification Results:" -ForegroundColor Cyan
    Write-Host "Successfully verified: $successCount/$($validUsers.Count)" -ForegroundColor Green
    if ($failCount -gt 0) {
        Write-Host "Verification failed: $failCount/$($validUsers.Count)" -ForegroundColor Yellow
    }
}
```

#### Remove for Department/Group
```powershell
# Enhanced AD group removal with comprehensive safety checks and cross-domain support
$groupName = "Marketing Department"
$addInName = "PromptEmail"

# Cross-domain configuration (uncomment and modify as needed)
# $domainController = "dc.trusted-domain.com"  # For groups in trusted domains
# $credential = Get-Credential -Message "Enter credentials for trusted domain"  # If different credentials needed

Write-Host "Removing add-in '$addInName' from AD group: $groupName" -ForegroundColor Red
Write-Host "⚠️  This will remove the add-in from all members of the group" -ForegroundColor Yellow

# Step 1: Validate AD group exists and get members (with cross-domain support)
try {
    Write-Host "`n📋 Step 1: Validating AD group..." -ForegroundColor Cyan
    
    # Standard domain (current domain)
    if (-not $domainController) {
        $adGroup = Get-ADGroup $groupName -ErrorAction Stop
        $groupMembers = Get-ADGroupMember $groupName -ErrorAction Stop
    } 
    # Cross-domain with domain controller
    elseif ($domainController -and -not $credential) {
        Write-Host "   Using domain controller: $domainController" -ForegroundColor Yellow
        $adGroup = Get-ADGroup $groupName -Server $domainController -ErrorAction Stop
        $groupMembers = Get-ADGroupMember $groupName -Server $domainController -ErrorAction Stop
    }
    # Cross-domain with credentials
    else {
        Write-Host "   Using domain controller: $domainController with alternate credentials" -ForegroundColor Yellow
        $adGroup = Get-ADGroup $groupName -Server $domainController -Credential $credential -ErrorAction Stop
        $groupMembers = Get-ADGroupMember $groupName -Server $domainController -Credential $credential -ErrorAction Stop
    }
    
    Write-Host "✅ AD group '$groupName' found" -ForegroundColor Green
    Write-Host "   Distinguished Name: $($adGroup.DistinguishedName)"
    Write-Host "   Domain: $($adGroup.DistinguishedName.Split(',') | Where-Object {$_ -like 'DC=*'} | Join-String -Separator '.')" -ForegroundColor Cyan
    Write-Host "✅ Group has $($groupMembers.Count) members" -ForegroundColor Green
    
} catch {
    Write-Error "❌ AD group '$groupName' not found or inaccessible"
    Write-Host "`nTroubleshooting cross-domain issues:" -ForegroundColor Yellow
    Write-Host "1. Verify domain controller is accessible: Test-NetConnection $domainController -Port 389" -ForegroundColor Cyan
    Write-Host "2. Check DNS resolution: Resolve-DnsName $domainController" -ForegroundColor Cyan
    Write-Host "3. Verify trust relationship: Get-ADTrust -Filter *" -ForegroundColor Cyan
    exit
}

# Step 2: Get user accounts with add-in currently installed (with cross-domain support)
Write-Host "`n📋 Step 2: Finding users with add-in installed..." -ForegroundColor Cyan
$usersWithAddIn = @()
$usersWithoutAddIn = @()

foreach ($member in $groupMembers) {
    try {
        # Get full user object with cross-domain support
        if (-not $domainController) {
            $adUser = Get-ADUser $member.SamAccountName -Properties UserPrincipalName -ErrorAction Stop
        } elseif ($domainController -and -not $credential) {
            $adUser = Get-ADUser $member.SamAccountName -Properties UserPrincipalName -Server $domainController -ErrorAction Stop
        } else {
            $adUser = Get-ADUser $member.SamAccountName -Properties UserPrincipalName -Server $domainController -Credential $credential -ErrorAction Stop
        }
        
        $upn = $adUser.UserPrincipalName
        if ($upn) {
            try {
                # Check if add-in is installed for this user
                $userApp = Get-App -Identity $addInName -Mailbox $upn -ErrorAction Stop
                $usersWithAddIn += $upn
                Write-Host "✅ $upn - add-in currently installed ($($userApp.Enabled ? 'enabled' : 'disabled'))" -ForegroundColor Yellow
            } catch {
                $usersWithoutAddIn += $upn
                Write-Host "ℹ️  $upn - add-in not installed" -ForegroundColor Gray
            }
        } else {
            Write-Warning "⚠️  $($member.Name) - no UserPrincipalName"
        }
    } catch {
        Write-Warning "⚠️  $($member.Name) - could not retrieve AD user details"
    }
}

# Step 3: Summary and confirmation
Write-Host "`n📊 Removal Summary:" -ForegroundColor Cyan
Write-Host "AD Group: $groupName" -ForegroundColor White
Write-Host "Total members: $($groupMembers.Count)" -ForegroundColor White
Write-Host "Users with add-in: $($usersWithAddIn.Count)" -ForegroundColor Yellow
Write-Host "Users without add-in: $($usersWithoutAddIn.Count)" -ForegroundColor Gray

if ($usersWithAddIn.Count -eq 0) {
    Write-Host "✅ No users in group '$groupName' have the add-in installed. Nothing to remove." -ForegroundColor Green
    exit
}

Write-Host "`n⚠️  USERS TO BE AFFECTED:" -ForegroundColor Red
$usersWithAddIn | ForEach-Object { Write-Host "   - $_" -ForegroundColor Red }

# Step 4: What-If preview
Write-Host "`n📋 Step 3: Previewing removal (What-If)..." -ForegroundColor Cyan
foreach ($user in $usersWithAddIn) {
    try {
        Remove-App -Identity $addInName -Mailbox $user -WhatIf
    } catch {
        Write-Warning "⚠️  Cannot preview removal for $user"
    }
}

# Step 5: Final confirmation with safety checks
Write-Host "`n⚠️  FINAL CONFIRMATION REQUIRED" -ForegroundColor Red
Write-Host "This will remove '$addInName' from $($usersWithAddIn.Count) users in group '$groupName'" -ForegroundColor Yellow
Write-Host "Domain: $($adGroup.DistinguishedName.Split(',') | Where-Object {$_ -like 'DC=*'} | Join-String -Separator '.')" -ForegroundColor Yellow

$firstConfirm = Read-Host "`nType 'REMOVE' to confirm removal (case-sensitive)"
if ($firstConfirm -ne 'REMOVE') {
    Write-Host "❌ Removal cancelled for safety." -ForegroundColor Green
    exit
}

$secondConfirm = Read-Host "Type the group name '$groupName' to confirm"
if ($secondConfirm -ne $groupName) {
    Write-Host "❌ Group name mismatch. Removal cancelled for safety." -ForegroundColor Green
    exit
}

# Step 6: Execute removal
Write-Host "`n📋 Step 4: Executing removal..." -ForegroundColor Cyan
$successCount = 0
$failCount = 0
$failedUsers = @()

foreach ($user in $usersWithAddIn) {
    try {
        Remove-App -Identity $addInName -Mailbox $user
        Write-Host "✅ $user - add-in removed successfully" -ForegroundColor Green
        $successCount++
    } catch {
        Write-Error "❌ $user - removal failed: $($_.Exception.Message)"
        $failedUsers += $user
        $failCount++
    }
}

# Step 7: Summary and logging
Write-Host "`n📊 Removal Results:" -ForegroundColor Cyan
Write-Host "Successfully removed: $successCount/$($usersWithAddIn.Count)" -ForegroundColor Green
if ($failCount -gt 0) {
    Write-Host "Failed removals: $failCount/$($usersWithAddIn.Count)" -ForegroundColor Red
    Write-Host "Failed users: $($failedUsers -join ', ')" -ForegroundColor Red
}

# Log the removal operation
$logEntry = @{
    Timestamp = Get-Date
    Action = "Group Removal"
    GroupName = $groupName
    GroupDN = $adGroup.DistinguishedName
    AddInName = $addInName
    TotalGroupMembers = $groupMembers.Count
    UsersWithAddIn = $usersWithAddIn.Count
    SuccessfulRemovals = $successCount
    FailedRemovals = $failCount
    FailedUsers = $failedUsers
    Domain = if ($domainController) { $domainController } else { "Current Domain" }
}
$logEntry | ConvertTo-Json | Out-File "PromptEmail-Removal-Log.json" -Append
Write-Host "📝 Removal logged to PromptEmail-Removal-Log.json" -ForegroundColor Cyan

Write-Host "`n✅ Group removal completed!" -ForegroundColor Green
```

## Troubleshooting

### Common Issues

#### Add-In Installed for All Users Instead of Target Group
**Problem**: Running group-based installation commands resulted in organization-wide deployment.

**Cause**: The Active Directory query returned no results, causing `$users` to be empty. When `-UserList` is empty or null, Exchange treats this as "install for all users".

**Solution**:
```powershell
# Always test your AD query first with cross-domain support
$groupName = "Marketing Department"

# Cross-domain configuration (uncomment and modify as needed)
# $domainController = "dc.trusted-domain.com"  # For groups in trusted domains
# $credential = Get-Credential -Message "Enter credentials for trusted domain"  # If different credentials needed

try {
    # Test AD group query with cross-domain support
    if (-not $domainController) {
        $testQuery = Get-ADGroupMember $groupName | Get-ADUser
    } elseif ($domainController -and -not $credential) {
        Write-Host "Using domain controller: $domainController" -ForegroundColor Yellow
        $testQuery = Get-ADGroupMember $groupName -Server $domainController | Get-ADUser -Server $domainController
    } else {
        Write-Host "Using domain controller: $domainController with alternate credentials" -ForegroundColor Yellow
        $testQuery = Get-ADGroupMember $groupName -Server $domainController -Credential $credential | Get-ADUser -Server $domainController -Credential $credential
    }
    
    if ($testQuery.Count -eq 0) {
        Write-Error "AD group query returned no results. Check group name and permissions."
        
        # List available groups to help with troubleshooting (with cross-domain support)
        Write-Host "Available AD groups containing 'Marketing':" -ForegroundColor Yellow
        if (-not $domainController) {
            Get-ADGroup -Filter "Name -like '*Marketing*'" | Select-Object Name, DistinguishedName
        } elseif ($domainController -and -not $credential) {
            Get-ADGroup -Filter "Name -like '*Marketing*'" -Server $domainController | Select-Object Name, DistinguishedName
        } else {
            Get-ADGroup -Filter "Name -like '*Marketing*'" -Server $domainController -Credential $credential | Select-Object Name, DistinguishedName
        }
        
        # Check if the group exists (with cross-domain support)
        try {
            if (-not $domainController) {
                Get-ADGroup $groupName
            } elseif ($domainController -and -not $credential) {
                Get-ADGroup $groupName -Server $domainController
            } else {
                Get-ADGroup $groupName -Server $domainController -Credential $credential
            }
            Write-Host "Group exists but has no members or insufficient permissions to read members." -ForegroundColor Yellow
        } catch {
            Write-Error "Group '$groupName' not found. Please verify the exact group name."
            if ($domainController) {
                Write-Host "`nTroubleshooting cross-domain issues:" -ForegroundColor Yellow
                Write-Host "1. Verify domain controller is accessible: Test-NetConnection $domainController -Port 389" -ForegroundColor Cyan
                Write-Host "2. Check DNS resolution: Resolve-DnsName $domainController" -ForegroundColor Cyan
                Write-Host "3. Verify trust relationship: Get-ADTrust -Filter *" -ForegroundColor Cyan
            }
        }
    } else {
        Write-Host "Query successful - found $($testQuery.Count) users" -ForegroundColor Green
        $users = $testQuery | Select-Object -ExpandProperty UserPrincipalName
    # Proceed with installation...
}
```

**Prevention**: Always validate your user list before running `New-App`:
```powershell
# Safe installation pattern
$users = Get-ADGroupMember "YourGroup" | Get-ADUser | Select-Object -ExpandProperty UserPrincipalName
if ($users -and $users.Count -gt 0) {
    Write-Host "Installing for $($users.Count) users: $($users -join ', ')"
    $proceed = Read-Host "Continue? (y/N)"
    if ($proceed -eq 'y') {
        New-App -OrganizationApp -Url "https://your-domain.com/manifest.xml" -DefaultStateForUser Enabled -UserList $users
    }
} else {
    Write-Error "No users found - installation aborted to prevent organization-wide deployment"
}
```

**Rollback if accidentally installed organization-wide**:
```powershell
# EMERGENCY ROLLBACK: Remove from all users and reinstall for intended group only
$addInName = "PromptEmail"
$manifestUrl = "https://your-domain.com/manifest.xml"
$intendedGroup = "Marketing Department"

Write-Host "🚨 EMERGENCY ROLLBACK PROCEDURE" -ForegroundColor Red
Write-Host "This will remove the add-in from ALL users and reinstall for intended group only" -ForegroundColor Yellow

# Step 1: Confirm the emergency rollback
$emergency = Read-Host "Is this an emergency rollback after accidental org-wide install? (yes/no)"
if ($emergency -ne 'yes') {
    Write-Host "❌ Emergency rollback cancelled." -ForegroundColor Red
    exit
}

# Step 2: Backup current state
Write-Host "`n📋 Step 1: Backing up current state..." -ForegroundColor Cyan
try {
    Get-App -Identity $addInName -OrganizationApp | Export-Clixml "Emergency-Backup-$(Get-Date -Format 'yyyyMMdd-HHmmss').xml"
    Write-Host "✅ Current state backed up" -ForegroundColor Green
} catch {
    Write-Warning "⚠️  Could not backup current state"
}

# Step 3: Remove from all users
Write-Host "`n📋 Step 2: Removing add-in from ALL users..." -ForegroundColor Cyan
Remove-App -Identity $addInName -OrganizationApp -WhatIf

$confirmRemove = Read-Host "Proceed with organization-wide removal? (yes/no)"
if ($confirmRemove -eq 'yes') {
    Remove-App -Identity $addInName -OrganizationApp -Confirm:$false
    Write-Host "✅ Add-in removed from all users" -ForegroundColor Green
    
    # Wait a moment for Exchange to process
    Write-Host "⏳ Waiting 10 seconds for Exchange to process removal..." -ForegroundColor Yellow
    Start-Sleep -Seconds 10
} else {
    Write-Host "❌ Rollback cancelled." -ForegroundColor Red
    exit
}

# Step 4: Reinstall for intended group only
Write-Host "`n📋 Step 3: Reinstalling for intended group '$intendedGroup'..." -ForegroundColor Cyan

# Cross-domain configuration (uncomment and modify as needed)
# $domainController = "dc.trusted-domain.com"  # For groups in trusted domains
# $credential = Get-Credential -Message "Enter credentials for trusted domain"  # If different credentials needed

# Get intended group members with validation and cross-domain support
try {
    if (-not $domainController) {
        $targetUsers = Get-ADGroupMember $intendedGroup | Get-ADUser | Select-Object -ExpandProperty UserPrincipalName
    } elseif ($domainController -and -not $credential) {
        Write-Host "   Using domain controller: $domainController" -ForegroundColor Yellow
        $targetUsers = Get-ADGroupMember $intendedGroup -Server $domainController | Get-ADUser -Server $domainController | Select-Object -ExpandProperty UserPrincipalName
    } else {
        Write-Host "   Using domain controller: $domainController with alternate credentials" -ForegroundColor Yellow
        $targetUsers = Get-ADGroupMember $intendedGroup -Server $domainController -Credential $credential | Get-ADUser -Server $domainController -Credential $credential | Select-Object -ExpandProperty UserPrincipalName
    }
    
    # Validate users have mailboxes
    $validUsers = @()
    foreach ($user in $targetUsers) {
        try {
            Get-Mailbox $user -ErrorAction Stop | Out-Null
            $validUsers += $user
        } catch {
            Write-Warning "⚠️  $user - no mailbox found, excluding"
        }
    }
    
    if ($validUsers -and $validUsers.Count -gt 0) {
        Write-Host "✅ Found $($validUsers.Count) valid users in group '$intendedGroup'" -ForegroundColor Green
        
        # Preview reinstallation
        New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Enabled -UserList $validUsers -WhatIf
        
        $confirmInstall = Read-Host "Reinstall for these $($validUsers.Count) users? (yes/no)"
        if ($confirmInstall -eq 'yes') {
            New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Enabled -UserList $validUsers
            Write-Host "✅ Add-in reinstalled for $($validUsers.Count) intended users only!" -ForegroundColor Green
            
            # Log the emergency rollback
            $rollbackLog = @{
                Timestamp = Get-Date
                Action = "Emergency Rollback"
                OriginalScope = "Organization-wide"
                NewScope = "Group: $intendedGroup"
                UsersCount = $validUsers.Count
                Users = $validUsers
            }
            $rollbackLog | ConvertTo-Json | Out-File "Emergency-Rollback-Log.json" -Append
            Write-Host "📝 Emergency rollback logged" -ForegroundColor Cyan
            
        } else {
            Write-Host "❌ Reinstallation cancelled." -ForegroundColor Red
        }
    } else {
        Write-Error "❌ No valid users found in group '$intendedGroup'"
    }
} catch {
    Write-Error "❌ Could not process group '$intendedGroup': $($_.Exception.Message)"
}

Write-Host "`n✅ Emergency rollback procedure completed!" -ForegroundColor Green
```

#### Add-In Not Appearing in Outlook
```powershell
# Check if add-in is properly installed
Get-App -Mailbox "user@company.com" | Where-Object {$_.DisplayName -like "*PromptEmail*"}

# Verify add-in is enabled
Get-App -Identity "PromptEmail" -Mailbox "user@company.com"
```

#### Manifest URL Issues
```powershell
# Test manifest URL accessibility
Invoke-WebRequest -Uri "https://your-domain.com/path/to/manifest.xml" -UseBasicParsing

# Install with verbose output for debugging
New-App -OrganizationApp -Url "https://your-domain.com/path/to/manifest.xml" -DefaultStateForUser Enabled -UserList "user@company.com" -Verbose
```

#### Permission Errors
Ensure you have the necessary Exchange roles:
- **Organization Management** (full admin)
- **App Marketplace** (for add-in management)
- **Mail Recipients** (for user-specific operations)

### Verification Steps

1. **Confirm Installation**:
   ```powershell
   Get-App -OrganizationApp | Where-Object {$_.DisplayName -eq "PromptEmail"}
   ```

2. **Check User Assignment**:
   ```powershell
   Get-App -Identity "PromptEmail" -Mailbox "user@company.com"
   ```

3. **Verify Manifest**:
   ```powershell
   (Get-App -Identity "PromptEmail" -OrganizationApp).ManifestUrl
   ```

## Security Considerations

- **Manifest Hosting**: Ensure the manifest XML is hosted on a secure, accessible server
- **HTTPS Required**: Exchange requires HTTPS for manifest URLs
- **Network Access**: Ensure Exchange servers can reach the manifest URL
- **User Permissions**: Consider which users should have access to AI-powered email features

## Advanced Scenarios

### Gradual Rollout
```powershell
# Safe gradual rollout with comprehensive validation and rollback capabilities
$manifestUrl = "https://your-domain.com/manifest.xml"
$addInName = "PromptEmail"

Write-Host "🚀 GRADUAL ROLLOUT PROCEDURE" -ForegroundColor Green
Write-Host "This implements a safe, phased approach to organization-wide deployment" -ForegroundColor Cyan

# Phase 1: Pilot group with comprehensive testing
Write-Host "`n📋 PHASE 1: Pilot Group" -ForegroundColor Yellow
$pilotGroup = "PromptEmail-Pilot"

# Validate pilot group and install
try {
    $pilotUsers = Get-ADGroupMember $pilotGroup | Get-ADUser | Select-Object -ExpandProperty UserPrincipalName
    
    if (-not $pilotUsers -or $pilotUsers.Count -eq 0) {
        Write-Error "❌ Pilot group '$pilotGroup' not found or empty. Create the group and add test users first."
        Write-Host "Recommended pilot group size: 5-10 users from different departments" -ForegroundColor Cyan
        exit
    }
    
    # Validate mailboxes exist
    $validPilotUsers = @()
    foreach ($user in $pilotUsers) {
        try {
            Get-Mailbox $user -ErrorAction Stop | Out-Null
            $validPilotUsers += $user
        } catch {
            Write-Warning "⚠️  Pilot user $user has no mailbox"
        }
    }
    
    Write-Host "Pilot group '$pilotGroup' has $($validPilotUsers.Count) valid users" -ForegroundColor Green
    Write-Host "Pilot users: $($validPilotUsers -join ', ')" -ForegroundColor White
    
    # Install for pilot
    New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Enabled -UserList $validPilotUsers -WhatIf
    $installPilot = Read-Host "`nInstall for pilot group? (y/N)"
    if ($installPilot -eq 'y') {
        New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Enabled -UserList $validPilotUsers
        Write-Host "✅ Phase 1 completed: Add-in installed for pilot group" -ForegroundColor Green
        
        Write-Host "`n⏳ Allow 1-2 weeks for pilot testing before proceeding to Phase 2" -ForegroundColor Yellow
        Write-Host "   - Gather feedback from pilot users" -ForegroundColor Cyan
        Write-Host "   - Monitor for any issues or concerns" -ForegroundColor Cyan
        Write-Host "   - Verify add-in functionality across different scenarios" -ForegroundColor Cyan
        
        $readyForPhase2 = Read-Host "`nIs pilot testing complete and successful? Proceed to Phase 2? (y/N)"
        if ($readyForPhase2 -ne 'y') {
            Write-Host "✅ Phase 1 completed. Stopping here for pilot evaluation." -ForegroundColor Green
            exit
        }
    } else {
        Write-Host "❌ Phase 1 cancelled." -ForegroundColor Red
        exit
    }
} catch {
    Write-Error "❌ Phase 1 failed: $($_.Exception.Message)"
    exit
}

# Phase 2: Department rollout
Write-Host "`n📋 PHASE 2: Department Rollout" -ForegroundColor Yellow
$targetDepartment = Read-Host "Enter department name for Phase 2 rollout (e.g., 'Marketing')"

if ($targetDepartment) {
    try {
        $deptUsers = Get-ADUser -Filter "Department -eq '$targetDepartment'" | Select-Object -ExpandProperty UserPrincipalName
        
        if (-not $deptUsers -or $deptUsers.Count -eq 0) {
            Write-Warning "⚠️  No users found in department '$targetDepartment'"
            Write-Host "Available departments:" -ForegroundColor Cyan
            Get-ADUser -Filter "Department -ne '$null'" | Group-Object Department | Select-Object Name, Count | Sort-Object Name | Format-Table
            $manualPhase2 = Read-Host "Continue with manual user selection? (y/N)"
            if ($manualPhase2 -ne 'y') { exit }
        } else {
            # Validate department users have mailboxes (excluding pilot users)
            $validDeptUsers = @()
            foreach ($user in $deptUsers) {
                if ($user -notin $validPilotUsers) {  # Exclude pilot users
                    try {
                        Get-Mailbox $user -ErrorAction Stop | Out-Null
                        $validDeptUsers += $user
                    } catch {
                        Write-Warning "⚠️  Department user $user has no mailbox"
                    }
                }
            }
            
            Write-Host "Department '$targetDepartment' has $($validDeptUsers.Count) valid users (excluding pilot)" -ForegroundColor Green
            
            if ($validDeptUsers.Count -gt 0) {
                # Show preview
                Write-Host "Phase 2 users (first 10): $($validDeptUsers[0..9] -join ', ')" -ForegroundColor White
                if ($validDeptUsers.Count -gt 10) {
                    Write-Host "... and $($validDeptUsers.Count - 10) more users" -ForegroundColor White
                }
                
                New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Enabled -UserList $validDeptUsers -WhatIf
                $installDept = Read-Host "`nInstall for department '$targetDepartment'? (y/N)"
                if ($installDept -eq 'y') {
                    New-App -OrganizationApp -Url $manifestUrl -DefaultStateForUser Enabled -UserList $validDeptUsers
                    Write-Host "✅ Phase 2 completed: Add-in installed for department '$targetDepartment'" -ForegroundColor Green
                } else {
                    Write-Host "❌ Phase 2 cancelled." -ForegroundColor Red
                    exit
                }
            }
        }
    } catch {
        Write-Error "❌ Phase 2 failed: $($_.Exception.Message)"
        exit
    }
} else {
    Write-Host "❌ Phase 2 skipped - no department specified." -ForegroundColor Yellow
}

# Phase 3: Organization-wide enablement
Write-Host "`n📋 PHASE 3: Organization-wide Enablement" -ForegroundColor Yellow
Write-Host "⚠️  This will enable the add-in for ALL remaining users in the organization" -ForegroundColor Red

$totalUsers = (Get-Mailbox -ResultSize Unlimited).Count
$currentlyInstalled = ($validPilotUsers.Count + $validDeptUsers.Count)
$remainingUsers = $totalUsers - $currentlyInstalled

Write-Host "Current status:" -ForegroundColor White
Write-Host "  Total organization users: $totalUsers" -ForegroundColor White
Write-Host "  Already installed (Phases 1-2): $currentlyInstalled" -ForegroundColor Green
Write-Host "  Remaining users: $remainingUsers" -ForegroundColor Yellow

$orgWideConfirm = Read-Host "`nProceed with organization-wide enablement for remaining $remainingUsers users? Type 'ENABLE-ALL' to confirm"
if ($orgWideConfirm -eq 'ENABLE-ALL') {
    Write-Host "`n🚀 Enabling add-in organization-wide..." -ForegroundColor Green
    Set-App -Identity $addInName -OrganizationApp -DefaultStateForUser Enabled -WhatIf
    
    $finalConfirm = Read-Host "Final confirmation for organization-wide enablement? (y/N)"
    if ($finalConfirm -eq 'y') {
        Set-App -Identity $addInName -OrganizationApp -DefaultStateForUser Enabled
        Write-Host "✅ Phase 3 completed: Add-in enabled organization-wide!" -ForegroundColor Green
        Write-Host "✅ GRADUAL ROLLOUT COMPLETED SUCCESSFULLY!" -ForegroundColor Green
        
        # Final summary
        Write-Host "`n📊 ROLLOUT SUMMARY:" -ForegroundColor Cyan
        Write-Host "Phase 1 (Pilot): $($validPilotUsers.Count) users" -ForegroundColor Green
        Write-Host "Phase 2 (Department): $($validDeptUsers.Count) users" -ForegroundColor Green
        Write-Host "Phase 3 (Organization): $remainingUsers users" -ForegroundColor Green
        Write-Host "Total Affected: $totalUsers users" -ForegroundColor Green
    } else {
        Write-Host "❌ Phase 3 cancelled. Add-in remains enabled only for Phases 1-2." -ForegroundColor Yellow
    }
} else {
    Write-Host "❌ Phase 3 cancelled. Add-in remains enabled only for Phases 1-2." -ForegroundColor Yellow
}
```

### Environment-Specific Deployments
```powershell
# Development environment
New-App -OrganizationApp -Url "https://dev.company.com/promptemail/manifest.xml" -DefaultStateForUser Enabled -UserList "dev-team@company.com"

# Production environment
New-App -OrganizationApp -Url "https://apps.company.com/promptemail/manifest.xml" -DefaultStateForUser Enabled
```

## Best Practices

1. **Test First**: Always test with a small pilot group before organization-wide deployment
2. **Version Management**: Keep track of manifest versions and update URLs as needed
3. **User Communication**: Inform users about new add-in availability and features
4. **Monitor Usage**: Use Exchange logs to monitor add-in adoption and issues
5. **Backup Strategy**: Document current add-in configurations before making changes
6. **Automation**: Consider using the OnPremExchangeAppMgmt framework for complex environments

## Support and Resources

- **PromptEmail Documentation**: See other files in this `docs/` folder
- **Exchange PowerShell Reference**: Microsoft's official Exchange cmdlet documentation

---

*This guide is specific to on-premises Exchange environments. For Exchange Online (Office 365), use the Microsoft 365 Admin Center or Exchange Online PowerShell with different cmdlets.*