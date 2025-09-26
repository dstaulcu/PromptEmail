# On-Premises Exchange Admin Guide for PromptEmail Add-In

This guide provides instructions for on-premises Exchange administrators to deploy, manage, and remove the PromptEmail add-in using PowerShell cmdlets.

## Prerequisites

- Exchange Management Shell access
- Exchange Server 2016 or later (for New-App and Remove-App cmdlets)
- Administrative permissions for add-in management
- Access to the PromptEmail manifest URL

## Quick Reference

### Install Add-In for a User
```powershell
New-App -OrganizationApp -Url "https://your-domain.com/path/to/manifest.xml" -DefaultStateForUser Enabled -UserList "user@domain.com"
```

### Remove Add-In for a User
```powershell
Remove-App -Identity "PromptEmail" -Mailbox "user@domain.com"
```

### Install Add-In for All Users
```powershell
New-App -OrganizationApp -Url "https://your-domain.com/path/to/manifest.xml" -DefaultStateForUser Enabled
```

## Detailed Instructions

### 1. Installing the Add-In

#### For Individual Users
```powershell
# Install for a single user
New-App -OrganizationApp -Url "https://your-domain.com/path/to/manifest.xml" -DefaultStateForUser Enabled -UserList "john.doe@company.com"

# Install for multiple specific users
New-App -OrganizationApp -Url "https://your-domain.com/path/to/manifest.xml" -DefaultStateForUser Enabled -UserList "user1@company.com","user2@company.com","user3@company.com"
```

#### For Distribution Groups
```powershell
# Install for members of a distribution group
New-App -OrganizationApp -Url "https://your-domain.com/path/to/manifest.xml" -DefaultStateForUser Enabled -UserList (Get-DistributionGroupMember "Sales Team" | Select-Object -ExpandProperty PrimarySmtpAddress)
```

#### For All Users in Organization
```powershell
# Install organization-wide (all users)
New-App -OrganizationApp -Url "https://your-domain.com/path/to/manifest.xml" -DefaultStateForUser Enabled

# Install organization-wide but disabled by default (users can enable themselves)
New-App -OrganizationApp -Url "https://your-domain.com/path/to/manifest.xml" -DefaultStateForUser Disabled
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
# Enable for specific user
Set-App -Identity "PromptEmail" -Mailbox "user@company.com" -Enabled $true

# Disable for specific user
Set-App -Identity "PromptEmail" -Mailbox "user@company.com" -Enabled $false

# Enable for multiple users
"user1@company.com","user2@company.com" | ForEach-Object { Set-App -Identity "PromptEmail" -Mailbox $_ -Enabled $true }
```

### 3. Removing the Add-In

#### Remove for Individual Users
```powershell
# Remove from specific user
Remove-App -Identity "PromptEmail" -Mailbox "user@company.com"

# Remove from multiple users
"user1@company.com","user2@company.com" | ForEach-Object { Remove-App -Identity "PromptEmail" -Mailbox $_ }
```

#### Remove Organization-Wide
```powershell
# Remove the organization app entirely (affects all users)
Remove-App -Identity "PromptEmail" -OrganizationApp

# Confirm removal
Remove-App -Identity "PromptEmail" -OrganizationApp -Confirm:$false
```

### 4. Bulk Operations

#### Install for Department/Group
```powershell
# Get users from AD group and install add-in
$users = Get-ADGroupMember "Marketing Department" | Get-ADUser | Select-Object -ExpandProperty UserPrincipalName
New-App -OrganizationApp -Url "https://your-domain.com/path/to/manifest.xml" -DefaultStateForUser Enabled -UserList $users
```

#### Install Based on User Attributes
```powershell
# Install for users in specific department
$users = Get-ADUser -Filter "Department -eq 'Sales'" | Select-Object -ExpandProperty UserPrincipalName
New-App -OrganizationApp -Url "https://your-domain.com/path/to/manifest.xml" -DefaultStateForUser Enabled -UserList $users
```

## Automated Management Framework

For advanced add-in lifecycle management, consider using our **Exchange Add-In Management Framework**:

🔗 **[OnPremExchangeAppMgmt](https://github.com/dstaulcu/OnPremExchangeAppMgmt)**

This framework provides:
- **Automated add-in management** based on Active Directory group membership
- **Environment separation** (dev/test/prod)
- **Comprehensive logging** and audit trails
- **State management** and change tracking
- **Mock servers** for development and testing

### Quick Start with Framework
```powershell
# Clone the framework
git clone https://github.com/dstaulcu/OnPremExchangeAppMgmt.git

# Configure for PromptEmail add-in
# Create AD group: app-exchangeaddin-promptemail-prod
# Add users to group who should have the add-in

# Run the framework
.\src\ExchangeAddInManager.ps1 -ExchangeServer "exchange.company.com" -Domain "company.com"
```

## Troubleshooting

### Common Issues

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
# Phase 1: Pilot group
New-App -OrganizationApp -Url "https://your-domain.com/manifest.xml" -DefaultStateForUser Enabled -UserList (Get-ADGroupMember "PromptEmail-Pilot" | Get-ADUser | Select-Object -ExpandProperty UserPrincipalName)

# Phase 2: Department rollout
New-App -OrganizationApp -Url "https://your-domain.com/manifest.xml" -DefaultStateForUser Enabled -UserList (Get-ADUser -Filter "Department -eq 'Marketing'" | Select-Object -ExpandProperty UserPrincipalName)

# Phase 3: Organization-wide
Set-App -Identity "PromptEmail" -OrganizationApp -DefaultStateForUser Enabled
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
- **OnPremExchangeAppMgmt Framework**: [GitHub Repository](https://github.com/dstaulcu/OnPremExchangeAppMgmt)

---

*This guide is specific to on-premises Exchange environments. For Exchange Online (Office 365), use the Microsoft 365 Admin Center or Exchange Online PowerShell with different cmdlets.*