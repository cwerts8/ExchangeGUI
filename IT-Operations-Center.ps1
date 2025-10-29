<#
.SYNOPSIS
    IT Operations Center - Comprehensive GUI-based management tool for IT infrastructure and services

.DESCRIPTION
    A comprehensive PowerShell GUI tool for managing Exchange Online mailbox permissions, calendar permissions,
    automatic replies (Out of Office), message tracking, and Active Directory group memberships. Features include:
    - Optional Exchange Online connection (connect when needed)
    - Visual connection status indicator with Connect/Disconnect controls
    - Mailbox permissions (Full Access & Send As)
    - Calendar permissions (7 levels)
    - Automatic Replies (OOF) with rich text editor
    - Message Trace / Tracking for email delivery troubleshooting
    - AD group member viewing and management
    - Excel export capabilities for all data
    - Double-click any user/group to view their AD properties
    - GUID resolution for AD groups in permission lists

.AUTHOR
    Created by: Craig Werts
    Department: Desktop Engineering

.VERSION HISTORY
    Version 1.0.0 - 10-06-25
    - Initial release
    - Basic mailbox and calendar permission management
    - AD group member viewing
    
    Version 2.0 - 10-08-25
    - Added Excel export functionality for all three modules
    - Added GUID-to-name resolution for AD groups in permissions
    - Added double-click functionality to view AD properties (ADUC-style)
    - AD Properties window shows 4 tabs: General, Contact Info, Organization, Account
    - Copy email to clipboard feature in AD Properties window
    
    Version 2.6 - 10-09-25
    - Improved error handling and logging
    - Added progress indicators for group member loading
    - Added "Copy Emails" button for group members (Outlook format)
    - Enhanced UI with status indicators
    - AutoMapping disabled by default for Full Access permissions
    - Added "Add_KeyDown" event handler to Mailbox and Calendar windows
    
    Version 2.6.1 - 10-09-25
    - Added Company Logo and dynamic version text
    
    Version 2.7.0 - 10-13-25
    - Added Automatic Replies (Out of Office) management module
    - Implemented rich text editor with formatting toolbar (Bold, Italic, Underline)
    - HTML-enabled message editor - no HTML knowledge required
    - Support for internal and external automatic reply messages
    - Scheduled automatic replies with date/time picker
    - Three reply states: Disabled, Enabled, Scheduled
    - Automatic HTML conversion from formatted text
    - Visual status indicators with color coding
    - External audience controls (All/Contacts only)
    
    Version 2.8.0 - 10-13-25
    - Implemented optional Exchange Online connection
    - GUI now launches immediately without requiring EXO connection
    - Added visual connection status indicator (Red/Green)
    - Connect/Disconnect buttons in GUI
    - Connection check before opening each module
    - Auto-restore GUI after successful authentication
    - Prompts user to connect when accessing modules while disconnected
    - Improved flexibility for users who don't need immediate EXO access
    
    Version 2.8.1 - 10-13-25
    - Fixed bug in mailbox permissions loading when only one delegate exists
    - Fixed bug in calendar permissions loading when only one delegate exists
    - Improved array handling in Get-CombinedMailboxPermissions function
    - Improved array handling in calendar permission retrieval
    - Added null/empty value checking for permission entries
    - Enhanced error logging for permission retrieval
    - Functions now properly return IEnumerable collections for DataGrid binding

    Version 2.9.0 - 10-16-25
    - Modified AD Group Members to use Active Directory module instead of Exchange Online
    - Now supports all AD group types (Security Groups, Distribution Groups, etc.)
    - No longer requires Exchange Online connection for AD Group Members feature
    - Added Group Scope display (DomainLocal, Global, Universal)
    - Enhanced member type support (Users, Groups, Computers, Contacts)
    - Added SAM Account Name to Excel export
    - Requires Active Directory PowerShell module (RSAT)

    Version 3.2.0 - 10-23-25 
    - Added Message Trace / Tracking module for email delivery troubleshooting
    - Search by sender, recipient, subject, message ID, or status
    - Flexible date range options (24 hours, 7 days, or custom up to 10 days)
    - View detailed message trace events and delivery timeline
    - Support for up to 5,000 results per search
    - Export message trace results to Excel
    - Copy Message ID to clipboard for further investigation
    - Status filtering (Delivered, Failed, Pending, Quarantined, FilteredAsSpam)
    - Real-time search progress indicators
    - Full message delivery event tracking with timestamps
    
    Version 3.3.0 - 10-27-25
    - Added IP Network Scanner module for network infrastructure management
    - Scan IP address ranges to discover active devices
    - **Parallel scanning with 50 concurrent threads for fast performance**
    - Test connectivity using Test-Connection (ping) with 1-second timeout
    - Automatic hostname resolution for online devices
    - MAC address detection via ARP table lookup
    - Response time measurement in milliseconds
    - Real-time scan progress with online/offline counters
    - Export scan results to Excel with full device details
    - Support for large IP ranges with progress indicators
    - Color-coded status display (Green=Online, Red=Offline)
    - Typical scan speed: ~50 IPs per second (vs ~1 IP/sec sequential)

        Version 3.4.0 - 10-28-25
    - Added Export Active Users Report module
    - Retrieves all enabled AD user accounts from Active and Consultants OUs
    - Automatically filters out test accounts (containing "test" or "t-" prefix)
    - Exports Name, SamAccountName, Email, Department, Title, and Office (desk location)
    - Results automatically sorted by Name in ascending order
    - User-selectable save location (file save dialog)
    - Excel export with auto-formatting, filters, and frozen header row
    - Requires Active Directory PowerShell module (RSAT)

    Version 3.5.0 - 10-28-25 (Current)
    - Added Intune Mobile Devices module for MDM device management
    - Retrieves all mobile devices (iOS, iPadOS, Android) from Microsoft Intune
    - Dashboard with real-time statistics (device counts by OS, compliance status)
    - Comprehensive device information (name, user, IMEI, serial, model, OS, storage)
    - Excel export with conditional formatting (compliant=green, non-compliant=red, jailbroken=orange)
    - IMEI and MEID formatted as numbers (no decimals) in Excel
    - Automatic Microsoft Graph API connection and consent management
    - Requires Microsoft.Graph.Authentication and Microsoft.Graph.DeviceManagement modules
    - Device enrollment dates, last sync times, and compliance states
    - Sortable grid with filtering capabilities
    - Added Export Active Users Report module
    - Retrieves all enabled AD user accounts from Active and Consultants OUs
    - Automatically filters out test accounts (containing "test" or "t-" prefix)
    - Exports Name, SamAccountName, Email, Department, Title, and Office (desk location)
    - Results automatically sorted by Name in ascending order
    - User-selectable save location (file save dialog)
    - Excel export with auto-formatting, filters, and frozen header row
    - Requires Active Directory PowerShell module (RSAT)

.NOTES
    File Name      : IT-Operations-Center.ps1
    Prerequisite   : ExchangeOnlineManagement module must be installed
    
.USAGE
    Simply run the script. It will:
    1. Check for required modules and install ImportExcel if needed
    2. Connect to Exchange Online (browser authentication)
    3. Launch the GUI management tool
    
    From the GUI you can:
    - Manage Mailboxes: Add/edit/remove Full Access and Send As permissions
    - Calendar Permissions: Add/edit/remove calendar delegation permissions
    - Automatic Replies: Configure Out of Office messages with rich text formatting
    - Message Trace: Track and troubleshoot email delivery issues
    - AD Group Members: View group members, export to Excel, copy email addresses
    - Double-click any user/group name in permission lists to view their AD properties

.FEATURES
    Connection Management:
    - Optional Exchange Online connection (launch GUI without connecting)
    - Visual status indicator (Red=Disconnected, Green=Connected)
    - Connect/Disconnect buttons in GUI
    - Console-based authentication with auto-restore GUI
    - Connection validation before accessing modules
    - Reconnect capability if session expires
    
    Mailbox Permissions:
    - Add Full Access and/or Send As permissions
    - View and edit existing permissions
    - Remove permissions selectively
    - Export permissions to Excel
    - GUID resolution for AD groups
    
    Calendar Permissions:
    - 7 permission levels (AvailabilityOnly to Owner)
    - Add, edit, and remove calendar access
    - Export permissions to Excel
    
    Automatic Replies (Out of Office):
    - Rich text editor with formatting toolbar (Bold, Italic, Underline)
    - Create formatted messages without HTML knowledge
    - Automatic HTML conversion and rendering
    - Internal and external message support
    - Three states: Disabled, Enabled (always on), Scheduled (date/time range)
    - Date/time picker for scheduled replies
    - External audience options (All senders or Contacts only)
    - Color-coded status display (Gray=Disabled, Green=Enabled, Orange=Scheduled)
    - Clear formatting button to remove all text formatting
    
    Message Trace / Tracking:
    - Search by sender, recipient, subject, message ID, or delivery status
    - Flexible date ranges (last 24 hours, last 7 days, custom up to 10 days)
    - Status filtering (All, Delivered, Failed, Pending, Quarantined, FilteredAsSpam)
    - Configurable result limits (100, 1000, 5000 messages)
    - Detailed message trace events with full delivery timeline
    - View complete transport events and routing information
    - Export search results to Excel with full message details
    - Copy Message ID to clipboard for further investigation
    - Real-time search progress indicators
    - Enter key support for quick searches
    - Note: Distribution group searches show emails sent TO the group address
      (external emails to groups are delivered to individual members and may not appear)
    
    AD Group Members:
    - Search by group name or email
    - View all members with enriched details (title, department)
    - Copy all member emails to clipboard (Outlook format)
    - Export to Excel with full member details
    
    AD Properties Viewer:
    - Double-click any user/group in any grid to view properties
    - General tab: Name, email, title, department, office, company
    - Contact Information tab: Phone, mobile, address
    - Organization tab: Manager, direct reports, group memberships
    - Account tab: UPN, SAM, DN, GUID, creation/modification dates

    IP Network Scanner:
    - Scan custom IP address ranges (start IP to end IP)
    - Real-time connectivity testing using Test-Connection
    - Automatic hostname resolution for discovered devices
    - MAC address lookup via ARP table
    - Response time measurement in milliseconds
    - Live scan progress with online/offline counters
    - Color-coded status display (Green=Online, Red=Offline)
    - Export complete scan results to Excel
    - Support for large IP ranges with progress tracking
    - Warning prompts for scans over 1,000 addresses
    
    Export Active Users Report:
    - One-click export of all enabled AD user accounts
    - Retrieves users from Active and Consultants OUs
    - Automatic filtering of test accounts (containing "test" or "t-" prefix)
    - Exports: Name, SamAccountName, Email, Department, Title, Office (desk location)
    - Results automatically sorted by Name
    - User-selectable save location with file save dialog
    - Excel export with professional formatting, filters, and frozen headers
    - Timestamped filenames for easy organization
    - Requires Active Directory PowerShell module (RSAT)v

.PROXY CONFIGURATION
    The script includes proxy configuration for corporate environments.
    Update the proxy URL in lines 209-211 if needed for your organization.

#>

# Update Script version
$ScriptVersion = "3.5.0"

# Load logo from file path
# Embedded logo (Base64 encoded)
$logoBase64 = @"
iVBORw0KGgoAAAANSUhEUgAABrgAAAGQCAYAAAD1DmBrAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAyZpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/I
iBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDkuMC1jMDAwIDc5LjE3MWMyN2ZhYiwgMjAyMi
8wOC8xNi0yMjozNTo0MSAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiI
geG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RSZWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20v
eGFwLzEuMC9zVHlwZS9SZXNvdXJjZVJlZiMiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIDI0LjEgKFdpbmRvd3MpIiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOkFBQzc0NzJEOTUyQTExRURCN
0Y1REIyODJFOUE1OUVCIiB4bXBNTTpEb2N1bWVudElEPSJ4bXAuZGlkOkFBQzc0NzJFOTUyQTExRURCN0Y1REIyODJFOUE1OUVCIj4gPHhtcE1NOkRlcml2ZWRGcm9tIHN0UmVmOmluc3RhbmNlSUQ9InhtcC5paW
Q6QUFDNzQ3MkI5NTJBMTFFREI3RjVEQjI4MkU5QTU5RUIiIHN0UmVmOmRvY3VtZW50SUQ9InhtcC5kaWQ6QUFDNzQ3MkM5NTJBMTFFREI3RjVEQjI4MkU5QTU5RUIiLz4gPC9yZGY6RGVzY3JpcHRpb24+IDwvcmR
mOlJERj4gPC94OnhtcG1ldGE+IDw/eHBhY2tldCBlbmQ9InIiPz4nPYG+AADI90lEQVR42uy9CZgcV33u/a+xZQvvwsbyvjAY2xgvgpsYDyQhLCEi4SYQQgJcEnMTAjE2mH1Nwr6EwAVCQjBg9HEJNwGyQABBAgQP
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
"@

# Function to load logo from Base64 string
function Get-LogoFromBase64 {
    param([string]$Base64String)
    
    try {
        # Convert Base64 to byte array
        $imageBytes = [Convert]::FromBase64String($Base64String)
        
        # Create a memory stream from the bytes
        $memoryStream = New-Object System.IO.MemoryStream
        $memoryStream.Write($imageBytes, 0, $imageBytes.Length)
        $memoryStream.Position = 0
        
        # Create BitmapImage from stream
        $logoImage = New-Object System.Windows.Media.Imaging.BitmapImage
        $logoImage.BeginInit()
        $logoImage.StreamSource = $memoryStream
        $logoImage.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $logoImage.EndInit()
        $logoImage.Freeze()
        
        return $logoImage
    } catch {
        Write-Log "Could not load embedded logo: $($_.Exception.Message)"
        return $null
    }
}

#Requires -Modules ExchangeOnlineManagement

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Host "Restarting in STA mode..." -ForegroundColor Yellow
    Start-Process powershell.exe -ArgumentList "-STA", "-NoProfile", "-File", "`"$PSCommandPath`"" -Wait
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12  

[system.net.webrequest]::defaultwebproxy = new-object system.net.webproxy('http://proxy.gellerco.com:8080')
[system.net.webrequest]::defaultwebproxy.credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[system.net.webrequest]::defaultwebproxy.BypassProxyOnLocal = $true



<# Pre-connection check - OPTIONAL NOW
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "IT Operations Center" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""
#>

# Import the module
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Check if ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
       Write-Host "ImportExcel module installed successfully!" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not install ImportExcel module. Excel export feature will not be available." -ForegroundColor Yellow
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Check current connection status
$existingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue

if ($null -ne $existingConnection -and $existingConnection.State -eq 'Connected') {
#    Write-Host "Existing Exchange Online connection detected!" -ForegroundColor Green
#    Write-Host "Connected as: $($existingConnection.UserPrincipalName)" -ForegroundColor Cyan
#    Write-Host "Connection State: $($existingConnection.State)" -ForegroundColor Cyan
#    Write-Host ""
} else {
#    Write-Host "No Exchange Online connection detected." -ForegroundColor Yellow
#    Write-Host "You can connect later from the GUI if needed." -ForegroundColor Yellow
#    Write-Host ""
}

# Write-Host "Launching GUI..." -ForegroundColor Green
Start-Sleep -Seconds 1

$syncHash = [hashtable]::Synchronized(@{})

function Show-ADPropertiesWindow {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Identity,
        
        [Parameter(Mandatory=$false)]
        [System.Windows.Window]$Owner
    )
    
    try {
        Write-Log "Loading AD properties for: $Identity"
        
        $recipient = Get-Recipient -Identity $Identity -ErrorAction Stop
        
        $isGroup = $false
        $detailedInfo = $null
        
        if ($recipient.RecipientType -like "*Group*") {
            $isGroup = $true
            try {
                $detailedInfo = Get-DistributionGroup -Identity $Identity -ErrorAction Stop
            } catch {
                $detailedInfo = Get-Group -Identity $Identity -ErrorAction Stop
            }
        } else {
            try {
                $detailedInfo = Get-User -Identity $Identity -ErrorAction Stop
            } catch {
                $detailedInfo = Get-Mailbox -Identity $Identity -ErrorAction SilentlyContinue
            }
        }
        
        [xml]$PropertiesXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Properties - $($recipient.DisplayName)" 
        Height="600" 
        Width="550" 
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <Border Grid.Row="0" Background="#233A4A" Padding="20,15">
            <StackPanel>
                <TextBlock Text="$($recipient.DisplayName)" FontSize="18" FontWeight="Bold" Foreground="White"/>
                <TextBlock Text="$($recipient.RecipientType)" FontSize="12" Foreground="#B0BEC5" Margin="0,5,0,0"/>
            </StackPanel>
        </Border>
        
        <TabControl Grid.Row="1" Margin="10">
            <TabItem Header="General">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <TextBlock Text="Display Name:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="DisplayNameBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Email Address:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="EmailBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock x:Name="TitleLabel" Text="Title:" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <TextBox x:Name="TitleBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5" Visibility="Collapsed"/>
                        
                        <TextBlock x:Name="DepartmentLabel" Text="Department:" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <TextBox x:Name="DepartmentBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5" Visibility="Collapsed"/>
                        
                        <TextBlock x:Name="OfficeLabel" Text="Office:" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <TextBox x:Name="OfficeBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5" Visibility="Collapsed"/>
                        
                        <TextBlock x:Name="CompanyLabel" Text="Company:" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <TextBox x:Name="CompanyBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5" Visibility="Collapsed"/>
                        
                        <TextBlock Text="Recipient Type:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="RecipientTypeBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock x:Name="ManagedByLabel" Text="Managed By:" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <TextBox x:Name="ManagedByBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5" Visibility="Collapsed"/>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <TabItem Header="Contact Information" x:Name="ContactTab">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <TextBlock Text="Phone:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="PhoneBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Mobile:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="MobileBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Fax:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="FaxBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Street Address:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="StreetBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="City:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="CityBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="State/Province:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="StateBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="ZIP/Postal Code:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="PostalBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Country:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="CountryBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <TabItem Header="Organization">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <TextBlock Text="Manager:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="ManagerBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock x:Name="DirectReportsLabel" Text="Direct Reports:" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <Border x:Name="DirectReportsBorder" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15" MaxHeight="150" Visibility="Collapsed">
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <ItemsControl x:Name="DirectReportsList" Padding="5"/>
                            </ScrollViewer>
                        </Border>
                        
                        <TextBlock x:Name="MemberOfLabel" Text="Member Of (Groups):" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <Border x:Name="MemberOfBorder" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15" MaxHeight="200" Visibility="Collapsed">
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <ItemsControl x:Name="MemberOfList" Padding="5"/>
                            </ScrollViewer>
                        </Border>
                        
                        <TextBlock x:Name="MembersLabel" Text="Group Members:" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <TextBlock x:Name="MembersCount" FontSize="11" Foreground="#666" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <Border x:Name="MembersBorder" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15" MaxHeight="200" Visibility="Collapsed">
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <ItemsControl x:Name="MembersList" Padding="5"/>
                            </ScrollViewer>
                        </Border>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <TabItem Header="Account">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <TextBlock Text="User Principal Name (UPN):" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="UPNBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="SAM Account Name:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="SAMBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Distinguished Name:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="DNBox" IsReadOnly="True" TextWrapping="Wrap" Margin="0,0,0,15" Padding="5" MinHeight="60"/>
                        
                        <TextBlock Text="Object GUID:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="GUIDBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Created:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="WhenCreatedBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Modified:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="WhenChangedBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>
        
        <Border Grid.Row="2" Background="#f8f9fa" BorderBrush="#dee2e6" BorderThickness="0,1,0,0" Padding="15">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="CopyEmailBtn" Content="Copy Email" Width="100" Height="30" Margin="0,0,10,0" Background="#6c757d" Foreground="White"/>
                <Button x:Name="CloseBtn" Content="Close" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
        
        $propReader = New-Object System.Xml.XmlNodeReader $PropertiesXAML
        $PropWindow = [Windows.Markup.XamlReader]::Load($propReader)
        if ($Owner) { $PropWindow.Owner = $Owner }
        
        $DisplayNameBox = $PropWindow.FindName("DisplayNameBox")
        $EmailBox = $PropWindow.FindName("EmailBox")
        $TitleLabel = $PropWindow.FindName("TitleLabel")
        $TitleBox = $PropWindow.FindName("TitleBox")
        $DepartmentLabel = $PropWindow.FindName("DepartmentLabel")
        $DepartmentBox = $PropWindow.FindName("DepartmentBox")
        $OfficeLabel = $PropWindow.FindName("OfficeLabel")
        $OfficeBox = $PropWindow.FindName("OfficeBox")
        $CompanyLabel = $PropWindow.FindName("CompanyLabel")
        $CompanyBox = $PropWindow.FindName("CompanyBox")
        $RecipientTypeBox = $PropWindow.FindName("RecipientTypeBox")
        $ManagedByLabel = $PropWindow.FindName("ManagedByLabel")
        $ManagedByBox = $PropWindow.FindName("ManagedByBox")
        
        $ContactTab = $PropWindow.FindName("ContactTab")
        $PhoneBox = $PropWindow.FindName("PhoneBox")
        $MobileBox = $PropWindow.FindName("MobileBox")
        $FaxBox = $PropWindow.FindName("FaxBox")
        $StreetBox = $PropWindow.FindName("StreetBox")
        $CityBox = $PropWindow.FindName("CityBox")
        $StateBox = $PropWindow.FindName("StateBox")
        $PostalBox = $PropWindow.FindName("PostalBox")
        $CountryBox = $PropWindow.FindName("CountryBox")
        
        $ManagerBox = $PropWindow.FindName("ManagerBox")
        $DirectReportsLabel = $PropWindow.FindName("DirectReportsLabel")
        $DirectReportsBorder = $PropWindow.FindName("DirectReportsBorder")
        $DirectReportsList = $PropWindow.FindName("DirectReportsList")
        $MemberOfLabel = $PropWindow.FindName("MemberOfLabel")
        $MemberOfBorder = $PropWindow.FindName("MemberOfBorder")
        $MemberOfList = $PropWindow.FindName("MemberOfList")
        $MembersLabel = $PropWindow.FindName("MembersLabel")
        $MembersCount = $PropWindow.FindName("MembersCount")
        $MembersBorder = $PropWindow.FindName("MembersBorder")
        $MembersList = $PropWindow.FindName("MembersList")
        
        $UPNBox = $PropWindow.FindName("UPNBox")
        $SAMBox = $PropWindow.FindName("SAMBox")
        $DNBox = $PropWindow.FindName("DNBox")
        $GUIDBox = $PropWindow.FindName("GUIDBox")
        $WhenCreatedBox = $PropWindow.FindName("WhenCreatedBox")
        $WhenChangedBox = $PropWindow.FindName("WhenChangedBox")
        
        $CopyEmailBtn = $PropWindow.FindName("CopyEmailBtn")
        $CloseBtn = $PropWindow.FindName("CloseBtn")
        
        $DisplayNameBox.Text = if ($recipient.DisplayName) { $recipient.DisplayName } else { "" }
        $EmailBox.Text = if ($recipient.PrimarySmtpAddress) { $recipient.PrimarySmtpAddress.ToString() } else { "N/A" }
        $RecipientTypeBox.Text = $recipient.RecipientType
        
        if (-not $isGroup) {
            if ($detailedInfo.Title) {
                $TitleLabel.Visibility = [System.Windows.Visibility]::Visible
                $TitleBox.Visibility = [System.Windows.Visibility]::Visible
                $TitleBox.Text = $detailedInfo.Title
            }
            if ($detailedInfo.Department) {
                $DepartmentLabel.Visibility = [System.Windows.Visibility]::Visible
                $DepartmentBox.Visibility = [System.Windows.Visibility]::Visible
                $DepartmentBox.Text = $detailedInfo.Department
            }
            if ($detailedInfo.Office) {
                $OfficeLabel.Visibility = [System.Windows.Visibility]::Visible
                $OfficeBox.Visibility = [System.Windows.Visibility]::Visible
                $OfficeBox.Text = $detailedInfo.Office
            }
            if ($detailedInfo.Company) {
                $CompanyLabel.Visibility = [System.Windows.Visibility]::Visible
                $CompanyBox.Visibility = [System.Windows.Visibility]::Visible
                $CompanyBox.Text = $detailedInfo.Company
            }
            
            $PhoneBox.Text = if ($detailedInfo.Phone) { $detailedInfo.Phone } else { "" }
            $MobileBox.Text = if ($detailedInfo.MobilePhone) { $detailedInfo.MobilePhone } else { "" }
            $FaxBox.Text = if ($detailedInfo.Fax) { $detailedInfo.Fax } else { "" }
            $StreetBox.Text = if ($detailedInfo.StreetAddress) { $detailedInfo.StreetAddress } else { "" }
            $CityBox.Text = if ($detailedInfo.City) { $detailedInfo.City } else { "" }
            $StateBox.Text = if ($detailedInfo.StateOrProvince) { $detailedInfo.StateOrProvince } else { "" }
            $PostalBox.Text = if ($detailedInfo.PostalCode) { $detailedInfo.PostalCode } else { "" }
            $CountryBox.Text = if ($detailedInfo.CountryOrRegion) { $detailedInfo.CountryOrRegion } else { "" }
            
            if ($detailedInfo.Manager) {
                try {
                    $mgr = Get-Recipient -Identity $detailedInfo.Manager -ErrorAction SilentlyContinue
                    $ManagerBox.Text = if ($mgr) { $mgr.DisplayName } else { $detailedInfo.Manager }
                } catch {
                    $ManagerBox.Text = $detailedInfo.Manager
                }
            }
            
            if ($detailedInfo.DirectReports -and $detailedInfo.DirectReports.Count -gt 0) {
				$DirectReportsLabel.Visibility = [System.Windows.Visibility]::Visible
				$DirectReportsBorder.Visibility = [System.Windows.Visibility]::Visible
				
				# Get and sort direct reports
				$sortedReports = @()
				foreach ($dr in $detailedInfo.DirectReports) {
					try {
						$drRecip = Get-Recipient -Identity $dr -ErrorAction SilentlyContinue
						$sortedReports += if ($drRecip) { $drRecip.DisplayName } else { $dr }
					} catch {
						$sortedReports += $dr
					}
				}
				
				# Sort alphabetically
				$sortedReports = $sortedReports | Sort-Object
				
				foreach ($name in $sortedReports) {
					$tb = New-Object System.Windows.Controls.TextBlock
					$tb.Text = "• $name"
					$tb.Margin = "0,2,0,2"
					$DirectReportsList.Items.Add($tb) | Out-Null
				}
			}
            
            try {
				$groups = Get-Recipient -Identity $Identity -ErrorAction Stop | Select-Object -ExpandProperty DistinguishedName | ForEach-Object {
					Get-Recipient -Filter "Members -eq '$_'" -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup -ErrorAction SilentlyContinue
				}
				
				if ($groups -and $groups.Count -gt 0) {
					$MemberOfLabel.Visibility = [System.Windows.Visibility]::Visible
					$MemberOfBorder.Visibility = [System.Windows.Visibility]::Visible
					
					# Sort groups alphabetically
					$sortedGroups = $groups | Sort-Object DisplayName
					
					foreach ($grp in $sortedGroups) {
						$tb = New-Object System.Windows.Controls.TextBlock
						$tb.Text = "• $($grp.DisplayName)"
						$tb.Margin = "0,2,0,2"
						$MemberOfList.Items.Add($tb) | Out-Null
					}
				}
			} catch {}
            
        } else {
            $ContactTab.Visibility = [System.Windows.Visibility]::Collapsed
            
            if ($detailedInfo.ManagedBy) {
                $ManagedByLabel.Visibility = [System.Windows.Visibility]::Visible
                $ManagedByBox.Visibility = [System.Windows.Visibility]::Visible
                try {
                    $mgr = Get-Recipient -Identity $detailedInfo.ManagedBy[0] -ErrorAction SilentlyContinue
                    $ManagedByBox.Text = if ($mgr) { $mgr.DisplayName } else { $detailedInfo.ManagedBy[0] }
                } catch {
                    $ManagedByBox.Text = $detailedInfo.ManagedBy[0]
                }
            }
            
            try {
				$members = Get-DistributionGroupMember -Identity $Identity -ErrorAction Stop
				if ($members -and $members.Count -gt 0) {
					$MembersLabel.Visibility = [System.Windows.Visibility]::Visible
					$MembersCount.Visibility = [System.Windows.Visibility]::Visible
					$MembersBorder.Visibility = [System.Windows.Visibility]::Visible
					$MembersCount.Text = "Total: $($members.Count) members"
					
					# Sort members alphabetically
					$sortedMembers = $members | Sort-Object DisplayName | Select-Object -First 50
					
					foreach ($mem in $sortedMembers) {
						$tb = New-Object System.Windows.Controls.TextBlock
						$tb.Text = "• $($mem.DisplayName)"
						$tb.Margin = "0,2,0,2"
						$MembersList.Items.Add($tb) | Out-Null
					}
					
					if ($members.Count -gt 50) {
						$tb = New-Object System.Windows.Controls.TextBlock
						$tb.Text = "... and $($members.Count - 50) more"
						$tb.Margin = "0,2,0,2"
						$tb.FontStyle = [System.Windows.FontStyles]::Italic
						$tb.Foreground = [System.Windows.Media.Brushes]::Gray
						$MembersList.Items.Add($tb) | Out-Null
					}
				}
			} catch {}
        }
        
        $UPNBox.Text = if ($recipient.UserPrincipalName) { $recipient.UserPrincipalName } else { "" }
        $SAMBox.Text = if ($detailedInfo.SamAccountName) { $detailedInfo.SamAccountName } else { "" }
        $DNBox.Text = if ($recipient.DistinguishedName) { $recipient.DistinguishedName } else { "" }
        $GUIDBox.Text = if ($recipient.Guid) { $recipient.Guid.ToString() } else { "" }
        $WhenCreatedBox.Text = if ($detailedInfo.WhenCreated) { $detailedInfo.WhenCreated.ToString() } else { "" }
        $WhenChangedBox.Text = if ($detailedInfo.WhenChanged) { $detailedInfo.WhenChanged.ToString() } else { "" }
        
        $CopyEmailBtn.Add_Click({
            if ($EmailBox.Text -and $EmailBox.Text -ne "N/A") {
                [System.Windows.Forms.Clipboard]::SetText($EmailBox.Text)
                [System.Windows.MessageBox]::Show("Email address copied to clipboard!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
        })
        
        $CloseBtn.Add_Click({ $PropWindow.Close() })
        
        $PropWindow.ShowDialog() | Out-Null
        
    } catch {
        Write-Log "Error loading properties: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Error loading properties:`n`n$($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="IT Operations Center" 
        Height="750" 
        Width="850" 
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanMinimize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="100"/>
        </Grid.RowDefinitions>
        
        <!-- Header Section with Logo -->
        <Border Grid.Row="0" Background="#233A4A" Padding="15">
            <DockPanel>
                <Image x:Name="CompanyLogo" 
                       DockPanel.Dock="Left"
                       Width="200" 
                       Height="50" 
                       Margin="0,0,15,0"
                       Stretch="Uniform"
                       VerticalAlignment="Center"/>
                <StackPanel DockPanel.Dock="Left" VerticalAlignment="Center">
                    <TextBlock Text="IT Operations Center" 
                              FontSize="20" 
                              FontWeight="Bold" 
                              Foreground="White"/>
                </StackPanel>
                <TextBlock x:Name="VersionText"
                          Text="v2.9.0"
                          FontSize="11"
                          Foreground="#B0BEC5"
                          VerticalAlignment="Bottom"
                          HorizontalAlignment="Right"
                          DockPanel.Dock="Right"/>
            </DockPanel>
        </Border>
        
        
        <!-- Main Content Area -->
        <Border Grid.Row="1" Background="WhiteSmoke" Padding="20">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <!-- Connection Status -->
                <GroupBox Grid.Row="0" Header="Connection Status" Padding="10" Margin="0,0,0,15">
                    <StackPanel>
                        <DockPanel Margin="0,0,0,10">
                            <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
                                <Ellipse x:Name="ConnectionStatusIndicator" 
                                        Width="12" 
                                        Height="12" 
                                        Fill="Red" 
                                        Margin="0,0,10,0"
                                        VerticalAlignment="Center"/>
                                <TextBlock x:Name="ConnectionStatusText" 
                                        Text="Not Connected" 
                                        FontWeight="Bold"
                                        VerticalAlignment="Center"
                                        Foreground="Red"/>
                            </StackPanel>
                        </DockPanel>
                        
                        <TextBlock x:Name="ConnectionInfoText" 
                                Text="Click 'Connect' to authenticate to Exchange Online" 
                                TextWrapping="Wrap" 
                                Margin="0,0,0,15"
                                FontSize="11"
                                Foreground="#666"/>
                        
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                            <Button x:Name="ConnectButton" 
                                Content="Connect to Exchange Online" 
                                Width="180" 
                                Height="35" 
                                Margin="5"
                                Background="#28a745"
                                Foreground="White"
                                FontWeight="Bold"
                                Cursor="Hand"/>
                            <Button x:Name="DisconnectButton" 
                                Content="Disconnect" 
                                Width="120" 
                                Height="35" 
                                Margin="5"
                                Background="#dc3545"
                                Foreground="White"
                                FontWeight="Bold"
                                Cursor="Hand"
                                IsEnabled="False"/>
                        </StackPanel>
                    </StackPanel>
                </GroupBox>
                
<!-- REORGANIZED MANAGEMENT OPTIONS -->
<!-- This replaces the Management Options section starting around line 1269 -->

                <GroupBox Grid.Row="1" Header="Management Options" 
                        x:Name="ManagementGroup" 
                        Padding="10">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel>
                            
                            <!-- Exchange Online Section -->
                            <Expander IsExpanded="False" Margin="0,0,0,5">
                                <Expander.Header>
                                    <StackPanel Orientation="Horizontal">
                                        <Ellipse Width="12" Height="12" Fill="#007bff" Margin="0,0,8,0" VerticalAlignment="Center"/>
                                        <TextBlock Text="Exchange Online" FontWeight="Bold" FontSize="13"/>
                                    </StackPanel>
                                </Expander.Header>
                                <StackPanel Margin="20,5,0,5">
                                    <Button x:Name="MailboxButton" 
                                        Content="Mailbox Permissions (Full Access &amp; Send As)" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="CalendarButton" 
                                        Content="Calendar Permissions" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="AutoRepliesButton" 
                                        Content="Automatic Replies (Out of Office)" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="MessageTraceButton" 
                                        Content="Message Trace / Tracking" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="SendOnBehalfButton" 
                                        Content="Send on Behalf Permissions" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                    <Button x:Name="ForwardingButton" 
                                        Content="Email Forwarding Management" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                    <Button x:Name="ResourceMailboxButton" 
                                        Content="Resource Mailbox Management" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                </StackPanel>
                            </Expander>
                            
                            <!-- Active Directory Section -->
                            <Expander IsExpanded="False" Margin="0,0,0,5">
                                <Expander.Header>
                                    <StackPanel Orientation="Horizontal">
                                        <Ellipse Width="12" Height="12" Fill="#28a745" Margin="0,0,8,0" VerticalAlignment="Center"/>
                                        <TextBlock Text="Active Directory" FontWeight="Bold" FontSize="13"/>
                                    </StackPanel>
                                </Expander.Header>
                                <StackPanel Margin="20,5,0,5">
                                    <Button x:Name="GroupMembersButton" 
                                        Content="AD Group Members Viewer" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="ExportActiveUsersButton" 
                                        Content="Export Active Users Report" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="DistributionGroupButton" 
                                        Content="Distribution List Management" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                </StackPanel>
                            </Expander>
                            
                            <!-- Reports & Analytics Section -->
                            <Expander IsExpanded="False" Margin="0,0,0,5">
                                <Expander.Header>
                                    <StackPanel Orientation="Horizontal">
                                        <Ellipse Width="12" Height="12" Fill="#17a2b8" Margin="0,0,8,0" VerticalAlignment="Center"/>
                                        <TextBlock Text="Reports &amp; Analytics" FontWeight="Bold" FontSize="13"/>
                                    </StackPanel>
                                </Expander.Header>
                                <StackPanel Margin="20,5,0,5">
                                    <Button x:Name="MailboxStatsButton" 
                                        Content="Mailbox Size &amp; Quota Report" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="PermissionAuditButton" 
                                        Content="Permission Audit Report" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                </StackPanel>
                            </Expander>
                            
                            <!-- Device Management Section -->
                            <Expander IsExpanded="False" Margin="0,0,0,5">
                                <Expander.Header>
                                    <StackPanel Orientation="Horizontal">
                                        <Ellipse Width="12" Height="12" Fill="#FF9800" Margin="0,0,8,0" VerticalAlignment="Center"/>
                                        <TextBlock Text="Device Management" FontWeight="Bold" FontSize="13"/>
                                    </StackPanel>
                                </Expander.Header>
                                <StackPanel Margin="20,5,0,5">
                                    <Button x:Name="IntuneMobileButton" 
                                        Content="Intune Mobile Devices" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="SCCMDevicesButton" 
                                        Content="SCCM Device Management" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                    <Button x:Name="IntuneComplianceButton" 
                                        Content="Compliance Policy Reports" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                </StackPanel>
                            </Expander>
                            
                            <!-- Network & Infrastructure Section -->
                            <Expander IsExpanded="False" Margin="0,0,0,5">
                                <Expander.Header>
                                    <StackPanel Orientation="Horizontal">
                                        <Ellipse Width="12" Height="12" Fill="#6f42c1" Margin="0,0,8,0" VerticalAlignment="Center"/>
                                        <TextBlock Text="Network &amp; Infrastructure" FontWeight="Bold" FontSize="13"/>
                                    </StackPanel>
                                </Expander.Header>
                                <StackPanel Margin="20,5,0,5">
                                    <Button x:Name="IPScannerButton" 
                                        Content="IP Network Scanner" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                </StackPanel>
                            </Expander>
                            
                            <!-- Compliance & Security Section -->
                            <Expander IsExpanded="False" Margin="0,0,0,5">
                                <Expander.Header>
                                    <StackPanel Orientation="Horizontal">
                                        <Ellipse Width="12" Height="12" Fill="#dc3545" Margin="0,0,8,0" VerticalAlignment="Center"/>
                                        <TextBlock Text="Compliance &amp; Security" FontWeight="Bold" FontSize="13"/>
                                    </StackPanel>
                                </Expander.Header>
                                <StackPanel Margin="20,5,0,5">
                                    <Button x:Name="LitigationHoldButton" 
                                        Content="Litigation Hold Management" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                </StackPanel>
                            </Expander>

                        </StackPanel>
                    </ScrollViewer>
                </GroupBox>
            </Grid>
        </Border>
        
        <!-- Status/Log Section -->
        <Border Grid.Row="2" Background="#f8f9fa" BorderBrush="#dee2e6" BorderThickness="0,1,0,0">
            <DockPanel Margin="10">
                <TextBlock Text="Activity Log:" 
                          DockPanel.Dock="Top" 
                          FontWeight="Bold" 
                          Margin="0,0,0,5"/>
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <TextBox x:Name="LogBox" 
                            IsReadOnly="True" 
                            Background="Transparent" 
                            BorderThickness="0"
                            TextWrapping="Wrap"/>
                </ScrollViewer>
            </DockPanel>
        </Border>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($reader)

$syncHash.Window = $Window

$syncHash.ConnectionStatusIndicator = $Window.FindName("ConnectionStatusIndicator")
$syncHash.ConnectionStatusText = $Window.FindName("ConnectionStatusText")
$syncHash.ConnectionInfoText = $Window.FindName("ConnectionInfoText")
$syncHash.ConnectButton = $Window.FindName("ConnectButton")
$syncHash.DisconnectButton = $Window.FindName("DisconnectButton")
$syncHash.StatusText = $Window.FindName("StatusText")
$syncHash.LogBox = $Window.FindName("LogBox")
$syncHash.ManagementGroup = $Window.FindName("ManagementGroup")
$syncHash.MailboxButton = $Window.FindName("MailboxButton")
$syncHash.CalendarButton = $Window.FindName("CalendarButton")
$syncHash.GroupMembersButton = $Window.FindName("GroupMembersButton")
$syncHash.VersionText = $Window.FindName("VersionText")
$syncHash.AutoRepliesButton = $Window.FindName("AutoRepliesButton")

# Mailbox Management - Future
$syncHash.SendOnBehalfButton = $Window.FindName("SendOnBehalfButton")
$syncHash.ForwardingButton = $Window.FindName("ForwardingButton")

# Calendar & Resources - Future
$syncHash.ResourceMailboxButton = $Window.FindName("ResourceMailboxButton")

# Groups & Distribution - Future
$syncHash.ExportActiveUsersButton = $Window.FindName("ExportActiveUsersButton")
$syncHash.DistributionGroupButton = $Window.FindName("DistributionGroupButton")

# Compliance & Security - Future
$syncHash.MessageTraceButton = $Window.FindName("MessageTraceButton")
$syncHash.LitigationHoldButton = $Window.FindName("LitigationHoldButton")

# Reports & Analytics - Future
$syncHash.MailboxStatsButton = $Window.FindName("MailboxStatsButton")
$syncHash.PermissionAuditButton = $Window.FindName("PermissionAuditButton")

# Network & Infrastructure
$syncHash.IPScannerButton = $Window.FindName("IPScannerButton")
$syncHash.IntuneMobileButton = $Window.FindName("IntuneMobileButton")
$syncHash.SCCMDevicesButton = $Window.FindName("SCCMDevicesButton")

$syncHash.IntuneComplianceButton = $Window.FindName("IntuneComplianceButton")



# Set dynamic version text
$syncHash.Window.Title = "IT Operations Center v$ScriptVersion"
$syncHash.VersionText.Text = "v$ScriptVersion"

function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logEntry = "[$timestamp] $Message`r`n"
    $syncHash.Window.Dispatcher.Invoke([action]{
        $syncHash.LogBox.AppendText($logEntry)
    })
}

# Function to update connection status in GUI
function Update-ConnectionStatus {
    $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
    
    if ($null -ne $connInfo -and $connInfo.State -eq 'Connected') {
        $syncHash.ConnectionStatusIndicator.Fill = [System.Windows.Media.Brushes]::Green
        $syncHash.ConnectionStatusText.Text = "Connected"
        $syncHash.ConnectionStatusText.Foreground = [System.Windows.Media.Brushes]::Green
        $syncHash.ConnectionInfoText.Text = "Connected as: $($connInfo.UserPrincipalName)"
        $syncHash.ConnectButton.IsEnabled = $false
        $syncHash.DisconnectButton.IsEnabled = $true
        Write-Log "Connected to Exchange Online as $($connInfo.UserPrincipalName)"
    } else {
        $syncHash.ConnectionStatusIndicator.Fill = [System.Windows.Media.Brushes]::Red
        $syncHash.ConnectionStatusText.Text = "Not Connected"
        $syncHash.ConnectionStatusText.Foreground = [System.Windows.Media.Brushes]::Red
        $syncHash.ConnectionInfoText.Text = "Click 'Connect' to authenticate to Exchange Online"
        $syncHash.ConnectButton.IsEnabled = $true
        $syncHash.DisconnectButton.IsEnabled = $false
        Write-Log "Not connected to Exchange Online"
    }
}

# Load embedded logo
try {
    $logoImage = Get-LogoFromBase64 -Base64String $logoBase64
    if ($null -ne $logoImage) {
        $CompanyLogo = $Window.FindName("CompanyLogo")
        $CompanyLogo.Source = $logoImage
        Write-Log "Company logo loaded from embedded data"
    }
} catch {
    Write-Log "Could not load embedded company logo: $($_.Exception.Message)"
}

# Update connection status on load
Update-ConnectionStatus

$syncHash.ConnectButton.Add_Click({
    Write-Log "Initiating Exchange Online connection..."
    
    # Minimize the GUI window
    $syncHash.Window.WindowState = [System.Windows.WindowState]::Minimized
    
    <#
    Show console message
    $host.UI.RawUI.ForegroundColor = "Cyan"
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "CONNECTING TO EXCHANGE ONLINE" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "A browser window will open for authentication..." -ForegroundColor Yellow
    Write-Host "Please complete the authentication process." -ForegroundColor Yellow
    Write-Host ""
    #>

    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        
#       Write-Host ""
       Write-Host "Successfully connected to Exchange Online!" -ForegroundColor Green
#        Write-Host ""
#        Write-Host "Returning to GUI..." -ForegroundColor Green

        # Restore the GUI window
        $syncHash.Window.WindowState = [System.Windows.WindowState]::Normal
        $syncHash.Window.Activate()
        
        # Update connection status
        Update-ConnectionStatus
        
        [System.Windows.MessageBox]::Show(
            "Successfully connected to Exchange Online!",
            "Connected",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
        
    } catch {
#        Write-Host ""
        Write-Host "Connection failed: $($_.Exception.Message)" -ForegroundColor Red
#        Write-Host ""
#        Write-Host "Press ENTER to return to the GUI..." -ForegroundColor Yellow
        $null = Read-Host
        
        # Restore the GUI window
        $syncHash.Window.WindowState = [System.Windows.WindowState]::Normal
        $syncHash.Window.Activate()
        
        Write-Log "Connection failed: $($_.Exception.Message)"
        
        [System.Windows.MessageBox]::Show(
            "Failed to connect to Exchange Online:`n`n$($_.Exception.Message)`n`nPlease try again.",
            "Connection Failed",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
})

$syncHash.DisconnectButton.Add_Click({
    Write-Log "Disconnecting from Exchange Online..."
    
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-Log "Successfully disconnected from Exchange Online"
        
        Update-ConnectionStatus
        
        [System.Windows.MessageBox]::Show(
            "Disconnected from Exchange Online.",
            "Disconnected",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
        
    } catch {
        Write-Log "Disconnect error: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show(
            "Error disconnecting: $($_.Exception.Message)",
            "Disconnect Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
    }
})

$syncHash.MailboxButton.Add_Click({
    # Check if connected
    $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
        $result = [System.Windows.MessageBox]::Show(
            "You are not connected to Exchange Online.`n`nWould you like to connect now?",
            "Connection Required",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            $syncHash.ConnectButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            
            # Wait a moment for connection to establish, then recheck
            Start-Sleep -Seconds 2
            $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
            
            if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
                Write-Log "Connection was not established. Please try again."
                return
            }
        } else {
            return
        }
    }
    
    Write-Log "Opening Mailbox Permissions window..."
    
    function Resolve-UserDisplayName {
        param($Identity)
        
        $identityStr = $Identity.ToString()
        
        if ($identityStr -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
            try {
                $group = Get-DistributionGroup -Identity $identityStr -ErrorAction SilentlyContinue
                if ($group) { return $group.DisplayName }
                
                $group = Get-Group -Identity $identityStr -ErrorAction SilentlyContinue
                if ($group) { return $group.DisplayName }
                
                $recipient = Get-Recipient -Identity $identityStr -ErrorAction SilentlyContinue
                if ($recipient) { return $recipient.DisplayName }
            } catch {
                Write-Log "Could not resolve GUID: $identityStr"
            }
        }
        
        return $identityStr
    }
    
    function Get-CombinedMailboxPermissions {
    param($MailboxIdentity)
    
    try {
        Write-Log "Retrieving Full Access permissions for $MailboxIdentity"
        $fullAccessPerms = @(Get-MailboxPermission -Identity $MailboxIdentity -ErrorAction Stop | 
            Where-Object {$_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-5-*" -and $_.IsInherited -eq $false -and $_.AccessRights -contains "FullAccess"})
        
        Write-Log "Retrieving Send As permissions for $MailboxIdentity"
        $sendAsPerms = @(Get-RecipientPermission -Identity $MailboxIdentity -ErrorAction Stop | 
            Where-Object {$_.Trustee -notlike "NT AUTHORITY\*" -and $_.Trustee -notlike "S-1-5-*" -and $_.AccessRights -contains "SendAs"})
        
        $allUsers = @{}
        
        # Process Full Access permissions
        if ($fullAccessPerms.Count -gt 0) {
            foreach ($perm in $fullAccessPerms) {
                try {
                    $userKey = $perm.User.ToString()
                    
                    # Skip if null or empty
                    if ([string]::IsNullOrWhiteSpace($userKey)) {
                        Write-Log "Warning: Skipping null/empty Full Access user"
                        continue
                    }
                    
                    $displayName = Resolve-UserDisplayName -Identity $userKey
                    
                    if (-not $allUsers.ContainsKey($userKey)) {
                        $allUsers[$userKey] = [PSCustomObject]@{
                            User = $displayName
                            UserIdentity = $userKey
                            HasFullAccess = $true
                            HasSendAs = $false
                        }
                    } else {
                        $allUsers[$userKey].HasFullAccess = $true
                    }
                } catch {
                    Write-Log "Warning: Could not process Full Access permission for user: $($_.Exception.Message)"
                    continue
                }
            }
        }
        
        # Process Send As permissions
        if ($sendAsPerms.Count -gt 0) {
            foreach ($perm in $sendAsPerms) {
                try {
                    $userKey = $perm.Trustee.ToString()
                    
                    # Skip if null or empty
                    if ([string]::IsNullOrWhiteSpace($userKey)) {
                        Write-Log "Warning: Skipping null/empty Send As trustee"
                        continue
                    }
                    
                    $displayName = Resolve-UserDisplayName -Identity $userKey
                    
                    if (-not $allUsers.ContainsKey($userKey)) {
                        $allUsers[$userKey] = [PSCustomObject]@{
                            User = $displayName
                            UserIdentity = $userKey
                            HasFullAccess = $false
                            HasSendAs = $true
                        }
                    } else {
                        $allUsers[$userKey].HasSendAs = $true
                    }
                } catch {
                    Write-Log "Warning: Could not process Send As permission for trustee: $($_.Exception.Message)"
                    continue
                }
            }
        }
        
        # Return sorted results, or empty array if no permissions
        if ($allUsers.Count -gt 0) {
            $result = @($allUsers.Values | Sort-Object User)
            return ,$result  # Comma forces PowerShell to return as array
        } else {
            Write-Log "No delegated permissions found for $MailboxIdentity"
            return @()
        }
        
    } catch {
        Write-Log "Error in Get-CombinedMailboxPermissions: $($_.Exception.Message)"
        throw
    }
}
    
    [xml]$MailboxXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Mailbox Permissions Management" 
        Height="600" 
        Width="700" 
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <TabControl Margin="10">
            <TabItem Header="Add Permission">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Margin="0,0,0,15">
                        <TextBlock Text="Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="MbxAddMailboxBox" Height="25" Padding="5"/>
                    </StackPanel>
                    
                    <StackPanel Grid.Row="1" Margin="0,0,0,15">
                        <TextBlock Text="User Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="MbxAddUserBox" Height="25" Padding="5"/>
                    </StackPanel>
                    
                    <StackPanel Grid.Row="2" Margin="0,0,0,15">
                        <TextBlock Text="Access Rights:" FontWeight="Bold" Margin="0,0,0,10"/>
                        <CheckBox x:Name="MbxFullAccessCheck" Content="Full Access" Margin="0,0,0,8" FontSize="13"/>
                        <CheckBox x:Name="MbxSendAsCheck" Content="Send As" Margin="0,0,0,8" FontSize="13"/>
                        <TextBlock Text="(Select one or both)" FontStyle="Italic" FontSize="11" Foreground="Gray" Margin="0,5,0,0"/>
                    </StackPanel>
                    
                    <Border Grid.Row="3" Background="#f8f9fa" BorderBrush="#dee2e6" BorderThickness="1" Padding="10" Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Margin="0,0,0,8">
                                <Bold>Full Access:</Bold> User can open and view the mailbox contents.
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11">
                                <Bold>Send As:</Bold> User can send emails as if they were the mailbox owner.
                            </TextBlock>
                        </StackPanel>
                    </Border>
                    
                    <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="MbxAddButton" Content="Add Permission" Width="120" Height="30" Margin="0,0,10,0" Background="#28a745" Foreground="White" FontWeight="Bold"/>
                        <Button x:Name="MbxAddCancelButton" Content="Cancel" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="View/Edit">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Margin="0,0,0,15">
                        <TextBlock Text="Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <DockPanel>
                            <Button x:Name="MbxLoadPermissionsButton" Content="Load" Width="80" Height="25" Margin="10,0,0,0" DockPanel.Dock="Right" Background="#007bff" Foreground="White" FontWeight="Bold"/>
                            <TextBox x:Name="MbxViewMailboxBox" Height="25" Padding="5"/>
                        </DockPanel>
                    </StackPanel>
                    
                    <Border Grid.Row="1" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15">
                        <DataGrid x:Name="MbxPermissionsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="User" Binding="{Binding User}" Width="2*"/>
                                <DataGridCheckBoxColumn Header="Full Access" Binding="{Binding HasFullAccess}" Width="*"/>
                                <DataGridCheckBoxColumn Header="Send As" Binding="{Binding HasSendAs}" Width="*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Border>
                    
                    <GroupBox Grid.Row="2" Header="Edit Selected" Padding="10" Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock x:Name="MbxEditUserLabel" Text="No user selected" Margin="0,0,0,10"/>
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                <CheckBox x:Name="MbxEditFullAccessCheck" Content="Full Access" Margin="0,0,20,0" IsEnabled="False"/>
                                <CheckBox x:Name="MbxEditSendAsCheck" Content="Send As" IsEnabled="False"/>
                            </StackPanel>
                            <Button x:Name="MbxUpdateButton" Content="Update Permissions" Width="140" Height="25" HorizontalAlignment="Left" Background="#ffc107" Foreground="Black" FontWeight="Bold" IsEnabled="False"/>
                        </StackPanel>
                    </GroupBox>
                    
                    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="MbxExportToExcelButton" Content="Export to Excel" Width="120" Height="30" Margin="0,0,10,0" Background="#17a2b8" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="MbxViewCloseButton" Content="Close" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="Remove">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Margin="0,0,0,15">
                        <TextBlock Text="Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <DockPanel>
                            <Button x:Name="MbxLoadRemovePermissionsButton" Content="Load" Width="80" Height="25" Margin="10,0,0,0" DockPanel.Dock="Right" Background="#007bff" Foreground="White" FontWeight="Bold"/>
                            <TextBox x:Name="MbxRemoveMailboxBox" Height="25" Padding="5"/>
                        </DockPanel>
                    </StackPanel>
                    
                    <Border Grid.Row="1" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15">
                        <DataGrid x:Name="MbxRemovePermissionsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="User" Binding="{Binding User}" Width="2*"/>
                                <DataGridCheckBoxColumn Header="Full Access" Binding="{Binding HasFullAccess}" Width="*"/>
                                <DataGridCheckBoxColumn Header="Send As" Binding="{Binding HasSendAs}" Width="*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Border>
                    
                    <StackPanel Grid.Row="2">
                        <TextBlock Text="Select which permissions to remove:" FontWeight="Bold" Margin="0,0,0,10"/>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                            <CheckBox x:Name="MbxRemoveFullAccessCheck" Content="Remove Full Access" Margin="0,0,20,0" IsEnabled="False"/>
                            <CheckBox x:Name="MbxRemoveSendAsCheck" Content="Remove Send As" IsEnabled="False"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                            <Button x:Name="MbxRemoveButton" Content="Remove Selected" Width="130" Height="30" Margin="0,0,10,0" Background="#dc3545" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                            <Button x:Name="MbxRemoveCloseButton" Content="Close" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
                        </StackPanel>
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@
    
    $mbxReader = New-Object System.Xml.XmlNodeReader $MailboxXAML
    $MbxWindow = [Windows.Markup.XamlReader]::Load($mbxReader)
    $MbxWindow.Owner = $syncHash.Window
    
    $MbxAddMailboxBox = $MbxWindow.FindName("MbxAddMailboxBox")
    $MbxAddUserBox = $MbxWindow.FindName("MbxAddUserBox")
    $MbxFullAccessCheck = $MbxWindow.FindName("MbxFullAccessCheck")
    $MbxSendAsCheck = $MbxWindow.FindName("MbxSendAsCheck")
    $MbxAddButton = $MbxWindow.FindName("MbxAddButton")
    $MbxAddCancelButton = $MbxWindow.FindName("MbxAddCancelButton")
    
    $MbxViewMailboxBox = $MbxWindow.FindName("MbxViewMailboxBox")
    $MbxLoadPermissionsButton = $MbxWindow.FindName("MbxLoadPermissionsButton")
    $MbxPermissionsGrid = $MbxWindow.FindName("MbxPermissionsGrid")
    $MbxEditUserLabel = $MbxWindow.FindName("MbxEditUserLabel")
    $MbxEditFullAccessCheck = $MbxWindow.FindName("MbxEditFullAccessCheck")
    $MbxEditSendAsCheck = $MbxWindow.FindName("MbxEditSendAsCheck")
    $MbxUpdateButton = $MbxWindow.FindName("MbxUpdateButton")
    $MbxExportToExcelButton = $MbxWindow.FindName("MbxExportToExcelButton")
    $MbxViewCloseButton = $MbxWindow.FindName("MbxViewCloseButton")
    
    $MbxRemoveMailboxBox = $MbxWindow.FindName("MbxRemoveMailboxBox")
    $MbxLoadRemovePermissionsButton = $MbxWindow.FindName("MbxLoadRemovePermissionsButton")
    $MbxRemovePermissionsGrid = $MbxWindow.FindName("MbxRemovePermissionsGrid")
    $MbxRemoveFullAccessCheck = $MbxWindow.FindName("MbxRemoveFullAccessCheck")
    $MbxRemoveSendAsCheck = $MbxWindow.FindName("MbxRemoveSendAsCheck")
    $MbxRemoveButton = $MbxWindow.FindName("MbxRemoveButton")
    $MbxRemoveCloseButton = $MbxWindow.FindName("MbxRemoveCloseButton")
    
    $MbxPermissionsGrid.Add_MouseDoubleClick({
        if ($MbxPermissionsGrid.SelectedItem) {
            $selectedUser = $MbxPermissionsGrid.SelectedItem.UserIdentity
            Show-ADPropertiesWindow -Identity $selectedUser -Owner $MbxWindow
        }
    })
    
    $MbxRemovePermissionsGrid.Add_MouseDoubleClick({
        if ($MbxRemovePermissionsGrid.SelectedItem) {
            $selectedUser = $MbxRemovePermissionsGrid.SelectedItem.UserIdentity
            Show-ADPropertiesWindow -Identity $selectedUser -Owner $MbxWindow
        }
    })
    
    $MbxAddCancelButton.Add_Click({ $MbxWindow.Close() })
    
    $MbxAddButton.Add_Click({
        $mailbox = $MbxAddMailboxBox.Text.Trim()
        $user = $MbxAddUserBox.Text.Trim()
        $addFullAccess = $MbxFullAccessCheck.IsChecked
        $addSendAs = $MbxSendAsCheck.IsChecked
        
        if ([string]::IsNullOrWhiteSpace($mailbox) -or [string]::IsNullOrWhiteSpace($user)) {
            [System.Windows.MessageBox]::Show("Please enter both mailbox and user email addresses", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not $addFullAccess -and -not $addSendAs) {
            [System.Windows.MessageBox]::Show("Please select at least one permission type", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        try {
            $MbxAddButton.IsEnabled = $false
            $successMessages = @()
            
            if ($addFullAccess) {
                Write-Log "Adding Full Access for $user on $mailbox"
                Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping $false -ErrorAction Stop
                $successMessages += "Full Access"
            }
            
            if ($addSendAs) {
                Write-Log "Adding Send As for $user on $mailbox"
                Add-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                $successMessages += "Send As"
            }
            
            Write-Log "Successfully added permissions: $($successMessages -join ', ')"
            [System.Windows.MessageBox]::Show("Permissions added successfully!`n`n$($successMessages -join ', ')", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            
            $MbxAddMailboxBox.Clear()
            $MbxAddUserBox.Clear()
            $MbxFullAccessCheck.IsChecked = $false
            $MbxSendAsCheck.IsChecked = $false
            
        } catch {
            Write-Log "Error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } finally {
            $MbxAddButton.IsEnabled = $true
        }
    })
    
    $MbxLoadPermissionsButton.Add_Click({
        $mailbox = $MbxViewMailboxBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($mailbox)) { return }
        
        try {
            $MbxLoadPermissionsButton.IsEnabled = $false
            Write-Log "Loading permissions for $mailbox"
            
            $perms = Get-CombinedMailboxPermissions -MailboxIdentity $mailbox
            $MbxPermissionsGrid.ItemsSource = $perms
            
            if ($perms -and $perms.Count -gt 0) {
                $MbxExportToExcelButton.IsEnabled = $true
            } else {
                $MbxExportToExcelButton.IsEnabled = $false
            }
            
            Write-Log "Loaded $($perms.Count) user permissions"
        } catch {
            Write-Log "Error loading permissions: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } finally {
            $MbxLoadPermissionsButton.IsEnabled = $true
        }
    })
    
    $MbxPermissionsGrid.Add_SelectionChanged({
        if ($MbxPermissionsGrid.SelectedItem) {
            $sel = $MbxPermissionsGrid.SelectedItem
            $MbxEditUserLabel.Text = "Editing: $($sel.User)"
            $MbxEditFullAccessCheck.IsEnabled = $true
            $MbxEditSendAsCheck.IsEnabled = $true
            $MbxEditFullAccessCheck.IsChecked = $sel.HasFullAccess
            $MbxEditSendAsCheck.IsChecked = $sel.HasSendAs
            $MbxUpdateButton.IsEnabled = $true
        }
    })
    
    $MbxUpdateButton.Add_Click({
        $mailbox = $MbxViewMailboxBox.Text.Trim()
        $sel = $MbxPermissionsGrid.SelectedItem
        $newFullAccess = $MbxEditFullAccessCheck.IsChecked
        $newSendAs = $MbxEditSendAsCheck.IsChecked
        
        if ($null -eq $sel) { return }
        
        if (-not $newFullAccess -and -not $newSendAs) {
            [System.Windows.MessageBox]::Show("At least one permission must be selected. Use the Remove tab to remove all permissions.", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $result = [System.Windows.MessageBox]::Show("Update permissions for $($sel.User)?", "Confirm", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                $MbxUpdateButton.IsEnabled = $false
                
                if ($sel.HasFullAccess -and -not $newFullAccess) {
                    Write-Log "Removing Full Access for $($sel.User)"
                    Remove-MailboxPermission -Identity $mailbox -User $sel.UserIdentity -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
                } elseif (-not $sel.HasFullAccess -and $newFullAccess) {
                    Write-Log "Adding Full Access for $($sel.User)"
                    Add-MailboxPermission -Identity $mailbox -User $sel.UserIdentity -AccessRights FullAccess -InheritanceType All -AutoMapping $false -ErrorAction Stop
                }
                
                if ($sel.HasSendAs -and -not $newSendAs) {
                    Write-Log "Removing Send As for $($sel.User)"
                    Remove-RecipientPermission -Identity $mailbox -Trustee $sel.UserIdentity -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                } elseif (-not $sel.HasSendAs -and $newSendAs) {
                    Write-Log "Adding Send As for $($sel.User)"
                    Add-RecipientPermission -Identity $mailbox -Trustee $sel.UserIdentity -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                }
                
                Write-Log "Successfully updated permissions"
                [System.Windows.MessageBox]::Show("Permissions updated successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                $MbxLoadPermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
                
            } catch {
                Write-Log "Error: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            } finally {
                $MbxUpdateButton.IsEnabled = $true
            }
        }
    })
	
	$MbxExportToExcelButton.Add_Click({
        $mailbox = $MbxViewMailboxBox.Text.Trim()
        $permissions = $MbxPermissionsGrid.ItemsSource
        
        if ($null -eq $permissions -or $permissions.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No permissions to export. Please load permissions first.", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            [System.Windows.MessageBox]::Show("ImportExcel module is not installed. Cannot export to Excel.`n`nPlease run: Install-Module -Name ImportExcel", "Module Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Write-Log "Export failed: ImportExcel module not installed"
            return
        }
        
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.Title = "Save Mailbox Permissions Report"
            $saveDialog.FileName = "Mailbox_Permissions_$($mailbox.Replace('@','_'))_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $MbxExportToExcelButton.IsEnabled = $false
                Write-Log "Exporting permissions to Excel: $excelPath"
                
                $exportData = @()
                foreach ($perm in $permissions) {
                    $exportData += [PSCustomObject]@{
                        'Mailbox' = $mailbox
                        'Delegate User' = $perm.User
                        'Full Access' = if ($perm.HasFullAccess) { "Yes" } else { "No" }
                        'Send As' = if ($perm.HasSendAs) { "Yes" } else { "No" }
                        'Export Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                $exportData | Export-Excel -Path $excelPath `
                    -WorksheetName "Mailbox Permissions" `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle Medium1 `
                    -TableName "MailboxPermissions"
                
                Write-Log "Successfully exported $($exportData.Count) permissions to Excel"
                
                [System.Windows.MessageBox]::Show(
                    "Mailbox permissions exported successfully!`n`nFile: $excelPath",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                ) | Out-Null
                
                # Open folder location in Explorer
                $folderPath = Split-Path $excelPath -Parent
                Invoke-Item $folderPath
            } else {
                Write-Log "Export cancelled by user"
            }
            
        } catch {
            Write-Log "Export error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error exporting to Excel:`n`n$($_.Exception.Message)",
                "Export Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        } finally {
            $MbxExportToExcelButton.IsEnabled = $true
        }
    })
    
    $MbxViewCloseButton.Add_Click({ $MbxWindow.Close() })
    
    $MbxLoadRemovePermissionsButton.Add_Click({
        $mailbox = $MbxRemoveMailboxBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($mailbox)) { return }
        
        try {
            $MbxLoadRemovePermissionsButton.IsEnabled = $false
            Write-Log "Loading permissions for $mailbox"
            
            $perms = Get-CombinedMailboxPermissions -MailboxIdentity $mailbox
            $MbxRemovePermissionsGrid.ItemsSource = $perms
            
            Write-Log "Loaded $($perms.Count) user permissions"
        } catch {
            Write-Log "Error loading permissions: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } finally {
            $MbxLoadRemovePermissionsButton.IsEnabled = $true
        }
    })
    
    $MbxRemovePermissionsGrid.Add_SelectionChanged({
        if ($MbxRemovePermissionsGrid.SelectedItem) {
            $sel = $MbxRemovePermissionsGrid.SelectedItem
            $MbxRemoveFullAccessCheck.IsEnabled = $sel.HasFullAccess
            $MbxRemoveSendAsCheck.IsEnabled = $sel.HasSendAs
            $MbxRemoveFullAccessCheck.IsChecked = $false
            $MbxRemoveSendAsCheck.IsChecked = $false
            $MbxRemoveButton.IsEnabled = $true
        } else {
            $MbxRemoveFullAccessCheck.IsEnabled = $false
            $MbxRemoveSendAsCheck.IsEnabled = $false
            $MbxRemoveButton.IsEnabled = $false
        }
    })
    
    $MbxRemoveButton.Add_Click({
        $mailbox = $MbxRemoveMailboxBox.Text.Trim()
        $sel = $MbxRemovePermissionsGrid.SelectedItem
        $removeFullAccess = $MbxRemoveFullAccessCheck.IsChecked
        $removeSendAs = $MbxRemoveSendAsCheck.IsChecked
        
        if ($null -eq $sel) { return }
        
        if (-not $removeFullAccess -and -not $removeSendAs) {
            [System.Windows.MessageBox]::Show("Please select at least one permission type to remove", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $permsToRemove = @()
        if ($removeFullAccess) { $permsToRemove += "Full Access" }
        if ($removeSendAs) { $permsToRemove += "Send As" }
        
        $result = [System.Windows.MessageBox]::Show(
            "Remove the following permissions for $($sel.User)?`n`n$($permsToRemove -join ', ')",
            "Confirm",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                $MbxRemoveButton.IsEnabled = $false
                $removedPerms = @()
                
                if ($removeFullAccess) {
                    Write-Log "Removing Full Access for $($sel.User)"
                    Remove-MailboxPermission -Identity $mailbox -User $sel.UserIdentity -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
                    $removedPerms += "Full Access"
                }
                
                if ($removeSendAs) {
                    Write-Log "Removing Send As for $($sel.User)"
                    Remove-RecipientPermission -Identity $mailbox -Trustee $sel.UserIdentity -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                    $removedPerms += "Send As"
                }
                
                Write-Log "Successfully removed: $($removedPerms -join ', ')"
                [System.Windows.MessageBox]::Show("Permissions removed successfully!`n`n$($removedPerms -join ', ')", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                $MbxLoadRemovePermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
                
            } catch {
                Write-Log "Error: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            } finally {
                $MbxRemoveButton.IsEnabled = $true
            }
        }
    })
    
    $MbxRemoveCloseButton.Add_Click({ $MbxWindow.Close() })
	
	# Add Enter key support for Mailbox Add tab
	$MbxAddMailboxBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$MbxAddUserBox.Focus()
		}
	})

	$MbxAddUserBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$MbxAddButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
		}
	})

	# Add Enter key support for Mailbox View/Edit tab
	$MbxViewMailboxBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$MbxLoadPermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
		}
	})

	# Add Enter key support for Mailbox Remove tab
	$MbxRemoveMailboxBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$MbxLoadRemovePermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
		}
	})
    
    $MbxWindow.ShowDialog() | Out-Null
})

$syncHash.CalendarButton.Add_Click({
    # Check if connected
    $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
        $result = [System.Windows.MessageBox]::Show(
            "You are not connected to Exchange Online.`n`nWould you like to connect now?",
            "Connection Required",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            $syncHash.ConnectButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            
            # Wait a moment for connection to establish, then recheck
            Start-Sleep -Seconds 2
            $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
            
            if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
                Write-Log "Connection was not established. Please try again."
                return
            }
        } else {
            return
        }
    }
    
    Write-Log "Opening Calendar Permissions window..."
    
    [xml]$CalendarXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Calendar Permissions Management" 
        Height="550" 
        Width="650" 
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <TabControl Margin="10">
            <TabItem Header="Add Permission">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Margin="0,0,0,15">
                        <TextBlock Text="Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="AddMailboxBox" Height="25" Padding="5"/>
                    </StackPanel>
                    
                    <StackPanel Grid.Row="1" Margin="0,0,0,15">
                        <TextBlock Text="User Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="AddUserBox" Height="25" Padding="5"/>
                    </StackPanel>
                    
                    <StackPanel Grid.Row="2" Margin="0,0,0,15">
                        <TextBlock Text="Access Rights:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <ComboBox x:Name="AddPermissionCombo" Height="25">
                            <ComboBoxItem Content="AvailabilityOnly" Tag="AvailabilityOnly"/>
                            <ComboBoxItem Content="LimitedDetails" Tag="LimitedDetails"/>
                            <ComboBoxItem Content="Reviewer" Tag="Reviewer"/>
                            <ComboBoxItem Content="Contributor" Tag="Contributor"/>
                            <ComboBoxItem Content="Author" Tag="Author"/>
                            <ComboBoxItem Content="Editor" Tag="Editor"/>
                            <ComboBoxItem Content="Owner" Tag="Owner"/>
                        </ComboBox>
                    </StackPanel>
                    
                    <Border Grid.Row="3" Background="#f8f9fa" BorderBrush="#dee2e6" BorderThickness="1" Padding="10" Margin="0,0,0,15">
                        <TextBlock TextWrapping="Wrap" FontSize="11" Text="Select the appropriate permission level for calendar access."/>
                    </Border>
                    
                    <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="AddButton" Content="Add Permission" Width="120" Height="30" Margin="0,0,10,0" Background="#28a745" Foreground="White" FontWeight="Bold"/>
                        <Button x:Name="AddCancelButton" Content="Cancel" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="View/Edit">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Margin="0,0,0,15">
                        <TextBlock Text="Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <DockPanel>
                            <Button x:Name="LoadPermissionsButton" Content="Load" Width="80" Height="25" Margin="10,0,0,0" DockPanel.Dock="Right" Background="#007bff" Foreground="White" FontWeight="Bold"/>
                            <TextBox x:Name="ViewMailboxBox" Height="25" Padding="5"/>
                        </DockPanel>
                    </StackPanel>
                    
                    <Border Grid.Row="1" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15">
                        <DataGrid x:Name="PermissionsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="User" Binding="{Binding User}" Width="*"/>
                                <DataGridTextColumn Header="Access Rights" Binding="{Binding AccessRights}" Width="*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Border>
                    
                    <GroupBox Grid.Row="2" Header="Edit Selected" Padding="10" Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock x:Name="EditUserLabel" Text="No permission selected" Margin="0,0,0,10"/>
                            <DockPanel>
                                <Button x:Name="UpdateButton" Content="Update" Width="80" Height="25" Margin="10,0,0,0" DockPanel.Dock="Right" Background="#ffc107" Foreground="Black" FontWeight="Bold" IsEnabled="False"/>
                                <ComboBox x:Name="EditPermissionCombo" Height="25" IsEnabled="False">
                                    <ComboBoxItem Content="AvailabilityOnly" Tag="AvailabilityOnly"/>
                                    <ComboBoxItem Content="LimitedDetails" Tag="LimitedDetails"/>
                                    <ComboBoxItem Content="Reviewer" Tag="Reviewer"/>
                                    <ComboBoxItem Content="Contributor" Tag="Contributor"/>
                                    <ComboBoxItem Content="Author" Tag="Author"/>
                                    <ComboBoxItem Content="Editor" Tag="Editor"/>
                                    <ComboBoxItem Content="Owner" Tag="Owner"/>
                                </ComboBox>
                            </DockPanel>
                        </StackPanel>
                    </GroupBox>
                    
                    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="ExportToExcelButton" Content="Export to Excel" Width="120" Height="30" Margin="0,0,10,0" Background="#17a2b8" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="ViewCloseButton" Content="Close" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="Remove">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Margin="0,0,0,15">
                        <TextBlock Text="Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <DockPanel>
                            <Button x:Name="LoadRemovePermissionsButton" Content="Load" Width="80" Height="25" Margin="10,0,0,0" DockPanel.Dock="Right" Background="#007bff" Foreground="White" FontWeight="Bold"/>
                            <TextBox x:Name="RemoveMailboxBox" Height="25" Padding="5"/>
                        </DockPanel>
                    </StackPanel>
                    
                    <Border Grid.Row="1" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15">
                        <DataGrid x:Name="RemovePermissionsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="User" Binding="{Binding User}" Width="*"/>
                                <DataGridTextColumn Header="Access Rights" Binding="{Binding AccessRights}" Width="*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Border>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="RemoveButton" Content="Remove Selected" Width="130" Height="30" Margin="0,0,10,0" Background="#dc3545" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="RemoveCloseButton" Content="Close" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@
    
    $calReader = New-Object System.Xml.XmlNodeReader $CalendarXAML
    $CalWindow = [Windows.Markup.XamlReader]::Load($calReader)
    $CalWindow.Owner = $syncHash.Window
    
    $AddMailboxBox = $CalWindow.FindName("AddMailboxBox")
    $AddUserBox = $CalWindow.FindName("AddUserBox")
    $AddPermissionCombo = $CalWindow.FindName("AddPermissionCombo")
    $AddButton = $CalWindow.FindName("AddButton")
    $AddCancelButton = $CalWindow.FindName("AddCancelButton")
    
    $ViewMailboxBox = $CalWindow.FindName("ViewMailboxBox")
    $LoadPermissionsButton = $CalWindow.FindName("LoadPermissionsButton")
    $PermissionsGrid = $CalWindow.FindName("PermissionsGrid")
    $EditUserLabel = $CalWindow.FindName("EditUserLabel")
    $EditPermissionCombo = $CalWindow.FindName("EditPermissionCombo")
    $UpdateButton = $CalWindow.FindName("UpdateButton")
    $ExportToExcelButton = $CalWindow.FindName("ExportToExcelButton")
    $ViewCloseButton = $CalWindow.FindName("ViewCloseButton")
    
    $RemoveMailboxBox = $CalWindow.FindName("RemoveMailboxBox")
    $LoadRemovePermissionsButton = $CalWindow.FindName("LoadRemovePermissionsButton")
    $RemovePermissionsGrid = $CalWindow.FindName("RemovePermissionsGrid")
    $RemoveButton = $CalWindow.FindName("RemoveButton")
    $RemoveCloseButton = $CalWindow.FindName("RemoveCloseButton")
    
    $AddPermissionCombo.SelectedIndex = 2
    
    $PermissionsGrid.Add_MouseDoubleClick({
        if ($PermissionsGrid.SelectedItem) {
            $selectedUser = $PermissionsGrid.SelectedItem.User
            Show-ADPropertiesWindow -Identity $selectedUser -Owner $CalWindow
        }
    })
    
    $RemovePermissionsGrid.Add_MouseDoubleClick({
        if ($RemovePermissionsGrid.SelectedItem) {
            $selectedUser = $RemovePermissionsGrid.SelectedItem.User
            Show-ADPropertiesWindow -Identity $selectedUser -Owner $CalWindow
        }
    })
    
    $AddCancelButton.Add_Click({ $CalWindow.Close() })
    
    $AddButton.Add_Click({
        $mailbox = $AddMailboxBox.Text.Trim()
        $user = $AddUserBox.Text.Trim()
        $perm = $AddPermissionCombo.SelectedItem.Tag
        
        if ([string]::IsNullOrWhiteSpace($mailbox) -or [string]::IsNullOrWhiteSpace($user) -or $null -eq $perm) {
            [System.Windows.MessageBox]::Show("Please fill all fields", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        try {
            $AddButton.IsEnabled = $false
            Write-Log "Adding $perm for $user on ${mailbox}:\Calendar"
            Add-MailboxFolderPermission -Identity "${mailbox}:\Calendar" -User $user -AccessRights $perm -ErrorAction Stop
            Write-Log "Successfully added permission"
            [System.Windows.MessageBox]::Show("Permission added successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            $AddMailboxBox.Clear(); $AddUserBox.Clear(); $AddPermissionCombo.SelectedIndex = 2
        } catch {
            Write-Log "Error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } finally {
            $AddButton.IsEnabled = $true
        }
    })
    
    $LoadPermissionsButton.Add_Click({
        $mailbox = $ViewMailboxBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($mailbox)) { return }
        
        try {
            $LoadPermissionsButton.IsEnabled = $false
            Write-Log "Loading permissions for $mailbox"
            
            # Get all permissions first
            $allPerms = Get-MailboxFolderPermission -Identity "${mailbox}:\Calendar" -ErrorAction Stop
            Write-Log "DEBUG: Retrieved $($allPerms.Count) total permissions"
            
            # Filter and format properly
            $perms = @()
            foreach ($perm in $allPerms) {
                $userName = $perm.User.DisplayName
                
                # Skip Default and Anonymous
                if ($userName -eq "Default" -or $userName -eq "Anonymous") {
                    Write-Log "DEBUG: Skipping $userName"
                    continue
                }
                
                Write-Log "DEBUG: Adding user: $userName with rights: $($perm.AccessRights)"
                
                $perms += [PSCustomObject]@{
                    User = $userName
                    AccessRights = ($perm.AccessRights -join ", ")
                }
            }
            
            Write-Log "Found $($perms.Count) delegated permissions (after filtering)"
            
            # Bind to grid
            if ($perms.Count -gt 0) {
                $PermissionsGrid.ItemsSource = $perms
                $ExportToExcelButton.IsEnabled = $true
            } else {
                $PermissionsGrid.ItemsSource = $null
                $ExportToExcelButton.IsEnabled = $false
                [System.Windows.MessageBox]::Show("No delegated calendar permissions found (only Default/Anonymous).", "No Permissions", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
            
        } catch {
            Write-Log "Error loading permissions: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } finally {
            $LoadPermissionsButton.IsEnabled = $true
        }
    })
    
    $PermissionsGrid.Add_SelectionChanged({
        if ($PermissionsGrid.SelectedItem) {
            $sel = $PermissionsGrid.SelectedItem
            $EditUserLabel.Text = "Editing: $($sel.User)"
            $EditPermissionCombo.IsEnabled = $true
            $UpdateButton.IsEnabled = $true
            $currentRight = ($sel.AccessRights -split ",")[0].Trim()
            for ($i = 0; $i -lt $EditPermissionCombo.Items.Count; $i++) {
                if ($EditPermissionCombo.Items[$i].Tag -eq $currentRight) {
                    $EditPermissionCombo.SelectedIndex = $i
                    break
                }
            }
        }
    })
    
    $UpdateButton.Add_Click({
        $mailbox = $ViewMailboxBox.Text.Trim()
        $sel = $PermissionsGrid.SelectedItem
        $newPerm = $EditPermissionCombo.SelectedItem.Tag
        
        if ($null -eq $sel -or $null -eq $newPerm) { return }
        
        $result = [System.Windows.MessageBox]::Show("Update permission for $($sel.User) to $newPerm?", "Confirm", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                $UpdateButton.IsEnabled = $false
                Write-Log "Updating permission for $($sel.User) to $newPerm"
                Set-MailboxFolderPermission -Identity "${mailbox}:\Calendar" -User $sel.User -AccessRights $newPerm -ErrorAction Stop
                Write-Log "Successfully updated"
                [System.Windows.MessageBox]::Show("Permission updated!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                $LoadPermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            } catch {
                [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            } finally {
                $UpdateButton.IsEnabled = $true
            }
        }
    })
	
	$ExportToExcelButton.Add_Click({
        $mailbox = $ViewMailboxBox.Text.Trim()
        $permissions = $PermissionsGrid.ItemsSource
        
        if ($null -eq $permissions -or $permissions.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No permissions to export. Please load permissions first.", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            [System.Windows.MessageBox]::Show("ImportExcel module is not installed. Cannot export to Excel.`n`nPlease run: Install-Module -Name ImportExcel", "Module Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Write-Log "Export failed: ImportExcel module not installed"
            return
        }
        
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.Title = "Save Calendar Permissions Report"
            $saveDialog.FileName = "Calendar_Permissions_$($mailbox.Replace('@','_'))_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $ExportToExcelButton.IsEnabled = $false
                Write-Log "Exporting permissions to Excel: $excelPath"
                
                $exportData = @()
                foreach ($perm in $permissions) {
                    $exportData += [PSCustomObject]@{
                        'Mailbox' = $mailbox
                        'Delegate User' = $perm.User
                        'Access Rights' = $perm.AccessRights
                        'Export Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                $exportData | Export-Excel -Path $excelPath `
                    -WorksheetName "Calendar Permissions" `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle Medium1 `
                    -TableName "CalendarPermissions"
                
                Write-Log "Successfully exported $($exportData.Count) permissions to Excel"
                
                [System.Windows.MessageBox]::Show(
                    "Calendar permissions exported successfully!`n`nFile: $excelPath",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                ) | Out-Null
                
                # Open folder location in Explorer
                $folderPath = Split-Path $excelPath -Parent
                Invoke-Item $folderPath
                
            } else {
                Write-Log "Export cancelled by user"
            }
            
        } catch {
            Write-Log "Export error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error exporting to Excel:`n`n$($_.Exception.Message)",
                "Export Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        } finally {
            $ExportToExcelButton.IsEnabled = $true
        }
    })
    
    $ViewCloseButton.Add_Click({ $CalWindow.Close() })
    
    $LoadRemovePermissionsButton.Add_Click({
        $mailbox = $RemoveMailboxBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($mailbox)) { return }
        
        try {
            $LoadRemovePermissionsButton.IsEnabled = $false
            Write-Log "Loading permissions for $mailbox"
            
            # Get all permissions first
            $allPerms = Get-MailboxFolderPermission -Identity "${mailbox}:\Calendar" -ErrorAction Stop
            Write-Log "DEBUG: Retrieved $($allPerms.Count) total permissions"
            
            # Filter and format properly
            $perms = @()
            foreach ($perm in $allPerms) {
                $userName = $perm.User.DisplayName
                
                # Skip Default and Anonymous
                if ($userName -eq "Default" -or $userName -eq "Anonymous") {
                    Write-Log "DEBUG: Skipping $userName"
                    continue
                }
                
                Write-Log "DEBUG: Adding user: $userName with rights: $($perm.AccessRights)"
                
                $perms += [PSCustomObject]@{
                    User = $userName
                    AccessRights = ($perm.AccessRights -join ", ")
                }
            }
            
            Write-Log "Found $($perms.Count) delegated permissions (after filtering)"
            
            # Bind to grid
            if ($perms.Count -gt 0) {
                $RemovePermissionsGrid.ItemsSource = $perms
            } else {
                $RemovePermissionsGrid.ItemsSource = $null
                [System.Windows.MessageBox]::Show("No delegated calendar permissions found (only Default/Anonymous).", "No Permissions", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
            
        } catch {
            Write-Log "Error loading permissions: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } finally {
            $LoadRemovePermissionsButton.IsEnabled = $true
        }
    })

    $RemovePermissionsGrid.Add_SelectionChanged({
        $RemoveButton.IsEnabled = ($null -ne $RemovePermissionsGrid.SelectedItem)
    })
    
    $RemoveButton.Add_Click({
        $mailbox = $RemoveMailboxBox.Text.Trim()
        $sel = $RemovePermissionsGrid.SelectedItem
        if ($null -eq $sel) { return }
        
        $result = [System.Windows.MessageBox]::Show("Remove permission for $($sel.User)?", "Confirm", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                $RemoveButton.IsEnabled = $false
                Write-Log "Removing permission for $($sel.User)"
                Remove-MailboxFolderPermission -Identity "${mailbox}:\Calendar" -User $sel.User -Confirm:$false -ErrorAction Stop
                Write-Log "Successfully removed"
                [System.Windows.MessageBox]::Show("Permission removed!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                $LoadRemovePermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            } catch {
                [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            } finally {
                $RemoveButton.IsEnabled = $true
            }
        }
    })
    
    $RemoveCloseButton.Add_Click({ $CalWindow.Close() })
	
	# Add Enter key support for Calendar Add tab
	$AddMailboxBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$AddUserBox.Focus()
		}
	})

	$AddUserBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$AddButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
		}
	})

	# Add Enter key support for Calendar View/Edit tab
	$ViewMailboxBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$LoadPermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
		}
	})

	# Add Enter key support for Calendar Remove tab
	$RemoveMailboxBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$LoadRemovePermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
		}
	})
    
    $CalWindow.ShowDialog() | Out-Null
})

$syncHash.GroupMembersButton.Add_Click({
    Write-Log "Opening AD Group Members window..."
    
    # Check if Active Directory module is available
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        [System.Windows.MessageBox]::Show(
            "Active Directory PowerShell module is not installed.`n`nThis feature requires the RSAT Active Directory module.`n`nPlease install it from:`nSettings > Apps > Optional Features > Add RSAT: Active Directory Domain Services",
            "Module Required",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        Write-Log "AD Group Members requires ActiveDirectory module"
        return
    }
    
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        [System.Windows.MessageBox]::Show(
            "Failed to load Active Directory module:`n`n$($_.Exception.Message)",
            "Module Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }
    
    [xml]$GroupMembersXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AD Group Members Management" 
        Height="550" 
        Width="700" 
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <GroupBox Grid.Row="0" Header="Group Information" Margin="15,15,15,10" Padding="15">
            <StackPanel>
                <TextBlock Text="Group Name, Email, or SAM Account Name:" FontWeight="Bold" Margin="0,0,0,5"/>
                <DockPanel Margin="0,0,0,10">
                    <Button x:Name="LoadGroupButton" 
                           Content="Load Members" 
                           Width="120" 
                           Height="30" 
                           Margin="10,0,0,0" 
                           DockPanel.Dock="Right" 
                           Background="#007bff" 
                           Foreground="White" 
                           FontWeight="Bold"/>
                    <TextBox x:Name="GroupIdentityBox" 
                            Height="30" 
                            Padding="5"
                            VerticalContentAlignment="Center"/>
                </DockPanel>
                
                <StackPanel x:Name="GroupInfoPanel" Visibility="Collapsed" Margin="0,10,0,0">
                    <Border Background="#e7f3ff" BorderBrush="#007bff" BorderThickness="1" Padding="10" CornerRadius="3">
                        <StackPanel>
                            <TextBlock x:Name="GroupNameText" FontWeight="Bold" Margin="0,0,0,5"/>
                            <TextBlock x:Name="GroupTypeText" FontSize="11" Foreground="#666" Margin="0,0,0,3"/>
                            <TextBlock x:Name="GroupEmailText" FontSize="11" Foreground="#666" Margin="0,0,0,3"/>
                            <TextBlock x:Name="GroupScopeText" FontSize="11" Foreground="#666" Margin="0,0,0,3"/>
                            <TextBlock x:Name="GroupMemberCountText" FontSize="11" Foreground="#666"/>
                        </StackPanel>
                    </Border>
                </StackPanel>
            </StackPanel>
        </GroupBox>
        
        <GroupBox Grid.Row="1" Header="Group Members" Margin="15,0,15,10" Padding="10">
            <Border BorderBrush="#dee2e6" BorderThickness="1">
                <DataGrid x:Name="MembersGrid" 
                         AutoGenerateColumns="False" 
                         IsReadOnly="True" 
                         SelectionMode="Extended"
                         AlternatingRowBackground="#f8f9fa">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="2*"/>
                        <DataGridTextColumn Header="Email Address" Binding="{Binding Email}" Width="2*"/>
                        <DataGridTextColumn Header="Object Type" Binding="{Binding ObjectClass}" Width="*"/>
                        <DataGridTextColumn Header="Title" Binding="{Binding Title}" Width="1.5*"/>
                        <DataGridTextColumn Header="Department" Binding="{Binding Department}" Width="1.5*"/>
                    </DataGrid.Columns>
                </DataGrid>
            </Border>
        </GroupBox>
        
        <Border Grid.Row="2" Background="#f8f9fa" BorderBrush="#dee2e6" BorderThickness="0,1,0,0" Padding="15">
            <DockPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Left" DockPanel.Dock="Left">
                    <TextBlock x:Name="StatusText" 
                              Text="Enter a group name or email to begin" 
                              VerticalAlignment="Center"
                              FontSize="11"
                              Foreground="#666"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="CopyEmailsButton" 
                           Content="Copy Emails" 
                           Width="110" 
                           Height="32" 
                           Margin="0,0,10,0" 
                           Background="#6c757d" 
                           Foreground="White" 
                           FontWeight="Bold"
                           IsEnabled="False"/>
                    <Button x:Name="ExportToExcelButton" 
                           Content="Export to Excel" 
                           Width="130" 
                           Height="32" 
                           Margin="0,0,10,0" 
                           Background="#28a745" 
                           Foreground="White" 
                           FontWeight="Bold"
                           IsEnabled="False"/>
                    <Button x:Name="CloseButton" 
                           Content="Close" 
                           Width="80" 
                           Height="32" 
                           Background="#6c757d" 
                           Foreground="White"/>
                </StackPanel>
            </DockPanel>
        </Border>
    </Grid>
</Window>
"@
    
    $grpReader = New-Object System.Xml.XmlNodeReader $GroupMembersXAML
    $GrpWindow = [Windows.Markup.XamlReader]::Load($grpReader)
    $GrpWindow.Owner = $syncHash.Window
    
    $GroupIdentityBox = $GrpWindow.FindName("GroupIdentityBox")
    $LoadGroupButton = $GrpWindow.FindName("LoadGroupButton")
    $GroupInfoPanel = $GrpWindow.FindName("GroupInfoPanel")
    $GroupNameText = $GrpWindow.FindName("GroupNameText")
    $GroupTypeText = $GrpWindow.FindName("GroupTypeText")
    $GroupEmailText = $GrpWindow.FindName("GroupEmailText")
    $GroupScopeText = $GrpWindow.FindName("GroupScopeText")
    $GroupMemberCountText = $GrpWindow.FindName("GroupMemberCountText")
    $MembersGrid = $GrpWindow.FindName("MembersGrid")
    $StatusText = $GrpWindow.FindName("StatusText")
    $CopyEmailsButton = $GrpWindow.FindName("CopyEmailsButton")
    $ExportToExcelButton = $GrpWindow.FindName("ExportToExcelButton")
    $CloseButton = $GrpWindow.FindName("CloseButton")
    
    $script:currentGroupInfo = $null
    $script:currentMembers = $null
    
    $MembersGrid.Add_MouseDoubleClick({
        if ($MembersGrid.SelectedItem) {
            $selectedMember = $MembersGrid.SelectedItem.Identity
            Show-ADPropertiesWindow -Identity $selectedMember -Owner $GrpWindow
        }
    })
    
    $LoadGroupButton.Add_Click({
        $groupIdentity = $GroupIdentityBox.Text.Trim()
        
        if ([string]::IsNullOrWhiteSpace($groupIdentity)) {
            [System.Windows.MessageBox]::Show("Please enter a group name, email, or SAM account name", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        try {
            $LoadGroupButton.IsEnabled = $false
            $StatusText.Text = "Loading group information..."
            Write-Log "Loading AD group: $groupIdentity"
            
            # Try to find the group using various methods
            $group = $null
            
            # Try by Identity (works for DN, GUID, SAM, etc.)
            try {
                $group = Get-ADGroup -Identity $groupIdentity -Properties * -ErrorAction Stop
                Write-Log "Found group by identity: $($group.Name)"
            } catch {
                # Try by email address
                try {
                    $group = Get-ADGroup -Filter "mail -eq '$groupIdentity'" -Properties * -ErrorAction Stop
                    Write-Log "Found group by email: $($group.Name)"
                } catch {
                    # Try by display name
                    try {
                        $group = Get-ADGroup -Filter "DisplayName -eq '$groupIdentity'" -Properties * -ErrorAction Stop
                        Write-Log "Found group by display name: $($group.Name)"
                    } catch {
                        throw "Group not found. Please verify the group name, email, or SAM account name."
                    }
                }
            }
            
            if ($null -eq $group) {
                throw "Group not found"
            }
            
            $StatusText.Text = "Loading members..."
            Write-Log "Retrieving members for: $($group.Name)"
            
            # Get group members using AD cmdlets
            $members = Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction Stop
            
            $enrichedMembers = @()
            $processedCount = 0
            $totalCount = $members.Count
            
            foreach ($member in $members) {
                $processedCount++
                $StatusText.Text = "Processing member $processedCount of $totalCount..."
                
                try {
                    $displayName = ""
                    $email = ""
                    $title = ""
                    $department = ""
                    $objectClass = $member.objectClass
                    
                    # Get detailed information based on object type
                    if ($member.objectClass -eq "user") {
                        try {
                            $adUser = Get-ADUser -Identity $member.DistinguishedName -Properties DisplayName, EmailAddress, Title, Department, mail -ErrorAction Stop
                            $displayName = if ($adUser.DisplayName) { $adUser.DisplayName } else { $adUser.Name }
                            $email = if ($adUser.EmailAddress) { $adUser.EmailAddress } elseif ($adUser.mail) { $adUser.mail } else { "" }
                            $title = if ($adUser.Title) { $adUser.Title } else { "" }
                            $department = if ($adUser.Department) { $adUser.Department } else { "" }
                        } catch {
                            Write-Log "Warning: Could not get full details for user $($member.Name)"
                            $displayName = $member.Name
                        }
                    } elseif ($member.objectClass -eq "group") {
                        try {
                            $adGroup = Get-ADGroup -Identity $member.DistinguishedName -Properties DisplayName, mail -ErrorAction Stop
                            $displayName = if ($adGroup.DisplayName) { $adGroup.DisplayName } else { $adGroup.Name }
                            $email = if ($adGroup.mail) { $adGroup.mail } else { "" }
                            $objectClass = "Group"
                        } catch {
                            Write-Log "Warning: Could not get full details for group $($member.Name)"
                            $displayName = $member.Name
                        }
                    } elseif ($member.objectClass -eq "computer") {
                        $displayName = $member.Name
                        $objectClass = "Computer"
                    } elseif ($member.objectClass -eq "contact") {
                        try {
                            $adContact = Get-ADObject -Identity $member.DistinguishedName -Properties DisplayName, mail -ErrorAction Stop
                            $displayName = if ($adContact.DisplayName) { $adContact.DisplayName } else { $member.Name }
                            $email = if ($adContact.mail) { $adContact.mail } else { "" }
                            $objectClass = "Contact"
                        } catch {
                            $displayName = $member.Name
                            $objectClass = "Contact"
                        }
                    } else {
                        $displayName = $member.Name
                    }
                    
                    $memberObj = [PSCustomObject]@{
                        DisplayName = $displayName
                        Email = $email
                        ObjectClass = $objectClass
                        Title = $title
                        Department = $department
                        Identity = $member.DistinguishedName
                        SAMAccountName = $member.SamAccountName
                    }
                    
                    $enrichedMembers += $memberObj
                } catch {
                    Write-Log "Warning: Error processing member $($member.Name): $($_.Exception.Message)"
                    $enrichedMembers += [PSCustomObject]@{
                        DisplayName = $member.Name
                        Email = ""
                        ObjectClass = $member.objectClass
                        Title = ""
                        Department = ""
                        Identity = $member.DistinguishedName
                        SAMAccountName = $member.SamAccountName
                    }
                }
            }
            
            # Sort by display name
            $enrichedMembers = $enrichedMembers | Sort-Object DisplayName
            
            $MembersGrid.ItemsSource = $enrichedMembers
            
            # Determine group type
            $groupType = "Unknown"
            if ($group.GroupCategory -eq "Security") {
                $groupType = "Security Group"
            } elseif ($group.GroupCategory -eq "Distribution") {
                $groupType = "Distribution Group"
            }
            
            # Update info panel
            $GroupNameText.Text = "Group: $($group.Name)"
            $GroupTypeText.Text = "Category: $groupType"
            $GroupScopeText.Text = "Scope: $($group.GroupScope)"
            $GroupEmailText.Text = "Email: $(if ($group.mail) { $group.mail } else { 'N/A' })"
            $GroupMemberCountText.Text = "Total Members: $($enrichedMembers.Count)"
            $GroupInfoPanel.Visibility = [System.Windows.Visibility]::Visible
            
            $script:currentGroupInfo = $group
            $script:currentMembers = $enrichedMembers
            
            $ExportToExcelButton.IsEnabled = $true
            $CopyEmailsButton.IsEnabled = $true
            
            $StatusText.Text = "Loaded $($enrichedMembers.Count) members successfully"
            Write-Log "Successfully loaded $($enrichedMembers.Count) members from $($group.Name)"
            
        } catch {
            Write-Log "Error loading group: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error loading group:`n`n$($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            $StatusText.Text = "Error loading group"
            $GroupInfoPanel.Visibility = [System.Windows.Visibility]::Collapsed
            $MembersGrid.ItemsSource = $null
            $ExportToExcelButton.IsEnabled = $false
            $CopyEmailsButton.IsEnabled = $false
        } finally {
            $LoadGroupButton.IsEnabled = $true
        }
    })
    
    $CopyEmailsButton.Add_Click({
        if ($null -eq $script:currentMembers -or $script:currentMembers.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No members to copy", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        try {
            $emailAddresses = $script:currentMembers | 
                Where-Object { -not [string]::IsNullOrWhiteSpace($_.Email) } | 
                Select-Object -ExpandProperty Email
            
            if ($emailAddresses.Count -eq 0) {
                [System.Windows.MessageBox]::Show("No email addresses found to copy", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            $emailList = $emailAddresses -join "; "
            [System.Windows.Forms.Clipboard]::SetText($emailList)
            
            Write-Log "Copied $($emailAddresses.Count) email addresses to clipboard"
            [System.Windows.MessageBox]::Show("Copied $($emailAddresses.Count) email addresses to clipboard!`n`nYou can now paste them into an email.", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            
        } catch {
            Write-Log "Error copying emails: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error copying emails:`n`n$($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    })
    
    $ExportToExcelButton.Add_Click({
        if ($null -eq $script:currentMembers -or $script:currentMembers.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No members to export", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            [System.Windows.MessageBox]::Show("ImportExcel module is not installed. Cannot export to Excel.`n`nPlease run: Install-Module -Name ImportExcel", "Module Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Write-Log "Export failed: ImportExcel module not installed"
            return
        }
        
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.Title = "Save Group Members Report"
            $groupNameSafe = $script:currentGroupInfo.Name -replace '[\\/:*?"<>|]', '_'
            $saveDialog.FileName = "AD_Group_Members_${groupNameSafe}_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $ExportToExcelButton.IsEnabled = $false
                $StatusText.Text = "Exporting to Excel..."
                Write-Log "Exporting group members to Excel: $excelPath"
                
                $exportData = @()
                foreach ($member in $script:currentMembers) {
                    $exportData += [PSCustomObject]@{
                        'Group Name' = $script:currentGroupInfo.Name
                        'Group Email' = if ($script:currentGroupInfo.mail) { $script:currentGroupInfo.mail } else { "N/A" }
                        'Group Category' = $script:currentGroupInfo.GroupCategory
                        'Group Scope' = $script:currentGroupInfo.GroupScope
                        'Member Display Name' = $member.DisplayName
                        'Member Email' = $member.Email
                        'Member SAM Account' = $member.SAMAccountName
                        'Object Type' = $member.ObjectClass
                        'Title' = $member.Title
                        'Department' = $member.Department
                        'Export Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                $exportData | Export-Excel -Path $excelPath `
                    -WorksheetName "AD Group Members" `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle Medium1 `
                    -TableName "ADGroupMembers"
                
                Write-Log "Successfully exported $($exportData.Count) members to Excel"
                $StatusText.Text = "Export completed successfully"
                
                [System.Windows.MessageBox]::Show(
                    "Group members exported successfully!`n`nFile: $excelPath",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                ) | Out-Null
                
                # Open folder location in Explorer
                $folderPath = Split-Path $excelPath -Parent
                Invoke-Item $folderPath
                
            } else {
                Write-Log "Export cancelled by user"
                $StatusText.Text = "Export cancelled"
            }
            
        } catch {
            Write-Log "Export error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error exporting to Excel:`n`n$($_.Exception.Message)",
                "Export Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            $StatusText.Text = "Export failed"
        } finally {
            $ExportToExcelButton.IsEnabled = $true
        }
    })
    
    $CloseButton.Add_Click({ $GrpWindow.Close() })
    
    $GroupIdentityBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $LoadGroupButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    $GrpWindow.ShowDialog() | Out-Null
})

$syncHash.ExportActiveUsersButton.Add_Click({
    Write-Log "Opening Export Active Users Report"
    
    # Check for Active Directory module
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        [System.Windows.MessageBox]::Show(
            "Active Directory module is not available.`n`nThis feature requires the Active Directory PowerShell module (RSAT).",
            "Module Required",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        Write-Log "Export Active Users failed: ActiveDirectory module not available"
        return
    }
    
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        [System.Windows.MessageBox]::Show(
            "Failed to load Active Directory module.`n`n$($_.Exception.Message)",
            "Module Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        Write-Log "Export Active Users failed: Could not import ActiveDirectory module"
        return
    }
    
    # Check for ImportExcel module
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        [System.Windows.MessageBox]::Show(
            "ImportExcel module is not installed.`n`nThis feature requires the ImportExcel module for Excel export.",
            "Module Required",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        Write-Log "Export Active Users failed: ImportExcel module not installed"
        return
    }
    
    Write-Log "Starting Active Users export..."
    
    # Show progress window
    $progressResult = [System.Windows.MessageBox]::Show(
        "This will retrieve all enabled user accounts from Active and Consultants OUs.`n`nThis may take a few moments. Continue?",
        "Export Active Users",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($progressResult -eq [System.Windows.MessageBoxResult]::No) {
        Write-Log "Export Active Users cancelled by user"
        return
    }
    
    try {
        # Define target OUs - modify these to match your organization's OU structure
        $OUList = @(
            "OU=Active,OU=Accts,DC=gellerco,DC=net", 
            "OU=Consultants,OU=Accts,DC=gellerco,DC=net",
            "OU=Vendor,OU=Accts,DC=gellerco,DC=net",
            "OU=Interns,OU=Accts,DC=gellerco,DC=net"
        )
        
        Write-Log "Retrieving users from OUs..."
        
        # Collect enabled users from all OUs
        $activeUsers = @()
        foreach ($OU in $OUList) {
            try {
                $users = Get-ADUser -Filter {Enabled -eq $True} -SearchBase $OU -Properties Name, SamAccountName, Enabled, EmailAddress, Department, Title, Office -ErrorAction Stop
                if ($users) {
                    $activeUsers += $users
                }
                Write-Log "Retrieved $($users.Count) users from $OU"
            } catch {
                Write-Log "Warning: Could not retrieve users from $OU - $($_.Exception.Message)"
            }
        }
        
        if ($activeUsers.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "No enabled users found in the specified OUs.",
                "No Data",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            Write-Log "No active users found"
            return
        }
        
        Write-Log "Found $($activeUsers.Count) total enabled users"
        
        # Exclude users with "test" or "t-" in SamAccountName
        $filteredUsers = $activeUsers | Where-Object {
            $_.SamAccountName -notmatch '(?i)test|^t-'
        }
        
        Write-Log "Filtered to $($filteredUsers.Count) users (excluding test accounts)"
        
        # Select properties to export and sort by Name in ascending order
        $selectedProperties = $filteredUsers | 
            Select-Object Name, SamAccountName, Enabled, EmailAddress, Department, Title, Office | 
            Sort-Object -Property Name
        
        # Prompt user for save location
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $saveDialog.Title = "Save Active Users Report"
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $saveDialog.FileName = "AD-ActiveAccounts_$timestamp.xlsx"
        
        if ($saveDialog.ShowDialog()) {
            $excelPath = $saveDialog.FileName
            
            Write-Log "Exporting $($selectedProperties.Count) users to Excel: $excelPath"
            
            # Import ImportExcel module
            Import-Module ImportExcel -ErrorAction Stop
            
            # Export to Excel
            $selectedProperties | Export-Excel -Path $excelPath `
                -AutoSize `
                -AutoFilter `
                -FreezeTopRow `
                -BoldTopRow `
                -TableStyle Medium1 `
                -WorksheetName "ActiveAccounts"
            
            Write-Log "Successfully exported active users report"
            
            [System.Windows.MessageBox]::Show(
                "Active users report exported successfully!`n`nFile: $excelPath`n`nExported $($selectedProperties.Count) users.",
                "Export Successful",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            ) | Out-Null
            
            # Open folder location in Explorer
            $folderPath = Split-Path $excelPath -Parent
            Invoke-Item $folderPath
            
        } else {
            Write-Log "Export cancelled by user"
        }
        
    } catch {
        Write-Log "Export error: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show(
            "Error exporting active users:`n`n$($_.Exception.Message)",
            "Export Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
})

$syncHash.AutoRepliesButton.Add_Click({
    # Check if connected
    $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
        $result = [System.Windows.MessageBox]::Show(
            "You are not connected to Exchange Online.`n`nWould you like to connect now?",
            "Connection Required",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            $syncHash.ConnectButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            
            # Wait a moment for connection to establish, then recheck
            Start-Sleep -Seconds 5
            $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
            
            if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
                Write-Log "Connection was not established. Please try again."
                return
            }
        } else {
            return
        }
    }
    
    Write-Log "Opening Automatic Replies window..."
    
    [xml]$AutoRepliesXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Automatic Replies (Out of Office) Management" 
        Height="650" 
        Width="750" 
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <GroupBox Grid.Row="0" Header="Mailbox Information" Margin="15,15,15,10" Padding="15">
            <StackPanel>
                <TextBlock Text="Mailbox Email Address:" FontWeight="Bold" Margin="0,0,0,5"/>
                <DockPanel Margin="0,0,0,10">
                    <Button x:Name="LoadAutoRepliesButton" 
                           Content="Load Settings" 
                           Width="120" 
                           Height="30" 
                           Margin="10,0,0,0" 
                           DockPanel.Dock="Right" 
                           Background="#007bff" 
                           Foreground="White" 
                           FontWeight="Bold"/>
                    <TextBox x:Name="MailboxIdentityBox" 
                            Height="30" 
                            Padding="5"
                            VerticalContentAlignment="Center"/>
                </DockPanel>
                
                <StackPanel x:Name="StatusPanel" Visibility="Collapsed" Margin="0,10,0,0">
                    <Border x:Name="StatusBorder" BorderThickness="1" Padding="10" CornerRadius="3">
                        <StackPanel>
                            <TextBlock x:Name="MailboxNameText" FontWeight="Bold" Margin="0,0,0,5"/>
                            <TextBlock x:Name="AutoReplyStateText" FontSize="11" Margin="0,0,0,3"/>
                            <TextBlock x:Name="ScheduledText" FontSize="11" Margin="0,0,0,3"/>
                        </StackPanel>
                    </Border>
                </StackPanel>
            </StackPanel>
        </GroupBox>
        
        <TabControl Grid.Row="1" Margin="15,0,15,10" x:Name="SettingsTabControl" IsEnabled="False">
            <TabItem Header="Auto Reply Settings">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="20">
                        <GroupBox Header="Status" Padding="10" Margin="0,0,0,15">
                            <StackPanel>
                                <RadioButton x:Name="DisabledRadio" Content="Disabled" GroupName="AutoReplyState" Margin="0,5" FontSize="13"/>
                                <RadioButton x:Name="EnabledRadio" Content="Enabled" GroupName="AutoReplyState" Margin="0,5" FontSize="13"/>
                                <RadioButton x:Name="ScheduledRadio" Content="Scheduled (Time Range)" GroupName="AutoReplyState" Margin="0,5" FontSize="13"/>
                            </StackPanel>
                        </GroupBox>
                        
                        <GroupBox Header="Schedule (Only for Scheduled)" Padding="10" Margin="0,0,0,15" x:Name="ScheduleGroup">
                            <StackPanel>
                                <StackPanel Orientation="Horizontal" Margin="0,5">
                                    <TextBlock Text="Start Date/Time:" Width="120" VerticalAlignment="Center" FontWeight="Bold"/>
                                    <DatePicker x:Name="StartDatePicker" Width="150" Margin="0,0,10,0"/>
                                    <TextBox x:Name="StartTimeBox" Width="80" Height="25" Padding="5" ToolTip="HH:mm format (e.g., 09:00)"/>
                                </StackPanel>
                                <StackPanel Orientation="Horizontal" Margin="0,5">
                                    <TextBlock Text="End Date/Time:" Width="120" VerticalAlignment="Center" FontWeight="Bold"/>
                                    <DatePicker x:Name="EndDatePicker" Width="150" Margin="0,0,10,0"/>
                                    <TextBox x:Name="EndTimeBox" Width="80" Height="25" Padding="5" ToolTip="HH:mm format (e.g., 17:00)"/>
                                </StackPanel>
                                <TextBlock Text="Time format: HH:mm (24-hour, e.g., 09:00 or 17:30)" 
                                          FontSize="10" 
                                          Foreground="Gray" 
                                          FontStyle="Italic" 
                                          Margin="120,5,0,0"/>
                            </StackPanel>
                        </GroupBox>
                        
                        <GroupBox Header="Internal Message (to people in your organization)" Padding="10" Margin="0,0,0,15">
                            <DockPanel>
                                <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,5" Background="#f0f0f0" Height="35">
                                    <Button x:Name="InternalBoldButton" Content="B" Width="30" Height="25" Margin="5,5,2,5" FontWeight="Bold" ToolTip="Bold"/>
                                    <Button x:Name="InternalItalicButton" Content="I" Width="30" Height="25" Margin="2,5" FontStyle="Italic" ToolTip="Italic"/>
                                    <Button x:Name="InternalUnderlineButton" Content="U" Width="30" Height="25" Margin="2,5" ToolTip="Underline">
                                        <Button.Template>
                                            <ControlTemplate TargetType="Button">
                                                <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1">
                                                    <TextBlock Text="{TemplateBinding Content}" TextDecorations="Underline" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                                </Border>
                                            </ControlTemplate>
                                        </Button.Template>
                                    </Button>
                                    <Separator Width="1" Margin="5,5"/>
                                    <Button x:Name="InternalClearButton" Content="Clear Format" Width="90" Height="25" Margin="5" ToolTip="Remove all formatting"/>
                                </StackPanel>
                                <Border BorderBrush="#dee2e6" BorderThickness="1" Padding="5">
                                    <RichTextBox x:Name="InternalRichTextBox" 
                                                Height="120" 
                                                VerticalScrollBarVisibility="Auto"
                                                Background="White"
                                                AcceptsReturn="True"/>
                                </Border>
                            </DockPanel>
                        </GroupBox>
                        
                        <GroupBox Header="External Message (to people outside your organization)" Padding="10" Margin="0,0,0,15">
                            <StackPanel>
                                <CheckBox x:Name="ExternalEnabledCheck" Content="Send automatic replies to external senders" Margin="0,0,0,10" FontWeight="Bold"/>
                                <RadioButton x:Name="ExternalAllRadio" Content="Send to all external senders" GroupName="ExternalAudience" Margin="0,5" IsEnabled="False"/>
                                <RadioButton x:Name="ExternalKnownRadio" Content="Send to external senders in my contacts only" GroupName="ExternalAudience" Margin="0,5" IsEnabled="False"/>
                                <DockPanel Margin="0,10,0,0">
                                    <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,5" Background="#f0f0f0" Height="35">
                                        <Button x:Name="ExternalBoldButton" Content="B" Width="30" Height="25" Margin="5,5,2,5" FontWeight="Bold" ToolTip="Bold" IsEnabled="False"/>
                                        <Button x:Name="ExternalItalicButton" Content="I" Width="30" Height="25" Margin="2,5" FontStyle="Italic" ToolTip="Italic" IsEnabled="False"/>
                                        <Button x:Name="ExternalUnderlineButton" Content="U" Width="30" Height="25" Margin="2,5" ToolTip="Underline" IsEnabled="False">
                                            <Button.Template>
                                                <ControlTemplate TargetType="Button">
                                                    <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1">
                                                        <TextBlock Text="{TemplateBinding Content}" TextDecorations="Underline" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                                    </Border>
                                                </ControlTemplate>
                                            </Button.Template>
                                        </Button>
                                        <Separator Width="1" Margin="5,5"/>
                                        <Button x:Name="ExternalClearButton" Content="Clear Format" Width="90" Height="25" Margin="5" ToolTip="Remove all formatting" IsEnabled="False"/>
                                    </StackPanel>
                                    <Border BorderBrush="#dee2e6" BorderThickness="1" Padding="5">
                                        <RichTextBox x:Name="ExternalRichTextBox" 
                                                    Height="100" 
                                                    VerticalScrollBarVisibility="Auto"
                                                    Background="White"
                                                    AcceptsReturn="True"
                                                    IsEnabled="False"/>
                                    </Border>
                                </DockPanel>
                            </StackPanel>
                        </GroupBox>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>
        
        <Border Grid.Row="2" Background="#f8f9fa" BorderBrush="#dee2e6" BorderThickness="0,1,0,0" Padding="15">
            <DockPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Left" DockPanel.Dock="Left">
                    <TextBlock x:Name="StatusTextBlock" 
                              Text="Enter a mailbox email address to begin" 
                              VerticalAlignment="Center"
                              FontSize="11"
                              Foreground="#666"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="SaveButton" 
                           Content="Save Changes" 
                           Width="130" 
                           Height="32" 
                           Margin="0,0,10,0" 
                           Background="#28a745" 
                           Foreground="White" 
                           FontWeight="Bold"
                           IsEnabled="False"/>
                    <Button x:Name="CloseButton" 
                           Content="Close" 
                           Width="80" 
                           Height="32" 
                           Background="#6c757d" 
                           Foreground="White"/>
                </StackPanel>
            </DockPanel>
        </Border>
    </Grid>
</Window>
"@
    
    $autoReader = New-Object System.Xml.XmlNodeReader $AutoRepliesXAML
    $AutoWindow = [Windows.Markup.XamlReader]::Load($autoReader)
    $AutoWindow.Owner = $syncHash.Window
    
    # Get controls
    $MailboxIdentityBox = $AutoWindow.FindName("MailboxIdentityBox")
    $LoadAutoRepliesButton = $AutoWindow.FindName("LoadAutoRepliesButton")
    $StatusPanel = $AutoWindow.FindName("StatusPanel")
    $StatusBorder = $AutoWindow.FindName("StatusBorder")
    $MailboxNameText = $AutoWindow.FindName("MailboxNameText")
    $AutoReplyStateText = $AutoWindow.FindName("AutoReplyStateText")
    $ScheduledText = $AutoWindow.FindName("ScheduledText")
    
    $SettingsTabControl = $AutoWindow.FindName("SettingsTabControl")
    $DisabledRadio = $AutoWindow.FindName("DisabledRadio")
    $EnabledRadio = $AutoWindow.FindName("EnabledRadio")
    $ScheduledRadio = $AutoWindow.FindName("ScheduledRadio")
    $ScheduleGroup = $AutoWindow.FindName("ScheduleGroup")
    $StartDatePicker = $AutoWindow.FindName("StartDatePicker")
    $StartTimeBox = $AutoWindow.FindName("StartTimeBox")
    $EndDatePicker = $AutoWindow.FindName("EndDatePicker")
    $EndTimeBox = $AutoWindow.FindName("EndTimeBox")
    
    $InternalRichTextBox = $AutoWindow.FindName("InternalRichTextBox")
    $InternalBoldButton = $AutoWindow.FindName("InternalBoldButton")
    $InternalItalicButton = $AutoWindow.FindName("InternalItalicButton")
    $InternalUnderlineButton = $AutoWindow.FindName("InternalUnderlineButton")
    $InternalClearButton = $AutoWindow.FindName("InternalClearButton")
    
    $ExternalEnabledCheck = $AutoWindow.FindName("ExternalEnabledCheck")
    $ExternalAllRadio = $AutoWindow.FindName("ExternalAllRadio")
    $ExternalKnownRadio = $AutoWindow.FindName("ExternalKnownRadio")
    $ExternalRichTextBox = $AutoWindow.FindName("ExternalRichTextBox")
    $ExternalBoldButton = $AutoWindow.FindName("ExternalBoldButton")
    $ExternalItalicButton = $AutoWindow.FindName("ExternalItalicButton")
    $ExternalUnderlineButton = $AutoWindow.FindName("ExternalUnderlineButton")
    $ExternalClearButton = $AutoWindow.FindName("ExternalClearButton")
    
    $StatusTextBlock = $AutoWindow.FindName("StatusTextBlock")
    $SaveButton = $AutoWindow.FindName("SaveButton")
    $CloseButton = $AutoWindow.FindName("CloseButton")
    
    $script:currentMailboxSettings = $null
    
    # Function to convert RichTextBox to HTML
    function Get-RichTextBoxHtml {
        param(
            [System.Windows.Controls.RichTextBox]$RichTextBox
        )
        
        $textRange = New-Object System.Windows.Documents.TextRange($RichTextBox.Document.ContentStart, $RichTextBox.Document.ContentEnd)
        $text = $textRange.Text
        
        if ([string]::IsNullOrWhiteSpace($text)) {
            return ""
        }
        
        $html = "<html><body style='font-family: Calibri, Arial, sans-serif; font-size: 11pt;'>"
        
        foreach ($block in $RichTextBox.Document.Blocks) {
            if ($block -is [System.Windows.Documents.Paragraph]) {
                $html += "<p>"
                
                foreach ($inline in $block.Inlines) {
                    if ($inline -is [System.Windows.Documents.Run]) {
                        $runText = [System.Net.WebUtility]::HtmlEncode($inline.Text)
                        
                        $isBold = $inline.FontWeight -eq [System.Windows.FontWeights]::Bold
                        $isItalic = $inline.FontStyle -eq [System.Windows.FontStyles]::Italic
                        $isUnderline = $inline.TextDecorations.Count -gt 0
                        
                        if ($isBold) { $runText = "<b>$runText</b>" }
                        if ($isItalic) { $runText = "<i>$runText</i>" }
                        if ($isUnderline) { $runText = "<u>$runText</u>" }
                        
                        $html += $runText
                    } elseif ($inline -is [System.Windows.Documents.LineBreak]) {
                        $html += "<br/>"
                    }
                }
                
                $html += "</p>"
            }
        }
        
        $html += "</body></html>"
        return $html
    }
    
    # Function to set RichTextBox from HTML
    function Set-RichTextBoxFromHtml {
        param(
            [System.Windows.Controls.RichTextBox]$RichTextBox,
            [string]$HtmlContent
        )
        
        $RichTextBox.Document.Blocks.Clear()
        
        if ([string]::IsNullOrWhiteSpace($HtmlContent)) {
            return
        }
        
        try {
            # Clean HTML
            $cleanHtml = $HtmlContent -replace '<html[^>]*>', '' -replace '</html>', ''
            $cleanHtml = $cleanHtml -replace '<body[^>]*>', '' -replace '</body>', ''
            $cleanHtml = $cleanHtml -replace '<head>.*?</head>', ''
            
            # Split by paragraphs
            $paragraphs = $cleanHtml -split '<p>|</p>' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            
            foreach ($paraText in $paragraphs) {
                $para = New-Object System.Windows.Documents.Paragraph
                $para.Margin = New-Object System.Windows.Thickness(0)
                
                # Process inline elements
                $currentText = $paraText
                $position = 0
                
                while ($position -lt $currentText.Length) {
                    # Check for tags
                    if ($currentText[$position] -eq '<') {
                        $tagEnd = $currentText.IndexOf('>', $position)
                        if ($tagEnd -gt $position) {
                            $tag = $currentText.Substring($position, $tagEnd - $position + 1)
                            
                            if ($tag -match '<(b|strong|i|em|u)>') {
                                $tagName = $matches[1]
                                $closeTag = "</$tagName>"
                                $closePos = $currentText.IndexOf($closeTag, $tagEnd)
                                
                                if ($closePos -gt $tagEnd) {
                                    $innerText = $currentText.Substring($tagEnd + 1, $closePos - $tagEnd - 1)
                                    $innerText = [System.Net.WebUtility]::HtmlDecode($innerText)
                                    
                                    $run = New-Object System.Windows.Documents.Run($innerText)
                                    
                                    if ($tagName -eq 'b' -or $tagName -eq 'strong') {
                                        $run.FontWeight = [System.Windows.FontWeights]::Bold
                                    }
                                    if ($tagName -eq 'i' -or $tagName -eq 'em') {
                                        $run.FontStyle = [System.Windows.FontStyles]::Italic
                                    }
                                    if ($tagName -eq 'u') {
                                        $run.TextDecorations = [System.Windows.TextDecorations]::Underline
                                    }
                                    
                                    $para.Inlines.Add($run)
                                    $position = $closePos + $closeTag.Length
                                    continue
                                }
                            } elseif ($tag -eq '<br>' -or $tag -eq '<br/>') {
                                $para.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
                                $position = $tagEnd + 1
                                continue
                            }
                            
                            $position = $tagEnd + 1
                        } else {
                            $position++
                        }
                    } else {
                        # Find next tag or end
                        $nextTag = $currentText.IndexOf('<', $position)
                        if ($nextTag -eq -1) { $nextTag = $currentText.Length }
                        
                        $plainText = $currentText.Substring($position, $nextTag - $position)
                        $plainText = [System.Net.WebUtility]::HtmlDecode($plainText)
                        
                        if (-not [string]::IsNullOrWhiteSpace($plainText)) {
                            $run = New-Object System.Windows.Documents.Run($plainText)
                            $para.Inlines.Add($run)
                        }
                        
                        $position = $nextTag
                    }
                }
                
                $RichTextBox.Document.Blocks.Add($para)
            }
            
        } catch {
            # Fallback: just add as plain text
            $plainText = $HtmlContent -replace '<[^>]+>', ''
            $plainText = [System.Net.WebUtility]::HtmlDecode($plainText)
            $para = New-Object System.Windows.Documents.Paragraph
            $para.Inlines.Add((New-Object System.Windows.Documents.Run($plainText)))
            $RichTextBox.Document.Blocks.Clear()
            $RichTextBox.Document.Blocks.Add($para)
        }
    }
    
    # Formatting button handlers for Internal message
    $InternalBoldButton.Add_Click({
        $selection = $InternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $currentWeight = $selection.GetPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty)
            if ($currentWeight -eq [System.Windows.FontWeights]::Bold) {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty, [System.Windows.FontWeights]::Normal)
            } else {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty, [System.Windows.FontWeights]::Bold)
            }
        }
        $InternalRichTextBox.Focus()
    })
    
    $InternalItalicButton.Add_Click({
        $selection = $InternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $currentStyle = $selection.GetPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty)
            if ($currentStyle -eq [System.Windows.FontStyles]::Italic) {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty, [System.Windows.FontStyles]::Normal)
            } else {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty, [System.Windows.FontStyles]::Italic)
            }
        }
        $InternalRichTextBox.Focus()
    })
    
    $InternalUnderlineButton.Add_Click({
        $selection = $InternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $currentDeco = $selection.GetPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty)
            if ($currentDeco -eq [System.Windows.TextDecorations]::Underline) {
                $selection.ApplyPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty, $null)
            } else {
                $selection.ApplyPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty, [System.Windows.TextDecorations]::Underline)
            }
        }
        $InternalRichTextBox.Focus()
    })
    
    $InternalClearButton.Add_Click({
        $selection = $InternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty, [System.Windows.FontWeights]::Normal)
            $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty, [System.Windows.FontStyles]::Normal)
            $selection.ApplyPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty, $null)
        }
        $InternalRichTextBox.Focus()
    })
    
    # Formatting button handlers for External message
    $ExternalBoldButton.Add_Click({
        $selection = $ExternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $currentWeight = $selection.GetPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty)
            if ($currentWeight -eq [System.Windows.FontWeights]::Bold) {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty, [System.Windows.FontWeights]::Normal)
            } else {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty, [System.Windows.FontWeights]::Bold)
            }
        }
        $ExternalRichTextBox.Focus()
    })
    
    $ExternalItalicButton.Add_Click({
        $selection = $ExternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $currentStyle = $selection.GetPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty)
            if ($currentStyle -eq [System.Windows.FontStyles]::Italic) {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty, [System.Windows.FontStyles]::Normal)
            } else {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty, [System.Windows.FontStyles]::Italic)
            }
        }
        $ExternalRichTextBox.Focus()
    })
    
    $ExternalUnderlineButton.Add_Click({
        $selection = $ExternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $currentDeco = $selection.GetPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty)
            if ($currentDeco -eq [System.Windows.TextDecorations]::Underline) {
                $selection.ApplyPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty, $null)
            } else {
                $selection.ApplyPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty, [System.Windows.TextDecorations]::Underline)
            }
        }
        $ExternalRichTextBox.Focus()
    })
    
    $ExternalClearButton.Add_Click({
        $selection = $ExternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty, [System.Windows.FontWeights]::Normal)
            $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty, [System.Windows.FontStyles]::Normal)
            $selection.ApplyPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty, $null)
        }
        $ExternalRichTextBox.Focus()
    })
    
    # Enable/disable schedule fields based on radio selection
    $DisabledRadio.Add_Checked({ 
        $ScheduleGroup.IsEnabled = $false 
    })
    $EnabledRadio.Add_Checked({ 
        $ScheduleGroup.IsEnabled = $false 
    })
    $ScheduledRadio.Add_Checked({ 
        $ScheduleGroup.IsEnabled = $true 
    })
    
    # Enable/disable external message fields
    $ExternalEnabledCheck.Add_Checked({
        $ExternalAllRadio.IsEnabled = $true
        $ExternalKnownRadio.IsEnabled = $true
        $ExternalRichTextBox.IsEnabled = $true
        $ExternalBoldButton.IsEnabled = $true
        $ExternalItalicButton.IsEnabled = $true
        $ExternalUnderlineButton.IsEnabled = $true
        $ExternalClearButton.IsEnabled = $true
        if (-not $ExternalAllRadio.IsChecked -and -not $ExternalKnownRadio.IsChecked) {
            $ExternalAllRadio.IsChecked = $true
        }
    })
    
    $ExternalEnabledCheck.Add_Unchecked({
        $ExternalAllRadio.IsEnabled = $false
        $ExternalKnownRadio.IsEnabled = $false
        $ExternalRichTextBox.IsEnabled = $false
        $ExternalBoldButton.IsEnabled = $false
        $ExternalItalicButton.IsEnabled = $false
        $ExternalUnderlineButton.IsEnabled = $false
        $ExternalClearButton.IsEnabled = $false
    })
    
    $LoadAutoRepliesButton.Add_Click({
        $mailboxIdentity = $MailboxIdentityBox.Text.Trim()
        
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity)) {
            [System.Windows.MessageBox]::Show("Please enter a mailbox email address", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        try {
            $LoadAutoRepliesButton.IsEnabled = $false
            $StatusTextBlock.Text = "Loading automatic reply settings..."
            Write-Log "Loading automatic reply settings for: $mailboxIdentity"
            
            # Get mailbox info
            $mailbox = Get-Mailbox -Identity $mailboxIdentity -ErrorAction Stop
            
            # Get automatic reply configuration
            $autoReplyConfig = Get-MailboxAutoReplyConfiguration -Identity $mailboxIdentity -ErrorAction Stop
            
            $script:currentMailboxSettings = $autoReplyConfig
            
            # Update status panel
            $MailboxNameText.Text = "Mailbox: $($mailbox.DisplayName)"
            
            switch ($autoReplyConfig.AutoReplyState) {
                "Disabled" {
                    $AutoReplyStateText.Text = "Status: Disabled"
                    $AutoReplyStateText.Foreground = [System.Windows.Media.Brushes]::Gray
                    $StatusBorder.Background = [System.Windows.Media.Brushes]::LightGray
                    $StatusBorder.BorderBrush = [System.Windows.Media.Brushes]::Gray
                    $DisabledRadio.IsChecked = $true
                }
                "Enabled" {
                    $AutoReplyStateText.Text = "Status: Enabled"
                    $AutoReplyStateText.Foreground = [System.Windows.Media.Brushes]::Green
                    $StatusBorder.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(230, 255, 230))
                    $StatusBorder.BorderBrush = [System.Windows.Media.Brushes]::Green
                    $EnabledRadio.IsChecked = $true
                }
                "Scheduled" {
                    $AutoReplyStateText.Text = "Status: Scheduled"
                    $AutoReplyStateText.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(255, 140, 0))
                    $StatusBorder.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(255, 248, 220))
                    $StatusBorder.BorderBrush = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(255, 140, 0))
                    $ScheduledRadio.IsChecked = $true
                }
            }
            
            if ($autoReplyConfig.AutoReplyState -eq "Scheduled") {
                $startLocal = $autoReplyConfig.StartTime.ToLocalTime()
                $endLocal = $autoReplyConfig.EndTime.ToLocalTime()
                $ScheduledText.Text = "Active: $($startLocal.ToString('g')) to $($endLocal.ToString('g'))"
                $ScheduledText.Visibility = [System.Windows.Visibility]::Visible
                
                $StartDatePicker.SelectedDate = $startLocal
                $StartTimeBox.Text = $startLocal.ToString("HH:mm")
                $EndDatePicker.SelectedDate = $endLocal
                $EndTimeBox.Text = $endLocal.ToString("HH:mm")
            } else {
                $ScheduledText.Visibility = [System.Windows.Visibility]::Collapsed
                $StartDatePicker.SelectedDate = (Get-Date).Date
                $StartTimeBox.Text = "09:00"
                $EndDatePicker.SelectedDate = (Get-Date).Date.AddDays(7)
                $EndTimeBox.Text = "17:00"
            }
            
            $StatusPanel.Visibility = [System.Windows.Visibility]::Visible
            
            # Load and render messages
            Set-RichTextBoxFromHtml -RichTextBox $InternalRichTextBox -HtmlContent $autoReplyConfig.InternalMessage
            Set-RichTextBoxFromHtml -RichTextBox $ExternalRichTextBox -HtmlContent $autoReplyConfig.ExternalMessage
            
            # External audience
            if ($autoReplyConfig.ExternalAudience -eq "None") {
                $ExternalEnabledCheck.IsChecked = $false
            } else {
                $ExternalEnabledCheck.IsChecked = $true
                if ($autoReplyConfig.ExternalAudience -eq "All") {
                    $ExternalAllRadio.IsChecked = $true
                } else {
                    $ExternalKnownRadio.IsChecked = $true
                }
            }
            
            $SettingsTabControl.IsEnabled = $true
            $SaveButton.IsEnabled = $true
            
            $StatusTextBlock.Text = "Settings loaded successfully"
            Write-Log "Successfully loaded automatic reply settings for $($mailbox.DisplayName)"
            
        } catch {
            Write-Log "Error loading automatic reply settings: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error loading settings:`n`n$($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            $StatusTextBlock.Text = "Error loading settings"
            $StatusPanel.Visibility = [System.Windows.Visibility]::Collapsed
            $SettingsTabControl.IsEnabled = $false
            $SaveButton.IsEnabled = $false
        } finally {
            $LoadAutoRepliesButton.IsEnabled = $true
        }
    })
    
    $SaveButton.Add_Click({
        $mailboxIdentity = $MailboxIdentityBox.Text.Trim()
        
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity)) {
            [System.Windows.MessageBox]::Show("No mailbox loaded", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        # Determine state
        $newState = "Disabled"
        if ($EnabledRadio.IsChecked) {
            $newState = "Enabled"
        } elseif ($ScheduledRadio.IsChecked) {
            $newState = "Scheduled"
        }
        
        # Validate scheduled dates if needed
        if ($newState -eq "Scheduled") {
            if (-not $StartDatePicker.SelectedDate -or -not $EndDatePicker.SelectedDate) {
                [System.Windows.MessageBox]::Show("Please select both start and end dates for scheduled automatic replies", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            if ([string]::IsNullOrWhiteSpace($StartTimeBox.Text) -or [string]::IsNullOrWhiteSpace($EndTimeBox.Text)) {
                [System.Windows.MessageBox]::Show("Please enter both start and end times in HH:mm format (e.g., 09:00)", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            # Validate time format
            if ($StartTimeBox.Text -notmatch '^\d{2}:\d{2}$' -or $EndTimeBox.Text -notmatch '^\d{2}:\d{2}$') {
                [System.Windows.MessageBox]::Show("Time must be in HH:mm format (e.g., 09:00 or 17:30)", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            try {
                $startTime = [datetime]::Parse($StartTimeBox.Text)
                $endTime = [datetime]::Parse($EndTimeBox.Text)
            } catch {
                [System.Windows.MessageBox]::Show("Invalid time format. Please use HH:mm (e.g., 09:00)", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            $startDateTime = $StartDatePicker.SelectedDate.Date.Add($startTime.TimeOfDay)
            $endDateTime = $EndDatePicker.SelectedDate.Date.Add($endTime.TimeOfDay)
            
            if ($endDateTime -le $startDateTime) {
                [System.Windows.MessageBox]::Show("End date/time must be after start date/time", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
        }
        
        # Determine external audience
        $externalAudience = "None"
        if ($ExternalEnabledCheck.IsChecked) {
            if ($ExternalAllRadio.IsChecked) {
                $externalAudience = "All"
            } else {
                $externalAudience = "Known"
            }
        }
        
        $result = [System.Windows.MessageBox]::Show("Save automatic reply settings for this mailbox?", "Confirm", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                $SaveButton.IsEnabled = $false
                $StatusTextBlock.Text = "Saving settings..."
                Write-Log "Saving automatic reply settings for $mailboxIdentity"
                
                # Convert RichTextBox content to HTML
                $internalHtml = Get-RichTextBoxHtml -RichTextBox $InternalRichTextBox
                $externalHtml = Get-RichTextBoxHtml -RichTextBox $ExternalRichTextBox
                
                $setParams = @{
                    Identity = $mailboxIdentity
                    AutoReplyState = $newState
                    InternalMessage = $internalHtml
                    ExternalMessage = $externalHtml
                    ExternalAudience = $externalAudience
                }
                
                if ($newState -eq "Scheduled") {
                    $setParams.StartTime = $startDateTime
                    $setParams.EndTime = $endDateTime
                }
                
                Set-MailboxAutoReplyConfiguration @setParams -ErrorAction Stop
                
                Write-Log "Successfully saved automatic reply settings"
                $StatusTextBlock.Text = "Settings saved successfully"
                [System.Windows.MessageBox]::Show("Automatic reply settings saved successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                
                # Reload to show updated status
                $LoadAutoRepliesButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
                
            } catch {
                Write-Log "Error saving automatic reply settings: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Error saving settings:`n`n$($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                $StatusTextBlock.Text = "Error saving settings"
            } finally {
                $SaveButton.IsEnabled = $true
            }
        }
    })
    
    $CloseButton.Add_Click({ $AutoWindow.Close() })
    
    # Add Enter key support
    $MailboxIdentityBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $LoadAutoRepliesButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    $AutoWindow.ShowDialog() | Out-Null
})

$syncHash.MessageTraceButton.Add_Click({
    # Check if connected
    $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
        $result = [System.Windows.MessageBox]::Show(
            "You are not connected to Exchange Online.`n`nWould you like to connect now?",
            "Connection Required",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            $syncHash.ConnectButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            
            # Wait a moment for connection to establish, then recheck
            Start-Sleep -Seconds 5
            $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
            
            if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
                Write-Log "Connection was not established. Please try again."
                return
            }
        } else {
            return
        }
    }
    
    Write-Log "Opening Message Trace window..."
    
    [xml]$MessageTraceXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Message Trace / Tracking" 
        Height="750" 
        Width="900" 
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <TabControl Margin="10">
            <TabItem Header="Search">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <GroupBox Grid.Row="0" Header="Search Criteria" Padding="15" Margin="0,0,0,15">
                        <StackPanel>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="15"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                
                                <StackPanel Grid.Column="0">
                                    <TextBlock Text="Sender Email Address:" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <TextBox x:Name="SenderEmailBox" MinHeight="30" Padding="5,8,5,5" VerticalContentAlignment="Center" Margin="0,0,0,15"/>

                                    
                                    <TextBlock Text="Recipient Email Address:" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <TextBox x:Name="RecipientEmailBox" MinHeight="30" Padding="5,8,5,5" VerticalContentAlignment="Center" Margin="0,0,0,15"/>
                                    
                                    <TextBlock Text="Subject Contains:" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <TextBox x:Name="SubjectBox" MinHeight="30" Padding="5,8,5,5" VerticalContentAlignment="Center"/>
                                </StackPanel>
                                
                                <StackPanel Grid.Column="2">
                                    <TextBlock Text="Message ID:" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <TextBox x:Name="MessageIdBox" MinHeight="30" Padding="5,8,5,5" VerticalContentAlignment="Center" Margin="0,0,0,15"/>

                                    <TextBlock Text="Status:" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <ComboBox x:Name="StatusCombo" MinHeight="30" Margin="0,0,0,15" SelectedIndex="0">
                                        <ComboBoxItem Content="All" Tag="All"/>
                                        <ComboBoxItem Content="Delivered" Tag="Delivered"/>
                                        <ComboBoxItem Content="Failed" Tag="Failed"/>
                                        <ComboBoxItem Content="Pending" Tag="Pending"/>
                                        <ComboBoxItem Content="Quarantined" Tag="Quarantined"/>
                                        <ComboBoxItem Content="FilteredAsSpam" Tag="FilteredAsSpam"/>
                                    </ComboBox>
                                    
                                    <TextBlock Text="Page Size (max results):" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <ComboBox x:Name="PageSizeCombo" MinHeight="30" SelectedIndex="1">
                                        <ComboBoxItem Content="100" Tag="100"/>
                                        <ComboBoxItem Content="1000" Tag="1000"/>
                                        <ComboBoxItem Content="5000" Tag="5000"/>
                                    </ComboBox>
                                </StackPanel>
                            </Grid>
                        </StackPanel>
                    </GroupBox>
                    
                    <GroupBox Grid.Row="1" Header="Date Range (Max 10 Days)" Padding="15" Margin="0,0,0,15">
                        <StackPanel>
                            <RadioButton x:Name="Last24HoursRadio" Content="Last 24 Hours" GroupName="DateRange" Margin="0,5" IsChecked="True"/>
                            <RadioButton x:Name="Last7DaysRadio" Content="Last 7 Days" GroupName="DateRange" Margin="0,5"/>
                            <RadioButton x:Name="CustomRangeRadio" Content="Custom Date Range" GroupName="DateRange" Margin="0,5"/>
                            
                            <StackPanel x:Name="CustomDatePanel" Margin="20,10,0,0" IsEnabled="False">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="15"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    
                                    <TextBlock Grid.Column="0" Text="Start:" VerticalAlignment="Center" Margin="0,0,10,0" FontWeight="Bold"/>
                                    <DatePicker Grid.Column="1" x:Name="StartDatePicker" Height="25"/>
                                    
                                    <TextBlock Grid.Column="3" Text="End:" VerticalAlignment="Center" Margin="0,0,10,0" FontWeight="Bold"/>
                                    <DatePicker Grid.Column="4" x:Name="EndDatePicker" Height="25"/>
                                </Grid>
                            </StackPanel>
                        </StackPanel>
                    </GroupBox>
                    
                    <Border Grid.Row="2" Background="#fff3cd" BorderBrush="#ffc107" BorderThickness="1" Padding="10" Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock FontWeight="Bold" Foreground="#856404" Margin="0,0,0,5">
                                <Run Text="&#x24D8;"/> Search Tips:
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#856404">
                                <Run Text="&#x2022;"/> At least one search criteria is required (sender, recipient, or message ID)
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#856404">
                                <Run Text="&#x2022;"/> Date range is limited to 10 days maximum
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#856404">
                                <Run Text="&#x2022;"/> Large searches may take several minutes to complete
                            </TextBlock>
                        </StackPanel>
                    </Border>
                    
                    <StackPanel Grid.Row="7" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
                        <Button x:Name="SearchButton" Content="Search Messages" Width="140" Height="35" Margin="0,0,10,0" Background="#007bff" Foreground="White" FontWeight="Bold"/>
                        <Button x:Name="ClearButton" Content="Clear" Width="80" Height="35" Margin="0,0,10,0" Background="#6c757d" Foreground="White"/>
                        <Button x:Name="SearchCloseButton" Content="Close" Width="80" Height="35" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="Results" x:Name="ResultsTab">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Margin="0,0,0,10">
                        <TextBlock x:Name="ResultsCountText" Text="No results yet. Use the Search tab to find messages." FontWeight="Bold" Margin="0,0,0,10"/>
                        <TextBlock x:Name="ResultsInfoText" Text="" FontSize="11" Foreground="#666" TextWrapping="Wrap"/>
                    </StackPanel>
                    
                    <Border Grid.Row="1" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15">
                        <DataGrid x:Name="ResultsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single" AlternatingRowBackground="#f8f9fa">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Received" Binding="{Binding Received}" Width="140"/>
                                <DataGridTextColumn Header="Sender" Binding="{Binding SenderAddress}" Width="200"/>
                                <DataGridTextColumn Header="Recipient" Binding="{Binding RecipientAddress}" Width="200"/>
                                <DataGridTextColumn Header="Subject" Binding="{Binding Subject}" Width="250"/>
                                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="100"/>
                                <DataGridTextColumn Header="Size (KB)" Binding="{Binding Size}" Width="80"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Border>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="ViewDetailsButton" Content="View Details" Width="120" Height="32" Margin="0,0,10,0" Background="#17a2b8" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="ExportResultsButton" Content="Export to Excel" Width="130" Height="32" Margin="0,0,10,0" Background="#28a745" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="ResultsCloseButton" Content="Close" Width="80" Height="32" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="Message Details" x:Name="DetailsTab">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <GroupBox Grid.Row="0" Header="Message Information" Padding="10" Margin="0,0,0,15">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Grid.Column="0" Text="Message ID:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="0" Grid.Column="1" x:Name="DetailMessageIdBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="1" Grid.Column="0" Text="Subject:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="1" Grid.Column="1" x:Name="DetailSubjectBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="2" Grid.Column="0" Text="From:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="2" Grid.Column="1" x:Name="DetailFromBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="3" Grid.Column="0" Text="To:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="3" Grid.Column="1" x:Name="DetailToBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="4" Grid.Column="0" Text="Received:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="4" Grid.Column="1" x:Name="DetailReceivedBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="5" Grid.Column="0" Text="Status:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="5" Grid.Column="1" x:Name="DetailStatusBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                        </Grid>
                    </GroupBox>
                    
                    <GroupBox Grid.Row="1" Header="Delivery Events" Padding="10" Margin="0,0,0,15">
                        <Border BorderBrush="#dee2e6" BorderThickness="1">
                            <DataGrid x:Name="DetailsGrid" AutoGenerateColumns="False" IsReadOnly="True" AlternatingRowBackground="#f8f9fa">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Timestamp" Binding="{Binding Timestamp}" Width="140"/>
                                    <DataGridTextColumn Header="Event" Binding="{Binding Event}" Width="150"/>
                                    <DataGridTextColumn Header="Detail" Binding="{Binding Detail}" Width="*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Border>
                    </GroupBox>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="CopyMessageIdButton" Content="Copy Message ID" Width="140" Height="32" Margin="0,0,10,0" Background="#6c757d" Foreground="White" IsEnabled="False"/>
                        <Button x:Name="DetailsCloseButton" Content="Close" Width="80" Height="32" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@
    
    $traceReader = New-Object System.Xml.XmlNodeReader $MessageTraceXAML
    $TraceWindow = [Windows.Markup.XamlReader]::Load($traceReader)
    $TraceWindow.Owner = $syncHash.Window
    
    # Search Tab Controls
    $SenderEmailBox = $TraceWindow.FindName("SenderEmailBox")
    $RecipientEmailBox = $TraceWindow.FindName("RecipientEmailBox")
    $SubjectBox = $TraceWindow.FindName("SubjectBox")
    $MessageIdBox = $TraceWindow.FindName("MessageIdBox")
    $StatusCombo = $TraceWindow.FindName("StatusCombo")
    $PageSizeCombo = $TraceWindow.FindName("PageSizeCombo")
    
    $Last24HoursRadio = $TraceWindow.FindName("Last24HoursRadio")
    $Last7DaysRadio = $TraceWindow.FindName("Last7DaysRadio")
    $CustomRangeRadio = $TraceWindow.FindName("CustomRangeRadio")
    $CustomDatePanel = $TraceWindow.FindName("CustomDatePanel")
    $StartDatePicker = $TraceWindow.FindName("StartDatePicker")
    $EndDatePicker = $TraceWindow.FindName("EndDatePicker")
    
    $SearchButton = $TraceWindow.FindName("SearchButton")
    $ClearButton = $TraceWindow.FindName("ClearButton")
    $SearchCloseButton = $TraceWindow.FindName("SearchCloseButton")
    
    # Results Tab Controls
    $ResultsTab = $TraceWindow.FindName("ResultsTab")
    $ResultsCountText = $TraceWindow.FindName("ResultsCountText")
    $ResultsInfoText = $TraceWindow.FindName("ResultsInfoText")
    $ResultsGrid = $TraceWindow.FindName("ResultsGrid")
    $ViewDetailsButton = $TraceWindow.FindName("ViewDetailsButton")
    $ExportResultsButton = $TraceWindow.FindName("ExportResultsButton")
    $ResultsCloseButton = $TraceWindow.FindName("ResultsCloseButton")
    
    # Details Tab Controls
    $DetailsTab = $TraceWindow.FindName("DetailsTab")
    $DetailMessageIdBox = $TraceWindow.FindName("DetailMessageIdBox")
    $DetailSubjectBox = $TraceWindow.FindName("DetailSubjectBox")
    $DetailFromBox = $TraceWindow.FindName("DetailFromBox")
    $DetailToBox = $TraceWindow.FindName("DetailToBox")
    $DetailReceivedBox = $TraceWindow.FindName("DetailReceivedBox")
    $DetailStatusBox = $TraceWindow.FindName("DetailStatusBox")
    $DetailsGrid = $TraceWindow.FindName("DetailsGrid")
    $CopyMessageIdButton = $TraceWindow.FindName("CopyMessageIdButton")
    $DetailsCloseButton = $TraceWindow.FindName("DetailsCloseButton")
    
    # Initialize date pickers
    $StartDatePicker.SelectedDate = (Get-Date).AddDays(-7)
    $EndDatePicker.SelectedDate = Get-Date
    
    # Store current results
    $script:currentTraceResults = $null
    $script:selectedMessage = $null
    
    # Enable/disable custom date panel
    $CustomRangeRadio.Add_Checked({ $CustomDatePanel.IsEnabled = $true })
    $Last24HoursRadio.Add_Checked({ $CustomDatePanel.IsEnabled = $false })
    $Last7DaysRadio.Add_Checked({ $CustomDatePanel.IsEnabled = $false })
    
    # Clear button
    $ClearButton.Add_Click({
        $SenderEmailBox.Clear()
        $RecipientEmailBox.Clear()
        $SubjectBox.Clear()
        $MessageIdBox.Clear()
        $StatusCombo.SelectedIndex = 0
        $PageSizeCombo.SelectedIndex = 1
        $Last24HoursRadio.IsChecked = $true
        Write-Log "Message trace search criteria cleared"
    })
    
    # Search button
    $SearchButton.Add_Click({
        $sender = $SenderEmailBox.Text.Trim()
        $recipient = $RecipientEmailBox.Text.Trim()
        $subject = $SubjectBox.Text.Trim()
        $messageId = $MessageIdBox.Text.Trim()
        
        # Validation
        if ([string]::IsNullOrWhiteSpace($sender) -and 
            [string]::IsNullOrWhiteSpace($recipient) -and 
            [string]::IsNullOrWhiteSpace($messageId)) {
            [System.Windows.MessageBox]::Show(
                "Please enter at least one search criteria:`n`n• Sender Email`n• Recipient Email`n• Message ID",
                "Validation",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
            return
        }
        
        # Get date range
        $startDate = $null
        $endDate = Get-Date
        
        if ($Last24HoursRadio.IsChecked) {
            $startDate = (Get-Date).AddHours(-24)
        } elseif ($Last7DaysRadio.IsChecked) {
            $startDate = (Get-Date).AddDays(-7)
        } else {
            if (-not $StartDatePicker.SelectedDate -or -not $EndDatePicker.SelectedDate) {
                [System.Windows.MessageBox]::Show("Please select both start and end dates", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            $startDate = $StartDatePicker.SelectedDate
            $endDate = $EndDatePicker.SelectedDate.AddDays(1).AddSeconds(-1)
            
            # Check 10-day limit
            $daysDiff = ($endDate - $startDate).TotalDays
            if ($daysDiff -gt 10) {
                [System.Windows.MessageBox]::Show("Date range cannot exceed 10 days", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
        }
        
        try {
            $SearchButton.IsEnabled = $false
            $ResultsCountText.Text = "Searching... This may take a few minutes for large result sets."
            $ResultsInfoText.Text = ""
            $ResultsGrid.ItemsSource = $null
            
            Write-Log "Starting message trace search..."
            Write-Log "  Sender: $sender"
            Write-Log "  Recipient: $recipient"
            Write-Log "  Start: $($startDate.ToString('yyyy-MM-dd HH:mm'))"
            Write-Log "  End: $($endDate.ToString('yyyy-MM-dd HH:mm'))"
            
            # Build parameters
            $traceParams = @{
                StartDate = $startDate
                EndDate = $endDate
                PageSize = [int]$PageSizeCombo.SelectedItem.Tag
            }
            
            if (-not [string]::IsNullOrWhiteSpace($sender)) {
                $traceParams.SenderAddress = $sender
            }
            if (-not [string]::IsNullOrWhiteSpace($recipient)) {
                $traceParams.RecipientAddress = $recipient
            }
            if (-not [string]::IsNullOrWhiteSpace($messageId)) {
                $traceParams.MessageId = $messageId
            }
            
            $status = $StatusCombo.SelectedItem.Tag
            if ($status -ne "All") {
                $traceParams.Status = $status
            }
            
            # Execute search
            $results = @(Get-MessageTrace @traceParams -ErrorAction Stop)
            
            Write-Log "Found $($results.Count) messages"
            
            if ($results.Count -eq 0) {
                $ResultsCountText.Text = "No messages found matching the search criteria"
                $ResultsInfoText.Text = "Try adjusting your search parameters or expanding the date range"
                $script:currentTraceResults = $null
                $ExportResultsButton.IsEnabled = $false
            } else {
                # Format results for display
                $displayResults = @()
                foreach ($msg in $results) {
                    $sizeKB = if ($msg.Size) { [math]::Round($msg.Size / 1KB, 2) } else { 0 }
                    
                    $displayResults += [PSCustomObject]@{
                        Received = $msg.Received.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
                        SenderAddress = $msg.SenderAddress
                        RecipientAddress = $msg.RecipientAddress
                        Subject = if ($msg.Subject) { $msg.Subject } else { "(No Subject)" }
                        Status = $msg.Status
                        Size = $sizeKB
                        MessageId = $msg.MessageId
                        MessageTraceId = $msg.MessageTraceId
                        FromIP = $msg.FromIP
                        ToIP = $msg.ToIP
                    }
                }
                
                $ResultsGrid.ItemsSource = $displayResults
                $script:currentTraceResults = $displayResults
                
                $ResultsCountText.Text = "Found $($results.Count) message(s)"
                $ResultsInfoText.Text = "Double-click a message or use 'View Details' to see full delivery information"
                $ExportResultsButton.IsEnabled = $true
                
                # Switch to Results tab
                $ResultsTab.IsSelected = $true
            }
            
        } catch {
            Write-Log "Error during message trace: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error searching messages:`n`n$($_.Exception.Message)",
                "Search Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            $ResultsCountText.Text = "Search failed"
            $ResultsInfoText.Text = $_.Exception.Message
        } finally {
            $SearchButton.IsEnabled = $true
        }
    })
    
    # Results grid selection changed
    $ResultsGrid.Add_SelectionChanged({
        $ViewDetailsButton.IsEnabled = ($null -ne $ResultsGrid.SelectedItem)
    })
    
    # Double-click to view details
    $ResultsGrid.Add_MouseDoubleClick({
        if ($ResultsGrid.SelectedItem) {
            $ViewDetailsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    # View Details button
    $ViewDetailsButton.Add_Click({
        $selected = $ResultsGrid.SelectedItem
        if ($null -eq $selected) { return }
        
        try {
            $ViewDetailsButton.IsEnabled = $false
            Write-Log "Loading message details for: $($selected.MessageId)"
            
            # Get message trace details
            $messageId = $selected.MessageId
            $recipient = $selected.RecipientAddress
            
            $details = @(Get-MessageTraceDetail -MessageTraceId $selected.MessageTraceId -RecipientAddress $recipient -ErrorAction Stop)
            
            Write-Log "Retrieved $($details.Count) detail events"
            
            # Populate message info
            $DetailMessageIdBox.Text = $selected.MessageId
            $DetailSubjectBox.Text = $selected.Subject
            $DetailFromBox.Text = $selected.SenderAddress
            $DetailToBox.Text = $selected.RecipientAddress
            $DetailReceivedBox.Text = $selected.Received
            $DetailStatusBox.Text = $selected.Status
            
            # Populate events
            $eventList = @()
            foreach ($detail in $details) {
                $eventList += [PSCustomObject]@{
                    Timestamp = $detail.Date.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
                    Event = $detail.Event
                    Detail = $detail.Detail
                }
            }
            
            $DetailsGrid.ItemsSource = $eventList | Sort-Object Timestamp
            $CopyMessageIdButton.IsEnabled = $true
            
            # Switch to Details tab
            $DetailsTab.IsSelected = $true
            
        } catch {
            Write-Log "Error loading message details: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error loading message details:`n`n$($_.Exception.Message)",
                "Details Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        } finally {
            $ViewDetailsButton.IsEnabled = $true
        }
    })
    
    # Export to Excel
    $ExportResultsButton.Add_Click({
        if ($null -eq $script:currentTraceResults -or $script:currentTraceResults.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No results to export", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            [System.Windows.MessageBox]::Show("ImportExcel module is not installed. Cannot export to Excel.`n`nPlease run: Install-Module -Name ImportExcel", "Module Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Write-Log "Export failed: ImportExcel module not installed"
            return
        }
        
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.Title = "Save Message Trace Report"
            $saveDialog.FileName = "MessageTrace_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $ExportResultsButton.IsEnabled = $false
                Write-Log "Exporting message trace results to Excel: $excelPath"
                
                $exportData = @()
                foreach ($msg in $script:currentTraceResults) {
                    $exportData += [PSCustomObject]@{
                        'Received' = $msg.Received
                        'Sender' = $msg.SenderAddress
                        'Recipient' = $msg.RecipientAddress
                        'Subject' = $msg.Subject
                        'Status' = $msg.Status
                        'Size (KB)' = $msg.Size
                        'Message ID' = $msg.MessageId
                        'From IP' = $msg.FromIP
                        'To IP' = $msg.ToIP
                        'Export Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                $exportData | Export-Excel -Path $excelPath `
                    -WorksheetName "Message Trace" `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle Medium1 `
                    -TableName "MessageTrace"
                
                Write-Log "Successfully exported $($exportData.Count) messages to Excel"
                
                [System.Windows.MessageBox]::Show(
                    "Message trace results exported successfully!`n`nFile: $excelPath",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                ) | Out-Null
                
                # Open folder location in Explorer
                $folderPath = Split-Path $excelPath -Parent
                Invoke-Item $folderPath
                
            } else {
                Write-Log "Export cancelled by user"
            }
            
        } catch {
            Write-Log "Export error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error exporting to Excel:`n`n$($_.Exception.Message)",
                "Export Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        } finally {
            $ExportResultsButton.IsEnabled = $true
        }
    })
    
    # Copy Message ID button
    $CopyMessageIdButton.Add_Click({
        if (-not [string]::IsNullOrWhiteSpace($DetailMessageIdBox.Text)) {
            [System.Windows.Forms.Clipboard]::SetText($DetailMessageIdBox.Text)
            [System.Windows.MessageBox]::Show("Message ID copied to clipboard!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
    })
    
    # Close buttons
    $SearchCloseButton.Add_Click({ $TraceWindow.Close() })
    $ResultsCloseButton.Add_Click({ $TraceWindow.Close() })
    $DetailsCloseButton.Add_Click({ $TraceWindow.Close() })
    
    # Enter key support
    $SenderEmailBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $SearchButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    $RecipientEmailBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $SearchButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    $MessageIdBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $SearchButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    $TraceWindow.ShowDialog() | Out-Null
})

# Future Feature Placeholders
$syncHash.SendOnBehalfButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Send on Behalf Permissions feature is planned for version 2.10.0",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})

$syncHash.ForwardingButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Email Forwarding Management feature is planned for version 2.11.0",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})

$syncHash.ResourceMailboxButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Resource Mailbox Management feature is planned for version 2.13.0",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})

$syncHash.DistributionGroupButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Distribution List Management feature is planned for version 2.15.0",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})



$syncHash.LitigationHoldButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Litigation Hold Management feature is planned for version 3.0.0",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})

$syncHash.MailboxStatsButton.Add_Click({
    # Check if connected
    $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
        $result = [System.Windows.MessageBox]::Show(
            "You are not connected to Exchange Online.`n`nWould you like to connect now?",
            "Connection Required",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            $syncHash.ConnectButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            
            # Wait for connection to complete
            Start-Sleep -Seconds 2
            $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
            
            if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
                Write-Log "Connection was not established. Please try again."
                return
            }
        } else {
            return
        }
    }
    
    Write-Log "Opening Mailbox Size & Quota Report window..."
    
    [xml]$MailboxStatsXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Mailbox Size &amp; Quota Report" 
        Height="700" 
        Width="1000" 
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <TabControl Margin="10">
            <TabItem Header="Scan">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <GroupBox Grid.Row="0" Header="Scan Options" Padding="15" Margin="0,0,0,15">
                        <StackPanel>
                            <RadioButton x:Name="ScanAllRadio" Content="Scan All Mailboxes" GroupName="ScanType" Margin="0,5" IsChecked="True" FontSize="13"/>
                            <RadioButton x:Name="ScanSpecificRadio" Content="Scan Specific Mailbox" GroupName="ScanType" Margin="0,5" FontSize="13"/>
                            
                            <StackPanel x:Name="SpecificMailboxPanel" Margin="25,10,0,0" IsEnabled="False">
                                <TextBlock Text="Enter email address:" FontWeight="Bold" Margin="0,0,0,5"/>
                                <TextBox x:Name="SpecificMailboxBox" MinHeight="30" Padding="5,8,5,5" VerticalContentAlignment="Center" Width="300" HorizontalAlignment="Left"/>
                            </StackPanel>
                        </StackPanel>
                    </GroupBox>
                    
                    <GroupBox Grid.Row="1" Header="Mailbox Type Filter" Padding="15" Margin="0,0,0,15">
                        <StackPanel Orientation="Horizontal">
                            <CheckBox x:Name="IncludeUserMailboxes" Content="User Mailboxes" Margin="0,0,20,0" IsChecked="True" FontSize="13"/>
                            <CheckBox x:Name="IncludeSharedMailboxes" Content="Shared Mailboxes" Margin="0,0,20,0" IsChecked="True" FontSize="13"/>
                            <CheckBox x:Name="IncludeArchives" Content="Archive Mailboxes" IsChecked="False" FontSize="13"/>
                        </StackPanel>
                    </GroupBox>
                    
                    <GroupBox Grid.Row="2" Header="Progress" Padding="15" Margin="0,0,0,15" x:Name="ProgressGroup" Visibility="Collapsed">
                        <StackPanel>
                            <TextBlock x:Name="ProgressText" Text="Ready to scan..." Margin="0,0,0,10" FontWeight="Bold"/>
                            <ProgressBar x:Name="ScanProgressBar" Height="25" Minimum="0" Maximum="100" Value="0"/>
                            <TextBlock x:Name="ProgressDetailText" Text="" Margin="0,10,0,0" FontSize="11" Foreground="#666" TextWrapping="Wrap"/>
                        </StackPanel>
                    </GroupBox>
                    
                    <Border Grid.Row="3" Background="#d1ecf1" BorderBrush="#17a2b8" BorderThickness="1" Padding="15" Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock FontWeight="Bold" Foreground="#0c5460" Margin="0,0,0,5" FontSize="13">
                                <Run Text="&#x24D8;"/> Scan Information:
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#0c5460" Margin="0,0,0,3">
                                <Run Text="&#x2022;"/> Scanning all mailboxes may take several minutes depending on organization size
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#0c5460" Margin="0,0,0,3">
                                <Run Text="&#x2022;"/> Results will show mailbox size, item count, quota, and percentage used
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#0c5460" Margin="0,0,0,3">
                                <Run Text="&#x2022;"/> Mailboxes over 80% quota will be highlighted in red
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#0c5460">
                                <Run Text="&#x2022;"/> Export to Excel for detailed analysis and reporting
                            </TextBlock>
                        </StackPanel>
                    </Border>
                    
                    <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="StartScanButton" Content="Start Scan" Width="120" Height="35" Margin="0,0,10,0" Background="#28a745" Foreground="White" FontWeight="Bold"/>
                        <Button x:Name="StopScanButton" Content="Stop Scan" Width="100" Height="35" Margin="0,0,10,0" Background="#dc3545" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="ScanCloseButton" Content="Close" Width="80" Height="35" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="Results" x:Name="ResultsTab">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <GroupBox Grid.Row="0" Header="Summary Statistics" Padding="15" Margin="0,0,0,15">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            
                            <StackPanel Grid.Column="0" Margin="0,0,15,0">
                                <TextBlock Text="Total Mailboxes" FontSize="11" Foreground="#666" Margin="0,0,0,5"/>
                                <TextBlock x:Name="TotalMailboxesText" Text="0" FontSize="24" FontWeight="Bold" Foreground="#007bff"/>
                            </StackPanel>
                            
                            <StackPanel Grid.Column="1" Margin="0,0,15,0">
                                <TextBlock Text="Total Size" FontSize="11" Foreground="#666" Margin="0,0,0,5"/>
                                <TextBlock x:Name="TotalSizeText" Text="0 GB" FontSize="24" FontWeight="Bold" Foreground="#28a745"/>
                            </StackPanel>
                            
                            <StackPanel Grid.Column="2" Margin="0,0,15,0">
                                <TextBlock Text="Average Size" FontSize="11" Foreground="#666" Margin="0,0,0,5"/>
                                <TextBlock x:Name="AverageSizeText" Text="0 GB" FontSize="24" FontWeight="Bold" Foreground="#17a2b8"/>
                            </StackPanel>
                            
                            <StackPanel Grid.Column="3">
                                <TextBlock Text="Over 80% Quota" FontSize="11" Foreground="#666" Margin="0,0,0,5"/>
                                <TextBlock x:Name="OverQuotaText" Text="0" FontSize="24" FontWeight="Bold" Foreground="#dc3545"/>
                            </StackPanel>
                        </Grid>
                    </GroupBox>
                    
                    <StackPanel Grid.Row="1" Margin="0,0,0,10">
                        <DockPanel>
                            <TextBlock DockPanel.Dock="Left" x:Name="ResultsCountText" Text="No results yet. Use the Scan tab to begin." FontWeight="Bold"/>
                            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" HorizontalAlignment="Right">
                                <TextBlock Text="Sort by:" VerticalAlignment="Center" Margin="0,0,10,0" FontSize="11"/>
                                <ComboBox x:Name="SortCombo" Width="150" Height="25" SelectedIndex="0">
                                    <ComboBoxItem Content="Display Name" Tag="DisplayName"/>
                                    <ComboBoxItem Content="Size (Largest First)" Tag="SizeDesc"/>
                                    <ComboBoxItem Content="Size (Smallest First)" Tag="SizeAsc"/>
                                    <ComboBoxItem Content="% Used (Highest First)" Tag="PercentDesc"/>
                                    <ComboBoxItem Content="% Used (Lowest First)" Tag="PercentAsc"/>
                                </ComboBox>
                            </StackPanel>
                        </DockPanel>
                    </StackPanel>
                    
                    <Border Grid.Row="2" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15">
                        <DataGrid x:Name="ResultsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single" AlternatingRowBackground="#f8f9fa">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="200"/>
                                <DataGridTextColumn Header="Email" Binding="{Binding PrimarySmtpAddress}" Width="220"/>
                                <DataGridTextColumn Header="Type" Binding="{Binding MailboxType}" Width="100"/>
                                <DataGridTextColumn Header="Size (GB)" Binding="{Binding SizeGB}" Width="90"/>
                                <DataGridTextColumn Header="Items" Binding="{Binding ItemCount}" Width="80"/>
                                <DataGridTextColumn Header="Quota (GB)" Binding="{Binding QuotaGB}" Width="90"/>
                                <DataGridTextColumn Header="% Used" Binding="{Binding PercentUsed}" Width="80"/>
                                <DataGridTextColumn Header="Status" Binding="{Binding QuotaStatus}" Width="100"/>
                            </DataGrid.Columns>
                            <DataGrid.RowStyle>
                                <Style TargetType="DataGridRow">
                                    <Style.Triggers>
                                        <DataTrigger Binding="{Binding QuotaWarning}" Value="Critical">
                                            <Setter Property="Background" Value="#f8d7da"/>
                                            <Setter Property="Foreground" Value="#721c24"/>
                                        </DataTrigger>
                                        <DataTrigger Binding="{Binding QuotaWarning}" Value="Warning">
                                            <Setter Property="Background" Value="#fff3cd"/>
                                            <Setter Property="Foreground" Value="#856404"/>
                                        </DataTrigger>
                                    </Style.Triggers>
                                </Style>
                            </DataGrid.RowStyle>
                        </DataGrid>
                    </Border>
                    
                    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="ViewFoldersButton" Content="View Folders" Width="120" Height="32" Margin="0,0,10,0" Background="#17a2b8" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="ExportResultsButton" Content="Export to Excel" Width="130" Height="32" Margin="0,0,10,0" Background="#28a745" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="ResultsCloseButton" Content="Close" Width="80" Height="32" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="Folder Details" x:Name="DetailsTab">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <GroupBox Grid.Row="0" Header="Mailbox Information" Padding="10" Margin="0,0,0,15">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="20"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Grid.Column="0" Text="Display Name:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="0" Grid.Column="1" x:Name="DetailDisplayNameBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="0" Grid.Column="3" Text="Email:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="0" Grid.Column="4" x:Name="DetailEmailBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="1" Grid.Column="0" Text="Total Size:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="1" Grid.Column="1" x:Name="DetailSizeBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="1" Grid.Column="3" Text="Total Items:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="1" Grid.Column="4" x:Name="DetailItemsBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="2" Grid.Column="0" Text="Quota:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="2" Grid.Column="1" x:Name="DetailQuotaBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="2" Grid.Column="3" Text="% Used:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="2" Grid.Column="4" x:Name="DetailPercentBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                        </Grid>
                    </GroupBox>
                    
                    <GroupBox Grid.Row="1" Header="Folder Breakdown" Padding="10" Margin="0,0,0,15">
                        <Border BorderBrush="#dee2e6" BorderThickness="1">
                            <DataGrid x:Name="FoldersGrid" AutoGenerateColumns="False" IsReadOnly="True" AlternatingRowBackground="#f8f9fa">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Folder Name" Binding="{Binding FolderName}" Width="300"/>
                                    <DataGridTextColumn Header="Items" Binding="{Binding ItemCount}" Width="100"/>
                                    <DataGridTextColumn Header="Size (MB)" Binding="{Binding SizeMB}" Width="120"/>
                                    <DataGridTextColumn Header="% of Total" Binding="{Binding PercentOfTotal}" Width="100"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Border>
                    </GroupBox>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="ExportFoldersButton" Content="Export Folders" Width="130" Height="32" Margin="0,0,10,0" Background="#28a745" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="DetailsCloseButton" Content="Close" Width="80" Height="32" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@
    
    $statsReader = New-Object System.Xml.XmlNodeReader $MailboxStatsXAML
    $StatsWindow = [Windows.Markup.XamlReader]::Load($statsReader)
    $StatsWindow.Owner = $syncHash.Window
    
    # Scan Tab Controls
    $ScanAllRadio = $StatsWindow.FindName("ScanAllRadio")
    $ScanSpecificRadio = $StatsWindow.FindName("ScanSpecificRadio")
    $SpecificMailboxPanel = $StatsWindow.FindName("SpecificMailboxPanel")
    $SpecificMailboxBox = $StatsWindow.FindName("SpecificMailboxBox")
    
    $IncludeUserMailboxes = $StatsWindow.FindName("IncludeUserMailboxes")
    $IncludeSharedMailboxes = $StatsWindow.FindName("IncludeSharedMailboxes")
    $IncludeArchives = $StatsWindow.FindName("IncludeArchives")
    
    $ProgressGroup = $StatsWindow.FindName("ProgressGroup")
    $ProgressText = $StatsWindow.FindName("ProgressText")
    $ScanProgressBar = $StatsWindow.FindName("ScanProgressBar")
    $ProgressDetailText = $StatsWindow.FindName("ProgressDetailText")
    
    $StartScanButton = $StatsWindow.FindName("StartScanButton")
    $StopScanButton = $StatsWindow.FindName("StopScanButton")
    $ScanCloseButton = $StatsWindow.FindName("ScanCloseButton")
    
    # Results Tab Controls
    $ResultsTab = $StatsWindow.FindName("ResultsTab")
    $TotalMailboxesText = $StatsWindow.FindName("TotalMailboxesText")
    $TotalSizeText = $StatsWindow.FindName("TotalSizeText")
    $AverageSizeText = $StatsWindow.FindName("AverageSizeText")
    $OverQuotaText = $StatsWindow.FindName("OverQuotaText")
    $ResultsCountText = $StatsWindow.FindName("ResultsCountText")
    $SortCombo = $StatsWindow.FindName("SortCombo")
    $ResultsGrid = $StatsWindow.FindName("ResultsGrid")
    $ViewFoldersButton = $StatsWindow.FindName("ViewFoldersButton")
    $ExportResultsButton = $StatsWindow.FindName("ExportResultsButton")
    $ResultsCloseButton = $StatsWindow.FindName("ResultsCloseButton")
    
    # Details Tab Controls
    $DetailsTab = $StatsWindow.FindName("DetailsTab")
    $DetailDisplayNameBox = $StatsWindow.FindName("DetailDisplayNameBox")
    $DetailEmailBox = $StatsWindow.FindName("DetailEmailBox")
    $DetailSizeBox = $StatsWindow.FindName("DetailSizeBox")
    $DetailItemsBox = $StatsWindow.FindName("DetailItemsBox")
    $DetailQuotaBox = $StatsWindow.FindName("DetailQuotaBox")
    $DetailPercentBox = $StatsWindow.FindName("DetailPercentBox")
    $FoldersGrid = $StatsWindow.FindName("FoldersGrid")
    $ExportFoldersButton = $StatsWindow.FindName("ExportFoldersButton")
    $DetailsCloseButton = $StatsWindow.FindName("DetailsCloseButton")
    
    # Store current results
    $script:currentMailboxStats = $null
    $script:shouldStopScan = $false
    
    # Enable/disable specific mailbox panel
    $ScanAllRadio.Add_Checked({ $SpecificMailboxPanel.IsEnabled = $false })
    $ScanSpecificRadio.Add_Checked({ $SpecificMailboxPanel.IsEnabled = $true; $SpecificMailboxBox.Focus() })
    
    # Start Scan button
    $StartScanButton.Add_Click({
        $scanAll = $ScanAllRadio.IsChecked
        $specificMailbox = $SpecificMailboxBox.Text.Trim()
        
        # Validation
        if (-not $scanAll -and [string]::IsNullOrWhiteSpace($specificMailbox)) {
            [System.Windows.MessageBox]::Show("Please enter a mailbox email address", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

        if ($scanAll) {
            # Only check mailbox type filters when scanning all
            if (-not $IncludeUserMailboxes.IsChecked -and -not $IncludeSharedMailboxes.IsChecked -and -not $IncludeArchives.IsChecked) {
                [System.Windows.MessageBox]::Show("Please select at least one mailbox type to scan", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
        }
        
        if (-not $IncludeUserMailboxes.IsChecked -and -not $IncludeSharedMailboxes.IsChecked -and -not $IncludeArchives.IsChecked) {
            [System.Windows.MessageBox]::Show("Please select at least one mailbox type to scan", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        try {
            $script:shouldStopScan = $false
            $StartScanButton.IsEnabled = $false
            $StopScanButton.IsEnabled = $true
            $ProgressGroup.Visibility = [System.Windows.Visibility]::Visible
            $ScanProgressBar.Value = 0
            
            Write-Log "Starting mailbox size scan..."
            
            # Get mailboxes to scan
            $mailboxesToScan = @()
            
            if ($scanAll) {
                $ProgressText.Text = "Retrieving mailbox list..."
                $ProgressDetailText.Text = "This may take a moment for large organizations..."
                
                Write-Log "Retrieving all mailboxes..."
                
                $recipientTypeDetails = @()
                if ($IncludeUserMailboxes.IsChecked) { $recipientTypeDetails += "UserMailbox" }
                if ($IncludeSharedMailboxes.IsChecked) { $recipientTypeDetails += "SharedMailbox" }
                
                if ($recipientTypeDetails.Count -gt 0) {
                    $mailboxesToScan += @(Get-Mailbox -RecipientTypeDetails $recipientTypeDetails -ResultSize Unlimited -ErrorAction Stop)
                }
                
                if ($IncludeArchives.IsChecked) {
                    Write-Log "Including archive mailboxes..."
                    $archiveMailboxes = @(Get-Mailbox -Archive -ResultSize Unlimited -ErrorAction Stop)
                    $mailboxesToScan += $archiveMailboxes
                }
                
                Write-Log "Found $($mailboxesToScan.Count) mailboxes to scan"
                
            } else {
                $ProgressText.Text = "Retrieving mailbox information..."
                Write-Log "Retrieving specific mailbox: $specificMailbox"
                
                $mailboxesToScan = @(Get-Mailbox -Identity $specificMailbox -ErrorAction Stop)
            }
            
            if ($mailboxesToScan.Count -eq 0) {
                throw "No mailboxes found matching the criteria"
            }
            
            $ProgressText.Text = "Scanning $($mailboxesToScan.Count) mailbox(es)..."
            $ProgressDetailText.Text = "Processing mailbox statistics..."
            
            # Collect stats
            $results = @()
            $processedCount = 0
            $totalCount = $mailboxesToScan.Count
            
            foreach ($mbx in $mailboxesToScan) {
                if ($script:shouldStopScan) {
                    Write-Log "Scan stopped by user"
                    break
                }
                
                $processedCount++
                $percentComplete = [Math]::Round(($processedCount / $totalCount) * 100)
                $ScanProgressBar.Value = $percentComplete
                $ProgressDetailText.Text = "Processing $($mbx.DisplayName) ($processedCount of $totalCount)"
                
                # Force UI update
                [System.Windows.Forms.Application]::DoEvents()
                
                try {
                    Write-Log "Processing: $($mbx.DisplayName)"
                    
                    # Get mailbox statistics
                    $stats = Get-MailboxStatistics -Identity $mbx.Identity -ErrorAction Stop
                    
                    # Calculate size in GB
                    $sizeBytes = 0
                    if ($stats.TotalItemSize) {
                        $sizeString = $stats.TotalItemSize.ToString()
                        if ($sizeString -match '([0-9,]+)\s*bytes') {
                            $sizeBytes = [int64]($matches[1] -replace ',', '')
                        }
                    }
                    $sizeGB = [Math]::Round($sizeBytes / 1GB, 2)
                    
                    # Get quota information
                    $quotaGB = 0
                    $quotaStatus = "N/A"
                    $percentUsed = 0
                    $quotaWarning = "Normal"
                    
                    # Safely get ProhibitSendQuota
                    $prohibitQuota = $null
                    try {
                        $prohibitQuota = $mbx.ProhibitSendQuota
                    } catch {
                        Write-Log "Could not retrieve ProhibitSendQuota for $($mbx.DisplayName)"
                    }
                    
                    if ($null -ne $prohibitQuota -and $prohibitQuota -ne "Unlimited" -and $prohibitQuota.ToString() -ne "Unlimited") {
                        $quotaString = $prohibitQuota.ToString()
                        if ($quotaString -match '([0-9.]+)\s*GB') {
                            $quotaGB = [Math]::Round([decimal]$matches[1], 2)
                        } elseif ($quotaString -match '([0-9.]+)\s*MB') {
                            $quotaGB = [Math]::Round([decimal]$matches[1] / 1024, 2)
                        }
                        
                        if ($quotaGB -gt 0) {
                            $percentUsed = [Math]::Round(($sizeGB / $quotaGB) * 100, 1)
                            
                            if ($percentUsed -ge 95) {
                                $quotaStatus = "Critical"
                                $quotaWarning = "Critical"
                            } elseif ($percentUsed -ge 80) {
                                $quotaStatus = "Warning"
                                $quotaWarning = "Warning"
                            } else {
                                $quotaStatus = "Normal"
                            }
                        }
                    } else {
                        $quotaGB = "Unlimited"
                        $quotaStatus = "N/A"
                    }
                    
                    # Determine mailbox type
                    $mailboxType = switch ($mbx.RecipientTypeDetails) {
                        "UserMailbox" { "User" }
                        "SharedMailbox" { "Shared" }
                        "RoomMailbox" { "Room" }
                        "EquipmentMailbox" { "Equipment" }
                        default { if ($mbx.RecipientTypeDetails) { $mbx.RecipientTypeDetails.ToString() } else { "Unknown" } }
                    }
                    
                    # Check for archive
                    try {
                        if ($mbx.ArchiveStatus -eq "Active") {
                            $mailboxType += " (Archive)"
                        }
                    } catch {
                        # ArchiveStatus not available, skip
                    }
                    
                    # Safely get email address
                    $emailAddress = "N/A"
                    try {
                        if ($mbx.PrimarySmtpAddress) {
                            $emailAddress = $mbx.PrimarySmtpAddress.ToString()
                        } elseif ($mbx.EmailAddresses) {
                            $smtpAddr = $mbx.EmailAddresses | Where-Object { $_ -like "smtp:*" } | Select-Object -First 1
                            if ($smtpAddr) {
                                $emailAddress = $smtpAddr.ToString() -replace '^smtp:', ''
                            }
                        }
                    } catch {
                        Write-Log "Could not retrieve email address for $($mbx.DisplayName)"
                    }
                    
                    # Safely get last logon time
                    $lastLogon = "Never"
                    try {
                        if ($stats.LastLogonTime) {
                            $lastLogon = $stats.LastLogonTime.ToString("yyyy-MM-dd HH:mm")
                        }
                    } catch {
                        # LastLogonTime not available
                    }
                    
                    $results += [PSCustomObject]@{
                        DisplayName = if ($mbx.DisplayName) { $mbx.DisplayName } else { "Unknown" }
                        PrimarySmtpAddress = $emailAddress
                        MailboxType = $mailboxType
                        SizeGB = $sizeGB
                        SizeBytes = $sizeBytes
                        ItemCount = if ($stats.ItemCount) { $stats.ItemCount } else { 0 }
                        QuotaGB = $quotaGB
                        PercentUsed = if ($quotaGB -eq "Unlimited") { "N/A" } else { "$percentUsed%" }
                        PercentValue = $percentUsed
                        QuotaStatus = $quotaStatus
                        QuotaWarning = $quotaWarning
                        Identity = $mbx.Identity
                        LastLogonTime = $lastLogon
                    }
                    
                } catch {
                    Write-Log "Error processing $($mbx.DisplayName): $($_.Exception.Message)"
                    # Continue to next mailbox instead of failing entire scan
                }
            }
            
            # Check if we got any results
            if ($null -eq $results -or $results.Count -eq 0) {
                throw "No mailbox statistics could be retrieved"
            }
            
            Write-Log "Successfully scanned $($results.Count) mailboxes"
            
            # Sort results by size (largest first) - filter out any null SizeBytes first
            try {
                $results = @($results | Where-Object { $null -ne $_.SizeBytes } | Sort-Object -Property SizeBytes -Descending)
            } catch {
                Write-Log "Warning: Could not sort results, displaying unsorted"
            }
            
            # Store results
            $script:currentMailboxStats = $results
            
            # Calculate summary statistics with null protection
            $totalMailboxes = $results.Count
            
            $totalSizeGB = 0
            try {
                $sizeSum = $results | Where-Object { $null -ne $_.SizeGB } | Measure-Object -Property SizeGB -Sum
                if ($null -ne $sizeSum -and $null -ne $sizeSum.Sum) {
                    $totalSizeGB = [Math]::Round($sizeSum.Sum, 2)
                }
            } catch {
                Write-Log "Warning: Could not calculate total size"
            }
            
            $avgSizeGB = 0
            try {
                if ($totalMailboxes -gt 0) {
                    $avgSizeGB = [Math]::Round($totalSizeGB / $totalMailboxes, 2)
                }
            } catch {
                Write-Log "Warning: Could not calculate average size"
            }
            
            $overQuotaCount = 0
            try {
                $overQuotaCount = @($results | Where-Object { 
                    $null -ne $_.PercentValue -and 
                    $_.PercentValue -ge 80 -and 
                    $_.PercentValue -ne 0 
                }).Count
            } catch {
                Write-Log "Warning: Could not calculate over-quota count"
            }
            
            # Update summary statistics
            try {
                $TotalMailboxesText.Text = $totalMailboxes.ToString()
                $TotalSizeText.Text = "$totalSizeGB GB"
                $AverageSizeText.Text = "$avgSizeGB GB"
                $OverQuotaText.Text = $overQuotaCount.ToString()
            } catch {
                Write-Log "Warning: Could not update summary text: $($_.Exception.Message)"
            }
            
            # Display results in grid
            try {
                $ResultsGrid.ItemsSource = $null
                $ResultsGrid.ItemsSource = $results
                $ResultsCountText.Text = "Showing $($results.Count) mailbox(es)"
                $ExportResultsButton.IsEnabled = $true
            } catch {
                Write-Log "Error binding results to grid: $($_.Exception.Message)"
                throw "Could not display results: $($_.Exception.Message)"
            }
            
            # Update progress
            $ProgressText.Text = "Scan complete!"
            $ProgressDetailText.Text = "Found $($results.Count) mailboxes. Total size: $totalSizeGB GB"
            $ScanProgressBar.Value = 100
            
            # Switch to Results tab
            $ResultsTab.IsSelected = $true

            
        } catch {
            Write-Log "Error during scan: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error scanning mailboxes:`n`n$($_.Exception.Message)",
                "Scan Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            $ProgressText.Text = "Scan failed"
            $ProgressDetailText.Text = $_.Exception.Message
        } finally {
            $StartScanButton.IsEnabled = $true
            $StopScanButton.IsEnabled = $false
        }
    })
    
    # Stop Scan button
    $StopScanButton.Add_Click({
        $script:shouldStopScan = $true
        $StopScanButton.IsEnabled = $false
        $ProgressText.Text = "Stopping scan..."
        Write-Log "Stop scan requested"
    })
    
    # Sort combo changed
    $SortCombo.Add_SelectionChanged({
        if ($null -eq $script:currentMailboxStats) { return }
        
        $sortTag = $SortCombo.SelectedItem.Tag
        $sorted = $null
        
        switch ($sortTag) {
            "DisplayName" { $sorted = $script:currentMailboxStats | Sort-Object DisplayName }
            "SizeDesc" { $sorted = $script:currentMailboxStats | Sort-Object -Property SizeBytes -Descending }
            "SizeAsc" { $sorted = $script:currentMailboxStats | Sort-Object -Property SizeBytes }
            "PercentDesc" { $sorted = $script:currentMailboxStats | Sort-Object -Property PercentValue -Descending }
            "PercentAsc" { $sorted = $script:currentMailboxStats | Sort-Object -Property PercentValue }
        }
        
        $ResultsGrid.ItemsSource = $sorted
    })
    
    # Results grid selection changed
    $ResultsGrid.Add_SelectionChanged({
        $ViewFoldersButton.IsEnabled = ($null -ne $ResultsGrid.SelectedItem)
    })
    
    # Double-click to view folders
    $ResultsGrid.Add_MouseDoubleClick({
        if ($ResultsGrid.SelectedItem) {
            $ViewFoldersButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    # View Folders button
    $ViewFoldersButton.Add_Click({
        $selected = $ResultsGrid.SelectedItem
        if ($null -eq $selected) { return }
        
        try {
            $ViewFoldersButton.IsEnabled = $false
            Write-Log "Loading folder statistics for: $($selected.DisplayName)"
            
            # Get folder statistics
            $folderStats = @(Get-MailboxFolderStatistics -Identity $selected.Identity -ErrorAction Stop)
            
            Write-Log "Retrieved $($folderStats.Count) folders"
            
            # Populate mailbox info
            $DetailDisplayNameBox.Text = $selected.DisplayName
            $DetailEmailBox.Text = $selected.PrimarySmtpAddress
            $DetailSizeBox.Text = "$($selected.SizeGB) GB"
            $DetailItemsBox.Text = $selected.ItemCount.ToString("N0")
            $DetailQuotaBox.Text = if ($selected.QuotaGB -eq "Unlimited") { "Unlimited" } else { "$($selected.QuotaGB) GB" }
            $DetailPercentBox.Text = $selected.PercentUsed
            
            # Format folder data
            $folderList = @()
            $totalSizeBytes = $selected.SizeBytes
            
            foreach ($folder in $folderStats) {
                $folderSizeBytes = 0
                if ($folder.FolderSize) {
                    $sizeString = $folder.FolderSize.ToString()
                    if ($sizeString -match '([0-9,]+)\s*bytes') {
                        $folderSizeBytes = [int64]($matches[1] -replace ',', '')
                    }
                }
                
                $folderSizeMB = [Math]::Round($folderSizeBytes / 1MB, 2)
                $percentOfTotal = if ($totalSizeBytes -gt 0) { 
                    [Math]::Round(($folderSizeBytes / $totalSizeBytes) * 100, 1) 
                } else { 
                    0 
                }
                
                $folderList += [PSCustomObject]@{
                    FolderName = $folder.Name
                    ItemCount = $folder.ItemsInFolder
                    SizeMB = $folderSizeMB
                    SizeBytes = $folderSizeBytes
                    PercentOfTotal = "$percentOfTotal%"
                }
            }
            
            # Sort by size (largest first)
            $folderList = $folderList | Sort-Object -Property SizeBytes -Descending
            
            $FoldersGrid.ItemsSource = $folderList
            $ExportFoldersButton.IsEnabled = $true
            
            # Switch to Details tab
            $DetailsTab.IsSelected = $true
            
        } catch {
            Write-Log "Error loading folder statistics: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error loading folder statistics:`n`n$($_.Exception.Message)",
                "Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        } finally {
            $ViewFoldersButton.IsEnabled = $true
        }
    })
    
    # Export Results to Excel
    $ExportResultsButton.Add_Click({
        if ($null -eq $script:currentMailboxStats -or $script:currentMailboxStats.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No results to export", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            [System.Windows.MessageBox]::Show("ImportExcel module is not installed. Cannot export to Excel.`n`nPlease run: Install-Module -Name ImportExcel", "Module Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Write-Log "Export failed: ImportExcel module not installed"
            return
        }
        
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.Title = "Save Mailbox Size Report"
            $saveDialog.FileName = "MailboxSizeReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $ExportResultsButton.IsEnabled = $false
                Write-Log "Exporting mailbox statistics to Excel: $excelPath"
                
                $exportData = @()
                foreach ($mbx in $script:currentMailboxStats) {
                    $exportData += [PSCustomObject]@{
                        'Display Name' = $mbx.DisplayName
                        'Email Address' = $mbx.PrimarySmtpAddress
                        'Mailbox Type' = $mbx.MailboxType
                        'Size (GB)' = $mbx.SizeGB
                        'Item Count' = $mbx.ItemCount
                        'Quota (GB)' = $mbx.QuotaGB
                        'Percent Used' = $mbx.PercentUsed
                        'Quota Status' = $mbx.QuotaStatus
                        'Last Logon' = $mbx.LastLogonTime
                        'Export Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                # Export with conditional formatting
                $exportData | Export-Excel -Path $excelPath `
                    -WorksheetName "Mailbox Sizes" `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle Medium1 `
                    -TableName "MailboxSizes"
                
                Write-Log "Successfully exported $($exportData.Count) mailboxes to Excel"
                
                [System.Windows.MessageBox]::Show(
                    "Mailbox size report exported successfully!`n`nFile: $excelPath",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                ) | Out-Null
                
                # Open folder location in Explorer
                $folderPath = Split-Path $excelPath -Parent
                Invoke-Item $folderPath
                
            } else {
                Write-Log "Export cancelled by user"
            }
            
        } catch {
            Write-Log "Export error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error exporting to Excel:`n`n$($_.Exception.Message)",
                "Export Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        } finally {
            $ExportResultsButton.IsEnabled = $true
        }
    })
    
    # Export Folders to Excel
    $ExportFoldersButton.Add_Click({
        $folders = $FoldersGrid.ItemsSource
        if ($null -eq $folders -or $folders.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No folder data to export", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            [System.Windows.MessageBox]::Show("ImportExcel module is not installed.", "Module Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }
        
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.Title = "Save Folder Breakdown Report"
            $mailboxName = $DetailDisplayNameBox.Text -replace '[\\/:*?"<>|]', '_'
            $saveDialog.FileName = "FolderBreakdown_${mailboxName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $exportData = @()
                foreach ($folder in $folders) {
                    $exportData += [PSCustomObject]@{
                        'Mailbox' = $DetailDisplayNameBox.Text
                        'Email' = $DetailEmailBox.Text
                        'Folder Name' = $folder.FolderName
                        'Item Count' = $folder.ItemCount
                        'Size (MB)' = $folder.SizeMB
                        'Percent of Total' = $folder.PercentOfTotal
                        'Export Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                $exportData | Export-Excel -Path $excelPath `
                    -WorksheetName "Folder Breakdown" `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle Medium1 `
                    -TableName "FolderBreakdown"
                
                Write-Log "Exported folder breakdown to Excel"
                
                [System.Windows.MessageBox]::Show(
                    "Folder breakdown exported successfully!`n`nFile: $excelPath",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                ) | Out-Null
                
                # Open folder location in Explorer
                $folderPath = Split-Path $excelPath -Parent
                Invoke-Item $folderPath
            }
            
        } catch {
            Write-Log "Export error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error exporting: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    })
    
    # Close buttons
    $ScanCloseButton.Add_Click({ $StatsWindow.Close() })
    $ResultsCloseButton.Add_Click({ $StatsWindow.Close() })
    $DetailsCloseButton.Add_Click({ $StatsWindow.Close() })
    
    # Enter key support
    $SpecificMailboxBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $StartScanButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    $StatsWindow.ShowDialog() | Out-Null
})

$syncHash.IPScannerButton.Add_Click({
    Write-Log "Opening IP Network Scanner"
    
    [xml]$ScannerXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="IP Network Scanner" 
        Height="700" 
        Width="1000" 
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#233A4A" Padding="20,15">
            <StackPanel>
                <TextBlock Text="IP Network Scanner" FontSize="20" FontWeight="Bold" Foreground="White"/>
                <TextBlock Text="Scan IP ranges to discover active devices on the network" FontSize="12" Foreground="#B0BEC5" Margin="0,5,0,0"/>
            </StackPanel>
        </Border>
        
        <!-- Main Content -->
        <Grid Grid.Row="1" Margin="20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            
            <!-- Scan Configuration -->
            <GroupBox Grid.Row="0" Header="Scan Configuration" Padding="15" Margin="0,0,0,15">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    
                    <TextBlock Grid.Column="0" Text="Start IP:" VerticalAlignment="Center" Margin="0,0,10,0" FontWeight="Bold"/>
                    <TextBox x:Name="StartIPBox" Grid.Column="1" Height="30" VerticalContentAlignment="Center" Padding="5" Margin="0,0,20,0"/>
                    
                    <TextBlock Grid.Column="2" Text="End IP:" VerticalAlignment="Center" Margin="0,0,10,0" FontWeight="Bold"/>
                    <TextBox x:Name="EndIPBox" Grid.Column="3" Height="30" VerticalContentAlignment="Center" Padding="5" Margin="0,0,20,0"/>
                    
                    <Button x:Name="StartScanButton" Grid.Column="4" Content="Start Scan" Width="120" Height="35" Background="#28a745" Foreground="White" FontWeight="Bold"/>
                </Grid>
            </GroupBox>
            
            <!-- Results -->
            <GroupBox Grid.Row="1" Header="Scan Results" Padding="10">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Progress and Stats -->
                    <StackPanel Grid.Row="0" Margin="0,0,0,10">
                        <TextBlock x:Name="ScanStatusText" Text="Ready to scan" FontWeight="Bold" Margin="0,0,0,5"/>
                        <ProgressBar x:Name="ScanProgressBar" Height="20" Minimum="0" Maximum="100" Value="0"/>
                        <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                            <TextBlock Text="Total Scanned: " FontWeight="Bold"/>
                            <TextBlock x:Name="TotalScannedText" Text="0" Margin="0,0,20,0"/>
                            <TextBlock Text="Online: " FontWeight="Bold" Foreground="Green"/>
                            <TextBlock x:Name="OnlineCountText" Text="0" Foreground="Green" Margin="0,0,20,0"/>
                            <TextBlock Text="Offline: " FontWeight="Bold" Foreground="Red"/>
                            <TextBlock x:Name="OfflineCountText" Text="0" Foreground="Red"/>
                        </StackPanel>
                    </StackPanel>
                    
                    <!-- Results Grid -->
                    <DataGrid Grid.Row="1" 
                            x:Name="ResultsGrid" 
                            AutoGenerateColumns="False" 
                            IsReadOnly="True"
                            SelectionMode="Extended"
                            GridLinesVisibility="All"
                            AlternatingRowBackground="#f8f9fa"
                            CanUserSortColumns="True"
                            CanUserResizeColumns="True">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="IP Address" Binding="{Binding IPAddress}" Width="150"/>
                            <DataGridTextColumn Header="Hostname" Binding="{Binding Hostname}" Width="*"/>
                            <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="100">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Style.Triggers>
                                            <Trigger Property="Text" Value="Online">
                                                <Setter Property="Foreground" Value="Green"/>
                                                <Setter Property="FontWeight" Value="Bold"/>
                                            </Trigger>
                                            <Trigger Property="Text" Value="Offline">
                                                <Setter Property="Foreground" Value="Red"/>
                                            </Trigger>
                                        </Style.Triggers>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn Header="MAC Address" Binding="{Binding MACAddress}" Width="150"/>
                            <DataGridTextColumn Header="Response Time (ms)" Binding="{Binding ResponseTime}" Width="150"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </GroupBox>
        </Grid>
        
        <!-- Footer -->
        <Border Grid.Row="2" Background="#f8f9fa" BorderBrush="#dee2e6" BorderThickness="0,1,0,0" Padding="15">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="ExportButton" Content="Export to Excel" Width="120" Height="30" Margin="0,0,10,0" Background="#17a2b8" Foreground="White" IsEnabled="False"/>
                <Button x:Name="CloseButton" Content="Close" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    
    $scanReader = New-Object System.Xml.XmlNodeReader $ScannerXAML
    $ScanWindow = [Windows.Markup.XamlReader]::Load($scanReader)
    $ScanWindow.Owner = $syncHash.Window
    
    # Get controls
    $StartIPBox = $ScanWindow.FindName("StartIPBox")
    $EndIPBox = $ScanWindow.FindName("EndIPBox")
    $StartScanButton = $ScanWindow.FindName("StartScanButton")
    $ScanStatusText = $ScanWindow.FindName("ScanStatusText")
    $ScanProgressBar = $ScanWindow.FindName("ScanProgressBar")
    $TotalScannedText = $ScanWindow.FindName("TotalScannedText")
    $OnlineCountText = $ScanWindow.FindName("OnlineCountText")
    $OfflineCountText = $ScanWindow.FindName("OfflineCountText")
    $ResultsGrid = $ScanWindow.FindName("ResultsGrid")
    $ExportButton = $ScanWindow.FindName("ExportButton")
    $CloseButton = $ScanWindow.FindName("CloseButton")
    
    # Function to validate IP address
    function Test-IPAddress {
        param([string]$IP)
        try {
            $null = [System.Net.IPAddress]::Parse($IP)
            return $true
        } catch {
            return $false
        }
    }
    
    # Function to convert IP to integer for range calculation
    function ConvertTo-IPInteger {
        param([string]$IP)
        $bytes = [System.Net.IPAddress]::Parse($IP).GetAddressBytes()
        [Array]::Reverse($bytes)
        return [BitConverter]::ToUInt32($bytes, 0)
    }
    
    # Function to convert integer back to IP
    function ConvertFrom-IPInteger {
        param([uint32]$Int)
        $bytes = [BitConverter]::GetBytes($Int)
        [Array]::Reverse($bytes)
        return [System.Net.IPAddress]::new($bytes).ToString()
    }
    
    # Function to get MAC address from ARP table
    function Get-MACFromARP {
        param([string]$IP)
        try {
            $arpResult = arp -a $IP 2>$null
            $macLine = $arpResult | Where-Object { $_ -match $IP }
            if ($macLine -match '([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})') {
                return $matches[0].ToUpper()
            }
        } catch {
            # Silently fail
        }
        return "N/A"
    }
    
    # Start Scan Button Click
    $StartScanButton.Add_Click({
        $startIP = $StartIPBox.Text.Trim()
        $endIP = $EndIPBox.Text.Trim()
        
        # Validate IPs
        if (-not (Test-IPAddress $startIP)) {
            [System.Windows.MessageBox]::Show("Invalid start IP address", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Test-IPAddress $endIP)) {
            [System.Windows.MessageBox]::Show("Invalid end IP address", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        Write-Log "Starting IP scan from $startIP to $endIP"
        
        # Disable scan button during scan
        $StartScanButton.IsEnabled = $false
        $ExportButton.IsEnabled = $false
        
        # Clear previous results
        $ResultsGrid.ItemsSource = $null
        $results = [System.Collections.Generic.List[object]]::new()
        
        # Convert IPs to integers for range
        $startInt = ConvertTo-IPInteger $startIP
        $endInt = ConvertTo-IPInteger $endIP
        
        if ($startInt -gt $endInt) {
            [System.Windows.MessageBox]::Show("Start IP must be less than or equal to End IP", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            $StartScanButton.IsEnabled = $true
            return
        }
        
        $totalIPs = $endInt - $startInt + 1
        
        if ($totalIPs -gt 1000) {
            $response = [System.Windows.MessageBox]::Show(
                "You are about to scan $totalIPs IP addresses. This may take a while. Continue?",
                "Large Scan Warning",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )
            if ($response -eq [System.Windows.MessageBoxResult]::No) {
                $StartScanButton.IsEnabled = $true
                return
            }
        }
        
        # Reset counters
        $onlineCount = 0
        $offlineCount = 0
        $scannedCount = 0
        
        $ScanStatusText.Text = "Scanning in progress..."
        $ScanProgressBar.Value = 0
        $TotalScannedText.Text = "0"
        $OnlineCountText.Text = "0"
        $OfflineCountText.Text = "0"
        
        # Create runspace pool for parallel scanning (50 concurrent threads)
        $runspacePool = [runspacefactory]::CreateRunspacePool(1, 50)
        $runspacePool.Open()
        $runspaces = New-Object System.Collections.ArrayList
        
        # Script block for each IP scan
        $scanScriptBlock = {
            param($IP)
            
            $result = [PSCustomObject]@{
                IPAddress = $IP
                Hostname = "N/A"
                Status = "Offline"
                MACAddress = "N/A"
                ResponseTime = "N/A"
            }
            
            try {
                # Test connection with BufferSize for faster ping
                $pingResult = Test-Connection -ComputerName $IP -Count 1 -BufferSize 32 -Quiet -ErrorAction Stop
                
                if ($pingResult) {
                    $result.Status = "Online"
                    
                    # Get response time
                    try {
                        $pingDetail = Test-Connection -ComputerName $IP -Count 1 -BufferSize 32 -ErrorAction Stop
                        if ($pingDetail.ResponseTime) {
                            $result.ResponseTime = $pingDetail.ResponseTime.ToString()
                        } elseif ($pingDetail.Latency) {
                            $result.ResponseTime = $pingDetail.Latency.ToString()
                        }
                    } catch {
                        $result.ResponseTime = "<1"
                    }
                    
                    # Try to resolve hostname
                    try {
                        $hostEntry = [System.Net.Dns]::GetHostEntry($IP)
                        $result.Hostname = $hostEntry.HostName
                    } catch {
                        $result.Hostname = "Unable to resolve"
                    }
                    
                    # Ping once more to populate ARP table
                    $null = Test-NetConnection -ComputerName $IP -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue 2>$null
                    
                    # Try to get MAC address from ARP
                    try {
                        Start-Sleep -Milliseconds 50
                        $arpResult = arp -a $IP 2>$null
                        if ($arpResult) {
                            $macLine = $arpResult | Where-Object { $_ -match $IP }
                            if ($macLine -match '([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})') {
                                $result.MACAddress = $matches[0].ToUpper()
                            }
                        }
                    } catch {
                        # Silently fail
                    }
                }
            } catch {
                # IP is offline or unreachable
            }
            
            return $result
        }
        
        # Launch all scan jobs
        for ($i = $startInt; $i -le $endInt; $i++) {
            $currentIP = ConvertFrom-IPInteger $i
            
            $powershell = [powershell]::Create().AddScript($scanScriptBlock).AddArgument($currentIP)
            $powershell.RunspacePool = $runspacePool
            
            [void]$runspaces.Add([PSCustomObject]@{
                Pipe = $powershell
                Status = $powershell.BeginInvoke()
            })
        }
        
        # Monitor progress and collect results
        $completed = 0
        while ($runspaces.Status.IsCompleted -contains $false) {
            $completedNow = ($runspaces.Status.IsCompleted -eq $true).Count
            
            if ($completedNow -gt $completed) {
                $completed = $completedNow
                $percentComplete = [math]::Round(($completed / $totalIPs) * 100)
                $ScanProgressBar.Value = $percentComplete
                $TotalScannedText.Text = $completed.ToString()
                $ScanStatusText.Text = "Scanning in progress... ($completed of $totalIPs)"
                $ScanWindow.Dispatcher.Invoke([Action]{}, [Windows.Threading.DispatcherPriority]::Background)
            }
            
            Start-Sleep -Milliseconds 100
        }
        
        # Collect all results
        foreach ($runspace in $runspaces) {
            $result = $runspace.Pipe.EndInvoke($runspace.Status)
            $results.Add($result)
            
            if ($result.Status -eq "Online") {
                $onlineCount++
            } else {
                $offlineCount++
            }
            
            $runspace.Pipe.Dispose()
        }
        
        # Cleanup
        $runspacePool.Close()
        $runspacePool.Dispose()
        
        # Update final counts
        $OnlineCountText.Text = $onlineCount.ToString()
        $OfflineCountText.Text = $offlineCount.ToString()
        
        # Update grid with results
        $ResultsGrid.ItemsSource = $results
        
        # Complete
        $ScanStatusText.Text = "Scan complete! Found $onlineCount online devices out of $totalIPs addresses scanned."
        $ScanProgressBar.Value = 100
        $StartScanButton.IsEnabled = $true
        $ExportButton.IsEnabled = $true
        
        Write-Log "IP scan complete: $onlineCount online, $offlineCount offline"
    })
    
    # Export to Excel
    $ExportButton.Add_Click({
        if ($null -eq $ResultsGrid.ItemsSource -or $ResultsGrid.Items.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No scan results to export", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            [System.Windows.MessageBox]::Show("ImportExcel module is not installed.", "Module Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Write-Log "Export failed: ImportExcel module not installed"
            return
        }
        
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.Title = "Save IP Scan Results"
            $saveDialog.FileName = "IPScan_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $ExportButton.IsEnabled = $false
                Write-Log "Exporting IP scan results to Excel: $excelPath"
                
                $exportData = @()
                foreach ($item in $ResultsGrid.ItemsSource) {
                    $exportData += [PSCustomObject]@{
                        'IP Address' = $item.IPAddress
                        'Hostname' = $item.Hostname
                        'Status' = $item.Status
                        'MAC Address' = $item.MACAddress
                        'Response Time (ms)' = $item.ResponseTime
                        'Scan Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                # Export with formatting
                $exportData | Export-Excel -Path $excelPath `
                    -WorksheetName "IP Scan Results" `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle Medium1 `
                    -TableName "IPScanResults"
                
                Write-Log "Successfully exported $($exportData.Count) IP scan results to Excel"
                
                [System.Windows.MessageBox]::Show(
                    "IP scan results exported successfully!`n`nFile: $excelPath",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                ) | Out-Null
                
                # Open folder location in Explorer
                $folderPath = Split-Path $excelPath -Parent
                Invoke-Item $folderPath
                
            } else {
                Write-Log "Export cancelled by user"
            }
            
        } catch {
            Write-Log "Export error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error exporting to Excel:`n`n$($_.Exception.Message)",
                "Export Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        } finally {
            $ExportButton.IsEnabled = $true
        }
    })
    
    # Close button
    $CloseButton.Add_Click({ $ScanWindow.Close() })
    
    # Enter key support for IP boxes
    $StartIPBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $StartScanButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    $EndIPBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $StartScanButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    # Set default IPs for convenience (can be customized)
    $StartIPBox.Text = "192.168.1.1"
    $EndIPBox.Text = "192.168.1.254"
    
    $ScanWindow.ShowDialog() | Out-Null
})



# =====================================================
# INTUNE MOBILE DEVICES MODULE
# =====================================================

function Show-IntuneMobileDevicesWindow {
    $xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Intune Mobile Devices" 
    Height="700" 
    Width="1200" 
    WindowStartupLocation="CenterScreen"
    Background="#F5F5F5">
    
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        
        <Style TargetType="DataGrid">
            <Setter Property="AutoGenerateColumns" Value="True"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="CanUserDeleteRows" Value="False"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="SelectionMode" Value="Extended"/>
            <Setter Property="GridLinesVisibility" Value="None"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="RowBackground" Value="White"/>
            <Setter Property="AlternatingRowBackground" Value="#F9F9F9"/>
            <Setter Property="BorderBrush" Value="#E0E0E0"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <StackPanel Grid.Row="0" Margin="0,0,0,15">
            <TextBlock Text="Intune Mobile Devices" FontSize="24" FontWeight="Bold" Foreground="#333"/>
            <TextBlock Text="View and export all mobile devices managed by Microsoft Intune" 
                      FontSize="13" Foreground="#666" Margin="0,5,0,0"/>
        </StackPanel>
        
        <!-- Control Panel -->
        <Border Grid.Row="1" Background="White" BorderBrush="#E0E0E0" BorderThickness="1" 
                Padding="15" Margin="0,0,0,15" CornerRadius="4">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0" Orientation="Horizontal">
                    <Button x:Name="LoadDevicesButton" Content="Load Devices" Width="130"/>
                    <Button x:Name="RefreshButton" Content="Refresh" Width="100" Background="#28a745"/>
                    <TextBlock x:Name="StatusText" Text="Ready" VerticalAlignment="Center" 
                              Margin="15,0,0,0" FontSize="13" Foreground="#666"/>
                </StackPanel>
                
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <TextBlock Text="Total Devices:" VerticalAlignment="Center" Margin="0,0,10,0" FontWeight="SemiBold"/>
                    <TextBlock x:Name="DeviceCountText" Text="0" VerticalAlignment="Center" 
                              FontSize="16" FontWeight="Bold" Foreground="#0078D4" Margin="0,0,20,0"/>
                    <Button x:Name="ExportButton" Content="Export to Excel" Width="140" Background="#28a745"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Summary Stats -->
        <Border Grid.Row="2" Background="White" BorderBrush="#E0E0E0" BorderThickness="1" 
                CornerRadius="4" Margin="0,0,0,15">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <!-- Stats Bar -->
                <Border Grid.Row="0" Background="#F8F9FA" BorderBrush="#E0E0E0" 
                       BorderThickness="0,0,0,1" Padding="15,10">
                    <StackPanel Orientation="Horizontal">
                        <StackPanel Orientation="Horizontal" Margin="0,0,30,0">
                            <TextBlock Text="iOS/iPadOS: " FontWeight="SemiBold" Foreground="#666"/>
                            <TextBlock x:Name="IosCountText" Text="0" FontWeight="Bold" Foreground="#0078D4"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,30,0">
                            <TextBlock Text="Android: " FontWeight="SemiBold" Foreground="#666"/>
                            <TextBlock x:Name="AndroidCountText" Text="0" FontWeight="Bold" Foreground="#3DDC84"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,30,0">
                            <TextBlock Text="Compliant: " FontWeight="SemiBold" Foreground="#666"/>
                            <TextBlock x:Name="CompliantCountText" Text="0" FontWeight="Bold" Foreground="#28a745"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Non-Compliant: " FontWeight="SemiBold" Foreground="#666"/>
                            <TextBlock x:Name="NonCompliantCountText" Text="0" FontWeight="Bold" Foreground="#dc3545"/>
                        </StackPanel>
                    </StackPanel>
                </Border>
                
                <!-- DataGrid -->
                <DataGrid x:Name="DevicesDataGrid" Grid.Row="1" Margin="15"/>
            </Grid>
        </Border>
        
        <!-- Close Button -->
        <Button Grid.Row="3" x:Name="CloseButton" Content="Close" Width="100" 
                HorizontalAlignment="Right" Background="#6c757d"/>
    </Grid>
</Window>
"@
    
    try {
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
        $window = [Windows.Markup.XamlReader]::Load($reader)
        $reader.Close()
        
        # Get controls
        $LoadDevicesButton = $window.FindName("LoadDevicesButton")
        $RefreshButton = $window.FindName("RefreshButton")
        $ExportButton = $window.FindName("ExportButton")
        $CloseButton = $window.FindName("CloseButton")
        $DevicesDataGrid = $window.FindName("DevicesDataGrid")
        $StatusText = $window.FindName("StatusText")
        $DeviceCountText = $window.FindName("DeviceCountText")
        $IosCountText = $window.FindName("IosCountText")
        $AndroidCountText = $window.FindName("AndroidCountText")
        $CompliantCountText = $window.FindName("CompliantCountText")
        $NonCompliantCountText = $window.FindName("NonCompliantCountText")
        
        # Script-level variable to store device data
        $script:intuneDevices = @()
        
        # Function to update statistics
        function Update-DeviceStats {
            $total = $script:intuneDevices.Count
            $DeviceCountText.Text = $total.ToString()
            
            if ($total -gt 0) {
                $iosCount = ($script:intuneDevices | Where-Object { $_.'Operating System' -in @('iOS', 'iPadOS') }).Count
                $androidCount = ($script:intuneDevices | Where-Object { $_.'Operating System' -eq 'Android' }).Count
                $compliantCount = ($script:intuneDevices | Where-Object { $_.'Compliance State' -eq 'compliant' }).Count
                $nonCompliantCount = ($script:intuneDevices | Where-Object { $_.'Compliance State' -eq 'noncompliant' }).Count
                
                $IosCountText.Text = $iosCount.ToString()
                $AndroidCountText.Text = $androidCount.ToString()
                $CompliantCountText.Text = $compliantCount.ToString()
                $NonCompliantCountText.Text = $nonCompliantCount.ToString()
            }
        }
        
        # Load Devices function
        function Load-IntuneDevices {
            try {
                $LoadDevicesButton.IsEnabled = $false
                $RefreshButton.IsEnabled = $false
                $StatusText.Text = "Connecting to Microsoft Graph..."
                $StatusText.Foreground = "#FF9800"
                
                Write-Log "Connecting to Microsoft Graph for Intune devices..."
                
                # Check if Microsoft.Graph modules are installed
                $requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.DeviceManagement')
                foreach ($module in $requiredModules) {
                    if (-not (Get-Module -ListAvailable -Name $module)) {
                        $StatusText.Text = "Installing $module module..."
                        Write-Log "Installing $module..."
                        Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
                    }
                }
                
                # Import modules
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
                Import-Module Microsoft.Graph.DeviceManagement -ErrorAction Stop
                
                # Check if already connected
                $context = Get-MgContext -ErrorAction SilentlyContinue
                if ($null -eq $context) {
                    # Connect to Microsoft Graph
                    Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All", "User.Read.All" -NoWelcome
                    Write-Log "Connected to Microsoft Graph"
                }
                
                $StatusText.Text = "Retrieving devices from Intune..."
                Write-Log "Retrieving managed devices..."
                
                # Get all managed devices using pagination
                $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
                $allDevices = @()
                
                do {
                    $response = Invoke-MgGraphRequest -Uri $uri -Method GET
                    $allDevices += $response.value
                    $uri = $response.'@odata.nextLink'
                    
                    $StatusText.Text = "Retrieved $($allDevices.Count) devices..."
                } while ($uri)
                
                Write-Log "Retrieved $($allDevices.Count) total devices"
                
                # Filter for mobile devices only
                $mobileDevices = $allDevices | Where-Object { 
                    $_.operatingSystem -in @('iOS', 'Android', 'iPadOS') 
                }
                
                Write-Log "Found $($mobileDevices.Count) mobile devices"
                
                $StatusText.Text = "Processing device information..."
                
                # Format device information
                $script:intuneDevices = foreach ($device in $mobileDevices) {
                    [PSCustomObject]@{
                        'Device Name' = $device.deviceName
                        'User Display Name' = $device.userDisplayName
                        'User Principal Name' = $device.userPrincipalName
                        'Operating System' = $device.operatingSystem
                        'OS Version' = $device.osVersion
                        'Model' = $device.model
                        'Manufacturer' = $device.manufacturer
                        'IMEI' = $device.imei
                        'Serial Number' = $device.serialNumber
                        'Phone Number' = $device.phoneNumber
                        'Enrollment Date' = if ($device.enrolledDateTime) { 
                            (Get-Date $device.enrolledDateTime).ToString('yyyy-MM-dd HH:mm:ss') 
                        } else { 'N/A' }
                        'Last Sync' = if ($device.lastSyncDateTime) { 
                            (Get-Date $device.lastSyncDateTime).ToString('yyyy-MM-dd HH:mm:ss') 
                        } else { 'N/A' }
                        'Compliance State' = $device.complianceState
                        'Management State' = $device.managementState
                        'Ownership' = $device.managedDeviceOwnerType
                        'Supervised' = $device.isSupervised
                        'Encrypted' = $device.isEncrypted
                        'Jail Broken' = $device.jailBroken
                        'Total Storage (GB)' = if ($device.totalStorageSpaceInBytes) { 
                            [math]::Round($device.totalStorageSpaceInBytes / 1GB, 2) 
                        } else { 'N/A' }
                        'Free Storage (GB)' = if ($device.freeStorageSpaceInBytes) { 
                            [math]::Round($device.freeStorageSpaceInBytes / 1GB, 2) 
                        } else { 'N/A' }
                    }
                }
                
                # Update DataGrid
                $DevicesDataGrid.ItemsSource = $script:intuneDevices
                
                # Update statistics
                Update-DeviceStats
                
                $StatusText.Text = "Loaded $($script:intuneDevices.Count) mobile devices"
                $StatusText.Foreground = "#28a745"
                Write-Log "Successfully loaded $($script:intuneDevices.Count) mobile devices"
                
                $ExportButton.IsEnabled = $true
                
            } catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Error loading Intune devices: $errorMsg"
                $StatusText.Text = "Error: $errorMsg"
                $StatusText.Foreground = "#dc3545"
                [System.Windows.MessageBox]::Show(
                    "Failed to load Intune devices:`n`n$errorMsg",
                    "Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
            } finally {
                $LoadDevicesButton.IsEnabled = $true
                $RefreshButton.IsEnabled = $true
            }
        }
        
        # Export to Excel function
        function Export-IntuneDevicesToExcel {
            if ($script:intuneDevices.Count -eq 0) {
                [System.Windows.MessageBox]::Show(
                    "No devices to export. Please load devices first.",
                    "No Data",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
                return
            }
            
            try {
                $ExportButton.IsEnabled = $false
                # Create SaveFileDialog
                $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
                $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
                $saveDialog.Title = "Save Intune Mobile Devices Report"
                $saveDialog.FileName = "IntuneMobileDevices_$(Get-Date -Format 'yyyyMMdd-HHmmss').xlsx"
                
                # Set initial directory to Documents\Reports\Intune if it exists
                $defaultFolder = Join-Path $env:USERPROFILE "Documents\Reports\Intune"
                if (Test-Path $defaultFolder) {
                    $saveDialog.InitialDirectory = $defaultFolder
                }
                
                if (-not $saveDialog.ShowDialog()) {
                    Write-Log "Export cancelled by user"
                    $ExportButton.IsEnabled = $true
                    return
                }
                
                $outputFile = $saveDialog.FileName
                $StatusText.Text = "Exporting to Excel..."
                $StatusText.Foreground = "#FF9800"
                
                Write-Log "Exporting to: $outputFile"
                
                # Export to Excel with formatting
                $script:intuneDevices | Export-Excel -Path $outputFile `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle "Medium1" `
                    -WorksheetName "Mobile Devices"
                
                # Add conditional formatting
                $excel = Open-ExcelPackage -Path $outputFile
                $worksheet = $excel.Workbook.Worksheets["Mobile Devices"]
                
                # Find columns
                $headers = 1..$worksheet.Dimension.Columns | ForEach-Object {
                    $worksheet.Cells[1, $_].Value
                }
                
                # Format Compliance State column
                $complianceColIndex = ($headers.IndexOf("Compliance State")) + 1
                if ($complianceColIndex -gt 0) {
                    $colName = Get-ExcelColumnName $complianceColIndex
                    
                    Add-ConditionalFormatting -Worksheet $worksheet `
                        -Range "${colName}2:${colName}$($worksheet.Dimension.Rows)" `
                        -RuleType Equal -ConditionValue "compliant" `
                        -BackgroundColor ([System.Drawing.Color]::LightGreen) `
                        -ForegroundColor ([System.Drawing.Color]::DarkGreen)
                    
                    Add-ConditionalFormatting -Worksheet $worksheet `
                        -Range "${colName}2:${colName}$($worksheet.Dimension.Rows)" `
                        -RuleType Equal -ConditionValue "noncompliant" `
                        -BackgroundColor ([System.Drawing.Color]::LightPink) `
                        -ForegroundColor ([System.Drawing.Color]::DarkRed)
                }
                
                # Format Jail Broken column
                $jailBrokenColIndex = ($headers.IndexOf("Jail Broken")) + 1
                if ($jailBrokenColIndex -gt 0) {
                    $colName = Get-ExcelColumnName $jailBrokenColIndex
                    Add-ConditionalFormatting -Worksheet $worksheet `
                        -Range "${colName}2:${colName}$($worksheet.Dimension.Rows)" `
                        -RuleType Equal -ConditionValue "Detected" `
                        -BackgroundColor ([System.Drawing.Color]::Orange) `
                        -ForegroundColor ([System.Drawing.Color]::DarkRed)
                }
                
                # Format IMEI column as number
                $imeiColIndex = ($headers.IndexOf("IMEI")) + 1
                if ($imeiColIndex -gt 0) {
                    $colName = Get-ExcelColumnName $imeiColIndex
                    $worksheet.Cells["${colName}2:${colName}$($worksheet.Dimension.Rows)"].Style.Numberformat.Format = "0"
                }
                
                Close-ExcelPackage $excel
                
                $StatusText.Text = "Export complete"
                $StatusText.Foreground = "#28a745"
                Write-Log "Export completed: $outputFile"
                
                # Ask to open file
                $result = [System.Windows.MessageBox]::Show(
                    "Export completed successfully!`n`n$outputFile`n`nWould you like to open the file?",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Information
                )
                
                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                    Start-Process $outputFile
                }
                
            } catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Export error: $errorMsg"
                $StatusText.Text = "Export failed"
                $StatusText.Foreground = "#dc3545"
                [System.Windows.MessageBox]::Show(
                    "Failed to export to Excel:`n`n$errorMsg",
                    "Export Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
            } finally {
                $ExportButton.IsEnabled = $true
            }
        }
        
        # Helper function for Excel column names
        function Get-ExcelColumnName {
            param([int]$ColumnNumber)
            $columnName = ""
            while ($ColumnNumber -gt 0) {
                $modulo = ($ColumnNumber - 1) % 26
                $columnName = [char](65 + $modulo) + $columnName
                $ColumnNumber = [math]::Floor(($ColumnNumber - $modulo) / 26)
            }
            return $columnName
        }
        
        # Button event handlers
        $LoadDevicesButton.Add_Click({ Load-IntuneDevices })
        $RefreshButton.Add_Click({ Load-IntuneDevices })
        $ExportButton.Add_Click({ Export-IntuneDevicesToExcel })
        $CloseButton.Add_Click({ $window.Close() })
        
        # Initial state
        $ExportButton.IsEnabled = $false
        
        # Show window
        $window.ShowDialog() | Out-Null
        
    } catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Error opening Intune Mobile Devices window: $errorMsg"
        [System.Windows.MessageBox]::Show(
            "Failed to open Intune Mobile Devices window:`n`n$errorMsg",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

# Button Click Handler
$syncHash.IntuneMobileButton.Add_Click({
    Write-Log "Opening Intune Mobile Devices window..."
    Show-IntuneMobileDevicesWindow
})



$syncHash.PermissionAuditButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Permission Audit Report feature is planned for version 3.6.0",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})



# Intune & SCCM - Future Features
$syncHash.SCCMDevicesButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "SCCM Device Management feature is planned for a future version.`n`nThis will include:`n- View all SCCM managed devices`n- Device collections`n- Deployment status`n- Hardware inventory",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})

$syncHash.IntuneComplianceButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Compliance Policy Reports feature is planned for a future version.`n`nThis will include:`n- Compliance status reports`n- Policy assignment details`n- Non-compliant device lists`n- Trend analysis",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})

Write-Log "IT Operations Center initialized"
Write-Log "Ready for IT operations management"

$Window.ShowDialog() | Out-Null