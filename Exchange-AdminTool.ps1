<#
.SYNOPSIS
    Exchange Online Management Tool - GUI-based management for Exchange Online permissions and group memberships

.DESCRIPTION
    A comprehensive PowerShell GUI tool for managing Exchange Online mailbox permissions, calendar permissions,
    and Active Directory group memberships. Features include:
    - Mailbox permissions (Full Access & Send As)
    - Calendar permissions (7 levels)
    - AD group member viewing and management
    - Excel export capabilities for all data
    - Double-click any user/group to view their AD properties
    - GUID resolution for AD groups in permission lists

.AUTHOR
    Created by: Craig Werts
    Company: Geller & Co.
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
    
    Version 2.5 - Current
    - Improved error handling and logging
    - Added progress indicators for group member loading
    - Added "Copy Emails" button for group members (Outlook format)
    - Enhanced UI with status indicators
    - AutoMapping disabled by default for Full Access permissions

.REQUIREMENTS
    - PowerShell 5.1 or higher
    - ExchangeOnlineManagement module
    - ImportExcel module (for Excel export functionality)
    - Appropriate Exchange Online administrator permissions
    - Corporate proxy configuration (if applicable)

.NOTES
    File Name      : Exchange-AdminTool.ps1
    Prerequisite   : ExchangeOnlineManagement module must be installed
    
.USAGE
    Simply run the script. It will:
    1. Check for required modules and install ImportExcel if needed
    2. Connect to Exchange Online (browser authentication)
    3. Launch the GUI management tool
    
    From the GUI you can:
    - Manage Mailboxes: Add/edit/remove Full Access and Send As permissions
    - Calendar Permissions: Add/edit/remove calendar delegation permissions
    - AD Group Members: View group members, export to Excel, copy email addresses
    - Double-click any user/group name in permission lists to view their AD properties

.FEATURES
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

.PROXY CONFIGURATION
    The script includes proxy configuration for corporate environments.
    Update the proxy URL in lines 21-23 if different from default:
    http://proxy.gellerco.com:8080

#>

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

# Uncomment this section and modify line 112 with your proxy settings if needed
<#
[system.net.webrequest]::defaultwebproxy = new-object system.net.webproxy('YOURDOMAINPROXY')  
[system.net.webrequest]::defaultwebproxy.credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials  
[system.net.webrequest]::defaultwebproxy.BypassProxyOnLocal = $true  
#>

Write-Host "======================================" -ForegroundColor Cyan
Write-Host "Exchange Online Management Tool" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

Import-Module ExchangeOnlineManagement -ErrorAction Stop

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

$existingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue

if ($null -eq $existingConnection -or $existingConnection.State -ne 'Connected') {
    Write-Host "No active Exchange Online connection detected." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "IMPORTANT: You must authenticate to Exchange Online before launching the GUI." -ForegroundColor Yellow
    Write-Host "This is due to Windows authentication limitations in GUI applications." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Connecting to Exchange Online (browser window will open)..." -ForegroundColor Green
    Write-Host ""
    
    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        
        Write-Host ""
        Write-Host "Successfully connected to Exchange Online!" -ForegroundColor Green
        Write-Host "Launching GUI..." -ForegroundColor Green
        Write-Host ""
        Start-Sleep -Seconds 2
        
    } catch {
        Write-Host ""
        Write-Host "Connection failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Write-Host "Troubleshooting tips:" -ForegroundColor Yellow
        Write-Host "1. Try running 'Connect-ExchangeOnline' manually first" -ForegroundColor Yellow
        Write-Host "2. Clear cached credentials: Settings > Accounts > Access work or school" -ForegroundColor Yellow
        Write-Host "3. Contact IT if authentication continues to fail" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Press ENTER to exit..." -ForegroundColor Red
        $null = Read-Host
        exit
    }
} else {
    Write-Host "Existing Exchange Online connection detected!" -ForegroundColor Green
    Write-Host "Connected as: $($existingConnection.UserPrincipalName)" -ForegroundColor Cyan
    Write-Host "Connection State: $($existingConnection.State)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Launching GUI..." -ForegroundColor Green
    Start-Sleep -Seconds 1
}

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
        Title="Exchange Online Management Tool" 
        Height="600" 
        Width="800" 
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanMinimize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="100"/>
        </Grid.RowDefinitions>
        
        <Border Grid.Row="0" Background="#233A4A" Padding="15">
            <StackPanel>
                <TextBlock Text="Exchange Online Management" 
                          FontSize="20" 
                          FontWeight="Bold" 
                          Foreground="White"/>
                <TextBlock x:Name="StatusText" 
                          Text="Status: Connected" 
                          FontSize="12" 
                          Foreground="#00FF00" 
                          Margin="0,5,0,0"/>
            </StackPanel>
        </Border>
        
        <Border Grid.Row="1" Background="WhiteSmoke" Padding="20">
            <StackPanel>
                <GroupBox Header="Connection" Padding="10" Margin="0,0,0,20">
                    <StackPanel>
                        <TextBlock x:Name="ConnectionInfoText" 
                                  Text="Connected and ready to manage Exchange Online" 
                                  TextWrapping="Wrap" 
                                  Margin="0,0,0,15"/>
                        <Button x:Name="DisconnectButton" 
                               Content="Disconnect" 
                               Width="120" 
                               Height="35" 
                               Margin="5"
                               Background="#dc3545"
                               Foreground="White"
                               FontWeight="Bold"
                               HorizontalAlignment="Center"
                               Cursor="Hand"/>
                    </StackPanel>
                </GroupBox>
                
                <GroupBox Header="Management Options" 
                         x:Name="ManagementGroup" 
                         Padding="10">
                    <StackPanel>
                        <Button x:Name="MailboxButton" 
                               Content="Manage Mailboxes" 
                               Height="30" 
                               Margin="0,5"
                               HorizontalAlignment="Left"
                               Width="200"/>
                        <Button x:Name="CalendarButton" 
                               Content="Calendar Permissions" 
                               Height="30" 
                               Margin="0,5"
                               HorizontalAlignment="Left"
                               Width="200"/>
                        <Button x:Name="GroupMembersButton" 
                               Content="AD Group Members" 
                               Height="30" 
                               Margin="0,5"
                               HorizontalAlignment="Left"
                               Width="200"/>
                    </StackPanel>
                </GroupBox>
            </StackPanel>
        </Border>
        
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
$syncHash.DisconnectButton = $Window.FindName("DisconnectButton")
$syncHash.StatusText = $Window.FindName("StatusText")
$syncHash.ConnectionInfoText = $Window.FindName("ConnectionInfoText")
$syncHash.LogBox = $Window.FindName("LogBox")
$syncHash.ManagementGroup = $Window.FindName("ManagementGroup")
$syncHash.MailboxButton = $Window.FindName("MailboxButton")
$syncHash.CalendarButton = $Window.FindName("CalendarButton")
$syncHash.GroupMembersButton = $Window.FindName("GroupMembersButton")

function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logEntry = "[$timestamp] $Message`r`n"
    $syncHash.Window.Dispatcher.Invoke([action]{
        $syncHash.LogBox.AppendText($logEntry)
    })
}

$connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
if ($connInfo) {
    $syncHash.ConnectionInfoText.Text = "Connected as: $($connInfo.UserPrincipalName)"
    Write-Log "Connected to Exchange Online as $($connInfo.UserPrincipalName)"
}

$syncHash.DisconnectButton.Add_Click({
    Write-Log "Disconnecting from Exchange Online..."
    
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-Log "Successfully disconnected from Exchange Online"
        
        [System.Windows.MessageBox]::Show(
            "Disconnected from Exchange Online.`n`nThe application will now close.",
            "Disconnected",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
        
        $syncHash.Window.Close()
        
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
        
        $fullAccessPerms = Get-MailboxPermission -Identity $MailboxIdentity -ErrorAction Stop | 
            Where-Object {$_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-5-*" -and $_.IsInherited -eq $false -and $_.AccessRights -contains "FullAccess"}
        
        $sendAsPerms = Get-RecipientPermission -Identity $MailboxIdentity -ErrorAction Stop | 
            Where-Object {$_.Trustee -notlike "NT AUTHORITY\*" -and $_.Trustee -notlike "S-1-5-*" -and $_.AccessRights -contains "SendAs"}
        
        $allUsers = @{}
        
        foreach ($perm in $fullAccessPerms) {
            $userKey = $perm.User.ToString()
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
        }
        
        foreach ($perm in $sendAsPerms) {
            $userKey = $perm.Trustee.ToString()
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
        }
        
        return $allUsers.Values | Sort-Object User
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
                        <TextBlock Text="Shared Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
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
                        <TextBlock Text="Shared Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
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
                        <TextBlock Text="Shared Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
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
                    -TableStyle Medium8 `
                    -TableName "MailboxPermissions"
                
                Write-Log "Successfully exported $($exportData.Count) permissions to Excel"
                
                $result = [System.Windows.MessageBox]::Show(
                    "Mailbox permissions exported successfully!`n`nFile: $excelPath`n`nWould you like to open the file now?",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Information
                )
                
                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                    Start-Process $excelPath
                }
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
    
    $MbxWindow.ShowDialog() | Out-Null
})

$syncHash.CalendarButton.Add_Click({
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
                        <TextBlock Text="Shared Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
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
                        <TextBlock Text="Shared Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
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
                        <TextBlock Text="Shared Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
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
            $perms = Get-MailboxFolderPermission -Identity "${mailbox}:\Calendar" -ErrorAction Stop |
                     Where-Object {$_.User -notlike "Default" -and $_.User -notlike "Anonymous"} |
                     Select-Object @{N="User";E={$_.User.DisplayName}}, @{N="AccessRights";E={$_.AccessRights -join ", "}}
            $PermissionsGrid.ItemsSource = $perms
            
            if ($perms -and $perms.Count -gt 0) {
                $ExportToExcelButton.IsEnabled = $true
            } else {
                $ExportToExcelButton.IsEnabled = $false
            }
            
            Write-Log "Loaded $($perms.Count) permissions"
        } catch {
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
                    -TableStyle Medium8 `
                    -TableName "CalendarPermissions"
                
                Write-Log "Successfully exported $($exportData.Count) permissions to Excel"
                
                $result = [System.Windows.MessageBox]::Show(
                    "Calendar permissions exported successfully!`n`nFile: $excelPath`n`nWould you like to open the file now?",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Information
                )
                
                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                    Start-Process $excelPath
                }
                
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
            $perms = Get-MailboxFolderPermission -Identity "${mailbox}:\Calendar" -ErrorAction Stop |
                     Where-Object {$_.User -notlike "Default" -and $_.User -notlike "Anonymous"} |
                     Select-Object @{N="User";E={$_.User.DisplayName}}, @{N="AccessRights";E={$_.AccessRights -join ", "}}
            $RemovePermissionsGrid.ItemsSource = $perms
            Write-Log "Loaded $($perms.Count) permissions"
        } catch {
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
    
    $CalWindow.ShowDialog() | Out-Null
})

$syncHash.GroupMembersButton.Add_Click({
    Write-Log "Opening AD Group Members window..."
    
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
                <TextBlock Text="Group Email or Name:" FontWeight="Bold" Margin="0,0,0,5"/>
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
                        <DataGridTextColumn Header="Type" Binding="{Binding RecipientType}" Width="*"/>
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
                              Text="Enter a group email or name to begin" 
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
            [System.Windows.MessageBox]::Show("Please enter a group email or name", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        try {
            $LoadGroupButton.IsEnabled = $false
            $StatusText.Text = "Loading group information..."
            Write-Log "Loading group: $groupIdentity"
            
            $group = $null
            try {
                $group = Get-DistributionGroup -Identity $groupIdentity -ErrorAction Stop
                $groupType = "Distribution Group"
            } catch {
                try {
                    $group = Get-Group -Identity $groupIdentity -ErrorAction Stop
                    $groupType = "Security Group"
                } catch {
                    throw "Group not found. Please verify the group name or email address."
                }
            }
            
            if ($null -eq $group) {
                throw "Group not found"
            }
            
            $StatusText.Text = "Loading members..."
            Write-Log "Retrieving members for: $($group.DisplayName)"
            
            $members = Get-DistributionGroupMember -Identity $group.Identity -ErrorAction Stop
            
            $enrichedMembers = @()
            $processedCount = 0
            $totalCount = $members.Count
            
            foreach ($member in $members) {
                $processedCount++
                $StatusText.Text = "Processing member $processedCount of $totalCount..."
                
                try {
                    $recipient = Get-Recipient -Identity $member.Identity -ErrorAction SilentlyContinue
                    
                    $memberObj = [PSCustomObject]@{
                        DisplayName = $member.DisplayName
                        Email = if ($member.PrimarySmtpAddress) { $member.PrimarySmtpAddress.ToString() } else { "N/A" }
                        RecipientType = $member.RecipientType
                        Title = if ($recipient.Title) { $recipient.Title } else { "" }
                        Department = if ($recipient.Department) { $recipient.Department } else { "" }
                        Identity = $member.Identity
                    }
                    
                    $enrichedMembers += $memberObj
                } catch {
                    Write-Log "Warning: Could not get full details for $($member.DisplayName)"
                    $enrichedMembers += [PSCustomObject]@{
                        DisplayName = $member.DisplayName
                        Email = if ($member.PrimarySmtpAddress) { $member.PrimarySmtpAddress.ToString() } else { "N/A" }
                        RecipientType = $member.RecipientType
                        Title = ""
                        Department = ""
                        Identity = $member.Identity
                    }
                }
            }
            
            $MembersGrid.ItemsSource = $enrichedMembers
            
            $GroupNameText.Text = "Group: $($group.DisplayName)"
            $GroupTypeText.Text = "Type: $groupType"
            $GroupEmailText.Text = "Email: $(if ($group.PrimarySmtpAddress) { $group.PrimarySmtpAddress } else { 'N/A' })"
            $GroupMemberCountText.Text = "Total Members: $($enrichedMembers.Count)"
            $GroupInfoPanel.Visibility = [System.Windows.Visibility]::Visible
            
            $script:currentGroupInfo = $group
            $script:currentMembers = $enrichedMembers
            
            $ExportToExcelButton.IsEnabled = $true
            $CopyEmailsButton.IsEnabled = $true
            
            $StatusText.Text = "Loaded $($enrichedMembers.Count) members successfully"
            Write-Log "Successfully loaded $($enrichedMembers.Count) members from $($group.DisplayName)"
            
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
                Where-Object { $_.Email -ne "N/A" } | 
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
            $groupNameSafe = $script:currentGroupInfo.DisplayName -replace '[\\/:*?"<>|]', '_'
            $saveDialog.FileName = "Group_Members_${groupNameSafe}_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $ExportToExcelButton.IsEnabled = $false
                $StatusText.Text = "Exporting to Excel..."
                Write-Log "Exporting group members to Excel: $excelPath"
                
                $exportData = @()
                foreach ($member in $script:currentMembers) {
                    $exportData += [PSCustomObject]@{
                        'Group Name' = $script:currentGroupInfo.DisplayName
                        'Group Email' = if ($script:currentGroupInfo.PrimarySmtpAddress) { $script:currentGroupInfo.PrimarySmtpAddress.ToString() } else { "N/A" }
                        'Member Display Name' = $member.DisplayName
                        'Member Email' = $member.Email
                        'Recipient Type' = $member.RecipientType
                        'Title' = $member.Title
                        'Department' = $member.Department
                        'Export Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                $exportData | Export-Excel -Path $excelPath `
                    -WorksheetName "Group Members" `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle Medium9 `
                    -TableName "GroupMembers"
                
                Write-Log "Successfully exported $($exportData.Count) members to Excel"
                $StatusText.Text = "Export completed successfully"
                
                $result = [System.Windows.MessageBox]::Show(
                    "Group members exported successfully!`n`nFile: $excelPath`n`nWould you like to open the file now?",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Information
                )
                
                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                    Start-Process $excelPath
                }
                
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

Write-Log "Exchange Online Management Tool initialized"
Write-Log "Ready to manage Exchange Online"


$Window.ShowDialog() | Out-Null


