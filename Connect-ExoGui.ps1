#Requires -Modules ExchangeOnlineManagement

# Ensure script runs in STA mode for WPF and authentication windows
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Host "Restarting in STA mode..." -ForegroundColor Yellow
    Start-Process powershell.exe -ArgumentList "-STA", "-NoProfile", "-File", "`"$PSCommandPath`"" -Wait
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12  

[system.net.webrequest]::defaultwebproxy = new-object system.net.webproxy('http://proxy.gellerco.com:8080')  
[system.net.webrequest]::defaultwebproxy.credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials  
[system.net.webrequest]::defaultwebproxy.BypassProxyOnLocal = $true  

# Pre-connection check
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "Exchange Online Management Tool" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

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

# Check if already connected
$existingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue

if ($null -eq $existingConnection -or $existingConnection.State -ne 'Connected') {
    Write-Host "No active Exchange Online connection detected." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "IMPORTANT: You must authenticate to Exchange Online before launching the GUI." -ForegroundColor Yellow
    Write-Host "This is due to Windows authentication limitations in GUI applications." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press ENTER to authenticate now (browser window will open)..." -ForegroundColor Green
    $null = Read-Host
    
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Write-Host ""
    
    try {
        # Attempt connection with DisableWAM
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        
        Write-Host ""
        Write-Host "Successfully connected to Exchange Online!" -ForegroundColor Green
        Write-Host ""
        Write-Host "Press ENTER to launch the GUI..." -ForegroundColor Green
        $null = Read-Host
        
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

# Create synchronized hashtable for thread-safe GUI updates
$syncHash = [hashtable]::Synchronized(@{})

# XAML for the main window
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
        
        <!-- Header Section -->
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
        
        <!-- Main Content Area -->
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
                    </StackPanel>
                </GroupBox>
            </StackPanel>
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

# Load XAML
$reader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Get UI elements
$syncHash.Window = $Window
$syncHash.DisconnectButton = $Window.FindName("DisconnectButton")
$syncHash.StatusText = $Window.FindName("StatusText")
$syncHash.ConnectionInfoText = $Window.FindName("ConnectionInfoText")
$syncHash.LogBox = $Window.FindName("LogBox")
$syncHash.ManagementGroup = $Window.FindName("ManagementGroup")
$syncHash.MailboxButton = $Window.FindName("MailboxButton")
$syncHash.CalendarButton = $Window.FindName("CalendarButton")

# Helper function to update log
function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logEntry = "[$timestamp] $Message`r`n"
    $syncHash.Window.Dispatcher.Invoke([action]{
        $syncHash.LogBox.AppendText($logEntry)
    })
}

# Update connection info on load
$connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
if ($connInfo) {
    $syncHash.ConnectionInfoText.Text = "Connected as: $($connInfo.UserPrincipalName)"
    Write-Log "Connected to Exchange Online as $($connInfo.UserPrincipalName)"
}

# Disconnect Button Click Event
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

# Placeholder events for management buttons
$syncHash.MailboxButton.Add_Click({
    Write-Log "Mailbox management feature - Coming soon!"
    [System.Windows.MessageBox]::Show(
        "Mailbox management feature will be implemented here.",
        "Info",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})

$syncHash.CalendarButton.Add_Click({
    Write-Log "Opening Calendar Permissions window..."
    
    # XAML for Calendar Permissions window
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
    
    # Get controls
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
    
    # ADD TAB
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
    
    # VIEW/EDIT TAB
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
            
            # Enable Export button if permissions were loaded
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
    
    # EXPORT TO EXCEL BUTTON
    $ExportToExcelButton.Add_Click({
        $mailbox = $ViewMailboxBox.Text.Trim()
        $permissions = $PermissionsGrid.ItemsSource
        
        if ($null -eq $permissions -or $permissions.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No permissions to export. Please load permissions first.", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        # Check if ImportExcel module is available
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            [System.Windows.MessageBox]::Show("ImportExcel module is not installed. Cannot export to Excel.`n`nPlease run: Install-Module -Name ImportExcel", "Module Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Write-Log "Export failed: ImportExcel module not installed"
            return
        }
        
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            # Create file dialog for save location
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.Title = "Save Calendar Permissions Report"
            $saveDialog.FileName = "Calendar_Permissions_$($mailbox.Replace('@','_'))_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $ExportToExcelButton.IsEnabled = $false
                Write-Log "Exporting permissions to Excel: $excelPath"
                
                # Prepare data for export
                $exportData = @()
                foreach ($perm in $permissions) {
                    $exportData += [PSCustomObject]@{
                        'Mailbox' = $mailbox
                        'Delegate User' = $perm.User
                        'Access Rights' = $perm.AccessRights
                        'Export Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                # Export to Excel with formatting
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
    
    # REMOVE TAB
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

# Initial log message
Write-Log "Exchange Online Management Tool initialized"
Write-Log "Click 'Connect to Exchange Online' to begin"

# Show the window
$Window.ShowDialog() | Out-Null