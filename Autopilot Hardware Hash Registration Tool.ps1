<#
.SYNOPSIS
Autopilot Hardware Hash Registration Tool - Modern WPF GUI (Offline + Styled)
Version: 1.2 (RichTextBox Logs)
Created by: Nitesh Kumar Solanki
Date: 03-Nov-2025
#>

# --- Hardcoded App-Based Authentication ---
$TenantID = "TID"
$ClientID = "CID"
$ClientSecret = "CS"


Add-Type -AssemblyName PresentationFramework, WindowsBase

# --- Ensure Admin Privileges ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    [System.Windows.MessageBox]::Show("Please run this tool as Administrator.", "Admin Privileges Required", 'OK', 'Warning') 
    exit
}

# --- Set Execution Policy for Process ---
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# --- Logging Function (RichTextBox + File) ---
$TempFolder = Join-Path $env:TEMP "AutopilotOffline"
if (-not (Test-Path $TempFolder)) { New-Item -Path $TempFolder -ItemType Directory -Force | Out-Null }
$LogFile = Join-Path $TempFolder "AutopilotTool.log"

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [bool]$IsManual = $false
    )
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$ts][$Level] $Message"

    # Write to file
    try { Add-Content -Path $LogFile -Value $line -ErrorAction SilentlyContinue } catch {}

    # Write to RichTextBox
    if ($global:txtOutput -and $global:txtOutput -is [System.Windows.Controls.RichTextBox]) {
        $para = New-Object System.Windows.Documents.Paragraph
        if ($IsManual) {
            $para.Margin = '0,10,0,10' 
            $run = New-Object System.Windows.Documents.Run $Message
            $run.FontWeight = 'Bold'
            $run.Foreground = [System.Windows.Media.Brushes]::OrangeRed
        } else {
            $run = New-Object System.Windows.Documents.Run $line
            switch ($Level.ToUpper()) {
                "SUCCESS" { $run.Foreground = [System.Windows.Media.Brushes]::LightGreen }
                "ERROR"   { $run.Foreground = [System.Windows.Media.Brushes]::OrangeRed }
                default   { $run.Foreground = [System.Windows.Media.Brushes]::LightBlue }
            }
        }
        $para.Inlines.Add($run)
        $global:txtOutput.Document.Blocks.Add($para)
        $global:txtOutput.ScrollToEnd()
    }
}

# ======================
# XAML UI SECTION
# ======================
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Autopilot Hardware Hash Registration Tool"
        Height="700" Width="650"
        Background="#F5F5F5"
        WindowStartupLocation="CenterScreen">
    <Grid Margin="15">
        <StackPanel>

            <!-- Header -->
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,0,0,15">
                <TextBlock Text="🖥️" FontSize="40" Margin="0,0,10,0"/>
                <TextBlock Text="Autopilot Hardware Hash Tool"
                           FontSize="24"
                           FontWeight="Bold"
                           Foreground="#2F4F4F"
                           VerticalAlignment="Center"/>
            </StackPanel>

            <!-- Registration Section -->
            <GroupBox Header="⚙️  Hardware Hash Registration"
                      FontWeight="Bold"
                      Padding="10"
                      Margin="0,0,0,10"
                      BorderBrush="#C0C0C0"
                      Background="#D3D3D3">
                <StackPanel>
                    <Button x:Name="btnRegister"
        Width="250" Height="38"
        FontWeight="SemiBold"
        FontSize="14"
        Background="#0078D7"     
        Foreground="White"       
        BorderBrush="#005A9E"
        Margin="0,5,0,0"
                            Content="⬆️  Register Hardware Hash"/>
                </StackPanel>
            </GroupBox>

<!-- Device Lookup Section -->
<GroupBox Header="🔍  Autopilot Device Lookup"
          FontWeight="Bold"
          Padding="10"
          Margin="0,0,0,10"
          BorderBrush="#C0C0C0"
          Background="#D3D3D3">

    <StackPanel>

        <TextBlock Text="Enter Serial Number:"
                   FontWeight="SemiBold"
                   Margin="0,0,0,5"/>

        <TextBox x:Name="LookupSerialBox"
                 Width="300"
                 Height="30"
                 FontSize="14"
                 VerticalContentAlignment="Center"
                 Margin="0,0,0,10"/>

        <Button x:Name="LookupBtn"
                Width="180"
                Height="32"
                FontWeight="SemiBold"
                FontSize="14"
                Background="#0078D7"     
        Foreground="White"       
        BorderBrush="#005A9E"
                Content="🔍  Lookup Device"/>
    </StackPanel>

</GroupBox>




            <!-- Log Output -->
            <RichTextBox x:Name="txtOutput"
                         Height="250"
                         IsReadOnly="True"
                         VerticalScrollBarVisibility="Auto"
                         Background="#1E1E1E"
                         FontFamily="Consolas"
                         FontSize="13"
                         Margin="0,5,0,5"/>

            <!-- Button Bar -->
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,5,0,0">
                <Button x:Name="btnClear" Content="🧹 Clear Logs" Width="120" Margin="5"/>
                <Button x:Name="btnInfo" Content="ℹ️ Info" Width="120" Margin="5"/>
                <Button x:Name="btnExit" Content="🚪 Exit" Width="120" Margin="5"/>
            </StackPanel>

        </StackPanel>
    </Grid>
</Window>
"@

# --- Parse XAML ---
$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$window = [Windows.Markup.XamlReader]::Load($reader)

# --- Reference Controls ---
$btnRegister = $window.FindName("btnRegister")
$btnClear    = $window.FindName("btnClear")
$btnInfo     = $window.FindName("btnInfo")
$btnExit     = $window.FindName("btnExit")
$global:txtOutput = $window.FindName("txtOutput")
$LookupBtn      = $window.FindName("LookupBtn")
$LookupSerialBox = $window.FindName("LookupSerialBox")

function Get-GraphToken {
    $tokenUrl = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientID
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }
    try {
        $resp = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body -ErrorAction Stop
        return $resp.access_token
    } catch {
        Write-Log "Get-GraphToken failed: $_" "ERROR"
        return $null
    }
}

function Graph {
    param([string]$Method,[string]$Url,$Body=$null)
    if (-not $global:GraphToken) {
        Write-Log "Obtaining Graph token..." "INFO"
        $global:GraphToken = Get-GraphToken
    }
    if (-not $global:GraphToken) { Write-Log "No Graph token available" "ERROR"; return $null }
    $headers = @{ Authorization = "Bearer $global:GraphToken"; "Content-Type" = "application/json" }
    try {
        if ($Body) {
            $json = $Body | ConvertTo-Json -Depth 10
            return Invoke-RestMethod -Method $Method -Uri $Url -Headers $headers -Body $json -ErrorAction Stop
        } else {
            return Invoke-RestMethod -Method $Method -Uri $Url -Headers $headers -ErrorAction Stop
        }
    } catch {
        # If token expired/invalid, clear token and try once
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 401) {
            Write-Log "Graph returned 401; refreshing token and retrying..." "INFO"
            $global:GraphToken = $null
            return Graph -Method $Method -Url $Url -Body $Body
        }
        Write-Log "Graph call error: $_" "ERROR"
        return $null
    }
}


# ======================
# EVENT HANDLERS
# ======================

# Register Hardware Hash
$btnRegister.Add_Click({
    Write-Log "Collecting hardware hash..." "INFO"
    try {
        $serial = (Get-CimInstance -Class Win32_BIOS).SerialNumber
        $devDetail = Get-CimInstance -Namespace root/cimv2/mdm/dmmap -Class MDM_DevDetail_Ext01 -Filter "InstanceID='Ext' AND ParentID='./DevDetail'"
        $hash = $devDetail.DeviceHardwareData
        if (-not $hash) { throw "Hardware hash not available" }

        $outputFile = "C:\${serial}_AutopilotHash.csv"
        [PSCustomObject]@{
            "Device Serial Number" = $serial
            "Hardware Hash" = $hash
        } | Export-Csv -Path $outputFile -NoTypeInformation -Force
        Write-Log "Hardware hash collected and saved to $outputFile" "SUCCESS"


        # ======= NO MODULE REQUIRED — USE GRAPH API DIRECTLY =======

        # 1. Get OAuth token
        $Body = @{
            grant_type    = "client_credentials"
            scope         = "https://graph.microsoft.com/.default"
            client_id     = $ClientID
            client_secret = $ClientSecret
        }

        $auth = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token" -Body $Body
        $Headers = @{ Authorization = "Bearer $($auth.access_token)" }

        # 2. Build Autopilot import JSON
        $payload = @{
            "@odata.type" = "#microsoft.graph.importedWindowsAutopilotDeviceIdentity"
            serialNumber  = $serial
            hardwareIdentifier = $hash
        }

        $json = $payload | ConvertTo-Json -Depth 5

        # 3. Upload hash
        Invoke-RestMethod `
            -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities" `
            -Headers $Headers `
            -Body $json `
            -ContentType "application/json"

        Write-Log "Device uploaded to Intune Autopilot Serial: $Serial successfully. Please notify niteshsolanki54@gmail.com to update the Group Tag and Hostname." "SUCCESS"
    }
    catch {
        Write-Log "Failed to upload device: $_" "ERROR"
        Write-Log "The hash file has been saved to C:\Temp\${serial}_Autopilot.csv. Please notify niteshsolanki54@gmail.com to register the device manually." "ERROR" $true
    }
})


# -------------------------
# Device Lookup (Check Device Details)
# -------------------------
$LookupBtn.Add_Click({

	Write-Log "Get Details: looking up Autopilot record..." "INFO"
    
    try{
	
    $serialCheck = $LookupSerialBox.Text.Trim()
    if (-not $serialCheck) { Write-Log "Please enter a serial number." "ERROR"; return }
    Write-Log "Checking Autopilot device details for serial: $serialCheck" "INFO"

    # -------------------------
    # AUTH
    # -------------------------
    $global:GraphToken = Get-GraphToken
        if (-not $global:GraphToken) { Write-Log "Auth failed - cannot get details" "ERROR"; return }

        $url = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,'$serialCheck')"
        $data = Graph -Method "GET" -Url $url
        if (-not $data -or -not $data.value -or $data.value.Count -eq 0) {
            Write-Log "No Autopilot record found for $serial" "ERROR"
            return
        }

        $dev = $data.value[0]

        # Basic fields
        $assignedName = $dev.displayName
        if (-not $assignedName) { $assignedName = $dev.assignedComputerName }
        if (-not $assignedName) { $assignedName = "(not set)" }

        # Requested field
        $dpStatus = $dev.deploymentProfileAssignmentStatus
        if (-not $dpStatus) { $dpStatus = "notAssigned" }

        Write-Log "Device Found:" "SUCCESS"
        Write-Log "Serial: $($dev.serialNumber)" "INFO"
        Write-Log "Group Tag: $($dev.groupTag)" "INFO"
        Write-Log "Assigned Name: $assignedName" "INFO"

        # -------------------------
        # STATUS LOGIC
        # -------------------------
        switch ($dpStatus) {

            "assignedUnknownSyncState" {
                Write-Log "Autopilot Profile is Assigned, Autopilot provisioning can be started." "SUCCESS"
                [System.Windows.MessageBox]::Show("Autopilot Profile is Assigned, Autopilot provisioning can be started.","Provisioning Ready","OK","Information") | Out-Null
            }

            "assignedUnkownSyncState" {   # tenant typo
                Write-Log "Autopilot Profile is Assigned, Autopilot provisioning can be started." "SUCCESS"
                [System.Windows.MessageBox]::Show("Autopilot Profile is Assigned, Autopilot provisioning can be started.","Provisioning Ready","OK","Information") | Out-Null
            }

            "assigned" {
                Write-Log "Autopilot Profile is Assigned, Autopilot provisioning can be started." "SUCCESS"
                [System.Windows.MessageBox]::Show("Autopilot Profile is Assigned, Autopilot provisioning can be started.","Provisioning Ready","OK","Information") | Out-Null
            }

            "notAssigned" {
                Write-Log "Autopilot Profile is not assigned. Contact Nitesh Kumar Solanki." "ERROR"
                [System.Windows.MessageBox]::Show("Autopilot Profile is not assigned. Contact Nitesh Kumar Solanki.","Autopilot Profile Missing","OK","Error") | Out-Null
            }
        } }
    
    catch {
        Write-Log "Failed fetching device details (Beta API): $_" "ERROR"
    }

}) | Out-Null


# Clear Logs
$btnClear.Add_Click({
    $txtOutput.Document.Blocks.Clear()
    Write-Log "Logs cleared." "INFO"
})

# Info
$btnInfo.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Autopilot Hardware Hash Registration Tool`nVersion: 1.2`nCreated by: Created By: Nitesh Kumar Solanki (niteshsolanki54@gmail.com)`n`nLog File:`n$LogFile",
        "About Tool",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})

# Exit
$btnExit.Add_Click({ $window.Close() })

# ======================
# SHOW WINDOW
# ======================
Write-Log "Tool launched successfully. Logging to: $LogFile" "INFO"
$window.ShowDialog() | Out-Null

