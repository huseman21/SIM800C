Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# SCRIPT-WIDE VARIABLES
# ====================================================================================
$script:baudRate = 115200
$script:serialPort = $null
$script:timeout = 10000 

# --- DARK MODE COLOR (Script-Scoped for access in event handlers) ---
$script:color_Background = [System.Drawing.ColorTranslator]::FromHtml("#1E1E1E") # Main Form Background
$script:color_Panel = [System.Drawing.ColorTranslator]::FromHtml("#2D2D30")      # Group Box Background (Original Status BG)
$script:color_Text = [System.Drawing.ColorTranslator]::FromHtml("#DCDCDC")       # Light Text (Original Status FG)
$script:color_Accent = [System.Drawing.ColorTranslator]::FromHtml("#007ACC")     # Azure Blue for emphasis
$script:color_ButtonBG = [System.Drawing.ColorTranslator]::FromHtml("#333333")    # Button background
$script:color_ButtonHover = [System.Drawing.ColorTranslator]::FromHtml("#444444") 
$script:color_FlashAccent = [System.Drawing.ColorTranslator]::FromHtml("#00FF7F") # Spring Green for Flash

# --- Signal Polling Timer ---
$signalTimer = New-Object System.Windows.Forms.Timer
$signalTimer.Interval = 10000 # Poll every 10 seconds

# --- Anchor ---
$anchorLeftRight = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$anchorTopLeft = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$anchorTopRight = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$anchorLeftRightTop = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top
$anchorAll = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom


# LOGGING AND VISUAL FEEDBACK 
# ====================================================================================
function Log-Output {
    param([string]$Message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $outputTextBox.AppendText("[$timestamp] $Message`r`n")
    $outputTextBox.SelectionStart = $outputTextBox.Text.Length;
    $outputTextBox.ScrollToCaret()
}

# NEW FUNCTION: Adds two blank lines for visual separation
function Log-Separator {
    if ($outputTextBox.Text.Length -gt 0) {
        $outputTextBox.AppendText("`r`n`r`n")
    }
    $outputTextBox.SelectionStart = $outputTextBox.Text.Length;
    $outputTextBox.ScrollToCaret()
}

function Flash-Status {
    # Set flash color using script scope variable
    $script:statusTextBox.BackColor = $script:color_FlashAccent 
    $script:statusTextBox.ForeColor = [System.Drawing.Color]::Black 
    
    # Use a non-blocking timer to revert the color
    $flashTimer = New-Object System.Windows.Forms.Timer
    $flashTimer.Interval = 200 # Flash duration: 200ms
    
    # Event handler script block must use $script: scope for external variables and controls.
    $handler_flash = {
        # Revert colors using script-level variables explicitly
        $script:statusTextBox.BackColor = $script:color_Panel
        $script:statusTextBox.ForeColor = $script:color_Text
        
        # Stop and dispose of the timer after it fires once
        $this.Enabled = $false
        $this.Dispose()
    }
    
    $flashTimer.Add_Tick($handler_flash)
    $flashTimer.Enabled = $true
}


# SERIAL
# ====================================================================================
function Open-Port {
    param ([string]$PortName)
    if ($script:serialPort) {
        if ($script:serialPort.IsOpen) { $script:serialPort.Close() }
        $script:serialPort.Dispose(); $script:serialPort = $null
    }
    Start-Sleep -Milliseconds 250
    
    try {
        $script:serialPort = New-Object System.IO.Ports.SerialPort $PortName, $script:baudRate, 'None', 8, 'One'
        $script:serialPort.ReadTimeout = 500; $script:serialPort.WriteTimeout = 500
        $script:serialPort.DtrEnable = $true; $script:serialPort.RtsEnable = $true
        $script:serialPort.Open()
        
        $statusTextBox.Text = "Port $PortName opened successfully."
        Log-Output "--- Port $PortName opened at $script:baudRate baud. ---"
        
        # *** STATUS UPDATE: ON ***
        $connectionIndicator.BackColor = [System.Drawing.Color]::Green
        $connectionIndicator.Text = "CONNECTED"
        # *************************
        
        # Initial Handshake
        Log-Output "Attempting initial handshake (Autobaud check)..."
        $script:serialPort.Write("AT`r`n"); Start-Sleep -Milliseconds 50
        $script:serialPort.Write("AT`r`n"); Start-Sleep -Milliseconds 50
        $script:serialPort.Write("AT`r`n"); Start-Sleep -Milliseconds 100 
        $script:serialPort.ReadExisting() | Out-Null # Flush buffer

        # Enable controls and start timer
        $connectButton.Enabled = $false; $disconnectButton.Enabled = $true
        $smsGroupBox.Enabled = $true; $httpGroupBox.Enabled = $true
        $extGroupBox.Enabled = $true; $btGroupBox.Enabled = $true 
        $customGroupBox.Enabled = $true
        $outputTextBox.Focus()
        
        $signalTimer.Enabled = $true # Start signal polling
        Update-SignalStatus # Run first check immediately
        
        return $true
    }
    catch {
        $statusTextBox.Text = "Error opening port: $($_.Exception.Message)"
        Log-Output "Error: Failed to open $PortName. $($_.Exception.Message)"
        
        # *** STATUS UPDATE: OFF ***
        $connectionIndicator.BackColor = [System.Drawing.Color]::Red
        $connectionIndicator.Text = "OFFLINE"
        return $false
    }
}

function Close-Port {
    if ($script:serialPort -and $script:serialPort.IsOpen) {
        $portName = $script:serialPort.PortName
        try {
            $script:serialPort.Close(); $script:serialPort.Dispose(); $script:serialPort = $null
            $statusTextBox.Text = "Port $portName closed."
            Log-Output "--- Port $portName closed. ---"
        }
        catch {
            $statusTextBox.Text = "Error closing port: $($_.Exception.Message)"
            Log-Output "Error closing port: $($_.Exception.Message)"
        }
    } else {
        $statusTextBox.Text = "Port is already closed or not initialized."
    }
    
    # *** STATUS UPDATE: OFF and stop timer/reset signal ***
    $connectionIndicator.BackColor = [System.Drawing.Color]::Red
    $connectionIndicator.Text = "OFFLINE"
    $signalTimer.Enabled = $false
    $signalValueLabel.Text = "Disconnected"
    $signalBarLabels | ForEach-Object { 
        $_.BackColor = $script:color_Panel
        $_.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    }
    # ******************************************************
    
    # Disable controls
    $connectButton.Enabled = $true; $disconnectButton.Enabled = $false
    $smsGroupBox.Enabled = $false; $httpGroupBox.Enabled = $false
    $extGroupBox.Enabled = $false; $btGroupBox.Enabled = $false
    $customGroupBox.Enabled = $false
}

function Send-AtCommand {
    param([string]$Command, [int]$PostOkReadMs = 1000, [switch]$IsSilent) 

    if (-not $script:serialPort -or -not $script:serialPort.IsOpen) {
        $statusTextBox.Text = "Error: Port not open."
        if (-not $IsSilent) { Log-Output "Error: Port not open. Cannot send command $Command." }
        return ""
    }
    
    # Flash Status Indicator for interactive commands
    if (-not $IsSilent) {
        Flash-Status 
    }
    
    $commandToSend = "$Command`r`n"; $response = ""; $buffer = ""
    
    try {
        $script:serialPort.ReadExisting() | Out-Null
        $script:serialPort.Write($commandToSend)
        $startTime = Get-Date; $foundFinalStatus = $false
        
        # Phase 1: Wait for Final Status (OK/ERROR/TIMEOUT)
        while (((Get-Date) - $startTime).TotalMilliseconds -lt $script:timeout) {
            $data = $script:serialPort.ReadExisting()
            if ($data -ne "") {
                $buffer += $data
                if ($buffer -match "`r`nOK`r`n" -or $buffer -match "`r`nERROR`r`n" -or $buffer -match "\+CMS ERROR: " -or $buffer -match "\+CME ERROR: ") {
                    $foundFinalStatus = $true; break
                }
            }
            Start-Sleep -Milliseconds 50
        }
        
        # Phase 2: Extended Read for URCs after OK
        if ($foundFinalStatus -and $buffer -match "OK" -and $PostOkReadMs -gt 0) {
            $postOkStartTime = Get-Date
            if (-not $IsSilent) { Log-Output ">>> Command accepted. Polling for async data for $($PostOkReadMs/1000)s..." } 
            while (((Get-Date) - $postOkStartTime).TotalMilliseconds -lt $PostOkReadMs) {
                $data = $script:serialPort.ReadExisting()
                if ($data -ne "") { $buffer += $data }
                Start-Sleep -Milliseconds 100 
            }
            if (-not $IsSilent) { Log-Output ">>> Asynchronous poll complete." } 
        }
        
        $response = $buffer
        # Update status box
        if ($response -match "ERROR" -or $response -match "CMS ERROR" -or $response -match "CME ERROR") {
            $statusTextBox.Text = "Command failed: $Command"
        } elseif ($response -match "OK") {
             $statusTextBox.Text = "Command successful: $Command"
        } else {
             $statusTextBox.Text = "Command completed without final status: $Command"
        }
        
        return $response.Trim()
    }
    catch [System.TimeoutException] {
        $statusTextBox.Text = "Timeout waiting for response for $Command."
        if (-not $IsSilent) { Log-Output "TIMEOUT: No response received after $script:timeout ms." } 
        return "TIMEOUT"
    }
    catch {
        $statusTextBox.Text = "Communication Error: $($_.Exception.Message)"
        if (-not $IsSilent) { Log-Output "Communication Error: $($_.Exception.Message)" } 
        return "ERROR"
    }
}

function Wait-ForHttpAction {
    param([int]$TimeoutMs = 15000)
    $startTime = Get-Date; $urcResponse = ""; $foundURC = $false
    
    while (((Get-Date) - $startTime).TotalMilliseconds -lt $TimeoutMs) {
        $data = $script:serialPort.ReadExisting()
        if ($data -ne "") {
            $urcResponse += $data
            $data.Split("`r`n") | Where-Object { $_ -ne "" } | ForEach-Object { Log-Output $_ }
            
            if ($urcResponse -match "\+HTTPACTION:") {
                $foundURC = $true
                $match = $urcResponse | Select-String -Pattern "\+HTTPACTION: \d+,(\d+),\d+"
                if ($match) {
                    $statusCode = $match.Matches.Groups[1].Value
                    if ($statusCode -eq "200") {
                        $statusTextBox.Text = "HTTP GET request completed successfully (Status 200). Ready to read."
                        return $true
                    } else {
                        $statusTextBox.Text = "HTTP GET request completed with error code $statusCode."
                        Log-Output "!!! ERROR: HTTP GET request failed with status code $statusCode."
                        return $false
                    }
                }
                break
            }
        }
        Start-Sleep -Milliseconds 100
    }
    if (-not $foundURC) {
        $statusTextBox.Text = "Timeout waiting for +HTTPACTION URC."
        Log-Output "TIMEOUT: HTTP GET failed to complete after $TimeoutMs ms."
        return $false
    }
}

# *** VISUAL LOGIC ***
# ------------------------------------------------------------------------------------
function Update-SignalStatus {
    # Check for open port explicitly, though the timer should be disabled on disconnect.
    if (-not $script:serialPort -or -not $script:serialPort.IsOpen) {
        return
    }

    # Flash the status box to provide visual feedback for the silent, periodic poll.
    Flash-Status 
    
    # Send AT+CSQ silently to prevent logging.
    $response = Send-AtCommand -Command "AT+CSQ" -PostOkReadMs 500 -IsSilent
    
    $rssiValue = -1
    
    # Removed the "^" anchor from the pattern to allow a match even if the response includes command echo or blank lines first.
    $rssiMatch = $response | Select-String -Pattern "\+CSQ:\s*(\d+),\s*\d+"
    
    if ($rssiMatch) {
        [int]::TryParse($rssiMatch.Matches.Groups[1].Value, [ref]$rssiValue) | Out-Null
    }
    
    $bars = 0; $color = [System.Drawing.Color]::Gray; $statusText = "No Service"
    
    # Map RSSI (0-31) to 5 visual bars
    if ($rssiValue -ge 1 -and $rssiValue -le 9) {
        $bars = 1; $color = [System.Drawing.Color]::DarkRed; $statusText = "Poor ($rssiValue RSSI)"
    } elseif ($rssiValue -ge 10 -and $rssiValue -le 14) {
        $bars = 2; $color = [System.Drawing.Color]::Orange; $statusText = "Fair ($rssiValue RSSI)"
    } elseif ($rssiValue -ge 15 -and $rssiValue -le 19) {
        $bars = 3; $color = [System.Drawing.Color]::Yellow; $statusText = "Good ($rssiValue RSSI)"
    } elseif ($rssiValue -ge 20 -and $rssiValue -le 25) {
        $bars = 4; $color = [System.Drawing.Color]::LightGreen; $statusText = "Excellent ($rssiValue RSSI)"
    } elseif ($rssiValue -ge 26 -and $rssiValue -le 31) {
        $bars = 5; $color = [System.Drawing.Color]::Green; $statusText = "Max ($rssiValue RSSI)"
    } elseif ($rssiValue -eq 99) {
        $bars = 0; $statusText = "No Service (99)"
    } elseif ($rssiValue -eq 0) {
        $bars = 0; $statusText = "Min (-113dBm)"
    }
    
    $signalValueLabel.Text = $statusText
    
    # Update the visual indicator bars
    for ($i = 0; $i -lt $signalBarLabels.Count; $i++) {
        if ($i -lt $bars) {
            # Bar is active
            $signalBarLabels[$i].BackColor = $color
            $signalBarLabels[$i].BorderStyle = [System.Windows.Forms.BorderStyle]::None 
        } else {
            # Bar is inactive
            $signalBarLabels[$i].BackColor = $script:color_Panel
            $signalBarLabels[$i].BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        }
    }
}
$signalTimer.Add_Tick({ Update-SignalStatus }) # Hook up the timer to the update function

function Send-CustomCommand {
    param ([string]$Command)
    if (-not $script:serialPort -or -not $script:serialPort.IsOpen) {
        $statusTextBox.Text = "Error: Port not open."
        Log-Output "Error: Port not open. Cannot send command."
        return
    }
    
    Log-Output ">>> Sending custom command: $Command"
    # This call is non-silent by default
    $response = Send-AtCommand -Command $Command 
    
    Log-Output "<<< Response received:"
    if ($response -ne "TIMEOUT" -and $response -ne "ERROR") {
        $response.Split("`r`n") | Where-Object { $_ -ne "" } | ForEach-Object {
            Log-Output $_
        }
    }
}

# FORM 
# ====================================================================================

# --- Main Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Professional GSM Module AT Command Interface"
# Size: 1040 (W) x 920 (H)
$form.Size = New-Object System.Drawing.Size(1040, 920) 
$form.MinimumSize = New-Object System.Drawing.Size(840, 700) 
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor = $script:color_Background # Apply Dark Mode Background
$form.ForeColor = $script:color_Text 

# --- Column Layout Constants ---
$paddingLeft = 15; $paddingCenter = 30; $c2Width = 400
# $c2X automatically uses the new $form.Width (1040)
$c2X = $form.Width - $c2Width - $paddingLeft 


# --- LEFT COLUMN: Connection (Y=15) ---
$connectGroupBox = New-Object System.Windows.Forms.GroupBox; $connectGroupBox.Text = "Connection Status"; $connectGroupBox.Location = New-Object System.Drawing.Size($paddingLeft, 15); $connectGroupBox.Height = 100
$connectGroupBox.Anchor = $anchorLeftRightTop; $connectGroupBox.Width = $form.Width - $paddingLeft - $c2Width - $paddingCenter - $paddingLeft
$connectGroupBox.BackColor = $script:color_Panel; $connectGroupBox.ForeColor = $script:color_Text;

# Com Port Input (All elements moved down by 5px to prevent clipping)
$comPortLabel = New-Object System.Windows.Forms.Label; $comPortLabel.Text = "COM Port:"; $comPortLabel.AutoSize = $true; $comPortLabel.Location = New-Object System.Drawing.Size(15, 25); $comPortLabel.ForeColor = $script:color_Text 
$comPortDropdown = New-Object System.Windows.Forms.ComboBox; $comPortDropdown.Location = New-Object System.Drawing.Size(15, 45); $comPortDropdown.Size = New-Object System.Drawing.Size(80, 20) 
$comPortDropdown.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $comPortDropdown.BackColor = $script:color_ButtonBG; $comPortDropdown.ForeColor = $script:color_Text

[System.IO.Ports.SerialPort]::GetPortNames() | ForEach-Object { [void]$comPortDropdown.Items.Add($_) }
if ($comPortDropdown.Items.Count -gt 0) { $comPortDropdown.SelectedIndex = 0 }

# Connect Button Styling 
$connectButton = New-Object System.Windows.Forms.Button; $connectButton.Text = "Connect"; $connectButton.Size = New-Object System.Drawing.Size(100, 30); $connectButton.Location = New-Object System.Drawing.Size(110, 45)
# Connect and Disconnect handlers do NOT use Log-Separator as they have their own visual separators (--- lines)
$connectButton.Add_Click({ Open-Port $comPortDropdown.SelectedItem }); $connectButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $connectButton.BackColor = $script:color_Accent; $connectButton.ForeColor = [System.Drawing.Color]::White

# Disconnect Button Styling 
$disconnectButton = New-Object System.Windows.Forms.Button; $disconnectButton.Text = "Disconnect"; $disconnectButton.Size = New-Object System.Drawing.Size(100, 30); $disconnectButton.Location = New-Object System.Drawing.Size(220, 45); $disconnectButton.Enabled = $false
$disconnectButton.Add_Click({ Close-Port }); $disconnectButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $disconnectButton.BackColor = $script:color_ButtonBG; $disconnectButton.ForeColor = $script:color_Text

# Dedicated Connection Indicator LED 
$connectionIndicator = New-Object System.Windows.Forms.Label; $connectionIndicator.Text = "OFFLINE"; $connectionIndicator.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$connectionIndicator.BackColor = [System.Drawing.Color]::Red; $connectionIndicator.ForeColor = [System.Drawing.Color]::White
$connectionIndicator.Location = New-Object System.Drawing.Size(340, 45); $connectionIndicator.Size = New-Object System.Drawing.Size(100, 30)
$connectionIndicator.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$connectionIndicator.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

$connectGroupBox.Controls.Add($comPortLabel); $connectGroupBox.Controls.Add($comPortDropdown); $connectGroupBox.Controls.Add($connectButton); $connectGroupBox.Controls.Add($disconnectButton); $connectGroupBox.Controls.Add($connectionIndicator)
$form.Controls.Add($connectGroupBox)

# --- LEFT COLUMN: Status Panel (Y=125) ---
$statusGroupBox = New-Object System.Windows.Forms.GroupBox; $statusGroupBox.Text = "Current Status"; $statusGroupBox.Height = 50
$statusGroupBox.Location = New-Object System.Drawing.Size($paddingLeft, 125); $statusGroupBox.Anchor = $anchorLeftRightTop; $statusGroupBox.Width = $connectGroupBox.Width 
$statusGroupBox.BackColor = $script:color_Panel; $statusGroupBox.ForeColor = $script:color_Text

$statusLabel = New-Object System.Windows.Forms.Label; $statusLabel.Text = "Status:"; $statusLabel.AutoSize = $true; $statusLabel.Location = New-Object System.Drawing.Size(10, 20); $statusLabel.ForeColor = $script:color_Text
$statusTextBox = New-Object System.Windows.Forms.TextBox; $statusTextBox.Text = "Ready. Select COM Port and Connect."; $statusTextBox.ReadOnly = $true
$statusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None; $statusTextBox.Location = New-Object System.Drawing.Size(60, 20); $statusTextBox.Anchor = $anchorLeftRight 
$statusTextBox.Width = 330; $statusTextBox.BackColor = $script:color_Panel; $statusTextBox.ForeColor = $script:color_Text

$statusGroupBox.Controls.Add($statusLabel); $statusGroupBox.Controls.Add($statusTextBox)
$form.Controls.Add($statusGroupBox)

# --- LEFT COLUMN: Custom Command Panel (Y=185) ---
$customGroupBox = New-Object System.Windows.Forms.GroupBox; $customGroupBox.Text = "Custom AT Command Sender"; $customGroupBox.Size = New-Object System.Drawing.Size($connectGroupBox.Width, 120) 
$customGroupBox.Location = New-Object System.Drawing.Size($paddingLeft, 185); $customGroupBox.Anchor = $anchorLeftRightTop; $customGroupBox.Enabled = $false
$customGroupBox.BackColor = $script:color_Panel; $customGroupBox.ForeColor = $script:color_Text

# Textbox height reduced by one line
$customCommandTextBox = New-Object System.Windows.Forms.TextBox; $customCommandTextBox.Multiline = $true; $customCommandTextBox.Size = New-Object System.Drawing.Size(380, 40)
$customCommandTextBox.Location = New-Object System.Drawing.Size(10, 25); $customCommandTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$customCommandTextBox.Text = "AT+BTSCAN?"; $customCommandTextBox.Anchor = $anchorLeftRight; $customCommandTextBox.BackColor = $script:color_ButtonBG; $customCommandTextBox.ForeColor = $script:color_Text

# Button position adjusted
$sendCommandButton = New-Object System.Windows.Forms.Button; $sendCommandButton.Text = "Send Command"; $sendCommandButton.Size = New-Object System.Drawing.Size(150, 35)
$sendCommandButton.Location = New-Object System.Drawing.Size(10, 75); 
$sendCommandButton.Add_Click({ Log-Separator; Send-CustomCommand $customCommandTextBox.Text }) # ADDED SEPARATOR
$sendCommandButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $sendCommandButton.BackColor = $script:color_ButtonBG; $sendCommandButton.ForeColor = $script:color_Text

$customGroupBox.Controls.Add($customCommandTextBox); $customGroupBox.Controls.Add($sendCommandButton)
$form.Controls.Add($customGroupBox)

# --- LEFT COLUMN: Log Control Panel (Y=315) ---
$logControlGroupBox = New-Object System.Windows.Forms.GroupBox; $logControlGroupBox.Text = "Log Controls"; $logControlGroupBox.Height = 50
$logControlGroupBox.Location = New-Object System.Drawing.Size($paddingLeft, 315); $logControlGroupBox.Anchor = $anchorLeftRightTop
$logControlGroupBox.Width = $connectGroupBox.Width; $logControlGroupBox.BackColor = $script:color_Panel; $logControlGroupBox.ForeColor = $script:color_Text

$clearLogButton = New-Object System.Windows.Forms.Button; $clearLogButton.Text = "Clear Log"; $clearLogButton.Size = New-Object System.Drawing.Size(140, 30); $clearLogButton.Location = New-Object System.Drawing.Size(10, 15)
$clearLogButton.Add_Click({ $outputTextBox.Clear(); $statusTextBox.Text = "Output Log Cleared." })
$clearLogButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $clearLogButton.BackColor = $script:color_ButtonBG; $clearLogButton.ForeColor = $script:color_Text

$saveLogButton = New-Object System.Windows.Forms.Button; $saveLogButton.Text = "Save Log to File"; $saveLogButton.Size = New-Object System.Drawing.Size(140, 30); $saveLogButton.Location = New-Object System.Drawing.Size(160, 15)
$saveLogButton.Add_Click({ 
    Add-Type -AssemblyName System.Windows.Forms
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Log Files (*.log)|*.log|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    $saveFileDialog.FileName = "GSM_Module_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try { $outputTextBox.Text | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8; $statusTextBox.Text = "Log saved successfully." } 
        catch { $statusTextBox.Text = "Error saving log: $($_.Exception.Message)" }
    } else { $statusTextBox.Text = "Log save cancelled." }
})
$saveLogButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $saveLogButton.BackColor = $script:color_ButtonBG; $saveLogButton.ForeColor = $script:color_Text

$logControlGroupBox.Controls.Add($clearLogButton); $logControlGroupBox.Controls.Add($saveLogButton)
$form.Controls.Add($logControlGroupBox)

# --- LEFT COLUMN: Output Panel (Y=375) ---
$outputGroupBox = New-Object System.Windows.Forms.GroupBox; $outputGroupBox.Text = "Output Log"; $outputGroupBox.Location = New-Object System.Drawing.Size($paddingLeft, 375) 
$outputGroupBox.Anchor = $anchorAll; $outputGroupBox.Width = $connectGroupBox.Width; $outputGroupBox.Height = $form.Height - 375 - 50 
$outputGroupBox.BackColor = $script:color_Panel; $outputGroupBox.ForeColor = $script:color_Text
 
$outputTextBox = New-Object System.Windows.Forms.TextBox; $outputTextBox.Multiline = $true; $outputTextBox.Location = New-Object System.Drawing.Size(10, 20)
$outputTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill; $outputTextBox.ScrollBars = "Vertical"; $outputTextBox.ReadOnly = $true
$outputTextBox.Font = New-Object System.Drawing.Font("Consolas", 10); $outputTextBox.BackColor = $script:color_ButtonBG; $outputTextBox.ForeColor = $script:color_Text # Dark Log Background

$outputGroupBox.Controls.Add($outputTextBox)
$form.Controls.Add($outputGroupBox)

# --- Column 2 (Right Side) ---
# ------------------------------------------------------------------------------------
$currentY = 15

$form.Add_Resize({
    $c2XNew = $form.ClientSize.Width - $c2Width - $paddingLeft
    $smsGroupBox.Location = New-Object System.Drawing.Size($c2XNew, $smsGroupBox.Location.Y)
    $httpGroupBox.Location = New-Object System.Drawing.Size($c2XNew, $httpGroupBox.Location.Y)
    $extGroupBox.Location = New-Object System.Drawing.Size($c2XNew, $extGroupBox.Location.Y)
    $btGroupBox.Location = New-Object System.Drawing.Size($c2XNew, $btGroupBox.Location.Y) 
    
    $c1Width = $form.ClientSize.Width - $c2Width - $paddingCenter - $paddingLeft
    $connectGroupBox.Width = $c1Width
    $statusGroupBox.Width = $c1Width
    $customGroupBox.Width = $c1Width
    $logControlGroupBox.Width = $c1Width
    $outputGroupBox.Width = $c1Width
})

# --- Apply dark theme 
function Style-Button {
    param($button, $useAccent = $false)
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.BackColor = if ($useAccent) { $script:color_Accent } else { $script:color_ButtonBG }
    $button.ForeColor = $script:color_Text
    if ($useAccent) { $button.ForeColor = [System.Drawing.Color]::White }
}

# --- 1. SMS and Calls  ---
$smsGroupBox = New-Object System.Windows.Forms.GroupBox; $smsGroupBox.Text = "1. SMS and Calls"; $smsGroupBox.Size = New-Object System.Drawing.Size($c2Width, 210) 
$smsGroupBox.Location = New-Object System.Drawing.Size($c2X, $currentY); $smsGroupBox.Anchor = $anchorTopRight; $smsGroupBox.Enabled = $false
$smsGroupBox.BackColor = $script:color_Panel; $smsGroupBox.ForeColor = $script:color_Text

$numberLabel = New-Object System.Windows.Forms.Label; $numberLabel.Text = "Number:"; $numberLabel.Location = New-Object System.Drawing.Size(10, 25); $numberLabel.AutoSize = $true; $numberLabel.ForeColor = $script:color_Text
$numberTextBox = New-Object System.Windows.Forms.TextBox; $numberTextBox.Location = New-Object System.Drawing.Size(80, 23); $numberTextBox.Width = 310; $numberTextBox.Anchor = $anchorLeftRight
$numberTextBox.Text = "+1234567890"; $numberTextBox.BackColor = $script:color_ButtonBG; $numberTextBox.ForeColor = $script:color_Text

$messageLabel = New-Object System.Windows.Forms.Label; $messageLabel.Text = "Message:"; $messageLabel.Location = New-Object System.Drawing.Size(10, 55); $messageLabel.AutoSize = $true; $messageLabel.ForeColor = $script:color_Text
$messageTextBox = New-Object System.Windows.Forms.TextBox; $messageTextBox.Multiline = $true; $messageTextBox.ScrollBars = "Vertical"
$messageTextBox.Location = New-Object System.Drawing.Size(10, 75); $messageTextBox.Size = New-Object System.Drawing.Size(380, 70); $messageTextBox.Anchor = $anchorLeftRight
$messageTextBox.Text = "Hello from PowerShell!"; $messageTextBox.BackColor = $script:color_ButtonBG; $messageTextBox.ForeColor = $script:color_Text

$sendSmsButton = New-Object System.Windows.Forms.Button; $sendSmsButton.Text = "Send SMS (CMGS)"; $sendSmsButton.Location = New-Object System.Drawing.Size(10, 155); $sendSmsButton.Width = 120; Style-Button $sendSmsButton -useAccent $true
$sendSmsButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+CMGF=1"; Send-CustomCommand "AT+CMGS=`"$($numberTextBox.Text)`"" }) # ADDED SEPARATOR

$dialButton = New-Object System.Windows.Forms.Button; $dialButton.Text = "Dial (ATD)"; $dialButton.Location = New-Object System.Drawing.Size(135, 155); $dialButton.Width = 80; Style-Button $dialButton
$dialButton.Add_Click({ Log-Separator; Send-CustomCommand "ATD$($numberTextBox.Text);" }) # ADDED SEPARATOR

$hangUpButton = New-Object System.Windows.Forms.Button; $hangUpButton.Text = "Hang Up (ATH)"; $hangUpButton.Location = New-Object System.Drawing.Size(220, 155); $hangUpButton.Width = 85; Style-Button $hangUpButton
$hangUpButton.Add_Click({ Log-Separator; Send-CustomCommand "ATH" }) # ADDED SEPARATOR

$answerButton = New-Object System.Windows.Forms.Button; $answerButton.Text = "Answer (ATA)"; $answerButton.Location = New-Object System.Drawing.Size(310, 155); $answerButton.Width = 85; Style-Button $answerButton
$answerButton.Add_Click({ Log-Separator; Send-CustomCommand "ATA" }) # ADDED SEPARATOR

$readSmsButton = New-Object System.Windows.Forms.Button; $readSmsButton.Text = "Read All SMS (CMGL)"; $readSmsButton.Location = New-Object System.Drawing.Size(10, 185); $readSmsButton.Width = 180; Style-Button $readSmsButton
$readSmsButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+CMGF=1"; Send-CustomCommand "AT+CMGL=`"ALL`"" }) # ADDED SEPARATOR

$delAllSmsButton = New-Object System.Windows.Forms.Button; $delAllSmsButton.Text = "Delete All SMS (CMGDA)"; $delAllSmsButton.Location = New-Object System.Drawing.Size(200, 185); $delAllSmsButton.Width = 195; Style-Button $delAllSmsButton
$delAllSmsButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+CMGDA=`"DEL ALL`"" }) # ADDED SEPARATOR

$smsGroupBox.Controls.Add($numberLabel); $smsGroupBox.Controls.Add($numberTextBox); $smsGroupBox.Controls.Add($messageLabel); $smsGroupBox.Controls.Add($messageTextBox)
$smsGroupBox.Controls.Add($sendSmsButton); $smsGroupBox.Controls.Add($dialButton); $smsGroupBox.Controls.Add($hangUpButton); $smsGroupBox.Controls.Add($answerButton)
$smsGroupBox.Controls.Add($readSmsButton); $smsGroupBox.Controls.Add($delAllSmsButton)
$form.Controls.Add($smsGroupBox)

$currentY += $smsGroupBox.Height + 10 

# --- 2. GPRS & HTTP Client Panel ---
$httpGroupBox = New-Object System.Windows.Forms.GroupBox; $httpGroupBox.Text = "2. GPRS & HTTP Client"; $httpGroupBox.Size = New-Object System.Drawing.Size($c2Width, 240)
$httpGroupBox.Location = New-Object System.Drawing.Size($c2X, $currentY); $httpGroupBox.Anchor = $anchorTopRight; $httpGroupBox.Enabled = $false
$httpGroupBox.BackColor = $script:color_Panel; $httpGroupBox.ForeColor = $script:color_Text

$gprsActivateButton = New-Object System.Windows.Forms.Button; $gprsActivateButton.Text = "1. GPRS Activate"; $gprsActivateButton.Location = New-Object System.Drawing.Size(10, 25); $gprsActivateButton.Width = 120; Style-Button $gprsActivateButton -useAccent $true
$gprsActivateButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+SAPBR=3,1,`"CONTYPE`",`"GPRS`""; Send-CustomCommand "AT+SAPBR=3,1,`"APN`",`"internet`""; Send-CustomCommand "AT+SAPBR=1,1" }) # ADDED SEPARATOR

$checkGprsIpButton = New-Object System.Windows.Forms.Button; $checkGprsIpButton.Text = "Check IP (SAPBR=2)"; $checkGprsIpButton.Location = New-Object System.Drawing.Size(140, 25); $checkGprsIpButton.Width = 120; Style-Button $checkGprsIpButton
$checkGprsIpButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+SAPBR=2,1" }) # ADDED SEPARATOR

$gprsDeactivateButton = New-Object System.Windows.Forms.Button; $gprsDeactivateButton.Text = "GPRS Deactivate"; $gprsDeactivateButton.Location = New-Object System.Drawing.Size(270, 25); $gprsDeactivateButton.Width = 120; Style-Button $gprsDeactivateButton
$gprsDeactivateButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+SAPBR=0,1" }) # ADDED SEPARATOR

$httpInitButton = New-Object System.Windows.Forms.Button; $httpInitButton.Text = "2. HTTP Init"; $httpInitButton.Location = New-Object System.Drawing.Size(10, 60); $httpInitButton.Width = 100; Style-Button $httpInitButton
$httpInitButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+HTTPINIT" }) # ADDED SEPARATOR

$httpTermButton = New-Object System.Windows.Forms.Button; $httpTermButton.Text = "HTTP Term"; $httpTermButton.Location = New-Object System.Drawing.Size(120, 60); $httpTermButton.Width = 100; Style-Button $httpTermButton
$httpTermButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+HTTPTERM" }) # ADDED SEPARATOR

$httpStatusButton = New-Object System.Windows.Forms.Button; $httpStatusButton.Text = "HTTP Status"; $httpStatusButton.Location = New-Object System.Drawing.Size(230, 60); $httpStatusButton.Width = 100; Style-Button $httpStatusButton
$httpStatusButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+HTTPSTATUS" }) # ADDED SEPARATOR

$urlLabel = New-Object System.Windows.Forms.Label; $urlLabel.Text = "URL:"; $urlLabel.Location = New-Object System.Drawing.Size(10, 95); $urlLabel.AutoSize = $true; $urlLabel.ForeColor = $script:color_Text
$urlTextBox = New-Object System.Windows.Forms.TextBox; $urlTextBox.Text = "http://httpbin.org/get"; $urlTextBox.Location = New-Object System.Drawing.Size(80, 93); $urlTextBox.Width = 310; $urlTextBox.Anchor = $anchorLeftRight; $urlTextBox.BackColor = $script:color_ButtonBG; $urlTextBox.ForeColor = $script:color_Text

$setUrlButton = New-Object System.Windows.Forms.Button; $setUrlButton.Text = "3. Set URL (HTTPPARA)"; $setUrlButton.Location = New-Object System.Drawing.Size(10, 125); $setUrlButton.Width = 140; Style-Button $setUrlButton
$setUrlButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+SAPBR=1,1"; Send-CustomCommand "AT+HTTPINIT"; Send-CustomCommand "AT+HTTPPARA=`"URL`",`"$($urlTextBox.Text)`"" }) # ADDED SEPARATOR

$sendGetButton = New-Object System.Windows.Forms.Button; $sendGetButton.Text = "4. Send GET and Read Data"; $sendGetButton.Location = New-Object System.Drawing.Size(160, 125); $sendGetButton.Width = 230; Style-Button $sendGetButton -useAccent $true
$sendGetButton.Add_Click({ 
    Log-Separator # ADDED SEPARATOR
    $initialResponse = Send-AtCommand "AT+HTTPACTION=0" -PostOkReadMs 0
    if ($initialResponse -notmatch "OK") { Log-Output "!!! ERROR: Command AT+HTTPACTION=0 failed. Aborting." ; return }
    if (Wait-ForHttpAction) { $oldTimeout = $script:timeout; $script:timeout = 15000; Send-CustomCommand "AT+HTTPREAD"; $script:timeout = $oldTimeout }
})

$postDataLabel = New-Object System.Windows.Forms.Label; $postDataLabel.Text = "POST Data:"; $postDataLabel.Location = New-Object System.Drawing.Size(10, 160); $postDataLabel.AutoSize = $true; $postDataLabel.ForeColor = $script:color_Text
$postDataTextBox = New-Object System.Windows.Forms.TextBox; $postDataTextBox.Text = "key1=value1"; $postDataTextBox.Location = New-Object System.Drawing.Size(10, 180); $postDataTextBox.Width = 380; $postDataTextBox.Anchor = $anchorLeftRight
$postDataTextBox.Multiline = $true; $postDataTextBox.Height = 25; $postDataTextBox.BackColor = $script:color_ButtonBG; $postDataTextBox.ForeColor = $script:color_Text

$setPostButton = New-Object System.Windows.Forms.Button; $setPostButton.Text = "Set POST Type"; $setPostButton.Location = New-Object System.Drawing.Size(10, 210); $setPostButton.Width = 120; Style-Button $setPostButton
$setPostButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+HTTPPARA=`"CONTENT`",`"application/x-www-form-urlencoded`"" }) # ADDED SEPARATOR

$sendPostButton = New-Object System.Windows.Forms.Button; $sendPostButton.Text = "4. Send POST (ACTION=1)"; $sendPostButton.Location = New-Object System.Drawing.Size(140, 210); $sendPostButton.Width = 135; Style-Button $sendPostButton -useAccent $true
$sendPostButton.Add_Click({ 
    Log-Separator # ADDED SEPARATOR
    $initialResponse = Send-AtCommand "AT+HTTPACTION=1" -PostOkReadMs 0
    if ($initialResponse -notmatch "OK") { Log-Output "!!! ERROR: Command AT+HTTPACTION=1 failed. Aborting." ; return }
    Wait-ForHttpAction 
})

$httpGroupBox.Controls.Add($gprsActivateButton); $httpGroupBox.Controls.Add($checkGprsIpButton); $httpGroupBox.Controls.Add($gprsDeactivateButton)
$httpGroupBox.Controls.Add($httpInitButton); $httpGroupBox.Controls.Add($httpTermButton); $httpGroupBox.Controls.Add($httpStatusButton)
$httpGroupBox.Controls.Add($urlLabel); $httpGroupBox.Controls.Add($urlTextBox); $httpGroupBox.Controls.Add($setUrlButton); $httpGroupBox.Controls.Add($sendGetButton)
$httpGroupBox.Controls.Add($postDataLabel); $httpGroupBox.Controls.Add($postDataTextBox); $httpGroupBox.Controls.Add($setPostButton); $httpGroupBox.Controls.Add($sendPostButton)
$form.Controls.Add($httpGroupBox)

$currentY += $httpGroupBox.Height + 10

# --- 3. Extended/Diagnostics & SIGNAL STATUS Panel ---
$extGroupBox = New-Object System.Windows.Forms.GroupBox; $extGroupBox.Text = "3. Diagnostics & Status (Live Signal)"; $extGroupBox.Size = New-Object System.Drawing.Size($c2Width, 200) 
$extGroupBox.Location = New-Object System.Drawing.Size($c2X, $currentY); $extGroupBox.Anchor = $anchorTopRight; $extGroupBox.Enabled = $false
$extGroupBox.BackColor = $script:color_Panel; $extGroupBox.ForeColor = $script:color_Text

# --- Signal Strength Indicator ---
$signalStatusLabel = New-Object System.Windows.Forms.Label; $signalStatusLabel.Text = "Signal Status:"; $signalStatusLabel.Location = New-Object System.Drawing.Size(10, 30); $signalStatusLabel.AutoSize = $true; $signalStatusLabel.ForeColor = $script:color_Text

$signalPanel = New-Object System.Windows.Forms.Panel; $signalPanel.Location = New-Object System.Drawing.Size(110, 25); $signalPanel.Size = New-Object System.Drawing.Size(150, 30); $signalPanel.BackColor = $script:color_ButtonBG; $signalPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$signalBarLabels = @()
[int]$barWidth = 20; [int]$barSpacing = 4

for ($i = 0; $i -lt 5; $i++) {
    $bar = New-Object System.Windows.Forms.Label
    [int]$barX = [int]$i * ([int]$barWidth + [int]$barSpacing) + 5
    [int]$barY = 5 + (([int]$i % 2) * 2) 
    $bar.Location = New-Object System.Drawing.Size($barX, $barY)
    [int]$barHeight = 20 - (([int]$i % 2) * 2)
    $bar.Size = New-Object System.Drawing.Size([int]$barWidth, $barHeight) 
    $bar.BackColor = $script:color_Panel
    $bar.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $signalPanel.Controls.Add($bar)
    $signalBarLabels += $bar
}

$signalValueLabel = New-Object System.Windows.Forms.Label; $signalValueLabel.Text = "Disconnected"; $signalValueLabel.AutoSize = $true; $signalValueLabel.Location = New-Object System.Drawing.Size(270, 32); $signalValueLabel.ForeColor = $script:color_Text; $signalValueLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$extGroupBox.Controls.Add($signalStatusLabel); $extGroupBox.Controls.Add($signalPanel); $extGroupBox.Controls.Add($signalValueLabel)

# Row 2: Status Checks (Moved down)
$imeiButton = New-Object System.Windows.Forms.Button; $imeiButton.Text = "Get IMEI (GSN)"; $imeiButton.Size = New-Object System.Drawing.Size(120, 30); $imeiButton.Location = New-Object System.Drawing.Size(10, 65); Style-Button $imeiButton
$imeiButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+GSN" }) # ADDED SEPARATOR

$getFirmwareButton = New-Object System.Windows.Forms.Button; $getFirmwareButton.Text = "Get Firmware (ATI)"; $getFirmwareButton.Size = New-Object System.Drawing.Size(120, 30); $getFirmwareButton.Location = New-Object System.Drawing.Size(140, 65); Style-Button $getFirmwareButton
$getFirmwareButton.Add_Click({ Log-Separator; Send-CustomCommand "ATI" }) # ADDED SEPARATOR

$checkPinButton = New-Object System.Windows.Forms.Button; $checkPinButton.Text = "Check SIM (CPIN)"; $checkPinButton.Size = New-Object System.Drawing.Size(120, 30); $checkPinButton.Location = New-Object System.Drawing.Size(270, 65); Style-Button $checkPinButton
$checkPinButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+CPIN?" }) # ADDED SEPARATOR

# Row 3: Maintenance
$testAtButton = New-Object System.Windows.Forms.Button; $testAtButton.Text = "Test AT"; $testAtButton.Size = New-Object System.Drawing.Size(120, 30); $testAtButton.Location = New-Object System.Drawing.Size(10, 105); Style-Button $testAtButton
$testAtButton.Add_Click({ Log-Separator; Send-CustomCommand "AT" }) # ADDED SEPARATOR

$resetButton = New-Object System.Windows.Forms.Button; $resetButton.Text = "Soft Reset (CFUN=1,1)"; $resetButton.Size = New-Object System.Drawing.Size(140, 30); $resetButton.Location = New-Object System.Drawing.Size(140, 105); Style-Button $resetButton
$resetButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+CFUN=1,1" }) # ADDED SEPARATOR

$shutdownButton = New-Object System.Windows.Forms.Button; $shutdownButton.Text = "Shutdown (CPOWD)"; $shutdownButton.Size = New-Object System.Drawing.Size(100, 30); $shutdownButton.Location = New-Object System.Drawing.Size(290, 105); Style-Button $shutdownButton
$shutdownButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+CPOWD=1" }) # ADDED SEPARATOR

# Row 4: GPS
$enableGpsButton = New-Object System.Windows.Forms.Button; $enableGpsButton.Text = "GPS Power ON"; $enableGpsButton.Size = New-Object System.Drawing.Size(120, 30); $enableGpsButton.Location = New-Object System.Drawing.Size(10, 145); Style-Button $enableGpsButton
$enableGpsButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+CGNSPWR=1" }) # ADDED SEPARATOR

$getGpsButton = New-Object System.Windows.Forms.Button; $getGpsButton.Text = "Get GPS Info"; $getGpsButton.Size = New-Object System.Drawing.Size(120, 30); $getGpsButton.Location = New-Object System.Drawing.Size(140, 145); Style-Button $getGpsButton
$getGpsButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+CGNSINF" }) # ADDED SEPARATOR

$disableGpsButton = New-Object System.Windows.Forms.Button; $disableGpsButton.Text = "GPS Power OFF"; $disableGpsButton.Size = New-Object System.Drawing.Size(120, 30); $disableGpsButton.Location = New-Object System.Drawing.Size(270, 145); Style-Button $disableGpsButton
$disableGpsButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+CGNSPWR=0" }) # ADDED SEPARATOR

$extGroupBox.Controls.Add($imeiButton); $extGroupBox.Controls.Add($getFirmwareButton); $extGroupBox.Controls.Add($checkPinButton)
$extGroupBox.Controls.Add($testAtButton); $extGroupBox.Controls.Add($resetButton); $extGroupBox.Controls.Add($shutdownButton)
$extGroupBox.Controls.Add($enableGpsButton); $extGroupBox.Controls.Add($getGpsButton); $extGroupBox.Controls.Add($disableGpsButton)
$form.Controls.Add($extGroupBox)

$currentY += $extGroupBox.Height + 10

# --- 4. Bluetooth Panel ---
$btGroupBox = New-Object System.Windows.Forms.GroupBox; $btGroupBox.Text = "4. Bluetooth Control (Diagnostics & Scan)"; $btGroupBox.Size = New-Object System.Drawing.Size($c2Width, 180) 
$btGroupBox.Location = New-Object System.Drawing.Size($c2X, $currentY); $btGroupBox.Anchor = $anchorTopRight; $btGroupBox.Enabled = $false
$btGroupBox.BackColor = $script:color_Panel; $btGroupBox.ForeColor = $script:color_Text

$btPowerOnButton = New-Object System.Windows.Forms.Button; $btPowerOnButton.Text = "Power ON (BTPOWER=1)"; $btPowerOnButton.Size = New-Object System.Drawing.Size(170, 30); $btPowerOnButton.Location = New-Object System.Drawing.Size(10, 25); Style-Button $btPowerOnButton -useAccent $true
$btPowerOnButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+BTPOWER=1" }) # ADDED SEPARATOR

$btPowerOffButton = New-Object System.Windows.Forms.Button; $btPowerOffButton.Text = "Power OFF (BTPOWER=0)"; $btPowerOffButton.Size = New-Object System.Drawing.Size(180, 30); $btPowerOffButton.Location = New-Object System.Drawing.Size(200, 25); Style-Button $btPowerOffButton
$btPowerOffButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+BTPOWER=0" }) # ADDED SEPARATOR

$btTestSupportButton = New-Object System.Windows.Forms.Button; $btTestSupportButton.Text = "Test BT Support (? Query)"; $btTestSupportButton.Size = New-Object System.Drawing.Size(170, 30); $btTestSupportButton.Location = New-Object System.Drawing.Size(10, 65); Style-Button $btTestSupportButton
$btTestSupportButton.Add_Click({ Log-Separator; Send-CustomCommand "AT+BTPOWER=?"; Send-CustomCommand "AT+BTNAME=?"; Send-CustomCommand "AT+BTSCAN=?" }) # ADDED SEPARATOR

$btScanStartButton = New-Object System.Windows.Forms.Button; $btScanStartButton.Text = "1. Start Scan (BTSCAN=1)"; $btScanStartButton.Size = New-Object System.Drawing.Size(170, 30); $btScanStartButton.Location = New-Object System.Drawing.Size(190, 65); Style-Button $btScanStartButton
$btScanStartButton.Add_Click({ 
    Log-Separator # ADDED SEPARATOR
    Log-Output ">>> Sending AT+BTSCAN=1. Polling for results for 10s..."
    $response = Send-AtCommand -Command "AT+BTSCAN=1" -PostOkReadMs 10000
    Log-Output "<<< Scan Start Response (including any immediate URCs):"
    $response.Split("`r`n") | Where-Object { $_ -ne "" } | ForEach-Object { Log-Output $_ }
})

$btMacLabel = New-Object System.Windows.Forms.Label; $btMacLabel.Text = "Headphone MAC/Handle:"; $btMacLabel.Location = New-Object System.Drawing.Size(10, 105); $btMacLabel.AutoSize = $true; $btMacLabel.ForeColor = $script:color_Text

$btMacTextBox = New-Object System.Windows.Forms.TextBox; $btMacTextBox.Text = "00:11:22:AA:BB:CC"; $btMacTextBox.Location = New-Object System.Drawing.Size(160, 102); $btMacTextBox.Width = 230; $btMacTextBox.Anchor = $anchorLeftRight; $btMacTextBox.BackColor = $script:color_ButtonBG; $btMacTextBox.ForeColor = $script:color_Text

$btConnectButton = New-Object System.Windows.Forms.Button; $btConnectButton.Text = "2. Connect Headphone (BTCONNECT)"; $btConnectButton.Size = New-Object System.Drawing.Size(380, 30)
$btConnectButton.Location = New-Object System.Drawing.Size(10, 135); $btConnectButton.Anchor = $anchorLeftRight; Style-Button $btConnectButton -useAccent $true
$btConnectButton.Add_Click({ 
    Log-Separator # ADDED SEPARATOR
    $mac = $btMacTextBox.Text -replace ':', ''
    Log-Output ">>> Attempting to connect to MAC: $mac (Service ID 1 - HFP/HSP)..."
    Send-CustomCommand "AT+BTCONNECT=$mac,1" 
})

$btGroupBox.Controls.Add($btPowerOnButton); $btGroupBox.Controls.Add($btPowerOffButton); 
$btGroupBox.Controls.Add($btTestSupportButton); $btGroupBox.Controls.Add($btScanStartButton);
$btGroupBox.Controls.Add($btMacLabel); $btGroupBox.Controls.Add($btMacTextBox); $btGroupBox.Controls.Add($btConnectButton); 
$form.Controls.Add($btGroupBox)


# --- END ---
$form.Add_Closing({
    if ($script:serialPort -and $script:serialPort.IsOpen) { Close-Port }
    $signalTimer.Dispose()
})

$form.PerformLayout()

[void]$form.ShowDialog()
