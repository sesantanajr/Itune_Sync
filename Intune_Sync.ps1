# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to install/update necessary modules
function Install-Update-Modules {
    $modules = @("Microsoft.Graph")
    foreach ($module in $modules) {
        $installed = Get-InstalledModule -Name $module -ErrorAction SilentlyContinue
        if (-not $installed) {
            Write-Host "Instalando o modulo $module..."
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
        } else {
            $latest = Find-Module -Name $module
            if ($installed.Version -lt $latest.Version) {
                Write-Host "Atualizando o modulo $module..."
                Update-Module -Name $module -Force
            } else {
                Write-Host "O modulo $module ja esta instalado e atualizado."
            }
        }
    }
}

# Function to connect to Microsoft Graph
function Connect-ToGraph {
    Write-Host "Conectando ao Microsoft Graph..."
    Connect-MgGraph -Scopes "DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementManagedDevices.PrivilegedOperations.All"
}

# Function to disconnect from Microsoft Graph
function Disconnect-FromGraph {
    Write-Host "Desconectando do Microsoft Graph..."
    Disconnect-MgGraph
}

# Function to enable dmwappushservice
function Ensure-Dmwappushservice {
    $service = Get-Service -Name dmwappushservice -ErrorAction SilentlyContinue
    if ($service -and $service.Status -ne 'Running') {
        Write-Host "Habilitando e iniciando o servico dmwappushservice..."
        Set-Service -Name dmwappushservice -StartupType Automatic
        Start-Service -Name dmwappushservice
    }
}

# Function to create directory if not exists
function Create-DirectoryIfNotExists {
    param ([string]$path)
    if (-not (Test-Path -Path $path)) {
        New-Item -ItemType Directory -Path $path -Force
    }
}

# Function to log actions
function Log-Action {
    param (
        [string]$logPath,
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $message"
    Add-Content -Path $logPath -Value $logEntry
}

# Function to sync devices
function Sync-Devices {
    param (
        [string]$filter,
        [string]$logPath,
        [string]$errorLogPath,
        [string]$deviceType
    )
    try {
        $devices = Get-MgDeviceManagementManagedDevice -Filter $filter -ErrorAction Stop
        if ($devices.Count -eq 0) {
            $message = "Nao existem dispositivos configurados para esta plataforma."
            Write-Host $message -ForegroundColor Yellow
            Log-Action -logPath $logPath -message $message
            return @()
        }
        $totalDevices = $devices.Count
        $counter = 0
        $syncResults = @()
        foreach ($device in $devices) {
            $counter++
            $percentage = [math]::Round(($counter / $totalDevices) * 100, 0)
            $progressBar.Value = $percentage
            $progressLabel.Text = "$percentage%"
            $statusLabel.Text = "Sincronizando dispositivo: $($device.DeviceName)"
            [System.Windows.Forms.Application]::DoEvents() # Update UI in real-time
            try {
                Ensure-Dmwappushservice
                Sync-MgDeviceManagementManagedDevice -ManagedDeviceId $device.Id
                $logEntry = [PSCustomObject]@{
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    Mensagem = "Sincronizacao bem-sucedida"
                    NomeDoDispositivo = $device.DeviceName
                    Email = $device.UserPrincipalName
                    ComplianceState = $device.ComplianceState
                }
                $syncResults += $logEntry
                $logEntry | ConvertTo-Csv -NoTypeInformation | Out-File -Append -FilePath $logPath
            } catch {
                $errorMessage = "Erro ao sincronizar dispositivo $($device.DeviceName): $_"
                $errorEntry = [PSCustomObject]@{
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    Mensagem = $errorMessage
                }
                $syncResults += $errorEntry
                $errorEntry | ConvertTo-Csv -NoTypeInformation | Out-File -Append -FilePath $errorLogPath
            }
            Start-Sleep -Milliseconds 500
        }
        $completionMessage = "Sincronizacao concluida para $deviceType."
        Log-Action -logPath $logPath -message $completionMessage
        return $syncResults
    } catch {
        $errorMessage = "Erro ao obter dispositivos: $_"
        Log-Action -logPath $errorLogPath -message $errorMessage
        $statusLabel.Text = $errorMessage
        return @()
    }
}

# Function to sync all platforms
function Sync-AllPlatforms {
    param (
        [string]$logPath,
        [string]$errorLogPath
    )
    $syncResults = @()
    $syncResults += Sync-Devices -filter "operatingSystem eq 'Windows'" -logPath $logPath -errorLogPath $errorLogPath -deviceType "Windows"
    $syncResults += Sync-Devices -filter "operatingSystem eq 'Android'" -logPath $logPath -errorLogPath $errorLogPath -deviceType "Android"
    $syncResults += Sync-Devices -filter "operatingSystem eq 'macOS'" -logPath $logPath -errorLogPath $errorLogPath -deviceType "macOS"
    $syncResults += Sync-Devices -filter "operatingSystem eq 'ChromeOS'" -logPath $logPath -errorLogPath $errorLogPath -deviceType "ChromeOS"
    return $syncResults
}

# Function to open URLs in the default browser
function Open-Url {
    param ([string]$url)
    Start-Process $url
}

# Create log directory if it doesn't exist
$logDirectory = "C:\J365_Intune"
Create-DirectoryIfNotExists -path $logDirectory

# Verify and install/update necessary modules
Install-Update-Modules

# Connect to Microsoft Graph
Connect-ToGraph

# Create main window
$form = New-Object System.Windows.Forms.Form
$form.Text = "Jornada 365 | Intune Sync"
$form.Size = New-Object System.Drawing.Size(1000, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# Add logo
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.ImageLocation = "https://jornada365.cloud/wp-content/uploads/2024/03/Logotipo-Jornada-365-Home.png"
$pictureBox.Size = New-Object System.Drawing.Size(165, 55)
$pictureBox.Location = New-Object System.Drawing.Point(20, 10)
$pictureBox.SizeMode = "Zoom"
$form.Controls.Add($pictureBox)

# Add title
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Jornada 365 | Intune Sync"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 24, [System.Drawing.FontStyle]::Bold)
$titleLabel.Size = New-Object System.Drawing.Size(960, 50)
$titleLabel.Location = New-Object System.Drawing.Point(0, 10)
$titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$titleLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($titleLabel)

# Create group box
$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Text = "Device Sync"
$groupBox.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$groupBox.Size = New-Object System.Drawing.Size(960, 300)
$groupBox.Location = New-Object System.Drawing.Point(20, 80) # Adjusted for better alignment
$groupBox.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($groupBox)

# Add CheckBoxes for options
$options = @(
    "Sincronizar todos os dispositivos Windows (Fisicos)",
    "Sincronizar todos os dispositivos Android",
    "Sincronizar todos os dispositivos macOS",
    "Sincronizar todos os dispositivos ChromeOS",
    "Sincronizar todas as plataformas",
    "Abrir Jornada 365",
    "Abrir Intune Portal"
)

$y = 30
$checkBoxes = @()

foreach ($option in $options) {
    $checkBox = New-Object System.Windows.Forms.CheckBox
    $checkBox.Text = $option
    $checkBox.Location = New-Object System.Drawing.Point(10, $y)
    $checkBox.Font = New-Object System.Drawing.Font("Segoe UI", 12)
    $checkBox.AutoSize = $true
    $groupBox.Controls.Add($checkBox)
    $checkBoxes += $checkBox
    $y += 30
}

# Add label to show devices being synchronized
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(170, 400)
$statusLabel.Size = New-Object System.Drawing.Size(660, 30)
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$statusLabel.BackColor = [System.Drawing.Color]::White
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($statusLabel)

# Add label to show progress percentage
$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Location = New-Object System.Drawing.Point(900, 440) # Centered horizontally
$progressLabel.Size = New-Object System.Drawing.Size(60, 20)
$progressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$progressLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$form.Controls.Add($progressLabel)

# Add a stylish and thin progress bar with neon effect
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(100, 440) # Centered horizontally
$progressBar.Size = New-Object System.Drawing.Size(800, 10) # Thin progress bar
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$form.Controls.Add($progressBar)

# Customize the progress bar with a neon effect
$progressBarPaintEventHandler = [System.Windows.Forms.PaintEventHandler]{
    param ($sender, $e)
    $progressRectangle = [System.Drawing.Rectangle]::new(0, 0, [int]($progressBar.Value * ($progressBar.Width / $progressBar.Maximum)), $progressBar.Height)
    $brush = [System.Drawing.Drawing2D.LinearGradientBrush]::new($progressRectangle, [System.Drawing.Color]::FromArgb(0, 255, 0), [System.Drawing.Color]::FromArgb(0, 128, 255), [System.Drawing.Drawing2D.LinearGradientMode]::ForwardDiagonal)
    $e.Graphics.FillRectangle($brush, $progressRectangle)
}
$progressBar.add_Paint($progressBarPaintEventHandler)

# Add stylish and centered black buttons
$executeButton = New-Object System.Windows.Forms.Button
$executeButton.Text = "Executar"
$executeButton.Location = New-Object System.Drawing.Point(350, 480) # Centered horizontally
$executeButton.Size = New-Object System.Drawing.Size(150, 40)
$executeButton.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$executeButton.BackColor = [System.Drawing.Color]::Black
$executeButton.ForeColor = [System.Drawing.Color]::White
$executeButton.FlatStyle = "Flat"
$executeButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($executeButton)

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "Fechar"
$closeButton.Location = New-Object System.Drawing.Point(530, 480) # Centered horizontally
$closeButton.Size = New-Object System.Drawing.Size(150, 40)
$closeButton.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$closeButton.BackColor = [System.Drawing.Color]::Black
$closeButton.ForeColor = [System.Drawing.Color]::White
$closeButton.FlatStyle = "Flat"
$closeButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($closeButton)

# Function to execute selected option
$executeButton.Add_Click({
    $selectedOptions = $checkBoxes | Where-Object { $_.Checked }
    if (-not $selectedOptions) {
        [System.Windows.Forms.MessageBox]::Show("Por favor, selecione uma opcao.")
        return
    }
    $logPath = "$logDirectory\sync_log_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    $errorLogPath = "$logDirectory\error_log_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    $syncResults = @()
    foreach ($selectedOption in $selectedOptions) {
        switch ($selectedOption.Text) {
            "Sincronizar todos os dispositivos Windows (Fisicos)" {
                Write-Host "Sincronizando todos os dispositivos Windows (Fisicos)..."
                $syncResults += Sync-Devices -filter "operatingSystem eq 'Windows'" -logPath $logPath -errorLogPath $errorLogPath -deviceType "Windows"
            }
            "Sincronizar todos os dispositivos Android" {
                Write-Host "Sincronizando todos os dispositivos Android..."
                $syncResults += Sync-Devices -filter "operatingSystem eq 'Android'" -logPath $logPath -errorLogPath $errorLogPath -deviceType "Android"
            }
            "Sincronizar todos os dispositivos macOS" {
                Write-Host "Sincronizando todos os dispositivos macOS..."
                $syncResults += Sync-Devices -filter "operatingSystem eq 'macOS'" -logPath $logPath -errorLogPath $errorLogPath -deviceType "macOS"
            }
            "Sincronizar todos os dispositivos ChromeOS" {
                Write-Host "Sincronizando todos os dispositivos ChromeOS..."
                $syncResults += Sync-Devices -filter "operatingSystem eq 'ChromeOS'" -logPath $logPath -errorLogPath $errorLogPath -deviceType "ChromeOS"
            }
            "Sincronizar todas as plataformas" {
                Write-Host "Sincronizando todas as plataformas..."
                $syncResults += Sync-AllPlatforms -logPath $logPath -errorLogPath $errorLogPath
            }
            "Abrir Jornada 365" {
                Write-Host "Abrindo Jornada 365..."
                Open-Url -url "https://jornada365.cloud"
            }
            "Abrir Intune Portal" {
                Write-Host "Abrindo Intune Portal..."
                Open-Url -url "https://intune.microsoft.com/"
            }
        }
        $selectedOption.Checked = $false
    }

    # Update GUI with synchronization results
    $resultForm = New-Object System.Windows.Forms.Form
    $resultForm.Text = "Resultados da Sincronizacao"
    $resultForm.Size = New-Object System.Drawing.Size(800, 600)
    $resultForm.StartPosition = "CenterScreen"
    $resultForm.BackColor = [System.Drawing.Color]::White

    $resultTextBox = New-Object System.Windows.Forms.TextBox
    $resultTextBox.Multiline = $true
    $resultTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $resultTextBox.Size = New-Object System.Drawing.Size(760, 500)
    $resultTextBox.Location = New-Object System.Drawing.Point(10, 10)
    $resultTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $resultTextBox.ReadOnly = $true
    $resultTextBox.BackColor = [System.Drawing.Color]::White
    $resultForm.Controls.Add($resultTextBox)

    foreach ($result in $syncResults) {
        $resultTextBox.AppendText("$($result.Timestamp) - $($result.Mensagem) - $($result.NomeDoDispositivo) - $($result.Email) - $($result.ComplianceState)`n")
    }

    # Add save button with centered text
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Salvar"
    $saveButton.Location = New-Object System.Drawing.Point(350, 520)
    $saveButton.Size = New-Object System.Drawing.Size(100, 40)
    $saveButton.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $saveButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $saveButton.BackColor = [System.Drawing.Color]::Black
    $saveButton.ForeColor = [System.Drawing.Color]::White
    $saveButton.FlatStyle = "Flat"
    $resultForm.Controls.Add($saveButton)

    # Function to save results in CSV
    $saveButton.Add_Click({
        $csvPath = "$logDirectory\sync_results_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        $syncResults | Export-Csv -Path $csvPath -NoTypeInformation -Force
        [System.Windows.Forms.MessageBox]::Show("Resultados salvos em: $csvPath")
    })

    $resultForm.ShowDialog()
})

# Function to close the form
$closeButton.Add_Click({
    $form.Close()
    Disconnect-FromGraph
})

# Show the form
[void]$form.ShowDialog()

# Disconnect from Microsoft Graph when form is closed
Disconnect-FromGraph
