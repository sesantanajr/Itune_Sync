# Adicionar bibliotecas necessárias
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Função para instalar/atualizar módulos necessários
function Install-Update-Modules {
    $modules = @("Microsoft.Graph", "PSWriteHTML", "WindowsCompatibility")
    foreach ($module in $modules) {
        $installed = Get-InstalledModule -Name $module -ErrorAction SilentlyContinue
        if (-not $installed) {
            Write-Host "Instalando o módulo $module..."
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
        } else {
            $latest = Find-Module -Name $module
            if ($installed.Version -lt $latest.Version) {
                Write-Host "Atualizando o módulo $module..."
                Update-Module -Name $module -Force
            } else {
                Write-Host "O módulo $module já está instalado e atualizado."
            }
        }
    }
}

# Função para conectar ao Microsoft Graph com autenticação interativa
function Connect-ToGraph {
    Write-Host "Conectando ao Microsoft Graph..."
    Connect-MgGraph -Scopes "DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementManagedDevices.PrivilegedOperations.All"
}

# Função para desconectar do Microsoft Graph
function Disconnect-FromGraph {
    Write-Host "Desconectando do Microsoft Graph..."
    Disconnect-MgGraph
}

# Função para habilitar o serviço dmwappushservice
function Ensure-Dmwappushservice {
    $service = Get-Service -Name dmwappushservice -ErrorAction SilentlyContinue
    if ($service -and $service.Status -ne 'Running') {
        Write-Host "Habilitando e iniciando o serviço dmwappushservice..."
        Set-Service -Name dmwappushservice -StartupType Automatic
        Start-Service -Name dmwappushservice
    }
}

# Função para criar diretório se não existir
function Create-DirectoryIfNotExists {
    param ([string]$path)
    if (-not (Test-Path -Path $path)) {
        New-Item -ItemType Directory -Path $path -Force
    }
}

# Função para gerar log
function Log-Action {
    param (
        [string]$logPath,
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $message"
    Add-Content -Path $logPath -Value $logEntry
}

# Funções para sincronizar dispositivos
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
            $message = "Não existem dispositivos configurados para esta plataforma."
            Write-Host $message -ForegroundColor Yellow
            Log-Action -logPath $logPath -message $message
            $statusLabel.Text = $message
            return
        }
        $totalDevices = $devices.Count
        $counter = 0
        foreach ($device in $devices) {
            $counter++
            $percentage = [math]::Round(($counter / $totalDevices) * 100)
            $progressBar.Value = $percentage
            $statusLabel.Text = "Sincronizando dispositivo: $($device.DeviceName) - $percentage%"
            try {
                Ensure-Dmwappushservice
                Sync-MgDeviceManagementManagedDevice -ManagedDeviceId $device.Id
                $logEntry = [PSCustomObject]@{
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    Mensagem = "Sincronização bem-sucedida"
                    NomeDoDispositivo = $device.DeviceName
                    Email = $device.UserPrincipalName
                    Licenca = $device.OperatingSystem
                    ComplianceState = $device.ComplianceState
                    Autopilot = $device.Autopilot
                }
                $logEntry | ConvertTo-Csv -NoTypeInformation | Out-File -Append -FilePath $logPath
            } catch {
                $errorMessage = "Erro ao sincronizar dispositivo $($device.DeviceName): $_"
                $errorEntry = [PSCustomObject]@{
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    Mensagem = $errorMessage
                }
                $errorEntry | ConvertTo-Csv -NoTypeInformation | Out-File -Append -FilePath $errorLogPath
            }
            Start-Sleep -Milliseconds 500
        }
        $completionMessage = "Sincronização concluída para $deviceType."
        Log-Action -logPath $logPath -message $completionMessage
    } catch {
        $errorMessage = "Erro ao obter dispositivos: $_"
        Log-Action -logPath $errorLogPath -message $errorMessage
        $statusLabel.Text = $errorMessage
    }
}

function Resync-NoncompliantOrOutdatedDevices {
    param (
        [string]$logPath,
        [string]$errorLogPath
    )
    try {
        $now = Get-Date
        $threshold = $now.AddHours(-12)
        $devices = Get-MgDeviceManagementManagedDevice | Where-Object {
            $_.ComplianceState -ne "compliant" -or ($_.LastSyncDateTime -lt $threshold)
        }
        if ($devices.Count -eq 0) {
            $message = "Nenhum dispositivo não conforme ou desatualizado encontrado para sincronizar."
            Write-Host $message -ForegroundColor Yellow
            Log-Action -logPath $logPath -message $message
            $statusLabel.Text = $message
            return
        }
        $totalDevices = $devices.Count
        $counter = 0
        foreach ($device in $devices) {
            $counter++
            $percentage = [math]::Round(($counter / $totalDevices) * 100)
            $progressBar.Value = $percentage
            $statusLabel.Text = "Reparando dispositivo: $($device.DeviceName) - $percentage%"
            Ensure-Dmwappushservice
            try {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($device.Id)/syncDevice"
                $logEntry = [PSCustomObject]@{
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    Mensagem = "Reparo bem-sucedido"
                    NomeDoDispositivo = $device.DeviceName
                    Email = $device.UserPrincipalName
                    Licenca = $device.OperatingSystem
                    ComplianceState = $device.ComplianceState
                    Autopilot = $device.Autopilot
                }
                $logEntry | ConvertTo-Csv -NoTypeInformation | Out-File -Append -FilePath $logPath
            } catch {
                $errorMessage = "Erro ao reparar dispositivo $($device.DeviceName): $_"
                $errorEntry = [PSCustomObject]@{
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    Mensagem = $errorMessage
                }
                $errorEntry | ConvertTo-Csv -NoTypeInformation | Out-File -Append -FilePath $errorLogPath
            }
            Start-Sleep -Milliseconds 500
        }
        $completionMessage = "Reparo concluído."
        Log-Action -logPath $logPath -message $completionMessage
    } catch {
        $errorMessage = "Erro ao obter dispositivos: $_"
        Log-Action -logPath $errorLogPath -message $errorMessage
        $statusLabel.Text = $errorMessage
    }
}

function Force-Sync-AllDevices {
    param (
        [string]$logPath,
        [string]$errorLogPath
    )
    try {
        $devices = Get-MgDeviceManagementManagedDevice -All -ErrorAction Stop
        if ($devices.Count -eq 0) {
            $message = "Nenhum dispositivo encontrado para sincronizar."
            Write-Host $message -ForegroundColor Yellow
            Log-Action -logPath $logPath -message $message
            $statusLabel.Text = $message
            return
        }
        $totalDevices = $devices.Count
        $counter = 0
        foreach ($device in $devices) {
            $counter++
            $percentage = [math]::Round(($counter / $totalDevices) * 100, 2)
            $progressBar.Value = $percentage
            $statusLabel.Text = "Forçando sincronização do dispositivo: $($device.DeviceName) - $percentage%"
            try {
                Ensure-Dmwappushservice
                Sync-MgDeviceManagementManagedDevice -ManagedDeviceId $device.Id
                $logEntry = [PSCustomObject]@{
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    Mensagem = "Sincronização forçada bem-sucedida"
                    NomeDoDispositivo = $device.DeviceName
                    Email = $device.UserPrincipalName
                    Licenca = $device.OperatingSystem
                    ComplianceState = $device.ComplianceState
                    Autopilot = $device.Autopilot
                }
                $logEntry | ConvertTo-Csv -NoTypeInformation | Out-File -Append -FilePath $logPath
            } catch {
                $errorMessage = "Erro ao forçar sincronização do dispositivo $($device.DeviceName): $_"
                $errorEntry = [PSCustomObject]@{
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    Mensagem = $errorMessage
                }
                $errorEntry | ConvertTo-Csv -NoTypeInformation | Out-File -Append -FilePath $errorLogPath
            }
            Start-Sleep -Milliseconds 500
        }
        $completionMessage = "Forçamento de sincronização concluído."
        Log-Action -logPath $logPath -message $completionMessage
    } catch {
        $errorMessage = "Erro ao obter dispositivos: $_"
        Log-Action -logPath $errorLogPath -message $errorMessage
        $statusLabel.Text = $errorMessage
    }
}

# Função para abrir URLs no navegador padrão
function Open-Url {
    param ([string]$url)
    Start-Process $url
}

# Função para gerar relatório HTML
function Generate-HTMLReport {
    param (
        [string]$title,
        [string]$logPath,
        [string]$errorLogPath
    )

    if (-not (Test-Path -Path $logPath)) {
        Write-Host "Arquivo de log não encontrado: $logPath" -ForegroundColor Red
        return
    }

    if (-not (Test-Path -Path $errorLogPath)) {
        Write-Host "Arquivo de log de erros não encontrado: $errorLogPath" -ForegroundColor Yellow
        New-Item -Path $errorLogPath -ItemType File -Force | Out-Null
    }

    $logContent = Import-Csv -Path $logPath
    $errorLogContent = Import-Csv -Path $errorLogPath

    $deviceCounts = @{
        Windows10 = 0
        Windows11 = 0
        Android = 0
        macOS = 0
        ChromeOS = 0
        Autopilot = 0
        Compliant = 0
        NonCompliant = 0
    }

    foreach ($logEntry in $logContent) {
        switch ($logEntry.Licenca) {
            "Windows 10" { $deviceCounts.Windows10++ }
            "Windows 11" { $deviceCounts.Windows11++ }
            "Android" { $deviceCounts.Android++ }
            "macOS" { $deviceCounts.macOS++ }
            "ChromeOS" { $deviceCounts.ChromeOS++ }
        }
        if ($logEntry.Autopilot -eq "True") { $deviceCounts.Autopilot++ }
        if ($logEntry.ComplianceState -eq "compliant") { $deviceCounts.Compliant++ }
        else { $deviceCounts.NonCompliant++ }
    }

    $htmlContent = @"
<!DOCTYPE html>
<html lang='pt-br'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>$title</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            font-size: 14px;
        }
        h1, h2 {
            color: #333;
            text-align: center;
        }
        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .logo {
            width: 150px;
            height: auto;
        }
        .search-bar {
            text-align: right;
            margin-bottom: 20px;
        }
        .search-bar input {
            padding: 8px;
            font-size: 14px;
            width: 300px;
        }
        .table-container {
            overflow-x: auto;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }
        table, th, td {
            border: 1px solid #ccc;
        }
        th, td {
            padding: 10px;
            text-align: left;
        }
        th {
            background-color: #4CAF50;
            color: white;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        .chart {
            width: 80%;
            margin: auto;
        }
        .additional-info {
            width: 50%;
            float: left;
        }
        .additional-info p {
            margin: 5px 0;
        }
    </style>
    <script src='https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.9.4/Chart.min.js'></script>
    <script>
        function searchTable() {
            var input, filter, table, tr, td, i, j, txtValue;
            input = document.getElementById('searchInput');
            filter = input.value.toUpperCase();
            table = document.getElementById('logTable');
            tr = table.getElementsByTagName('tr');
            for (i = 1; i < tr.length; i++) {
                tr[i].style.display = 'none';
                td = tr[i].getElementsByTagName('td');
                for (j = 0; j < td.length; j++) {
                    if (td[j]) {
                        txtValue = td[j].textContent || td[j].innerText;
                        if (txtValue.toUpperCase().indexOf(filter) > -1) {
                            tr[i].style.display = '';
                            break;
                        }
                    }
                }
            }
        }
    </script>
</head>
<body>
    <div class='header'>
        <img src='https://jornada365.cloud/wp-content/uploads/2024/03/Logotipo-Jornada-365-Home.png' class='logo' alt='Jornada 365 Logo'>
        <h1>Jornada Intune Sync Report</h1>
        <div class='search-bar'>
            <input type='text' id='searchInput' onkeyup='searchTable()' placeholder='Pesquisar...'>
        </div>
    </div>
    <div class='table-container'>
        <table id='logTable'>
            <thead>
                <tr>
                    <th>Hostname</th>
                    <th>Email</th>
                    <th>Licença</th>
                    <th>Status</th>
                    <th>Data/Hora</th>
                </tr>
            </thead>
            <tbody>
"@

    $logEntries = @{}
    foreach ($logEntry in $logContent) {
        if ($logEntry.Mensagem -eq "Sincronização bem-sucedida") {
            if (-not $logEntries[$logEntry.NomeDoDispositivo]) {
                $logEntries[$logEntry.NomeDoDispositivo] = $true
                $htmlContent += "<tr><td>$($logEntry.NomeDoDispositivo)</td><td>$($logEntry.Email)</td><td>$($logEntry.Licenca)</td><td>Sucesso</td><td>$($logEntry.Timestamp)</td></tr>"
            }
        }
    }

    $htmlContent += @"
            </tbody>
        </table>
    </div>
    <div class='chart'>
        <canvas id='syncChart'></canvas>
    </div>
    <script>
        var ctx = document.getElementById('syncChart').getContext('2d');
        var chart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['Sucesso', 'Erro'],
                datasets: [{
                    label: 'Sincronizações',
                    data: [$($logContent.Count), $($errorLogContent.Count)],
                    backgroundColor: ['#4CAF50', '#F44336']
                }]
            },
            options: {
                responsive: true,
                scales: {
                    yAxes: [{
                        ticks: {
                            beginAtZero: true
                        }]
                    }]
                }
            }
        });
    </script>
    <div class='additional-info'>
        <h2>Dispositivos por Plataforma</h2>
        <p>Total de dispositivos Windows 10: $($deviceCounts.Windows10)</p>
        <p>Total de dispositivos Windows 11: $($deviceCounts.Windows11)</p>
        <p>Total de dispositivos Android: $($deviceCounts.Android)</p>
        <p>Total de dispositivos macOS: $($deviceCounts.macOS)</p>
        <p>Total de dispositivos ChromeOS: $($deviceCounts.ChromeOS)</p>
        <p>Total de dispositivos no Autopilot: $($deviceCounts.Autopilot)</p>
        <p>Total de dispositivos em conformidade: $($deviceCounts.Compliant)</p>
        <p>Total de dispositivos fora de conformidade: $($deviceCounts.NonCompliant)</p>
    </div>
    <div class='chart'>
        <canvas id='osChart'></canvas>
    </div>
    <script>
        var osCtx = document.getElementById('osChart').getContext('2d');
        var osChart = new Chart(osCtx, {
            type: 'pie',
            data: {
                labels: ['Windows 10', 'Windows 11', 'Android', 'macOS', 'ChromeOS'],
                datasets: [{
                    label: 'Dispositivos por SO',
                    data: [$($deviceCounts.Windows10), $($deviceCounts.Windows11), $($deviceCounts.Android), $($deviceCounts.macOS), $($deviceCounts.ChromeOS)],
                    backgroundColor: ['#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF']
                }]
            },
            options: {
                responsive: true
            }
        });
    </script>
</body>
</html>
"@

    $outputPath = "$logDirectory\relatorio_$(Get-Date -Format 'yyyyMMdd_HHmm').html"
    $htmlContent | Out-File -FilePath $outputPath -Encoding utf8
    Start-Process $outputPath
}

# Criar diretório de log se não existir
$logDirectory = "C:\J365_Intune"
Create-DirectoryIfNotExists -path $logDirectory

# Verificar e instalar/atualizar os módulos necessários
Install-Update-Modules

# Conectar ao Microsoft Graph
Connect-ToGraph

# Criar a janela principal
$form = New-Object System.Windows.Forms.Form
$form.Text = "Jornada Intune Sync"
$form.Size = New-Object System.Drawing.Size(1000, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# Adicionar a logo
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.ImageLocation = "https://jornada365.cloud/wp-content/uploads/2024/03/Logotipo-Jornada-365-Home.png"
$pictureBox.Size = New-Object System.Drawing.Size(150, 50)
$pictureBox.Location = New-Object System.Drawing.Point(10, 10)
$pictureBox.SizeMode = "Zoom"
$form.Controls.Add($pictureBox)

# Adicionar título
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Jornada Intune Sync"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 24, [System.Drawing.FontStyle]::Bold)
$titleLabel.Location = New-Object System.Drawing.Point(170, 10)
$titleLabel.Size = New-Object System.Drawing.Size(660, 50)
$titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($titleLabel)

# Criar a group box
$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Text = "Device Sync"
$groupBox.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$groupBox.Size = New-Object System.Drawing.Size(960, 300)
$groupBox.Location = New-Object System.Drawing.Point(10, 80)
$groupBox.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($groupBox)

# Adicionar CheckBoxes para as opções
$options = @(
    "Sincronizar todos os dispositivos Windows (Físicos)",
    "Sincronizar todos os dispositivos Android",
    "Sincronizar todos os dispositivos macOS",
    "Sincronizar todos os dispositivos ChromeOS",
    "Reparar dispositivos com problemas de sincronismo",
    "Forçar sincronização de todos os dispositivos",
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

# Adicionar label para mostrar dispositivos sendo sincronizados
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(170, 400)
$statusLabel.Size = New-Object System.Drawing.Size(660, 30)
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$statusLabel.BackColor = [System.Drawing.Color]::White
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($statusLabel)

# Adicionar barra de progresso estilosa e fina
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(70, 440)
$progressBar.Size = New-Object System.Drawing.Size(860, 10)  # Barra de progresso mais fina
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$progressBar.ForeColor = [System.Drawing.Color]::RoyalBlue
$form.Controls.Add($progressBar)

# Adicionar botões pretos e estilosos
$executeButton = New-Object System.Windows.Forms.Button
$executeButton.Text = "Executar"
$executeButton.Location = New-Object System.Drawing.Point(320, 480)
$executeButton.Size = New-Object System.Drawing.Size(150, 40)
$executeButton.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$executeButton.BackColor = [System.Drawing.Color]::Black
$executeButton.ForeColor = [System.Drawing.Color]::White
$executeButton.FlatStyle = "Flat"
$form.Controls.Add($executeButton)

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "Fechar"
$closeButton.Location = New-Object System.Drawing.Point(490, 480)
$closeButton.Size = New-Object System.Drawing.Size(150, 40)
$closeButton.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$closeButton.BackColor = [System.Drawing.Color]::Black
$closeButton.ForeColor = [System.Drawing.Color]::White
$closeButton.FlatStyle = "Flat"
$form.Controls.Add($closeButton)

# Função para executar a opção selecionada
$executeButton.Add_Click({
    $selectedOptions = $checkBoxes | Where-Object { $_.Checked }
    if (-not $selectedOptions) {
        [System.Windows.Forms.MessageBox]::Show("Por favor, selecione uma opção.")
        return
    }
    $logPath = "$logDirectory\sync_log_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    $errorLogPath = "$logDirectory\error_log_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    foreach ($selectedOption in $selectedOptions) {
        switch ($selectedOption.Text) {
            "Sincronizar todos os dispositivos Windows (Físicos)" {
                Write-Host "Sincronizando todos os dispositivos Windows (Físicos)..."
                Sync-Devices -filter "operatingSystem eq 'Windows'" -logPath $logPath -errorLogPath $errorLogPath -deviceType "Windows"
            }
            "Sincronizar todos os dispositivos Android" {
                Write-Host "Sincronizando todos os dispositivos Android..."
                Sync-Devices -filter "operatingSystem eq 'Android'" -logPath $logPath -errorLogPath $errorLogPath -deviceType "Android"
            }
            "Sincronizar todos os dispositivos macOS" {
                Write-Host "Sincronizando todos os dispositivos macOS..."
                Sync-Devices -filter "operatingSystem eq 'macOS'" -logPath $logPath -errorLogPath $errorLogPath -deviceType "macOS"
            }
            "Sincronizar todos os dispositivos ChromeOS" {
                Write-Host "Sincronizando todos os dispositivos ChromeOS..."
                Sync-Devices -filter "operatingSystem eq 'ChromeOS'" -logPath $logPath -errorLogPath $errorLogPath -deviceType "ChromeOS"
            }
            "Reparar dispositivos com problemas de sincronismo" {
                Write-Host "Reparando dispositivos com problemas de sincronismo..."
                Resync-NoncompliantOrOutdatedDevices -logPath $logPath -errorLogPath $errorLogPath
            }
            "Forçar sincronização de todos os dispositivos" {
                Write-Host "Forçando sincronização de todos os dispositivos..."
                Force-Sync-AllDevices -logPath $logPath -errorLogPath $errorLogPath
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
    Generate-HTMLReport -title "Jornada Intune Sync Report" -logPath $logPath -errorLogPath $errorLogPath
})

# Função para fechar o formulário
$closeButton.Add_Click({
    $form.Close()
    Disconnect-FromGraph
})

# Mostrar o formulário
[void]$form.ShowDialog()

# Desconectar do Microsoft Graph ao fechar o formulário
Disconnect-FromGraph
