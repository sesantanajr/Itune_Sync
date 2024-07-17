# Funcao para verificar e instalar/atualizar o modulo Microsoft Graph Intune
function Install-Update-GraphModule {
    $moduleName = "Microsoft.Graph.Intune"
    $module = Get-InstalledModule -Name $moduleName -ErrorAction SilentlyContinue

    if ($null -eq $module) {
        Write-Host "Instalando o modulo $moduleName..."
        Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber
    } else {
        $latestModule = Find-Module -Name $moduleName
        if ($module.Version -lt $latestModule.Version) {
            Write-Host "Atualizando o modulo $moduleName..."
            Update-Module -Name $moduleName -Force
        } else {
            Write-Host "O modulo $moduleName ja esta instalado e atualizado."
        }
    }
}

# Funcao para conectar ao Microsoft Graph com autenticacao interativa
function Connect-ToGraph {
    Write-Host "Conectando ao Microsoft Graph..."
    Connect-MgGraph -Scopes "DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementManagedDevices.PrivilegedOperations.All"
}

# Funcao para desconectar do Microsoft Graph
function Disconnect-FromGraph {
    Write-Host "Desconectando do Microsoft Graph..."
    Disconnect-MgGraph
}

# Funcao para verificar e habilitar o servico dmwappushservice
function Ensure-Dmwappushservice {
    $service = Get-Service -Name dmwappushservice -ErrorAction SilentlyContinue
    if ($service -and $service.Status -ne 'Running') {
        Write-Host "Habilitando e iniciando o servico dmwappushservice..."
        Set-Service -Name dmwappushservice -StartupType Automatic
        Start-Service -Name dmwappushservice
    }
}

# Funcao para criar diretorio se nao existir
function Create-DirectoryIfNotExists {
    param (
        [string]$path
    )
    if (-not (Test-Path -Path $path)) {
        New-Item -ItemType Directory -Path $path -Force
    }
}

# Funcao para gerar log
function Log-Action {
    param (
        [string]$logPath,
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $message"
    Add-Content -Path $logPath -Value $logEntry
}

# Funcao para sincronizar dispositivos por sistema operacional
function Sync-Devices {
    param (
        [string]$filter,
        [string]$logPath,
        [string]$errorLogPath
    )
    try {
        $devices = Get-MgDeviceManagementManagedDevice -Filter $filter -ErrorAction Stop
        if ($devices.Count -eq 0) {
            $message = "Nao existem dispositivos configurados para esta plataforma."
            Write-Host $message -ForegroundColor Yellow
            Log-Action -logPath $logPath -message $message
            Write-Host "Pressione Enter para voltar ao menu principal..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        }
        $totalDevices = $devices.Count
        $counter = 0

        foreach ($device in $devices) {
            $counter++
            $percentage = [math]::Round(($counter / $totalDevices) * 100)
            Write-Progress -Activity "Sincronizando dispositivos" -Status "$percentage% completo" -PercentComplete $percentage -CurrentOperation $device.DeviceName
            try {
                Sync-MgDeviceManagementManagedDevice -ManagedDeviceId $device.Id
                $logMessage = "Sincronizando dispositivo: $($device.DeviceName) (ID: $($device.Id))"
                Write-Host $logMessage
                Log-Action -logPath $logPath -message $logMessage
            } catch {
                $errorMessage = "Erro ao sincronizar dispositivo $($device.DeviceName): $_"
                Write-Host $errorMessage -ForegroundColor Red
                Log-Action -logPath $errorLogPath -message $errorMessage
            }
            Start-Sleep -Milliseconds 500 # Pausa para visualizacao gradual
        }
        Write-Host "Sincronizacao concluida. O arquivo de log foi salvo em $logPath"
        Write-Host "Pressione Enter para voltar ao menu principal..."
    } catch {
        $errorMessage = "Erro ao obter dispositivos: $_"
        Write-Host $errorMessage -ForegroundColor Red
        Log-Action -logPath $errorLogPath -message $errorMessage
        Write-Host "Pressione Enter para continuar..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Funcao para sincronizar dispositivos nao conformes e desatualizados
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
            $message = "Nenhum dispositivo nao conforme ou desatualizado encontrado para sincronizar."
            Write-Host $message -ForegroundColor Yellow
            Log-Action -logPath $logPath -message $message
            Write-Host "Pressione Enter para voltar ao menu principal..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        }
        
        $totalDevices = $devices.Count
        $counter = 0

        foreach ($device in $devices) {
            $counter++
            $percentage = [math]::Round(($counter / $totalDevices) * 100)
            Write-Progress -Activity "Reparando dispositivos" -Status "$percentage% completo" -PercentComplete $percentage -CurrentOperation $device.DeviceName
            Ensure-Dmwappushservice
            try {
                # Correção do cmdlet
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($device.Id)/syncDevice"
                $logMessage = "Reparando e sincronizando dispositivo: $($device.DeviceName) (ID: $($device.Id))"
                Write-Host $logMessage
                Log-Action -logPath $logPath -message $logMessage
            } catch {
                $errorMessage = "Erro ao reparar dispositivo $($device.DeviceName): $_"
                Write-Host $errorMessage -ForegroundColor Red
                Log-Action -logPath $errorLogPath -message $errorMessage
            }
            Start-Sleep -Milliseconds 500 # Pausa para visualizacao gradual
        }
        Write-Host "Reparo concluido. O arquivo de log foi salvo em $logPath"
        Write-Host "Pressione Enter para voltar ao menu principal..."
    } catch {
        $errorMessage = "Erro ao obter dispositivos: $_"
        Write-Host $errorMessage -ForegroundColor Red
        Log-Action -logPath $errorLogPath -message $errorMessage
        Write-Host "Pressione Enter para continuar..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Funcao para forcar sincronizacao de todos os dispositivos
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
            Write-Host "Pressione Enter para voltar ao menu principal..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        }
        $totalDevices = $devices.Count
        $counter = 0

        foreach ($device in $devices) {
            $counter++
            $percentage = [math]::Round(($counter / $totalDevices) * 100)
            Write-Progress -Activity "Forcando sincronizacao de todos os dispositivos" -Status "$percentage% completo" -PercentComplete $percentage -CurrentOperation $device.DeviceName
            try {
                Sync-MgDeviceManagementManagedDevice -ManagedDeviceId $device.Id
                $logMessage = "Forcando sincronizacao do dispositivo: $($device.DeviceName) (ID: $($device.Id))"
                Write-Host $logMessage
                Log-Action -logPath $logPath -message $logMessage
            } catch {
                $errorMessage = "Erro ao forcar sincronizacao do dispositivo $($device.DeviceName): $_"
                Write-Host $errorMessage -ForegroundColor Red
                Log-Action -logPath $errorLogPath -message $errorMessage
            }
            Start-Sleep -Milliseconds 500 # Pausa para visualizacao gradual
        }
        Write-Host "Forcamento de sincronizacao concluido. O arquivo de log foi salvo em $logPath"
        Write-Host "Pressione Enter para voltar ao menu principal..."
    } catch {
        $errorMessage = "Erro ao obter dispositivos: $_"
        Write-Host $errorMessage -ForegroundColor Red
        Log-Action -logPath $errorLogPath -message $errorMessage
        Write-Host "Pressione Enter para continuar..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Funcao para abrir URLs no navegador padrao
function Open-Url {
    param (
        [string]$url
    )
    Start-Process $url
}

# Funcao para exportar ajuda para arquivo txt
function Export-Help {
    param (
        [string]$helpText
    )
    $helpFileName = "C:\J365_Intune\Jornada365_Ajuda.txt"
    $counter = 1
    while (Test-Path $helpFileName) {
        $helpFileName = "C:\J365_Intune\Jornada365_Ajuda_$counter.txt"
        $counter++
    }
    $helpText | Out-File -FilePath $helpFileName
    Write-Host "Ajuda exportada para o arquivo $helpFileName"
}

# Funcao de ajuda detalhada
function Show-Help {
    Clear-Host
    $helpText = @"
============================================
                  Ajuda
============================================
Este script ajuda a gerenciar e sincronizar dispositivos no Intune.

Opcoes disponiveis:
1. Sincronizar todos os dispositivos Windows (Fisicos):
   Sincroniza todos os dispositivos com sistema operacional Windows que sao fisicos.
   Comando utilizado: Sync-Devices -filter "operatingSystem eq 'Windows'" -logPath \$logPath -errorLogPath \$errorLogPath

2. Sincronizar todos os dispositivos Android:
   Sincroniza todos os dispositivos com sistema operacional Android.
   Comando utilizado: Sync-Devices -filter "operatingSystem eq 'Android'" -logPath \$logPath -errorLogPath \$errorLogPath

3. Sincronizar todos os dispositivos macOS:
   Sincroniza todos os dispositivos com sistema operacional macOS.
   Comando utilizado: Sync-Devices -filter "operatingSystem eq 'macOS'" -logPath \$logPath -errorLogPath \$errorLogPath

4. Sincronizar todos os dispositivos ChromeOS:
   Sincroniza todos os dispositivos com sistema operacional ChromeOS.
   Comando utilizado: Sync-Devices -filter "operatingSystem eq 'ChromeOS'" -logPath \$logPath -errorLogPath \$errorLogPath

5. Reparar dispositivos com problemas de sincronismo:
   Repara e sincroniza dispositivos que nao estao em conformidade ou que nao sincronizaram nas ultimas 12 horas.
   Comando utilizado: Resync-NoncompliantOrOutdatedDevices -logPath \$logPath -errorLogPath \$errorLogPath

6. Forcar sincronizacao de todos os dispositivos:
   Forca a sincronizacao de todos os dispositivos, independentemente do estado.
   Comando utilizado: Force-Sync-AllDevices -logPath \$logPath -errorLogPath \$errorLogPath

7. Abrir Jornada 365:
   Abre o site Jornada 365 no navegador padrao.
   Comando utilizado: Open-Url -url "https://jornada365.cloud"

8. Abrir Intune Portal:
   Abre o portal do Intune no navegador padrao.
   Comando utilizado: Open-Url -url "https://intune.microsoft.com/"

9. Ajuda:
   Exibe este menu de ajuda detalhada.
   Comando utilizado: Show-Help

10. Sair e desconectar:
   Sai do script e desconecta do Microsoft Graph.
   Comando utilizado: Disconnect-FromGraph

Os logs de todas as acoes realizadas sao salvos em C:\J365_Intune com data e hora no nome do arquivo.
============================================
"@
    Write-Host $helpText
    Write-Host "Deseja exportar essa ajuda para um arquivo txt? (S/N)"
    $response = Read-Host
    if ($response -eq "S") {
        Export-Help -helpText $helpText
    }
    Write-Host "Pressione Enter para voltar ao menu principal..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Tela de boas-vindas
function Show-WelcomeScreen {
    Clear-Host
    Write-Host "============================================"
    Write-Host "          Jornada 365 - Intune Sync          "
    Write-Host "============================================"
    Write-Host ""
    Write-Host "Bem-vindo ao script para sincronizar dispositivos Intune."
    Write-Host "Foi criado um diretorio em C:\J365_Intune para armazenar logs detalhados de todas as operacoes realizadas."
    Write-Host ""
    Write-Host "Acesse o site: https://jornada365.cloud"
    Write-Host "Learn Microsoft: https://learn.microsoft.com/pt-br/mem/intune/"
    Write-Host ""
    Write-Host "============================================"
    Write-Host ""
    Write-Host "Pressione qualquer tecla para continuar..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Menu de opcoes
function Show-Menu {
    Clear-Host
    Write-Host "============================================"
    Write-Host "               Menu de Opcoes               "
    Write-Host "============================================"
    Write-Host "1. Sincronizar todos os dispositivos Windows (Fisicos)"
    Write-Host "2. Sincronizar todos os dispositivos Android"
    Write-Host "3. Sincronizar todos os dispositivos macOS"
    Write-Host "4. Sincronizar todos os dispositivos ChromeOS"
    Write-Host "5. Reparar dispositivos com problemas de sincronismo"
    Write-Host "6. Forcar sincronizacao de todos os dispositivos"
    Write-Host "7. Abrir Jornada 365"
    Write-Host "8. Abrir Intune Portal"
    Write-Host "9. Ajuda"
    Write-Host "10. Sair e desconectar"
    Write-Host "============================================"
    $choice = Read-Host "Digite o numero da sua escolha"
    return $choice
}

# Criar diretorio de log se nao existir
$logDirectory = "C:\J365_Intune"
Create-DirectoryIfNotExists -path $logDirectory

# Caminho do arquivo de log
$logPath = "$logDirectory\sync_log_$(Get-Date -Format 'yyyyMMdd_HHmm').txt"
$errorLogPath = "$logDirectory\error_log_$(Get-Date -Format 'yyyyMMdd_HHmm').txt"

# Verificar e instalar/atualizar o modulo Microsoft Graph Intune
Install-Update-GraphModule

# Conectar ao Microsoft Graph
Connect-ToGraph

# Exibir tela de boas-vindas
Show-WelcomeScreen

# Loop principal
do {
    $choice = Show-Menu
    switch ($choice) {
        1 {
            Write-Host "Sincronizando todos os dispositivos Windows (Fisicos)..."
            Sync-Devices -filter "operatingSystem eq 'Windows'" -logPath $logPath -errorLogPath $errorLogPath
        }
        2 {
            Write-Host "Sincronizando todos os dispositivos Android..."
            Sync-Devices -filter "operatingSystem eq 'Android'" -logPath $logPath -errorLogPath $errorLogPath
        }
        3 {
            Write-Host "Sincronizando todos os dispositivos macOS..."
            Sync-Devices -filter "operatingSystem eq 'macOS'" -logPath $logPath -errorLogPath $errorLogPath
        }
        4 {
            Write-Host "Sincronizando todos os dispositivos ChromeOS..."
            Sync-Devices -filter "operatingSystem eq 'ChromeOS'" -logPath $logPath -errorLogPath $errorLogPath
        }
        5 {
            Write-Host "Reparando dispositivos com problemas de sincronismo..."
            Resync-NoncompliantOrOutdatedDevices -logPath $logPath -errorLogPath $errorLogPath
        }
        6 {
            Write-Host "Forcando sincronizacao de todos os dispositivos..."
            Force-Sync-AllDevices -logPath $logPath -errorLogPath $errorLogPath
        }
        7 {
            Write-Host "Abrindo Jornada 365..."
            Open-Url -url "https://jornada365.cloud"
        }
        8 {
            Write-Host "Abrindo Intune Portal..."
            Open-Url -url "https://intune.microsoft.com/"
        }
        9 {
            Show-Help
        }
        10 {
            Write-Host "Saindo e desconectando..."
            Disconnect-FromGraph
            exit
        }
        default {
            Write-Host "Opcao invalida. Tente novamente."
        }
    }
} while ($choice -ne 10)

Write-Host "Sincronizacao concluida. O arquivo de log foi salvo em $logPath"
Disconnect-FromGraph
