# Jornada 365 - Intune Sync

## Visão Geral

O script "Jornada 365 - Intune Sync" foi desenvolvido para facilitar a sincronização e gerenciamento de dispositivos no Microsoft Intune. Este script automatiza várias tarefas de administração de dispositivos, como sincronização, reparo de dispositivos com problemas de sincronismo e forçar a sincronização de todos os dispositivos. Ele é especialmente útil para administradores de TI que gerenciam um grande número de dispositivos e precisam de uma maneira eficiente e automatizada para garantir que todos os dispositivos estejam atualizados e em conformidade.

https://github.com/sesantanajr/Itune_Sync/blob/main/intune_sync_menu.png

## Funcionalidades

- **Instalação e Atualização do Módulo Microsoft Graph Intune**: Verifica se o módulo Microsoft Graph Intune está instalado e atualizado. Se não estiver, o módulo é instalado ou atualizado automaticamente.
- **Conexão e Desconexão do Microsoft Graph**: Conecta e desconecta do Microsoft Graph com as permissões necessárias para gerenciar dispositivos no Intune.
- **Sincronização de Dispositivos**: Sincroniza dispositivos baseados em diferentes sistemas operacionais, como Windows, Android, macOS e ChromeOS.
- **Reparo de Dispositivos com Problemas de Sincronismo**: Repara e sincroniza dispositivos que não estão em conformidade ou que não foram sincronizados nas últimas 12 horas.
- **Forçar Sincronização de Todos os Dispositivos**: Força a sincronização de todos os dispositivos, independentemente do seu estado atual.
- **Abertura de URLs**: Abre URLs importantes, como o site Jornada 365 e o portal do Intune, no navegador padrão.
- **Geração de Logs**: Gera logs detalhados de todas as operações realizadas, armazenados em um diretório específico.

## Menu de Opções

### 1. Sincronizar todos os dispositivos Windows (Físicos)

Sincroniza todos os dispositivos com sistema operacional Windows que são físicos.
- **Comando Utilizado**: `Sync-Devices -filter "operatingSystem eq 'Windows'" -logPath $logPath -errorLogPath $errorLogPath`

### 2. Sincronizar todos os dispositivos Android

Sincroniza todos os dispositivos com sistema operacional Android.
- **Comando Utilizado**: `Sync-Devices -filter "operatingSystem eq 'Android'" -logPath $logPath -errorLogPath $errorLogPath`

### 3. Sincronizar todos os dispositivos macOS

Sincroniza todos os dispositivos com sistema operacional macOS.
- **Comando Utilizado**: `Sync-Devices -filter "operatingSystem eq 'macOS'" -logPath $logPath -errorLogPath $errorLogPath`

### 4. Sincronizar todos os dispositivos ChromeOS

Sincroniza todos os dispositivos com sistema operacional ChromeOS.
- **Comando Utilizado**: `Sync-Devices -filter "operatingSystem eq 'ChromeOS'" -logPath $logPath -errorLogPath $errorLogPath`

### 5. Reparar dispositivos com problemas de sincronismo

Repara e sincroniza dispositivos que não estão em conformidade ou que não sincronizaram nas últimas 12 horas.
- **Comando Utilizado**: `Resync-NoncompliantOrOutdatedDevices -logPath $logPath -errorLogPath $errorLogPath`

### 6. Forçar sincronização de todos os dispositivos

Força a sincronização de todos os dispositivos, independentemente do estado.
- **Comando Utilizado**: `Force-Sync-AllDevices -logPath $logPath -errorLogPath $errorLogPath`

### 7. Abrir Jornada 365

Abre o site Jornada 365 no navegador padrão.
- **Comando Utilizado**: `Open-Url -url "https://jornada365.cloud"`

### 8. Abrir Intune Portal

Abre o portal do Intune no navegador padrão.
- **Comando Utilizado**: `Open-Url -url "https://intune.microsoft.com/"`

### 9. Ajuda

Exibe o menu de ajuda detalhada.
- **Comando Utilizado**: `Show-Help`

### 10. Sair e desconectar

Sai do script e desconecta do Microsoft Graph.
- **Comando Utilizado**: `Disconnect-FromGraph`

## Logs

Todos os logs de operações realizadas são salvos no diretório `C:\J365_Intune` com data e hora no nome do arquivo, permitindo fácil rastreamento e auditoria das ações realizadas pelo script.

## Como Usar

1. **Clonar o Repositório**:
   ```sh
   git clone https://github.com/sesantanajr/Itune_Sync.git
   cd Itune_Sync
   ```

2. **Executar o Script**:
   Abra o PowerShell como Administrador e execute o script `Intune_Sync.ps1`:
   ```sh
   .\Intune_Sync.ps1
   ```

3. **Seguir as Instruções na Tela**:
   O script exibirá um menu de opções. Escolha a opção desejada digitando o número correspondente e pressionando Enter.

## Requisitos

- **PowerShell 5.1 ou superior**
- **Módulo Microsoft.Graph.Intune**
- **Permissões apropriadas no Microsoft Intune**

## Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para abrir uma issue ou enviar um pull request com melhorias, correções de bugs ou novas funcionalidades.
Faça parte desta Jornada você também - Jornada 365 - https://jornada365.cloud
