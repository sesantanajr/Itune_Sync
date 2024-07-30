# Caminhos e configurações
$CompanyLogo = "https://jornada365.cloud/wp-content/uploads/2024/03/Logotipo-Jornada-365-Home.png"
$ReportSavePath = "C:\Relatorio_J365\"
$ErrorLog = Join-Path -Path $ReportSavePath -ChildPath "ErrorLog.txt"
$DateTimeNow = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

# Criar caminho para salvar o relatório se não existir
if (-not (Test-Path -Path $ReportSavePath)) {
    New-Item -ItemType Directory -Path $ReportSavePath -Force
}

# Função para garantir que os módulos necessários estão instalados
function Ensure-Module {
    param (
        [string]$ModuleName,
        [string]$ModuleVersion = 'latest'
    )
    try {
        if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
            Write-Host "Instalando módulo $ModuleName..."
            Install-Module -Name $ModuleName -Force -Scope CurrentUser
        } else {
            $currentVersion = (Get-Module -ListAvailable -Name $ModuleName).Version
            if ($ModuleVersion -ne 'latest' -and $currentVersion -lt [version]$ModuleVersion) {
                Write-Host "Atualizando módulo $ModuleName para a versão $ModuleVersion..."
                Install-Module -Name $ModuleName -Force -Scope CurrentUser -RequiredVersion $ModuleVersion
            } else {
                Write-Host "$ModuleName já está instalado e atualizado."
            }
        }
    } catch {
        Write-Error "Erro ao instalar módulo $($ModuleName): $_"
        Add-Content -Path $ErrorLog -Value "Erro ao instalar módulo $($ModuleName): $_"
    }
}

Ensure-Module -ModuleName "Microsoft.Graph"
Ensure-Module -ModuleName "PSWriteHTML"

# Autenticação interativa
function Connect-ToMicrosoftGraph {
    try {
        Write-Host "Solicitando credenciais para conectar ao Microsoft Graph" -ForegroundColor Yellow
        Connect-MgGraph -Scopes `
            "User.Read.All", `
            "Group.Read.All", `
            "Reports.Read.All", `
            "Sites.Read.All", `
            "DeviceManagementManagedDevices.Read.All", `
            "SecurityEvents.Read.All", `
            "Files.Read.All", `
            "Directory.ReadWrite.All", `
            "RoleManagement.ReadWrite.Directory", `
            "Domain.ReadWrite.All", `
            "ServiceHealth.Read.All", `
            "Mail.ReadWrite", `
            "Mail.ReadWrite.Shared", `
            "Mail.Send", `
            "Mail.Send.Shared"
        Write-Host "Conectado ao Microsoft Graph com sucesso." -ForegroundColor Green
    } catch {
        $errorMessage = "Erro ao conectar ao Microsoft Graph: $_"
        Write-Error $errorMessage
        Add-Content -Path $ErrorLog -Value $errorMessage
    }
}

Connect-ToMicrosoftGraph

# Função para obter nomes amigáveis das licenças
function Get-FriendlySkuNames {
    $skuUrl = "https://raw.githubusercontent.com/MicrosoftDocs/entra-docs/main/docs/identity/users/licensing-service-plan-reference.md"
    $skuNames = @{}
    try {
        Write-Host "Obtendo nomes amigáveis de SKU..."
        $content = Invoke-WebRequest -Uri $skuUrl -UseBasicParsing
        $lines = $content.Content -split "`n"
        foreach ($line in $lines) {
            if ($line -match "\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|") {
                $skuNames[$matches[3].Trim()] = $matches[1].Trim()
            }
        }
    } catch {
        $errorMessage = "Erro ao obter nomes amigáveis de SKU: $($_.Exception.Message)"
        Write-Error $errorMessage
        Add-Content -Path $ErrorLog -Value $errorMessage
    }
    Write-Host "Nomes amigáveis de SKU obtidos com sucesso."
    return $skuNames
}

# Função para listar licenças disponíveis
function Get-AvailableSkus {
    [array]$Skus = Get-MgSubscribedSku
    $SkuList = [System.Collections.Generic.List[Object]]::new()
    $friendlySkuNames = Get-FriendlySkuNames

    foreach ($Sku in $Skus) {
        $SkuAvailable = ($Sku.PrepaidUnits.Enabled - $Sku.ConsumedUnits)
        $SkuName = if ($friendlySkuNames.ContainsKey($Sku.SkuId)) { $friendlySkuNames[$Sku.SkuId] } else { $Sku.SkuPartNumber }
        $ReportLine = [PSCustomObject]@{
            SkuId         = $Sku.SkuId
            NomeDoProduto = $SkuName
            Consumido     = $Sku.ConsumedUnits
            Pago          = $Sku.PrepaidUnits.Enabled
            Disponivel    = $SkuAvailable
        }
        $SkuList.Add($ReportLine)
    }

    return $SkuList
}

# Função para obter o status do MFA
function Get-MFAStatus {
    param (
        [string]$UserPrincipalName
    )
    try {
        $mfa = Get-MgUserAuthenticationMethod -UserId $UserPrincipalName
        if ($mfa -ne $null) {
            return "Enabled"
        } else {
            return "Disabled"
        }
    } catch {
        return "Error"
    }
}

# Função para obter uso da caixa de correio
function Get-MailboxUsage {
    param ([string]$UPN)
    try {
        $mailFolder = Get-MgUserMailFolder -UserId $UPN -MailFolderId 'Inbox'
        $usage = $mailFolder.TotalItemSize
        if ($usage -match '\d+') {
            $usageInBytes = [int64]$usage
            $usageInGB = [math]::Round($usageInBytes / 1GB, 2)
            if ($usageInGB -ge 1) {
                return "$usageInGB GB"
            } else {
                $usageInMB = [math]::Round($usageInBytes / 1MB, 2)
                return "$usageInMB MB"
            }
        }
        return "Dados de uso não disponíveis"
    } catch {
        if ($_.Exception.ErrorCode -eq "ErrorItemNotFound" -or $_.Exception.ErrorCode -eq "MailboxNotEnabledForRESTAPI") {
            $message = "Caixa de correio não encontrada ou não habilitada para o usuário ${UPN}"
            Write-Warning $message
            Add-Content -Path $ErrorLog -Value $message
            return "Dados de uso não disponíveis"
        } else {
            $message = "Não foi possível obter o uso da caixa de correio para ${UPN}: $_"
            Write-Warning $message
            Add-Content -Path $ErrorLog -Value $message
            return "Dados de uso não disponíveis"
        }
    }
}

# Funções para obter dados de vários serviços
function Get-EntraIDData {
    try {
        $entraIDUsers = Get-MgUser -All
        Write-Host "Dados do Microsoft Entra ID extraídos: $($entraIDUsers.Count)" -ForegroundColor Green
        return $entraIDUsers
    } catch {
        $errorMessage = "Erro ao obter dados do Microsoft Entra ID: $_"
        Write-Error $errorMessage
        Add-Content -Path $ErrorLog -Value $errorMessage
        return $null
    }
}

function Get-SharePointData {
    try {
        $sharePointSites = Get-MgSite -All
        Write-Host "Dados do SharePoint extraídos: $($sharePointSites.Count)" -ForegroundColor Green
        return $sharePointSites
    } catch {
        $errorMessage = "Erro ao obter dados do SharePoint: $_"
        Write-Error $errorMessage
        Add-Content -Path $ErrorLog -Value $errorMessage
        return $null
    }
}

function Get-ExchangeData {
    try {
        $mailboxes = @()
        $users = Get-EntraIDData
        foreach ($user in $users) {
            try {
                $mailboxUsage = Get-MailboxUsage -UPN $user.UserPrincipalName
                $mailboxes += [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName = $user.DisplayName
                    MailboxUsage = $mailboxUsage
                }
            } catch {
                $message = "Erro ao obter a caixa de correio para o usuário $($user.UserPrincipalName)"
                Write-Warning $message
                Add-Content -Path $ErrorLog -Value $message
            }
        }
        Write-Host "Dados do Exchange extraídos: $($mailboxes.Count)" -ForegroundColor Green
        return $mailboxes
    } catch {
        $errorMessage = "Erro ao obter dados do Exchange: $_"
        Write-Error $errorMessage
        Add-Content -Path $ErrorLog -Value $errorMessage
        return $null
    }
}

function Get-PurviewData {
    try {
        $purviewData = Get-MgCompliancePolicy -All
        Write-Host "Dados do Purview extraídos: $($purviewData.Count)" -ForegroundColor Green
        return $purviewData
    } catch {
        $errorMessage = "Erro ao obter dados do Purview: $_"
        Write-Error $errorMessage
        Add-Content -Path $ErrorLog -Value $errorMessage
        return $null
    }
}

function Get-DefenderData {
    try {
        $defenderReports = Get-MgSecurityAlert -All
        Write-Host "Dados do Defender extraídos: $($defenderReports.Count)" -ForegroundColor Green
        return $defenderReports
    } catch {
        $errorMessage = "Erro ao obter dados do Defender: $_"
        Write-Error $errorMessage
        Add-Content -Path $ErrorLog -Value $errorMessage
        return $null
    }
}

function Get-AdminCenterData {
    try {
        $adminCenterReports = Get-MgServiceAnnouncementIssue -All
        Write-Host "Dados do Microsoft 365 Admin Center extraídos: $($adminCenterReports.Count)" -ForegroundColor Green
        return $adminCenterReports
    } catch {
        $errorMessage = "Erro ao obter dados do Microsoft 365 Admin Center: $_"
        Write-Error $errorMessage
        Add-Content -Path $ErrorLog -Value $errorMessage
        return $null
    }
}

function Get-HealthData {
    try {
        $healthData = Get-MgServiceHealth -All
        Write-Host "Dados do Microsoft Health extraídos: $($healthData.Count)" -ForegroundColor Green
        return $healthData
    } catch {
        $errorMessage = "Erro ao obter dados do Microsoft Health: $_"
        Write-Error $errorMessage
        Add-Content -Path $ErrorLog -Value $errorMessage
        return $null
    }
}

function Get-OneDriveData {
    try {
        $oneDriveUsers = @()
        $users = Get-EntraIDData
        foreach ($user in $users) {
            try {
                $drive = Get-MgUserDrive -UserId $user.Id
                if ($null -ne $drive) {
                    $oneDriveUsers += [PSCustomObject]@{
                        Tipo = "OneDrive"
                        Nome = $user.DisplayName
                        Uso = $drive.Quota.Used
                        Disponivel = $drive.Quota.Remaining
                        Total = $drive.Quota.Total
                        LastModified = $drive.LastModifiedDateTime
                    }
                }
            } catch {
                $message = "OneDrive não encontrado para o usuário $($user.UserPrincipalName)"
                Write-Warning $message
                Add-Content -Path $ErrorLog -Value $message
            }
        }
        Write-Host "Dados do OneDrive extraídos: $($oneDriveUsers.Count)" -ForegroundColor Green
        return $oneDriveUsers
    } catch {
        $errorMessage = "Erro ao obter dados do OneDrive: $_"
        Write-Error $errorMessage
        Add-Content -Path $ErrorLog -Value $errorMessage
        return $null
    }
}

# Extração de dados
function Extract-Data {
    try {
        Write-Host "Extraindo dados..." -ForegroundColor Green
        $Users = Get-EntraIDData
        $Groups = Get-MgGroup -All
        $Licenses = Get-AvailableSkus
        $Sites = Get-SharePointData
        $Mailboxes = Get-ExchangeData
        $IntuneDevices = Get-MgDeviceManagementManagedDevice -All
        $DefenderReports = Get-DefenderData
        $Domains = Get-MgDomain -All
        $ServiceHealthOverview = Get-MgServiceAnnouncementHealthOverview
        $ServiceHealthIssues = Get-AdminCenterData
        $HealthData = Get-HealthData
        $OneDriveUsers = Get-OneDriveData
        $PurviewData = Get-PurviewData

        Write-Host "Dados extraídos com sucesso." -ForegroundColor Green
        return @{
            Users = $Users
            Groups = $Groups
            Licenses = $Licenses
            Sites = $Sites
            Mailboxes = $Mailboxes
            IntuneDevices = $IntuneDevices
            DefenderReports = $DefenderReports
            Domains = $Domains
            ServiceHealthOverview = $ServiceHealthOverview
            ServiceHealthIssues = $ServiceHealthIssues
            HealthData = $HealthData
            OneDriveUsers = $OneDriveUsers
            PurviewData = $PurviewData
        }
    } catch {
        $errorMessage = "Erro ao extrair dados do Microsoft Graph: $_"
        Write-Error $errorMessage
        Add-Content -Path $ErrorLog -Value $errorMessage
        return $null
    }
}

$data = Extract-Data
if ($data -eq $null) {
    Write-Host "Falha na extração de dados. Saindo do script." -ForegroundColor Red
    exit
}

# Função para combinar dados em uma única tabela
function Combine-Data {
    $combinedData = @()

    # Adicionando usuários licenciados e não licenciados
    $data.Users | ForEach-Object {
        $licenses = ($_.AssignedLicenses | ForEach-Object { 
            if ((Get-FriendlySkuNames).ContainsKey($_.SkuId)) {
                (Get-FriendlySkuNames)[$_.SkuId]
            } else {
                "Desconhecida"
            }
        }) -join ' | '
        if ($licenses -ne '') {
            $mailboxUsage = Get-MailboxUsage -UPN $_.UserPrincipalName
            $mfaStatus = Get-MFAStatus -UserPrincipalName $_.UserPrincipalName
            $combinedData += [PSCustomObject]@{
                Tipo = "Mailbox"
                NomeCompleto = $_.DisplayName
                Email = $_.UserPrincipalName
                MailboxUsage = $mailboxUsage
                Licenca = $licenses
                MFA = $mfaStatus
            }
        } else {
            $combinedData += [PSCustomObject]@{
                Tipo = "Unlicensed"
                NomeCompleto = $_.DisplayName
                UPN = $_.UserPrincipalName
                DataRemovida = "Desconhecida"
            }
        }
    }

    # Adicionando grupos
    $data.Groups | ForEach-Object {
        $memberCount = (Get-MgGroupMember -GroupId $_.Id).Count
        $combinedData += [PSCustomObject]@{
            Tipo = "Group"
            NomeDoGrupo = $_.DisplayName
            TipoDeGrupo = if ($_.GroupTypes -contains "Unified") {"MS 365 Group"} elseif ($_.MailEnabled) {"Mail Group"} elseif ($_.SecurityEnabled) {"Security Group"} else {"Other"}
            Membros = $memberCount
        }
    }

    # Adicionando licenças
    $data.Licenses | ForEach-Object {
        $combinedData += [PSCustomObject]@{
            Tipo = "License"
            NomeDoProduto = $_.NomeDoProduto
            QuantidadeTotal = $_.Pago
            QuantidadeEmUso = $_.Consumido
            LicencasDisponiveis = $_.Disponivel
        }
    }

    # Adicionando sites
    $data.Sites | ForEach-Object {
        $combinedData += [PSCustomObject]@{
            Tipo = "Site"
            NomeDoSite = $_.DisplayName
            TipoDeSite = "Desconhecido" # Substituir pelo tipo de site correto
            AdministradoresDoSite = "Desconhecido" # Substituir pelos administradores do site
            Membros = "Desconhecido" # Substituir pela quantidade de membros
            StorageUsed = $_.Quota.Used
            StorageTotal = $_.Quota.Total
        }
    }

    # Adicionando dispositivos Intune
    $data.IntuneDevices | ForEach-Object {
        $combinedData += [PSCustomObject]@{
            Tipo = "IntuneDevice"
            NomeDoDispositivo = $_.DeviceName
            NomeDoUsuario = $_.UserDisplayName
            TipoDeDispositivo = $_.ManagementType
            MicrosoftEntra = $_.EnrollmentType
        }
    }

    # Adicionando relatórios do Defender
    $data.DefenderReports | ForEach-Object {
        $combinedData += [PSCustomObject]@{
            Tipo = "DefenderReport"
            DisplayName = $_.Name
            Details = $_
        }
    }

    # Adicionando usuários do OneDrive
    $data.OneDriveUsers | ForEach-Object {
        $combinedData += [PSCustomObject]@{
            Tipo = "OneDrive"
            Nome = $_.Nome
            Uso = $_.Uso
            Disponivel = $_.Disponivel
            Total = $_.Total
        }
    }

    # Adicionando detalhes de domínios
    $data.Domains | ForEach-Object {
        $combinedData += [PSCustomObject]@{
            Tipo = "Domain"
            DomainName = $_.Id
            VerificationStatus = $_.VerificationStatus
            Default = $_.IsDefault
            DKIMActive = $_.IsAdminManaged
        }
    }

    # Adicionando admins
    $data.Admins | ForEach-Object {
        $combinedData += $_
    }

    # Adicionando visão geral da saúde dos serviços
    $data.ServiceHealthIssues | ForEach-Object {
        $combinedData += $_
    }

    return $combinedData
}

$combinedData = Combine-Data

# Verificando se os dados foram recuperados
Write-Host "Usuários recuperados: $($data.Users.Count)"
Write-Host "Grupos recuperados: $($data.Groups.Count)"
Write-Host "Licenças recuperadas: $($data.Licenses.Count)"
Write-Host "Sites recuperados: $($data.Sites.Count)"
Write-Host "Dispositivos Intune recuperados: $($data.IntuneDevices.Count)"
Write-Host "Relatórios do Defender recuperados: $($data.DefenderReports.Count)"
Write-Host "Usuários do OneDrive recuperados: $($data.OneDriveUsers.Count)"
Write-Host "Domínios recuperados: $($data.Domains.Count)"
Write-Host "Admins recuperados: $($data.Admins.Count)"
Write-Host "Problemas de saúde dos serviços recuperados: $($data.ServiceHealthIssues.Count)"

# Exportando dados para CSV
try {
    $combinedData | Export-Csv -Path "$ReportSavePath\CombinedData.csv" -NoTypeInformation -Force
    Write-Host "Dados exportados para CSV com sucesso." -ForegroundColor Green
} catch {
    $errorMessage = "Erro ao exportar dados para CSV: $_"
    Write-Error $errorMessage
    Add-Content -Path $ErrorLog -Value $errorMessage
}

# Gerar relatório HTML

# CSS para estilização
$css = @"
<style>
    body {
        font-family: Arial, sans-serif;
        margin: 0;
        padding: 0;
        background-color: #f4f4f4;
    }
    .container {
        width: 90%;
        margin: auto;
        overflow: hidden;
    }
    header {
        background: #333;
        color: #fff;
        padding: 10px 0;
        border-bottom: #0078D4 3px solid;
        text-align: center;
    }
    header #branding {
        display: inline-block;
        vertical-align: middle;
    }
    header #branding img {
        width: 150px;
    }
    header #title {
        display: inline-block;
        font-size: 24px;
        font-weight: bold;
        vertical-align: middle;
        margin-left: 20px;
    }
    header #datetime {
        display: inline-block;
        font-size: 14px;
        vertical-align: middle;
        float: right;
    }
    .navbar {
        display: flex;
        justify-content: center;
        background-color: #0078D4;
        flex-wrap: wrap;
        position: relative;
    }
    .navbar a {
        color: white;
        padding: 10px 15px;
        text-decoration: none;
        text-align: center;
    }
    .navbar a:hover {
        background-color: #005bb5;
    }
    .tab-content {
        display: none;
        padding: 20px;
    }
    .tab-content.active {
        display: block;
    }
    .content-table {
        width: 100%;
        border-collapse: collapse;
    }
    .content-table th, .content-table td {
        border: 1px solid #ddd;
        padding: 8px;
    }
    .content-table th {
        background-color: #0078D4;
        color: #fff;
        text-align: left;
    }
    .content-table td {
        background-color: #f9f9f9;
    }
    .navbar #datetime {
        position: absolute;
        right: 0;
        padding: 10px 15px;
        color: white;
    }
</style>
<script>
    function openTab(tabName) {
        var i;
        var x = document.getElementsByClassName("tab-content");
        for (i = 0; i < x.length; i++) {
            x[i].classList.remove('active');
        }
        document.getElementById(tabName).classList.add('active');
    }
</script>
"@

# Função para criar gráficos com PowerShell
function Add-Chart {
    param (
        [string]$Title,
        [string[]]$Labels,
        [int[]]$Values
    )
    $chartHTML = @"
<div class='chart-container'>
    <canvas id='$Title'></canvas>
</div>
<script>
    var ctx = document.getElementById('$Title').getContext('2d');
    var chart = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: ["@($Labels -join '","')"],
            datasets: [{
                label: '$Title',
                data: ["@($Values -join ',')"],
                backgroundColor: [
                    '#0078D4',
                    '#1A73E8',
                    '#4285F4',
                    '#66A1D2',
                    '#80B3F2',
                    '#99C2FF'
                ]
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: {
                    position: 'top',
                },
                title: {
                    display: true,
                    text: '$Title'
                }
            }
        }
    });
</script>
"@
    return $chartHTML
}

# Função para ler CSV e converter para tabela HTML
function Get-HTMLTableFromCSV {
    param (
        [string]$CSVPath,
        [string]$TypeFilter
    )
    try {
        Write-Host "Gerando tabela HTML para ${TypeFilter}..."
        $csvData = Import-Csv -Path $CSVPath | Where-Object { $_.Tipo -eq $TypeFilter }
        if ($csvData.Count -gt 0) {
            $tableHTML = "<table class='content-table'><thead><tr>"
            $tableHTML += ($csvData[0].PSObject.Properties.Name | ForEach-Object { "<th>$_</th>" }) -join ""
            $tableHTML += "</tr></thead><tbody>"
            foreach ($row in $csvData) {
                $tableHTML += "<tr>"
                $tableHTML += ($row.PSObject.Properties.Value | ForEach-Object { "<td>$_</td>" }) -join ""
                $tableHTML += "</tr>"
            }
            $tableHTML += "</tbody></table>"
            return $tableHTML
        } else {
            $message = "Nenhum dado disponível para ${TypeFilter}"
            Write-Warning $message
            Add-Content -Path $ErrorLog -Value $message
            return "<p>$message</p>"
        }
    } catch {
        $errorMessage = "Erro ao gerar tabela HTML para ${TypeFilter}: $_"
        Write-Error $errorMessage
        Add-Content -Path $ErrorLog -Value $errorMessage
        return "<p>Erro ao gerar tabela</p>"
    }
}

# Função para abrir a página HTML
function Get-HTMLOpenPage {
    param (
        [string]$TitleText,
        [string]$LeftLogoString,
        [string]$DateTimeNow
    )
    return @"
<!DOCTYPE html>
<html lang='pt-BR'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>$TitleText</title>
    $css
</head>
<body>
<header>
    <div id='branding'>
        <img src='$CompanyLogo' alt='Company Logo'>
    </div>
    <div id='title'>$TitleText</div>
</header>
<nav class='navbar'>
    <a href='javascript:void(0)' onclick='openTab("Dashboard")'>Dashboard</a>
    <a href='javascript:void(0)' onclick='openTab("Contas")'>Contas</a>
    <a href='javascript:void(0)' onclick='openTab("Groups")'>Groups</a>
    <a href='javascript:void(0)' onclick='openTab("Licencas")'>Licencas</a>
    <a href='javascript:void(0)' onclick='openTab("Health")'>Health</a>
    <a href='javascript:void(0)' onclick='openTab("SharePoint")'>SharePoint</a>
    <a href='javascript:void(0)' onclick='openTab("Intune")'>Intune</a>
    <a href='javascript:void(0)' onclick='openTab("OneDrive")'>OneDrive</a>
    <div id='datetime'>Relatório: $DateTimeNow</div>
</nav>
<div class='container'>
"@
}

# Função para fechar a página HTML
function Get-HTMLClosePage {
    return @"
</div>
</body>
</html>
"@
}

# Função para abrir conteúdo de aba HTML
function Get-HTMLTabContentOpen {
    param (
        [string]$TabName,
        [string]$TabHeading
    )
    return @"
<div id='$TabName' class='tab-content'>
    <h2>$TabHeading</h2>
"@
}

# Função para fechar conteúdo de aba HTML
function Get-HTMLTabContentClose {
    return @"
</div>
"@
}

# Função para abrir conteúdo HTML
function Get-HTMLContentOpen {
    param (
        [string]$HeaderText
    )
    return @"
<div class='content'>
    <h2>$HeaderText</h2>
"@
}

# Função para fechar conteúdo HTML
function Get-HTMLContentClose {
    return @"
</div>
"@
}

# Gerando o relatório HTML
try {
    $rpt = New-Object 'System.Collections.Generic.List[System.Object]'
    $rpt += Get-HTMLOpenPage -TitleText 'Jornada 365 | Microsoft 365 Reports' -LeftLogoString $null -DateTimeNow $DateTimeNow

    # Adicionando CSS
    $rpt += $css

    # Dashboard
    $rpt += Get-HTMLTabContentOpen -TabName 'Dashboard' -TabHeading "Microsoft 365 Dashboard"
    $rpt += Get-HTMLContentOpen -HeaderText "Microsoft 365 Dashboard"

    # Adicionando Informações da Empresa
    $rpt += "<h2>Informações da Empresa</h2>"
    $rpt += Get-HTMLTableFromCSV -CSVPath "$ReportSavePath\CombinedData.csv" -TypeFilter "CompanyInformation"

    # Adicionando informações dos administradores
    $rpt += "<h2>Administradores</h2>"
    $rpt += Get-HTMLTableFromCSV -CSVPath "$ReportSavePath\CombinedData.csv" -TypeFilter "Admin"

    # Adicionando Users MFA
    $rpt += "<h2>Users MFA</h2>"
    $rpt += "<div style='display: flex;'>"
    $rpt += "<div style='width: 50%; padding-right: 10px;'>"
    $rpt += "<h3>Contas com MFA Ativo</h3>"
    $rpt += Get-HTMLTableFromCSV -CSVPath "$ReportSavePath\CombinedData.csv" -TypeFilter "Mailbox" | Where-Object { $_.MFA -eq "Enabled" }
    $rpt += "</div>"
    $rpt += "<div style='width: 50%; padding-left: 10px;'>"
    $rpt += "<h3>Contas sem MFA</h3>"
    $rpt += Get-HTMLTableFromCSV -CSVPath "$ReportSavePath\CombinedData.csv" -TypeFilter "Mailbox" | Where-Object { $_.MFA -ne "Enabled" }
    $rpt += "</div>"
    $rpt += "</div>"

    # Adicionando Contas Licenciadas e Não Licenciadas
    $rpt += "<h2>Contas</h2>"
    $licensedCount = ($data.Users | Where-Object { $_.AssignedLicenses.Count -gt 0 }).Count
    $unlicensedCount = ($data.Users | Where-Object { $_.AssignedLicenses.Count -eq 0 }).Count
    $rpt += "<p>Contas Licenciadas: $licensedCount</p>"
    $rpt += "<p>Contas Não Licenciadas: $unlicensedCount</p>"

    # Adicionando Domínios
    $rpt += "<h2>Domínios</h2>"
    $domainDetails = @()
    foreach ($domain in $data.Domains) {
        $domainDetails += [PSCustomObject]@{
            DomainName = $domain.Id
            VerificationStatus = $domain.VerificationStatus
            Default = $domain.IsDefault
            DKIMActive = $domain.IsAdminManaged
        }
    }
    $rpt += Get-HTMLTableFromCSV -CSVPath "$ReportSavePath\CombinedData.csv" -TypeFilter "Domain"

    $rpt += Get-HTMLContentClose
    $rpt += Get-HTMLTabContentClose

    # Contas
    $rpt += Get-HTMLTabContentOpen -TabName 'Contas' -TabHeading "Microsoft 365 Contas"
    $rpt += Get-HTMLContentOpen -HeaderText "Contas"
    $rpt += Get-HTMLTableFromCSV -CSVPath "$ReportSavePath\CombinedData.csv" -TypeFilter "Mailbox"
    $rpt += Get-HTMLContentClose
    $rpt += Get-HTMLTabContentClose

    # Grupos
    $rpt += Get-HTMLTabContentOpen -TabName 'Groups' -TabHeading "Microsoft 365 Groups"
    $rpt += Get-HTMLContentOpen -HeaderText "Groups"
    $rpt += Get-HTMLTableFromCSV -CSVPath "$ReportSavePath\CombinedData.csv" -TypeFilter "Group"
    $rpt += Get-HTMLContentClose
    $rpt += Get-HTMLTabContentClose

    # Licenças
    $rpt += Get-HTMLTabContentOpen -TabName 'Licencas' -TabHeading "Microsoft 365 Licencas"
    $rpt += Get-HTMLContentOpen -HeaderText "Licencas"
    $rpt += Get-HTMLTableFromCSV -CSVPath "$ReportSavePath\CombinedData.csv" -TypeFilter "License"
    $rpt += Get-HTMLContentClose
    $rpt += Get-HTMLTabContentClose

    # Health
    $rpt += Get-HTMLTabContentOpen -TabName 'Health' -TabHeading "Microsoft 365 Health"
    $rpt += Get-HTMLContentOpen -HeaderText "Health"
    $rpt += Get-HTMLTableFromCSV -CSVPath "$ReportSavePath\CombinedData.csv" -TypeFilter "Health"
    $rpt += Get-HTMLContentClose
    $rpt += Get-HTMLTabContentClose

    # SharePoint
    $rpt += Get-HTMLTabContentOpen -TabName 'SharePoint' -TabHeading "Microsoft 365 SharePoint"
    $rpt += Get-HTMLContentOpen -HeaderText "SharePoint"
    $rpt += Get-HTMLTableFromCSV -CSVPath "$ReportSavePath\CombinedData.csv" -TypeFilter "SharePoint"
    $rpt += Get-HTMLContentClose
    $rpt += Get-HTMLTabContentClose

    # Intune
    $rpt += Get-HTMLTabContentOpen -TabName 'Intune' -TabHeading "Microsoft 365 Intune"
    $rpt += Get-HTMLContentOpen -HeaderText "Intune"
    $rpt += Get-HTMLTableFromCSV -CSVPath "$ReportSavePath\CombinedData.csv" -TypeFilter "IntuneDevice"
    $rpt += Get-HTMLContentClose
    $rpt += Get-HTMLTabContentClose

    # OneDrive
    $rpt += Get-HTMLTabContentOpen -TabName 'OneDrive' -TabHeading "Microsoft 365 OneDrive"
    $rpt += Get-HTMLContentOpen -HeaderText "OneDrive"
    $rpt += Get-HTMLTableFromCSV -CSVPath "$ReportSavePath\CombinedData.csv" -TypeFilter "OneDrive"
    $rpt += Get-HTMLContentClose
    $rpt += Get-HTMLTabContentClose

    # Fechando a página
    $rpt += Get-HTMLClosePage

    # Salvando relatório
    $rptString = $rpt -join "`r`n"
    $htmlPath = "$ReportSavePath\Microsoft365Report.html"
    Set-Content -Path $htmlPath -Value $rptString
    Write-Host "Relatório HTML gerado com sucesso." -ForegroundColor Green

    # Abrindo o relatório HTML no navegador padrão
    Start-Process $htmlPath

} catch {
    $errorMessage = "Erro ao gerar relatório HTML: $_"
    Write-Error $errorMessage
    Add-Content -Path $ErrorLog -Value $errorMessage
}

# Desconectando do Microsoft Graph
Disconnect-MgGraph
Write-Host "Desconectado do Microsoft Graph com sucesso." -ForegroundColor Green
