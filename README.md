# Jornada 365 | Intune Sync

Bem-vindo ao **Jornada 365 | Intune Sync**! Este script PowerShell foi desenvolvido para facilitar a sincronização de dispositivos gerenciados pelo Microsoft Intune, garantindo que todas as plataformas estejam atualizadas e em conformidade. Além disso, o script oferece uma interface gráfica intuitiva para simplificar a interação com os usuários.

## Funcionalidades

- **Instalação/Atualização de Módulos Necessários**: Verifica e instala ou atualiza os módulos do PowerShell necessários.
- **Conexão ao Microsoft Graph**: Estabelece uma conexão segura com o Microsoft Graph para executar operações de gerenciamento de dispositivos.
- **Sincronização de Dispositivos**: Sincroniza dispositivos Windows, Android, macOS e ChromeOS, garantindo que todos estejam em conformidade.
- **Logs Detalhados**: Gera logs detalhados de todas as ações executadas, permitindo rastrear e solucionar problemas facilmente.
- **Interface Gráfica Intuitiva**: Apresenta uma interface gráfica moderna e fácil de usar, facilitando a seleção e execução de tarefas.
- **Acesso Rápido**: Opções para abrir rapidamente o portal Jornada 365 e o portal Intune diretamente da interface.

  ![MENU SCRIPT](https://github.com/sesantanajr/Itune_Sync/blob/main/Jornada%20365%20Intune%20Sync.png)

## Como Usar

### Pré-requisitos

- **PowerShell 5.1 ou superior**
- **Conexão com a Internet** para instalar módulos e conectar ao Microsoft Graph
- **Permissões Adequadas** no Microsoft Intune para leitura e escrita de dispositivos gerenciados

### Passo a Passo

1. **Clone o Repositório**:
    ```sh
    git clone https://github.com/sesantanajr/Itune_Sync.git
    cd Itune_Sync
    ```

2. **Execute o Script**:
    - Abra o PowerShell como Administrador
    - Navegue até o diretório do script
    - Execute o script:
    ```sh
    powershell -ExecutionPolicy Bypass -File .\Itune_Sync.ps1
    ```

3. **Utilize a Interface Gráfica**:
    - Selecione as opções desejadas para sincronização de dispositivos.
    - Clique em "Executar" para iniciar a sincronização.
    - Acompanhe o progresso através da barra de progresso e das mensagens de status.

### Logs e Resultados

Os logs são gerados automaticamente no diretório `C:\J365_Intune`, contendo detalhes das ações executadas e quaisquer erros encontrados durante a sincronização. Um formulário adicional exibe os resultados da sincronização, que podem ser salvos em formato CSV para futuras referências.

### Interface Gráfica

- **Checkboxes**: Selecione as plataformas de dispositivos que deseja sincronizar.
- **Barra de Progresso**: Acompanhe o progresso da sincronização em tempo real.
- **Botões**: Utilize os botões "Executar" para iniciar a sincronização e "Fechar" para sair da aplicação.

## Vantagens de Usar Este Script

1. **Automatização**: Automatiza a sincronização de dispositivos, reduzindo a necessidade de intervenção manual.
2. **Eficiência**: Garante que todos os dispositivos estejam atualizados e em conformidade de forma rápida e eficiente.
3. **Facilidade de Uso**: A interface gráfica intuitiva torna o processo de sincronização simples, mesmo para usuários com pouca experiência em PowerShell.
4. **Logs Detalhados**: Fornece logs completos e detalhados, facilitando a identificação e resolução de problemas.
5. **Acesso Rápido**: Permite acesso rápido aos portais Jornada 365 e Intune diretamente da interface, melhorando a produtividade.

## Contribuição

Contribuições são bem-vindas! Faça parte dessa jornada você também.
