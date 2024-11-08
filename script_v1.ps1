###############
###VARIAVEIS###
###############

###############
###FUNCTIONS###
###############

functionar criaTarefa {

    $scriptPath = "${SYSTEMROOT}\Temp\script_v1.ps1"
    $taskName = "checarProgresso"
    $taskDescription = "Tarefa para checar progresso do script pos formatação"
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$scriptPath`""
    Register-ScheduledTask -Action $action -Trigger $trigger -TaskName $taskName -Description $taskDescription -RunLevel Highest

}

function VerificarLog {
    $caminhoLog = "${SYSTEMROOT}\Temp\pws.log"
    $caminhoTemp = "${SYSTEMROOT}\Temp"

    if (Test-Path $caminhoTemp) {
        Write-Host "Pasta Temp não encontrada. Criando a pasta..."
        New-Item -Path $caminhoTemp -ItemType Directory -Force
    } else {
        Write-Host "Pasta Temp já existe."
    }

    
    # Verifica se o arquivo existe
    if (Test-Path $caminhoLog) {

        # Lê o valor do arquivo
        $valorLog = Get-Content -Path $caminhoLog
        
        # Compara o valor lido com 7
        if ($valorLog -eq "7") {
            Write-Host "Valor encontrado: 7. Finalizando execução."
            #finalizar a tarefa
            return
        } else {
            Write-Host "valor contido no arquivo ainda nao e o valor final. Continuando execução...."
        }
    } else {
        Write-Host "Arquivo de log não encontrado."
        New-Item -Path "${SYSTEMROOT}\Temp\pws.log" -ItemType File -Force
        Set-Content -Path "C:\Temp\pws.log" -Value "0" ### adicionar o valor 0 no arquivo###
        #cria tarefa
    }
}

function InstalaAppx {

    Add-AppProvisionedPackage -Online -PackagePath "${SystemRoot}\Windows\Setup\MSTeams-x64.msix" -SkipLicense;
    Set-Content -Path "C:\Temp\pws.log" -Value "1" ### 2= Instalação dos appx/msi ###

}

function InstalaVPN{

    $bateria = Get-WmiObject -Class Win32_Battery

    if ($bateria.BatteryStatus) {
        Write-Output "Dispositivo é um notebook. Iniciando instalação do FortiClientVPN, Aguarde......"

        $installerPath = "${SystemRoot}\Windows\Setup\Files\FortiClientVPN.exe"
        $installerArgs = "/quiet /norestart"  # ajuste de acordo com o instalador

        Start-Process -FilePath $installerPath -ArgumentList $installerArgs -Wait
        Write-Output "Instalação concluída."
    }
    else {

        Write-Output "Dispositivo não é um notebook. Programa não será instalado."

         }
    Set-Content -Path "C:\Temp\pws.log" -Value "3" ### 3= instalação da VPN ###

}

function selecionaFabricante {

    $fab = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -Property Manufacturer;
    echo $fab;
    $fabL = $fab.Manufacturer.ToLower();

    if($fabL.Contains("lenovo")){
    Copy-Item "${SYSTEMROOT}\Windows\Setup\Files\lenovoupdate.exe" -Destination "${SYSTEMROOT}\Temp"
    echo "lenovo";
    }

    elseif($fabL.Contains("dell")){
    Copy-Item "${SYSTEMROOT}\Windows\Setup\Files\dellupdate.exe" -Destination "${SYSTEMROOT}\Temp"
    echo "dell";
    }
    Set-Content -Path "C:\Temp\pws.log" -Value "4" ### 4= deixar o dellUpdate x lenovoUpdate na temp ###
}

function RemoveInstaladores {

    Remove-Item -Path "${SYSTEMROOT}\Windows\Setup\Files\lenovoupdate.exe"
    Remove-Item -Path "${SYSTEMROOT}\Windows\Setup\Files\dellupdate.exe"
    Remove-Item -Path "${SYSTEMROOT}\Windows\Setup\MSTeams-x64.msix"
    Remove-Item -Path "${SYSTEMROOT}\Windows\Setup\Files\FortiClientVPN.exe"
    Remove-Item -Path "${SYSTEMROOT}\Windows\Setup\Files\Appload Setup.exe";
    Remove-Item -Path "${SYSTEMROOT}\Windows\Setup\Files\setupneto32.exe"
    Remove-Item -Path "${SYSTEMROOT}\Windows\Setup\Files\USB Drivers Installer OPL9728.exe"
    Set-Content -Path "C:\Temp\pws.log" -Value "7" ### 7= caso value igual a 7 script finalizado ###

}

function AdicionaAreadetrabalho {

    # Caminho base para procurar pela pasta MSTeams
    $basePath = "C:\Program Files\WindowsApps"

    # Encontrando a pasta que começa com "MSTeams"
    $msTeamsPath = Get-ChildItem -Path $basePath -Directory | Where-Object { $_.Name -like "MSTeams*" } | Select-Object -First 1

    # Verifica se encontrou a pasta
    if ($msTeamsPath) {
        # Caminho do executável dentro da pasta encontrada
        $targetPath = Join-Path -Path $msTeamsPath.FullName -ChildPath "ms-teams.exe"

        # Caminho para a área de trabalho pública (para todos os usuários)
        $desktopPath = "$env:PUBLIC\Desktop"

        # Nome do atalho
        $shortcutName = "Teams.lnk"

        # Caminho completo do atalho
        $shortcutPath = Join-Path -Path $desktopPath -ChildPath $shortcutName

        # Criando o objeto de atalho
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)

        # Configurando o atalho
        $shortcut.TargetPath = $targetPath
        $shortcut.WorkingDirectory = $msTeamsPath.FullName
        $shortcut.Save()

        Write-Output "Atalho criado em: $shortcutPath"
        }
        else {

            Write-Output "A pasta MSTeams não foi encontrada em $basePath."

        }
        Set-Content -Path "C:\Temp\pws.log" -Value "5" ### 5= adicionar o teams na area de trabalho ###
}

function InstalaColetor {

    Write-Output "UTILITÁRIO DE INSTALAÇÃO DE PROGRAMAS PARA LOJA (APPLOAD) (NET032) E DRIVERS"
    $op=0

    while($op -ne 1){
      $op = Read-Host 'DIGITE "1" PARA INSTALAR OU "0" PARA CANCELAR';
        if($op -eq 1){

             $installerPath = "${SystemRoot}\Windows\Setup\Files\USB Drivers Installer OPL9728.exe"
             Start-Process -FilePath $installerPath -Wait
             Write-Output "Instalação concluída Drivers."
             
             $installerPath = "${SystemRoot}\Windows\Setup\Files\Appload Setup.exe"
             Start-Process -FilePath $installerPath -Wait
             Write-Output "Instalação concluída Appload." 
             
             $installerPath = "${SystemRoot}\Windows\Setup\Files\setupneto32.exe"
             Start-Process -FilePath $installerPath -Wait
             Write-Output "Instalação concluída Neto32."           

        } 
        elseif($op -eq 0){

            $op=1;

        }
        else{

            Write-Output "Por favor digitar um numero valido, 1 ou 0!!";

        }
    }
    Set-Content -Path "C:\Temp\pws.log" -Value "6" ### 6= instalação de driver e aplicativos coletor###
}

#######################
#####CHAMARFUNÇÕES#####
#######################

VerificarLog
InstalaAppx
InstalaVPN
selecionaFabricante
AdicionaAreadetrabalho
InstalaColetor
RemoveInstaladores
