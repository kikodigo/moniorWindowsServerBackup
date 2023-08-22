function main {
    $serverNames = @("ad-server","ng-server","servidor-gestse")
    $logsToSend = "";
    $ObjectToSendMail = [ObjectToSend]::new()
   
    try 
    {
        foreach ($currentServerName in $serverNames) {
            CopyFile -serverName $currentServerName
            $resultObject = SearchLogs -servername $currentServerName

            $ObjectToSendMail.eventLog = $($resultObject.logMessages)
            $ObjectToSendMail.hasError = $($resultObject.hasError)

            SendMail -serverName $currentServerName -ObjectToSendMail $ObjectToSendMail 
            

            $resultObject = $null
            $ObjectToSendMail.logToSystem = $null
            $ObjectToSendMail.eventLog = $null
            $ObjectToSendMail.hasError = $true
        }
    }
    catch 
    {
        $ObjectToSendMail.logToSystem =+ "Exceção capturada: $_| "
        $ObjectToSendMail.hasError = $false
        Write-Host "Exceção capturada: $_| "
    }
}

function CopyFile {
    param (
        [string] $serverName
    )
    $result = [ObjectToSend]::new()
    $logPathRemote = "\\" + $serverName + "\c$\Windows\System32\winevt\Logs\Microsoft-Windows-Backup.evtx"
    $localCopyPath = "C:\WinServBackup_Monitor\" + $serverName + "\Microsoft-Windows-Backup.evtx"

    $maxRetries = 3
    $retryDelaySeconds = 5  # 1 minuto em segundos

    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
        try {
            # Tentar copiar o arquivo de log para a máquina local
            Copy-Item -Path $logPathRemote -Destination $localCopyPath -Force
            
            $result.logToSystem = "Cópia bem-sucedida do server $serverName na tentativa $attempt| "
            return $result
        }
        catch {
            if ($attempt -lt $maxRetries) {
                Write-Host "Aguardando 1 minuto"
                Start-Sleep -Seconds $retryDelaySeconds
            }
            else {
                $errorMessage = "Número máximo de tentativas para copiar no servidor $serverName| "
                throw $errorMessage
            }
        }
    }
}

function SearchLogs {
    param (
        [string]$servername
    )

    $logPath = "C:\WinServBackup_Monitor\" + $serverName + "\Microsoft-Windows-Backup.evtx"
    $eventIDs = @(5, 19, 9, 49, 517, 561, 20, 4)
    
    # Obter a data de hoje à meia-noite
    $todayMidnight = (Get-Date).Date
    
    # Subtrair 1 dia para obter a data do dia anterior
    $yesterdayMidnight = $todayMidnight.AddDays(-1)
    
    # Subtrair 1 segundo para obter o último segundo do dia anterior
    $yesterdayEnd = $yesterdayMidnight.AddDays(1).AddSeconds(-1)
    
    $backupErrors = Get-WinEvent -Path $logPath -ErrorAction SilentlyContinue |
        Where-Object { $_.TimeCreated -ge $yesterdayMidnight -and $_.TimeCreated -lt $yesterdayEnd -and $eventIDs -contains $_.Id }    

    $logMessages = @()

    $hasError = $false

    if ($backupErrors.Count -gt 0) {
        foreach ($errorEvent in $backupErrors) {
            $eventTime = $errorEvent.TimeCreated
            $eventMessage = "$($eventTime.ToString('yyyy-MM-dd HH:mm:ss')) - $($errorEvent.Message) - EventID $($errorEvent.Id.ToString())"
            $logMessages += $eventMessage
            
            if ($errorEvent.Id -ne 4) {
                $hasError = $true
            }
        }
    } else {
        $logMessages = "Nenhum evento de erro do backup encontrado nas últimas 2 horas."
    }

    $result = [PSCustomObject]@{
        logMessages = $logMessages
        hasError = $hasError
    }
    # Retorna as mensagens de erro reportadas
    return $result
}

function SendMail {
    param (
        [string] $serverName,
        [ObjectToSend] $ObjectToSendMail
    )    
  
    $conclusao ="sucesso"
    if($ObjectToSendMail.hasError){
        $conclusao = "problema"
    }

    if ($ObjectToSendMail.eventLog -ne  $null -and $ObjectToSendMail.eventLog.Count -gt 0) {
        for ($i = 0; $i -lt $ObjectToSendMail.eventLog.Count; $i++) {
            $ObjectToSendMail.eventLog[$i] = "<li><span class=`"bullet`"></span>" + $ObjectToSendMail.eventLog[$i] + "</li><br>"
        }
    }

    if ($ObjectToSendMail.logToSystem -ne  $null -and $ObjectToSendMail.logToSystem.Count -gt 0) {
        for ($i = 0; $i -lt $ObjectToSendMail.logToSystem.Count; $i++) {
            $ObjectToSendMail.logToSystem[$i] = "<li><span class=`"bullet`"></span>" + $ObjectToSendMail.logToSystem[$i] + "</li><br>"
        }        
    }

    # Enviar E-mail
    $emailSmtpServer = "mail.gestservi.com.br"
    $emailSmtpServerPort = "587"
    $emailSmtpUser = "backup@gestservi.com.br"
    $emailSmtpPass = "bla"

    $emailMessage = New-Object System.Net.Mail.MailMessage
    $emailMessage.From = "BackupGestServi<backup@gestservi.com.br>"
    $emailMessage.To.Add( "backup@gestservi.com.br" )
    $emailMessage.Subject = "Backup com "+$conclusao+" no servidor " +$serverName 
    $emailMessage.IsBodyHtml = $true
    $emailMessage.Body = 
    " <style>
        ul {
            list-style-type: none; /* Remove os marcadores padrão da lista */
            padding-left: 0; /* Remove o espaçamento padrão da lista */
        }

        .bullet {
            display: inline-block;
            width: 6px;
            height: 6px;
            background-color: black;
            border-radius: 50%;
            margin-right: 5px; /* Espaçamento entre a bolinha e o texto */
            vertical-align: middle; /* Alinhamento vertical ao texto */
        }
    </style>    
    <ul>
    <h2>Dados do Evento</h2>
    <br>
    " + $ObjectToSendMail.eventLog + "
    <br>
    Logs Gerados
    <br>" +  $ObjectToSendMail.logToSystem +
    "</ul>"

    $SMPTClient = New-Object System.Net.Mail.SmtpClient( $emailSmtpServer, $emailSmtpServerPort)
    $SMPTClient.EnableSsl = $false
    $SMPTClient.Credentials = New-Object System.Net.NetworkCredential( $emailSmtpUser, $emailSmtpPass)
    $SMPTClient.Send($emailMessage)
}

class ObjectToSend {
    [string[]] $eventLog
    [string[]] $logToSystem
    [bool] $hasError
}

main