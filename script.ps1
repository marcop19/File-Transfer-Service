Function LogWrite
{
   Param (
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")][String]$Level = "INFO",
	[string]$logstring
   )
   
   $Stamp = (Get-Date).toString("yyyy-MM-dd HH:mm:ss")
   
   Add-content $Logfile -value "$Stamp [$Level] $logstring"
}

Function Connection {
    Param ([string]$hostname, [string]$username, $pwd, $typeservice, $portnumber, $autosecurityaccept, $timeout)
    try {
            # Carico WinSCP .NET assembly
            Add-Type -Path "WinSCPnet.dll"
 
            # Parametri di connessione FTP
            If ($typeservice -eq "ftp") {
                $protocol = [WinSCP.Protocol]::Ftp

                $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
                    Protocol = $protocol
                    HostName = $hostname
                    PortNumber = $portnumber
                    UserName = $username
                    Password = $pwd
                    TimeoutInMilliseconds = $timeout
                }
            }

            # Parametri di connessione SFTP
            If ($typeservice -eq "sftp") {
                $protocol = [WinSCP.Protocol]::Ftp

                $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
                    Protocol = $protocol
                    HostName = $hostname
                    PortNumber = $portnumber
                    UserName = $username
                    Password = $pwd
                    TimeoutInMilliseconds = $timeout
                    GiveUpSecurityAndAcceptAnySshHostKey = True
                }
            }

            $session = New-Object WinSCP.Session

            # Effettuo la connessione
            $session.Open($sessionOptions)

            return $session
    }
    Catch {
            LogWrite "ERROR" "$($_.Exception.Message)"
            exit 1
    }
}

Function Upload {
        
    # Inizializzo la connessione
    $session = Connection $xml.Configuration.Connection.Hostname `
      $xml.Configuration.Connection.Username `
      $xml.Configuration.Connection.Password `
      $xml.Configuration.Connection.Typeservice `
      $xml.Configuration.Connection.PortNumber `
      $xml.Configuration.Connection.Autosecurityaccept `
      $xml.Configuration.Connection.Timeout

    Try {
        # Imposto il resume ad off
        $transferOptions = New-Object WinSCP.TransferOptions
        $transferOptions.ResumeSupport.State = [WinSCP.TransferResumeSupportState]::Off

        $suffix = "_filepart"

        $files = Get-ChildItem ($xml.Configuration.Directory.ExportLocalPath + "*.$($xml.Configuration.Service.FileExtension)")
        foreach ($fileInfo in $files) { 

            # Effettuo l'upload e calcolo il tempo che impiega
            $time = (Measure-Command { 
              $session.PutFiles($fileInfo, `
                $xml.Configuration.Directory.ExportRemotePath + $fileInfo.Name + $suffix, `
                $False, `
                $transferOptions).Check()  
            }).TotalSeconds
            
            LogWrite "INFO" "Upload $($fileInfo.Name + $suffix) OK ($($time) seconds)"

            # Rinomimo il file ripristinando il suffisso originale
            $time = (Measure-Command { 
              $a = $xml.Configuration.Directory.ExportRemotePath + $fileInfo.Name + $suffix
              $b = $xml.Configuration.Directory.ExportRemotePath + $fileInfo.Name

              $session.MoveFile($a, $b)
            }).TotalSeconds
        
            #Trasferimento di upload OK"
            LogWrite "INFO" "Rename $($fileInfo.Name + $suffix) in $($fileInfo.Name) OK ($($time) seconds)"
                            
            # Provo a spostare il sorgente nella cartella di backup
            Move-Item $fileInfo $xml.Configuration.Directory.BackupExportLocalPath -Force -ErrorAction Stop
        }
    }
   
    catch {
        LogWrite "ERROR" "$($_.Exception.Message)"
        exit 1
    }

    finally {
        # Disconnessione
        $session.Dispose()
    }
}

Function Download {
    
    # Inizializzo la connessione
    $session = Connection $xml.Configuration.Connection.Hostname `
      $xml.Configuration.Connection.Username `
      $xml.Configuration.Connection.Password `
      $xml.Configuration.Connection.Typeservice `
      $xml.Configuration.Connection.PortNumber `
      $xml.Configuration.Connection.Autosecurityaccept `
      $xml.Configuration.Connection.Timeout

    $suffix = "_filepart"

    Try {
        $files = $session.EnumerateRemoteFiles($xml.Configuration.Directory.ImportRemotePath, `
          "*.$($xml.Configuration.Service.FileExtension)" , [WinSCP.EnumerationOptions]::None)

        foreach ($fileInfo in $files) {

            # Download file
            $time = (Measure-Command { 
                $session.GetFiles($xml.Configuration.Directory.ImportRemotePath + $fileInfo, `
                $xml.Configuration.Directory.ImportLocalPath + $fileInfo + $suffix).Check()
        
                # Se esiste un file con lo stesso nome lo cancello
                if (Test-Path -Path ($xml.Configuration.Directory.ImportLocalPath + $fileInfo)) {
                    Remove-Item ($xml.Configuration.Directory.ImportLocalPath + $fileInfo)
                }
            
                Rename-Item ($xml.Configuration.Directory.ImportLocalPath + $fileInfo + $suffix) `
                ($xml.Configuration.Directory.ImportLocalPath + $fileInfo)
            }).TotalSeconds

            LogWrite "INFO" "Download $($fileInfo) OK ($($time) seconds)"

            # Cancello il file
            $time = (Measure-Command {
                $session.RemoveFiles($fileInfo.FullName)
            }).TotalSeconds

            LogWrite "INFO" "Remove remote $($fileInfo) OK ($($time) seconds)"
        }
    }

    catch {
        LogWrite "ERROR" "$($_.Exception.Message)"
        exit 1
    }

    finally {
        # Disconnessione
        $session.Dispose()
    }

}

# Recupero i parametri da file di configurazione
$XMLPath = "d:\progetti\AG_Ftpservice\config.xml"
$xml = [xml](Get-Content $XMLPath)

# Inizializzo file di log
$Logfile = $xml.Configuration.Directory.LogPath + "$(Get-Date -f yyyy-MM-dd).log"

# Avvio esportazione se attivo in configurazione
If ($xml.Configuration.Service.ExportService -eq "Yes") {
    Upload
}

# Avvio importazione se attivo in configurazione
If ($xml.Configuration.Service.ImportService -eq "Yes") {
    Download
}

