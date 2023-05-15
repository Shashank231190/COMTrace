<#
   1. Collect WPR Trace
   2. Collect COM Trace
   3. Collect Event logs
   3. Event based triggered

#>
$XperfPath = "c:\Xperf"
$LogDrive = "C:\"
$LogDir = "MSComTraces"
$Logs = $LogDrive + $LogDir
$EventLogFile = "EventLogs.log"
if (-Not(Test-Path -Path $Logs)) {
    New-item -Path $LogDrive -Name $LogDir -ItemType Directory
}

$GetEventLogs = {

    $timeStamp = get-date -Format "MM_dd_yyyy_HH_mm"
    
    $systemLogPath = $timeStamp + "_System.evtx"
    $systemLogPath = $Logs + "\" + $systemLogPath
    
    $applicationLogPath = $timeStamp + "_Application.evtx"
    $applicationLogPath = $Logs + "\" + $applicationLogPath
      
    wevtutil.exe epl System $systemLogPath
    wevtutil.exe epl Application $applicationLogPath

    #get-winevent -path  $applicationLogPath -MaxEvents 10 | out-file -FilePath $Logs"\"$EventLogFile -append
  
    
}

$StartComTrace = {

    "Enabling COM Trace.."

    reg add HKEY_LOCAL_MACHINE\Software\Microsoft\OLE\Tracing /v ExecutablesToTrace /t REG_MULTI_SZ /d * /f
    logman create trace "com_complus" -ow -o $Logs"\com_complus.etl" -p `{A0C4702B-51F7-4EA9-9C74-E39952C694B8`} 0xffffffffffffffff 0xff -nb 16 16 -bs 1024 -mode Circular -f bincirc -max 4096 -ets
    logman update trace "com_complus" -p `{53201895-60E8-4FB0-9643-3F80762D658F`} 0xffffffffffffffff 0xff -ets
    logman update trace "com_complus" -p `{B46FA1AD-B22D-4362-B072-9F5BA07B046D`} 0xffffffffffffffff 0xff -ets
    logman update trace "com_complus" -p `{9474A749-A98D-4F52-9F45-5B20247E4F01`} 0xffffffffffffffff 0xff -ets
    logman update trace "com_complus" -p `{BDA92AE8-9F11-4D49-BA1D-A4C2ABCA692E`} 0xffffffffffffffff 0xff -ets

}

$StopComTrace = {

    $LogDrive = "C:\"
    $LogDir = "MSComTraces"
    $Logs = $LogDrive + $LogDir
    "Stop COM Trace.."
    logman stop "com_complus" -ets
}


$StartXperfTrace = {

    "Xperf Trace.."
    $timeStamp = get-date -Format HH_mm_ss
    $XperfKernelETL = "Start_Xperf" + $timeStamp + "_kernel.etl"
    $XperfDir = $Logs + "\" + $XperfKernelETL
    Set-Location -Path $XperfPath
    .\xperf -on PROC_THREAD+LOADER+FLT_IO_INIT+FLT_IO+FLT_FASTIO+FLT_IO_FAILURE+FILENAME+FILE_IO+FILE_IO_INIT+DISK_IO+HARD_FAULTS+DPC+INTERRUPT+CSWITCH+PROFILE+DRIVERS+Latency+DISPATCHER -stackwalk MiniFilterPreOpInit+MiniFilterPostOpInit+CSwitch+ReadyThread+ThreadCreate+Profile+DiskReadInit+DiskWriteInit+DiskFlushInit+FileCreate+FileCleanup+FileClose+FileRead+FileWrite+FileFlush -BufferSize 4096 -MaxBuffers 4096 -MaxFile 4096 -FileMode Circular -f $XperfDir 

}

$StopXperf = {

    $LogDrive = "C:\"
    $LogDir = "MSComTraces"
    $Logs = $LogDrive + $LogDir
    set-location -Path $XperfPath
    .\xperf -d $Logs"\Xperf_Wait.etl"
}

$Debug = 0

#Filter
$EventLogName = "System"
$EventLevel = "2"
$EventId = "7034"
$EventEntryType = "Error"
$EventProviderName = "Service Control Manager"
$MessageToMonitor = "The CTACServer Service service terminated unexpectedly"

$FinalEventFilter = @{

    LogName      = $EventLogName
    ProviderName = $EventProviderName
    Level        = $EventLevel
    ID           = $EventId
}



$DummyEvent = {

    Write-EventLog -logname $EventLogName -source $EventProviderName -EntryType $EventEntryType -EventID $EventId -Message "How Dummy I am ?" 
}



if ($Debug -eq 1) {
    .$DummyEvent
    $ErrorMessage = Get-WinEvent -FilterHashtable $FinalEventFilter -MaxEvents 1 -ErrorAction SilentlyContinue
    Write-host "$($ErrorMessage.Message)"
}


$XperfTrace = {

    param($LogName, $ProviderName, $Level, $Id, $MesssageToMonitor)
    
    $StartTime = (get-date).AddMinutes(0)


    $Filter = @{
        LogName      = $LogName
        ProviderName = $ProviderName
        StartTime    = $StartTime
        Level        = $Level
        ID           = $ID

    }

    $StopXperf = {
        $LogDrive = "C:\"
        $LogDir = "MSComTraces"
        $Logs = $LogDrive + $LogDir
        $XperfPath = "c:\Xperf"
        set-location -Path $XperfPath
        .\xperf -d $Logs"\Xperf_Wait.etl"
    }


    Write-Host -ForegroundColor yellow "Xperf start at $StartTime"

    try {
        while (1) {
            $ErrorMessage = Get-WinEvent -FilterHashtable $Filter -MaxEvents 1 -ErrorAction SilentlyContinue
            [string]$Message = $($ErrorMessage.Message) | out-string
            if ($Message.Contains($MesssageToMonitor)) {
                "/////////////////////////////////////////Xperf Trace Collected\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                .$StopXperf
                break
            }

        }
    }
    catch {
        .$StopXperf
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "Xperf : Line at which exception occured => $($_.InvocationInfo.Line)" -ForegroundColor yellow
        
    }



}

$COMTrace = {

    param($LogName, $ProviderName, $Level, $Id, $MesssageToMonitor)
    $StartTime = (get-date).AddMinutes(0)

 
    $Filter = @{
        LogName      = $LogName
        ProviderName = $ProviderName
        StartTime    = $StartTime
        Level        = $Level
        ID           = $ID

    }

    
    $StopComTrace = {

        $LogDrive = "C:\"
        $LogDir = "MSComTraces"
        $Logs = $LogDrive + $LogDir

        "Stop COM Trace.."
        logman stop "com_complus" -ets
    }

    Write-Host -ForegroundColor yellow "COM trace start at $StartTime"

    try {
        while (1) {
            $ErrorMessage = Get-WinEvent -FilterHashtable $Filter -MaxEvents 1 -ErrorAction SilentlyContinue
            [string]$Message = $($ErrorMessage.Message) | out-string
            if ($Message.Contains($MesssageToMonitor)) {
                "/////////////////////////////////////////COM Trace Collected\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                .$StopComTrace
                break
            }

        }
    }
    catch {
        .$StopComTrace
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "COMTrace : Line at which exception occured => $($_.InvocationInfo.Line)" -ForegroundColor yellow
        
    }
}


function Get-ComTrace {
    param(
        [Parameter(Mandatory = $true)]
        $LogPath,
        [Parameter(Mandatory = $true)]
        $XperfLocation
    )

    try {

        $AssertLogPath = Get-ChildItem $LogPath -ErrorAction Stop
        $AssertXperfLocation = Get-ChildItem $XperfLocation -ErrorAction Stop
        "Starting Trace"
        .$StartXperfTrace
        .$StartComTrace

        $XperfJob = Start-Job -ScriptBlock $XperfTrace -ArgumentList @($EventLogName, $EventProviderName, $EventLevel, $EventId, $MessageToMonitor) -name "Xperf"

        $ComJob = Start-Job -ScriptBlock $COMTrace -ArgumentList @($EventLogName, $EventProviderName, $EventLevel, $EventId, $MessageToMonitor) -name "COM"

        do {

            Write-Host -ForegroundColor Yellow "Press 1 - Job Status & Stop (Script will stop, when XperfJob and ComJob shows Status completed)"
            Write-Host -ForegroundColor Yellow "Press 2 - Terminate "
            $Command = read-host "Option"
            switch ($Command) {

                1 {
                    "Job {0} - status {1}" -f (Get-job -Id $XperfJob.Id).Name, (Get-job -id $XperfJob.Id).State
                    "Job {0} - status {1}" -f (Get-job -Id $ComJob.Id).Name, (Get-job -id $ComJob.Id).State
                    $ErrorMessage = Get-WinEvent -FilterHashtable $FinalEventFilter -MaxEvents 1 -ErrorAction SilentlyContinue
                    Write-host "$($ErrorMessage.Message)"
                  
                    if ((get-Job -id $XperfJob.Id).State -eq "Completed" -and (get-job -id $ComJob.Id).State -eq "Completed") {
                        Write-Output $(get-date) | out-file -FilePath $Logs"\JobCompletedSuccessfully.txt"
                        receive-job -job $XperfJob -keep
                        receive-job -job $ComJob -keep
                        $Command = 2    
                        break
                    }
                }
                2 {


                    if ((get-Job -id $XperfJob.Id).State -eq "Completed" -and (get-job -id $ComJob.Id).State -eq "Completed") {
                        Write-Output $(get-date) | out-file -FilePath $Logs"\JobCompletedSuccessfully.txt"
                        receive-job -job $XperfJob -keep
                        receive-job -job $ComJob -keep
                        $Command = 2    
                        break
                    }
                
                    #Writing a dummy event
                    .$DummyEvent
                    Write-Output $(get-date) | out-file -FilePath $Logs"\TracesStoppedForceFully.txt"
                    $Que = read-host "Do you want to terminate the traces forceFully [y/n]"
                    if ($Que -eq 'y') {
                    }
                    else {
                        $Command = 1
                    }
                                       
                }
                
            }#Switch
    

        }while ($Command -ne 2)
    }
    finally {

        .$StopXperf
        .$StopComTrace

        @{ XperfJob       = Receive-Job -job $XperfJob -keep
            ComJob        = Receive-Job -job $ComJob -keep
            XperfJobState = (get-job -id $XperfJob.id).State
            ComJobState   = (get-job -id $ComJob.id).State
             
        } | export-clixml -path $Logs"\diagCatch.xml"

        get-job
        get-job -id $XperfJob.id | Remove-Job -force
        get-job -id $ComJob.id | Remove-Job -force
       
        $logDirName = get-date -Format "MM_dd_yyyy_HH_mm"
        $nameHost = HOSTNAME
        $logDirName = $logDirName + "_" + $nameHost
        New-Item -Name $logDirName -ItemType dir -Path $Logs"\"
        
        $ErrorMessage = Get-WinEvent -FilterHashtable $FinalEventFilter -MaxEvents 1 -ErrorAction SilentlyContinue
        Write-Output $($ErrorMessage.Message) | Out-File -FilePath $Logs"\"$EventLogFile
        .$GetEventLogs
    
        set-location -Path $Logs
        Move-Item -Path $Logs"\*.*" -Destination $Logs"\"$logDirName 
        Write-Host -ForegroundColor white "`n////////////////////////////Data Collected Under $Logs Directory\\\\\\\\\\\\\\\\\\\\\\\\\\\"
    }
}

Get-ComTrace -LogPath $Logs -XperfLocation $XperfPath