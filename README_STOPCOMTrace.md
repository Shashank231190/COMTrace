Set of files

Name
----
Application_Windows Error Reporting_1001.xml = profiler
StartLogic.ps1 = Start all the traces
StopLogic.ps1  = Create the scheduled task using profiler "Application_Windows Error Reporting_1001.xml"
TerminateLogic.ps1 = Terminate the traces forcefully.


NOTE:  Application_Windows Error Reporting_1001.xml, is just a sample profiler.
Similar profiler needs to be created to monitor event of choice and required changes needs to be done in mentioned filter,  Filter source :  StopLogic.ps1 


#Filter
$EventLogName = "Application"
$EventLevel = "4"
$EventId = "1001"
$EventEntryType = "Information"
$EventProviderName = "Windows Error Reporting"
$MessageToMonitor = "...."

Must actions:


1.Keep xperf toolkit on c:\ drive
2.Update $User and $Pwd field with Local administrator and password
3.Keep StopComTrace directory on C:\