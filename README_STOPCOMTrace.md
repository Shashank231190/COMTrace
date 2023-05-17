1.Set of files

Name
----
Application_Windows Error Reporting_1001.xml
StartLogic.ps1 = will start all the traces
StopLogic.ps1  = will create the scheduled task using profiler "Application_Windows Error Reporting_1001.xml"
TerminateLogic.ps1 = this will be used in case script needs to be terminated forcefully.


NOTE:  Application_Windows Error Reporting_1001.xml, is just a sample profiler.
Similar profiler needs to be created to monitor event of choice and required changes needs to be done in mentioned filter


 Filter source :  StopLogic.ps1 
#Filter
$EventLogName = "Application"
$EventLevel = "4"
$EventId = "1001"
$EventEntryType = "Information"
$EventProviderName = "Windows Error Reporting"
$MessageToMonitor = "The CTACServer Service service terminated unexpectedly"

Xperf tool kit needs to be kept under c:\ drive.
$Pwd  field needs to be updated in StopLogic source.
