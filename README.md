# COMTrace
Collect COM Trace and WPR
Keep Xperf Directort on C:\Drive
This trace spawn 2 different jobs to collect COM trace and WPR trace
This could cause perf issues, if run on low CPU machine.
V2 - will use threads, rather than spawing a new process.