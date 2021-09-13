[EvtxDevicesAnalyzer]  

List of information related to devices   

- OS: Windows 10 64Bit  
- Source: C:\Windows\System32\winevt\\*.evtx  
- Supported event log:  
System.evtx  
Microsoft-Windows-Partition%4Diagnostic.evtx  
Microsoft-Windows-Ntfs%4Operational.evtx  
Microsoft-Windows-Kernel-PnP%4Configuration  
Microsoft-Windows-DriverFrameworks-UserMode%4Operational.evtx  
- Usage:  
`EvtxDevicesAnalyzer.exe -f Microsoft-Windows-Partition%4Diagnostic.evtx` //Analyze a file  
`EvtxDevicesAnalyzer.exe -d EventlogFolder` //Analyze files  
`EvtxDevicesAnalyzer.exe -f Microsoft-Windows-Partition%4Diagnostic.evtx -local` //Time value is displayed in local time  
- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/132985678-2eab1626-9613-469b-aaf0-a52e48833914.png)  

