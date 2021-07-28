[WinDevicesAnalyzer]  

List of connected devices information  

- OS: Windows 10 64Bit  
- Source: %SystemRoot%\System32\config\SYSTEM   
%SystemRoot%\System32\config\SOFTWARE   
%UserProfile%\NTUSER.DAT    

- Usage:  
`WinDevicesAnalyzer.exe -f SYSTEM` //Analyze a file   
`WinDevicesAnalyzer.exe -f SYSTEM -sw SOFTWARE` //Additional Analysis for volume name  
`WinDevicesAnalyzer.exe -f SYSTEM -nt NTUSER.DAT` //Additional Analysis for volume name  
`WinDevicesAnalyzer.exe -f SYSTEM -local` //Time value is displayed in local time   
- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/127318904-9d30b405-2e47-474c-89d8-e3ef5ce01a1e.png)  
