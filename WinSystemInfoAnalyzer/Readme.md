[WinSystemInfoAnalyzer]  

List of Windows system information  

- OS: Windows 10 64Bit  
- Source: %SystemRoot%\System32\config\SOFTWARE   
%SystemRoot%\System32\config\SYSTEM  
%UserProfile%\NTUSER.DAT  
- Usage:  
`WinSystemInfoAnalyzer.exe -f SOFTWARE` //Analyze a file  
`WinSystemInfoAnalyzer.exe -f SYSTEM` //Analyze a file  
`WinSystemInfoAnalyzer.exe -f NTUSER.DAT  ` //Analyze a file  
`WinSystemInfoAnalyzer.exe -d RegistryFolder` //Analyze folder that contains NTUSER.DAT, SOFTWARE, SYSTEM  
`WinSystemInfoAnalyzer.exe -f SYSTEM -local` //Time value is displayed in local time  
- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/126628888-9bc33137-82f8-4e67-9436-ce6dd508f11b.png)  
