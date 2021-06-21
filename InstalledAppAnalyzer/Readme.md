[InstalledAppAnalyzer]  

List of installed programs         

- OS: Windows 10 64Bit  
- Source: %UserProfile%\NTUSER.DAT   
%SystemRoot%\System32\config\SOFTWARE  
- Usage:  
`InstalledAppAnalyzer.exe -f NTUSER.DAT` //Analyze a file   
`InstalledAppAnalyzer.exe -f SOFTWARE` //Analyze a file    
`InstalledAppAnalyzer.exe -d RegistryFolder` //Analyze folder that contains NTUSER.DAT, SOFTWARE  
`InstalledAppAnalyzer.exe -f NTUSER.DAT -local` //Time value is displayed in local time   
- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/122758877-448e2c80-d2d4-11eb-9ffc-6ceaf2bbb23a.png)  
