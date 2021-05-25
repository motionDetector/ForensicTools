[AppCompatFlagsAnalyzer]  

List of executed programs    

- OS: Windows 10 64Bit  
- Source: %UserProfile%\NTUSER.DAT   
%SystemRoot%\System32\config\SOFTWARE  
- Usage:  
`AppCompatFlagsAnalyzer.exe -f NTUSER.DAT` //Analyze a file   
`AppCompatFlagsAnalyzer.exe -f SOFTWARE` //Analyze a file    
`AppCompatFlagsAnalyzer.exe -d RegistryFolder` //Analyze folder that contains NTUSER.DAT, SOFTWARE  
`AppCompatFlagsAnalyzer.exe -f NTUSER.DAT -local` //Time value is displayed in local time   
- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/119479914-17f5fc00-bd8c-11eb-8c4e-1a9a713cc4be.png)  


[AppCompatCacheAnalyzer]  

List of executed programs   

- OS: Windows 10 64Bit  
- Source: %SystemRoot%\System32\config\SYSTEM  
- Usage:  
`AppCompatCacheAnalyzer.exe -f SYSTEM` //Analyze a file  
`AppCompatCacheAnalyzer.exe -f SYSTEM -local` //Time value is displayed in local time  
- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/119341433-b6238c80-bcce-11eb-9d5b-80d5a286edec.png)  
