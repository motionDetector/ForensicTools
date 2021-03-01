[JumplistAnalyzer]  

Analyze jumplist files(.automaticDestinations-ms or .customDestinations-ms)  

- OS: Windows 10 64Bit  
- Source: %AppData%\Microsoft\Windows\Recent\AutomaticDestinations\\*.automaticDestinations-ms  
%AppData%\Microsoft\Windows\Recent\CustomDestinations\\*.customDestinations-ms  

- Usage:  
`JumplistAnalyzer.exe -f 469m4o7t8ion4d4.automaticDestinations-ms` //Analyze a jumplist file   
`JumplistAnalyzer.exe -d C:\JumplistFolder` //Analyze jumplist files   
`JumplistAnalyzer.exe -d C:\JumplistFolder -local` //Time value is displayed in local time  

- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/109501970-b51dfd00-7adb-11eb-958f-d35f86e096e7.png)  
