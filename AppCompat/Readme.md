[AppCompatFlagsView]

List of program path

- OS: Windows 10 64Bit
- Source: HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags  
HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags  
- Usage:  
`AppCompatFlagsView.exe` //print results  
`AppCompatFlagsView.exe -csv` //export csv file  
- Screenshot  

![image](https://user-images.githubusercontent.com/69110090/95338643-1d323380-08ee-11eb-8eea-1bb011ccdcd4.png)   

[AppCompatCacheAnalyzer]  

List of program path and mod time  

- OS: Windows 10 64Bit  
- Source: %SystemRoot%\System32\config\SYSTEM  
- Usage:  
`AppCompatCacheAnalyzer.exe -f SYSTEM` //Analyze a file  
`AppCompatCacheAnalyzer.exe -f SYSTEM -local` //Time value is displayed in local time  
- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/119341433-b6238c80-bcce-11eb-9d5b-80d5a286edec.png)  
