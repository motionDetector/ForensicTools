[MftExtractor]  

Export $MFT of NTFS  

- OS: Windows 10 64Bit (**Need to Run as Administrator**)  
- Source: $MFT  
- Usage: Input NTFS drive letter (ex. C:) and output path  
`MftExtractor.exe -d C: -o C:\OutputFolder`  
- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/100465706-27fd8d80-3113-11eb-9a51-349bd8246b29.png)  


[MftAnalyzer]  

Analyze $MFT of NTFS  

- OS: Windows 10 64Bit  
- Source: $MFT   
- Usage:  
`MftAnalyzer.exe -f $MFT`  
`MftAnalyzer.exe -f $MFT -local` //Time value is displayed in local time  
- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/102706445-4fd8af00-42d5-11eb-869d-5bb1b7e4ed1d.png)  

[LogFileExtractor]  

Export $LogFile of NTFS  

- OS: Windows 10 64Bit (**Need to Run as Administrator**)  
- Source: $LogFile  
- Usage: Input NTFS drive letter (ex. C:) and output path  
`LogFileExtractor.exe -d C: -o C:\OutputFolder`  
- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/102885199-52cecd80-4496-11eb-90dc-c75f996b16c6.png)  

[UsnJournalExtractor]  

Export $J($UsnJrnl) of NTFS  

- OS: Windows 10 64Bit (**Need to Run as Administrator**)  
- Sometimes it takes a long time to find a $J record  
- Source: $J  
- Usage: Input NTFS drive letter (ex. C:) and output path  
`UsnJournalExtractor.exe -d C: -o C:\OutputFolder`  
- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/103439351-815f5c00-4c7f-11eb-82dc-0f38c44115d9.png)  

[UsnJournalAnalyzer]  

Analyze $J of NTFS  

- OS: Windows 10 64Bit  
- Source: $J   
- Usage:  
`UsnJournalAnalyzer.exe -f $J`  
`UsnJournalAnalyzer.exe -f $J -local` //Time value is displayed in local time  
- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/103553000-27190200-4ef0-11eb-871d-541a3e030bae.png)  


