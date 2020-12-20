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


