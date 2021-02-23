[PrefetchDecompressor]  

Decompress prefetch file(.pf) that compressed as Xpress Huffman Algorithm  

- OS: Windows 10 64Bit  
- Source: C:\Windows\Prefetch\\*.pf  
- Usage:  
`PrefetchDecompressor.exe -f HELLOWORLD.EXE-AB22E9A6.pf`  

- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/108587323-59929780-7396-11eb-9cea-d23073db3340.png)  

[PrefetchAnalyzer]  

Analyze prefetch files(.pf) 

- OS: Windows 10 64Bit  
- Source: C:\Windows\Prefetch\\*.pf  
- Usage:  
`PrefetchAnalyzer.exe -f HELLOWORLD.EXE-AB22E9A6.pf` //Analyze a prefetch file and print info  
`PrefetchAnalyzer.exe -f HELLOWORLD.EXE-AB22E9A6.pf -csv` //Analyze a prefetch file and export csv file  
`PrefetchAnalyzer.exe -d C:\PrefetchFolder` //Analyze prefetch files and export csv file  
`PrefetchAnalyzer.exe -d C:\PrefetchFolder -local` //Time value is displayed in local time  

- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/108844417-72fa4480-761f-11eb-8acf-fb15c2f6d1e8.png)  
