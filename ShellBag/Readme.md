[ShellBagMruAnalyzer]  

List of accessed folders  

- OS: Windows 10 64Bit  
- Source: %UserProfile%\NTUSER.DAT  
%LocalAppData%\Microsoft\Windows\UsrClass.dat  
- Usage:  
`ShellBagMruAnalyzer.exe -f UsrClass.dat` //Analyze a file  
`ShellBagMruAnalyzer.exe -f NTUSER.DAT` //Analyze a file  
`ShellBagMruAnalyzer.exe -d RegistryFolder` //Analyze folder that contains NTUSER.DAT, UsrClass.dat  
`ShellBagMruAnalyzer.exe -f NTUSER.DAT -local` //Time value is displayed in local time  
- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/121765725-0a13e980-cb88-11eb-8daf-8a6105a19b42.png)  

                      
                          
   
[ShellBagsDesktopAnalyzer]  

List of folders and fildes in Desktop    

- OS: Windows 10 64Bit  
- Source: %UserProfile%\NTUSER.DAT  
- Usage:  
`ShellBagsDesktopAnalyzer.exe -f NTUSER.DAT` //Analyze a file  
`ShellBagsDesktopAnalyzer.exe -f NTUSER.DAT -local` //Time value is displayed in local time  
- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/122631272-9cc30400-d105-11eb-9409-ae6740f8bc38.png)  
