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

                      
                          
   
[ShellBagsDesktopView]  

List of folders and fildes in Desktop folder  

- OS: Windows 10 64Bit  
- Source: HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\Shell\Bags\1\Desktop  
- Usage:  
`ShellBagsDesktopView.exe` //print info  
`ShellBagsDesktopView.exe -csv` //export csv file  
- Screenshot  
![image](https://user-images.githubusercontent.com/69110090/99638835-2fc19000-2a8a-11eb-8d88-a1100134cf3d.png)
