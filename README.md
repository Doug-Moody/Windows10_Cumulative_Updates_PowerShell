# Windows10_Cumulative_Updates_PowerShell
This powershell script can be ran on a system and will identify if a system is patched for CVE-2020-0601 "Curveball" and if not will then download the appropriate patch and execute.  Only works for Windows 10 1507-1909, didn't include for arm based CPUs or embedded versions.  Will update for Server 2016 later. 

This is a cumualtive update so downloads all security related updates in one rollup.


Alternatives:

There has been a write-up for two other options using Powershell -

https://www.virtualizationhowto.com/2020/01/automate-curveball-crypt32-dll-patching/

