@echo off
:start
cls
echo.
echo. Choose between Windows and Office activation:
echo. win       	- Windows
echo. off       	- Office
echo.

set /p ed=Activation: 
if %ed%==win (goto win)
if %ed%==off (goto off)
goto start

:win
cls
echo.
echo. Choose a Windows Version:
echo. win10 		- Windows 10
echo. win81 		- Windows 8.1
echo. win8  		- Windows 8
echo. win7  		- Windows 7
echo. winvista  	- Windows Vista
echo. winserv   	- Windows Server
echo.

set /p ed=Version: 
if %ed%==win10 (goto win10)
if %ed%==win81 (goto win81)
if %ed%==win8 (goto win8)
if %ed%==win7 (goto win7)
if %ed%==winvista (goto winvista)
if %ed%==winserv (goto winserv)
goto start

:win10
cls
echo.
echo. Choose an edition of Windows 10:
echo. win10home 	- Windows 10 Home
echo. win10homen	- Windows 10 Home N
echo. win10homesl 	- Windows 10 Home Single Language
echo. win10homecs 	- Windows 10 Home Country Specific
echo. win10pro  	- Windows 10 Professional
echo. win10pron  	- Windows 10 Professional N
echo. win10profw	- Windows 10 Professional for Workstations
echo. win10profwn	- Windows 10 Professional for Workstations N
echo. win10proedu	- Windows 10 Professional Education
echo. win10proedun	- Windows 10 Professional Education N
echo. win10edu  	- Windows 10 Education
echo. win10edun  	- Windows 10 Education N
echo. win10ent  	- Windows 10 Enterprise
echo. win10entn  	- Windows 10 Enterprise N
echo. win10entg 	- Windows 10 Enterprise G
echo. win10entgn	- Windows 10 Enterprise G N
echo. win10ent15 	- Windows 10 Enterprise 2015 LTSB
echo. win10ent15n  	- Windows 10 Enterprise 2015 LTSB N
echo. win10ent16  	- Windows 10 Enterprise 2016 LTSB
echo. win10ent16n  	- Windows 10 Enterprise 2016 LTSB N
echo. win10ent19	- Windows 10 Enterprise 2019 LTSC
echo. win10ent19n	- Windows 10 Enterprise 2019 LTSC N
echo.
echo. Enter "list" to get a list of all Product Keys of Windows 10
echo. 

set /p ed=Edition: 
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==win10home (goto win10home)
if %ed%==win10homen (goto win10homen)
if %ed%==win10homesl (goto win10homesl)
if %ed%==win10homecs (goto win10homecs)
if %ed%==win10pro (goto win10pro)
if %ed%==win10pron (goto win10pron)
if %ed%==win10profw (goto win10profw)
if %ed%==win10profwn (goto win10profwn)
if %ed%==win10proedu (goto win10proedu)
if %ed%==win10proedun (goto win10proedun)
if %ed%==win10edu (goto win10edu)
if %ed%==win10edun (goto win10edun)
if %ed%==win10ent (goto win10ent)
if %ed%==win10entn (goto win10entn)
if %ed%==win10entg (goto win10entg)
if %ed%==win10entgn (goto win10entgn)
if %ed%==win10ent15 (goto win10ent15)
if %ed%==win10ent15n (goto win10ent15n)
if %ed%==win10ent16 (goto win10ent16)
if %ed%==win10ent16n (goto win10ent16n)
if %ed%==win10ent19 (goto win10ent19)
if %ed%==win10ent19n (goto win10ent19n)
if %ed%==list (goto listwin10)
goto start

:win81
cls
echo.
echo. Choose an edition of Windows 8.1:
echo. win81pro		- Windows 8.1 Professional
echo. win81pron		- Windows 8.1 Professional N
echo. win81ent		- Windows 8.1 Enterprise
echo. win81entn		- Windows 8.1 Enterprise N
echo.
echo. Enter "list" to get a list of all Product Keys of Windows 8.1
echo. 

set /p ed=Edition: 
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==win81pro (goto win81pro)
if %ed%==win81pron (goto win81pron)
if %ed%==win81ent (goto win81ent)
if %ed%==win81entn (goto win81entn)
if %ed%==list (goto listwin81)
goto start

:win8
cls
echo.
echo. Choose an edition of Windows 8:
echo. win8pro		- Windows 8 Professional
echo. win8pron		- Windows 8 Professional N
echo. win8ent		- Windows 8 Enterprise
echo. win8entn		- Windows 8 Enterprise N
echo.
echo. Enter "list" to get a list of all Product Keys of Windows 8
echo. 

set /p ed=Edition: 
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==win8pro (goto win8pro)
if %ed%==win8pron (goto win8pron)
if %ed%==win8ent (goto win8ent)
if %ed%==win8entn (goto win8entn)
if %ed%==list (goto listwin8)
goto start

:win7
cls
echo.
echo. Choose an edition of Windows 7:
echo. win7pro		- Windows 7 Professional
echo. win7pron		- Windows 7 Professional N
echo. win7proe		- Windows 7 Professional E
echo. win7ent		- Windows 7 Enterprise
echo. win7entn		- Windows 7 Enterprise N
echo. win7ente		- Windows 7 Enterprise E
echo.
echo. Enter "list" to get a list of all Product Keys of Windows 7
echo. 

set /p ed=Edition: 
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==win7pro (goto win7pro)
if %ed%==win7pron (goto win7pron)
if %ed%==win7proe (goto win7proe)
if %ed%==win7ent (goto win7ent)
if %ed%==win7entn (goto win7entn)
if %ed%==win7ente (goto win7ente)
if %ed%==list (goto listwin7)
goto start

:winvista
cls
echo.
echo. Choose an edition of Windows Vista:
echo. winvistabus	- Windows Vista Business
echo. winvistabusn	- Windows Vista Business N
echo. winvistaent	- Windows Vista Enterprise
echo. winvistaentn	- Windows Vista Enterprise N
echo.
echo. Enter "list" to get a list of all Product Keys of Windows Vista
echo. 

set /p ed=Edition: 
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==winvistabus (goto winvistabus)
if %ed%==winvistabusn (goto winvistabusn)
if %ed%==winvistaent (goto winvistaent)
if %ed%==winvistaentn (goto winvistaentn)
if %ed%==list (goto listwinvista)
goto start

:winserv
cls
echo.
echo. Choose a Version of Windows Server:
echo. winserv19 	- Windows Server 2019
echo. winserv16 	- Windows Server 2016
echo. winserv1909	- Windows Server SAC Version 1909, 1903 and 1809
echo. winserv1803	- Windows Server SAC Version 1803
echo. winserv1709	- Windows Server SAC Version 1709
echo. winserv12r2	- Windows Server 2012 R2
echo. winserv12 	- Windows Server 2012
echo. winserv08r2	- Windows Server 2008 R2
echo. winserv08 	- Windows Server 2008
echo. 

set /p ed=Version: 
if %ed%==winserv19 (goto winserv19)
if %ed%==winserv16 (goto winserv16)
if %ed%==winserv1909 (goto winserv1909)
if %ed%==winserv1803 (goto winserv1803)
if %ed%==winserv1709 (goto winserv1709)
if %ed%==winserv12r2 (goto winserv12r2)
if %ed%==winserv12 (goto winserv12)
if %ed%==winserv08r2 (goto winserv08r2)
if %ed%==winserv08 (goto winserv08)
goto start

:winserv19
cls
echo.
echo. Choose an edition of Windows Server 2019:
echo. winserv19dat	- Windows Server 2019 Datacenter
echo. winserv19stan	- Windows Server 2019 Standart
echo. winserv19ess	- Windows Server 2019 Essentials
echo.
echo. Enter "list" to get a list of all Product Keys of Windows Server 2019
echo. 

set /p ed=Edition: 
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==winserv19dat (goto winserv19dat)
if %ed%==winserv19stan (goto winserv19stan)
if %ed%==winserv19ess (goto winserv19ess)
if %ed%==list (goto listwinserv19)
goto start

:winserv16
cls
echo.
echo. Choose an edition of Windows Server 2016:
echo. winserv16dat	- Windows Server 2016 Datacenter
echo. winserv16stan	- Windows Server 2016 Standart
echo. winserv16ess	- Windows Server 2016 Essentials
echo.
echo. Enter "list" to get a list of all Product Keys of Windows Server 2016
echo. 

set /p ed=Edition: 
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==winserv16dat (goto winserv16dat)
if %ed%==winserv16stan (goto winserv16stan)
if %ed%==winserv16ess (goto winserv16ess)
if %ed%==list (goto listwinserv16)
goto start

:winserv1909
cls
echo.
echo. Choose an edition of Windows Server Version 1909, 1903 and 1809:
echo. winserv1909dat	- Windows Server Datacenter
echo. winserv1909stan	- Windows Server Standart
echo.
echo. Enter "list" to get a list of all Product Keys of Windows Server Version 1909, 1903 and 1809
echo. 

set /p ed=Edition: 
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==winserv1909dat (goto winserv1909dat)
if %ed%==winserv1909stan (goto winserv1909stan)
if %ed%==list (goto listwinserv1909)
goto start

:winserv1803
cls
echo.
echo. Choose an edition of Windows Server Version 1803:
echo. winserv1803dat 	- Windows Server Datacenter
echo. winserv1803stan	- Windows Server Standart
echo.
echo. Enter "list" to get a list of all Product Keys of Windows Server Version 1803
echo. 

set /p ed=Edition: 
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==winserv1803dat (goto winserv1803dat)
if %ed%==winserv1803stan (goto winserv1803stan)
if %ed%==list (goto listwinserv1803)
goto start

:winserv1709
cls
echo.
echo. Choose an edition of Windows Server Version 1709:
echo. winserv1709dat 	- Windows Server Datacenter
echo. winserv1709stan	- Windows Server Standart
echo.
echo. Enter "list" to get a list of all Product Keys of Windows Server Version 1709
echo. 

set /p ed=Edition: 
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==winserv1709dat (goto winserv1709dat)
if %ed%==winserv1709stan (goto winserv1709stan)
if %ed%==list (goto listwinserv1709)
goto start

:winserv12r2
cls
echo.
echo. Choose an edition of Windows Server 2012 R2:
echo. winserv12r2ssta	- Windows Server 2012 R2 Server Standart
echo. winserv12r2dat	- Windows Server 2012 R2 Datacenter
echo. winserv12r2ess	- Windows Server 2012 R2 Essentials
echo.
echo. Enter "list" to get a list of all Product Keys of Windows Server 2012 R2
echo. 

set /p ed=Edition: 
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==winserv12r2ssta (goto winserv12r2ssta)
if %ed%==winserv12r2dat (goto winserv12r2dat)
if %ed%==winserv12r2ess (goto winserv12r2ess)
if %ed%==list (goto listwinserv12r2)
goto start

:winserv12
cls
echo.
echo. Choose an edition of Windows Server 2012:
echo. winserv12a	- Windows Server 2012
echo. winserv12an	- Windows Server 2012 N
echo. winserv12sl	- Windows Server 2012 Single Language
echo. winserv12cs	- Windows Server 2012 Country Specific
echo. winserv12ssta	- Windows Server 2012 Server Standart
echo. winserv12mps	- Windows Server 2012 MultiPoint Standard
echo. winserv12mpp	- Windows Server 2012 MultiPoint Premium
echo. winserv12dat	- Windows Server 2012 Datacenter
echo.
echo. Enter "list" to get a list of all Product Keys of Windows Server 2012 R2
echo. 

set /p ed=Edition: 
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==winserv12a (goto winserv12a)
if %ed%==winserv12an (goto winserv12an)
if %ed%==winserv12sl (goto winserv12sl)
if %ed%==winserv12cs (goto winserv12cs)
if %ed%==winserv12ssta (goto winserv12ssta)
if %ed%==winserv12mps (goto winserv12mps)
if %ed%==winserv12mpp (goto winserv12mpp)
if %ed%==winserv12dat (goto winserv12dat)
if %ed%==list (goto listwinserv12)
goto start

:winserv08r2
cls
echo.
echo. Choose an edition of Windows Server 2008 R2:
echo. winserv08r2web	- Windows Server 2008 R2 Web
echo. winserv08r2hpc	- Windows Server 2008 R2 HPC edition
echo. winserv08r2stan	- Windows Server 2008 R2 Standard
echo. winserv08r2ent	- Windows Server 2008 R2 Enterprise
echo. winserv08r2dat	- Windows Server 2008 R2 Datacenter
echo. winserv08r2fibs 	- Windows Server 2008 R2 for Itanium-based Systems
echo.
echo. Enter "list" to get a list of all Product Keys of Windows Server 2008 R2
echo. 

set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==winserv08r2web (goto winserv08r2web)
if %ed%==winserv08r2hpc (goto winserv08r2hpc)
if %ed%==winserv08r2stan (goto winserv08r2stan)
if %ed%==winserv08r2ent (goto winserv08r2ent)
if %ed%==winserv08r2dat (goto winserv08r2dat)
if %ed%==winserv08r2fibs (goto winserv08r2fibs)   
if %ed%==list (goto listwinserv08r2)
goto start

:winserv08
cls
echo.
echo. Choose an edition of Windows Server 2008:
echo. winserv08web   	- Windows Server 2008 Web
echo. winserv08stan  	- Windows Server 2008 Standard
echo. winserv08stanwhv 	- Windows Server 2008 Standard without Hyper-V
echo. winserv08ent   	- Windows Server 2008 Enterprise
echo. winserv08entwhv 	- Windows Server 2008 Enterprise without Hyper-V
echo. winserv08hpc   	- Windows Server 2008 HPC edition
echo. winserv08dat   	- Windows Server 2008 Datacenter
echo. winserv08datwhv 	- Windows Server 2008 Datacenter without Hyper-V
echo. winserv08fibs  	- Windows Server 2008 for Itanium-based Systems
echo.
echo. Enter "list" to get a list of all Product Keys of Windows Server 2008
echo. 

set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==winserv08web (goto winserv08web)
if %ed%==winserv08stan (goto winserv08stan)
if %ed%==winserv08stanwhv (goto winserv08stanwhv)
if %ed%==winserv08ent (goto winserv08ent)
if %ed%==winserv08entwhv (goto winserv08entwhv)
if %ed%==winserv08hpc (goto winserv08hpc)
if %ed%==winserv08dat (goto winserv08dat)
if %ed%==winserv08datwhv (goto winserv08datwhv)
if %ed%==winserv08fibs (goto winserv08fibs)   
if %ed%==list (goto listwinserv08)
goto start

:off
cls
echo.
echo. Choose an Office-Program:
echo. off 			- Microsoft Office
echo. wor 			- Word
echo. sky 			- Skype
echo. pub			- Publisher
echo. pow			- PowerPoint
echo. out			- Outlook
echo. one			- One Note
echo. exc			- Excel
echo. acc			- Access
echo. vis 			- Visio
echo. pro 			- Project
echo. lyn			- Lync
echo. inf			- InfoPath
echo. sha 			- SharePoint
echo.

set /p ed=Edition:
if %ed%==off (goto offa)
if %ed%==wor (goto wor)
if %ed%==sky (goto sky)
if %ed%==pub (goto pub)
if %ed%==pow (goto pow)
if %ed%==out (goto out)
if %ed%==one (goto one)
if %ed%==exc (goto exc)
if %ed%==acc (goto acc)
if %ed%==vis (goto vis)
if %ed%==pro (goto pro)
if %ed%==lyn (goto lyn)
if %ed%==inf (goto inf)
if %ed%==sha (goto sha)
goto start

:offa
cls
echo.
echo. Choose an Office-Edition:
echo. off10sta		- Office 2010 Standard
echo. off10prop		- Office 2010 Professional Plus
echo. off13sta		- Office 2013 Standard
echo. off13prop		- Office 2013 Professional Plus
echo. off16sta		- Office 2016 Standard
echo. off16prop		- Office 2016 Professional Plus
echo. off19sta		- Office 2019 Standard
echo. off19prop		- Office 2019 Professional Plus
echo.
echo. Enter "list" to get a list of all product keys of Office
echo. 

set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==off10sta (goto off10sta)
if %ed%==off10prop (goto off10prop)
if %ed%==off13sta (goto off13sta)
if %ed%==off13prop (goto off13prop)
if %ed%==off16sta (goto off16sta)
if %ed%==off16prop (goto off16prop)
if %ed%==off19sta (goto off19sta)
if %ed%==off19prop (goto off19prop)
if %ed%==list (goto listoff)
goto start

:wor
cls
echo.
echo. Choose a Word-Edition:
echo. wor10			- Word 2010
echo. wor13			- Word 2013
echo. wor16			- Word 2016
echo. wor19			- Word 2019
echo.
echo. Enter "list" to get a list of all product keys of Word
echo. 
set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==wor10 (goto wor10)
if %ed%==wor13 (goto wor13)
if %ed%==wor16 (goto wor16)
if %ed%==wor19 (goto wor19)
if %ed%==list (goto listwor)
goto start

:sky
cls
echo.
echo. Choose a Skype-Edition:
echo. sky16			- Skype for Business 2016
echo. sky19			- Skype for Business 2019
echo.
echo. Enter "list" to get a list of all product keys of Skype
echo. 
set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==sky16 (goto sky16)
if %ed%==sky19 (goto sky19)
if %ed%==list (goto listsky)
goto start

:pub
cls
echo.
echo. Choose a Publisher-Edition:
echo. pub10			- Publisher 2010
echo. pub13			- Publisher 2013
echo. pub16			- Publisher 2016
echo. pub19			- Publisher 2019
echo.
echo. Enter "list" to get a list of all product keys of Publisher
echo. 
set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==pub10 (goto pub10)
if %ed%==pub13 (goto pub13)
if %ed%==pub16 (goto pub16)
if %ed%==pub19 (goto pub19)
if %ed%==list (goto listpub)
goto start

:pow
cls
echo.
echo. Choose a PowerPoint-Edition:
echo. pow10			- PowerPoint 2010
echo. pow13			- PowerPoint 2013
echo. pow16			- PowerPoint 2016
echo. pow19			- PowerPoint 2019
echo.
echo. Enter "list" to get a list of all product keys of PowerPoint
echo. 
set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==pow10 (goto pow10)
if %ed%==pow13 (goto pow13)
if %ed%==pow16 (goto pow16)
if %ed%==pow19 (goto pow19)
if %ed%==list (goto listpow)
goto start

:out
cls
echo.
echo. Choose an Outlook-Edition:
echo. out10			- Outlook 2010
echo. out13			- Outlook 2013
echo. out16			- Outlook 2016
echo. out19			- Outlook 2019
echo.
echo. Enter "list" to get a list of all product keys of Outlook
echo. 
set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==out10 (goto out10)
if %ed%==out13 (goto out13)
if %ed%==out16 (goto out16)
if %ed%==out19 (goto out19)
if %ed%==list (goto listout)
goto start

:one
cls
echo.
echo. Choose an OneNote-Edition:
echo. one10			- OneNote 2010
echo. one13			- OneNote 2013
echo. one16			- OneNote 2016
echo.
echo. Enter "list" to get a list of all product keys of OneNote
echo. 
set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==one10 (goto one10)
if %ed%==one13 (goto one13)
if %ed%==one16 (goto one16)
if %ed%==list (goto listone)
goto start

:exc
cls
echo.
echo. Choose an Excel-Edition:
echo. exc10			- Excel 2010
echo. exc13			- Excel 2013
echo. exc16			- Excel 2016
echo. exc19			- Excel 2019
echo.
echo. Enter "list" to get a list of all product keys of Excel
echo. 
set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==exc10 (goto exc10)
if %ed%==exc13 (goto exc13)
if %ed%==exc16 (goto exc16)
if %ed%==exc19 (goto exc19)
if %ed%==list (goto listexc)
goto start

:acc
cls
echo.
echo. Choose an Access-Edition:
echo. acc10			- Access 2010
echo. acc13			- Access 2013
echo. acc16			- Access 2016
echo. acc19			- Access 2019
echo.
echo. Enter "list" to get a list of all product keys of Access
echo. 
set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==acc10 (goto acc10)
if %ed%==acc13 (goto acc13)
if %ed%==acc16 (goto acc16)
if %ed%==acc19 (goto acc19)
if %ed%==list (goto listacc)
goto start

:vis
cls
echo.
echo. Choose a Visio-Edition:
echo. vis10sta		- Visio 2010 Standard
echo. vis10pro 		- Visio 2010 Professional
echo. vis10pre 		- Visio 2010 Premium
echo. vis13sta		- Visio 2013 Standard
echo. vis13pro 		- Visio 2013 Professional
echo. vis16sta		- Visio 2016 Standard
echo. vis16pro 		- Visio 2016 Professional
echo. vis19sta 		- Visio 2019 Standard
echo. vis19pro 		- Visio 2019 Professional
echo.
echo. Enter "list" to get a list of all product keys of Visio
echo. 
set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==vis10sta (goto vis10sta)
if %ed%==vis10pro (goto vis10pro)
if %ed%==vis10pre (goto vis10pre)
if %ed%==vis13sta (goto vis13sta)
if %ed%==vis13pro (goto vis13pro)
if %ed%==vis16sta (goto vis16sta)
if %ed%==vis16pro (goto vis16pro)
if %ed%==vis19sta (goto vis19sta)
if %ed%==vis19pro (goto vis19pro)
if %ed%==list (goto listvis)
goto start

:pro
cls
echo.
echo. Choose a Project-Edition:
echo. pro10sta		- Project 2010 Standard
echo. pro10pro 		- Project 2010 Professional
echo. pro13sta		- Project 2013 Standard
echo. pro13pro 		- Project 2013 Professional
echo. pro16sta		- Project 2016 Standard
echo. pro16pro 		- Project 2016 Professional
echo. pro19sta 		- Project 2019 Standard
echo. pro19pro 		- Project 2019 Professional
echo.
echo. Enter "list" to get a list of all product keys of Project
echo. 
set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==pro10sta (goto pro10sta)
if %ed%==pro10pro (goto pro10pro)
if %ed%==pro13sta (goto pro13sta)
if %ed%==pro13pro (goto pro13pro)
if %ed%==pro16sta (goto pro16sta)
if %ed%==pro16pro (goto pro16pro)
if %ed%==pro19sta (goto pro19sta)
if %ed%==pro19pro (goto pro19pro)
if %ed%==list (goto listpro)
goto start

:lyn
cls
echo.
echo. Choose a Lync-Edition:
echo. lyn13			- Lync 2013
echo.
echo. Enter "list" to get a list of all product keys of Lync
echo. 
set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==lyn13 (goto lyn13)
if %ed%==list (goto listlyn)
goto start

:inf
cls
echo.
echo. Choose an InfoPath-Edition:
echo. inf10			- InfoPath 2010
echo. inf13			- InfoPath 2013
echo.
echo. Enter "list" to get a list of all product keys of InfoPath
echo. 
set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==inf10 (goto inf10)
if %ed%==inf13 (goto inf13)
if %ed%==list (goto listinf)
goto start

:sha
cls
echo.
echo. Choose a SharePoint-Edition:
echo. sha10			- SharePoint Workspace 2010
echo.
echo. Enter "list" to get a list of all product keys of SharePoint
echo. 
set /p ed=Edition:
cscript.exe C:\Windows\system32\slmgr.vbs /dlv
cscript.exe C:\Windows\system32\slmgr.vbs /upk
if %ed%==sha10 (goto sha10)
if %ed%==list (goto listsha)
goto start


:win10home
cscript.exe C:\Windows\system32\slmgr.vbs /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
goto exit
:win10homen
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 3KHY7-WNT83-DGQKR-F7HPR-844BM
goto exit
:win10homesl
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
goto exit
:win10homecs
cscript.exe C:\Windows\system32\slmgr.vbs /ipk PVMJN-6DFY6-9CCP6-7BKTT-D3WVR
goto exit
:win10pro
cscript.exe C:\Windows\system32\slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
goto exit
:win10pron
cscript.exe C:\Windows\system32\slmgr.vbs /ipk MH37W-N47XK-V7XM9-C7227-GCQG9
goto exit
:win10profw
cscript.exe C:\Windows\system32\slmgr.vbs /ipk NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
goto exit
:win10profwn
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 9FNHH-K3HBT-3W4TD-6383H-6XYWF
goto exit
:win10proedu
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 6TP4R-GNPTD-KYYHQ-7B7DP-J447Y
goto exit
:win10proedun
cscript.exe C:\Windows\system32\slmgr.vbs /ipk YVWGF-BXNMC-HTQYQ-CPQ99-66QFC
goto exit
:win10edu
cscript.exe C:\Windows\system32\slmgr.vbs /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
goto exit
:win10edun
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
goto exit
:win10ent
cscript.exe C:\Windows\system32\slmgr.vbs /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43
goto exit
:win10entn
cscript.exe C:\Windows\system32\slmgr.vbs /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4
goto exit
:win10entg
cscript.exe C:\Windows\system32\slmgr.vbs /ipk YYVX9-NTFWV-6MDM3-9PT4T-4M68B
goto exit
:win10entgn
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 44RPN-FTY23-9VTTB-MP9BX-T84FV
goto exit
:win10ent15
cscript.exe C:\Windows\system32\slmgr.vbs /ipk WNMTR-4C88C-JK8YV-HQ7T2-76DF9
goto exit
:win10ent15n
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 2F77B-TNFGY-69QQF-B8YKP-D69TJ
goto exit
:win10ent16
cscript.exe C:\Windows\system32\slmgr.vbs /ipk DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ
goto exit
:win10ent16n
cscript.exe C:\Windows\system32\slmgr.vbs /ipk QFFDN-GRT3P-VKWWX-X7T3R-8B639
goto exit
:win10ent19
cscript.exe C:\Windows\system32\slmgr.vbs /ipk M7XTQ-FN8P6-TTKYV-9D4CC-J462D
goto exit
:win10ent19n
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 92NFX-8DJQP-P6BBQ-THF9C-7CG2H
goto exit
:win81pro
cscript.exe C:\Windows\system32\slmgr.vbs /ipk GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
goto exit
:win81pron
cscript.exe C:\Windows\system32\slmgr.vbs /ipk HMCNV-VVBFX-7HMBH-CTY9B-B4FXY
goto exit
:win81ent
cscript.exe C:\Windows\system32\slmgr.vbs /ipk MHF9N-XY6XB-WVXMC-BTDCT-MKKG7
goto exit
:win81entn
cscript.exe C:\Windows\system32\slmgr.vbs /ipk TT4HM-HN7YT-62K67-RGRQJ-JFFXW
goto exit
:win8pro
cscript.exe C:\Windows\system32\slmgr.vbs /ipk NG4HW-VH26C-733KW-K6F98-J8CK4
goto exit
:win8pron
cscript.exe C:\Windows\system32\slmgr.vbs /ipk XCVCF-2NXM9-723PB-MHCB7-2RYQQ
goto exit
:win8ent
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 2JNW-9KQ84-P47T8-D8GGY-CWCK7
goto exit
:win8entn
cscript.exe C:\Windows\system32\slmgr.vbs /ipk JMNMF-RHW7P-DMY6X-RF3DR-X2BQT
goto exit
:win7pro
cscript.exe C:\Windows\system32\slmgr.vbs /ipk FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
goto exit
:win7pron
cscript.exe C:\Windows\system32\slmgr.vbs /ipk MRPKT-YTG23-K7D7T-X2JMM-QY7MG
goto exit
:win7proe
cscript.exe C:\Windows\system32\slmgr.vbs /ipk W82YF-2Q76Y-63HXB-FGJG9-GF7QX
goto exit
:win7ent
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 33PXH-7Y6KF-2VJC9-XBBR8-HVTHH
goto exit
:win7entn
cscript.exe C:\Windows\system32\slmgr.vbs /ipk YDRBP-3D83W-TY26F-D46B2-XCKRJ
goto exit
:win7ente
cscript.exe C:\Windows\system32\slmgr.vbs /ipk C29WB-22CC8-VJ326-GHFJW-H9DH4
goto exit
:winvistabus
cscript.exe C:\Windows\system32\slmgr.vbs /ipk YFKBB-PQJJV-G996G-VWGXY-2V3X8
goto exit
:winvistabusn
cscript.exe C:\Windows\system32\slmgr.vbs /ipk HMBQG-8H2RH-C77VX-27R82-VMQBT
goto exit
:winvistaent
cscript.exe C:\Windows\system32\slmgr.vbs /ipk VKK3X-68KWM-X2YGT-QR4M6-4BWMV
goto exit
:winvistaentn
cscript.exe C:\Windows\system32\slmgr.vbs /ipk VTC42-BM838-43QHV-84HX6-XJXKV
goto exit
:winserv19dat
cscript.exe C:\Windows\system32\slmgr.vbs /ipk WMDGN-G9PQG-XVVXX-R3X43-63DFG
:winserv19stan
cscript.exe C:\Windows\system32\slmgr.vbs /ipk N69G4-B89J2-4G8F4-WWYCC-J464C
goto exit
:winserv19ess
cscript.exe C:\Windows\system32\slmgr.vbs /ipk WVDHN-86M7X-466P6-VHXV7-YY726
goto exit
:winserv16dat
cscript.exe C:\Windows\system32\slmgr.vbs /ipk CB7KF-BWN84-R7R2Y-793K2-8XDDG
goto exit
:winserv16stan
cscript.exe C:\Windows\system32\slmgr.vbs /ipk WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY
goto exit
:winserv16ess
cscript.exe C:\Windows\system32\slmgr.vbs /ipk JCKRF-N37P4-C2D82-9YXRT-4M63B
goto exit
:winserv1909dat
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 6NMRW-2C8FM-D24W7-TQWMY-CWH2D
goto exit
:winserv1909stan
cscript.exe C:\Windows\system32\slmgr.vbs /ipk N2KJX-J94YW-TQVFB-DG9YT-724CC
goto exit
:winserv1803dat
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG
goto exit
:winserv1803stan
cscript.exe C:\Windows\system32\slmgr.vbs /ipk PTXN8-JFHJM-4WC78-MPCBR-9W4KR
goto exit
:winserv1709dat
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 6Y6KB-N82V8-D8CQV-23MJW-BWTG6
goto exit
:winserv1709stan
cscript.exe C:\Windows\system32\slmgr.vbs /ipk DPCNP-XQFKJ-BJF7R-FRC8D-GF6G4
goto exit
:winserv12r2ssta
cscript.exe C:\Windows\system32\slmgr.vbs /ipk D2N9P-3P6X9-2R39C-7RTCD-MDVJX
goto exit
:winserv12r2dat
cscript.exe C:\Windows\system32\slmgr.vbs /ipk W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
goto exit
:winserv12r2ess
cscript.exe C:\Windows\system32\slmgr.vbs /ipk KNC87-3J2TX-XB4WP-VCPJV-M4FWM
goto exit
:winserv12a
cscript.exe C:\Windows\system32\slmgr.vbs /ipk BN3D2-R7TKB-3YPBD-8DRP2-27GG4
goto exit
:winserv12an
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 8N2M2-HWPGY-7PGT9-HGDD8-GVGGY
goto exit
:winserv12sl
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 2WN2H-YGCQR-KFX6K-CD6TF-84YXQ
goto exit
:winserv12cs
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 4K36P-JN4VD-GDC6V-KDT89-DYFKP
goto exit
:winserv12ssta
cscript.exe C:\Windows\system32\slmgr.vbs /ipk XC9B7-NBPP2-83J2H-RHMBY-92BT4
goto exit
:winserv12mps
cscript.exe C:\Windows\system32\slmgr.vbs /ipk HM7DN-YVMH3-46JC3-XYTG7-CYQJJ
goto exit
:winserv12mpp
cscript.exe C:\Windows\system32\slmgr.vbs /ipk XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G
goto exit
:winserv12dat
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 48HP8-DN98B-MYWDG-T2DCC-8W83P
goto exit
:winserv08r2web
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 6TPJF-RBVHG-WBW2R-86QPH-6RTM4
goto exit
:winserv08r2hpc
cscript.exe C:\Windows\system32\slmgr.vbs /ipk TT8MH-CG224-D3D7Q-498W2-9QCTX
goto exit
:winserv08r2stan
cscript.exe C:\Windows\system32\slmgr.vbs /ipk YC6KT-GKW9T-YTKYR-T4X34-R7VHC
goto exit
:winserv08r2ent
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 489J6-VHDMP-X63PK-3K798-CPX3Y
cscript.exe C:\Windows\system32\slmgr.vbs /skms dimanyakms.sytes.net:1688
echo. This may take a while:
cscript.exe
:winserv08r2dat
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 74YFP-3QFB3-KQT8W-PMXWJ-7M648
goto exit
:winserv08r2fibs
cscript.exe C:\Windows\system32\slmgr.vbs /ipk GT63C-RJFQ3-4GMB6-BRFB9-CB83V
goto exit
:winserv08web
cscript.exe C:\Windows\system32\slmgr.vbs /ipk WYR28-R7TFJ-3X2YQ-YCY4H-M249D
goto exit
:winserv08stan
cscript.exe C:\Windows\system32\slmgr.vbs /ipk TM24T-X9RMF-VWXK6-X8JC9-BFGM2
goto exit
:winserv08stanwhv
cscript.exe C:\Windows\system32\slmgr.vbs /ipk W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ
goto exit
:winserv08ent
cscript.exe C:\Windows\system32\slmgr.vbs /ipk YQGMW-MPWTJ-34KDK-48M3W-X4Q6V
goto exit
:winserv08entwhv
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 39BXF-X8Q23-P2WWT-38T2F-G3FPG
goto exit
:winserv08hpc
cscript.exe C:\Windows\system32\slmgr.vbs /ipk RCTX3-KWVHP-BR6TB-RB6DM-6X7HP
goto exit
:winserv08dat
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 7M67G-PC374-GR742-YH8V4-TCBY3
goto exit
:winserv08datwhv
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 22XQ2-VRXRG-P8D42-K34TD-G3QQC
goto exit
:winserv08fibs
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 4DWFP-JF3DJ-B7DTH-78FJB-PDRHK
goto exit

:off10sta
cscript.exe C:\Windows\system32\slmgr.vbs /ipk V7QKV-4XVVR-XYV4D-F7DFM-8R6BM
goto exit
:off10prop
cscript.exe C:\Windows\system32\slmgr.vbs /ipk VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
goto exit
:off13sta
cscript.exe C:\Windows\system32\slmgr.vbs /ipk KBKQT-2NMXY-JJWGP-M62JB-92CD4
goto exit
:off13prop
cscript.exe C:\Windows\system32\slmgr.vbs /ipk YC7DK-G2NP3-2QQC3-J6H88-GVGXT
goto exit
:off16sta
cscript.exe C:\Windows\system32\slmgr.vbs /ipk JNRGM-WHDWX-FJJG3-K47QV-DRTFM
goto exit
:off16prop
cscript.exe C:\Windows\system32\slmgr.vbs /ipk XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
goto exit
:off19sta
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK
goto exit
:off19prop
cscript.exe C:\Windows\system32\slmgr.vbs /ipk NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
goto exit
:wor10
cscript.exe C:\Windows\system32\slmgr.vbs /ipk HVHB3-C6FV7-KQX9W-YQG79-CRY7T
goto exit
:wor13
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7
goto exit
:wor16
cscript.exe C:\Windows\system32\slmgr.vbs /ipk WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6
goto exit
:wor19
cscript.exe C:\Windows\system32\slmgr.vbs /ipk PBX3G-NWMT6-Q7XBW-PYJGG-WXD33 
goto exit
:sky16
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 869NQ-FJ69K-466HW-QYCP2-DDBV6 
goto exit
:sky19
cscript.exe C:\Windows\system32\slmgr.vbs /ipk NCJ33-JHBBY-HTK98-MYCV8-HMKHJ
goto exit
:pub10
cscript.exe C:\Windows\system32\slmgr.vbs /ipk BFK7F-9MYHM-V68C7-DRQ66-83YTP
goto exit
:pub13
cscript.exe C:\Windows\system32\slmgr.vbs /ipk PN2WF-29XG2-T9HJ7-JQPJR-FCXK4
goto exit
:pub16
cscript.exe C:\Windows\system32\slmgr.vbs /ipk F47MM-N3XJP-TQXJ9-BP99D-8K837
goto exit
:pub19
cscript.exe C:\Windows\system32\slmgr.vbs /ipk G2KWX-3NW6P-PY93R-JXK2T-C9Y9V
goto exit
:pow10
cscript.exe C:\Windows\system32\slmgr.vbs /ipk RC8FX-88JRY-3PF7C-X8P67-P4VTT
goto exit
:pow13
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 4NT99-8RJFH-Q2VDH-KYG2C-4RD4F
goto exit
:pow16
cscript.exe C:\Windows\system32\slmgr.vbs /ipk J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6
goto exit
:pow19
cscript.exe C:\Windows\system32\slmgr.vbs /ipk RRNCX-C64HY-W2MM7-MCH9G-TJHMQ 
goto exit
:out10
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ
goto exit
:out13
cscript.exe C:\Windows\system32\slmgr.vbs /ipk QPN8Q-BJBTJ-334K3-93TGY-2PMBT
goto exit
:out16
cscript.exe C:\Windows\system32\slmgr.vbs /ipk R69KK-NTPKF-7M3Q4-QYBHW-6MT9B 
goto exit
:out19
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK
goto exit
:one10
cscript.exe C:\Windows\system32\slmgr.vbs /ipk Q4Y4M-RHWJM-PY37F-MTKWH-D3XHX
goto exit
:one13
cscript.exe C:\Windows\system32\slmgr.vbs /ipk TGN6P-8MMBC-37P2F-XHXXK-P34VW
goto exit
:one16
cscript.exe C:\Windows\system32\slmgr.vbs /ipk DR92N-9HTF2-97XKM-XW2WJ-XW3J6
goto exit
:exc10
cscript.exe C:\Windows\system32\slmgr.vbs /ipk H62QG-HXVKF-PP4HP-66KMR-CW9BM
goto exit
:exc13
cscript.exe C:\Windows\system32\slmgr.vbs /ipk VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB
goto exit
:exc16
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 9C2PK-NWTVB-JMPW8-BFT28-7FTBF
goto exit
:exc19
cscript.exe C:\Windows\system32\slmgr.vbs /ipk TMJWT-YYNMB-3BKTF-644FC-RVXBD
goto exit
:acc10
cscript.exe C:\Windows\system32\slmgr.vbs /ipk V7Y44-9T38C-R2VJK-666HK-T7DDX
goto exit
:acc13
cscript.exe C:\Windows\system32\slmgr.vbs /ipk NG2JY-H4JBT-HQXYP-78QH9-4JM2D
goto exit
:acc16
cscript.exe C:\Windows\system32\slmgr.vbs /ipk GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW
goto exit
:acc19
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT 
goto exit
:vis10sta
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 767HD-QGMWX-8QTDB-9G3R2-KHFGJ 
goto exit
:vis10pro
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 7MCW8-VRQVK-G677T-PDJCM-Q8TCP
goto exit
:vis10pre
cscript.exe C:\Windows\system32\slmgr.vbs /ipk D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ
goto exit
:vis13sta
cscript.exe C:\Windows\system32\slmgr.vbs /ipk J484Y-4NKBF-W2HMG-DBMJC-PGWR7
goto exit
:vis13pro
cscript.exe C:\Windows\system32\slmgr.vbs /ipk C2FG9-N6J68-H8BTJ-BW3QX-RM3B3
goto exit
:vis16sta
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 7WHWN-4T7MP-G96JF-G33KR-W8GF4
goto exit
:vis16pro
cscript.exe C:\Windows\system32\slmgr.vbs /ipk PD3PC-RHNGV-FXJ29-8JK7D-RJRJK
goto exit
:vis19sta
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2
goto exit
:vis19pro
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 9BGNQ-K37YR-RQHF2-38RQ3-7VCBB 
goto exit
:pro10sta
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 4HP3K-88W3F-W2K3D-6677X-F9PGB
goto exit
:pro10pro
cscript.exe C:\Windows\system32\slmgr.vbs /ipk YGX6F-PGV49-PGW3J-9BTGG-VHKC6
goto exit
:pro13sta
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 6NTH3-CW976-3G3Y2-JK3TX-8QHTT
goto exit
:pro13pro
cscript.exe C:\Windows\system32\slmgr.vbs /ipk FN8TT-7WMH6-2D4X9-M337T-2342K
goto exit
:pro16sta
cscript.exe C:\Windows\system32\slmgr.vbs /ipk GNFHQ-F6YQM-KQDGJ-327XX-KQBVC
goto exit
:pro16pro
cscript.exe C:\Windows\system32\slmgr.vbs /ipk YG9NW-3K39V-2T3HJ-93F3Q-G83KT
goto exit
:pro19sta
cscript.exe C:\Windows\system32\slmgr.vbs /ipk C4F7P-NCP8C-6CQPT-MQHV9-JXD2M
goto exit
:pro19pro
cscript.exe C:\Windows\system32\slmgr.vbs /ipk B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
goto exit
:lyn13
cscript.exe C:\Windows\system32\slmgr.vbs /ipk 2MG3G-3BNTT-3MFW9-KDQW3-TCK7R
goto exit
:inf10
cscript.exe C:\Windows\system32\slmgr.vbs /ipk K96W8-67RPQ-62T9Y-J8FQJ-BT37T
goto exit
:inf13
cscript.exe C:\Windows\system32\slmgr.vbs /ipk DKT8B-N7VXH-D963P-Q4PHY-F8894
goto exit
:sha10
cscript.exe C:\Windows\system32\slmgr.vbs /ipk QYYW6-QP4CB-MBV6G-HYMCJ-4T3J4
goto exit

:listwin10
cls
echo.
echo. Windows 10 Home: TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
echo. Windows 10 Home N: 3KHY7-WNT83-DGQKR-F7HPR-844BM
echo. Windows 10 Home Single Language: 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
echo. Windows 10 Home Country Specific: PVMJN-6DFY6-9CCP6-7BKTT-D3WVR
echo. Windows 10 Professional: W269N-WFGWX-YVC9B-4J6C9-T83GX
echo. Windows 10 Professional N: MH37W-N47XK-V7XM9-C7227-GCQG9
echo. Windows 10 Pro for Workstations: NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
echo. Windows 10 Pro for Workstations N: 9FNHH-K3HBT-3W4TD-6383H-6XYWF
echo. Windows 10 Pro Education: 6TP4R-GNPTD-KYYHQ-7B7DP-J447Y
echo. Windows 10 Pro Education N: YVWGF-BXNMC-HTQYQ-CPQ99-66QFC
echo. Windows 10 Education: NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
echo. Windows 10 Education N: 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
echo. Windows 10 Enterprise: NPPR9-FWDCX-D2C8J-H872K-2YT43
echo. Windows 10 Enterprise N: DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4
echo. Windows 10 Enterprise G: YYVX9-NTFWV-6MDM3-9PT4T-4M68B
echo. Windows 10 Enterprise G N: 44RPN-FTY23-9VTTB-MP9BX-T84FV
echo. Windows 10 Enterprise 2015 LTSB: WNMTR-4C88C-JK8YV-HQ7T2-76DF9
echo. Windows 10 Enterprise 2015 LTSB N: 2F77B-TNFGY-69QQF-B8YKP-D69TJ
echo. Windows 10 Enterprise 2016 LTSB: DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ
echo. Windows 10 Enterprise 2016 LTSB N: QFFDN-GRT3P-VKWWX-X7T3R-8B639
echo. Windows 10 Enterprise 2019 LTSC: M7XTQ-FN8P6-TTKYV-9D4CC-J462D
echo. Windows 10 Enterprise 2019 LTSC N: 92NFX-8DJQP-P6BBQ-THF9C-7CG2H
echo.
echo. Press any key to continue...
pause>nul
goto start
:listwin81
cls
echo.
echo. Windows 8.1 Pro: GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
echo. Windows 8.1 Pro N: HMCNV-VVBFX-7HMBH-CTY9B-B4FXY
echo. Windows 8.1 Enterprise: MHF9N-XY6XB-WVXMC-BTDCT-MKKG7
echo. Windows 8.1 Enterprise N: TT4HM-HN7YT-62K67-RGRQJ-JFFXW
echo.
echo. Press any key to continue...
pause>nul
goto start
:listwin8
cls
echo.
echo. Windows 8 Pro: NG4HW-VH26C-733KW-K6F98-J8CK4
echo. Windows 8 Pro N: XCVCF-2NXM9-723PB-MHCB7-2RYQQ
echo. Windows 8 Enterprise: 32JNW-9KQ84-P47T8-D8GGY-CWCK7
echo. Windows 8 Enterprise N: JMNMF-RHW7P-DMY6X-RF3DR-X2BQT
echo.
echo. Press any key to continue...
pause>nul
goto start
:listwin7
cls
echo.
echo. Windows 7 Professional: FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
echo. Windows 7 Professional N: MRPKT-YTG23-K7D7T-X2JMM-QY7MG
echo. Windows 7 Professional E: W82YF-2Q76Y-63HXB-FGJG9-GF7QX
echo. Windows 7 Enterprise: 33PXH-7Y6KF-2VJC9-XBBR8-HVTHH
echo. Windows 7 Enterprise N: YDRBP-3D83W-TY26F-D46B2-XCKRJ
echo. Windows 7 Enterprise E: C29WB-22CC8-VJ326-GHFJW-H9DH4
echo.
echo. Press any key to continue...
pause>nul
goto start
:listwinvista
cls
echo.
echo. Windows Vista Business: YFKBB-PQJJV-G996G-VWGXY-2V3X8
echo. Windows Vista Business N: HMBQG-8H2RH-C77VX-27R82-VMQBT
echo. Windows Vista Enterprise: VKK3X-68KWM-X2YGT-QR4M6-4BWMV
echo. Windows Vista Enterprise N: VTC42-BM838-43QHV-84HX6-XJXKV
echo.
echo. Press any key to continue...
pause>nul
goto start
:listwinserv19
cls
echo.
echo. Windows Server 2019 Datacenter: WMDGN-G9PQG-XVVXX-R3X43-63DFG
echo. Windows Server 2019 Standard: N69G4-B89J2-4G8F4-WWYCC-J464C
echo. Windows Server 2019 Essentials: WVDHN-86M7X-466P6-VHXV7-YY726
echo.
echo. Press any key to continue...
pause>nul
goto start
:listwinserv16
cls
echo.
echo. Windows Server 2016 Datacenter: CB7KF-BWN84-R7R2Y-793K2-8XDDG
echo. Windows Server 2016 Standard: WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY
echo. Windows Server 2016 Essentials: JCKRF-N37P4-C2D82-9YXRT-4M63B
echo.
echo. Press any key to continue...
pause>nul
goto start
:listwinserv1909
cls
echo.
echo. Windows Server Datacenter: 6NMRW-2C8FM-D24W7-TQWMY-CWH2D
echo. Windows Server Standard: N2KJX-J94YW-TQVFB-DG9YT-724CC
echo.
echo. Press any key to continue...
pause>nul
goto start
:listwinserv1803
cls
echo.
echo. Windows Server Datacenter: 2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG
echo. Windows Server Standard: PTXN8-JFHJM-4WC78-MPCBR-9W4KR
echo.
echo. Press any key to continue...
pause>nul
goto start
:listwinserv1709
cls
echo.
echo. Windows Server Datacenter: 6Y6KB-N82V8-D8CQV-23MJW-BWTG6
echo. Windows Server Standard: DPCNP-XQFKJ-BJF7R-FRC8D-GF6G4
echo.
echo. Press any key to continue...
pause>nul
goto start
:listwinserv12r2
cls
echo.
echo. Windows Server 2012 R2 Server Standard: D2N9P-3P6X9-2R39C-7RTCD-MDVJX
echo. Windows Server 2012 R2 Datacenter: W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
echo. Windows Server 2012 R2 Essentials: KNC87-3J2TX-XB4WP-VCPJV-M4FWM
echo.
echo. Press any key to continue...
pause>nul
goto start
:listwinserv12
cls
echo.
echo. Windows Server 2012: BN3D2-R7TKB-3YPBD-8DRP2-27GG4
echo. Windows Server 2012 N: 8N2M2-HWPGY-7PGT9-HGDD8-GVGGY
echo. Windows Server 2012 Single Language: 2WN2H-YGCQR-KFX6K-CD6TF-84YXQ
echo. Windows Server 2012 Country Specific: 4K36P-JN4VD-GDC6V-KDT89-DYFKP
echo. Windows Server 2012 Server Standard: XC9B7-NBPP2-83J2H-RHMBY-92BT4
echo. Windows Server 2012 MultiPoint Standard: HM7DN-YVMH3-46JC3-XYTG7-CYQJJ
echo. Windows Server 2012 MultiPoint Premium: XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G
echo. Windows Server 2012 Datacenter: 48HP8-DN98B-MYWDG-T2DCC-8W83P
echo.
echo. Press any key to continue...
pause>nul
goto start
:listwinserv08r2
cls
echo.
echo. Windows Server 2008 R2 Web: 6TPJF-RBVHG-WBW2R-86QPH-6RTM4
echo. Windows Server 2008 R2 HPC edition: TT8MH-CG224-D3D7Q-498W2-9QCTX
echo. Windows Server 2008 R2 Standard: YC6KT-GKW9T-YTKYR-T4X34-R7VHC
echo. Windows Server 2008 R2 Enterprise: 489J6-VHDMP-X63PK-3K798-CPX3Y
echo. Windows Server 2008 R2 Datacenter: 74YFP-3QFB3-KQT8W-PMXWJ-7M648
echo. Windows Server 2008 R2 for Itanium-based Systems: GT63C-RJFQ3-4GMB6-BRFB9-CB83V
echo.
echo. Press any key to continue...
pause>nul
goto start
:listwinserv08
cls
echo.
echo. Windows Server 2008 Web: WYR28-R7TFJ-3X2YQ-YCY4H-M249D
echo. Windows Server 2008 Standard: TM24T-X9RMF-VWXK6-X8JC9-BFGM2
echo. Windows Server 2008 Standard without Hyper-V: W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ
echo. Windows Server 2008 Enterprise: YQGMW-MPWTJ-34KDK-48M3W-X4Q6V
echo. Windows Server 2008 Enterprise without Hyper-V: 39BXF-X8Q23-P2WWT-38T2F-G3FPG
echo. Windows Server 2008 HPC: RCTX3-KWVHP-BR6TB-RB6DM-6X7HP
echo. Windows Server 2008 Datacenter: 7M67G-PC374-GR742-YH8V4-TCBY3
echo. Windows Server 2008 Datacenter without Hyper-V: 22XQ2-VRXRG-P8D42-K34TD-G3QQC
echo. Windows Server 2008 for Itanium-Based Systems: 4DWFP-JF3DJ-B7DTH-78FJB-PDRHK
echo.
echo. Press any key to continue...
pause>nul
goto start

:listoff
cls
echo.
echo. Office Standard 2010: V7QKV-4XVVR-XYV4D-F7DFM-8R6BM
echo. Office Professional Plus 2010: VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
echo. Office 2013 Standard: KBKQT-2NMXY-JJWGP-M62JB-92CD4
echo. Office 2013 Professional Plus: YC7DK-G2NP3-2QQC3-J6H88-GVGXT
echo. Office Standard 2016: JNRGM-WHDWX-FJJG3-K47QV-DRTFM
echo. Office Professional Plus 2016: XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
echo. Office Standard 2019: 6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK
echo. Office Professional Plus 2019: NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
echo.
echo. Press any key to continue...
pause>nul
goto start
:listwor
cls
echo.
echo. Word 2010: HVHB3-C6FV7-KQX9W-YQG79-CRY7T
echo. Word 2013: 6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7
echo. Word 2016: WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6
echo. Word 2019: PBX3G-NWMT6-Q7XBW-PYJGG-WXD33 
echo.
echo. Press any key to continue...
pause>nul
goto start
:listsky
cls
echo.
echo. Skype for Business 2016: 869NQ-FJ69K-466HW-QYCP2-DDBV6
echo. Skype for Business 2019: NCJ33-JHBBY-HTK98-MYCV8-HMKHJ 
echo.
echo. Press any key to continue...
pause>nul
goto start
:listpub
cls
echo.
echo. Publisher 2010: BFK7F-9MYHM-V68C7-DRQ66-83YTP
echo. Publisher 2013: PN2WF-29XG2-T9HJ7-JQPJR-FCXK4
echo. Publisher 2016: F47MM-N3XJP-TQXJ9-BP99D-8K837
echo. Publisher 2019: G2KWX-3NW6P-PY93R-JXK2T-C9Y9V
echo.
echo. Press any key to continue...
pause>nul
goto start
:listpow
cls
echo.
echo. PowerPoint 2010: RC8FX-88JRY-3PF7C-X8P67-P4VTT
echo. PowerPoint 2013: 4NT99-8RJFH-Q2VDH-KYG2C-4RD4F
echo. PowerPoint 2016: J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6
echo. PowerPoint 2019: RRNCX-C64HY-W2MM7-MCH9G-TJHMQ
echo.
echo. Press any key to continue...
pause>nul
goto start
:listout
cls
echo.
echo. Outlook 2010: 7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ
echo. Outlook 2013: QPN8Q-BJBTJ-334K3-93TGY-2PMBT
echo. Outlook 2016: R69KK-NTPKF-7M3Q4-QYBHW-6MT9B
echo. Outlook 2019: 7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK 
echo.
echo. Press any key to continue...
pause>nul
goto start
:listone
cls
echo.
echo. OneNote 2010: Q4Y4M-RHWJM-PY37F-MTKWH-D3XHX
echo. OneNote 2013: TGN6P-8MMBC-37P2F-XHXXK-P34VW
echo. OneNote 2016: DR92N-9HTF2-97XKM-XW2WJ-XW3J6
echo.
echo. Press any key to continue...
pause>nul
goto start
:listexc
cls
echo.
echo. Excel 2010: H62QG-HXVKF-PP4HP-66KMR-CW9BM
echo. Excel 2013: VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB
echo. Excel 2016: 9C2PK-NWTVB-JMPW8-BFT28-7FTBF
echo. Excel 2019: TMJWT-YYNMB-3BKTF-644FC-RVXBD
echo.
echo. Press any key to continue...
pause>nul
goto start
:listacc
cls
echo.
echo. Access 2010: V7Y44-9T38C-R2VJK-666HK-T7DDX
echo. Access 2013: NG2JY-H4JBT-HQXYP-78QH9-4JM2D
echo. Access 2016: GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW
echo. Access 2019: 9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT 
echo.
echo. Press any key to continue...
pause>nul
goto start
:listvis
cls
echo.
echo. Visio 2010 Standard: 767HD-QGMWX-8QTDB-9G3R2-KHFGJ
echo. Visio 2010 Professional: 7MCW8-VRQVK-G677T-PDJCM-Q8TCP
echo. Visio 2010 Premium: D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ
echo. Visio 2013 Standard: J484Y-4NKBF-W2HMG-DBMJC-PGWR7
echo. Visio 2013 Professional: C2FG9-N6J68-H8BTJ-BW3QX-RM3B3
echo. Visio 2016 Standard: 7WHWN-4T7MP-G96JF-G33KR-W8GF4
echo. Visio 2016 Professional: PD3PC-RHNGV-FXJ29-8JK7D-RJRJK
echo. Visio 2019 Standard: 7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2
echo. Visio 2019 Professional: 9BGNQ-K37YR-RQHF2-38RQ3-7VCBB 
echo.
echo. Press any key to continue...
pause>nul
goto start
:listpro
cls
echo.
echo. Project 2010 Standard: 4HP3K-88W3F-W2K3D-6677X-F9PGB
echo. Project 2010 Professional: YGX6F-PGV49-PGW3J-9BTGG-VHKC6
echo. Project 2013 Standard: 6NTH3-CW976-3G3Y2-JK3TX-8QHTT
echo. Project 2013 Professional: FN8TT-7WMH6-2D4X9-M337T-2342K
echo. Project 2016 Standard: GNFHQ-F6YQM-KQDGJ-327XX-KQBVC
echo. Project 2016 Professional: YG9NW-3K39V-2T3HJ-93F3Q-G83KT
echo. Project 2019 Standard: C4F7P-NCP8C-6CQPT-MQHV9-JXD2M
echo. Project 2019 Professional: B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
echo.
echo. Press any key to continue...
pause>nul
goto start
:listlyn
cls
echo.
echo. Lync 2013: 2MG3G-3BNTT-3MFW9-KDQW3-TCK7R
echo.
echo. Press any key to continue...
pause>nul
goto start
:listinf
cls
echo.
echo. InfoPath 2010: K96W8-67RPQ-62T9Y-J8FQJ-BT37T
echo. InfoPath 2013: DKT8B-N7VXH-D963P-Q4PHY-F8894
echo.
echo. Press any key to continue...
pause>nul
goto start
:listsha
cls
echo.
echo. SharePoint Workspace 2010: QYYW6-QP4CB-MBV6G-HYMCJ-4T3J4
echo.
echo. Press any key to continue...
pause>nul
goto start

:exit
cscript.exe C:\Windows\system32\slmgr.vbs /skms dimanyakms.sytes.net:1688
echo. This may take a while:
cscript.exe C:\Windows\system32\slmgr.vbs /ato
echo. Press any key to exit...
pause>nul
exit