@echo off


:: Sunucu Adresi ve Portu
SET SERVER=192.168.2.20   
SET PORT=1688
cd /d %~dp0
cls
GOTO WINDOWS
GOTO OFFICE




:WINDOWS
	SET COMMON="%windir%\system32\slmgr.vbs"
	IF EXIST %COMMON% (
		CLS
		ECHO Windows Etkinlestiriliyor....
		cscript  %COMMON% //B /skms %SERVER%:%PORT%
		cscript  %COMMON% //B /ato
		

	) 	

	cscript  %COMMON% /dlv | findstr /I /B /C:"Lisans Durumu:" |findstr /I "Lisansl"
	SET FLAG=%ERRORLEVEL%
		IF %FLAG%==0 (
		ECHO WINDOWS AKTIF EDILDI....
	)
	cscript  %COMMON% /dlv | findstr /I /B /C:"License Status:" |findstr /I "Licensed"
	SET FLAG=%ERRORLEVEL%
	IF %FLAG%==0 (
		ECHO WINDOWS AKTIF EDILDI.... 
	)
	
:OFFICE
	echo.
	ECHO =====================================================
	echo.
	:2010
		SET COMMON="%ProgramFiles%\Microsoft Office\Office14\OSPP.VBS"

		IF EXIST %COMMON% (
			ECHO Office 2010 bulundu
			ECHO office 2010 icin Kurulum basliyor. 
			cscript //B %COMMON% /sethst:%SERVER%
			cscript //B  %COMMON% /setprt:%PORT%
			cscript //B  %COMMON% /act 
			cls
			echo ...OFFICE 2010 AKTIF EDILDI.....

		)
		cscript %COMMON% /dstatus |findstr /I /B /C:"LICENSE STATUS:"|findstr /I  "LICENSED"
		SET FLAG=%ERRORLEVEL%
		IF %FLAG%==0 (
			echo ...OFFICE 2010 AKTIF EDILDI.....
		)
		
		SET X64_32="%ProgramFiles(x86)%\Microsoft Office\Office14\OSPP.VBS"  
		IF EXIST %X64_32% (
			ECHO Office 2010 bulundu
			ECHO office 2010 icin Kurulum basliyor. 
			cscript //B  %X64_32% /sethst:%SERVER%
			cscript //B  %X64_32% /setprt:%PORT%
			cscript //B  %X64_32% /act
			cls
			

		)
	   cscript %X64_32% /dstatus |findstr /I /B /C:"LICENSE STATUS:"|findstr /I  "LICENSED"
	   SET FLAG=%ERRORLEVEL%
	   IF %FLAG%==0 (
			echo ...OFFICE 2010 AKTIF EDILDI.....
	   )
	 :: =======================================================================




	::::  OFFICE 2013 x32 x64 BIT 

	:2013
		SET COMMON="%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS"

		IF EXIST %COMMON% (
			ECHO Office 2013 bulundu
			ECHO office 2013 icin Kurulum basliyor. 
			cscript //B  %COMMON% /sethst:%SERVER%
			cscript //B  %COMMON% /setprt:%PORT%
			cscript //B  %COMMON% /act
			
			
		   
		   )
		cscript %COMMON% /dstatus |findstr /I /B /C:"LICENSE STATUS:"|findstr /I  "LICENSED"
		SET FLAG=%ERRORLEVEL%
		IF %FLAG%==0 (
			echo ...OFFICE 2013 AKTIF EDILDI.....
		)
		   
		 SET X64_32="%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS"  
		 IF EXIST %X64_32% (
			ECHO Office 2013 bulundu
			ECHO office 2013 icin Kurulum basliyor. 
			cscript //B  %X64_32% /sethst:%SERVER%
			cscript //B  %X64_32% /setprt:%PORT%
			cscript //B  %X64_32% /act
			
			
		   
		   )
		cscript %X64_32% /dstatus |findstr /I /B /C:"LICENSE STATUS:"|findstr /I  "LICENSED"
		SET FLAG=%ERRORLEVEL%
		IF %FLAG%==0 (
			echo ...OFFICE 2013 AKTIF EDILDI.....
		)   
	 :: =======================================================================

	 
	 
	::::  OFFICE 2016 x32 x64 BIT 

	:2016
			SET COMMON="%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS"

		IF EXIST %COMMON% (
			ECHO Office 2016 bulundu
			ECHO office 2016 icin Kurulum basliyor. 
			cscript //B  %COMMON% /sethst:%SERVER%
			cscript //B  %COMMON% /setprt:%PORT%
			cscript //B  %COMMON% /act
			
			
		   
		   )
		cscript %COMMON% /dstatus |findstr /I /B /C:"LICENSE STATUS:"|findstr /I  "LICENSED"
		SET FLAG=%ERRORLEVEL%
		IF %FLAG%==0 (
			echo ...OFFICE 2016 AKTIF EDILDI.....
		)
		   
		 SET X64_32="%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS"  
		 IF EXIST %X64_32% (
			ECHO Office 2016 bulundu
			ECHO office 2016 icin Kurulum basliyor. 
			cscript //B  %X64_32% /sethst:%SERVER%
			cscript //B  %X64_32% /setprt:%PORT%
			cscript //B  %X64_32% /act
			
			
		   
		   )
		cscript %X64_32% /dstatus |findstr /I /B /C:"LICENSE STATUS:"|findstr /I  "LICENSED"
		SET FLAG=%ERRORLEVEL%
		IF %FLAG%==0 (
			echo ...OFFICE 2016 AKTIF EDILDI.....
		)   
	   
	 :: =======================================================================


 
pause
