@echo off
set RETAIL=1
rem *** set value ***
set arcname=ppxqjsT11.zip
set readme=ppxqjs.txt
set srcname=ppxqjsrc.7z
set exename=ppxqjs

rem *** main ***
del /q *.zip 2> NUL
del /q .obj\*.obj 2> NUL
del /q .obj\*.res 2> NUL
del /q !!*.* 2> NUL

cmd /c "mingw32.bat && make.exe -E RELEASE=1"
cmd /c "mingw64.bat && make.exe -E RELEASE=1"

tfilesign sp %exename%32.DLL %exename%32.DLL
tfilesign sp %exename%64.DLL %exename%64.DLL
rem tfilesign sp %exename%64A.DLL %exename%64A.DLL

rem *** Source Archive ***
if %RETAIL%==0 goto :skipsource

for %%i in (*) do CT %readme% %%i

ppb /c %%u/7-zip32.dll,a %srcname% -hide -mx=9 makefile *.bat *.c *.cpp pp*.h torowin.h *.RC *.RH *.ICO *.DEF vs2015\ppxscr.vcxproj vs2015\ppxscr.sln
CT %readme% %srcname%
:skipsource

for %%i in (%exename%*.dll) do CT %readme% %%i
for /R %%i in (*.sln) do CT %readme% %%i
for /R %%i in (*.vc*) do CT %readme% %%i
for %%i in (sample*.*) do CT %readme% %%i

ppb /c %%u/7-ZIP32.DLL,a -tzip -hide -mx=7 %arcname% %readme% %exename%*.dll sample*.* %srcname%
del /q %srcname% 2> NUL
tfilesign s %arcname% %arcname%
CT %readme% %arcname%
