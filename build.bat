@echo off
if not exist dist mkdir dist
cl /nologo /EHsc /O2 /MT /D_WIN32_WINNT=0x0601 /LD BKDRCore.cpp /link /DEF:BKDRCore.def /OUT:dist\BKDRCore.dll /IMPLIB:dist\BKDRCore.lib
if errorlevel 1 exit /b 1
cl /nologo /EHsc /O2 /MT /D_WIN32_WINNT=0x0601 BKDRCLI.cpp dist\BKDRCore.lib /link /OUT:dist\BKDRCLI.exe
if errorlevel 1 exit /b 1
copy BKDRCore.h dist\
