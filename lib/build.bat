@echo off
setlocal
:: call "C:\Program Files (x86)\Microsoft Visual Studio 14.0\VC\bin\vcvars32.bat"
set cl_exe=cl.exe
set bin_dir=..\bin

pushd %~dp0

%cl_exe% /MDd /LD /Tppithy\pithy.c -I. -DPITHY_UNALIGNED_LOADS_AND_STORES -DNDEBUG /Fedebug_pithy.dll /link /def:debug_pithy.def
::%cl_exe% /MDd /LD /Tppithy\pithy.c -I. -DPITHY_UNALIGNED_LOADS_AND_STORES /Fedebug_pithy.dll /Zi /link /def:debug_pithy.def /DEBUG /INCREMENTAL:NO
if errorlevel 1 goto :eof

copy debug_pithy.dll %bin_dir% > nul
copy debug_pithy.pdb %bin_dir% > nul

:cleanup
del /q *.exp *.lib *.obj *.dll *.pdb *.ilk ~$* 2> nul

popd
