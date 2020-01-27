@echo off
Title Converter Office 2016 Retail to Volume

echo Press Enter to start VL-Conversion...
echo.
pause
echo.

for /f "tokens=6 delims=[]. " %%G in ('ver') do set win=%%G

set LICPATH=%ProgramFiles(x86)%\Microsoft Office\root\Licenses16

echo path %LICPATH%

if %win% GEQ 9200 (
cd /D "%SystemRoot%\System32"
cscript slmgr.vbs /ilc "%LICPATH%\ProPlusVL_KMS_Client-ppd.xrm-ms"
cscript slmgr.vbs /ilc "%LICPATH%\ProPlusVL_KMS_Client-ul.xrm-ms"
cscript slmgr.vbs /ilc "%LICPATH%\ProPlusVL_KMS_Client-ul-oob.xrm-ms"
cscript slmgr.vbs /ilc "%LICPATH%\client-issuance-bridge-office.xrm-ms"
cscript slmgr.vbs /ilc "%LICPATH%\client-issuance-root.xrm-ms"
cscript slmgr.vbs /ilc "%LICPATH%\client-issuance-root-bridge-test.xrm-ms"
cscript slmgr.vbs /ilc "%LICPATH%\client-issuance-stil.xrm-ms"
cscript slmgr.vbs /ilc "%LICPATH%\client-issuance-ul.xrm-ms"
cscript slmgr.vbs /ilc "%LICPATH%\client-issuance-ul-oob.xrm-ms"
cscript slmgr.vbs /ilc "%LICPATH%\pkeyconfig-office.xrm-ms"
cscript slmgr.vbs /ipk:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
)
if %win% LSS 9200 (
cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /inslic:"%LICPATH%\ProPlusVL_KMS_Client-ppd.xrm-ms"
cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /inslic:"%LICPATH%\ProPlusVL_KMS_Client-ul.xrm-ms"
cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /inslic:"%LICPATH%\ProPlusVL_KMS_Client-ul-oob.xrm-ms"
cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /inslic:"%LICPATH%\pkeyconfig-office.xrm-ms"
cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
)
echo.
echo Retail to Volume License conversion finished.
echo.
pause
