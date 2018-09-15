:copy .\*.ocx %windir%\SysWow64
:c:
:cd %windir%\SysWow64
regsvr32 comdlg32.ocx
regsvr32 mscomctl.ocx