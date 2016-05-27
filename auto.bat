cd C:\ProgramData\SafeNet Sentinel\Sentinel RMS Development Kit\System
del *.* /F /A /Q

@ping 127.0.0.1 -n 5 -w 1000 > nul

cd C:\Program Files (x86)\HP\Unified Functional Testing\bin
start instdemo.exe
