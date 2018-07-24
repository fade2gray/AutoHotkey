# AutoHotkey-HandyTM [Handy Tray Menu]

### "Control all running scripts with one master menu"

**HandyTM-original.ahk** - Original script by SKAN - https://preview.tinyurl.com/HandyTM-SKAN

**HandyTM-modified.ahk** - I modified the original to include use with 64Bit OS running AutoHotkey v1.1.29.01

If you wish to run the script without compiling, you will either need to use a base AHK install of ANSI 32 Bit or use a bootstrap script to run HandyTm.ahk, e.g.
```AutoHotkey
Run, AutoHotkeyA32.exe HandyTM.ahk
```
...or comple using the 'Script to EXE Converter' and select the 'ANSI 32 Bit' base option.

**Note:** The following line of code had to be commented out in the original and modified scripts to prevent AutoHotkey v1.1.29.01 crashing when run on either 32Bit or 64Bit sytems -
```AutoHotkey
DllCall( "LocalFree", UInt,pSid )
```
Only tested on Windows 7 64Bit PC and Windows 7 32Bit VirtualBox VM.
