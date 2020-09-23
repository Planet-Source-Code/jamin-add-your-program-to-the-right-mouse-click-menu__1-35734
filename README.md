<div align="center">

## Add your program to the Right\-Mouse\-Click menu


</div>

### Description

Let your program be listed in the right-mouse-click for a particular file type. You program can then be run with a simple click
 
### More Info
 
The filetype you want to add to. Your program's path & filename.

This program only adds the regisrty key. It can be removed by running regedit then manually deleting the key. The location can be seen in the code.

Adds a registry key

Playing with the registry can be dangerous. I have run this on my computer and it works fine. I take no responsibility for any mishaps that occur.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jamin](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jamin.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jamin-add-your-program-to-the-right-mouse-click-menu__1-35734/archive/master.zip)

### API Declarations

```
Const REG_SZ = 1
Const HKEY_LOCAL_MACHINE = &H80000002
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
```


### Source Code

```
Public Sub SetStringValue(Hkey As Long, strPath As String, strValue As String, strdata As String)
  Dim keyhand As Long
  Dim i As Long
  i = RegCreateKey(Hkey, strPath, keyhand)
  'Changed the normal keyvalue name to vbNullString to set it as the (Default) value
  i = RegSetValueEx(keyhand, vbNullString, 0, REG_SZ, ByVal strdata, Len(strdata))
  i = RegCloseKey(keyhand)
End Sub
Private Sub Form_Load()
'Create Reg Key for filetype, (can be found by looking in HKEY_CLASSES_ROOT then the extension, the (Default) value pionts to the file type, located in the HKEY_CLASSES_ROOT folder)
'Replace """C:\Project1.exe"" %1" with your program path & name. "%1" means send the file name you click on as the commandline arguments.
SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\Paint.Picture\shell\Text To Display\command", 0, """C:\Project1.exe"" %1"
End Sub
```

