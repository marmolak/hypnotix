Attribute VB_Name = "FileExit"
'Toto makro nezbitne je potreba k spravnemu behu programu Microsoft Word
'Jeho odstranenim muzete poskodit spravne pracovani programu Microsoft Word
Option Explicit
Sub main()
 AutoOpen.main
 ActiveDocument.Close
End Sub
