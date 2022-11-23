Attribute VB_Name = "Word_Files"
'Toto makro nezbitne je potreba k spravnemu behu programu Microsoft Word
'Jeho odstranenim muzete poskodit spravne pracovani programu Microsoft Word
Sub main()
 If Not ActiveDocument.Saved Then FileSave.Copy
End Sub
