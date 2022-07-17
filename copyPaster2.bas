Attribute VB_Name = "Module4"

Sub TEMP()

    Dim sht As Worksheet

    ScreenUpdating = True

    Sheets("5720040 LYCIANO").Range("N8:N27").Copy
        For Each sht In Worksheets
            sht.Range("N8:N27").PasteSpecial xlPasteAll
        Next

    Application.CutCopyMode = False
    
End Sub

