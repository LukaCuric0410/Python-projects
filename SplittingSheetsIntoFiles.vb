Sub ExportSheetsToCSV()
    Dim ws As Worksheet
    Dim FilePath As String
    Dim FileName As String
    Dim ValidFileName As String
    Dim FileNum As Integer
    Dim RowNum As Long
    Dim ColNum As Long
    Dim Line As String
    
    ' Postavite putanju gdje će se spremiti CSV datoteke
    FilePath = "/Users/CuricL/Desktop/AnamariaTest/"
    
    ' Provjera i dodavanje završnog backslasha ako nedostaje
    If Right(FilePath, 1) <> "/" Then
        FilePath = FilePath & "/"
    End If

    ' Provjerite postoji li putanja
    If Dir(FilePath, vbDirectory) = "" Then
        MsgBox "Putanja ne postoji! Molimo unesite ispravnu putanju.", vbExclamation
        Exit Sub
    End If

    ' Petlja kroz svaki tab (worksheet) u aktivnoj radnoj knjizi
    For Each ws In ThisWorkbook.Worksheets
        ' Osiguraj ispravno ime fajla (bez neispravnih znakova za naziv fajla)
        ValidFileName = Replace(ws.Name, "/", "_")
        ValidFileName = Replace(ValidFileName, "\", "_")
        ValidFileName = Replace(ValidFileName, ":", "_")
        ValidFileName = Replace(ValidFileName, "?", "_")
        ValidFileName = Replace(ValidFileName, "*", "_")
        ValidFileName = Replace(ValidFileName, "[", "_")
        ValidFileName = Replace(ValidFileName, "]", "_")
        
        ' Kreiraj puni naziv datoteke
        FileName = FilePath & ValidFileName & ".csv"
        
        ' Otvori tekstualni fajl za pisanje
        FileNum = FreeFile
        Open FileName For Output As FileNum
        
        ' Petlja kroz svaki redak i stupac kako bi zapisao podatke u CSV
        For RowNum = 1 To ws.UsedRange.Rows.Count
            Line = ""
            For ColNum = 1 To ws.UsedRange.Columns.Count
                ' Dodaj podatke iz svake ćelije
                Line = Line & IIf(ColNum > 1, ",", "") & ws.Cells(RowNum, ColNum).Text
            Next ColNum
            ' Zapiši redak u CSV fajl
            Print #FileNum, Line
        Next RowNum
        
        ' Zatvori fajl
        Close FileNum
    Next ws
    
    MsgBox "I dalje ti Luka sve omogucuje!", vbInformation
End Sub
