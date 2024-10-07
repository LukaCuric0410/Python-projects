Sub ExportSheetsToCSVIfSpecificTextFound()
    Dim ws As Worksheet
    Dim FilePath As String
    Dim FileName As String
    Dim ValidFileName As String
    Dim FileNum As Integer
    Dim RowNum As Long
    Dim ColNum As Long
    Dim Line As String
    Dim CheckCellContent As String
    Dim SpecificText As String
    Dim TextFound As Boolean
    
    ' Prikazuje InputBox za unos specifi?nog teksta
    SpecificText = InputBox("Unesite tekst koji zelite traziti u cijelom radnom listu:", "Unos tra?enog teksta")
    
    ' Ako korisnik ne unese ni?ta, iza?i iz procedure
    If SpecificText = "" Then
        MsgBox "Niste unijeli tekst.", vbExclamation
        Exit Sub
    End If
    
    ' Postavite putanju gdje ?e se spremiti CSV datoteke
    FilePath = "/Users/CuricL/Desktop"
    
    ' Provjera i dodavanje zavr?nog backslasha ako nedostaje
    If Right(FilePath, 1) <> "/" Then
        FilePath = FilePath & "/"
    End If

    ' Provjerite postoji li putanja
    If Dir(FilePath, vbDirectory) = "" Then
        MsgBox "Putanja ne postoji! Molimo unesite ispravnu putanju.", vbExclamation
        Exit Sub
    Else
        MsgBox "Putanja je prepoznata: " & FilePath, vbInformation
    End If

    ' Petlja kroz svaki tab (worksheet) u aktivnoj radnoj knjizi
    For Each ws In ThisWorkbook.Worksheets
        TextFound = False ' Resetiraj oznaku za svaki radni list
        
        ' Petlja kroz sve ?elije na radnom listu
        For RowNum = 1 To ws.UsedRange.Rows.Count
            For ColNum = 1 To ws.UsedRange.Columns.Count
                CheckCellContent = ws.Cells(RowNum, ColNum).Value ' ?itanje vrijednosti iz svake ?elije
                
                ' Provjera ima li ?elija specifi?an tekst
                If CheckCellContent = SpecificText Then
                    TextFound = True ' Prona?en specifi?an tekst
                    Exit For ' Iza?i iz petlje za stupce
                End If
            Next ColNum
            
            ' Ako je tekst prona?en, iza?i i iz petlje za redove
            If TextFound Then Exit For
        Next RowNum
        
        ' Ako je prona?en specifi?an tekst na radnom listu, kreiraj CSV datoteku
        If TextFound Then
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
                    ' Dodaj podatke iz svake ?elije
                    Line = Line & IIf(ColNum > 1, ",", "") & ws.Cells(RowNum, ColNum).Text
                Next ColNum
                ' Zapi?i redak u CSV fajl
                Print #FileNum, Line
            Next RowNum
            
            ' Zatvori fajl
            Close FileNum
        End If
    Next ws
    
    MsgBox "I dalje ti Luka sve omogucuje!", vbInformation
End Sub

