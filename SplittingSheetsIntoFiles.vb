Sub ExportSheetsToCSVIfMultipleTextsFound()
    Dim ws As Worksheet
    Dim FilePath As String
    Dim FileName As String
    Dim ValidFileName As String
    Dim FileNum As Integer
    Dim RowNum As Long
    Dim ColNum As Long
    Dim Line As String
    Dim CheckCellContent As String
    Dim SpecificTexts As String
    Dim TextArray() As String
    Dim TextFound As Boolean
    Dim i As Long
    
    ' Prikazuje InputBox za unos vi?e tekstova, odvojenih zarezima
    SpecificTexts = InputBox("Unesite tekstove koje zelite traziti u cijelom radnom listu, odvojene zarezima:", "Unos trazenih tekstova")
    
    ' Ako korisnik ne unese ni?ta, iza?i iz procedure
    If SpecificTexts = "" Then
        MsgBox "Niste unijeli nijedan tekst.", vbExclamation
        Exit Sub
    End If
    
    ' Razdvajanje unesenih tekstova na temelju zareza
    TextArray = Split(SpecificTexts, ",")
    
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
                
                ' Provjera ima li ?elija bilo koji od unesenih tekstova
                For i = LBound(TextArray) To UBound(TextArray)
                    ' Uklanja nepotrebne razmake s po?etka i kraja teksta
                    If Trim(CheckCellContent) = Trim(TextArray(i)) Then
                        TextFound = True ' Prona?en jedan od tra?enih tekstova
                        Exit For ' Iza?i iz petlje za pretragu teksta
                    End If
                Next i
                
                ' Ako je tekst prona?en, iza?i iz petlje za stupce
                If TextFound Then Exit For
            Next ColNum
            
            ' Ako je tekst prona?en, iza?i iz petlje za redove
            If TextFound Then Exit For
        Next RowNum
        
        ' Ako je prona?en bilo koji tra?eni tekst na radnom listu, kreiraj CSV datoteku
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
    
    MsgBox "Zapamti,Luka ti je sve omogucio!", vbInformation
End Sub

