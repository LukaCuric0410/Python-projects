Sub SplitWorksheetsIntoFiles()
    Dim ws As Worksheet
    Dim NewWb As Workbook
    Dim FilePath As String
    Dim FileName As String
    Dim ValidFileName As String
    
    ' Ovdje se postavlja putanja di se trebaju spremiti file-ovi(PROMIJENI ZA SEBE)
    FilePath = "/FolderPath/"
    
    ' Na ovome jos radim
    If Right(FilePath, 1) <> "/" Then
        FilePath = FilePath & "/"
    End If

    ' provjera putanje jeli postojana 
    If Dir(FilePath, vbDirectory) = "" Then
        MsgBox "Putanja ne postoji! Molimo unesite ispravnu putanju.", vbExclamation
        Exit Sub
    End If

    ' Excel ima naviku otvarati nove stvorene workbook-ove pa ih ovo zatvara
    Application.ScreenUpdating = False

    ' Petlja koja prolazi kroz svaki worksheet u workbook-u
    For Each ws In ThisWorkbook.Worksheets
        ' Kreiranje novih Workbook-ova od svakog worksheet-a
        Set NewWb = Workbooks.Add(xlWBATWorksheet)
        
        ' Kopiraj trenutni worksheet u novi workbook
        ws.Copy Before:=NewWb.Sheets(1)
        
        ' Izbriši prazni sheet koji se automatski generira
        Application.DisplayAlerts = False
        NewWb.Sheets(2).Delete
        Application.DisplayAlerts = True
        
        ' Osiguraj ispravno ime fajla (bez neispravnih znakova za naziv fajla)
        ValidFileName = Replace(ws.Name, "/", "_")
        ValidFileName = Replace(ValidFileName, "\", "_")
        ValidFileName = Replace(ValidFileName, ":", "_")
        ValidFileName = Replace(ValidFileName, "?", "_")
        ValidFileName = Replace(ValidFileName, "*", "_")
        ValidFileName = Replace(ValidFileName, "[", "_")
        ValidFileName = Replace(ValidFileName, "]", "_")
        
        ' Spremi workbook pod imenom taba
        FileName = FilePath & ValidFileName & ".xlsx"
        
        ' Spremi novi workbook bez otvaranja
        NewWb.SaveAs FileName:=FileName, FileFormat:=xlOpenXMLWorkbook
        
        ' Zatvori workbook bez prikazivanja
        NewWb.Close False
    Next ws
    
    ' Ponovno uključi prikaz ekrana
    Application.ScreenUpdating = True

    MsgBox "Zapamti da ti je Luka sve omogucio !", vbInformation
End Sub