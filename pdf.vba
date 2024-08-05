Sub CopyPdfColumnsToNewSheetAndExportToPDF()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim col As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim targetCol As Long
    Dim wsName As String
    Dim pdfFileName As String
    Dim folderPath As String
    Dim currentDate As String

    ' Setze den Namen des Quellarbeitsblatts
    wsName = "Ship 1 " ' Ersetze diesen Namen mit dem korrekten Namen deines Arbeitsblatts
    
    ' Überprüfe, ob das Arbeitsblatt existiert
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets(wsName)
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        MsgBox "Das Arbeitsblatt '" & wsName & "' existiert nicht.", vbCritical
        Exit Sub
    End If
    
    ' Erstelle ein neues Arbeitsblatt für die PDF-Daten
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets("PDF_Export")
    On Error GoTo 0
    
    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Sheets.Add
        wsTarget.Name = "PDF_Export" ' Du kannst den Namen ändern, wenn du möchtest
    Else
        wsTarget.Cells.Clear ' Lösche vorhandene Daten im Zielarbeitsblatt
    End If
    
    ' Finde die letzte benutzte Spalte in Zeile 6
    lastCol = wsSource.Cells(6, wsSource.Columns.Count).End(xlToLeft).Column
    targetCol = 1
    
    ' Durchlaufe alle Spalten in Zeile 6
    For Each col In wsSource.Range(wsSource.Cells(6, 1), wsSource.Cells(6, lastCol))
        If LCase(col.Value) = "pdf" Then
            ' Finde die letzte benutzte Zeile der aktuellen Spalte
            lastRow = wsSource.Cells(wsSource.Rows.Count, col.Column).End(xlUp).Row
            
            ' Kopiere die Zeilen 1-3, inklusive aller Formate
            wsSource.Range(wsSource.Cells(1, col.Column), wsSource.Cells(3, col.Column)).Copy
            wsTarget.Cells(1, targetCol).PasteSpecial Paste:=xlPasteAll
            
            ' Kopiere die Zeilen ab 7 bis zur letzten Zeile, inklusive aller Formate
            If lastRow > 6 Then
                wsSource.Range(wsSource.Cells(7, col.Column), wsSource.Cells(lastRow, col.Column)).Copy
                wsTarget.Cells(7, targetCol).PasteSpecial Paste:=xlPasteAll
            End If
            
            ' Erhöhe die Zielspalte
            targetCol = targetCol + 1
        End If
    Next col
    
    ' Setze den Text "Timeline" in Zelle M7
    wsTarget.Cells(7, 13).Value = "Timeline" ' Spalte M ist die 13. Spalte
    
    ' Überprüfe, ob das Zielarbeitsblatt Daten enthält
    If Application.WorksheetFunction.CountA(wsTarget.Cells) = 0 Then
        MsgBox "Keine Daten zum Exportieren vorhanden.", vbExclamation
        Exit Sub
    End If
    
    ' Setze das Arbeitsblatt auf Querformat und passe den Maßstab an
    With wsTarget.PageSetup
        .Orientation = xlLandscape
        .Zoom = False ' Deaktiviere automatisches Zoom, um Maßstab anzupassen
        .FitToPagesWide = 1 ' Passe an, damit es auf eine Seite in der Breite passt
        .FitToPagesTall = False ' Ignoriere die Höhe
        .PrintArea = wsTarget.UsedRange.Address ' Setze den Druckbereich auf den genutzten Bereich
        .CenterHorizontally = True ' Zentriere horizontal
        .CenterVertically = True ' Zentriere vertikal
    End With
    
    ' AutoAnpassen der Zeilenhöhe für Lesbarkeit
    wsTarget.Rows.AutoFit
    
    ' Setze spezifische Spaltenbreiten
    wsTarget.Columns("I").ColumnWidth = 65.17
    wsTarget.Columns("L").ColumnWidth = 19.33
    
    ' Festlegen des Ordners, in dem die PDF gespeichert werden soll
    folderPath = "/Users/bersintekmen/Library/CloudStorage/OneDrive-FreigegebeneBibliotheken–newboxesGmbH/Operations - RRS_012_2024 F126 Automation/30 Arbeitsdokumente/Part Tracking Lists"
    folderPath = "/Users/bersintekmen/Documents" ' Ersetzen falls nötig
    
    ' Debug-Ausgabe: Überprüfe, ob der Ordner existiert
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Der angegebene Ordner existiert nicht: " & folderPath, vbCritical
        Exit Sub
    End If
    
    ' Erstellen des Dateinamens mit aktuellem Datum
    currentDate = Format(Date, "yyyymmdd")
    pdfFileName = folderPath & Application.PathSeparator & currentDate & "_F126 Parts Tracking List_Damen.pdf"
    
    ' Debug-Ausgabe: Überprüfe den vollständigen Dateipfad
    Debug.Print "PDF wird gespeichert unter: " & pdfFileName
    
    ' Exportiere das Zielarbeitsblatt als PDF
    On Error Resume Next
    wsTarget.ExportAsFixedFormat Type:=xlTypePDF, FileName:=pdfFileName, Quality:=xlQualityStandard
    If Err.Number <> 0 Then
        MsgBox "Fehler beim Erstellen der PDF. Bitte überprüfen Sie den Dateipfad und die Berechtigungen.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Benachrichtige den Benutzer über den Abschluss
    MsgBox "Die PDF-Datei wurde erfolgreich erstellt und gespeichert unter:" & vbCrLf & pdfFileName, vbInformation
End Sub