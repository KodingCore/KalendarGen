Attribute VB_Name = "Module1"
Sub Kalender()
    SupprimerToutesLesFeuilles

    Dim thisSheet As Worksheet
    
    Dim startYear As Integer
    startYear = 2025
    Dim year As Integer
    year = startYear
    
    Dim startMonth As Integer
    Dim howMonths As Integer
    Dim endMonth As Integer
    startMonth = 9
    howMonths = 15
    endMonth = startMonth + howMonths - 1
    
    Dim textThisMonth As String
    
    
    For i = startMonth To endMonth
        Dim j As Integer
        j = i
        If j > 12 Then
            j = j - 12
            If j = 1 Then
                year = year + 1
            End If
        End If
        Select Case j
            Case 1
                CreateNewSheet j, "Janvier", CStr(year)
            Case 2
                CreateNewSheet j, "Février", CStr(year)
            Case 3
                CreateNewSheet j, "Mars", CStr(year)
            Case 4
                CreateNewSheet j, "Avril", CStr(year)
            Case 5
                CreateNewSheet j, "Mai", CStr(year)
            Case 6
                CreateNewSheet j, "Juin", CStr(year)
            Case 7
                CreateNewSheet j, "Juillet", CStr(year)
            Case 8
                CreateNewSheet j, "Août", CStr(year)
            Case 9
                CreateNewSheet j, "Septembre", CStr(year)
            Case 10
                CreateNewSheet j, "Octobre", CStr(year)
            Case 11
                CreateNewSheet j, "Novembre", CStr(year)
            Case 12
                CreateNewSheet j, "Décembre", CStr(year)
            Case Else
                MsgBox "j est hors context"
                ' Code à exécuter si aucune correspondance trouvée
        End Select
        
    Next i
    
End Sub


Sub CreateNewSheet(ByVal month As Integer, ByVal monthStr As String, ByVal year As String)
    
    Dim sheetName As String
    Dim classeur As Workbook
    Dim newSheet As Worksheet
    Dim color As Long
    
    ' Obtenir le classeur actif
    Set classeur = ActiveWorkbook
    
    ' Ajouter une nouvelle feuille
    Set newSheet = classeur.Sheets.Add
    
    ' Renommer la nouvelle feuille si nécessaire
    sheetName = month & year
    newSheet.Name = sheetName
    
    color = RandomColor
    
    SetTitle monthStr, year, color, newSheet
    SetTableau month, year, color, newSheet

End Sub

Sub SetTableau(ByVal month As Integer, ByVal yearStr As String, ByVal color As Long, ByVal sheet As Worksheet)
    Dim year As Integer
    Dim premierJourDuMois As Integer
    Dim firstBox As Range
    Set firstBox = sheet.Range("B3")
    
    
    Dim nombreJours As Integer
    

    Dim plageTableHeader As Range
    Dim plageTableBody As Range
    
    Set plageTableHeader = sheet.Range("B2:H2")
    Set plageTableBody = sheet.Range("B3:H9")
    
    year = Val(yearStr)
    nombreJours = NombreJoursDansMois(month, year)
    
    plageTableHeader.Rows.RowHeight = 40
    plageTableBody.Rows.RowHeight = 100
    
    plageTableHeader(1).Value = "Lundi"
    plageTableHeader(2).Value = "Mardi"
    plageTableHeader(3).Value = "Mercredi"
    plageTableHeader(4).Value = "Jeudi"
    plageTableHeader(5).Value = "Vendredi"
    plageTableHeader(6).Value = "Samedi"
    plageTableHeader(7).Value = "Dimanche"
    
    plageTableHeader.HorizontalAlignment = xlCenter
    plageTableHeader.VerticalAlignment = xlCenter
    plageTableHeader.Font.color = color
    With plageTableHeader.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    premierJourDuMois = FirstDay(month, year)
    
    Set firstBox = firstBox.Offset(0, premierJourDuMois - 1)
    
    If premierJourDuMois > 1 Then
        Dim nombreJoursMoisPrecedent As Integer
        nombreJoursMoisPrecedent = NombreJoursDansMois(month - 1, year)
        Dim j As Integer
        For j = premierJourDuMois - 1 To 1 Step -1
            Set firstBox = firstBox.Offset(0, -1)
            firstBox.Value = nombreJoursMoisPrecedent
            firstBox.Font.color = RGB(150, 150, 150)
            firstBox.HorizontalAlignment = xlLeft
            firstBox.VerticalAlignment = xlTop
            firstBox.BorderAround ColorIndex:=1, Weight:=xlThin, LineStyle:=xlDot
            nombreJoursMoisPrecedent = nombreJoursMoisPrecedent - 1
        Next j
    
    End If
    
    Set firstBox = firstBox.Offset(0, premierJourDuMois - 1)
    firstBox.Value = "1"
    firstBox.HorizontalAlignment = xlLeft
    firstBox.VerticalAlignment = xlTop
    firstBox.BorderAround ColorIndex:=1, Weight:=xlThin, LineStyle:=xlDash
    
    Dim i As Integer
    
    For i = 2 To nombreJours
        Set firstBox = firstBox.Offset(0, 1)
        
        If firstBox.Column = 9 Then
            
            Set firstBox = firstBox.Offset(1, -7)
        End If
        firstBox.Value = CStr(i)
        firstBox.HorizontalAlignment = xlLeft
        firstBox.VerticalAlignment = xlTop
        firstBox.BorderAround ColorIndex:=1, Weight:=xlThin, LineStyle:=xlDash
    Next i
    
    i = 1
    While firstBox.Column < 8
        Set firstBox = firstBox.Offset(0, 1)
        
        firstBox.Value = CStr(i)
        firstBox.Font.color = RGB(150, 150, 150)
        firstBox.HorizontalAlignment = xlLeft
        firstBox.VerticalAlignment = xlTop
        firstBox.BorderAround ColorIndex:=1, Weight:=xlThin, LineStyle:=xlDot
        i = i + 1
    Wend
    
    
    
    sheet.Columns("A").ColumnWidth = 2
    sheet.Columns("I").ColumnWidth = 2
    
    sheet.PageSetup.PrintArea = "$A$1:$I$8"
    
End Sub


Function NombreJoursDansMois(ByVal mois As Integer, ByVal annee As Integer) As Integer
    ' Créer la date au premier jour du mois suivant
    Dim dateUse As Date
    dateUse = DateSerial(annee, mois + 1, 1)
    
    ' Soustraire un jour pour obtenir le dernier jour du mois actuel
    Dim dernierJour As Date
    dernierJour = dateUse - 1
    
    ' Extraire le jour du mois
    NombreJoursDansMois = Day(dernierJour)
End Function

Function FirstDay(ByVal month As Integer, ByVal year As Integer) As Integer

    Dim premierJour As Date
    premierJour = DateSerial(year, month, 0)
    FirstDay = Weekday(premierJour)
    
End Function


Sub SetTitle(ByVal month As String, ByVal year As String, ByVal color As Long, ByVal sheet As Worksheet)
    Dim plageTitre As Range
    Set plageTitre = sheet.Range("B1:H1")
    plageTitre.Merge
    plageTitre.Value = month & " " & year
    sheet.Rows(1).RowHeight = 100
    plageTitre.HorizontalAlignment = xlCenter
    plageTitre.VerticalAlignment = xlCenter
    
    plageTitre.Font.Size = 36
    plageTitre.Font.Name = "MS Gothic"
    
    plageTitre.Font.color = color
    
    plageTitre.HorizontalAlignment = xlCenterAcrossSelection
End Sub

Function RandomColor() As Long
    Dim randomRed As Integer
    Dim randomGreen As Integer
    Dim randomBlue As Integer
    
    ' Générer un nombre aléatoire entre 1 et 100
    randomRed = Int((255 - 1 + 1) * Rnd + 1)
    randomGreen = Int((255 - 1 + 1) * Rnd + 1)
    randomBlue = Int((255 - 1 + 1) * Rnd + 1)
    RandomColor = RGB(randomRed, randomGreen, randomBlue)
End Function

Sub SupprimerToutesLesFeuilles()
    Dim ws As Worksheet
    
    Application.DisplayAlerts = False ' Désactiver les alertes de suppression
    
    ' Parcourir toutes les feuilles et les supprimer
    Dim i As Integer
    i = ThisWorkbook.Sheets.Count
    For Each ws In ThisWorkbook.Sheets
    If i > 0 And ThisWorkbook.Sheets.Count > 1 Then
        ws.Delete
    End If
        i = i - 1
    Next ws
    
    Application.DisplayAlerts = True ' Réactiver les alertes
    
    ' Ajouter une nouvelle feuille vide
    ThisWorkbook.Sheets.Add
End Sub
