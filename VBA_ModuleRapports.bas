Attribute VB_Name = "ModuleRapports"
' ========================================
' MODULE RAPPORTS ET STATISTIQUES
' ========================================
' Description: Fonctions pour générer des rapports

Option Explicit

' ========================================
' GENERER RAPPORT MENSUEL
' ========================================
Sub GenererRapportMensuel(mois As Integer, annee As Integer)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim dateDebut As Date, dateFin As Date
    Dim ligne As Long
    
    ' Créer les dates de début et fin du mois
    dateDebut = DateSerial(annee, mois, 1)
    dateFin = DateSerial(annee, mois + 1, 0)
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RAPPORTS)
    ws.Cells.Clear
    
    ' En-tête du rapport
    ws.Cells(1, 1).Value = "RAPPORT MENSUEL - " & Format(dateDebut, "mmmm yyyy")
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    
    ligne = 3
    
    ' Statistiques générales
    ws.Cells(ligne, 1).Value = "STATISTIQUES GÉNÉRALES"
    ws.Cells(ligne, 1).Font.Bold = True
    ligne = ligne + 2
    
    ws.Cells(ligne, 1).Value = "Nombre de réservations :"
    ws.Cells(ligne, 2).Value = CompterReservationsPeriode(dateDebut, dateFin)
    ligne = ligne + 1
    
    ws.Cells(ligne, 1).Value = "Chiffre d'affaires :"
    ws.Cells(ligne, 2).Value = Format(CalculerChiffreAffaires(dateDebut, dateFin), "0.00") & "€"
    ligne = ligne + 1
    
    ws.Cells(ligne, 1).Value = "Taux d'occupation moyen :"
    ws.Cells(ligne, 2).Value = Format(CalculerTauxOccupation(dateDebut, dateFin), "0.00") & "%"
    ligne = ligne + 3
    
    ' Formatage
    ws.Columns("A:B").AutoFit
    
    MsgBox "Rapport mensuel généré avec succès !", vbInformation, APP_NAME
    ws.Activate
    Exit Sub
    
ErrHandler:
    MsgBox "Erreur lors de la génération du rapport : " & Err.Description, vbCritical, APP_NAME
End Sub

' ========================================
' COMPTER RESERVATIONS SUR UNE PERIODE
' ========================================
Function CompterReservationsPeriode(dateDebut As Date, dateFin As Date) As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    Dim compteur As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    compteur = 0
    For i = 2 To derniereLigne
        If ws.Cells(i, 9).Value >= dateDebut And ws.Cells(i, 9).Value <= dateFin Then
            compteur = compteur + 1
        End If
    Next i
    
    CompterReservationsPeriode = compteur
End Function

' ========================================
' CALCULER TAUX D'OCCUPATION
' ========================================
Function CalculerTauxOccupation(dateDebut As Date, dateFin As Date) As Double
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    Dim totalNuits As Long
    Dim nuitsOccupees As Long
    Dim nbChambres As Long
    Dim nbJours As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Compter le nombre de chambres
    nbChambres = CompterNombreChambres()
    nbJours = dateFin - dateDebut + 1
    totalNuits = nbChambres * nbJours
    
    nuitsOccupees = 0
    For i = 2 To derniereLigne
        If ws.Cells(i, 8).Value = "Confirmée" Then
            ' Vérifier si la réservation chevauche avec la période
            If ws.Cells(i, 4).Value <= dateFin And ws.Cells(i, 5).Value >= dateDebut Then
                nuitsOccupees = nuitsOccupees + ws.Cells(i, 6).Value
            End If
        End If
    Next i
    
    If totalNuits > 0 Then
        CalculerTauxOccupation = (nuitsOccupees / totalNuits) * 100
    Else
        CalculerTauxOccupation = 0
    End If
End Function

' ========================================
' COMPTER NOMBRE DE CHAMBRES
' ========================================
Function CompterNombreChambres() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CHAMBRES)
    CompterNombreChambres = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1
End Function
