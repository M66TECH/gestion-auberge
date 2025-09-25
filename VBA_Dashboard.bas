Attribute VB_Name = "Dashboard"
' ========================================
' MODULE DASHBOARD
' ========================================
' Description: Configuration et gestion du tableau de bord

Option Explicit

' ========================================
' CONFIGURER FEUILLE DASHBOARD
' ========================================
Sub ConfigurerFeuilleDashboard(ws As Worksheet)
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    With ws
        .Cells.Clear
        
        ' Titre principal
        .Cells(1, 1).Value = "🏨 GESTION AUBERGE - TABLEAU DE BORD"
        .Cells(1, 1).Font.Size = 18
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Color = RGB(68, 114, 196)
        .Range("A1:H1").Merge
        .Range("A1").HorizontalAlignment = xlCenter
        
        ' Date du jour
        .Cells(3, 1).Value = "Date : " & Format(Date, "dddd dd mmmm yyyy")
        .Cells(3, 1).Font.Size = 12
        .Cells(3, 1).Font.Bold = True
        
        ' Section Réservations du jour
        .Cells(5, 1).Value = "📅 RÉSERVATIONS DU JOUR"
        .Cells(5, 1).Font.Size = 14
        .Cells(5, 1).Font.Bold = True
        .Cells(5, 1).Interior.Color = RGB(217, 225, 242)
        
        ' Arrivées
        .Cells(7, 1).Value = "Arrivées :"
        .Cells(7, 1).Font.Bold = True
        Call AfficherArriveesDuJour(ws, 8, 1)
        
        ' Départs
        .Cells(12, 1).Value = "Départs :"
        .Cells(12, 1).Font.Bold = True
        Call AfficherDepartsDuJour(ws, 13, 1)
        
        ' Section Statistiques
        .Cells(5, 5).Value = "📊 STATISTIQUES"
        .Cells(5, 5).Font.Size = 14
        .Cells(5, 5).Font.Bold = True
        .Cells(5, 5).Interior.Color = RGB(217, 225, 242)
        
        Call AfficherStatistiques(ws, 7, 5)
        
        ' Section Boutons d'action
        .Cells(18, 1).Value = "🔧 ACTIONS RAPIDES"
        .Cells(18, 1).Font.Size = 14
        .Cells(18, 1).Font.Bold = True
        .Cells(18, 1).Interior.Color = RGB(217, 225, 242)
        
        Call CreerBoutonsDashboard(ws)
        
        ' Formatage général
        .Columns("A:H").AutoFit
        .Range("A1:H30").Font.Name = "Arial"
        
    End With
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de la configuration du dashboard : " & Err.Description, vbCritical, APP_NAME
End Sub

' ========================================
' AFFICHER ARRIVEES DU JOUR
' ========================================
Sub AfficherArriveesDuJour(ws As Worksheet, ligneDebut As Long, colonne As Long)
    Dim arrivees As Variant
    Dim i As Integer
    
    arrivees = ObtenirArriveesDuJour()
    
    For i = 0 To UBound(arrivees)
        ws.Cells(ligneDebut + i, colonne).Value = arrivees(i)
        If arrivees(i) <> "Aucune arrivée aujourd'hui" Then
            ws.Cells(ligneDebut + i, colonne).Font.Color = RGB(0, 128, 0)
        End If
    Next i
End Sub

' ========================================
' AFFICHER DEPARTS DU JOUR
' ========================================
Sub AfficherDepartsDuJour(ws As Worksheet, ligneDebut As Long, colonne As Long)
    Dim departs As Variant
    Dim i As Integer
    
    departs = ObtenirDepartsDuJour()
    
    For i = 0 To UBound(departs)
        ws.Cells(ligneDebut + i, colonne).Value = departs(i)
        If departs(i) <> "Aucun départ aujourd'hui" Then
            ws.Cells(ligneDebut + i, colonne).Font.Color = RGB(255, 0, 0)
        End If
    Next i
End Sub

' ========================================
' AFFICHER STATISTIQUES
' ========================================
Sub AfficherStatistiques(ws As Worksheet, ligneDebut As Long, colonne As Long)
    Dim ligne As Long
    ligne = ligneDebut
    
    ' Chambres libres
    ws.Cells(ligne, colonne).Value = "Chambres libres :"
    ws.Cells(ligne, colonne + 1).Value = CompterChambresLibres()
    ws.Cells(ligne, colonne).Font.Bold = True
    ligne = ligne + 1
    
    ' Chambres occupées
    ws.Cells(ligne, colonne).Value = "Chambres occupées :"
    ws.Cells(ligne, colonne + 1).Value = CompterChambresOccupees()
    ws.Cells(ligne, colonne).Font.Bold = True
    ligne = ligne + 1
    
    ' Taux d'occupation
    ws.Cells(ligne, colonne).Value = "Taux d'occupation :"
    ws.Cells(ligne, colonne + 1).Value = Format(CalculerTauxOccupationJour(), "0.00") & "%"
    ws.Cells(ligne, colonne).Font.Bold = True
    ligne = ligne + 2
    
    ' Chiffre d'affaires du mois
    ws.Cells(ligne, colonne).Value = "CA du mois :"
    ws.Cells(ligne, colonne + 1).Value = Format(CalculerChiffreAffaires(DateSerial(Year(Date), Month(Date), 1), Date), "0.00") & "€"
    ws.Cells(ligne, colonne).Font.Bold = True
End Sub

' ========================================
' COMPTER CHAMBRES LIBRES
' ========================================
Function CompterChambresLibres() As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    Dim compteur As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CHAMBRES)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    compteur = 0
    For i = 2 To derniereLigne
        If ws.Cells(i, 4).Value = "Libre" Then
            compteur = compteur + 1
        End If
    Next i
    
    CompterChambresLibres = compteur
End Function

' ========================================
' COMPTER CHAMBRES OCCUPEES
' ========================================
Function CompterChambresOccupees() As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    Dim compteur As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CHAMBRES)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    compteur = 0
    For i = 2 To derniereLigne
        If ws.Cells(i, 4).Value = "Occupée" Then
            compteur = compteur + 1
        End If
    Next i
    
    CompterChambresOccupees = compteur
End Function

' ========================================
' CALCULER TAUX OCCUPATION DU JOUR
' ========================================
Function CalculerTauxOccupationJour() As Double
    Dim totalChambres As Long
    Dim chambresOccupees As Long
    
    totalChambres = CompterNombreChambres()
    chambresOccupees = CompterChambresOccupees()
    
    If totalChambres > 0 Then
        CalculerTauxOccupationJour = (chambresOccupees / totalChambres) * 100
    Else
        CalculerTauxOccupationJour = 0
    End If
End Function

' ========================================
' CREER BOUTONS DASHBOARD
' ========================================
Sub CreerBoutonsDashboard(ws As Worksheet)
    ' Note: En VBA Excel, les boutons sont généralement créés via l'interface
    ' Ici nous créons des "boutons" textuels avec formatage
    
    Dim ligne As Long
    ligne = 20
    
    ' Bouton Nouvelle Réservation
    ws.Cells(ligne, 1).Value = "➕ Nouvelle Réservation"
    ws.Cells(ligne, 1).Font.Bold = True
    ws.Cells(ligne, 1).Interior.Color = RGB(146, 208, 80)
    ws.Cells(ligne, 1).Font.Color = RGB(255, 255, 255)
    ligne = ligne + 1
    
    ' Bouton Gestion Chambres
    ws.Cells(ligne, 1).Value = "🛏️ Gestion Chambres"
    ws.Cells(ligne, 1).Font.Bold = True
    ws.Cells(ligne, 1).Interior.Color = RGB(68, 114, 196)
    ws.Cells(ligne, 1).Font.Color = RGB(255, 255, 255)
    ligne = ligne + 1
    
    ' Bouton Gestion Clients
    ws.Cells(ligne, 1).Value = "👤 Gestion Clients"
    ws.Cells(ligne, 1).Font.Bold = True
    ws.Cells(ligne, 1).Interior.Color = RGB(255, 192, 0)
    ws.Cells(ligne, 1).Font.Color = RGB(0, 0, 0)
    ligne = ligne + 1
    
    ' Bouton Rapports
    ws.Cells(ligne, 1).Value = "📊 Rapports"
    ws.Cells(ligne, 1).Font.Bold = True
    ws.Cells(ligne, 1).Interior.Color = RGB(112, 48, 160)
    ws.Cells(ligne, 1).Font.Color = RGB(255, 255, 255)
End Sub

' ========================================
' ACTUALISER DASHBOARD
' ========================================
Sub ActualiserDashboard()
    Call ConfigurerFeuilleDashboard(ThisWorkbook.Worksheets(FEUILLE_DASHBOARD))
    MsgBox "Dashboard actualisé !", vbInformation, APP_NAME
End Sub

' ========================================
' CONFIGURER FEUILLE RAPPORTS
' ========================================
Sub ConfigurerFeuilleRapports(ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "📊 RAPPORTS ET STATISTIQUES"
        .Cells(1, 1).Font.Size = 16
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Color = RGB(68, 114, 196)
        
        .Cells(3, 1).Value = "Utilisez les fonctions du module Rapports pour générer des statistiques."
        .Cells(4, 1).Value = "Exemple : GenererRapportMensuel(12, 2024) pour décembre 2024"
        
        .Columns("A:F").AutoFit
    End With
End Sub
