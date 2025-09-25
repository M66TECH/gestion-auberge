Attribute VB_Name = "ModulePrincipal"
' ========================================
' MODULE PRINCIPAL - GESTION AUBERGE
' ========================================
' Auteur: Application VBA Gestion Auberge
' Date: 2024
' Description: Module principal contenant les fonctions de base

Option Explicit

' Variables globales
Public Const APP_NAME As String = "Gestion Auberge v1.0"
Public Const FEUILLE_CHAMBRES As String = "Chambres"
Public Const FEUILLE_CLIENTS As String = "Clients"
Public Const FEUILLE_RESERVATIONS As String = "Reservations"
Public Const FEUILLE_PAIEMENTS As String = "Paiements"
Public Const FEUILLE_PARAMETRES As String = "Parametres"
Public Const FEUILLE_DASHBOARD As String = "Dashboard"
Public Const FEUILLE_RAPPORTS As String = "Rapports"

' ========================================
' INITIALISATION DE L'APPLICATION
' ========================================
Sub InitialiserApplication()
    On Error GoTo ErrHandler
    
    ' Masquer les alertes Excel
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' Vérifier et créer les feuilles si nécessaire
    Call VerifierStructureFeuilles
    
    ' Initialiser les données de base
    Call InitialiserDonneesBase
    
    ' Afficher le dashboard
    Call AfficherDashboard
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Application initialisée avec succès !", vbInformation, APP_NAME
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Erreur lors de l'initialisation : " & Err.Description, vbCritical, APP_NAME
End Sub

' ========================================
' VERIFICATION DE LA STRUCTURE
' ========================================
Sub VerifierStructureFeuilles()
    Dim ws As Worksheet
    Dim feuillesRequises As Variant
    Dim i As Integer
    
    feuillesRequises = Array(FEUILLE_CHAMBRES, FEUILLE_CLIENTS, FEUILLE_RESERVATIONS, _
                            FEUILLE_PAIEMENTS, FEUILLE_PARAMETRES, FEUILLE_DASHBOARD, FEUILLE_RAPPORTS)
    
    For i = 0 To UBound(feuillesRequises)
        If Not FeuilleExiste(feuillesRequises(i)) Then
            Set ws = ThisWorkbook.Worksheets.Add
            ws.Name = feuillesRequises(i)
            Call ConfigurerFeuille(ws)
        End If
    Next i
End Sub

' ========================================
' VERIFICATION EXISTENCE FEUILLE
' ========================================
Function FeuilleExiste(nomFeuille As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    FeuilleExiste = Not ws Is Nothing
    On Error GoTo 0
End Function

' ========================================
' CONFIGURATION DES FEUILLES
' ========================================
Sub ConfigurerFeuille(ws As Worksheet)
    Select Case ws.Name
        Case FEUILLE_CHAMBRES
            Call ConfigurerFeuilleChambre(ws)
        Case FEUILLE_CLIENTS
            Call ConfigurerFeuilleClients(ws)
        Case FEUILLE_RESERVATIONS
            Call ConfigurerFeuilleReservations(ws)
        Case FEUILLE_PAIEMENTS
            Call ConfigurerFeuillePaiements(ws)
        Case FEUILLE_PARAMETRES
            Call ConfigurerFeuilleParametres(ws)
        Case FEUILLE_DASHBOARD
            Call ConfigurerFeuilleDashboard(ws)
        Case FEUILLE_RAPPORTS
            Call ConfigurerFeuilleRapports(ws)
    End Select
End Sub

' ========================================
' CONFIGURATION FEUILLE CHAMBRES
' ========================================
Sub ConfigurerFeuilleChambre(ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "NumChambre"
        .Cells(1, 2).Value = "TypeChambre"
        .Cells(1, 3).Value = "TarifNuit"
        .Cells(1, 4).Value = "Statut"
        .Cells(1, 5).Value = "Description"
        .Cells(1, 6).Value = "Equipements"
        
        ' Formatage des en-têtes
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:F1").Font.Color = RGB(255, 255, 255)
        .Range("A1:F1").Borders.LineStyle = xlContinuous
        
        ' Ajustement automatique des colonnes
        .Columns("A:F").AutoFit
    End With
End Sub

' ========================================
' CONFIGURATION FEUILLE CLIENTS
' ========================================
Sub ConfigurerFeuilleClients(ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "IDClient"
        .Cells(1, 2).Value = "Nom"
        .Cells(1, 3).Value = "Prenom"
        .Cells(1, 4).Value = "Telephone"
        .Cells(1, 5).Value = "Email"
        .Cells(1, 6).Value = "Adresse"
        .Cells(1, 7).Value = "DateCreation"
        
        ' Formatage des en-têtes
        .Range("A1:G1").Font.Bold = True
        .Range("A1:G1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:G1").Font.Color = RGB(255, 255, 255)
        .Range("A1:G1").Borders.LineStyle = xlContinuous
        
        ' Ajustement automatique des colonnes
        .Columns("A:G").AutoFit
    End With
End Sub

' ========================================
' CONFIGURATION FEUILLE RESERVATIONS
' ========================================
Sub ConfigurerFeuilleReservations(ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "IDReservation"
        .Cells(1, 2).Value = "IDClient"
        .Cells(1, 3).Value = "NumChambre"
        .Cells(1, 4).Value = "DateArrivee"
        .Cells(1, 5).Value = "DateDepart"
        .Cells(1, 6).Value = "NbNuits"
        .Cells(1, 7).Value = "MontantTotal"
        .Cells(1, 8).Value = "Statut"
        .Cells(1, 9).Value = "DateReservation"
        .Cells(1, 10).Value = "Commentaires"
        
        ' Formatage des en-têtes
        .Range("A1:J1").Font.Bold = True
        .Range("A1:J1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:J1").Font.Color = RGB(255, 255, 255)
        .Range("A1:J1").Borders.LineStyle = xlContinuous
        
        ' Ajustement automatique des colonnes
        .Columns("A:J").AutoFit
    End With
End Sub

' ========================================
' CONFIGURATION FEUILLE PAIEMENTS
' ========================================
Sub ConfigurerFeuillePaiements(ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "IDPaiement"
        .Cells(1, 2).Value = "IDReservation"
        .Cells(1, 3).Value = "Montant"
        .Cells(1, 4).Value = "ModePaiement"
        .Cells(1, 5).Value = "DatePaiement"
        .Cells(1, 6).Value = "TypePaiement"
        .Cells(1, 7).Value = "Statut"
        
        ' Formatage des en-têtes
        .Range("A1:G1").Font.Bold = True
        .Range("A1:G1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:G1").Font.Color = RGB(255, 255, 255)
        .Range("A1:G1").Borders.LineStyle = xlContinuous
        
        ' Ajustement automatique des colonnes
        .Columns("A:G").AutoFit
    End With
End Sub

' ========================================
' CONFIGURATION FEUILLE PARAMETRES
' ========================================
Sub ConfigurerFeuilleParametres(ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "Parametre"
        .Cells(1, 2).Value = "Valeur"
        .Cells(1, 3).Value = "Description"
        
        ' Formatage des en-têtes
        .Range("A1:C1").Font.Bold = True
        .Range("A1:C1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:C1").Font.Color = RGB(255, 255, 255)
        .Range("A1:C1").Borders.LineStyle = xlContinuous
        
        ' Ajustement automatique des colonnes
        .Columns("A:C").AutoFit
    End With
End Sub

' ========================================
' AFFICHAGE DU DASHBOARD
' ========================================
Sub AfficherDashboard()
    ThisWorkbook.Worksheets(FEUILLE_DASHBOARD).Activate
End Sub

' ========================================
' INITIALISATION DES DONNEES DE BASE
' ========================================
Sub InitialiserDonneesBase()
    Call InitialiserParametres
    Call InitialiserChambresExemple
End Sub

' ========================================
' INITIALISATION DES PARAMETRES
' ========================================
Sub InitialiserParametres()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(FEUILLE_PARAMETRES)
    
    ' Vérifier si les paramètres existent déjà
    If ws.Cells(2, 1).Value = "" Then
        ws.Cells(2, 1).Value = "NomAuberge"
        ws.Cells(2, 2).Value = "Auberge du Bon Repos"
        ws.Cells(2, 3).Value = "Nom de l'établissement"
        
        ws.Cells(3, 1).Value = "AdresseAuberge"
        ws.Cells(3, 2).Value = "123 Rue de la Paix, 75000 Paris"
        ws.Cells(3, 3).Value = "Adresse complète"
        
        ws.Cells(4, 1).Value = "TelephoneAuberge"
        ws.Cells(4, 2).Value = "01 23 45 67 89"
        ws.Cells(4, 3).Value = "Numéro de téléphone"
        
        ws.Cells(5, 1).Value = "EmailAuberge"
        ws.Cells(5, 2).Value = "contact@auberge-bonrepos.fr"
        ws.Cells(5, 3).Value = "Adresse email"
        
        ws.Cells(6, 1).Value = "TauxTVA"
        ws.Cells(6, 2).Value = "10"
        ws.Cells(6, 3).Value = "Taux de TVA en pourcentage"
    End If
End Sub

' ========================================
' INITIALISATION CHAMBRES EXEMPLE
' ========================================
Sub InitialiserChambresExemple()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CHAMBRES)
    
    ' Vérifier si les chambres existent déjà
    If ws.Cells(2, 1).Value = "" Then
        ' Chambres simples
        ws.Cells(2, 1).Value = "101"
        ws.Cells(2, 2).Value = "Simple"
        ws.Cells(2, 3).Value = 65
        ws.Cells(2, 4).Value = "Libre"
        ws.Cells(2, 5).Value = "Chambre simple avec vue jardin"
        ws.Cells(2, 6).Value = "TV, WiFi, Salle de bain privée"
        
        ws.Cells(3, 1).Value = "102"
        ws.Cells(3, 2).Value = "Simple"
        ws.Cells(3, 3).Value = 65
        ws.Cells(3, 4).Value = "Libre"
        ws.Cells(3, 5).Value = "Chambre simple standard"
        ws.Cells(3, 6).Value = "TV, WiFi, Salle de bain privée"
        
        ' Chambres doubles
        ws.Cells(4, 1).Value = "201"
        ws.Cells(4, 2).Value = "Double"
        ws.Cells(4, 3).Value = 85
        ws.Cells(4, 4).Value = "Libre"
        ws.Cells(4, 5).Value = "Chambre double avec balcon"
        ws.Cells(4, 6).Value = "TV, WiFi, Balcon, Salle de bain privée"
        
        ws.Cells(5, 1).Value = "202"
        ws.Cells(5, 2).Value = "Double"
        ws.Cells(5, 3).Value = 85
        ws.Cells(5, 4).Value = "Libre"
        ws.Cells(5, 5).Value = "Chambre double standard"
        ws.Cells(5, 6).Value = "TV, WiFi, Salle de bain privée"
        
        ' Suites
        ws.Cells(6, 1).Value = "301"
        ws.Cells(6, 2).Value = "Suite"
        ws.Cells(6, 3).Value = 120
        ws.Cells(6, 4).Value = "Libre"
        ws.Cells(6, 5).Value = "Suite familiale avec salon"
        ws.Cells(6, 6).Value = "TV, WiFi, Salon, Balcon, Salle de bain privée"
    End If
End Sub
