Attribute VB_Name = "ModuleChambres"
' ========================================
' MODULE GESTION DES CHAMBRES
' ========================================
' Description: Fonctions pour la gestion des chambres

Option Explicit

' ========================================
' AJOUTER UNE NOUVELLE CHAMBRE
' ========================================
Function AjouterChambre(numChambre As String, typeChambre As String, tarif As Double, _
                       description As String, equipements As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim derniereLigne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CHAMBRES)
    
    ' Vérifier si la chambre existe déjà
    If ChambreExiste(numChambre) Then
        MsgBox "La chambre " & numChambre & " existe déjà !", vbExclamation, APP_NAME
        AjouterChambre = False
        Exit Function
    End If
    
    ' Trouver la dernière ligne
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Ajouter les données
    ws.Cells(derniereLigne, 1).Value = numChambre
    ws.Cells(derniereLigne, 2).Value = typeChambre
    ws.Cells(derniereLigne, 3).Value = tarif
    ws.Cells(derniereLigne, 4).Value = "Libre"
    ws.Cells(derniereLigne, 5).Value = description
    ws.Cells(derniereLigne, 6).Value = equipements
    
    ' Formatage
    ws.Range("A" & derniereLigne & ":F" & derniereLigne).Borders.LineStyle = xlContinuous
    ws.Columns("A:F").AutoFit
    
    AjouterChambre = True
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de l'ajout de la chambre : " & Err.Description, vbCritical, APP_NAME
    AjouterChambre = False
End Function

' ========================================
' VERIFIER SI UNE CHAMBRE EXISTE
' ========================================
Function ChambreExiste(numChambre As String) As Boolean
    Dim ws As Worksheet
    Dim plage As Range
    Dim cellule As Range
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CHAMBRES)
    Set plage = ws.Range("A2:A" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)
    
    For Each cellule In plage
        If cellule.Value = numChambre Then
            ChambreExiste = True
            Exit Function
        End If
    Next cellule
    
    ChambreExiste = False
End Function

' ========================================
' MODIFIER UNE CHAMBRE
' ========================================
Function ModifierChambre(numChambre As String, typeChambre As String, tarif As Double, _
                        statut As String, description As String, equipements As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim ligne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CHAMBRES)
    ligne = TrouverLigneChambre(numChambre)
    
    If ligne = 0 Then
        MsgBox "Chambre " & numChambre & " non trouvée !", vbExclamation, APP_NAME
        ModifierChambre = False
        Exit Function
    End If
    
    ' Modifier les données
    ws.Cells(ligne, 2).Value = typeChambre
    ws.Cells(ligne, 3).Value = tarif
    ws.Cells(ligne, 4).Value = statut
    ws.Cells(ligne, 5).Value = description
    ws.Cells(ligne, 6).Value = equipements
    
    ws.Columns("A:F").AutoFit
    
    ModifierChambre = True
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de la modification : " & Err.Description, vbCritical, APP_NAME
    ModifierChambre = False
End Function

' ========================================
' TROUVER LA LIGNE D'UNE CHAMBRE
' ========================================
Function TrouverLigneChambre(numChambre As String) As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CHAMBRES)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 1).Value = numChambre Then
            TrouverLigneChambre = i
            Exit Function
        End If
    Next i
    
    TrouverLigneChambre = 0
End Function

' ========================================
' CHANGER LE STATUT D'UNE CHAMBRE
' ========================================
Function ChangerStatutChambre(numChambre As String, nouveauStatut As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim ligne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CHAMBRES)
    ligne = TrouverLigneChambre(numChambre)
    
    If ligne = 0 Then
        MsgBox "Chambre " & numChambre & " non trouvée !", vbExclamation, APP_NAME
        ChangerStatutChambre = False
        Exit Function
    End If
    
    ' Valider le statut
    If nouveauStatut <> "Libre" And nouveauStatut <> "Occupée" And nouveauStatut <> "Maintenance" Then
        MsgBox "Statut invalide ! Utilisez : Libre, Occupée ou Maintenance", vbExclamation, APP_NAME
        ChangerStatutChambre = False
        Exit Function
    End If
    
    ws.Cells(ligne, 4).Value = nouveauStatut
    
    ' Colorier selon le statut
    Select Case nouveauStatut
        Case "Libre"
            ws.Range("A" & ligne & ":F" & ligne).Interior.Color = RGB(144, 238, 144) ' Vert clair
        Case "Occupée"
            ws.Range("A" & ligne & ":F" & ligne).Interior.Color = RGB(255, 182, 193) ' Rose clair
        Case "Maintenance"
            ws.Range("A" & ligne & ":F" & ligne).Interior.Color = RGB(255, 255, 0) ' Jaune
    End Select
    
    ChangerStatutChambre = True
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors du changement de statut : " & Err.Description, vbCritical, APP_NAME
    ChangerStatutChambre = False
End Function

' ========================================
' OBTENIR LA LISTE DES CHAMBRES LIBRES
' ========================================
Function ObtenirChambresLibres() As Variant
    Dim ws As Worksheet
    Dim chambresLibres() As String
    Dim i As Long, j As Long
    Dim derniereLigne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CHAMBRES)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ReDim chambresLibres(0 To 0)
    j = 0
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 4).Value = "Libre" Then
            ReDim Preserve chambresLibres(0 To j)
            chambresLibres(j) = ws.Cells(i, 1).Value & " - " & ws.Cells(i, 2).Value & " (" & ws.Cells(i, 3).Value & "€)"
            j = j + 1
        End If
    Next i
    
    If j = 0 Then
        ReDim chambresLibres(0 To 0)
        chambresLibres(0) = "Aucune chambre libre"
    End If
    
    ObtenirChambresLibres = chambresLibres
End Function

' ========================================
' OBTENIR LE TARIF D'UNE CHAMBRE
' ========================================
Function ObtenirTarifChambre(numChambre As String) As Double
    Dim ws As Worksheet
    Dim ligne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CHAMBRES)
    ligne = TrouverLigneChambre(numChambre)
    
    If ligne > 0 Then
        ObtenirTarifChambre = ws.Cells(ligne, 3).Value
    Else
        ObtenirTarifChambre = 0
    End If
End Function

' ========================================
' OBTENIR LE TYPE D'UNE CHAMBRE
' ========================================
Function ObtenirTypeChambre(numChambre As String) As String
    Dim ws As Worksheet
    Dim ligne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CHAMBRES)
    ligne = TrouverLigneChambre(numChambre)
    
    If ligne > 0 Then
        ObtenirTypeChambre = ws.Cells(ligne, 2).Value
    Else
        ObtenirTypeChambre = ""
    End If
End Function

' ========================================
' VERIFIER DISPONIBILITE CHAMBRE
' ========================================
Function ChambreDisponible(numChambre As String, dateArrivee As Date, dateDepart As Date) As Boolean
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    Dim dateArrRes As Date, dateDepartRes As Date
    
    ' Vérifier d'abord si la chambre existe et est libre
    If Not ChambreExiste(numChambre) Then
        ChambreDisponible = False
        Exit Function
    End If
    
    ' Vérifier le statut actuel
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CHAMBRES)
    If ws.Cells(TrouverLigneChambre(numChambre), 4).Value <> "Libre" Then
        ChambreDisponible = False
        Exit Function
    End If
    
    ' Vérifier les réservations existantes
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 3).Value = numChambre And ws.Cells(i, 8).Value = "Confirmée" Then
            dateArrRes = ws.Cells(i, 4).Value
            dateDepartRes = ws.Cells(i, 5).Value
            
            ' Vérifier s'il y a conflit de dates
            If Not (dateDepart <= dateArrRes Or dateArrivee >= dateDepartRes) Then
                ChambreDisponible = False
                Exit Function
            End If
        End If
    Next i
    
    ChambreDisponible = True
End Function

' ========================================
' SUPPRIMER UNE CHAMBRE
' ========================================
Function SupprimerChambre(numChambre As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim ligne As Long
    Dim reponse As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CHAMBRES)
    ligne = TrouverLigneChambre(numChambre)
    
    If ligne = 0 Then
        MsgBox "Chambre " & numChambre & " non trouvée !", vbExclamation, APP_NAME
        SupprimerChambre = False
        Exit Function
    End If
    
    ' Vérifier s'il y a des réservations actives
    If ChambreAReservations(numChambre) Then
        MsgBox "Impossible de supprimer la chambre " & numChambre & " car elle a des réservations actives !", _
               vbExclamation, APP_NAME
        SupprimerChambre = False
        Exit Function
    End If
    
    ' Demander confirmation
    reponse = MsgBox("Êtes-vous sûr de vouloir supprimer la chambre " & numChambre & " ?", _
                     vbYesNo + vbQuestion, APP_NAME)
    
    If reponse = vbYes Then
        ws.Rows(ligne).Delete
        SupprimerChambre = True
    Else
        SupprimerChambre = False
    End If
    
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de la suppression : " & Err.Description, vbCritical, APP_NAME
    SupprimerChambre = False
End Function

' ========================================
' VERIFIER SI UNE CHAMBRE A DES RESERVATIONS
' ========================================
Function ChambreAReservations(numChambre As String) As Boolean
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 3).Value = numChambre And _
           (ws.Cells(i, 8).Value = "Confirmée" Or ws.Cells(i, 8).Value = "En attente") Then
            ChambreAReservations = True
            Exit Function
        End If
    Next i
    
    ChambreAReservations = False
End Function
