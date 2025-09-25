Attribute VB_Name = "ModuleReservations"
' ========================================
' MODULE GESTION DES RESERVATIONS
' ========================================
' Description: Fonctions pour la gestion des réservations

Option Explicit

' ========================================
' CREER UNE NOUVELLE RESERVATION
' ========================================
Function CreerReservation(idClient As Long, numChambre As String, dateArrivee As Date, _
                         dateDepart As Date, commentaires As String) As Long
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim nouvelID As Long
    Dim nbNuits As Long
    Dim montantTotal As Double
    Dim tarifNuit As Double
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    
    ' Validations
    If dateArrivee >= dateDepart Then
        MsgBox "La date d'arrivée doit être antérieure à la date de départ !", vbExclamation, APP_NAME
        CreerReservation = 0
        Exit Function
    End If
    
    If dateArrivee < Date Then
        MsgBox "La date d'arrivée ne peut pas être dans le passé !", vbExclamation, APP_NAME
        CreerReservation = 0
        Exit Function
    End If
    
    ' Vérifier la disponibilité de la chambre
    If Not ChambreDisponible(numChambre, dateArrivee, dateDepart) Then
        MsgBox "La chambre " & numChambre & " n'est pas disponible pour ces dates !", vbExclamation, APP_NAME
        CreerReservation = 0
        Exit Function
    End If
    
    ' Calculer le nombre de nuits et le montant
    nbNuits = dateDepart - dateArrivee
    tarifNuit = ObtenirTarifChambre(numChambre)
    montantTotal = nbNuits * tarifNuit
    
    ' Générer un nouvel ID
    nouvelID = ObtenirProchainIDReservation()
    
    ' Trouver la dernière ligne
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Ajouter les données
    ws.Cells(derniereLigne, 1).Value = nouvelID
    ws.Cells(derniereLigne, 2).Value = idClient
    ws.Cells(derniereLigne, 3).Value = numChambre
    ws.Cells(derniereLigne, 4).Value = dateArrivee
    ws.Cells(derniereLigne, 5).Value = dateDepart
    ws.Cells(derniereLigne, 6).Value = nbNuits
    ws.Cells(derniereLigne, 7).Value = montantTotal
    ws.Cells(derniereLigne, 8).Value = "En attente"
    ws.Cells(derniereLigne, 9).Value = Date
    ws.Cells(derniereLigne, 10).Value = commentaires
    
    ' Formatage
    ws.Range("A" & derniereLigne & ":J" & derniereLigne).Borders.LineStyle = xlContinuous
    ws.Columns("A:J").AutoFit
    
    CreerReservation = nouvelID
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de la création de la réservation : " & Err.Description, vbCritical, APP_NAME
    CreerReservation = 0
End Function

' ========================================
' OBTENIR LE PROCHAIN ID RESERVATION
' ========================================
Function ObtenirProchainIDReservation() As Long
    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim maxID As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    maxID = 0
    For i = 2 To derniereLigne
        If ws.Cells(i, 1).Value > maxID Then
            maxID = ws.Cells(i, 1).Value
        End If
    Next i
    
    ObtenirProchainIDReservation = maxID + 1
End Function

' ========================================
' CONFIRMER UNE RESERVATION
' ========================================
Function ConfirmerReservation(idReservation As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim ligne As Long
    Dim numChambre As String
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    ligne = TrouverLigneReservation(idReservation)
    
    If ligne = 0 Then
        MsgBox "Réservation " & idReservation & " non trouvée !", vbExclamation, APP_NAME
        ConfirmerReservation = False
        Exit Function
    End If
    
    ' Vérifier le statut actuel
    If ws.Cells(ligne, 8).Value = "Confirmée" Then
        MsgBox "Cette réservation est déjà confirmée !", vbInformation, APP_NAME
        ConfirmerReservation = True
        Exit Function
    End If
    
    If ws.Cells(ligne, 8).Value = "Annulée" Then
        MsgBox "Impossible de confirmer une réservation annulée !", vbExclamation, APP_NAME
        ConfirmerReservation = False
        Exit Function
    End If
    
    ' Changer le statut
    ws.Cells(ligne, 8).Value = "Confirmée"
    
    ' Mettre à jour le statut de la chambre si c'est pour aujourd'hui
    numChambre = ws.Cells(ligne, 3).Value
    If ws.Cells(ligne, 4).Value <= Date And ws.Cells(ligne, 5).Value > Date Then
        Call ChangerStatutChambre(numChambre, "Occupée")
    End If
    
    ' Colorier la ligne en vert
    ws.Range("A" & ligne & ":J" & ligne).Interior.Color = RGB(144, 238, 144)
    
    ConfirmerReservation = True
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de la confirmation : " & Err.Description, vbCritical, APP_NAME
    ConfirmerReservation = False
End Function

' ========================================
' ANNULER UNE RESERVATION
' ========================================
Function AnnulerReservation(idReservation As Long, motif As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim ligne As Long
    Dim numChambre As String
    Dim reponse As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    ligne = TrouverLigneReservation(idReservation)
    
    If ligne = 0 Then
        MsgBox "Réservation " & idReservation & " non trouvée !", vbExclamation, APP_NAME
        AnnulerReservation = False
        Exit Function
    End If
    
    ' Vérifier le statut actuel
    If ws.Cells(ligne, 8).Value = "Annulée" Then
        MsgBox "Cette réservation est déjà annulée !", vbInformation, APP_NAME
        AnnulerReservation = True
        Exit Function
    End If
    
    ' Demander confirmation
    reponse = MsgBox("Êtes-vous sûr de vouloir annuler la réservation " & idReservation & " ?", _
                     vbYesNo + vbQuestion, APP_NAME)
    
    If reponse = vbNo Then
        AnnulerReservation = False
        Exit Function
    End If
    
    ' Changer le statut
    ws.Cells(ligne, 8).Value = "Annulée"
    ws.Cells(ligne, 10).Value = ws.Cells(ligne, 10).Value & " [ANNULÉE: " & motif & "]"
    
    ' Libérer la chambre si elle était occupée
    numChambre = ws.Cells(ligne, 3).Value
    Call ChangerStatutChambre(numChambre, "Libre")
    
    ' Colorier la ligne en rouge
    ws.Range("A" & ligne & ":J" & ligne).Interior.Color = RGB(255, 182, 193)
    
    AnnulerReservation = True
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de l'annulation : " & Err.Description, vbCritical, APP_NAME
    AnnulerReservation = False
End Function

' ========================================
' TROUVER LA LIGNE D'UNE RESERVATION
' ========================================
Function TrouverLigneReservation(idReservation As Long) As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 1).Value = idReservation Then
            TrouverLigneReservation = i
            Exit Function
        End If
    Next i
    
    TrouverLigneReservation = 0
End Function

' ========================================
' RECHERCHER RESERVATIONS PAR CLIENT
' ========================================
Function RechercherReservationsParClient(idClient As Long) As Variant
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim derniereLigne As Long
    Dim reservations() As String
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ReDim reservations(0 To 0)
    j = 0
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 2).Value = idClient Then
            ReDim Preserve reservations(0 To j)
            reservations(j) = "Rés. " & ws.Cells(i, 1).Value & " - Ch." & ws.Cells(i, 3).Value & _
                             " du " & Format(ws.Cells(i, 4).Value, "dd/mm/yyyy") & _
                             " au " & Format(ws.Cells(i, 5).Value, "dd/mm/yyyy") & _
                             " (" & ws.Cells(i, 8).Value & ") - " & ws.Cells(i, 7).Value & "€"
            j = j + 1
        End If
    Next i
    
    If j = 0 Then
        ReDim reservations(0 To 0)
        reservations(0) = "Aucune réservation trouvée"
    End If
    
    RechercherReservationsParClient = reservations
End Function

' ========================================
' RECHERCHER RESERVATIONS PAR DATE
' ========================================
Function RechercherReservationsParDate(dateRecherche As Date) As Variant
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim derniereLigne As Long
    Dim reservations() As String
    Dim clientInfo As Variant
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ReDim reservations(0 To 0)
    j = 0
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 4).Value <= dateRecherche And ws.Cells(i, 5).Value > dateRecherche And _
           ws.Cells(i, 8).Value = "Confirmée" Then
            
            clientInfo = RechercherClientParID(ws.Cells(i, 2).Value)
            ReDim Preserve reservations(0 To j)
            reservations(j) = "Ch." & ws.Cells(i, 3).Value & " - " & clientInfo(1) & " " & clientInfo(2) & _
                             " (Rés. " & ws.Cells(i, 1).Value & ")"
            j = j + 1
        End If
    Next i
    
    If j = 0 Then
        ReDim reservations(0 To 0)
        reservations(0) = "Aucune réservation pour cette date"
    End If
    
    RechercherReservationsParDate = reservations
End Function

' ========================================
' OBTENIR RESERVATIONS DU JOUR
' ========================================
Function ObtenirReservationsDuJour() As Variant
    ObtenirReservationsDuJour = RechercherReservationsParDate(Date)
End Function

' ========================================
' OBTENIR ARRIVEES DU JOUR
' ========================================
Function ObtenirArriveesDuJour() As Variant
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim derniereLigne As Long
    Dim arrivees() As String
    Dim clientInfo As Variant
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ReDim arrivees(0 To 0)
    j = 0
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 4).Value = Date And ws.Cells(i, 8).Value = "Confirmée" Then
            clientInfo = RechercherClientParID(ws.Cells(i, 2).Value)
            ReDim Preserve arrivees(0 To j)
            arrivees(j) = "Ch." & ws.Cells(i, 3).Value & " - " & clientInfo(1) & " " & clientInfo(2) & _
                         " (Rés. " & ws.Cells(i, 1).Value & ") - " & ws.Cells(i, 6).Value & " nuit(s)"
            j = j + 1
        End If
    Next i
    
    If j = 0 Then
        ReDim arrivees(0 To 0)
        arrivees(0) = "Aucune arrivée aujourd'hui"
    End If
    
    ObtenirArriveesDuJour = arrivees
End Function

' ========================================
' OBTENIR DEPARTS DU JOUR
' ========================================
Function ObtenirDepartsDuJour() As Variant
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim derniereLigne As Long
    Dim departs() As String
    Dim clientInfo As Variant
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ReDim departs(0 To 0)
    j = 0
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 5).Value = Date And ws.Cells(i, 8).Value = "Confirmée" Then
            clientInfo = RechercherClientParID(ws.Cells(i, 2).Value)
            ReDim Preserve departs(0 To j)
            departs(j) = "Ch." & ws.Cells(i, 3).Value & " - " & clientInfo(1) & " " & clientInfo(2) & _
                        " (Rés. " & ws.Cells(i, 1).Value & ") - " & ws.Cells(i, 7).Value & "€"
            j = j + 1
        End If
    Next i
    
    If j = 0 Then
        ReDim departs(0 To 0)
        departs(0) = "Aucun départ aujourd'hui"
    End If
    
    ObtenirDepartsDuJour = departs
End Function

' ========================================
' MODIFIER UNE RESERVATION
' ========================================
Function ModifierReservation(idReservation As Long, dateArrivee As Date, dateDepart As Date, _
                            commentaires As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim ligne As Long
    Dim numChambre As String
    Dim nbNuits As Long
    Dim montantTotal As Double
    Dim tarifNuit As Double
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    ligne = TrouverLigneReservation(idReservation)
    
    If ligne = 0 Then
        MsgBox "Réservation " & idReservation & " non trouvée !", vbExclamation, APP_NAME
        ModifierReservation = False
        Exit Function
    End If
    
    ' Vérifications
    If ws.Cells(ligne, 8).Value = "Annulée" Then
        MsgBox "Impossible de modifier une réservation annulée !", vbExclamation, APP_NAME
        ModifierReservation = False
        Exit Function
    End If
    
    If dateArrivee >= dateDepart Then
        MsgBox "La date d'arrivée doit être antérieure à la date de départ !", vbExclamation, APP_NAME
        ModifierReservation = False
        Exit Function
    End If
    
    ' Recalculer les valeurs
    numChambre = ws.Cells(ligne, 3).Value
    nbNuits = dateDepart - dateArrivee
    tarifNuit = ObtenirTarifChambre(numChambre)
    montantTotal = nbNuits * tarifNuit
    
    ' Modifier les données
    ws.Cells(ligne, 4).Value = dateArrivee
    ws.Cells(ligne, 5).Value = dateDepart
    ws.Cells(ligne, 6).Value = nbNuits
    ws.Cells(ligne, 7).Value = montantTotal
    ws.Cells(ligne, 10).Value = commentaires
    
    ModifierReservation = True
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de la modification : " & Err.Description, vbCritical, APP_NAME
    ModifierReservation = False
End Function

' ========================================
' EFFECTUER CHECK-IN
' ========================================
Function EffectuerCheckIn(idReservation As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim ligne As Long
    Dim numChambre As String
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    ligne = TrouverLigneReservation(idReservation)
    
    If ligne = 0 Then
        MsgBox "Réservation " & idReservation & " non trouvée !", vbExclamation, APP_NAME
        EffectuerCheckIn = False
        Exit Function
    End If
    
    ' Vérifications
    If ws.Cells(ligne, 8).Value <> "Confirmée" Then
        MsgBox "La réservation doit être confirmée pour effectuer le check-in !", vbExclamation, APP_NAME
        EffectuerCheckIn = False
        Exit Function
    End If
    
    If ws.Cells(ligne, 4).Value > Date Then
        MsgBox "Il est trop tôt pour effectuer le check-in !", vbExclamation, APP_NAME
        EffectuerCheckIn = False
        Exit Function
    End If
    
    ' Marquer la chambre comme occupée
    numChambre = ws.Cells(ligne, 3).Value
    Call ChangerStatutChambre(numChambre, "Occupée")
    
    ' Ajouter une note dans les commentaires
    ws.Cells(ligne, 10).Value = ws.Cells(ligne, 10).Value & " [CHECK-IN: " & Format(Now, "dd/mm/yyyy hh:mm") & "]"
    
    MsgBox "Check-in effectué avec succès pour la chambre " & numChambre & " !", vbInformation, APP_NAME
    EffectuerCheckIn = True
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors du check-in : " & Err.Description, vbCritical, APP_NAME
    EffectuerCheckIn = False
End Function

' ========================================
' EFFECTUER CHECK-OUT
' ========================================
Function EffectuerCheckOut(idReservation As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim ligne As Long
    Dim numChambre As String
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    ligne = TrouverLigneReservation(idReservation)
    
    If ligne = 0 Then
        MsgBox "Réservation " & idReservation & " non trouvée !", vbExclamation, APP_NAME
        EffectuerCheckOut = False
        Exit Function
    End If
    
    ' Libérer la chambre
    numChambre = ws.Cells(ligne, 3).Value
    Call ChangerStatutChambre(numChambre, "Libre")
    
    ' Ajouter une note dans les commentaires
    ws.Cells(ligne, 10).Value = ws.Cells(ligne, 10).Value & " [CHECK-OUT: " & Format(Now, "dd/mm/yyyy hh:mm") & "]"
    
    MsgBox "Check-out effectué avec succès pour la chambre " & numChambre & " !", vbInformation, APP_NAME
    EffectuerCheckOut = True
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors du check-out : " & Err.Description, vbCritical, APP_NAME
    EffectuerCheckOut = False
End Function
