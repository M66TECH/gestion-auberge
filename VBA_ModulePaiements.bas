Attribute VB_Name = "ModulePaiements"
' ========================================
' MODULE GESTION DES PAIEMENTS
' ========================================
' Description: Fonctions pour la gestion des paiements et facturation

Option Explicit

' ========================================
' ENREGISTRER UN PAIEMENT
' ========================================
Function EnregistrerPaiement(idReservation As Long, montant As Double, modePaiement As String, _
                            typePaiement As String) As Long
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim nouvelID As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_PAIEMENTS)
    
    ' Validations
    If montant <= 0 Then
        MsgBox "Le montant doit être supérieur à zéro !", vbExclamation, APP_NAME
        EnregistrerPaiement = 0
        Exit Function
    End If
    
    If Not ReservationExiste(idReservation) Then
        MsgBox "Réservation " & idReservation & " non trouvée !", vbExclamation, APP_NAME
        EnregistrerPaiement = 0
        Exit Function
    End If
    
    ' Vérifier que le montant ne dépasse pas le montant dû
    If MontantDejaPayé(idReservation) + montant > MontantTotalReservation(idReservation) Then
        MsgBox "Le montant total des paiements ne peut pas dépasser le montant de la réservation !", _
               vbExclamation, APP_NAME
        EnregistrerPaiement = 0
        Exit Function
    End If
    
    ' Générer un nouvel ID
    nouvelID = ObtenirProchainIDPaiement()
    
    ' Trouver la dernière ligne
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Ajouter les données
    ws.Cells(derniereLigne, 1).Value = nouvelID
    ws.Cells(derniereLigne, 2).Value = idReservation
    ws.Cells(derniereLigne, 3).Value = montant
    ws.Cells(derniereLigne, 4).Value = modePaiement
    ws.Cells(derniereLigne, 5).Value = Date
    ws.Cells(derniereLigne, 6).Value = typePaiement
    ws.Cells(derniereLigne, 7).Value = "Validé"
    
    ' Formatage
    ws.Range("A" & derniereLigne & ":G" & derniereLigne).Borders.LineStyle = xlContinuous
    ws.Columns("A:G").AutoFit
    
    EnregistrerPaiement = nouvelID
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de l'enregistrement du paiement : " & Err.Description, vbCritical, APP_NAME
    EnregistrerPaiement = 0
End Function

' ========================================
' OBTENIR LE PROCHAIN ID PAIEMENT
' ========================================
Function ObtenirProchainIDPaiement() As Long
    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim maxID As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_PAIEMENTS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    maxID = 0
    For i = 2 To derniereLigne
        If ws.Cells(i, 1).Value > maxID Then
            maxID = ws.Cells(i, 1).Value
        End If
    Next i
    
    ObtenirProchainIDPaiement = maxID + 1
End Function

' ========================================
' VERIFIER SI UNE RESERVATION EXISTE
' ========================================
Function ReservationExiste(idReservation As Long) As Boolean
    ReservationExiste = (TrouverLigneReservation(idReservation) > 0)
End Function

' ========================================
' CALCULER LE MONTANT DEJA PAYE
' ========================================
Function MontantDejaPayé(idReservation As Long) As Double
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    Dim total As Double
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_PAIEMENTS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    total = 0
    For i = 2 To derniereLigne
        If ws.Cells(i, 2).Value = idReservation And ws.Cells(i, 7).Value = "Validé" Then
            total = total + ws.Cells(i, 3).Value
        End If
    Next i
    
    MontantDejaPayé = total
End Function

' ========================================
' OBTENIR LE MONTANT TOTAL D'UNE RESERVATION
' ========================================
Function MontantTotalReservation(idReservation As Long) As Double
    Dim ws As Worksheet
    Dim ligne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    ligne = TrouverLigneReservation(idReservation)
    
    If ligne > 0 Then
        MontantTotalReservation = ws.Cells(ligne, 7).Value
    Else
        MontantTotalReservation = 0
    End If
End Function

' ========================================
' CALCULER LE MONTANT RESTANT A PAYER
' ========================================
Function MontantRestantAPayer(idReservation As Long) As Double
    MontantRestantAPayer = MontantTotalReservation(idReservation) - MontantDejaPayé(idReservation)
End Function

' ========================================
' VERIFIER SI UNE RESERVATION EST SOLDEE
' ========================================
Function ReservationSoldee(idReservation As Long) As Boolean
    ReservationSoldee = (MontantRestantAPayer(idReservation) <= 0)
End Function

' ========================================
' OBTENIR L'HISTORIQUE DES PAIEMENTS
' ========================================
Function ObtenirHistoriquePaiements(idReservation As Long) As Variant
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim derniereLigne As Long
    Dim paiements() As String
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_PAIEMENTS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ReDim paiements(0 To 0)
    j = 0
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 2).Value = idReservation Then
            ReDim Preserve paiements(0 To j)
            paiements(j) = Format(ws.Cells(i, 5).Value, "dd/mm/yyyy") & " - " & _
                          ws.Cells(i, 3).Value & "€ (" & ws.Cells(i, 4).Value & ") - " & _
                          ws.Cells(i, 6).Value & " [" & ws.Cells(i, 7).Value & "]"
            j = j + 1
        End If
    Next i
    
    If j = 0 Then
        ReDim paiements(0 To 0)
        paiements(0) = "Aucun paiement enregistré"
    End If
    
    ObtenirHistoriquePaiements = paiements
End Function

' ========================================
' ANNULER UN PAIEMENT
' ========================================
Function AnnulerPaiement(idPaiement As Long, motif As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim ligne As Long
    Dim reponse As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_PAIEMENTS)
    ligne = TrouverLignePaiement(idPaiement)
    
    If ligne = 0 Then
        MsgBox "Paiement " & idPaiement & " non trouvé !", vbExclamation, APP_NAME
        AnnulerPaiement = False
        Exit Function
    End If
    
    If ws.Cells(ligne, 7).Value = "Annulé" Then
        MsgBox "Ce paiement est déjà annulé !", vbInformation, APP_NAME
        AnnulerPaiement = True
        Exit Function
    End If
    
    ' Demander confirmation
    reponse = MsgBox("Êtes-vous sûr de vouloir annuler le paiement de " & _
                     ws.Cells(ligne, 3).Value & "€ ?", vbYesNo + vbQuestion, APP_NAME)
    
    If reponse = vbYes Then
        ws.Cells(ligne, 7).Value = "Annulé"
        ' Ajouter le motif dans une colonne de commentaires (si elle existe)
        
        ' Colorier la ligne en rouge
        ws.Range("A" & ligne & ":G" & ligne).Interior.Color = RGB(255, 182, 193)
        
        AnnulerPaiement = True
    Else
        AnnulerPaiement = False
    End If
    
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de l'annulation du paiement : " & Err.Description, vbCritical, APP_NAME
    AnnulerPaiement = False
End Function

' ========================================
' TROUVER LA LIGNE D'UN PAIEMENT
' ========================================
Function TrouverLignePaiement(idPaiement As Long) As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_PAIEMENTS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 1).Value = idPaiement Then
            TrouverLignePaiement = i
            Exit Function
        End If
    Next i
    
    TrouverLignePaiement = 0
End Function

' ========================================
' GENERER UNE FACTURE
' ========================================
Function GenererFacture(idReservation As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim wsFacture As Worksheet
    Dim wsReservation As Worksheet
    Dim wsClient As Worksheet
    Dim wsParametres As Worksheet
    Dim ligneRes As Long
    Dim clientInfo As Variant
    Dim nomFeuille As String
    
    ' Vérifier que la réservation existe
    If Not ReservationExiste(idReservation) Then
        MsgBox "Réservation " & idReservation & " non trouvée !", vbExclamation, APP_NAME
        GenererFacture = False
        Exit Function
    End If
    
    Set wsReservation = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    Set wsClient = ThisWorkbook.Worksheets(FEUILLE_CLIENTS)
    Set wsParametres = ThisWorkbook.Worksheets(FEUILLE_PARAMETRES)
    
    ligneRes = TrouverLigneReservation(idReservation)
    clientInfo = RechercherClientParID(wsReservation.Cells(ligneRes, 2).Value)
    
    ' Créer une nouvelle feuille pour la facture
    nomFeuille = "Facture_" & idReservation
    
    ' Supprimer la feuille si elle existe déjà
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(nomFeuille).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler
    
    Set wsFacture = ThisWorkbook.Worksheets.Add
    wsFacture.Name = nomFeuille
    
    ' En-tête de la facture
    With wsFacture
        .Cells(1, 1).Value = "FACTURE"
        .Cells(1, 1).Font.Size = 20
        .Cells(1, 1).Font.Bold = True
        
        .Cells(3, 1).Value = "Facture N° : " & idReservation
        .Cells(4, 1).Value = "Date : " & Format(Date, "dd/mm/yyyy")
        
        ' Informations de l'auberge
        .Cells(6, 1).Value = ObtenirParametre("NomAuberge")
        .Cells(7, 1).Value = ObtenirParametre("AdresseAuberge")
        .Cells(8, 1).Value = "Tél : " & ObtenirParametre("TelephoneAuberge")
        .Cells(9, 1).Value = "Email : " & ObtenirParametre("EmailAuberge")
        
        ' Informations du client
        .Cells(6, 5).Value = "FACTURÉ À :"
        .Cells(6, 5).Font.Bold = True
        .Cells(7, 5).Value = clientInfo(1) & " " & clientInfo(2)
        .Cells(8, 5).Value = clientInfo(5) ' Adresse
        .Cells(9, 5).Value = "Tél : " & clientInfo(3)
        .Cells(10, 5).Value = "Email : " & clientInfo(4)
        
        ' Détails de la réservation
        .Cells(12, 1).Value = "DÉTAILS DE LA RÉSERVATION"
        .Cells(12, 1).Font.Bold = True
        
        .Cells(14, 1).Value = "Description"
        .Cells(14, 2).Value = "Quantité"
        .Cells(14, 3).Value = "Prix unitaire"
        .Cells(14, 4).Value = "Total"
        .Range("A14:D14").Font.Bold = True
        .Range("A14:D14").Borders.LineStyle = xlContinuous
        
        .Cells(15, 1).Value = "Chambre " & wsReservation.Cells(ligneRes, 3).Value & " (" & _
                             ObtenirTypeChambre(wsReservation.Cells(ligneRes, 3).Value) & ")"
        .Cells(15, 2).Value = wsReservation.Cells(ligneRes, 6).Value & " nuit(s)"
        .Cells(15, 3).Value = ObtenirTarifChambre(wsReservation.Cells(ligneRes, 3).Value) & "€"
        .Cells(15, 4).Value = wsReservation.Cells(ligneRes, 7).Value & "€"
        
        .Cells(16, 1).Value = "Du " & Format(wsReservation.Cells(ligneRes, 4).Value, "dd/mm/yyyy") & _
                             " au " & Format(wsReservation.Cells(ligneRes, 5).Value, "dd/mm/yyyy")
        
        ' Totaux
        Dim montantHT As Double
        Dim tauxTVA As Double
        Dim montantTVA As Double
        Dim montantTTC As Double
        
        tauxTVA = Val(ObtenirParametre("TauxTVA")) / 100
        montantTTC = wsReservation.Cells(ligneRes, 7).Value
        montantHT = montantTTC / (1 + tauxTVA)
        montantTVA = montantTTC - montantHT
        
        .Cells(18, 3).Value = "Sous-total HT :"
        .Cells(18, 4).Value = Format(montantHT, "0.00") & "€"
        
        .Cells(19, 3).Value = "TVA (" & ObtenirParametre("TauxTVA") & "%) :"
        .Cells(19, 4).Value = Format(montantTVA, "0.00") & "€"
        
        .Cells(20, 3).Value = "TOTAL TTC :"
        .Cells(20, 4).Value = Format(montantTTC, "0.00") & "€"
        .Range("C20:D20").Font.Bold = True
        .Range("C18:D20").Borders.LineStyle = xlContinuous
        
        ' Paiements
        .Cells(22, 1).Value = "PAIEMENTS"
        .Cells(22, 1).Font.Bold = True
        
        Dim paiements As Variant
        Dim i As Integer
        paiements = ObtenirHistoriquePaiements(idReservation)
        
        For i = 0 To UBound(paiements)
            .Cells(23 + i, 1).Value = paiements(i)
        Next i
        
        Dim montantRestant As Double
        montantRestant = MontantRestantAPayer(idReservation)
        
        .Cells(25 + UBound(paiements), 1).Value = "SOLDE RESTANT : " & Format(montantRestant, "0.00") & "€"
        .Cells(25 + UBound(paiements), 1).Font.Bold = True
        
        If montantRestant <= 0 Then
            .Cells(25 + UBound(paiements), 1).Font.Color = RGB(0, 128, 0) ' Vert
        Else
            .Cells(25 + UBound(paiements), 1).Font.Color = RGB(255, 0, 0) ' Rouge
        End If
        
        ' Formatage général
        .Columns("A:F").AutoFit
        .Range("A1:F30").Font.Name = "Arial"
        .Range("A1:F30").Font.Size = 10
        
    End With
    
    ' Activer la feuille de facture
    wsFacture.Activate
    
    MsgBox "Facture générée avec succès dans la feuille '" & nomFeuille & "' !", vbInformation, APP_NAME
    GenererFacture = True
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de la génération de la facture : " & Err.Description, vbCritical, APP_NAME
    GenererFacture = False
End Function

' ========================================
' OBTENIR UN PARAMETRE
' ========================================
Function ObtenirParametre(nomParametre As String) As String
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_PARAMETRES)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 1).Value = nomParametre Then
            ObtenirParametre = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    
    ObtenirParametre = ""
End Function

' ========================================
' CALCULER LE CHIFFRE D'AFFAIRES
' ========================================
Function CalculerChiffreAffaires(dateDebut As Date, dateFin As Date) As Double
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    Dim total As Double
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_PAIEMENTS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    total = 0
    For i = 2 To derniereLigne
        If ws.Cells(i, 5).Value >= dateDebut And ws.Cells(i, 5).Value <= dateFin And _
           ws.Cells(i, 7).Value = "Validé" Then
            total = total + ws.Cells(i, 3).Value
        End If
    Next i
    
    CalculerChiffreAffaires = total
End Function

' ========================================
' OBTENIR RESERVATIONS NON SOLDEES
' ========================================
Function ObtenirReservationsNonSoldees() As Variant
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim derniereLigne As Long
    Dim reservations() As String
    Dim clientInfo As Variant
    Dim montantRestant As Double
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ReDim reservations(0 To 0)
    j = 0
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 8).Value = "Confirmée" Then
            montantRestant = MontantRestantAPayer(ws.Cells(i, 1).Value)
            If montantRestant > 0 Then
                clientInfo = RechercherClientParID(ws.Cells(i, 2).Value)
                ReDim Preserve reservations(0 To j)
                reservations(j) = "Rés. " & ws.Cells(i, 1).Value & " - " & clientInfo(1) & " " & clientInfo(2) & _
                                 " - Ch." & ws.Cells(i, 3).Value & " - Reste : " & Format(montantRestant, "0.00") & "€"
                j = j + 1
            End If
        End If
    Next i
    
    If j = 0 Then
        ReDim reservations(0 To 0)
        reservations(0) = "Toutes les réservations sont soldées"
    End If
    
    ObtenirReservationsNonSoldees = reservations
End Function
