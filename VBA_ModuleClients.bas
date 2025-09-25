Attribute VB_Name = "ModuleClients"
' ========================================
' MODULE GESTION DES CLIENTS
' ========================================
' Description: Fonctions pour la gestion des clients

Option Explicit

' ========================================
' AJOUTER UN NOUVEAU CLIENT
' ========================================
Function AjouterClient(nom As String, prenom As String, telephone As String, _
                      email As String, adresse As String) As Long
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim nouvelID As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CLIENTS)
    
    ' Vérifier si le client existe déjà
    If ClientExiste(nom, prenom) Then
        MsgBox "Un client avec ce nom et prénom existe déjà !", vbExclamation, APP_NAME
        AjouterClient = 0
        Exit Function
    End If
    
    ' Trouver la dernière ligne et générer un nouvel ID
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    nouvelID = ObtenirProchainIDClient()
    
    ' Ajouter les données
    ws.Cells(derniereLigne, 1).Value = nouvelID
    ws.Cells(derniereLigne, 2).Value = nom
    ws.Cells(derniereLigne, 3).Value = prenom
    ws.Cells(derniereLigne, 4).Value = telephone
    ws.Cells(derniereLigne, 5).Value = email
    ws.Cells(derniereLigne, 6).Value = adresse
    ws.Cells(derniereLigne, 7).Value = Date
    
    ' Formatage
    ws.Range("A" & derniereLigne & ":G" & derniereLigne).Borders.LineStyle = xlContinuous
    ws.Columns("A:G").AutoFit
    
    AjouterClient = nouvelID
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de l'ajout du client : " & Err.Description, vbCritical, APP_NAME
    AjouterClient = 0
End Function

' ========================================
' OBTENIR LE PROCHAIN ID CLIENT
' ========================================
Function ObtenirProchainIDClient() As Long
    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim maxID As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CLIENTS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    maxID = 0
    For i = 2 To derniereLigne
        If ws.Cells(i, 1).Value > maxID Then
            maxID = ws.Cells(i, 1).Value
        End If
    Next i
    
    ObtenirProchainIDClient = maxID + 1
End Function

' ========================================
' VERIFIER SI UN CLIENT EXISTE
' ========================================
Function ClientExiste(nom As String, prenom As String) As Boolean
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CLIENTS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To derniereLigne
        If UCase(ws.Cells(i, 2).Value) = UCase(nom) And _
           UCase(ws.Cells(i, 3).Value) = UCase(prenom) Then
            ClientExiste = True
            Exit Function
        End If
    Next i
    
    ClientExiste = False
End Function

' ========================================
' RECHERCHER UN CLIENT PAR ID
' ========================================
Function RechercherClientParID(idClient As Long) As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    Dim clientInfo(6) As Variant
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CLIENTS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 1).Value = idClient Then
            clientInfo(0) = ws.Cells(i, 1).Value ' ID
            clientInfo(1) = ws.Cells(i, 2).Value ' Nom
            clientInfo(2) = ws.Cells(i, 3).Value ' Prénom
            clientInfo(3) = ws.Cells(i, 4).Value ' Téléphone
            clientInfo(4) = ws.Cells(i, 5).Value ' Email
            clientInfo(5) = ws.Cells(i, 6).Value ' Adresse
            clientInfo(6) = ws.Cells(i, 7).Value ' Date création
            
            RechercherClientParID = clientInfo
            Exit Function
        End If
    Next i
    
    ' Client non trouvé
    RechercherClientParID = Empty
End Function

' ========================================
' RECHERCHER CLIENTS PAR NOM
' ========================================
Function RechercherClientsParNom(nomRecherche As String) As Variant
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim derniereLigne As Long
    Dim clients() As String
    Dim nomComplet As String
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CLIENTS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ReDim clients(0 To 0)
    j = 0
    
    For i = 2 To derniereLigne
        nomComplet = ws.Cells(i, 2).Value & " " & ws.Cells(i, 3).Value
        If InStr(1, UCase(nomComplet), UCase(nomRecherche)) > 0 Then
            ReDim Preserve clients(0 To j)
            clients(j) = ws.Cells(i, 1).Value & " - " & nomComplet & " (" & ws.Cells(i, 4).Value & ")"
            j = j + 1
        End If
    Next i
    
    If j = 0 Then
        ReDim clients(0 To 0)
        clients(0) = "Aucun client trouvé"
    End If
    
    RechercherClientsParNom = clients
End Function

' ========================================
' MODIFIER UN CLIENT
' ========================================
Function ModifierClient(idClient As Long, nom As String, prenom As String, _
                       telephone As String, email As String, adresse As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim ligne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CLIENTS)
    ligne = TrouverLigneClient(idClient)
    
    If ligne = 0 Then
        MsgBox "Client avec ID " & idClient & " non trouvé !", vbExclamation, APP_NAME
        ModifierClient = False
        Exit Function
    End If
    
    ' Modifier les données
    ws.Cells(ligne, 2).Value = nom
    ws.Cells(ligne, 3).Value = prenom
    ws.Cells(ligne, 4).Value = telephone
    ws.Cells(ligne, 5).Value = email
    ws.Cells(ligne, 6).Value = adresse
    
    ws.Columns("A:G").AutoFit
    
    ModifierClient = True
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de la modification : " & Err.Description, vbCritical, APP_NAME
    ModifierClient = False
End Function

' ========================================
' TROUVER LA LIGNE D'UN CLIENT
' ========================================
Function TrouverLigneClient(idClient As Long) As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CLIENTS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 1).Value = idClient Then
            TrouverLigneClient = i
            Exit Function
        End If
    Next i
    
    TrouverLigneClient = 0
End Function

' ========================================
' SUPPRIMER UN CLIENT
' ========================================
Function SupprimerClient(idClient As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim ligne As Long
    Dim reponse As VbMsgBoxResult
    Dim clientInfo As Variant
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CLIENTS)
    ligne = TrouverLigneClient(idClient)
    
    If ligne = 0 Then
        MsgBox "Client avec ID " & idClient & " non trouvé !", vbExclamation, APP_NAME
        SupprimerClient = False
        Exit Function
    End If
    
    ' Vérifier s'il y a des réservations actives
    If ClientAReservations(idClient) Then
        MsgBox "Impossible de supprimer le client car il a des réservations actives !", _
               vbExclamation, APP_NAME
        SupprimerClient = False
        Exit Function
    End If
    
    ' Obtenir les infos du client pour confirmation
    clientInfo = RechercherClientParID(idClient)
    
    ' Demander confirmation
    reponse = MsgBox("Êtes-vous sûr de vouloir supprimer le client " & _
                     clientInfo(1) & " " & clientInfo(2) & " ?", _
                     vbYesNo + vbQuestion, APP_NAME)
    
    If reponse = vbYes Then
        ws.Rows(ligne).Delete
        SupprimerClient = True
    Else
        SupprimerClient = False
    End If
    
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de la suppression : " & Err.Description, vbCritical, APP_NAME
    SupprimerClient = False
End Function

' ========================================
' VERIFIER SI UN CLIENT A DES RESERVATIONS
' ========================================
Function ClientAReservations(idClient As Long) As Boolean
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 2).Value = idClient And _
           (ws.Cells(i, 8).Value = "Confirmée" Or ws.Cells(i, 8).Value = "En attente") Then
            ClientAReservations = True
            Exit Function
        End If
    Next i
    
    ClientAReservations = False
End Function

' ========================================
' OBTENIR L'HISTORIQUE D'UN CLIENT
' ========================================
Function ObtenirHistoriqueClient(idClient As Long) As Variant
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim derniereLigne As Long
    Dim historique() As String
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ReDim historique(0 To 0)
    j = 0
    
    For i = 2 To derniereLigne
        If ws.Cells(i, 2).Value = idClient Then
            ReDim Preserve historique(0 To j)
            historique(j) = "Rés. " & ws.Cells(i, 1).Value & " - Ch." & ws.Cells(i, 3).Value & _
                           " du " & Format(ws.Cells(i, 4).Value, "dd/mm/yyyy") & _
                           " au " & Format(ws.Cells(i, 5).Value, "dd/mm/yyyy") & _
                           " (" & ws.Cells(i, 8).Value & ")"
            j = j + 1
        End If
    Next i
    
    If j = 0 Then
        ReDim historique(0 To 0)
        historique(0) = "Aucun historique"
    End If
    
    ObtenirHistoriqueClient = historique
End Function

' ========================================
' OBTENIR LA LISTE DE TOUS LES CLIENTS
' ========================================
Function ObtenirListeClients() As Variant
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim derniereLigne As Long
    Dim clients() As String
    
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CLIENTS)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If derniereLigne < 2 Then
        ReDim clients(0 To 0)
        clients(0) = "Aucun client enregistré"
        ObtenirListeClients = clients
        Exit Function
    End If
    
    ReDim clients(0 To derniereLigne - 2)
    j = 0
    
    For i = 2 To derniereLigne
        clients(j) = ws.Cells(i, 1).Value & " - " & ws.Cells(i, 2).Value & " " & _
                     ws.Cells(i, 3).Value & " (" & ws.Cells(i, 4).Value & ")"
        j = j + 1
    Next i
    
    ObtenirListeClients = clients
End Function

' ========================================
' VALIDER EMAIL
' ========================================
Function ValiderEmail(email As String) As Boolean
    If Len(email) = 0 Then
        ValiderEmail = True ' Email optionnel
        Exit Function
    End If
    
    If InStr(email, "@") > 0 And InStr(email, ".") > 0 And Len(email) > 5 Then
        ValiderEmail = True
    Else
        ValiderEmail = False
    End If
End Function

' ========================================
' VALIDER TELEPHONE
' ========================================
Function ValiderTelephone(telephone As String) As Boolean
    Dim i As Integer
    Dim caractere As String
    
    If Len(telephone) < 10 Then
        ValiderTelephone = False
        Exit Function
    End If
    
    For i = 1 To Len(telephone)
        caractere = Mid(telephone, i, 1)
        If Not (IsNumeric(caractere) Or caractere = " " Or caractere = "." Or _
                caractere = "-" Or caractere = "(" Or caractere = ")") Then
            ValiderTelephone = False
            Exit Function
        End If
    Next i
    
    ValiderTelephone = True
End Function
