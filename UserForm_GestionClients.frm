VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_GestionClients 
   Caption         =   "👤 Gestion des Clients"
   ClientHeight    =   550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   750
   OleObjectBlob   =   "UserForm_GestionClients.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_GestionClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ========================================
' USERFORM GESTION CLIENTS - DESIGN MODERNE
' ========================================

Private Sub UserForm_Initialize()
    Call InitialiserInterface
End Sub

Private Sub InitialiserInterface()
    Call AppliquerStyleModerne(Me)
    
    Me.Width = 750
    Me.Height = 550
    Me.Caption = "👤 Gestion des Clients"
    
    Call CreerEnTeteClients
    Call CreerFormulaireClient
    Call CreerRechercheClient
    Call CreerHistoriqueClient
    Call CreerBarreActionsClients
End Sub

Private Sub CreerEnTeteClients()
    ' Titre principal
    Dim lblTitre As MSForms.Label
    Set lblTitre = Me.Controls.Add("Forms.Label.1", "lblTitreClients")
    
    With lblTitre
        .Top = 15
        .Left = 20
        .Width = 710
        .Height = 35
        .Caption = "👤 GESTION DES CLIENTS"
        .TextAlign = fmTextAlignCenter
    End With
    Call StylerTitre(lblTitre)
    
    ' Statistiques rapides
    Dim lblStats As MSForms.Label
    Set lblStats = Me.Controls.Add("Forms.Label.1", "lblStatsClients")
    
    With lblStats
        .Top = 55
        .Left = 20
        .Width = 710
        .Height = 20
        .Caption = "📊 Total clients : " & CompterTotalClients() & " | Nouveaux ce mois : " & CompterNouveauxClients()
        .TextAlign = fmTextAlignCenter
    End With
    Call StylerLabelSecondaire(lblStats)
End Sub

Private Sub CreerFormulaireClient()
    ' Panneau formulaire client
    Dim panneauForm As MSForms.Frame
    Set panneauForm = Me.Controls.Add("Forms.Frame.1", "panneauFormulaireClient")
    
    With panneauForm
        .Top = 85
        .Left = 20
        .Width = 350
        .Height = 320
        .Caption = "  ➕ Fiche Client"
        .Font.Bold = True
    End With
    Call StylerPanneau(panneauForm)
    
    ' Champ ID (lecture seule)
    Call CreerChampClient(panneauForm, "ID Client :", "txtIDClient", 20, 40, 150, True)
    
    ' Informations personnelles
    Call CreerChampClient(panneauForm, "Nom :", "txtNom", 20, 80, 150)
    Call CreerChampClient(panneauForm, "Prénom :", "txtPrenom", 180, 80, 150)
    
    ' Coordonnées
    Call CreerChampClient(panneauForm, "Téléphone :", "txtTelephone", 20, 120, 150)
    Call CreerChampClient(panneauForm, "Email :", "txtEmail", 180, 120, 150)
    
    ' Adresse complète
    Call CreerChampTexteMultiligneClient(panneauForm, "Adresse :", "txtAdresse", 20, 160, 310, 50)
    
    ' Informations système
    Dim lblDateCreation As MSForms.Label
    Set lblDateCreation = Me.Controls.Add("Forms.Label.1", "lblDateCreation")
    
    With lblDateCreation
        .Top = 225
        .Left = 40
        .Width = 290
        .Height = 15
        .Caption = "Date de création : " & Format(Date, "dd/mm/yyyy")
    End With
    Call StylerLabelSecondaire(lblDateCreation)
    
    ' Boutons d'action
    Call CreerBoutonClient(panneauForm, "💾 Enregistrer", "btnEnregistrerClient", 20, 260, 90, 35)
    Call CreerBoutonClient(panneauForm, "✏️ Modifier", "btnModifierClient", 120, 260, 90, 35)
    Call CreerBoutonClient(panneauForm, "🗑️ Supprimer", "btnSupprimerClient", 220, 260, 90, 35)
End Sub

Private Sub CreerChampClient(parent As MSForms.Frame, label As String, nomControle As String, _
                            posX As Integer, posY As Integer, largeur As Integer, Optional lectureSeule As Boolean = False)
    
    ' Label
    Dim lbl As MSForms.Label
    Set lbl = parent.Controls.Add("Forms.Label.1", "lbl" & nomControle)
    
    With lbl
        .Top = posY - 18
        .Left = posX
        .Width = largeur
        .Height = 15
        .Caption = label
    End With
    Call StylerLabelNormal(lbl)
    
    ' Champ de saisie
    Dim txt As MSForms.TextBox
    Set txt = parent.Controls.Add("Forms.TextBox.1", nomControle)
    
    With txt
        .Top = posY
        .Left = posX
        .Width = largeur
        .Height = 25
        .Locked = lectureSeule
    End With
    Call StylerChampTexte(txt)
    
    If lectureSeule Then
        txt.BackColor = COLOR_LIGHT_GRAY
    End If
End Sub

Private Sub CreerChampTexteMultiligneClient(parent As MSForms.Frame, label As String, nomControle As String, _
                                           posX As Integer, posY As Integer, largeur As Integer, hauteur As Integer)
    
    ' Label
    Dim lbl As MSForms.Label
    Set lbl = parent.Controls.Add("Forms.Label.1", "lbl" & nomControle)
    
    With lbl
        .Top = posY - 18
        .Left = posX
        .Width = largeur
        .Height = 15
        .Caption = label
    End With
    Call StylerLabelNormal(lbl)
    
    ' Zone de texte
    Dim txt As MSForms.TextBox
    Set txt = parent.Controls.Add("Forms.TextBox.1", nomControle)
    
    With txt
        .Top = posY
        .Left = posX
        .Width = largeur
        .Height = hauteur
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
    End With
    Call StylerChampTexte(txt)
End Sub

Private Sub CreerBoutonClient(parent As MSForms.Frame, texte As String, nomBouton As String, _
                             posX As Integer, posY As Integer, largeur As Integer, hauteur As Integer)
    
    Dim btn As MSForms.CommandButton
    Set btn = parent.Controls.Add("Forms.CommandButton.1", nomBouton)
    
    With btn
        .Top = posY
        .Left = posX
        .Width = largeur
        .Height = hauteur
        .Caption = texte
    End With
    
    ' Style selon le type
    If InStr(nomBouton, "Enregistrer") > 0 Or InStr(nomBouton, "Modifier") > 0 Then
        Call StylerBoutonSucces(btn)
    ElseIf InStr(nomBouton, "Supprimer") > 0 Then
        Call StylerBoutonDanger(btn)
    Else
        Call StylerBoutonPrimaire(btn)
    End If
End Sub

Private Sub CreerRechercheClient()
    ' Panneau de recherche
    Dim panneauRecherche As MSForms.Frame
    Set panneauRecherche = Me.Controls.Add("Forms.Frame.1", "panneauRechercheClient")
    
    With panneauRecherche
        .Top = 85
        .Left = 390
        .Width = 340
        .Height = 150
        .Caption = "  🔍 Recherche Client"
        .Font.Bold = True
    End With
    Call StylerPanneau(panneauRecherche)
    
    ' Champ de recherche
    Dim txtRecherche As MSForms.TextBox
    Set txtRecherche = panneauRecherche.Controls.Add("Forms.TextBox.1", "txtRechercheClient")
    
    With txtRecherche
        .Top = 30
        .Left = 15
        .Width = 200
        .Height = 25
    End With
    Call StylerChampTexte(txtRecherche)
    
    ' Bouton rechercher
    Dim btnRechercher As MSForms.CommandButton
    Set btnRechercher = panneauRecherche.Controls.Add("Forms.CommandButton.1", "btnRechercherClient")
    
    With btnRechercher
        .Top = 30
        .Left = 225
        .Width = 100
        .Height = 25
        .Caption = "🔍 Rechercher"
    End With
    Call StylerBoutonPrimaire(btnRechercher)
    
    ' Liste des résultats
    Dim lstResultats As MSForms.ListBox
    Set lstResultats = panneauRecherche.Controls.Add("Forms.ListBox.1", "lstResultatsRecherche")
    
    With lstResultats
        .Top = 70
        .Left = 15
        .Width = 310
        .Height = 60
        .Font.Name = FONT_MONOSPACE
        .Font.Size = 8
    End With
    Call StylerChampTexte(lstResultats)
End Sub

Private Sub CreerHistoriqueClient()
    ' Panneau historique
    Dim panneauHistorique As MSForms.Frame
    Set panneauHistorique = Me.Controls.Add("Forms.Frame.1", "panneauHistoriqueClient")
    
    With panneauHistorique
        .Top = 255
        .Left = 390
        .Width = 340
        .Height = 150
        .Caption = "  📋 Historique des Séjours"
        .Font.Bold = True
    End With
    Call StylerPanneauAccent(panneauHistorique)
    
    ' Liste historique
    Dim lstHistorique As MSForms.ListBox
    Set lstHistorique = panneauHistorique.Controls.Add("Forms.ListBox.1", "lstHistoriqueClient")
    
    With lstHistorique
        .Top = 30
        .Left = 15
        .Width = 310
        .Height = 80
        .Font.Name = FONT_MONOSPACE
        .Font.Size = 8
    End With
    Call StylerChampTexte(lstHistorique)
    
    ' Statistiques client
    Dim lblStatsClient As MSForms.Label
    Set lblStatsClient = Me.Controls.Add("Forms.Label.1", "lblStatsClient")
    
    With lblStatsClient
        .Top = 120
        .Left = 405
        .Width = 310
        .Height = 20
        .Caption = "💰 Total dépensé : 0,00 € | 🛏️ Séjours : 0 | ⭐ Client depuis : --"
        .TextAlign = fmTextAlignCenter
    End With
    Call StylerLabelSecondaire(lblStatsClient)
End Sub

Private Sub CreerBarreActionsClients()
    ' Barre d'actions
    Dim panneauActions As MSForms.Frame
    Set panneauActions = Me.Controls.Add("Forms.Frame.1", "panneauActionsClients")
    
    With panneauActions
        .Top = 425
        .Left = 20
        .Width = 710
        .Height = 60
        .Caption = ""
    End With
    Call StylerPanneauAccent(panneauActions)
    
    ' Boutons d'actions
    Call CreerBoutonClient(panneauActions, "🏠 Menu Principal", "btnRetourMenuClients", 20, 15, 120, 30)
    Call CreerBoutonClient(panneauActions, "📊 Rapport Clients", "btnRapportClients", 160, 15, 120, 30)
    Call CreerBoutonClient(panneauActions, "📧 Envoyer Email", "btnEnvoyerEmailClient", 300, 15, 120, 30)
    Call CreerBoutonClient(panneauActions, "💾 Exporter", "btnExporterClients", 440, 15, 100, 30)
    Call CreerBoutonClient(panneauActions, "🔄 Actualiser", "btnActualiserClients", 560, 15, 80, 30)
    Call CreerBoutonClient(panneauActions, "❓ Aide", "btnAideClients", 660, 15, 50, 30)
End Sub

' ========================================
' ÉVÉNEMENTS ET INTERACTIONS
' ========================================

Private Sub btnEnregistrerClient_Click()
    If ValiderFormulaireClient() Then
        Call EnregistrerNouveauClient
        Call AfficherMessageSucces(Me, "Client enregistré avec succès !")
        Call ViderFormulaireClient
        Call ActualiserStatsClients
    End If
End Sub

Private Sub btnModifierClient_Click()
    If Me.Controls("txtIDClient").Value <> "" Then
        If ValiderFormulaireClient() Then
            Call ModifierClientExistant
            Call AfficherMessageSucces(Me, "Client modifié avec succès !")
            Call ActualiserHistoriqueClient
        End If
    Else
        Call AfficherMessageErreur(Me, "Veuillez d'abord sélectionner un client")
    End If
End Sub

Private Sub btnSupprimerClient_Click()
    If Me.Controls("txtIDClient").Value <> "" Then
        Dim reponse As VbMsgBoxResult
        reponse = MsgBox("Êtes-vous sûr de vouloir supprimer ce client ?", vbYesNo + vbQuestion, "Confirmation")
        
        If reponse = vbYes Then
            Call SupprimerClientSelectionne
            Call AfficherMessageSucces(Me, "Client supprimé avec succès !")
            Call ViderFormulaireClient
            Call ActualiserStatsClients
        End If
    Else
        Call AfficherMessageErreur(Me, "Veuillez d'abord sélectionner un client")
    End If
End Sub

Private Sub btnRechercherClient_Click()
    Dim critereRecherche As String
    critereRecherche = Me.Controls("txtRechercheClient").Value
    
    If critereRecherche <> "" Then
        Call EffectuerRechercheClient(critereRecherche)
    Else
        Call AfficherMessageErreur(Me, "Veuillez saisir un critère de recherche")
    End If
End Sub

Private Sub lstResultatsRecherche_Click()
    ' Charger le client sélectionné dans le formulaire
    Dim selectionIndex As Integer
    selectionIndex = Me.Controls("lstResultatsRecherche").ListIndex
    
    If selectionIndex >= 0 Then
        Call ChargerClientDansFormulaire(selectionIndex)
    End If
End Sub

Private Sub txtEmail_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Validation de l'email en temps réel
    Dim email As String
    email = Me.Controls("txtEmail").Value
    
    If email <> "" And Not ValiderEmail(email) Then
        Call AfficherMessageErreur(Me, "Format d'email invalide")
        Cancel = True
    End If
End Sub

Private Sub txtTelephone_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Validation du téléphone en temps réel
    Dim telephone As String
    telephone = Me.Controls("txtTelephone").Value
    
    If telephone <> "" And Not ValiderTelephone(telephone) Then
        Call AfficherMessageErreur(Me, "Format de téléphone invalide")
        Cancel = True
    End If
End Sub

Private Sub btnRetourMenuClients_Click()
    Me.Hide
    UserForm_MenuPrincipal.Show
End Sub

' ========================================
' FONCTIONS UTILITAIRES
' ========================================

Private Function ValiderFormulaireClient() As Boolean
    ' Validation des champs obligatoires
    If Trim(Me.Controls("txtNom").Value) = "" Then
        Call AfficherMessageErreur(Me, "Le nom est obligatoire")
        ValiderFormulaireClient = False
        Exit Function
    End If
    
    If Trim(Me.Controls("txtPrenom").Value) = "" Then
        Call AfficherMessageErreur(Me, "Le prénom est obligatoire")
        ValiderFormulaireClient = False
        Exit Function
    End If
    
    ' Validation de l'email si renseigné
    Dim email As String
    email = Me.Controls("txtEmail").Value
    If email <> "" And Not ValiderEmail(email) Then
        Call AfficherMessageErreur(Me, "Format d'email invalide")
        ValiderFormulaireClient = False
        Exit Function
    End If
    
    ' Validation du téléphone si renseigné
    Dim telephone As String
    telephone = Me.Controls("txtTelephone").Value
    If telephone <> "" And Not ValiderTelephone(telephone) Then
        Call AfficherMessageErreur(Me, "Format de téléphone invalide")
        ValiderFormulaireClient = False
        Exit Function
    End If
    
    ValiderFormulaireClient = True
End Function

Private Sub EnregistrerNouveauClient()
    Dim idClient As Long
    
    idClient = AjouterClient( _
        Me.Controls("txtNom").Value, _
        Me.Controls("txtPrenom").Value, _
        Me.Controls("txtTelephone").Value, _
        Me.Controls("txtEmail").Value, _
        Me.Controls("txtAdresse").Value _
    )
    
    If idClient > 0 Then
        Me.Controls("txtIDClient").Value = idClient
        Me.Controls("lblDateCreation").Caption = "Date de création : " & Format(Date, "dd/mm/yyyy")
    End If
End Sub

Private Sub ModifierClientExistant()
    Dim idClient As Long
    idClient = CLng(Me.Controls("txtIDClient").Value)
    
    Call ModifierClient( _
        idClient, _
        Me.Controls("txtNom").Value, _
        Me.Controls("txtPrenom").Value, _
        Me.Controls("txtTelephone").Value, _
        Me.Controls("txtEmail").Value, _
        Me.Controls("txtAdresse").Value _
    )
End Sub

Private Sub SupprimerClientSelectionne()
    Dim idClient As Long
    idClient = CLng(Me.Controls("txtIDClient").Value)
    
    Call SupprimerClient(idClient)
End Sub

Private Sub ViderFormulaireClient()
    Me.Controls("txtIDClient").Value = ""
    Me.Controls("txtNom").Value = ""
    Me.Controls("txtPrenom").Value = ""
    Me.Controls("txtTelephone").Value = ""
    Me.Controls("txtEmail").Value = ""
    Me.Controls("txtAdresse").Value = ""
    Me.Controls("lblDateCreation").Caption = "Date de création : " & Format(Date, "dd/mm/yyyy")
    
    ' Vider l'historique
    Me.Controls("lstHistoriqueClient").Clear
    Me.Controls("lblStatsClient").Caption = "💰 Total dépensé : 0,00 € | 🛏️ Séjours : 0 | ⭐ Client depuis : --"
End Sub

Private Sub EffectuerRechercheClient(critere As String)
    Dim resultats As Variant
    Dim lst As MSForms.ListBox
    Dim i As Integer
    
    Set lst = Me.Controls("lstResultatsRecherche")
    lst.Clear
    
    resultats = RechercherClientsParNom(critere)
    
    For i = 0 To UBound(resultats)
        lst.AddItem resultats(i)
    Next i
End Sub

Private Sub ChargerClientDansFormulaire(index As Integer)
    ' Extraire l'ID du client depuis la sélection
    Dim selection As String
    Dim idClient As Long
    
    selection = Me.Controls("lstResultatsRecherche").List(index)
    idClient = CLng(Left(selection, InStr(selection, " ") - 1))
    
    ' Charger les données du client
    Dim clientInfo As Variant
    clientInfo = RechercherClientParID(idClient)
    
    If Not IsEmpty(clientInfo) Then
        Me.Controls("txtIDClient").Value = clientInfo(0)
        Me.Controls("txtNom").Value = clientInfo(1)
        Me.Controls("txtPrenom").Value = clientInfo(2)
        Me.Controls("txtTelephone").Value = clientInfo(3)
        Me.Controls("txtEmail").Value = clientInfo(4)
        Me.Controls("txtAdresse").Value = clientInfo(5)
        Me.Controls("lblDateCreation").Caption = "Date de création : " & Format(clientInfo(6), "dd/mm/yyyy")
        
        Call ActualiserHistoriqueClient
    End If
End Sub

Private Sub ActualiserHistoriqueClient()
    If Me.Controls("txtIDClient").Value <> "" Then
        Dim idClient As Long
        Dim historique As Variant
        Dim lst As MSForms.ListBox
        Dim i As Integer
        
        idClient = CLng(Me.Controls("txtIDClient").Value)
        Set lst = Me.Controls("lstHistoriqueClient")
        lst.Clear
        
        historique = ObtenirHistoriqueClient(idClient)
        
        For i = 0 To UBound(historique)
            lst.AddItem historique(i)
        Next i
        
        ' Mettre à jour les statistiques
        Call ActualiserStatsClientSelectionne(idClient)
    End If
End Sub

Private Sub ActualiserStatsClientSelectionne(idClient As Long)
    ' Calculer les statistiques du client
    Dim nbSejours As Integer
    Dim totalDepense As Double
    Dim clientDepuis As String
    
    ' Simuler le calcul (à implémenter avec les vraies données)
    nbSejours = 3
    totalDepense = 1250.5
    clientDepuis = "Jan 2024"
    
    Me.Controls("lblStatsClient").Caption = "💰 Total dépensé : " & Format(totalDepense, "0.00") & " € | 🛏️ Séjours : " & nbSejours & " | ⭐ Client depuis : " & clientDepuis
End Sub

Private Sub ActualiserStatsClients()
    Me.Controls("lblStatsClients").Caption = "📊 Total clients : " & CompterTotalClients() & " | Nouveaux ce mois : " & CompterNouveauxClients()
End Sub

Private Function CompterTotalClients() As Integer
    ' Compter le nombre total de clients
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CLIENTS)
    CompterTotalClients = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1
End Function

Private Function CompterNouveauxClients() As Integer
    ' Compter les nouveaux clients du mois
    CompterNouveauxClients = 5 ' Simulé
End Function
