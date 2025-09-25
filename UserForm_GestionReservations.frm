VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_GestionReservations 
   Caption         =   "📅 Gestion des Réservations"
   ClientHeight    =   600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   800
   OleObjectBlob   =   "UserForm_GestionReservations.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_GestionReservations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ========================================
' USERFORM GESTION RESERVATIONS - UX MODERNE
' ========================================

Private Sub UserForm_Initialize()
    Call InitialiserInterface
End Sub

Private Sub InitialiserInterface()
    Call AppliquerStyleModerne(Me)
    
    Me.Width = 800
    Me.Height = 600
    Me.Caption = "📅 Gestion des Réservations"
    
    Call CreerEnTeteReservations
    Call CreerFormulaireReservation
    Call CreerListeReservations
    Call CreerBarreActions
    
    ' Initialiser l'accessibilité
    Call InitialiserAccessibilite
    
    ' Animation d'entrée
    Call AnimerApparitionFormulaire(Me)
End Sub

Private Sub InitialiserAccessibilite()
    ' Définir l'ordre de tabulation
    Call DefinirOrdreTabulation(Me)
    
    ' Créer les tooltips
    Call CreerTooltip(Me.btnEnregistrer, "Enregistrer la réservation")
    Call CreerTooltip(Me.btnNouveau, "Créer une nouvelle réservation")
    Call CreerTooltip(Me.btnAnnuler, "Annuler les modifications")
    Call CreerTooltip(Me.btnRechercher, "Rechercher une réservation")
    Call CreerTooltip(Me.btnModifier, "Modifier la réservation sélectionnée")
    Call CreerTooltip(Me.btnSupprimer, "Supprimer la réservation")
    Call CreerTooltip(Me.btnActualiser, "Actualiser la liste")
End Sub

Private Sub CreerEnTeteReservations()
    ' Titre avec icône
    Dim lblTitre As MSForms.Label
    Set lblTitre = Me.Controls.Add("Forms.Label.1", "lblTitreReservations")
    
    With lblTitre
        .Top = 15
        .Left = 20
        .Width = 760
        .Height = 35
        .Caption = "📅 GESTION DES RÉSERVATIONS"
        .TextAlign = fmTextAlignCenter
    End With
    Call StylerTitre(lblTitre)
    
    ' Navigation breadcrumb
    Dim lblBreadcrumb As MSForms.Label
    Set lblBreadcrumb = Me.Controls.Add("Forms.Label.1", "lblBreadcrumb")
    
    With lblBreadcrumb
        .Top = 55
        .Left = 20
        .Width = 760
        .Height = 20
        .Caption = "🏠 Accueil > 📅 Réservations"
        .TextAlign = fmTextAlignLeft
    End With
    Call StylerLabelSecondaire(lblBreadcrumb)
End Sub

Private Sub CreerFormulaireReservation()
    ' Panneau du formulaire
    Dim panneauForm As MSForms.Frame
    Set panneauForm = Me.Controls.Add("Forms.Frame.1", "panneauFormulaire")
    
    With panneauForm
        .Top = 85
        .Left = 20
        .Width = 380
        .Height = 400
        .Caption = "  ➕ Nouvelle Réservation"
        .Font.Bold = True
    End With
    Call StylerPanneau(panneauForm)
    
    ' Sélection du client
    Call CreerChampFormulaire(panneauForm, "Client :", "cmbClient", 20, 40, 340, True)
    
    ' Sélection de la chambre
    Call CreerChampFormulaire(panneauForm, "Chambre :", "cmbChambre", 20, 90, 340, True)
    
    ' Dates avec calendrier visuel
    Call CreerChampDate(panneauForm, "Date d'arrivée :", "txtDateArrivee", 20, 140, 160)
    Call CreerChampDate(panneauForm, "Date de départ :", "txtDateDepart", 200, 140, 160)
    
    ' Calcul automatique des nuits
    Dim lblNuits As MSForms.Label
    Set lblNuits = Me.Controls.Add("Forms.Label.1", "lblNombreNuits")
    
    With lblNuits
        .Top = 190
        .Left = 40
        .Width = 320
        .Height = 25
        .Caption = "Nombre de nuits : 0 | Montant total : 0,00 €"
        .TextAlign = fmTextAlignCenter
        .Font.Bold = True
    End With
    Call StylerLabelNormal(lblNuits)
    lblNuits.ForeColor = COLOR_ACCENT
    
    ' Commentaires
' ========================================
' VALIDATION ET CALCULS AMÉLIORÉS
' ========================================

Private Sub txtDateArrivee_Change()
    Call CalculerEtValiderDates
End Sub

Private Sub txtDateDepart_Change()
    Call CalculerEtValiderDates
End Sub

Private Sub CalculerEtValiderDates()
    On Error Resume Next
    
    Dim dateArrivee As Date
    Dim dateDepart As Date
    Dim nombreNuits As Integer
    Dim montantTotal As Double
    
    ' Valider les dates
    If IsDate(Me.txtDateArrivee.Value) And IsDate(Me.txtDateDepart.Value) Then
        dateArrivee = CDate(Me.txtDateArrivee.Value)
        dateDepart = CDate(Me.txtDateDepart.Value)
        
        If dateDepart > dateArrivee Then
            nombreNuits = DateDiff("d", dateArrivee, dateDepart)
            
            ' Calculer le montant (prix moyen par nuit)
            montantTotal = nombreNuits * 85 ' Prix moyen d'une chambre
            
            ' Mettre à jour l'affichage
            Me.lblNombreNuits.Caption = "Nombre de nuits : " & nombreNuits & " | Montant total : " & Format(montantTotal, "0.00") & " €"
            
            ' Vérifier la disponibilité
            If Not VerifierDisponibiliteChambre(dateArrivee, dateDepart) Then
                Call AfficherNotificationAvancee(Me, "Attention : chambre peut-être indisponible", "warning", 3)
            End If
        Else
            Me.lblNombreNuits.Caption = "Date de départ doit être après la date d'arrivée"
            Call AfficherMessageErreur(Me, "Date de départ invalide")
        End If
    Else
        Me.lblNombreNuits.Caption = "Format de date invalide"
    End If
End Sub

Private Function VerifierDisponibiliteChambre(dateArrivee As Date, dateDepart As Date) As Boolean
    ' Simulation de vérification de disponibilité
    ' En production, ceci devrait interroger la base de données
    VerifierDisponibiliteChambre = True ' Simulation
End Function

' ========================================
' ÉVÉNEMENTS DES BOUTONS AVEC VALIDATION
' ========================================

Private Sub btnEnregistrer_Click()
    ' Créer les règles de validation
    Dim regles As New Collection
    
    regles.Add Array("obligatoire", "Client"), "cmbClient"
    regles.Add Array("obligatoire", "Chambre"), "cmbChambre"
    regles.Add Array("date", "Date d'arrivée"), "txtDateArrivee"
    regles.Add Array("date", "Date de départ"), "txtDateDepart"
    
    If ValiderFormulaireComplet(Me, regles) Then
        Call AfficherNotificationAvancee(Me, "Réservation enregistrée avec succès !", "success", 3)
        Call ReinitialiserFormulaire
    End If
End Sub

Private Sub btnNouveau_Click()
    Call ReinitialiserFormulaire
    Call AfficherNotificationAvancee(Me, "Nouveau formulaire de réservation", "info", 2)
End Sub

Private Sub btnAnnuler_Click()
    If MsgBox("Voulez-vous vraiment annuler les modifications ?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
        Call ReinitialiserFormulaire
        Call AfficherNotificationAvancee(Me, "Modifications annulées", "info", 2)
    End If
End Sub

Private Sub ReinitialiserFormulaire()
    Me.cmbClient.Value = ""
    Me.cmbChambre.Value = ""
    Me.txtDateArrivee.Value = Format(Date, "dd/mm/yyyy")
    Me.txtDateDepart.Value = Format(Date + 1, "dd/mm/yyyy")
    Me.txtCommentaires.Value = ""
    Call CalculerEtValiderDates
End Sub

' ========================================
' NAVIGATION CLAVIER
' ========================================

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub cmbClient_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub cmbChambre_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub txtDateArrivee_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub txtDateDepart_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub txtCommentaires_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub btnEnregistrer_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub btnNouveau_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub btnAnnuler_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

' ========================================
' EFFETS DE SURVOL AVEC TOOLTIPS
' ========================================

Private Sub btnEnregistrer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnEnregistrer, True)
    Call AfficherTooltip(Me.btnEnregistrer, True)
End Sub

Private Sub btnNouveau_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnNouveau, True)
    Call AfficherTooltip(Me.btnNouveau, True)
End Sub

Private Sub btnAnnuler_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnAnnuler, True)
    Call AfficherTooltip(Me.btnAnnuler, True)
End Sub

Private Sub btnRechercher_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnRechercher, True)
    Call AfficherTooltip(Me.btnRechercher, True)
End Sub

Private Sub btnModifier_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnModifier, True)
    Call AfficherTooltip(Me.btnModifier, True)
End Sub

Private Sub btnSupprimer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnSupprimer, True)
    Call AfficherTooltip(Me.btnSupprimer, True)
End Sub

Private Sub btnActualiser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnActualiser, True)
    Call AfficherTooltip(Me.btnActualiser, True)
End Sub

' Masquer les tooltips
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call AfficherTooltip(Me.btnEnregistrer, False)
    Call AfficherTooltip(Me.btnNouveau, False)
    Call AfficherTooltip(Me.btnAnnuler, False)
    Call AfficherTooltip(Me.btnRechercher, False)
    Call AfficherTooltip(Me.btnModifier, False)
    Call AfficherTooltip(Me.btnSupprimer, False)
    Call AfficherTooltip(Me.btnActualiser, False)
End Sub

Private Sub CreerChampFormulaire(parent As MSForms.Frame, label As String, nomControle As String, _
                                posX As Integer, posY As Integer, largeur As Integer, Optional estCombo As Boolean = False)
    
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
    
    ' Contrôle de saisie
    If estCombo Then
        Dim cmb As MSForms.ComboBox
        Set cmb = parent.Controls.Add("Forms.ComboBox.1", nomControle)
        
        With cmb
            .Top = posY
            .Left = posX
            .Width = largeur
            .Height = 25
        End With
        Call StylerComboBox(cmb)
        
        ' Remplir les données selon le type
        If nomControle = "cmbClient" Then
            Call RemplirComboClients(cmb)
        ElseIf nomControle = "cmbChambre" Then
            Call RemplirComboChambres(cmb)
        End If
    Else
        Dim txt As MSForms.TextBox
        Set txt = parent.Controls.Add("Forms.TextBox.1", nomControle)
        
        With txt
            .Top = posY
            .Left = posX
            .Width = largeur
            .Height = 25
        End With
        Call StylerChampTexte(txt)
    End If
End Sub

Private Sub CreerChampDate(parent As MSForms.Frame, label As String, nomControle As String, _
                          posX As Integer, posY As Integer, largeur As Integer)
    
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
    
    ' Champ de date
    Dim txt As MSForms.TextBox
    Set txt = parent.Controls.Add("Forms.TextBox.1", nomControle)
    
    With txt
        .Top = posY
        .Left = posX
        .Width = largeur - 30
        .Height = 25
        .Value = Format(Date, "dd/mm/yyyy")
    End With
    Call StylerChampTexte(txt)
    
    ' Bouton calendrier
    Dim btnCal As MSForms.CommandButton
    Set btnCal = parent.Controls.Add("Forms.CommandButton.1", "btn" & nomControle)
    
    With btnCal
        .Top = posY
        .Left = posX + largeur - 25
        .Width = 25
        .Height = 25
        .Caption = "📅"
        .Font.Size = 8
    End With
    Call StylerBoutonSecondaire(btnCal)
End Sub

Private Sub CreerChampTexteMultiligne(parent As MSForms.Frame, label As String, nomControle As String, _
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
    
    ' Zone de texte multiligne
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

Private Sub CreerBoutonAction(parent As MSForms.Frame, texte As String, nomBouton As String, _
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
    
    ' Style selon le type de bouton
    If InStr(nomBouton, "Enregistrer") > 0 Then
        Call StylerBoutonSucces(btn)
    ElseIf InStr(nomBouton, "Annuler") > 0 Then
        Call StylerBoutonDanger(btn)
    Else
        Call StylerBoutonSecondaire(btn)
    End If
End Sub

Private Sub CreerListeReservations()
    ' Panneau de la liste
    Dim panneauListe As MSForms.Frame
    Set panneauListe = Me.Controls.Add("Forms.Frame.1", "panneauListe")
    
    With panneauListe
        .Top = 85
        .Left = 420
        .Width = 360
        .Height = 400
        .Caption = "  📋 Réservations Actuelles"
        .Font.Bold = True
    End With
    Call StylerPanneau(panneauListe)
    
    ' Filtres rapides
    Call CreerFiltresRapides(panneauListe)
    
    ' Liste des réservations
    Dim lstReservations As MSForms.ListBox
    Set lstReservations = panneauListe.Controls.Add("Forms.ListBox.1", "lstReservations")
    
    With lstReservations
        .Top = 80
        .Left = 15
        .Width = 330
        .Height = 250
        .Font.Name = FONT_MONOSPACE
        .Font.Size = 8
    End With
    Call StylerChampTexte(lstReservations)
    
    ' Boutons de gestion
    Call CreerBoutonAction(panneauListe, "✏️ Modifier", "btnModifier", 15, 345, 80, 30)
    Call CreerBoutonAction(panneauListe, "✅ Confirmer", "btnConfirmer", 105, 345, 80, 30)
    Call CreerBoutonAction(panneauListe, "❌ Annuler", "btnAnnulerRes", 195, 345, 80, 30)
    Call CreerBoutonAction(panneauListe, "🔄 Actualiser", "btnActualiserListe", 285, 345, 60, 30)
    
    ' Charger les données
    Call ChargerListeReservations
End Sub

Private Sub CreerFiltresRapides(parent As MSForms.Frame)
    ' Filtre par statut
    Dim cmbFiltre As MSForms.ComboBox
    Set cmbFiltre = parent.Controls.Add("Forms.ComboBox.1", "cmbFiltreStatut")
    
    With cmbFiltre
        .Top = 30
        .Left = 15
        .Width = 120
        .Height = 25
        .AddItem "Toutes"
        .AddItem "Confirmées"
        .AddItem "En attente"
        .AddItem "Annulées"
        .Value = "Toutes"
    End With
    Call StylerComboBox(cmbFiltre)
    
    ' Filtre par date
    Dim cmbFiltreDate As MSForms.ComboBox
    Set cmbFiltreDate = parent.Controls.Add("Forms.ComboBox.1", "cmbFiltreDate")
    
    With cmbFiltreDate
        .Top = 30
        .Left = 145
        .Width = 120
        .Height = 25
        .AddItem "Toutes les dates"
        .AddItem "Aujourd'hui"
        .AddItem "Cette semaine"
        .AddItem "Ce mois"
        .Value = "Toutes les dates"
    End With
    Call StylerComboBox(cmbFiltreDate)
    
    ' Bouton recherche
    Dim btnRecherche As MSForms.CommandButton
    Set btnRecherche = parent.Controls.Add("Forms.CommandButton.1", "btnRecherche")
    
    With btnRecherche
        .Top = 30
        .Left = 275
        .Width = 70
        .Height = 25
        .Caption = "🔍 Filtrer"
    End With
    Call StylerBoutonPrimaire(btnRecherche)
End Sub

Private Sub CreerBarreActions()
    ' Barre d'actions en bas
    Dim panneauActions As MSForms.Frame
    Set panneauActions = Me.Controls.Add("Forms.Frame.1", "panneauActions")
    
    With panneauActions
        .Top = 505
        .Left = 20
        .Width = 760
        .Height = 60
        .Caption = ""
    End With
    Call StylerPanneauAccent(panneauActions)
    
    ' Boutons d'actions globales
    Call CreerBoutonAction(panneauActions, "🏠 Menu Principal", "btnRetourMenu", 20, 15, 120, 30)
    Call CreerBoutonAction(panneauActions, "📊 Rapport Réservations", "btnRapportRes", 160, 15, 140, 30)
    Call CreerBoutonAction(panneauActions, "📧 Envoyer Confirmation", "btnEnvoyerEmail", 320, 15, 140, 30)
    Call CreerBoutonAction(panneauActions, "💾 Exporter Liste", "btnExporter", 480, 15, 100, 30)
    Call CreerBoutonAction(panneauActions, "❓ Aide", "btnAideRes", 600, 15, 60, 30)
End Sub

' ========================================
' FONCTIONS UTILITAIRES
' ========================================

Private Sub RemplirComboClients(cmb As MSForms.ComboBox)
    Dim clients As Variant
    Dim i As Integer
    
    clients = ObtenirListeClients()
    
    cmb.Clear
    For i = 0 To UBound(clients)
        cmb.AddItem clients(i)
    Next i
End Sub

Private Sub RemplirComboChambres(cmb As MSForms.ComboBox)
    Dim chambres As Variant
    Dim i As Integer
    
    chambres = ObtenirChambresLibres()
    
    cmb.Clear
    For i = 0 To UBound(chambres)
        cmb.AddItem chambres(i)
    Next i
End Sub

Private Sub ChargerListeReservations()
    Dim lst As MSForms.ListBox
    Set lst = Me.Controls("lstReservations")
    
    ' Simuler le chargement des réservations
    lst.Clear
    lst.AddItem "Rés. 001 - Dupont Jean - Ch.101 - 25/12/2024 [Confirmée]"
    lst.AddItem "Rés. 002 - Martin Marie - Ch.201 - 26/12/2024 [En attente]"
    lst.AddItem "Rés. 003 - Bernard Pierre - Ch.301 - 27/12/2024 [Confirmée]"
    ' ... autres réservations
End Sub

' ========================================
' ÉVÉNEMENTS DES CONTRÔLES
' ========================================

Private Sub txtDateArrivee_Change()
    Call CalculerMontantTotal
End Sub

Private Sub txtDateDepart_Change()
    Call CalculerMontantTotal
End Sub

Private Sub cmbChambre_Change()
    Call CalculerMontantTotal
End Sub

Private Sub CalculerMontantTotal()
    On Error Resume Next
    
    Dim dateArr As Date, dateDep As Date
    Dim nbNuits As Integer
    Dim tarifNuit As Double
    Dim montantTotal As Double
    
    ' Récupérer les dates
    dateArr = CDate(Me.Controls("txtDateArrivee").Value)
    dateDep = CDate(Me.Controls("txtDateDepart").Value)
    
    If dateDep > dateArr Then
        nbNuits = dateDep - dateArr
        
        ' Récupérer le tarif de la chambre sélectionnée
        Dim chambreSelectionnee As String
        chambreSelectionnee = Me.Controls("cmbChambre").Value
        
        If chambreSelectionnee <> "" Then
            ' Extraire le numéro de chambre
            Dim numChambre As String
            numChambre = Left(chambreSelectionnee, 3)
            tarifNuit = ObtenirTarifChambre(numChambre)
            montantTotal = nbNuits * tarifNuit
            
            Me.Controls("lblNombreNuits").Caption = "Nombre de nuits : " & nbNuits & " | Montant total : " & Format(montantTotal, "0.00") & " €"
        End If
    End If
End Sub

Private Sub btnEnregistrer_Click()
    ' Validation et enregistrement
    If ValiderFormulaire() Then
        Call EnregistrerReservation
        Call AfficherMessageSucces(Me, "Réservation enregistrée avec succès !")
        Call ViderFormulaire
        Call ChargerListeReservations
    End If
End Sub

Private Sub btnNouveau_Click()
    Call ViderFormulaire
End Sub

Private Sub btnRetourMenu_Click()
    Me.Hide
    UserForm_MenuPrincipal.Show
End Sub

Private Function ValiderFormulaire() As Boolean
    ' Validation des champs obligatoires
    If Me.Controls("cmbClient").Value = "" Then
        Call AfficherMessageErreur(Me, "Veuillez sélectionner un client")
        ValiderFormulaire = False
        Exit Function
    End If
    
    If Me.Controls("cmbChambre").Value = "" Then
        Call AfficherMessageErreur(Me, "Veuillez sélectionner une chambre")
        ValiderFormulaire = False
        Exit Function
    End If
    
    ' Validation des dates
    On Error GoTo ErreurDate
    Dim dateArr As Date, dateDep As Date
    dateArr = CDate(Me.Controls("txtDateArrivee").Value)
    dateDep = CDate(Me.Controls("txtDateDepart").Value)
    
    If dateDep <= dateArr Then
        Call AfficherMessageErreur(Me, "La date de départ doit être postérieure à l'arrivée")
        ValiderFormulaire = False
        Exit Function
    End If
    
    ValiderFormulaire = True
    Exit Function
    
ErreurDate:
    Call AfficherMessageErreur(Me, "Format de date invalide")
    ValiderFormulaire = False
End Function

Private Sub EnregistrerReservation()
    ' Logique d'enregistrement de la réservation
    ' Appeler les fonctions du module Reservations
End Sub

Private Sub ViderFormulaire()
    Me.Controls("cmbClient").Value = ""
    Me.Controls("cmbChambre").Value = ""
    Me.Controls("txtDateArrivee").Value = Format(Date, "dd/mm/yyyy")
    Me.Controls("txtDateDepart").Value = Format(Date + 1, "dd/mm/yyyy")
    Me.Controls("txtCommentaires").Value = ""
    Me.Controls("lblNombreNuits").Caption = "Nombre de nuits : 0 | Montant total : 0,00 €"
End Sub
