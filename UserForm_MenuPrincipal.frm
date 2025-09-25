VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_MenuPrincipal 
   Caption         =   "🏨 Gestion Auberge - Menu Principal"
   ClientHeight    =   480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   640
   OleObjectBlob   =   "UserForm_MenuPrincipal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_MenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ========================================
' USERFORM MENU PRINCIPAL - DESIGN MODERNE
' ========================================
' Interface principale avec navigation intuitive

Private Sub UserForm_Initialize()
    Call InitialiserInterface
End Sub

Private Sub InitialiserInterface()
    ' Appliquer le style moderne au formulaire
    Call AppliquerStyleModerne(Me)
    
    ' Configuration du formulaire
    Me.Width = 640
    Me.Height = 480
    Me.Caption = "🏨 Gestion Auberge - Menu Principal"
    
    ' Créer l'interface
    Call CreerEnTete
    Call CreerMenuNavigation
    Call CreerTableauBordRapide
    Call CreerPiedPage
    
    ' Initialiser l'accessibilité
    Call InitialiserAccessibilite
    
    ' Animation d'entrée
    Call AnimerApparitionFormulaire(Me)
End Sub

Private Sub InitialiserAccessibilite()
    ' Définir l'ordre de tabulation
    Call DefinirOrdreTabulation(Me)
    
    ' Créer les tooltips pour les boutons
    Call CreerTooltip(Me.btnChambres, "Gérer les chambres de l'auberge")
    Call CreerTooltip(Me.btnClients, "Gérer la base clients")
    Call CreerTooltip(Me.btnReservations, "Gérer les réservations")
    Call CreerTooltip(Me.btnPaiements, "Traiter les paiements")
    Call CreerTooltip(Me.btnRapports, "Consulter les rapports")
    Call CreerTooltip(Me.btnParametres, "Configurer l'application")
    Call CreerTooltip(Me.btnActualiser, "Actualiser les données")
    Call CreerTooltip(Me.btnAide, "Afficher l'aide")
End Sub

Private Sub CreerEnTete()
    ' Logo et titre principal
    Dim lblTitre As MSForms.Label
    Set lblTitre = Me.Controls.Add("Forms.Label.1", "lblTitrePrincipal")
    
    With lblTitre
        .Top = 20
        .Left = 20
        .Width = 600
        .Height = 40
        .Caption = "🏨 GESTION AUBERGE"
        .TextAlign = fmTextAlignCenter
    End With
    Call StylerTitre(lblTitre)
    
    ' Sous-titre avec informations
    Dim lblSousTitre As MSForms.Label
    Set lblSousTitre = Me.Controls.Add("Forms.Label.1", "lblSousTitre")
    
    With lblSousTitre
        .Top = 65
        .Left = 20
        .Width = 600
        .Height = 20
        .Caption = "Tableau de bord principal - " & Format(Date, "dddd dd mmmm yyyy")
        .TextAlign = fmTextAlignCenter
    End With
    Call StylerLabelSecondaire(lblSousTitre)
    
    ' Ligne de séparation
    Dim ligneSeparation As MSForms.Label
    Set ligneSeparation = Me.Controls.Add("Forms.Label.1", "ligneSeparation")
    
    With ligneSeparation
        .Top = 90
        .Left = 20
        .Width = 600
        .Height = 2
        .BackColor = COLOR_ACCENT
        .BorderStyle = fmBorderStyleNone
    End With
End Sub

Private Sub CreerMenuNavigation()
    ' Panneau de navigation principal
    Dim panneauNav As MSForms.Frame
    Set panneauNav = Me.Controls.Add("Forms.Frame.1", "panneauNavigation")
    
    With panneauNav
        .Top = 110
        .Left = 20
        .Width = 600
        .Height = 200
        .Caption = ""
    End With
    Call StylerPanneau(panneauNav)
    
    ' Boutons de navigation avec icônes
    Call CreerBoutonNavigation(panneauNav, "🛏️ Gestion Chambres", "btnChambres", 20, 20, 180, 50)
    Call CreerBoutonNavigation(panneauNav, "👤 Gestion Clients", "btnClients", 220, 20, 180, 50)
    Call CreerBoutonNavigation(panneauNav, "📅 Réservations", "btnReservations", 420, 20, 180, 50)
    
    Call CreerBoutonNavigation(panneauNav, "💳 Paiements", "btnPaiements", 20, 90, 180, 50)
    Call CreerBoutonNavigation(panneauNav, "📊 Rapports", "btnRapports", 220, 90, 180, 50)
    Call CreerBoutonNavigation(panneauNav, "⚙️ Paramètres", "btnParametres", 420, 90, 180, 50)
    
    ' Bouton d'actualisation
    Call CreerBoutonNavigation(panneauNav, "🔄 Actualiser", "btnActualiser", 220, 150, 180, 30)
End Sub

Private Sub CreerBoutonNavigation(parent As MSForms.Frame, texte As String, nom As String, _
                                 posX As Integer, posY As Integer, largeur As Integer, hauteur As Integer)
    
    Dim btn As MSForms.CommandButton
    Set btn = parent.Controls.Add("Forms.CommandButton.1", nom)
    
    With btn
        .Top = posY
        .Left = posX
        .Width = largeur
        .Height = hauteur
        .Caption = texte
        .Font.Size = 10
    End With
    
    Call StylerBoutonPrimaire(btn)
End Sub

Private Sub CreerTableauBordRapide()
    ' Panneau des statistiques rapides
    Dim panneauStats As MSForms.Frame
    Set panneauStats = Me.Controls.Add("Forms.Frame.1", "panneauStatistiques")
    
    With panneauStats
        .Top = 330
        .Left = 20
        .Width = 600
        .Height = 100
        .Caption = "📈 Aperçu Rapide"
    End With
    Call StylerPanneauAccent(panneauStats)
    
    ' Statistiques en temps réel
    Call AfficherStatistiqueRapide(panneauStats, "Chambres Libres", CompterChambresLibres(), 20, 30)
    Call AfficherStatistiqueRapide(panneauStats, "Arrivées Aujourd'hui", CompterArriveesDuJour(), 150, 30)
    Call AfficherStatistiqueRapide(panneauStats, "Départs Aujourd'hui", CompterDepartsDuJour(), 280, 30)
    Call AfficherStatistiqueRapide(panneauStats, "Taux Occupation", Format(CalculerTauxOccupationJour(), "0") & "%", 410, 30)
    
    ' Jauge de taux d'occupation
    Call CreerJaugeOccupation(panneauStats, CalculerTauxOccupationJour())
End Sub

Private Sub AfficherStatistiqueRapide(parent As MSForms.Frame, titre As String, valeur As String, _
                                     posX As Integer, posY As Integer)
    
    ' Label titre
    Dim lblTitre As MSForms.Label
    Set lblTitre = parent.Controls.Add("Forms.Label.1", "lbl" & titre)
    
    With lblTitre
        .Top = posY
        .Left = posX
        .Width = 120
        .Height = 15
        .Caption = titre
        .TextAlign = fmTextAlignCenter
    End With
    Call StylerLabelSecondaire(lblTitre)
    
    ' Label valeur
    Dim lblValeur As MSForms.Label
    Set lblValeur = parent.Controls.Add("Forms.Label.1", "lblVal" & titre)
    
    With lblValeur
        .Top = posY + 18
        .Left = posX
        .Width = 120
        .Height = 25
        .Caption = valeur
        .TextAlign = fmTextAlignCenter
        .Font.Size = 14
        .Font.Bold = True
    End With
    Call StylerLabelNormal(lblValeur)
    lblValeur.ForeColor = COLOR_PRIMARY
End Sub

Private Sub CreerJaugeOccupation(parent As MSForms.Frame, taux As Double)
    ' Jauge visuelle du taux d'occupation
    Dim jaugeFond As MSForms.Label
    Set jaugeFond = parent.Controls.Add("Forms.Label.1", "jaugeFond")
    
    With jaugeFond
        .Top = 60
        .Left = 50
        .Width = 500
        .Height = 8
        .BackColor = COLOR_LIGHT_GRAY
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_MEDIUM_GRAY
    End With
    
    ' Barre de progression
    Dim jaugeRemplie As MSForms.Label
    Set jaugeRemplie = parent.Controls.Add("Forms.Label.1", "jaugeRemplie")
    
    With jaugeRemplie
        .Top = 62
        .Left = 52
        .Width = (496 * taux) / 100
        .Height = 4
        .BorderStyle = fmBorderStyleNone
    End With
    
    ' Couleur selon le taux
    If taux < 50 Then
        jaugeRemplie.BackColor = COLOR_SUCCESS
    ElseIf taux < 80 Then
        jaugeRemplie.BackColor = COLOR_WARNING
    Else
        jaugeRemplie.BackColor = COLOR_DANGER
    End If
End Sub

Private Sub CreerPiedPage()
    ' Informations de version et aide
    Dim lblVersion As MSForms.Label
    Set lblVersion = Me.Controls.Add("Forms.Label.1", "lblVersion")
    
    With lblVersion
        .Top = 450
        .Left = 20
        .Width = 300
        .Height = 15
        .Caption = "Gestion Auberge v1.0 - Développé avec Excel VBA"
    End With
    Call StylerLabelSecondaire(lblVersion)
    
    ' Bouton d'aide
    Dim btnAide As MSForms.CommandButton
    Set btnAide = Me.Controls.Add("Forms.CommandButton.1", "btnAide")
    
    With btnAide
        .Top = 445
        .Left = 520
        .Width = 100
        .Height = 25
        .Caption = "❓ Aide"
    End With
    Call StylerBoutonSecondaire(btnAide)
End Sub

' ========================================
' ÉVÉNEMENTS AMÉLIORÉS AVEC ACCESSIBILITÉ
' ========================================

Private Sub btnChambres_Click()
    Call AfficherNotificationAvancee(Me, "Ouverture de la gestion des chambres...", "info", 2)
    Me.Hide
    UserForm_GestionChambres.Show
End Sub

Private Sub btnClients_Click()
    Call AfficherNotificationAvancee(Me, "Ouverture de la gestion clients...", "info", 2)
    Me.Hide
    UserForm_GestionClients.Show
End Sub

Private Sub btnReservations_Click()
    Call AfficherNotificationAvancee(Me, "Ouverture des réservations...", "info", 2)
    Me.Hide
    UserForm_GestionReservations.Show
End Sub

Private Sub btnPaiements_Click()
    Call AfficherNotificationAvancee(Me, "Ouverture des paiements...", "info", 2)
    Me.Hide
    UserForm_GestionPaiements.Show
End Sub

Private Sub btnRapports_Click()
    Call AfficherNotificationAvancee(Me, "Génération des rapports...", "info", 2)
    Me.Hide
    UserForm_Rapports.Show
End Sub

Private Sub btnParametres_Click()
    Call AfficherNotificationAvancee(Me, "Ouverture des paramètres...", "info", 2)
    Me.Hide
    UserForm_Parametres.Show
End Sub

Private Sub btnActualiser_Click()
    Call InitialiserInterface
    Call AfficherNotificationAvancee(Me, "Interface actualisée avec succès !", "success", 3)
End Sub

Private Sub btnAide_Click()
    Call AfficherAideContextuelle(Me)
End Sub

' ========================================
' GESTION DU CLAVIER ET ACCESSIBILITÉ
' ========================================

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub btnChambres_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub btnClients_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub btnReservations_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub btnPaiements_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub btnRapports_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub btnParametres_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub btnActualiser_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

Private Sub btnAide_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call NaviguerAuClavier(Me, KeyCode, Shift)
End Sub

' ========================================
' EFFETS DE SURVOL AMÉLIORÉS
' ========================================

Private Sub btnChambres_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnChambres, True)
    Call AfficherTooltip(Me.btnChambres, True)
End Sub

Private Sub btnClients_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnClients, True)
    Call AfficherTooltip(Me.btnClients, True)
End Sub

Private Sub btnReservations_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnReservations, True)
    Call AfficherTooltip(Me.btnReservations, True)
End Sub

Private Sub btnPaiements_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnPaiements, True)
    Call AfficherTooltip(Me.btnPaiements, True)
End Sub

Private Sub btnRapports_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnRapports, True)
    Call AfficherTooltip(Me.btnRapports, True)
End Sub

Private Sub btnParametres_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnParametres, True)
    Call AfficherTooltip(Me.btnParametres, True)
End Sub

Private Sub btnActualiser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnActualiser, True)
    Call AfficherTooltip(Me.btnActualiser, True)
End Sub

Private Sub btnAide_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnAide, True)
    Call AfficherTooltip(Me.btnAide, True)
End Sub

' Masquer les tooltips quand la souris quitte les boutons
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EffetSurvol(Me.btnChambres, False)
    Call AfficherTooltip(Me.btnChambres, False)
    Call EffetSurvol(Me.btnClients, False)
    Call AfficherTooltip(Me.btnClients, False)
    Call EffetSurvol(Me.btnReservations, False)
    Call AfficherTooltip(Me.btnReservations, False)
    Call EffetSurvol(Me.btnPaiements, False)
    Call AfficherTooltip(Me.btnPaiements, False)
    Call EffetSurvol(Me.btnRapports, False)
    Call AfficherTooltip(Me.btnRapports, False)
    Call EffetSurvol(Me.btnParametres, False)
    Call AfficherTooltip(Me.btnParametres, False)
    Call EffetSurvol(Me.btnActualiser, False)
    Call AfficherTooltip(Me.btnActualiser, False)
    Call EffetSurvol(Me.btnAide, False)
    Call AfficherTooltip(Me.btnAide, False)
End Sub

' ========================================
' FONCTIONS UTILITAIRES
' ========================================

Private Function CompterArriveesDuJour() As Integer
    Dim arrivees As Variant
    arrivees = ObtenirArriveesDuJour()
    If IsArray(arrivees) Then
        If arrivees(0) = "Aucune arrivée aujourd'hui" Then
            CompterArriveesDuJour = 0
        Else
            CompterArriveesDuJour = UBound(arrivees) + 1
        End If
    Else
        CompterArriveesDuJour = 0
    End If
End Function

Private Function CompterDepartsDuJour() As Integer
    Dim departs As Variant
    departs = ObtenirDepartsDuJour()
    If IsArray(departs) Then
        If departs(0) = "Aucun départ aujourd'hui" Then
            CompterDepartsDuJour = 0
        Else
            CompterDepartsDuJour = UBound(departs) + 1
        End If
    Else
        CompterDepartsDuJour = 0
    End If
End Function
