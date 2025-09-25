VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_DashboardModerne 
   Caption         =   "📊 Dashboard Moderne - Gestion Auberge"
   ClientHeight    =   650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   900
   OleObjectBlob   =   "UserForm_DashboardModerne.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_DashboardModerne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ========================================
' DASHBOARD MODERNE AVEC GRAPHIQUES ET JAUGES
' ========================================

Private Sub UserForm_Initialize()
    Call InitialiserDashboardModerne
End Sub

Private Sub InitialiserDashboardModerne()
    Call AppliquerStyleModerne(Me)
    
    Me.Width = 900
    Me.Height = 650
    Me.Caption = "📊 Dashboard Moderne - Gestion Auberge"
    
    Call CreerEnTeteDashboard
    Call CreerJaugesPerformance
    Call CreerGraphiquesVisuels
    Call CreerTableauActivites
    Call CreerBarreNavigationRapide
End Sub

Private Sub CreerEnTeteDashboard()
    ' Titre avec gradient visuel
    Dim lblTitre As MSForms.Label
    Set lblTitre = Me.Controls.Add("Forms.Label.1", "lblTitreDashboard")
    
    With lblTitre
        .Top = 10
        .Left = 20
        .Width = 860
        .Height = 40
        .Caption = "📊 DASHBOARD AUBERGE - VUE D'ENSEMBLE"
        .TextAlign = fmTextAlignCenter
        .Font.Size = 18
        .Font.Bold = True
        .ForeColor = COLOR_PRIMARY
        .BackColor = COLOR_LIGHT_GRAY
    End With
    
    ' Informations temps réel
    Dim lblTempsReel As MSForms.Label
    Set lblTempsReel = Me.Controls.Add("Forms.Label.1", "lblTempsReel")
    
    With lblTempsReel
        .Top = 55
        .Left = 20
        .Width = 860
        .Height = 20
        .Caption = "🕒 Dernière mise à jour : " & Format(Now, "dd/mm/yyyy hh:mm") & " | 🌡️ Statut système : Opérationnel"
        .TextAlign = fmTextAlignCenter
    End With
    Call StylerLabelSecondaire(lblTempsReel)
End Sub

Private Sub CreerJaugesPerformance()
    ' Panneau des jauges
    Dim panneauJauges As MSForms.Frame
    Set panneauJauges = Me.Controls.Add("Forms.Frame.1", "panneauJauges")
    
    With panneauJauges
        .Top = 85
        .Left = 20
        .Width = 860
        .Height = 120
        .Caption = "  📈 Indicateurs de Performance"
        .Font.Bold = True
    End With
    Call StylerPanneauAccent(panneauJauges)
    
    ' Jauge taux d'occupation
    Call CreerJaugeCirculaire(panneauJauges, CalculerTauxOccupationJour(), 100, "Taux Occupation")
    
    ' Jauge revenus du jour
    Call CreerJaugeLineaire(panneauJauges, "Revenus Jour", "1,250€", "2,000€", 62, 200, 30)
    
    ' Jauge satisfaction client (simulée)
    Call CreerJaugeLineaire(panneauJauges, "Satisfaction", "4.2/5", "5.0", 84, 450, 30)
    
    ' Compteurs rapides
    Call CreerCompteurRapide(panneauJauges, "🛏️", CompterChambresLibres(), "Chambres Libres", 650, 30)
    Call CreerCompteurRapide(panneauJauges, "👥", CompterArriveesDuJour(), "Arrivées", 750, 30)
End Sub

Private Sub CreerJaugeLineaire(parent As MSForms.Frame, titre As String, valeurActuelle As String, _
                              valeurMax As String, pourcentage As Integer, posX As Integer, posY As Integer)
    
    ' Titre de la jauge
    Dim lblTitre As MSForms.Label
    Set lblTitre = parent.Controls.Add("Forms.Label.1", "lblTitre" & titre)
    
    With lblTitre
        .Top = posY
        .Left = posX
        .Width = 150
        .Height = 15
        .Caption = titre
        .TextAlign = fmTextAlignCenter
        .Font.Size = 9
        .Font.Bold = True
    End With
    
    ' Barre de fond
    Dim barreFond As MSForms.Label
    Set barreFond = parent.Controls.Add("Forms.Label.1", "barreFond" & titre)
    
    With barreFond
        .Top = posY + 20
        .Left = posX
        .Width = 150
        .Height = 8
        .BackColor = COLOR_LIGHT_GRAY
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_MEDIUM_GRAY
    End With
    
    ' Barre de progression
    Dim barreProgres As MSForms.Label
    Set barreProgres = parent.Controls.Add("Forms.Label.1", "barreProgres" & titre)
    
    With barreProgres
        .Top = posY + 22
        .Left = posX + 2
        .Width = (146 * pourcentage) / 100
        .Height = 4
        .BorderStyle = fmBorderStyleNone
    End With
    
    ' Couleur selon performance
    If pourcentage >= 80 Then
        barreProgres.BackColor = COLOR_SUCCESS
    ElseIf pourcentage >= 60 Then
        barreProgres.BackColor = COLOR_WARNING
    Else
        barreProgres.BackColor = COLOR_DANGER
    End If
    
    ' Valeurs
    Dim lblValeurs As MSForms.Label
    Set lblValeurs = parent.Controls.Add("Forms.Label.1", "lblValeurs" & titre)
    
    With lblValeurs
        .Top = posY + 35
        .Left = posX
        .Width = 150
        .Height = 15
        .Caption = valeurActuelle & " / " & valeurMax
        .TextAlign = fmTextAlignCenter
        .Font.Size = 8
    End With
    Call StylerLabelSecondaire(lblValeurs)
End Sub

Private Sub CreerCompteurRapide(parent As MSForms.Frame, icone As String, valeur As Integer, _
                               titre As String, posX As Integer, posY As Integer)
    
    ' Icône
    Dim lblIcone As MSForms.Label
    Set lblIcone = parent.Controls.Add("Forms.Label.1", "lblIcone" & titre)
    
    With lblIcone
        .Top = posY
        .Left = posX
        .Width = 30
        .Height = 30
        .Caption = icone
        .TextAlign = fmTextAlignCenter
        .Font.Size = 16
    End With
    
    ' Valeur
    Dim lblValeur As MSForms.Label
    Set lblValeur = parent.Controls.Add("Forms.Label.1", "lblValeur" & titre)
    
    With lblValeur
        .Top = posY + 5
        .Left = posX + 35
        .Width = 40
        .Height = 20
        .Caption = valeur
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = COLOR_PRIMARY
    End With
    
    ' Titre
    Dim lblTitre As MSForms.Label
    Set lblTitre = parent.Controls.Add("Forms.Label.1", "lblTitre" & titre)
    
    With lblTitre
        .Top = posY + 35
        .Left = posX
        .Width = 75
        .Height = 15
        .Caption = titre
        .TextAlign = fmTextAlignCenter
        .Font.Size = 8
    End With
    Call StylerLabelSecondaire(lblTitre)
End Sub

Private Sub CreerGraphiquesVisuels()
    ' Panneau graphiques
    Dim panneauGraphiques As MSForms.Frame
    Set panneauGraphiques = Me.Controls.Add("Forms.Frame.1", "panneauGraphiques")
    
    With panneauGraphiques
        .Top = 220
        .Left = 20
        .Width = 860
        .Height = 200
        .Caption = "  📊 Analyse Visuelle"
        .Font.Bold = True
    End With
    Call StylerPanneau(panneauGraphiques)
    
    ' Graphique occupation (simulation ASCII)
    Call CreerGraphiqueOccupation(panneauGraphiques, 20, 30)
    
    ' Graphique revenus
    Call CreerGraphiqueRevenus(panneauGraphiques, 450, 30)
End Sub

Private Sub CreerGraphiqueOccupation(parent As MSForms.Frame, posX As Integer, posY As Integer)
    Dim lblTitreGraph As MSForms.Label
    Set lblTitreGraph = parent.Controls.Add("Forms.Label.1", "lblTitreOccupation")
    
    With lblTitreGraph
        .Top = posY
        .Left = posX
        .Width = 400
        .Height = 20
        .Caption = "📈 Taux d'Occupation - 7 Derniers Jours"
        .Font.Bold = True
    End With
    
    ' Simulation d'un graphique avec des barres ASCII
    Dim txtGraphique As MSForms.TextBox
    Set txtGraphique = parent.Controls.Add("Forms.TextBox.1", "txtGraphiqueOccupation")
    
    With txtGraphique
        .Top = posY + 25
        .Left = posX
        .Width = 400
        .Height = 120
        .MultiLine = True
        .Font.Name = FONT_MONOSPACE
        .Font.Size = 8
        .Locked = True
        .BackColor = COLOR_WHITE
    End With
    
    ' Données simulées
    txtGraphique.Value = "Lun  ████████████████░░░░  80%" & vbCrLf & _
                        "Mar  ██████████████████░░  90%" & vbCrLf & _
                        "Mer  ████████████░░░░░░░░  60%" & vbCrLf & _
                        "Jeu  ██████████████████░░  90%" & vbCrLf & _
                        "Ven  ████████████████████  100%" & vbCrLf & _
                        "Sam  ████████████████████  100%" & vbCrLf & _
                        "Dim  ██████████████░░░░░░  70%"
End Sub

Private Sub CreerGraphiqueRevenus(parent As MSForms.Frame, posX As Integer, posY As Integer)
    Dim lblTitreRev As MSForms.Label
    Set lblTitreRev = parent.Controls.Add("Forms.Label.1", "lblTitreRevenus")
    
    With lblTitreRev
        .Top = posY
        .Left = posX
        .Width = 380
        .Height = 20
        .Caption = "💰 Évolution des Revenus - Tendance Mensuelle"
        .Font.Bold = True
    End With
    
    ' Graphique linéaire simulé
    Dim txtRevGraph As MSForms.TextBox
    Set txtRevGraph = parent.Controls.Add("Forms.TextBox.1", "txtGraphiqueRevenus")
    
    With txtRevGraph
        .Top = posY + 25
        .Left = posX
        .Width = 380
        .Height = 120
        .MultiLine = True
        .Font.Name = FONT_MONOSPACE
        .Font.Size = 8
        .Locked = True
        .BackColor = COLOR_WHITE
    End With
    
    txtRevGraph.Value = "Semaine 1:  1,200€  ▲ +5%" & vbCrLf & _
                       "Semaine 2:  1,450€  ▲ +20%" & vbCrLf & _
                       "Semaine 3:  1,380€  ▼ -5%" & vbCrLf & _
                       "Semaine 4:  1,620€  ▲ +17%" & vbCrLf & vbCrLf & _
                       "📊 Moyenne: 1,412€/semaine" & vbCrLf & _
                       "🎯 Objectif: 1,500€/semaine"
End Sub

Private Sub CreerTableauActivites()
    ' Panneau activités récentes
    Dim panneauActivites As MSForms.Frame
    Set panneauActivites = Me.Controls.Add("Forms.Frame.1", "panneauActivites")
    
    With panneauActivites
        .Top = 440
        .Left = 20
        .Width = 860
        .Height = 140
        .Caption = "  🔔 Activités Récentes & Alertes"
        .Font.Bold = True
    End With
    Call StylerPanneauAccent(panneauActivites)
    
    ' Liste des activités
    Dim lstActivites As MSForms.ListBox
    Set lstActivites = panneauActivites.Controls.Add("Forms.ListBox.1", "lstActivitesRecentes")
    
    With lstActivites
        .Top = 25
        .Left = 15
        .Width = 830
        .Height = 90
        .Font.Name = FONT_PRIMARY
        .Font.Size = 9
    End With
    Call StylerChampTexte(lstActivites)
    
    ' Charger activités simulées
    Call ChargerActivitesRecentes(lstActivites)
End Sub

Private Sub ChargerActivitesRecentes(lst As MSForms.ListBox)
    lst.AddItem "🟢 " & Format(Now - 0.1, "hh:mm") & " - Nouvelle réservation: Dupont Jean, Chambre 201"
    lst.AddItem "🔵 " & Format(Now - 0.2, "hh:mm") & " - Paiement reçu: 450€ pour réservation #123"
    lst.AddItem "🟡 " & Format(Now - 0.3, "hh:mm") & " - Check-in effectué: Chambre 103, Martin Marie"
    lst.AddItem "🟠 " & Format(Now - 0.5, "hh:mm") & " - Alerte: Chambre 205 nécessite maintenance"
    lst.AddItem "🔴 " & Format(Now - 1, "hh:mm") & " - Annulation: Réservation #118 annulée par client"
End Sub

Private Sub CreerBarreNavigationRapide()
    ' Barre de navigation
    Dim panneauNav As MSForms.Frame
    Set panneauNav = Me.Controls.Add("Forms.Frame.1", "panneauNavigation")
    
    With panneauNav
        .Top = 595
        .Left = 20
        .Width = 860
        .Height = 45
        .Caption = ""
    End With
    Call StylerPanneau(panneauNav)
    
    ' Boutons navigation rapide
    Call CreerBoutonNavRapide(panneauNav, "🏠 Accueil", "btnAccueil", 20, 10, 100)
    Call CreerBoutonNavRapide(panneauNav, "📅 Réservations", "btnNavReservations", 140, 10, 120)
    Call CreerBoutonNavRapide(panneauNav, "👤 Clients", "btnNavClients", 280, 10, 100)
    Call CreerBoutonNavRapide(panneauNav, "💳 Paiements", "btnNavPaiements", 400, 10, 100)
    Call CreerBoutonNavRapide(panneauNav, "📊 Rapports", "btnNavRapports", 520, 10, 100)
    Call CreerBoutonNavRapide(panneauNav, "🔄 Actualiser", "btnActualiserDash", 640, 10, 100)
    Call CreerBoutonNavRapide(panneauNav, "⚙️ Config", "btnConfig", 760, 10, 80)
End Sub

Private Sub CreerBoutonNavRapide(parent As MSForms.Frame, texte As String, nom As String, _
                                posX As Integer, posY As Integer, largeur As Integer)
    
    Dim btn As MSForms.CommandButton
    Set btn = parent.Controls.Add("Forms.CommandButton.1", nom)
    
    With btn
        .Top = posY
        .Left = posX
        .Width = largeur
        .Height = 25
        .Caption = texte
        .Font.Size = 8
    End With
    Call StylerBoutonPrimaire(btn)
End Sub

' ========================================
' ÉVÉNEMENTS DE NAVIGATION
' ========================================

Private Sub btnAccueil_Click()
    Me.Hide
    UserForm_MenuPrincipal.Show
End Sub

Private Sub btnNavReservations_Click()
    Me.Hide
    UserForm_GestionReservations.Show
End Sub

Private Sub btnNavClients_Click()
    Me.Hide
    UserForm_GestionClients.Show
End Sub

Private Sub btnActualiserDash_Click()
    Call InitialiserDashboardModerne
    Call AfficherMessageSucces(Me, "Dashboard actualisé !")
End Sub
