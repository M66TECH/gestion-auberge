Attribute VB_Name = "ComposantsInterfaceAvances"
' ========================================
' MODULE COMPOSANTS D'INTERFACE AVANCÉS
' ========================================
' Description: Composants UX modernes pour une expérience utilisateur professionnelle

Option Explicit

' ========================================
' INDICATEURS DE STATUT INTELLIGENTS
' ========================================

' Créer un indicateur de statut avec animation
Sub CreerIndicateurStatut(frm As Object, statut As String, message As String, Optional posX As Integer = 10, Optional posY As Integer = 10)
    On Error Resume Next

    ' Supprimer l'ancien indicateur s'il existe
    Call SupprimerIndicateurStatut(frm)

    Dim indicateur As Object
    Dim lblMessage As Object
    Dim couleurFond As Long
    Dim couleurTexte As Long
    Dim icone As String

    ' Définir les propriétés selon le statut
    Select Case LCase(statut)
        Case "success", "succès"
            couleurFond = COLOR_SUCCESS
            couleurTexte = COLOR_WHITE
            icone = "✓"
        Case "error", "erreur"
            couleurFond = COLOR_DANGER
            couleurTexte = COLOR_WHITE
            icone = "⚠"
        Case "warning", "avertissement"
            couleurFond = COLOR_WARNING
            couleurTexte = COLOR_WHITE
            icone = "⚠"
        Case "info", "information"
            couleurFond = COLOR_PRIMARY
            couleurTexte = COLOR_WHITE
            icone = "ℹ"
        Case "loading", "chargement"
            couleurFond = COLOR_INFO
            couleurTexte = COLOR_WHITE
            icone = "⟳"
        Case Else
            couleurFond = COLOR_SECONDARY
            couleurTexte = COLOR_TEXT_PRIMARY
            icone = "●"
    End Select

    ' Créer le container
    Set indicateur = frm.Controls.Add("Forms.Frame.1", "frameIndicateurStatut")
    With indicateur
        .Top = posY
        .Left = posX
        .Width = 300
        .Height = 50
        .BackColor = couleurFond
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectRaised
    End With

    ' Créer l'icône
    Set lblMessage = frm.Controls.Add("Forms.Label.1", "lblIconeStatut")
    With lblMessage
        .Top = posY + 10
        .Left = posX + 10
        .Width = 30
        .Height = 25
        .Caption = icone
        .ForeColor = couleurTexte
        .Font.Name = FONT_PRIMARY
        .Font.Size = 12
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
    End With

    ' Créer le message
    Set lblMessage = frm.Controls.Add("Forms.Label.1", "lblMessageStatut")
    With lblMessage
        .Top = posY + 10
        .Left = posX + 45
        .Width = 240
        .Height = 25
        .Caption = message
        .ForeColor = couleurTexte
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .BackStyle = fmBackStyleTransparent
    End With

    ' Animation d'entrée
    Call AnimerGlissementLateral(indicateur, indicateur.Left, "gauche")

    ' Animation spécifique selon le type
    If statut = "loading" Then
        Call AnimerRotation(frm.Controls("lblIconeStatut"))
    End If
End Sub

' Supprimer l'indicateur de statut
Sub SupprimerIndicateurStatut(frm As Object)
    On Error Resume Next

    Dim ctrl As Object
    For Each ctrl In frm.Controls
        If InStr(ctrl.Name, "IndicateurStatut") > 0 Or InStr(ctrl.Name, "Statut") > 0 Then
            frm.Controls.Remove ctrl.Name
        End If
    Next ctrl
End Sub

' ========================================
' CARTES D'INFORMATION MODERNES
' ========================================

' Créer une carte d'information
Sub CreerCarteInformation(frm As Object, titre As String, contenu As String, Optional posX As Integer = 10, Optional posY As Integer = 10)
    On Error Resume Next

    Dim carte As Object
    Dim lblTitre As Object
    Dim lblContenu As Object
    Dim hauteurContenu As Integer

    ' Calculer la hauteur du contenu
    hauteurContenu = Len(contenu) / 50 * 20 + 40 ' Estimation approximative

    ' Créer la carte
    Set carte = frm.Controls.Add("Forms.Frame.1", "frameCarte_" & Replace(titre, " ", "_"))
    With carte
        .Top = posY
        .Left = posX
        .Width = 280
        .Height = hauteurContenu + 60
        .BackColor = COLOR_WHITE
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_BORDER
        .SpecialEffect = fmSpecialEffectSunken
    End With

    ' Titre de la carte
    Set lblTitre = frm.Controls.Add("Forms.Label.1", "lblTitreCarte_" & Replace(titre, " ", "_"))
    With lblTitre
        .Top = posY + 10
        .Left = posX + 15
        .Width = 250
        .Height = 25
        .Caption = titre
        .ForeColor = COLOR_TEXT_PRIMARY
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_HEADING
        .Font.Bold = True
        .BackStyle = fmBackStyleTransparent
    End With

    ' Contenu de la carte
    Set lblContenu = frm.Controls.Add("Forms.Label.1", "lblContenuCarte_" & Replace(titre, " ", "_"))
    With lblContenu
        .Top = posY + 40
        .Left = posX + 15
        .Width = 250
        .Height = hauteurContenu
        .Caption = contenu
        .ForeColor = COLOR_TEXT_SECONDARY
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .BackStyle = fmBackStyleTransparent
        .WordWrap = True
    End With

    ' Effet d'ombre simulé
    Call AjouterOmbreCarte(carte)
End Sub

' Ajouter un effet d'ombre à une carte
Sub AjouterOmbreCarte(carte As Object)
    On Error Resume Next

    Dim ombre As Object
    Set ombre = carte.Parent.Controls.Add("Forms.Frame.1", "ombre_" & carte.Name)
    With ombre
        .Top = carte.Top + 3
        .Left = carte.Left + 3
        .Width = carte.Width
        .Height = carte.Height
        .BackColor = COLOR_TEXT_MUTED
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With

    ' Envoyer l'ombre derrière la carte
    ombre.ZOrder fmSendToBack
End Sub

' ========================================
' BOUTONS D'ACTION AVANCÉS
' ========================================

' Créer un bouton d'action flottant (FAB)
Sub CreerBoutonActionFlottant(frm As Object, icone As String, tooltip As String, Optional posX As Integer = 0, Optional posY As Integer = 0)
    On Error Resume Next

    Dim fab As Object
    Dim lblTooltip As Object

    ' Position par défaut
    If posX = 0 Then posX = frm.Width - 70
    If posY = 0 Then posY = frm.Height - 70

    ' Créer le bouton FAB
    Set fab = frm.Controls.Add("Forms.CommandButton.1", "btnFAB")
    With fab
        .Top = posY
        .Left = posX
        .Width = 50
        .Height = 50
        .Caption = icone
        .Font.Size = 14
        .Font.Bold = True
        .BackColor = COLOR_ACCENT
        .ForeColor = COLOR_WHITE
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .PicturePosition = fmPicturePositionCenter
    End With

    ' Créer le tooltip
    Set lblTooltip = frm.Controls.Add("Forms.Label.1", "lblTooltipFAB")
    With lblTooltip
        .Top = posY - 30
        .Left = posX - 50
        .Width = 150
        .Height = 25
        .Caption = tooltip
        .BackColor = COLOR_TEXT_MUTED
        .ForeColor = COLOR_WHITE
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_TEXT_MUTED
        .TextAlign = fmTextAlignCenter
        .Visible = False
    End With

    ' Stocker la référence du tooltip
    fab.Tag = lblTooltip

    ' Ajouter les événements
    ' Note: Les événements doivent être ajoutés dans le formulaire
End Sub

' ========================================
' GRAPHIQUES ET VISUALISATIONS
' ========================================

' Créer un graphique circulaire simple
Sub CreerGraphiqueCirculaire(frm As Object, valeurs As Variant, labels As Variant, Optional posX As Integer = 10, Optional posY As Integer = 10)
    On Error Resume Next

    Dim i As Integer
    Dim total As Double
    Dim angleDebut As Double
    Dim angleFin As Double
    Dim couleurs As Variant

    ' Couleurs pour les segments
    couleurs = Array(COLOR_PRIMARY, COLOR_SUCCESS, COLOR_WARNING, COLOR_DANGER, COLOR_INFO)

    ' Calculer le total
    total = 0
    For i = LBound(valeurs) To UBound(valeurs)
        total = total + valeurs(i)
    Next i

    ' Créer les segments
    angleDebut = 0
    For i = LBound(valeurs) To UBound(valeurs)
        angleFin = angleDebut + (valeurs(i) / total) * 360

        Call CreerSegmentCirculaire(frm, posX, posY, 50, angleDebut, angleFin, couleurs(i Mod 5))

        ' Ajouter le label
        Call CreerLabelSegment(frm, labels(i), valeurs(i), total, posX, posY, angleDebut, angleFin, couleurs(i Mod 5))

        angleDebut = angleFin
    Next i
End Sub

' Créer un segment circulaire (simulation avec formes)
Sub CreerSegmentCirculaire(frm As Object, centreX As Integer, centreY As Integer, rayon As Integer, angleDebut As Double, angleFin As Double, couleur As Long)
    On Error Resume Next

    Dim segment As Object
    Set segment = frm.Controls.Add("Forms.Frame.1", "segment_" & angleDebut & "_" & angleFin)

    With segment
        .Top = centreY - rayon
        .Left = centreX - rayon
        .Width = rayon * 2
        .Height = rayon * 2
        .BackColor = couleur
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_TEXT_MUTED
        .SpecialEffect = fmSpecialEffectRaised
    End With
End Sub

' Créer un label pour un segment
Sub CreerLabelSegment(frm As Object, label As String, valeur As Double, total As Double, centreX As Integer, centreY As Integer, angleDebut As Double, angleFin As Double, couleur As Long)
    On Error Resume Next

    Dim lbl As Object
    Dim angleMilieu As Double
    Dim posX As Integer
    Dim posY As Integer

    angleMilieu = (angleDebut + angleFin) / 2
    posX = centreX + (60 * Cos(angleMilieu * 3.14159 / 180)) - 30
    posY = centreY + (60 * Sin(angleMilieu * 3.14159 / 180)) - 15

    Set lbl = frm.Controls.Add("Forms.Label.1", "lblSegment_" & label)
    With lbl
        .Top = posY
        .Left = posX
        .Width = 60
        .Height = 30
        .Caption = label & vbCrLf & Format(valeur / total * 100, "0") & "%"
        .ForeColor = COLOR_WHITE
        .Font.Size = 8
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleOpaque
        .BackColor = couleur
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_WHITE
    End With
End Sub

' ========================================
' LISTES ET GRILLES AMÉLIORÉES
' ========================================

' Créer une liste avec style moderne
Sub CreerListeModerne(frm As Object, donnees As Variant, enTetes As Variant, Optional posX As Integer = 10, Optional posY As Integer = 10)
    On Error Resume Next

    Dim i As Integer, j As Integer
    Dim lblEnTete As Object
    Dim lblDonnee As Object
    Dim ligne As Object

    ' Créer les en-têtes
    For i = LBound(enTetes) To UBound(enTetes)
        Set lblEnTete = frm.Controls.Add("Forms.Label.1", "entete_" & i)
        With lblEnTete
            .Top = posY
            .Left = posX + i * 100
            .Width = 95
            .Height = 25
            .Caption = enTetes(i)
            .ForeColor = COLOR_WHITE
            .Font.Bold = True
            .TextAlign = fmTextAlignCenter
            .BackColor = COLOR_PRIMARY
            .BorderStyle = fmBorderStyleSingle
        End With
    Next i

    ' Créer les données
    For i = LBound(donnees) To UBound(donnees)
        ' Ligne de fond alternée
        If i Mod 2 = 0 Then
            Set ligne = frm.Controls.Add("Forms.Frame.1", "ligne_" & i)
            With ligne
                .Top = posY + 25 + i * 25
                .Left = posX
                .Width = 100 * (UBound(enTetes) - LBound(enTetes) + 1)
                .Height = 25
                .BackColor = COLOR_BACKGROUND_HOVER
                .BorderStyle = fmBorderStyleNone
            End With
        End If

        For j = LBound(donnees(i)) To UBound(donnees(i))
            Set lblDonnee = frm.Controls.Add("Forms.Label.1", "donnee_" & i & "_" & j)
            With lblDonnee
                .Top = posY + 25 + i * 25
                .Left = posX + j * 100
                .Width = 95
                .Height = 20
                .Caption = donnees(i)(j)
                .ForeColor = COLOR_TEXT_PRIMARY
                .TextAlign = fmTextAlignLeft
                .BackStyle = fmBackStyleTransparent
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = COLOR_BORDER
            End With
        Next j
    Next i
End Sub

' ========================================
' MENUS CONTEXTUELS
' ========================================

' Créer un menu contextuel
Sub CreerMenuContextuel(frm As Object, options As Variant, posX As Integer, posY As Integer)
    On Error Resume Next

    Dim menu As Object
    Dim lblOption As Object
    Dim i As Integer

    ' Créer le container du menu
    Set menu = frm.Controls.Add("Forms.Frame.1", "frameMenuContextuel")
    With menu
        .Top = posY
        .Left = posX
        .Width = 150
        .Height = 25 * UBound(options) + 10
        .BackColor = COLOR_WHITE
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_TEXT_MUTED
        .SpecialEffect = fmSpecialEffectRaised
    End With

    ' Créer les options
    For i = LBound(options) To UBound(options)
        Set lblOption = frm.Controls.Add("Forms.Label.1", "option_" & i)
        With lblOption
            .Top = posY + 5 + i * 25
            .Left = posX + 5
            .Width = 140
            .Height = 20
            .Caption = options(i)
            .ForeColor = COLOR_TEXT_PRIMARY
            .Font.Size = FONT_SIZE_BODY
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleNone
            .MousePointer = fmMousePointerHand
        End With

        ' Stocker l'index de l'option
        lblOption.Tag = i
    Next i

    ' Animer l'apparition
    Call AnimerFonduEntree(menu)
End Sub

' ========================================
' MODALES ET DIALOGUES
' ========================================

' Créer une modale personnalisée
Sub CreerModale(frm As Object, titre As String, contenu As String, boutons As Variant)
    On Error Resume Next

    Dim modale As Object
    Dim fond As Object
    Dim lblTitre As Object
    Dim lblContenu As Object
    Dim btn As Object
    Dim i As Integer

    ' Fond semi-transparent
    Set fond = frm.Controls.Add("Forms.Frame.1", "fondModale")
    With fond
        .Top = 0
        .Left = 0
        .Width = frm.Width
        .Height = frm.Height
        .BackColor = COLOR_TEXT_MUTED
        .BorderStyle = fmBorderStyleNone
    End With

    ' Container de la modale
    Set modale = frm.Controls.Add("Forms.Frame.1", "frameModale")
    With modale
        .Top = (frm.Height - 200) / 2
        .Left = (frm.Width - 400) / 2
        .Width = 400
        .Height = 200
        .BackColor = COLOR_WHITE
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_TEXT_MUTED
        .SpecialEffect = fmSpecialEffectRaised
    End With

    ' Titre
    Set lblTitre = frm.Controls.Add("Forms.Label.1", "lblTitreModale")
    With lblTitre
        .Top = modale.Top + 10
        .Left = modale.Left + 15
        .Width = 370
        .Height = 30
        .Caption = titre
        .ForeColor = COLOR_TEXT_PRIMARY
        .Font.Size = FONT_SIZE_HEADING
        .Font.Bold = True
        .BackStyle = fmBackStyleTransparent
    End With

    ' Contenu
    Set lblContenu = frm.Controls.Add("Forms.Label.1", "lblContenuModale")
    With lblContenu
        .Top = modale.Top + 50
        .Left = modale.Left + 15
        .Width = 370
        .Height = 60
        .Caption = contenu
        .ForeColor = COLOR_TEXT_SECONDARY
        .Font.Size = FONT_SIZE_BODY
        .BackStyle = fmBackStyleTransparent
        .WordWrap = True
    End With

    ' Boutons
    For i = LBound(boutons) To UBound(boutons)
        Set btn = frm.Controls.Add("Forms.CommandButton.1", "btnModale_" & i)
        With btn
            .Top = modale.Top + 130
            .Left = modale.Left + 15 + i * 120
            .Width = 100
            .Height = 30
            .Caption = boutons(i)
            .BackColor = COLOR_PRIMARY
            .ForeColor = COLOR_WHITE
            .Font.Bold = True
        End With
    Next i

    ' Envoyer le fond derrière
    fond.ZOrder fmSendToBack

    ' Animation d'entrée
    Call AnimerFonduEntree(modale)
End Sub

' ========================================
' ANIMATIONS ET TRANSITIONS
' ========================================

' Animation de rebond
Sub AnimerRebond(ctrl As Object, Optional amplitude As Integer = 10)
    On Error Resume Next

    Dim positionOriginale As Integer
    Dim i As Integer

    positionOriginale = ctrl.Top

    For i = 1 To 8
        ctrl.Top = positionOriginale - amplitude + (amplitude / 4) * i
        DoEvents
        Sleep 50
    Next i

    ctrl.Top = positionOriginale
End Sub

' Animation de secousse
Sub AnimerSecousse(ctrl As Object)
    On Error Resume Next

    Dim positionOriginale As Integer
    Dim i As Integer

    positionOriginale = ctrl.Left

    For i = 1 To 10
        If i Mod 2 = 0 Then
            ctrl.Left = positionOriginale + 5
        Else
            ctrl.Left = positionOriginale - 5
        End If
        DoEvents
        Sleep 30
    Next i

    ctrl.Left = positionOriginale
End Sub
