Attribute VB_Name = "DesignSystem"
' ========================================
' MODULE DESIGN SYSTEM - UX MODERNE
' ========================================
' Description: Système de design cohérent pour toutes les UserForms

Option Explicit

' ========================================
' PALETTE DE COULEURS MODERNE
' ========================================
Public Const COLOR_PRIMARY As Long = &H8B4513        ' Bleu nuit élégant
Public Const COLOR_SECONDARY As Long = &HF5F5DC      ' Beige doux
Public Const COLOR_ACCENT As Long = &H6B8E23         ' Vert olive
Public Const COLOR_SUCCESS As Long = &H228B22        ' Vert succès
Public Const COLOR_WARNING As Long = &H1E90FF        ' Orange attention
Public Const COLOR_DANGER As Long = &H4169E1         ' Rouge erreur
Public Const COLOR_WHITE As Long = &HFFFFFF          ' Blanc pur
Public Const COLOR_LIGHT_GRAY As Long = &HF8F9FA     ' Gris très clair
Public Const COLOR_MEDIUM_GRAY As Long = &HE9ECEF    ' Gris moyen
Public Const COLOR_DARK_GRAY As Long = &H6C757D      ' Gris foncé
Public Const COLOR_TEXT_MUTED As Long = &H868E96      ' Texte atténué pour accessibilité
Public Const COLOR_BORDER_FOCUS As Long = &H007BFF    ' Couleur de focus pour accessibilité
Public Const COLOR_BACKGROUND_HOVER As Long = &HF8F9FA ' Fond au survol

' ========================================
' ACCESSIBILITÉ ET NAVIGATION CLAVIER
' ========================================
Public Const FOCUS_BORDER_WIDTH As Integer = 2
Public Const MIN_CLICKABLE_SIZE As Integer = 44        ' Taille minimum recommandée WCAG
Public Const CONTRAST_RATIO_MIN As Double = 4.5       ' Ratio de contraste minimum

' Gestionnaire de focus global
Private focusActif As Object
Private ordreTabOriginal As Collection

' ========================================
' POLICES MODERNES
' ========================================
Public Const FONT_PRIMARY As String = "Segoe UI"
Public Const FONT_SECONDARY As String = "Calibri"
Public Const FONT_MONOSPACE As String = "Consolas"

' Tailles de police
Public Const FONT_SIZE_TITLE As Integer = 16
Public Const FONT_SIZE_SUBTITLE As Integer = 12
Public Const FONT_SIZE_BODY As Integer = 10
Public Const FONT_SIZE_SMALL As Integer = 8

' ========================================
' DIMENSIONS ET ESPACEMENTS
' ========================================
Public Const SPACING_XS As Integer = 4
Public Const SPACING_SM As Integer = 8
Public Const SPACING_MD As Integer = 16
Public Const SPACING_LG As Integer = 24
Public Const SPACING_XL As Integer = 32

Public Const BUTTON_HEIGHT As Integer = 32
Public Const INPUT_HEIGHT As Integer = 24
Public Const FORM_PADDING As Integer = 20

' ========================================
' APPLIQUER LE STYLE MODERNE À UN USERFORM
' ========================================
Sub AppliquerStyleModerne(frm As Object)
    On Error Resume Next
    
    With frm
        .BackColor = COLOR_WHITE
        .BorderStyle = fmBorderStyleSingle
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .ForeColor = COLOR_TEXT_PRIMARY
    End With
End Sub

' ========================================
' CRÉER UN BOUTON MODERNE
' ========================================
Sub StylerBoutonPrimaire(btn As Object)
    On Error Resume Next
    
    With btn
        .BackColor = COLOR_PRIMARY
        .ForeColor = COLOR_WHITE
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = True
        .Height = BUTTON_HEIGHT
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_PRIMARY
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

Sub StylerBoutonSecondaire(btn As Object)
    On Error Resume Next
    
    With btn
        .BackColor = COLOR_SECONDARY
        .ForeColor = COLOR_TEXT_PRIMARY
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .Height = BUTTON_HEIGHT
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_MEDIUM_GRAY
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

Sub StylerBoutonSucces(btn As Object)
    On Error Resume Next
    
    With btn
        .BackColor = COLOR_SUCCESS
        .ForeColor = COLOR_WHITE
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = True
        .Height = BUTTON_HEIGHT
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_SUCCESS
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

Sub StylerBoutonDanger(btn As Object)
    On Error Resume Next
    
    With btn
        .BackColor = COLOR_DANGER
        .ForeColor = COLOR_WHITE
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = True
        .Height = BUTTON_HEIGHT
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_DANGER
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

' ========================================
' STYLER LES CHAMPS DE SAISIE
' ========================================
Sub StylerChampTexte(txt As Object)
    On Error Resume Next
    
    With txt
        .BackColor = COLOR_WHITE
        .ForeColor = COLOR_TEXT_PRIMARY
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .Height = INPUT_HEIGHT
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_MEDIUM_GRAY
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

Sub StylerComboBox(cmb As Object)
    On Error Resume Next
    
    With cmb
        .BackColor = COLOR_WHITE
        .ForeColor = COLOR_TEXT_PRIMARY
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .Height = INPUT_HEIGHT
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_MEDIUM_GRAY
        .SpecialEffect = fmSpecialEffectFlat
        .Style = fmStyleDropDownCombo
    End With
End Sub

' ========================================
' STYLER LES LABELS
' ========================================
Sub StylerTitre(lbl As Object)
    On Error Resume Next
    
    With lbl
        .ForeColor = COLOR_PRIMARY
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_TITLE
        .Font.Bold = True
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

Sub StylerSousTitre(lbl As Object)
    On Error Resume Next
    
    With lbl
        .ForeColor = COLOR_TEXT_PRIMARY
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_SUBTITLE
        .Font.Bold = True
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

Sub StylerLabelNormal(lbl As Object)
    On Error Resume Next
    
    With lbl
        .ForeColor = COLOR_TEXT_PRIMARY
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

Sub StylerLabelSecondaire(lbl As Object)
    On Error Resume Next
    
    With lbl
        .ForeColor = COLOR_TEXT_SECONDARY
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_SMALL
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

' ========================================
' CRÉER DES PANNEAUX AVEC OMBRE
' ========================================
Sub StylerPanneau(frm As Object)
    On Error Resume Next
    
    With frm
        .BackColor = COLOR_WHITE
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_MEDIUM_GRAY
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

Sub StylerPanneauAccent(frm As Object)
    On Error Resume Next
    
    With frm
        .BackColor = COLOR_LIGHT_GRAY
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_ACCENT
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

' ========================================
' ANIMATIONS ET EFFETS VISUELS
' ========================================
Sub EffetSurvol(ctrl As Object, survol As Boolean)
    On Error Resume Next
    
    If survol Then
        ' Effet de survol - assombrir légèrement
        Select Case ctrl.BackColor
            Case COLOR_PRIMARY
                ctrl.BackColor = &H7A3F0F ' Bleu plus foncé
            Case COLOR_SUCCESS
                ctrl.BackColor = &H1E7B1E ' Vert plus foncé
            Case COLOR_SECONDARY
                ctrl.BackColor = &HE5E5DC ' Beige plus foncé
        End Select
    Else
        ' Retour à la couleur normale
        Select Case ctrl.BackColor
            Case &H7A3F0F
                ctrl.BackColor = COLOR_PRIMARY
            Case &H1E7B1E
                ctrl.BackColor = COLOR_SUCCESS
            Case &HE5E5DC
                ctrl.BackColor = COLOR_SECONDARY
        End Select
    End If
End Sub

' ========================================
' MESSAGES DE FEEDBACK VISUELS AMÉLIORÉS
' ========================================
Sub AfficherMessageSucces(frm As Object, message As String)
    On Error Resume Next
    
    ' Nettoyer les anciens messages
    Call NettoyerMessagesTemporaires(frm)
    
    ' Créer ou mettre à jour le label de message
    Dim lblMessage As Object
    Set lblMessage = frm.Controls("lblMessageFeedback")
    If lblMessage Is Nothing Then
        Set lblMessage = frm.Controls.Add("Forms.Label.1", "lblMessageFeedback")
        lblMessage.Top = frm.Height - 60
        lblMessage.Left = FORM_PADDING
        lblMessage.Width = frm.Width - (FORM_PADDING * 2)
        lblMessage.Height = 25
        lblMessage.TextAlign = fmTextAlignCenter
    End If
    
    With lblMessage
        .Caption = "✓ " & message
        .ForeColor = COLOR_WHITE
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = True
        .BackColor = COLOR_SUCCESS
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_SUCCESS
        .Visible = True
    End With
    
    ' Animation d'entrée
    Call AnimerGlissementLateral(lblMessage, lblMessage.Left, "gauche")
    
    ' Auto-masquer après 4 secondes
    Application.OnTime Now + TimeValue("00:00:04"), "MasquerMessageFeedback['" & frm.Name & "']"
End Sub

Sub AfficherMessageErreur(frm As Object, message As String)
    On Error Resume Next
    
    ' Nettoyer les anciens messages
    Call NettoyerMessagesTemporaires(frm)
    
    Dim lblMessage As Object
    Set lblMessage = frm.Controls("lblMessageFeedback")
    If lblMessage Is Nothing Then
        Set lblMessage = frm.Controls.Add("Forms.Label.1", "lblMessageFeedback")
        lblMessage.Top = frm.Height - 60
        lblMessage.Left = frm.Width - (FORM_PADDING * 2) - 300
        lblMessage.Width = 300
        lblMessage.Height = 25
        lblMessage.TextAlign = fmTextAlignCenter
    End If
    
    With lblMessage
        .Caption = "⚠ " & message
        .ForeColor = COLOR_WHITE
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = True
        .BackColor = COLOR_DANGER
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_DANGER
        .Visible = True
    End With
    
    ' Animation d'entrée avec tremblement
    Call AnimerGlissementLateral(lblMessage, lblMessage.Left, "gauche")
    Call EffetValidationEchec(lblMessage)
    
    ' Auto-masquer après 6 secondes
    Application.OnTime Now + TimeValue("00:00:06"), "MasquerMessageFeedback['" & frm.Name & "']"
End Sub

Sub AfficherMessageAvertissement(frm As Object, message As String)
    On Error Resume Next
    
    ' Nettoyer les anciens messages
    Call NettoyerMessagesTemporaires(frm)
    
    Dim lblMessage As Object
    Set lblMessage = frm.Controls("lblMessageFeedback")
    If lblMessage Is Nothing Then
        Set lblMessage = frm.Controls.Add("Forms.Label.1", "lblMessageFeedback")
        lblMessage.Top = frm.Height - 60
        lblMessage.Left = frm.Width - (FORM_PADDING * 2) - 350
        lblMessage.Width = 350
        lblMessage.Height = 25
        lblMessage.TextAlign = fmTextAlignCenter
    End If
    
    With lblMessage
        .Caption = "⚠ " & message
        .ForeColor = COLOR_WHITE
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = True
        .BackColor = COLOR_WARNING
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_WARNING
        .Visible = True
    End With
    
    ' Animation d'entrée
    Call AnimerGlissementLateral(lblMessage, lblMessage.Left, "gauche")
    
    ' Auto-masquer après 5 secondes
    Application.OnTime Now + TimeValue("00:00:05"), "MasquerMessageFeedback['" & frm.Name & "']"
End Sub

Sub AfficherMessageInformation(frm As Object, message As String)
    On Error Resume Next
    
    ' Nettoyer les anciens messages
    Call NettoyerMessagesTemporaires(frm)
    
    Dim lblMessage As Object
    Set lblMessage = frm.Controls("lblMessageFeedback")
    If lblMessage Is Nothing Then
        Set lblMessage = frm.Controls.Add("Forms.Label.1", "lblMessageFeedback")
        lblMessage.Top = frm.Height - 60
        lblMessage.Left = frm.Width - (FORM_PADDING * 2) - 280
        lblMessage.Width = 280
        lblMessage.Height = 25
        lblMessage.TextAlign = fmTextAlignCenter
    End If
    
    With lblMessage
        .Caption = "ℹ " & message
        .ForeColor = COLOR_WHITE
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .BackColor = COLOR_PRIMARY
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_PRIMARY
        .Visible = True
    End With
    
    ' Animation d'entrée
    Call AnimerGlissementLateral(lblMessage, lblMessage.Left, "gauche")
    
    ' Auto-masquer après 3 secondes
    Application.OnTime Now + TimeValue("00:00:03"), "MasquerMessageFeedback['" & frm.Name & "']"
End Sub

' Nettoyer les messages temporaires
Sub NettoyerMessagesTemporaires(frm As Object)
    On Error Resume Next
    
    Dim ctrl As Object
    For Each ctrl In frm.Controls
        If InStr(ctrl.Name, "lblMessageFeedback") > 0 Then
            frm.Controls.Remove ctrl.Name
            Exit For
        End If
    Next ctrl
End Sub

' Masquer le message de feedback
Sub MasquerMessageFeedback(nomForm As String)
    On Error Resume Next
    
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = nomForm Then
            Call NettoyerMessagesTemporaires(frm)
            Exit For
        End If
    Next frm
End Sub

' ========================================
' CRÉER UNE BARRE DE PROGRESSION
' ========================================
Sub CreerBarreProgression(frm As Object, valeur As Integer, maximum As Integer)
    Dim barreContainer As Object
    Dim barreFond As Object
    Dim barreRemplie As Object
    Dim lblPourcentage As Object
    
    On Error Resume Next
    
    ' Container principal
    Set barreContainer = frm.Controls("barreProgressionContainer")
    If barreContainer Is Nothing Then
        Set barreContainer = frm.Controls.Add("Forms.Frame.1", "barreProgressionContainer")
        barreContainer.Top = 100
        barreContainer.Left = FORM_PADDING
        barreContainer.Width = 200
        barreContainer.Height = 30
        barreContainer.BackColor = COLOR_LIGHT_GRAY
        barreContainer.BorderStyle = fmBorderStyleSingle
        barreContainer.BorderColor = COLOR_MEDIUM_GRAY
    End If
    
    ' Barre de progression
    Set barreRemplie = frm.Controls("barreProgressionRemplie")
    If barreRemplie Is Nothing Then
        Set barreRemplie = frm.Controls.Add("Forms.Label.1", "barreProgressionRemplie")
        barreRemplie.Top = barreContainer.Top + 2
        barreRemplie.Left = barreContainer.Left + 2
        barreRemplie.Height = barreContainer.Height - 4
        barreRemplie.BackColor = COLOR_SUCCESS
    End If
    
    ' Calculer et appliquer la largeur
    Dim pourcentage As Double
    pourcentage = (valeur / maximum) * 100
    barreRemplie.Width = ((barreContainer.Width - 4) * valeur) / maximum
    
    ' Label pourcentage
    Set lblPourcentage = frm.Controls("lblPourcentageProgression")
    If lblPourcentage Is Nothing Then
        Set lblPourcentage = frm.Controls.Add("Forms.Label.1", "lblPourcentageProgression")
        lblPourcentage.Top = barreContainer.Top + 5
        lblPourcentage.Left = barreContainer.Left + barreContainer.Width + 10
        lblPourcentage.Width = 50
        lblPourcentage.Height = 20
    End If
    
    lblPourcentage.Caption = Format(pourcentage, "0") & "%"
    Call StylerLabelNormal(lblPourcentage)
End Sub

' ========================================
' CRÉER UNE JAUGE CIRCULAIRE (APPROXIMATION)
' ========================================
Sub CreerJaugeCirculaire(frm As Object, valeur As Integer, maximum As Integer, titre As String)
    Dim jaugeContainer As Object
    Dim lblTitre As Object
    Dim lblValeur As Object
    
    On Error Resume Next
    
    ' Container de la jauge
    Set jaugeContainer = frm.Controls("jaugeContainer")
    If jaugeContainer Is Nothing Then
        Set jaugeContainer = frm.Controls.Add("Forms.Frame.1", "jaugeContainer")
        jaugeContainer.Top = 50
        jaugeContainer.Left = FORM_PADDING
        jaugeContainer.Width = 120
        jaugeContainer.Height = 120
        jaugeContainer.BackColor = COLOR_LIGHT_GRAY
        jaugeContainer.BorderStyle = fmBorderStyleSingle
        jaugeContainer.BorderColor = COLOR_ACCENT
        jaugeContainer.SpecialEffect = fmSpecialEffectFlat
    End If
    
    ' Titre de la jauge
    Set lblTitre = frm.Controls("lblTitreJauge")
    If lblTitre Is Nothing Then
        Set lblTitre = frm.Controls.Add("Forms.Label.1", "lblTitreJauge")
        lblTitre.Top = jaugeContainer.Top + 10
        lblTitre.Left = jaugeContainer.Left + 10
        lblTitre.Width = jaugeContainer.Width - 20
        lblTitre.Height = 20
        lblTitre.TextAlign = fmTextAlignCenter
    End If
    
    lblTitre.Caption = titre
    Call StylerLabelSecondaire(lblTitre)
    
    ' Valeur centrale
    Set lblValeur = frm.Controls("lblValeurJauge")
    If lblValeur Is Nothing Then
        Set lblValeur = frm.Controls.Add("Forms.Label.1", "lblValeurJauge")
        lblValeur.Top = jaugeContainer.Top + 40
        lblValeur.Left = jaugeContainer.Left + 10
        lblValeur.Width = jaugeContainer.Width - 20
        lblValeur.Height = 40
        lblValeur.TextAlign = fmTextAlignCenter
    End If
    
    Dim pourcentage As Double
    pourcentage = (valeur / maximum) * 100
    lblValeur.Caption = Format(pourcentage, "0.0") & "%"
    
    With lblValeur
        .Font.Name = FONT_PRIMARY
        .Font.Size = 18
        .Font.Bold = True
        .ForeColor = COLOR_PRIMARY
        .BackStyle = fmBackStyleTransparent
    End With

' Animation d'entrée progressive pour les UserForms
Sub AnimerApparitionFormulaire(frm As Object)
    On Error Resume Next
    
    Dim i As Integer
    Dim opacite As Double
    
    ' Commencer transparent
    frm.BackColor = COLOR_WHITE
    
    ' Animation de fondu
    For i = 0 To 10
        opacite = i / 10
        ' Note: En VBA pur, on simule l'opacité avec des effets visuels
        DoEvents
        Application.Wait Now + TimeValue("0:00:00.05")
    Next i
End Sub

' Créer un indicateur de progression circulaire
Sub CreerIndicateurCirculaire(frm As Object, valeur As Integer, maximum As Integer, posX As Integer, posY As Integer)
    On Error Resume Next
    
    Dim container As Object
    Dim lblValeur As Object
    Dim lblTitre As Object
    
    ' Container principal
    Set container = frm.Controls.Add("Forms.Frame.1", "containerCirculaire")
    With container
        .Top = posY
        .Left = posX
        .Width = 80
        .Height = 80
        .BackColor = COLOR_LIGHT_GRAY
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_ACCENT
    End With
    
    ' Titre
    Set lblTitre = frm.Controls.Add("Forms.Label.1", "lblTitreCirculaire")
    With lblTitre
        .Top = posY + 5
        .Left = posX + 5
        .Width = 70
        .Height = 15
        .Caption = "Taux"
        .TextAlign = fmTextAlignCenter
    End With
    Call StylerLabelSecondaire(lblTitre)
    
    ' Valeur centrale
    Set lblValeur = frm.Controls.Add("Forms.Label.1", "lblValeurCirculaire")
    With lblValeur
        .Top = posY + 25
        .Left = posX + 5
        .Width = 70
        .Height = 30
        .TextAlign = fmTextAlignCenter
        .Font.Size = 14
        .Font.Bold = True
    End With
    
    Dim pourcentage As Double
    pourcentage = (valeur / maximum) * 100
    lblValeur.Caption = Format(pourcentage, "0") & "%"
    
    If pourcentage >= 75 Then
        lblValeur.ForeColor = COLOR_SUCCESS
    ElseIf pourcentage >= 50 Then
        lblValeur.ForeColor = COLOR_WARNING
    Else
        lblValeur.ForeColor = COLOR_DANGER
    End If
End Sub

' Créer un système de notifications toast amélioré
Sub AfficherNotificationAvancee(frm As Object, message As String, typeNotif As String, duree As Integer)
    On Error Resume Next
    
    Dim toast As Object
    Dim lblMessage As Object
    Dim lblIcone As Object
    
    Set toast = frm.Controls.Add("Forms.Frame.1", "toast_" & Timer)
    
    With toast
        .Top = 20
        .Left = frm.Width - 320
        .Width = 300
        .Height = 50
        .BorderStyle = fmBorderStyleSingle
    End With
    
    ' Style selon le type
    Select Case typeNotif
        Case "success"
            toast.BackColor = COLOR_SUCCESS
            toast.BorderColor = COLOR_SUCCESS
        Case "error"
            toast.BackColor = COLOR_DANGER
            toast.BorderColor = COLOR_DANGER
        Case "warning"
            toast.BackColor = COLOR_WARNING
            toast.BorderColor = COLOR_WARNING
        Case "info"
            toast.BackColor = COLOR_PRIMARY
            toast.BorderColor = COLOR_PRIMARY
    End Select
    
    ' Icône
    Set lblIcone = toast.Controls.Add("Forms.Label.1", "iconeToast")
    With lblIcone
        .Top = 12
        .Left = 10
        .Width = 25
        .Height = 25
        .TextAlign = fmTextAlignCenter
        .Font.Size = 12
        .ForeColor = COLOR_WHITE
    End With
    
    Select Case typeNotif
        Case "success": lblIcone.Caption = "✓"
        Case "error": lblIcone.Caption = "✗"
        Case "warning": lblIcone.Caption = "⚠"
        Case "info": lblIcone.Caption = "ℹ"
    End Select
    
    ' Message
    Set lblMessage = toast.Controls.Add("Forms.Label.1", "msgToast")
    With lblMessage
        .Top = 12
        .Left = 40
        .Width = 250
        .Height = 25
        .Caption = message
        .ForeColor = COLOR_WHITE
        .Font.Name = FONT_PRIMARY
        .Font.Size = 9
    End With
    
    ' Animation d'entrée
    toast.Left = frm.Width
    Call AnimerGlissementLateral(toast, frm.Width - 320, "gauche")
    
    ' Auto-suppression
    Application.OnTime Now + TimeValue("0:00:" & duree), "SupprimerToastParNom['" & toast.Name & "']"
End Sub

' Fonction pour supprimer un toast spécifique
Sub SupprimerToastParNom(nomToast As String)
    On Error Resume Next
    
    Dim frm As Object
    Dim ctrl As Object
    
    ' Trouver la forme active
    For Each frm In VBA.UserForms
        If frm.Controls.Exists(nomToast) Then
            frm.Controls.Remove nomToast
            Exit For
        End If
    Next frm
End Sub

' ========================================
' VALIDATION AMÉLIORÉE ET SÉCURITÉ
' ========================================

' Validation complète d'un formulaire
Function ValiderFormulaireComplet(frm As Object, regles As Collection) As Boolean
    On Error Resume Next
    
    ValiderFormulaireComplet = True
    
    Dim ctrl As Object
    Dim regle As Variant
    
    For Each ctrl In frm.Controls
        If regles.Exists(ctrl.Name) Then
            regle = regles(ctrl.Name)
            
            Select Case regle("type")
                Case "obligatoire"
                    If Trim(ctrl.Value) = "" Then
                        Call AfficherMessageErreur(frm, "Le champ '" & regle("label") & "' est obligatoire")
                        Call AppliquerFocusVisuel(ctrl, True)
                        ValiderFormulaireComplet = False
                        Exit Function
                    End If
                Case "email"
                    If Not ValiderFormatEmailAvance(ctrl.Value) Then
                        Call AfficherMessageErreur(frm, "Format d'email invalide")
                        Call AppliquerFocusVisuel(ctrl, True)
                        ValiderFormulaireComplet = False
                        Exit Function
                    End If
                Case "telephone"
                    If Not ValiderFormatTelephoneAvance(ctrl.Value) Then
                        Call AfficherMessageErreur(frm, "Format de téléphone invalide")
                        Call AppliquerFocusVisuel(ctrl, True)
                        ValiderFormulaireComplet = False
                        Exit Function
                    End If
                Case "date"
                    If Not IsDate(ctrl.Value) Then
                        Call AfficherMessageErreur(frm, "Format de date invalide")
                        Call AppliquerFocusVisuel(ctrl, True)
                        ValiderFormulaireComplet = False
                        Exit Function
                    End If
            End Select
        End If
    Next ctrl
End Function

' Validation d'email avancée
Function ValiderFormatEmailAvance(email As String) As Boolean
    Dim regex As Object
    
    On Error Resume Next
    
    ' Expression régulière simple pour la validation d'email
    ValiderFormatEmailAvance = email Like "*@*.*" And Len(email) >= 5
    
    If ValiderFormatEmailAvance Then
        ' Vérifications supplémentaires
        Dim parties As Variant
        parties = Split(email, "@")
        
        If UBound(parties) = 1 Then
            ValiderFormatEmailAvance = Len(parties(0)) > 0 And Len(parties(1)) > 0
        End If
    End If
End Function

' Validation de téléphone avancée
Function ValiderFormatTelephoneAvance(tel As String) As Boolean
    On Error Resume Next
    
    ' Supprimer tous les caractères non numériques
    Dim chiffresSeulement As String
    Dim i As Integer
    
    For i = 1 To Len(tel)
        If IsNumeric(Mid(tel, i, 1)) Then
            chiffresSeulement = chiffresSeulement & Mid(tel, i, 1)
        End If
    Next i
    
    ValiderFormatTelephoneAvance = Len(chiffresSeulement) = 10
End Function
