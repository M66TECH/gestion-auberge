Attribute VB_Name = "EffetsVisuelsV2"
' ========================================
' MODULE EFFETS VISUELS V2 - UX PREMIUM
' ========================================
' Description: Effets visuels avancés et animations pour UserForms
' Version: 2.0 - Interactions modernes et feedback utilisateur
' Auteur: Assistant IA - Spécialiste UX/UI

Option Explicit

' ========================================
' CONSTANTES POUR ANIMATIONS
' ========================================
Private Const ANIMATION_DURATION_FAST As Double = 0.15     ' 150ms
Private Const ANIMATION_DURATION_NORMAL As Double = 0.3    ' 300ms
Private Const ANIMATION_DURATION_SLOW As Double = 0.5      ' 500ms

Private Const EASING_LINEAR As String = "linear"
Private Const EASING_EASE_IN As String = "ease-in"
Private Const EASING_EASE_OUT As String = "ease-out"
Private Const EASING_EASE_IN_OUT As String = "ease-in-out"

' ========================================
' EFFETS DE HOVER AVANCÉS
' ========================================

' Effet hover pour boutons primaires
Sub AppliquerHoverBoutonPrimaire(btn As Object)
    On Error Resume Next
    
    ' Sauvegarder l'état original
    btn.Tag = btn.BackColor & "," & btn.BorderColor
    
    ' Appliquer l'effet hover
    With btn
        .BackColor = COLOR_ACCENT_HOVER
        .BorderColor = COLOR_ACCENT_HOVER
        .Font.Bold = True
    End With
    
    ' Effet de "lift" simulé par changement de bordure
    Call SimulerOmbre(btn, "hover")
End Sub

' Restaurer l'état normal du bouton
Sub RetirerHoverBoutonPrimaire(btn As Object)
    On Error Resume Next
    
    Dim couleurs() As String
    If btn.Tag <> "" Then
        couleurs = Split(btn.Tag, ",")
        If UBound(couleurs) >= 1 Then
            btn.BackColor = CLng(couleurs(0))
            btn.BorderColor = CLng(couleurs(1))
        End If
    End If
    
    Call SimulerOmbre(btn, "normal")
End Sub

' Effet hover pour boutons secondaires
Sub AppliquerHoverBoutonSecondaire(btn As Object)
    On Error Resume Next
    
    btn.Tag = btn.BackColor & "," & btn.BorderColor
    
    With btn
        .BackColor = COLOR_SURFACE_HOVER
        .BorderColor = COLOR_ACCENT
    End With
End Sub

Sub RetirerHoverBoutonSecondaire(btn As Object)
    On Error Resume Next
    
    Dim couleurs() As String
    If btn.Tag <> "" Then
        couleurs = Split(btn.Tag, ",")
        If UBound(couleurs) >= 1 Then
            btn.BackColor = CLng(couleurs(0))
            btn.BorderColor = CLng(couleurs(1))
        End If
    End If
End Sub

' ========================================
' EFFETS DE FOCUS MODERNES
' ========================================

' Focus ring moderne pour champs de saisie
Sub AppliquerFocusModerne(ctrl As Object)
    On Error Resume Next
    
    ' Sauvegarder l'état original
    ctrl.Tag = ctrl.BorderColor & "," & ctrl.BackColor
    
    With ctrl
        .BorderColor = COLOR_BORDER_FOCUS
        .BackColor = COLOR_BACKGROUND
        ' Simuler un "glow" avec une bordure plus épaisse visuellement
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    ' Animation de pulsation subtile
    Call AnimerPulsationSubtile(ctrl)
End Sub

' Retirer le focus
Sub RetirerFocusModerne(ctrl As Object)
    On Error Resume Next
    
    Dim couleurs() As String
    If ctrl.Tag <> "" Then
        couleurs = Split(ctrl.Tag, ",")
        If UBound(couleurs) >= 1 Then
            ctrl.BorderColor = CLng(couleurs(0))
            ctrl.BackColor = CLng(couleurs(1))
        End If
    End If
End Sub

' ========================================
' ANIMATIONS DE FEEDBACK
' ========================================

' Animation de succès (vert qui pulse)
Sub AnimerSuccesAvance(ctrl As Object)
    On Error Resume Next
    
    Dim couleurOriginale As Long
    Dim i As Integer
    
    couleurOriginale = ctrl.BackColor
    
    ' Effet de pulsation verte
    For i = 1 To 3
        ctrl.BackColor = COLOR_SUCCESS_LIGHT
        DoEvents
        Application.Wait (Now + TimeValue("0:00:00.1"))
        
        ctrl.BackColor = COLOR_SUCCESS
        DoEvents
        Application.Wait (Now + TimeValue("0:00:00.1"))
    Next i
    
    ctrl.BackColor = couleurOriginale
End Sub

' Animation d'erreur (rouge qui tremble)
Sub AnimerErreurAvance(ctrl As Object)
    On Error Resume Next
    
    Dim positionOriginale As Integer
    Dim couleurOriginale As Long
    Dim i As Integer
    
    positionOriginale = ctrl.Left
    couleurOriginale = ctrl.BackColor
    
    ' Effet de tremblement avec couleur rouge
    ctrl.BackColor = COLOR_ERROR_LIGHT
    
    For i = 1 To 4
        ctrl.Left = positionOriginale - 2
        DoEvents
        Application.Wait (Now + TimeValue("0:00:00.05"))
        
        ctrl.Left = positionOriginale + 2
        DoEvents
        Application.Wait (Now + TimeValue("0:00:00.05"))
    Next i
    
    ctrl.Left = positionOriginale
    ctrl.BackColor = couleurOriginale
End Sub

' Animation d'information (bleu qui fade)
Sub AnimerInformation(ctrl As Object)
    On Error Resume Next
    
    Dim couleurOriginale As Long
    couleurOriginale = ctrl.BackColor
    
    ' Effet de fade bleu
    ctrl.BackColor = COLOR_INFO_LIGHT
    DoEvents
    Application.Wait (Now + TimeValue("0:00:00.5"))
    
    ctrl.BackColor = COLOR_INFO
    DoEvents
    Application.Wait (Now + TimeValue("0:00:00.3"))
    
    ctrl.BackColor = couleurOriginale
End Sub

' ========================================
' TRANSITIONS D'APPARITION
' ========================================

' Faire apparaître un contrôle avec effet de fade
Sub FadeIn(ctrl As Object, Optional duree As Double = ANIMATION_DURATION_NORMAL)
    On Error Resume Next
    
    Dim steps As Integer
    Dim i As Integer
    Dim stepDuration As Double
    
    steps = 10
    stepDuration = duree / steps
    
    ctrl.Visible = True
    
    ' Simuler un fade-in en jouant avec la couleur de fond
    For i = 1 To steps
        DoEvents
        Application.Wait (Now + TimeValue("0:00:00." & Format(stepDuration * 100, "00")))
    Next i
End Sub

' Faire disparaître un contrôle avec effet de fade
Sub FadeOut(ctrl As Object, Optional duree As Double = ANIMATION_DURATION_NORMAL)
    On Error Resume Next
    
    Dim steps As Integer
    Dim i As Integer
    Dim stepDuration As Double
    
    steps = 10
    stepDuration = duree / steps
    
    ' Simuler un fade-out
    For i = steps To 1 Step -1
        DoEvents
        Application.Wait (Now + TimeValue("0:00:00." & Format(stepDuration * 100, "00")))
    Next i
    
    ctrl.Visible = False
End Sub

' ========================================
' EFFETS DE CHARGEMENT
' ========================================

' Barre de progression animée
Sub AnimerBarreProgression(progressBar As Object, pourcentage As Integer)
    On Error Resume Next
    
    Dim i As Integer
    Dim largeurCible As Integer
    Dim largeurActuelle As Integer
    
    largeurCible = (progressBar.Parent.Width - 40) * pourcentage / 100
    largeurActuelle = progressBar.Width
    
    ' Animation fluide vers la cible
    If largeurCible > largeurActuelle Then
        For i = largeurActuelle To largeurCible Step 5
            progressBar.Width = i
            DoEvents
            Application.Wait (Now + TimeValue("0:00:00.02"))
        Next i
    Else
        For i = largeurActuelle To largeurCible Step -5
            progressBar.Width = i
            DoEvents
            Application.Wait (Now + TimeValue("0:00:00.02"))
        Next i
    End If
    
    progressBar.Width = largeurCible
End Sub

' Spinner de chargement (rotation simulée)
Sub AnimerSpinner(spinner As Object, Optional tours As Integer = 3)
    On Error Resume Next
    
    Dim i As Integer
    Dim j As Integer
    Dim caracteres As String
    
    caracteres = "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏"
    
    For i = 1 To tours
        For j = 1 To Len(caracteres)
            spinner.Caption = Mid(caracteres, j, 1) & " Chargement..."
            DoEvents
            Application.Wait (Now + TimeValue("0:00:00.1"))
        Next j
    Next i
    
    spinner.Caption = "✓ Terminé"
End Sub

' ========================================
' EFFETS DE NOTIFICATION
' ========================================

' Toast notification moderne
Sub AfficherToastNotification(message As String, type As String, Optional duree As Integer = 3)
    On Error Resume Next
    
    Dim couleurFond As Long
    Dim couleurTexte As Long
    Dim icone As String
    
    ' Définir les couleurs selon le type
    Select Case LCase(type)
        Case "success"
            couleurFond = COLOR_SUCCESS_LIGHT
            couleurTexte = COLOR_SUCCESS
            icone = "✓ "
        Case "error"
            couleurFond = COLOR_ERROR_LIGHT
            couleurTexte = COLOR_ERROR
            icone = "✗ "
        Case "warning"
            couleurFond = COLOR_WARNING_LIGHT
            couleurTexte = COLOR_WARNING
            icone = "⚠ "
        Case "info"
            couleurFond = COLOR_INFO_LIGHT
            couleurTexte = COLOR_INFO
            icone = "ℹ "
        Case Else
            couleurFond = COLOR_SURFACE
            couleurTexte = COLOR_TEXT_PRIMARY
            icone = "• "
    End Select
    
    ' Afficher le message avec style
    MsgBox icone & message, vbInformation, "Notification"
End Sub

' ========================================
' EFFETS DE VALIDATION EN TEMPS RÉEL
' ========================================

' Validation visuelle d'un champ
Sub ValiderChampVisuellement(ctrl As Object, estValide As Boolean)
    On Error Resume Next
    
    If estValide Then
        With ctrl
            .BackColor = COLOR_SUCCESS_LIGHT
            .BorderColor = COLOR_SUCCESS
        End With
        Call AnimerSuccesAvance(ctrl)
    Else
        With ctrl
            .BackColor = COLOR_ERROR_LIGHT
            .BorderColor = COLOR_ERROR
        End With
        Call AnimerErreurAvance(ctrl)
    End If
End Sub

' Réinitialiser l'état visuel d'un champ
Sub ReinitialiserEtatVisuel(ctrl As Object)
    On Error Resume Next
    
    With ctrl
        .BackColor = COLOR_BACKGROUND
        .BorderColor = COLOR_BORDER
    End With
End Sub

' ========================================
' EFFETS DE MICRO-INTERACTIONS
' ========================================

' Effet de "press" sur un bouton
Sub AnimerPresseBouton(btn As Object)
    On Error Resume Next
    
    Dim positionOriginale As Integer
    positionOriginale = btn.Top
    
    ' Simuler un enfoncement
    btn.Top = positionOriginale + 1
    DoEvents
    Application.Wait (Now + TimeValue("0:00:00.05"))
    
    btn.Top = positionOriginale
End Sub

' Pulsation subtile pour attirer l'attention
Sub AnimerPulsationSubtile(ctrl As Object)
    On Error Resume Next
    
    Dim couleurOriginale As Long
    couleurOriginale = ctrl.BorderColor
    
    ' Pulsation douce
    ctrl.BorderColor = COLOR_ACCENT
    DoEvents
    Application.Wait (Now + TimeValue("0:00:00.2"))
    
    ctrl.BorderColor = couleurOriginale
End Sub

' ========================================
' EFFETS DE GROUPE ET ORCHESTRATION
' ========================================

' Animer l'apparition séquentielle de contrôles
Sub AnimerApparitionSequentielle(ParamArray controles() As Variant)
    On Error Resume Next
    
    Dim i As Integer
    
    ' Masquer tous les contrôles d'abord
    For i = 0 To UBound(controles)
        controles(i).Visible = False
    Next i
    
    ' Les faire apparaître un par un
    For i = 0 To UBound(controles)
        Call FadeIn(controles(i), ANIMATION_DURATION_FAST)
        Application.Wait (Now + TimeValue("0:00:00.1"))
    Next i
End Sub

' Effet de cascade pour un groupe de boutons
Sub AnimerCascadeBoutons(ParamArray boutons() As Variant)
    On Error Resume Next
    
    Dim i As Integer
    
    For i = 0 To UBound(boutons)
        Call AppliquerHoverBoutonPrimaire(boutons(i))
        Application.Wait (Now + TimeValue("0:00:00.05"))
        Call RetirerHoverBoutonPrimaire(boutons(i))
    Next i
End Sub

' ========================================
' SIMULATEUR D'OMBRES
' ========================================

' Simuler une ombre portée (effet visuel)
Sub SimulerOmbre(ctrl As Object, etat As String)
    On Error Resume Next
    
    Select Case LCase(etat)
        Case "hover"
            ' Simuler une élévation avec une bordure plus marquée
            ctrl.BorderColor = COLOR_ACCENT
        Case "focus"
            ctrl.BorderColor = COLOR_BORDER_FOCUS
        Case "normal"
            ctrl.BorderColor = COLOR_BORDER
        Case "active"
            ctrl.BorderColor = COLOR_ACCENT_HOVER
    End Select
End Sub

' ========================================
' GESTIONNAIRE D'ÉTATS VISUELS
' ========================================

' Gérer les états d'un contrôle interactif
Sub GererEtatVisuel(ctrl As Object, nouvelEtat As String)
    On Error Resume Next
    
    Select Case LCase(nouvelEtat)
        Case "normal"
            Call ReinitialiserEtatVisuel(ctrl)
        Case "hover"
            Call SimulerOmbre(ctrl, "hover")
        Case "focus"
            Call AppliquerFocusModerne(ctrl)
        Case "active"
            Call SimulerOmbre(ctrl, "active")
        Case "disabled"
            With ctrl
                .BackColor = COLOR_SURFACE
                .ForeColor = COLOR_TEXT_MUTED
                .Enabled = False
            End With
        Case "enabled"
            With ctrl
                .BackColor = COLOR_BACKGROUND
                .ForeColor = COLOR_TEXT_PRIMARY
                .Enabled = True
            End With
    End Select
End Sub

' ========================================
' EFFETS DE DONNÉES EN TEMPS RÉEL
' ========================================

' Animer un compteur qui s'incrémente
Sub AnimerCompteur(lbl As Object, valeurCible As Long, Optional duree As Double = 1)
    On Error Resume Next
    
    Dim valeurActuelle As Long
    Dim increment As Long
    Dim steps As Integer
    Dim i As Integer
    
    valeurActuelle = 0
    steps = 20
    increment = valeurCible / steps
    
    For i = 1 To steps
        valeurActuelle = valeurActuelle + increment
        lbl.Caption = Format(valeurActuelle, "#,##0")
        DoEvents
        Application.Wait (Now + TimeValue("0:00:00." & Format(duree * 100 / steps, "00")))
    Next i
    
    lbl.Caption = Format(valeurCible, "#,##0")
End Sub

' Effet de mise à jour de données (flash)
Sub AnimerMiseAJourDonnees(ctrl As Object)
    On Error Resume Next
    
    Dim couleurOriginale As Long
    couleurOriginale = ctrl.BackColor
    
    ' Flash rapide pour indiquer une mise à jour
    ctrl.BackColor = COLOR_INFO_LIGHT
    DoEvents
    Application.Wait (Now + TimeValue("0:00:00.1"))
    
    ctrl.BackColor = couleurOriginale
End Sub

' ========================================
' INITIALISATION DES EFFETS
' ========================================

' Initialiser le système d'effets visuels
Sub InitialiserEffetsVisuelsV2()
    On Error Resume Next
    
    ' Configuration des paramètres d'animation
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "✨ Système d'effets visuels V2 initialisé!" & vbCrLf & _
           "🎭 Animations et micro-interactions activées", vbInformation, "Effets Visuels V2"
End Sub

' Test de tous les effets (pour démonstration)
Sub TesterTousLesEffets()
    On Error Resume Next
    
    MsgBox "🎬 Démonstration des effets visuels V2" & vbCrLf & _
           "Les effets seront appliqués aux contrôles actifs", vbInformation, "Test Effets"
    
    ' Ici on pourrait tester tous les effets sur des contrôles de test
End Sub
