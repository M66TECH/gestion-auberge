Attribute VB_Name = "EffetsVisuels"
' ========================================
' MODULE EFFETS VISUELS AVANCÉS
' ========================================
' Description: Animations et effets pour une UX moderne

Option Explicit

' ========================================
' ANIMATIONS DE TRANSITION
' ========================================

' Animation de fondu d'entrée optimisée
Sub AnimerFonduEntree(ctrl As Object, Optional duree As Double = 0.5)
    On Error Resume Next
    
    Dim i As Integer
    Dim pas As Integer
    Dim delai As Double
    
    pas = 15 ' Réduit pour de meilleures performances
    delai = duree / pas
    
    ' Commencer invisible
    ctrl.Visible = False
    
    For i = 0 To pas
        ctrl.Visible = True
        ' Simuler l'opacité avec la couleur de fond
        If TypeName(ctrl) = "Label" Or TypeName(ctrl) = "CommandButton" Then
            Dim couleurOriginale As Long
            couleurOriginale = ctrl.BackColor
            
            ' Effet de fondu progressif
            Dim facteurOpacite As Double
            facteurOpacite = i / pas
            ctrl.BackColor = MelangerCouleurs(COLOR_WHITE, couleurOriginale, facteurOpacite)
        End If
        
        DoEvents
        Sleep 10 ' Plus rapide
    Next i
End Sub

' Animation de glissement latéral optimisée
Sub AnimerGlissementLateral(ctrl As Object, positionFinale As Integer, Optional direction As String = "droite")
    On Error Resume Next
    
    Dim positionInitiale As Integer
    Dim i As Integer
    Dim pas As Integer
    Dim increment As Integer
    
    pas = 25 ' Plus de fluidité
    positionInitiale = ctrl.Left
    
    If direction = "droite" Then
        ctrl.Left = positionInitiale - 200
        increment = (positionFinale - (positionInitiale - 200)) / pas
    Else
        ctrl.Left = positionInitiale + 200
        increment = (positionFinale - (positionInitiale + 200)) / pas
    End If
    
    For i = 1 To pas
        ctrl.Left = ctrl.Left + increment
        DoEvents
        Sleep 8 ' Plus rapide
    Next i
    
    ctrl.Left = positionFinale
End Sub

' Animation de pulsation optimisée
Sub AnimerPulsation(ctrl As Object, Optional cycles As Integer = 3)
    On Error Resume Next
    
    Dim couleurOriginale As Long
    Dim tailleOriginale As Integer
    Dim i As Integer, j As Integer
    
    If TypeName(ctrl) = "CommandButton" Or TypeName(ctrl) = "Label" Then
        couleurOriginale = ctrl.BackColor
        
        For i = 1 To cycles
            ' Pulsation couleur optimisée
            For j = 0 To 8 Step 2 ' Moins de pas pour plus de fluidité
                ctrl.BackColor = MelangerCouleurs(couleurOriginale, COLOR_ACCENT, j / 10)
                DoEvents
                Sleep 30
            Next j
            
            For j = 8 To 0 Step -2
                ctrl.BackColor = MelangerCouleurs(couleurOriginale, COLOR_ACCENT, j / 10)
                DoEvents
                Sleep 30
            Next j
        Next i
        
        ctrl.BackColor = couleurOriginale
    End If
End Sub

' ========================================
' EFFETS DE SURVOL AVANCÉS
' ========================================

' ========================================
' EFFETS DE SURVOL OPTIMISÉS
' ========================================

' Effet de survol avec ombre portée (simulation optimisée)
Sub EffetSurvolOmbre(ctrl As Object, activer As Boolean)
    On Error Resume Next
    
    Static couleursOriginales As Object
    
    If couleursOriginales Is Nothing Then
        Set couleursOriginales = CreateObject("Scripting.Dictionary")
    End If
    
    If activer Then
        ' Stocker la couleur originale si pas déjà fait
        If Not couleursOriginales.Exists(ctrl.Name) Then
            couleursOriginales.Add ctrl.Name, ctrl.BackColor
        End If
        
        ' Simuler une ombre en modifiant la couleur de fond
        Select Case TypeName(ctrl)
            Case "CommandButton"
                ctrl.BackColor = AssombrirCouleur(ctrl.BackColor, 0.15)
                ctrl.Font.Bold = True
            Case "Label"
                ctrl.ForeColor = COLOR_ACCENT
                ctrl.Font.Bold = True
        End Select
    Else
        ' Restaurer l'apparence normale
        If couleursOriginales.Exists(ctrl.Name) Then
            Select Case TypeName(ctrl)
                Case "CommandButton"
                    ctrl.BackColor = couleursOriginales(ctrl.Name)
                    Call StylerBoutonPrimaire(ctrl)
                Case "Label"
                    ctrl.ForeColor = couleursOriginales(ctrl.Name)
                    Call StylerLabelNormal(ctrl)
            End Select
        End If
    End If
End Sub

' Effet de zoom au survol optimisé
Sub EffetSurvolZoom(ctrl As Object, activer As Boolean)
    On Error Resume Next
    
    Static taillesOriginales As Object
    
    If taillesOriginales Is Nothing Then
        Set taillesOriginales = CreateObject("Scripting.Dictionary")
    End If
    
    If activer Then
        ' Stocker les dimensions originales
        If Not taillesOriginales.Exists(ctrl.Name) Then
            taillesOriginales.Add ctrl.Name, Array(ctrl.Height, ctrl.Width)
        End If
        
        Dim dimensions As Variant
        dimensions = taillesOriginales(ctrl.Name)
        
        ' Agrandir légèrement
        ctrl.Height = dimensions(0) * 1.08
        ctrl.Width = dimensions(1) * 1.08
    Else
        ' Restaurer la taille
        If taillesOriginales.Exists(ctrl.Name) Then
            Dim dimensionsOrig As Variant
            dimensionsOrig = taillesOriginales(ctrl.Name)
            ctrl.Height = dimensionsOrig(0)
            ctrl.Width = dimensionsOrig(1)
        End If
    End If
End Sub

' ========================================
' INDICATEURS VISUELS DYNAMIQUES
' ========================================

' Créer un indicateur de chargement rotatif (simulation)
Sub CreerIndicateurChargement(parent As Object, posX As Integer, posY As Integer)
    On Error Resume Next
    
    Dim lblChargement As Object
    Set lblChargement = parent.Controls.Add("Forms.Label.1", "lblChargement")
    
    With lblChargement
        .Top = posY
        .Left = posX
        .Width = 30
        .Height = 30
        .Caption = "⟳"
        .Font.Size = 16
        .TextAlign = fmTextAlignCenter
        .ForeColor = COLOR_ACCENT
    End With
    
    ' Animation de rotation (simulation avec différents caractères)
    Call AnimerRotation(lblChargement)
End Sub

' Animation de rotation simulée optimisée
Sub AnimerRotation(ctrl As Object)
    Dim symboles As Variant
    Dim i As Integer, j As Integer
    
    symboles = Array("⟳", "⟲", "⟳", "⟲")
    
    For j = 1 To 8 ' Moins de cycles pour de meilleures performances
        For i = 0 To UBound(symboles)
            ctrl.Caption = symboles(i)
            DoEvents
            Sleep 150
        Next i
    Next j
End Sub

' Barre de progression avec animation fluide
Sub CreerBarreProgressionAnimee(parent As Object, valeur As Integer, maximum As Integer, _
                                posX As Integer, posY As Integer, largeur As Integer)
    On Error Resume Next
    
    ' Container
    Dim container As Object
    Set container = parent.Controls.Add("Forms.Frame.1", "containerProgression")
    
    With container
        .Top = posY
        .Left = posX
        .Width = largeur
        .Height = 25
        .BackColor = COLOR_LIGHT_GRAY
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_MEDIUM_GRAY
    End With
    
    ' Barre de progression
    Dim barre As Object
    Set barre = parent.Controls.Add("Forms.Label.1", "barreProgression")
    
    With barre
        .Top = posY + 2
        .Left = posX + 2
        .Height = 21
        .BackColor = COLOR_SUCCESS
        .BorderStyle = fmBorderStyleNone
    End With
    
    ' Animation progressive
    Call AnimerProgression(barre, largeur - 4, valeur, maximum)
End Sub

' Animation de progression fluide optimisée
Sub AnimerProgression(barre As Object, largeurMax As Integer, valeur As Integer, maximum As Integer)
    Dim largeurFinale As Integer
    Dim i As Integer
    Dim pas As Integer
    
    largeurFinale = (largeurMax * valeur) / maximum
    pas = 40 ' Plus de fluidité
    
    For i = 0 To pas
        barre.Width = (largeurFinale * i) / pas
        DoEvents
        Sleep 15
    Next i
    
    barre.Width = largeurFinale
End Sub

' ========================================
' NOTIFICATIONS TOAST MODERNES
' ========================================

' Créer une notification toast
Sub AfficherNotificationToast(parent As Object, message As String, typeNotif As String)
    On Error Resume Next
    
    Dim toast As Object
    Set toast = parent.Controls.Add("Forms.Frame.1", "toastNotification")
    
    ' Position en haut à droite
    With toast
        .Top = 10
        .Left = parent.Width - 300
        .Width = 280
        .Height = 60
        .Caption = ""
    End With
    
    ' Style selon le type
    Select Case typeNotif
        Case "success"
            toast.BackColor = COLOR_SUCCESS
            Call AjouterIconeToast(toast, "✓", COLOR_WHITE)
        Case "error"
            toast.BackColor = COLOR_DANGER
            Call AjouterIconeToast(toast, "✗", COLOR_WHITE)
        Case "warning"
            toast.BackColor = COLOR_WARNING
            Call AjouterIconeToast(toast, "⚠", COLOR_WHITE)
        Case Else
            toast.BackColor = COLOR_PRIMARY
            Call AjouterIconeToast(toast, "ℹ", COLOR_WHITE)
    End Select
    
    ' Message
    Dim lblMessage As Object
    Set lblMessage = toast.Controls.Add("Forms.Label.1", "lblToastMessage")
    
    With lblMessage
        .Top = 15
        .Left = 40
        .Width = 230
        .Height = 30
        .Caption = message
        .ForeColor = COLOR_WHITE
        .Font.Name = FONT_PRIMARY
        .Font.Size = 9
        .WordWrap = True
    End With
    
    ' Animation d'entrée
    Call AnimerGlissementLateral(toast, toast.Left, "gauche")
    
    ' Auto-suppression après 3 secondes
    Application.OnTime Now + TimeValue("0:00:03"), "SupprimerToast"
End Sub

' Ajouter une icône au toast
Sub AjouterIconeToast(toast As Object, icone As String, couleur As Long)
    Dim lblIcone As Object
    Set lblIcone = toast.Controls.Add("Forms.Label.1", "lblIconeToast")
    
    With lblIcone
        .Top = 15
        .Left = 10
        .Width = 25
        .Height = 25
        .Caption = icone
        .ForeColor = couleur
        .Font.Size = 14
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
    End With
End Sub

' ========================================
' EFFETS DE VALIDATION VISUELLE
' ========================================

' Effet de validation réussie
Sub EffetValidationReussie(ctrl As Object)
    On Error Resume Next
    
    Dim couleurOriginale As Long
    couleurOriginale = ctrl.BackColor
    
    ' Flash vert
    ctrl.BackColor = COLOR_SUCCESS
    DoEvents
    Application.Wait Now + TimeValue("0:00:00.2")
    
    ctrl.BackColor = couleurOriginale
    
    ' Ajouter une coche temporaire si c'est un TextBox
    If TypeName(ctrl) = "TextBox" Then
        Dim ancienTexte As String
        ancienTexte = ctrl.Text
        ctrl.Text = ctrl.Text & " ✓"
        DoEvents
        Application.Wait Now + TimeValue("0:00:01")
        ctrl.Text = ancienTexte
    End If
End Sub

' Effet d'erreur de validation
Sub EffetValidationEchec(ctrl As Object)
    On Error Resume Next
    
    Dim couleurOriginale As Long
    Dim positionOriginale As Integer
    
    couleurOriginale = ctrl.BackColor
    positionOriginale = ctrl.Left
    
    ' Flash rouge
    ctrl.BackColor = COLOR_DANGER
    
    ' Effet de tremblement
    Dim i As Integer
    For i = 1 To 5
        ctrl.Left = positionOriginale - 2
        DoEvents
        Application.Wait Now + TimeValue("0:00:00.05")
        ctrl.Left = positionOriginale + 2
        DoEvents
        Application.Wait Now + TimeValue("0:00:00.05")
    Next i
    
    ctrl.Left = positionOriginale
    ctrl.BackColor = couleurOriginale
End Sub

' ========================================
' FONCTIONS UTILITAIRES COULEURS
' ========================================

' Mélanger deux couleurs
Function MelangerCouleurs(couleur1 As Long, couleur2 As Long, ratio As Double) As Long
    Dim r1 As Integer, g1 As Integer, b1 As Integer
    Dim r2 As Integer, g2 As Integer, b2 As Integer
    Dim rFinal As Integer, gFinal As Integer, bFinal As Integer
    
    ' Extraire les composantes RGB de la première couleur
    r1 = couleur1 And &HFF
    g1 = (couleur1 And &HFF00) \ &H100
    b1 = (couleur1 And &HFF0000) \ &H10000
    
    ' Extraire les composantes RGB de la deuxième couleur
    r2 = couleur2 And &HFF
    g2 = (couleur2 And &HFF00) \ &H100
    b2 = (couleur2 And &HFF0000) \ &H10000
    
    ' Calculer les composantes finales
    rFinal = r1 + (r2 - r1) * ratio
    gFinal = g1 + (g2 - g1) * ratio
    bFinal = b1 + (b2 - b1) * ratio
    
    ' Recombiner en couleur
    MelangerCouleurs = RGB(rFinal, gFinal, bFinal)
End Function

' Assombrir une couleur
Function AssombrirCouleur(couleur As Long, facteur As Double) As Long
    AssombrirCouleur = MelangerCouleurs(couleur, &H0, facteur)
End Function

' Éclaircir une couleur
Function EclaircirCouleur(couleur As Long, facteur As Double) As Long
    EclaircirCouleur = MelangerCouleurs(couleur, &HFFFFFF, facteur)
End Function

' ========================================
' NETTOYAGE ET SUPPRESSION
' ========================================

' Supprimer les éléments temporaires
Sub SupprimerToast()
    On Error Resume Next
    ' Cette fonction sera appelée par le timer pour supprimer les toasts
    ' En pratique, il faudrait maintenir une référence aux toasts actifs
End Sub

' Nettoyer tous les effets visuels
Sub NettoyerEffetsVisuels(parent As Object)
    On Error Resume Next
    
    ' Supprimer les contrôles temporaires
    Dim ctrl As Object
    For Each ctrl In parent.Controls
        If InStr(ctrl.Name, "toast") > 0 Or _
           InStr(ctrl.Name, "Chargement") > 0 Or _
           InStr(ctrl.Name, "Progression") > 0 Then
            parent.Controls.Remove ctrl.Name
        End If
    Next ctrl
End Sub
