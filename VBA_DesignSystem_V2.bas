Attribute VB_Name = "DesignSystemV2"
' ========================================
' MODULE DESIGN SYSTEM V2 - UX PREMIUM
' ========================================
' Description: Système de design moderne et professionnel
' Version: 2.0 - Améliorations UX avancées
' Auteur: Assistant IA - Spécialiste UX/UI

Option Explicit

' ========================================
' PALETTE DE COULEURS PREMIUM 2024
' ========================================
' Couleurs principales (Thème Corporate Moderne)
Public Const COLOR_PRIMARY As Long = &H2D3748         ' Bleu-gris foncé moderne
Public Const COLOR_PRIMARY_LIGHT As Long = &H4A5568   ' Variant clair
Public Const COLOR_PRIMARY_DARK As Long = &H1A202C    ' Variant foncé

' Couleurs d'accent (Gradient moderne)
Public Const COLOR_ACCENT As Long = &H3182CE          ' Bleu professionnel
Public Const COLOR_ACCENT_LIGHT As Long = &H63B3ED    ' Bleu clair
Public Const COLOR_ACCENT_HOVER As Long = &H2C5282    ' Bleu hover

' Couleurs de surface
Public Const COLOR_BACKGROUND As Long = &HFFFFFF      ' Blanc pur
Public Const COLOR_SURFACE As Long = &HF7FAFC         ' Gris très clair
Public Const COLOR_SURFACE_HOVER As Long = &HEDF2F7   ' Surface hover
Public Const COLOR_BORDER As Long = &HE2E8F0          ' Bordure subtile
Public Const COLOR_BORDER_FOCUS As Long = &H3182CE    ' Bordure focus

' États et feedback
Public Const COLOR_SUCCESS As Long = &H38A169         ' Vert moderne
Public Const COLOR_SUCCESS_LIGHT As Long = &H68D391   ' Vert clair
Public Const COLOR_WARNING As Long = &HED8936         ' Orange moderne
Public Const COLOR_WARNING_LIGHT As Long = &HF6AD55   ' Orange clair
Public Const COLOR_ERROR As Long = &HE53E3E           ' Rouge moderne
Public Const COLOR_ERROR_LIGHT As Long = &HFC8181     ' Rouge clair
Public Const COLOR_INFO As Long = &H3182CE            ' Bleu info
Public Const COLOR_INFO_LIGHT As Long = &H63B3ED      ' Bleu info clair

' Texte et typographie
Public Const COLOR_TEXT_PRIMARY As Long = &H1A202C    ' Texte principal
Public Const COLOR_TEXT_SECONDARY As Long = &H4A5568  ' Texte secondaire
Public Const COLOR_TEXT_MUTED As Long = &H718096      ' Texte atténué
Public Const COLOR_TEXT_INVERSE As Long = &HFFFFFF    ' Texte inversé

' ========================================
' SYSTÈME TYPOGRAPHIQUE MODERNE
' ========================================
Public Const FONT_FAMILY_PRIMARY As String = "Segoe UI"
Public Const FONT_FAMILY_SECONDARY As String = "Inter"
Public Const FONT_FAMILY_MONOSPACE As String = "JetBrains Mono"

' Échelle typographique harmonieuse
Public Const FONT_SIZE_H1 As Integer = 24             ' Titres principaux
Public Const FONT_SIZE_H2 As Integer = 20             ' Sous-titres
Public Const FONT_SIZE_H3 As Integer = 16             ' Titres de section
Public Const FONT_SIZE_BODY As Integer = 14           ' Corps de texte
Public Const FONT_SIZE_BODY_SM As Integer = 12        ' Texte petit
Public Const FONT_SIZE_CAPTION As Integer = 10        ' Légendes

' ========================================
' SYSTÈME D'ESPACEMENT (Échelle 4px)
' ========================================
Public Const SPACE_1 As Integer = 4                   ' 0.25rem
Public Const SPACE_2 As Integer = 8                   ' 0.5rem
Public Const SPACE_3 As Integer = 12                  ' 0.75rem
Public Const SPACE_4 As Integer = 16                  ' 1rem
Public Const SPACE_5 As Integer = 20                  ' 1.25rem
Public Const SPACE_6 As Integer = 24                  ' 1.5rem
Public Const SPACE_8 As Integer = 32                  ' 2rem
Public Const SPACE_10 As Integer = 40                 ' 2.5rem
Public Const SPACE_12 As Integer = 48                 ' 3rem

' ========================================
' COMPOSANTS STANDARDISÉS
' ========================================
Public Const BUTTON_HEIGHT_SM As Integer = 28         ' Bouton petit
Public Const BUTTON_HEIGHT_MD As Integer = 36         ' Bouton moyen
Public Const BUTTON_HEIGHT_LG As Integer = 44         ' Bouton large

Public Const INPUT_HEIGHT_SM As Integer = 28          ' Champ petit
Public Const INPUT_HEIGHT_MD As Integer = 36          ' Champ moyen
Public Const INPUT_HEIGHT_LG As Integer = 44          ' Champ large

Public Const BORDER_RADIUS_SM As Integer = 4          ' Rayon petit
Public Const BORDER_RADIUS_MD As Integer = 6          ' Rayon moyen
Public Const BORDER_RADIUS_LG As Integer = 8          ' Rayon large
Public Const BORDER_RADIUS_XL As Integer = 12         ' Rayon extra-large

' ========================================
' OMBRES ET ÉLÉVATIONS
' ========================================
Public Const SHADOW_SM As String = "0 1px 2px rgba(0,0,0,0.05)"
Public Const SHADOW_MD As String = "0 4px 6px rgba(0,0,0,0.07)"
Public Const SHADOW_LG As String = "0 10px 15px rgba(0,0,0,0.1)"
Public Const SHADOW_XL As String = "0 20px 25px rgba(0,0,0,0.15)"

' ========================================
' FONCTIONS UTILITAIRES AVANCÉES
' ========================================

' Appliquer le thème moderne à un UserForm
Sub AppliquerThemeModerne(frm As Object)
    On Error Resume Next
    
    With frm
        .BackColor = COLOR_BACKGROUND
        .BorderStyle = fmBorderStyleNone
        .Font.Name = FONT_FAMILY_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .ForeColor = COLOR_TEXT_PRIMARY
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

' ========================================
' BOUTONS MODERNES AVEC VARIANTES
' ========================================

' Bouton primaire (action principale)
Sub StylerBoutonPrimaire(btn As Object, Optional taille As String = "md")
    On Error Resume Next
    
    Dim hauteur As Integer
    Select Case taille
        Case "sm": hauteur = BUTTON_HEIGHT_SM
        Case "lg": hauteur = BUTTON_HEIGHT_LG
        Case Else: hauteur = BUTTON_HEIGHT_MD
    End Select
    
    With btn
        .BackColor = COLOR_ACCENT
        .ForeColor = COLOR_TEXT_INVERSE
        .Font.Name = FONT_FAMILY_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = True
        .Height = hauteur
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_ACCENT
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

' Bouton secondaire (action secondaire)
Sub StylerBoutonSecondaire(btn As Object, Optional taille As String = "md")
    On Error Resume Next
    
    Dim hauteur As Integer
    Select Case taille
        Case "sm": hauteur = BUTTON_HEIGHT_SM
        Case "lg": hauteur = BUTTON_HEIGHT_LG
        Case Else: hauteur = BUTTON_HEIGHT_MD
    End Select
    
    With btn
        .BackColor = COLOR_SURFACE
        .ForeColor = COLOR_TEXT_PRIMARY
        .Font.Name = FONT_FAMILY_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .Height = hauteur
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_BORDER
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

' Bouton de succès (confirmation)
Sub StylerBoutonSucces(btn As Object, Optional taille As String = "md")
    On Error Resume Next
    
    Dim hauteur As Integer
    Select Case taille
        Case "sm": hauteur = BUTTON_HEIGHT_SM
        Case "lg": hauteur = BUTTON_HEIGHT_LG
        Case Else: hauteur = BUTTON_HEIGHT_MD
    End Select
    
    With btn
        .BackColor = COLOR_SUCCESS
        .ForeColor = COLOR_TEXT_INVERSE
        .Font.Name = FONT_FAMILY_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = True
        .Height = hauteur
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_SUCCESS
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

' Bouton de danger (suppression/annulation)
Sub StylerBoutonDanger(btn As Object, Optional taille As String = "md")
    On Error Resume Next
    
    Dim hauteur As Integer
    Select Case taille
        Case "sm": hauteur = BUTTON_HEIGHT_SM
        Case "lg": hauteur = BUTTON_HEIGHT_LG
        Case Else: hauteur = BUTTON_HEIGHT_MD
    End Select
    
    With btn
        .BackColor = COLOR_ERROR
        .ForeColor = COLOR_TEXT_INVERSE
        .Font.Name = FONT_FAMILY_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = True
        .Height = hauteur
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_ERROR
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

' ========================================
' CHAMPS DE SAISIE MODERNES
' ========================================

' Champ de texte standard
Sub StylerChampTexte(txt As Object, Optional taille As String = "md")
    On Error Resume Next
    
    Dim hauteur As Integer
    Select Case taille
        Case "sm": hauteur = INPUT_HEIGHT_SM
        Case "lg": hauteur = INPUT_HEIGHT_LG
        Case Else: hauteur = INPUT_HEIGHT_MD
    End Select
    
    With txt
        .BackColor = COLOR_BACKGROUND
        .ForeColor = COLOR_TEXT_PRIMARY
        .Font.Name = FONT_FAMILY_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .Height = hauteur
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_BORDER
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

' Champ avec état d'erreur
Sub StylerChampErreur(txt As Object)
    On Error Resume Next
    
    With txt
        .BackColor = COLOR_ERROR_LIGHT
        .BorderColor = COLOR_ERROR
        .ForeColor = COLOR_TEXT_PRIMARY
    End With
End Sub

' Champ avec état de succès
Sub StylerChampSucces(txt As Object)
    On Error Resume Next
    
    With txt
        .BackColor = COLOR_SUCCESS_LIGHT
        .BorderColor = COLOR_SUCCESS
        .ForeColor = COLOR_TEXT_PRIMARY
    End With
End Sub

' ========================================
' LABELS ET TITRES
' ========================================

' Titre principal (H1)
Sub StylerTitrePrincipal(lbl As Object)
    On Error Resume Next
    
    With lbl
        .Font.Name = FONT_FAMILY_PRIMARY
        .Font.Size = FONT_SIZE_H1
        .Font.Bold = True
        .ForeColor = COLOR_TEXT_PRIMARY
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

' Sous-titre (H2)
Sub StylerSousTitre(lbl As Object)
    On Error Resume Next
    
    With lbl
        .Font.Name = FONT_FAMILY_PRIMARY
        .Font.Size = FONT_SIZE_H2
        .Font.Bold = True
        .ForeColor = COLOR_TEXT_PRIMARY
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

' Titre de section (H3)
Sub StylerTitreSection(lbl As Object)
    On Error Resume Next
    
    With lbl
        .Font.Name = FONT_FAMILY_PRIMARY
        .Font.Size = FONT_SIZE_H3
        .Font.Bold = True
        .ForeColor = COLOR_TEXT_SECONDARY
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

' Texte de corps
Sub StylerTexteCorps(lbl As Object)
    On Error Resume Next
    
    With lbl
        .Font.Name = FONT_FAMILY_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .ForeColor = COLOR_TEXT_PRIMARY
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

' Texte secondaire
Sub StylerTexteSecondaire(lbl As Object)
    On Error Resume Next
    
    With lbl
        .Font.Name = FONT_FAMILY_PRIMARY
        .Font.Size = FONT_SIZE_BODY_SM
        .ForeColor = COLOR_TEXT_SECONDARY
        .BackStyle = fmBackStyleTransparent
    End Sub
End Sub

' ========================================
' PANNEAUX ET CONTENEURS
' ========================================

' Panneau principal (carte)
Sub StylerPanneauPrincipal(frm As Object)
    On Error Resume Next
    
    With frm
        .BackColor = COLOR_BACKGROUND
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_BORDER
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

' Panneau secondaire
Sub StylerPanneauSecondaire(frm As Object)
    On Error Resume Next
    
    With frm
        .BackColor = COLOR_SURFACE
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_BORDER
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

' ========================================
' LISTES ET TABLEAUX
' ========================================

' Liste moderne
Sub StylerListeModerne(lst As Object)
    On Error Resume Next
    
    With lst
        .BackColor = COLOR_BACKGROUND
        .ForeColor = COLOR_TEXT_PRIMARY
        .Font.Name = FONT_FAMILY_PRIMARY
        .Font.Size = FONT_SIZE_BODY
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COLOR_BORDER
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

' ========================================
' UTILITAIRES DE MISE EN PAGE
' ========================================

' Appliquer un espacement cohérent
Sub AppliquerEspacement(ctrl As Object, top As Integer, left As Integer, Optional marge As Integer = SPACE_4)
    On Error Resume Next
    
    With ctrl
        .top = top + marge
        .left = left + marge
    End With
End Sub

' Aligner des contrôles horizontalement
Sub AlignerHorizontalement(ParamArray controles() As Variant)
    On Error Resume Next
    
    Dim i As Integer
    Dim positionY As Integer
    
    If UBound(controles) >= 0 Then
        positionY = controles(0).top
        
        For i = 1 To UBound(controles)
            controles(i).top = positionY
        Next i
    End If
End Sub

' Distribuer des contrôles avec espacement égal
Sub DistribuerAvecEspacement(espacement As Integer, ParamArray controles() As Variant)
    On Error Resume Next
    
    Dim i As Integer
    Dim positionActuelle As Integer
    
    If UBound(controles) >= 0 Then
        positionActuelle = controles(0).left
        
        For i = 1 To UBound(controles)
            positionActuelle = positionActuelle + controles(i - 1).Width + espacement
            controles(i).left = positionActuelle
        Next i
    End If
End Sub

' ========================================
' ANIMATIONS ET TRANSITIONS (Simulées)
' ========================================

' Effet de focus sur un contrôle
Sub AnimerFocus(ctrl As Object)
    On Error Resume Next
    
    ' Simulation d'animation par changement de couleur
    Dim couleurOriginale As Long
    couleurOriginale = ctrl.BorderColor
    
    ctrl.BorderColor = COLOR_ACCENT
    DoEvents
    Application.Wait (Now + TimeValue("0:00:01"))
    ctrl.BorderColor = couleurOriginale
End Sub

' Effet de validation réussie
Sub AnimerSucces(ctrl As Object)
    On Error Resume Next
    
    Dim couleurOriginale As Long
    couleurOriginale = ctrl.BackColor
    
    ctrl.BackColor = COLOR_SUCCESS_LIGHT
    DoEvents
    Application.Wait (Now + TimeValue("0:00:01"))
    ctrl.BackColor = couleurOriginale
End Sub

' Effet d'erreur
Sub AnimerErreur(ctrl As Object)
    On Error Resume Next
    
    Dim couleurOriginale As Long
    couleurOriginale = ctrl.BackColor
    
    ctrl.BackColor = COLOR_ERROR_LIGHT
    DoEvents
    Application.Wait (Now + TimeValue("0:00:02"))
    ctrl.BackColor = couleurOriginale
End Sub

' ========================================
' INITIALISATION DU SYSTÈME
' ========================================

' Initialiser le design system pour l'application
Sub InitialiserDesignSystemV2()
    On Error Resume Next
    
    ' Configuration globale de l'application
    Application.ScreenUpdating = False
    
    ' Ici on pourrait configurer des paramètres globaux
    ' comme les thèmes, les préférences utilisateur, etc.
    
    Application.ScreenUpdating = True
    
    MsgBox "🎨 Design System V2 initialisé avec succès!" & vbCrLf & _
           "✨ Interface moderne activée", vbInformation, "Design System V2"
End Sub
