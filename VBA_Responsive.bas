Attribute VB_Name = "ResponsiveDesign"
' ========================================
' MODULE RESPONSIVE DESIGN & ADAPTATION
' ========================================
' Description: Adaptation intelligente de l'interface selon la taille d'écran et les préférences utilisateur

Option Explicit

' ========================================
' DÉTECTION DE L'ENVIRONNEMENT
' ========================================
Private resolutionEcran As String
Private tailleEcran As String
Private preferencesUtilisateur As Object

' Initialiser les paramètres responsive
Sub InitialiserResponsive()
    On Error Resume Next

    Call DetecterResolutionEcran
    Call DetecterTailleEcran
    Call ChargerPreferencesUtilisateur
    Call AdapterInterfaceGlobale
End Sub

' Détecter la résolution de l'écran
Sub DetecterResolutionEcran()
    On Error Resume Next

    Dim largeur As Integer
    Dim hauteur As Integer

    ' Obtenir la résolution de l'écran principal
    largeur = Application.Width
    hauteur = Application.Height

    ' Classifier la résolution
    If largeur >= 1920 And hauteur >= 1080 Then
        resolutionEcran = "4K"
    ElseIf largeur >= 1366 And hauteur >= 768 Then
        resolutionEcran = "HD"
    ElseIf largeur >= 1024 And hauteur >= 768 Then
        resolutionEcran = "Tablette"
    Else
        resolutionEcran = "Mobile"
    End If

    Debug.Print "Résolution détectée : " & resolutionEcran & " (" & largeur & "x" & hauteur & ")"
End Sub

' Détecter la taille d'écran logique
Sub DetecterTailleEcran()
    On Error Resume Next

    Dim largeur As Integer
    Dim hauteur As Integer

    largeur = Application.Width
    hauteur = Application.Height

    If largeur >= 1920 Then
        tailleEcran = "TresGrand"
    ElseIf largeur >= 1366 Then
        tailleEcran = "Grand"
    ElseIf largeur >= 1024 Then
        tailleEcran = "Moyen"
    ElseIf largeur >= 768 Then
        tailleEcran = "Petit"
    Else
        tailleEcran = "TresPetit"
    End If
End Sub

' Charger les préférences utilisateur
Sub ChargerPreferencesUtilisateur()
    On Error Resume Next

    If preferencesUtilisateur Is Nothing Then
        Set preferencesUtilisateur = CreateObject("Scripting.Dictionary")
    End If

    ' Préférences par défaut
    preferencesUtilisateur("theme") = "clair"
    preferencesUtilisateur("taillePolice") = "normale"
    preferencesUtilisateur("animations") = True
    preferencesUtilisateur("contrasteEleve") = False
    preferencesUtilisateur("reductionMouvement") = False

    ' En production, charger depuis un fichier de configuration ou base de données
End Sub

' ========================================
' ADAPTATION DE L'INTERFACE
' ========================================

' Adapter l'interface selon la taille d'écran
Sub AdapterInterfaceGlobale()
    On Error Resume Next

    Select Case tailleEcran
        Case "TresGrand"
            Call AdapterInterfaceGrande
        Case "Grand"
            Call AdapterInterfaceGrande
        Case "Moyen"
            Call AdapterInterfaceMoyenne
        Case "Petit"
            Call AdapterInterfaceCompacte
        Case "TresPetit"
            Call AdapterInterfaceMinimale
    End Select
End Sub

' Interface pour grands écrans
Sub AdapterInterfaceGrande()
    On Error Resume Next

    ' Grandes marges et espacement
    FORM_PADDING = 25
    ELEMENT_SPACING = 20
    FONT_SIZE_HEADING = 14
    FONT_SIZE_BODY = 11
    FONT_SIZE_BUTTON = 10

    ' Boutons plus grands
    BUTTON_WIDTH = 140
    BUTTON_HEIGHT = 40

    ' Plus d'informations visibles
    Call ActiverModeDetaille
End Sub

' Interface pour écrans moyens
Sub AdapterInterfaceMoyenne()
    On Error Resume Next

    FORM_PADDING = 20
    ELEMENT_SPACING = 15
    FONT_SIZE_HEADING = 13
    FONT_SIZE_BODY = 10
    FONT_SIZE_BUTTON = 9

    BUTTON_WIDTH = 120
    BUTTON_HEIGHT = 35
End Sub

' Interface compacte
Sub AdapterInterfaceCompacte()
    On Error Resume Next

    FORM_PADDING = 15
    ELEMENT_SPACING = 12
    FONT_SIZE_HEADING = 12
    FONT_SIZE_BODY = 9
    FONT_SIZE_BUTTON = 8

    BUTTON_WIDTH = 100
    BUTTON_HEIGHT = 30

    ' Réduire les animations
    If preferencesUtilisateur("reductionMouvement") Then
        Call DesactiverAnimations
    End If
End Sub

' Interface minimale pour petits écrans
Sub AdapterInterfaceMinimale()
    On Error Resume Next

    FORM_PADDING = 10
    ELEMENT_SPACING = 8
    FONT_SIZE_HEADING = 11
    FONT_SIZE_BODY = 8
    FONT_SIZE_BUTTON = 7

    BUTTON_WIDTH = 80
    BUTTON_HEIGHT = 25

    ' Interface très compacte
    Call ActiverModeMinimal
    Call DesactiverAnimations
End Sub

' Activer le mode détaillé
Sub ActiverModeDetaille()
    On Error Resume Next

    ' Afficher plus d'informations
    ' Activer les tooltips détaillés
    ' Afficher les descriptions complètes
End Sub

' Activer le mode minimal
Sub ActiverModeMinimal()
    On Error Resume Next

    ' Masquer les éléments non essentiels
    ' Réduire les textes
    ' Utiliser des icônes
End Sub

' ========================================
' ADAPTATION DES FORMULAIRES
' ========================================

' Adapter un formulaire selon la résolution
Sub AdapterFormulaire(frm As Object)
    On Error Resume Next

    Dim nouvelleLargeur As Integer
    Dim nouvelleHauteur As Integer

    ' Calculer les nouvelles dimensions
    Select Case tailleEcran
        Case "TresGrand"
            nouvelleLargeur = 900
            nouvelleHauteur = 700
        Case "Grand"
            nouvelleLargeur = 800
            nouvelleHauteur = 600
        Case "Moyen"
            nouvelleLargeur = 700
            nouvelleHauteur = 500
        Case "Petit"
            nouvelleLargeur = 600
            nouvelleHauteur = 400
        Case "TresPetit"
            nouvelleLargeur = 500
            nouvelleHauteur = 350
    End Select

    ' Appliquer les nouvelles dimensions
    If frm.Width > nouvelleLargeur Then frm.Width = nouvelleLargeur
    If frm.Height > nouvelleHauteur Then frm.Height = nouvelleHauteur

    ' Centrer le formulaire
    frm.StartUpPosition = 1 ' Centrer sur l'écran

    ' Adapter les contrôles
    Call AdapterControlesFormulaire(frm)
End Sub

' Adapter les contrôles d'un formulaire
Sub AdapterControlesFormulaire(frm As Object)
    On Error Resume Next

    Dim ctrl As Object
    Dim scaleFactor As Double

    scaleFactor = frm.Width / 800 ' Base 800px

    For Each ctrl In frm.Controls
        Select Case TypeName(ctrl)
            Case "CommandButton"
                ctrl.Width = BUTTON_WIDTH * scaleFactor
                ctrl.Height = BUTTON_HEIGHT * scaleFactor
                ctrl.Font.Size = FONT_SIZE_BUTTON * scaleFactor

            Case "Label"
                If InStr(ctrl.Name, "Titre") > 0 Then
                    ctrl.Font.Size = FONT_SIZE_HEADING * scaleFactor
                Else
                    ctrl.Font.Size = FONT_SIZE_BODY * scaleFactor
                End If

            Case "TextBox", "ComboBox"
                ctrl.Width = ctrl.Width * scaleFactor
                ctrl.Height = ctrl.Height * scaleFactor
                ctrl.Font.Size = FONT_SIZE_BODY * scaleFactor

            Case "Frame"
                ctrl.Width = ctrl.Width * scaleFactor
                ctrl.Height = ctrl.Height * scaleFactor
        End Select
    Next ctrl
End Sub

' ========================================
' ADAPTATION TYPOGRAPHIQUE
' ========================================

' Adapter la taille des polices selon les préférences
Sub AdapterTypographie()
    On Error Resume Next

    Dim facteur As Double

    Select Case preferencesUtilisateur("taillePolice")
        Case "tresPetite"
            facteur = 0.8
        Case "petite"
            facteur = 0.9
        Case "normale"
            facteur = 1.0
        Case "grande"
            facteur = 1.1
        Case "tresGrande"
            facteur = 1.2
    End Select

    FONT_SIZE_HEADING = 13 * facteur
    FONT_SIZE_BODY = 10 * facteur
    FONT_SIZE_BUTTON = 9 * facteur
End Sub

' ========================================
' ADAPTATION DES COULEURS
' ========================================

' Adapter les couleurs selon le thème et les préférences
Sub AdapterSchemaCouleurs()
    On Error Resume Next

    ' Thème sombre/clair
    If preferencesUtilisateur("theme") = "sombre" Then
        Call ActiverThemeSombre
    Else
        Call ActiverThemeClair
    End If

    ' Contraste élevé
    If preferencesUtilisateur("contrasteEleve") Then
        Call ActiverContrasteEleve
    End If
End Sub

' Activer le thème sombre
Sub ActiverThemeSombre()
    On Error Resume Next

    ' Ajuster les couleurs pour le thème sombre
    COLOR_BACKGROUND = &H2D2D30
    COLOR_TEXT_PRIMARY = &HE0E0E0
    COLOR_TEXT_SECONDARY = &HA0A0A0
    COLOR_BORDER = &H404040
End Sub

' Activer le thème clair
Sub ActiverThemeClair()
    On Error Resume Next

    ' Restaurer les couleurs claires
    COLOR_BACKGROUND = &HF8F9FA
    COLOR_TEXT_PRIMARY = &H212529
    COLOR_TEXT_SECONDARY = &H6C757D
    COLOR_BORDER = &HDEE2E6
End Sub

' Activer le contraste élevé
Sub ActiverContrasteEleve()
    On Error Resume Next

    ' Couleurs à fort contraste
    COLOR_TEXT_PRIMARY = &H000000
    COLOR_TEXT_SECONDARY = &H000000
    COLOR_BACKGROUND = &HFFFFFF
    COLOR_PRIMARY = &H0000FF
    COLOR_SUCCESS = &H008000
    COLOR_DANGER = &HFF0000
    COLOR_WARNING = &HFFA500
End Sub

' ========================================
' ADAPTATION DES ANIMATIONS
' ========================================

' Adapter les animations selon les préférences
Sub AdapterAnimations()
    On Error Resume Next

    If preferencesUtilisateur("reductionMouvement") Then
        Call DesactiverAnimations
    ElseIf preferencesUtilisateur("animations") Then
        Call ActiverAnimations
    End If
End Sub

' Désactiver les animations
Sub DesactiverAnimations()
    On Error Resume Next

    ' Remplacer les animations par des transitions instantanées
    ' Réduire les effets visuels
End Sub

' Activer les animations
Sub ActiverAnimations()
    On Error Resume Next

    ' Activer tous les effets visuels
    ' Utiliser les animations complètes
End Sub

' ========================================
' GESTION DES PÉRIPHÉRIQUES
' ========================================

' Détecter et adapter selon le périphérique d'entrée
Sub AdapterPeripheriques()
    On Error Resume Next

    ' Détecter si une souris est présente
    Dim sourisPresente As Boolean
    sourisPresente = True ' Simulation

    If Not sourisPresente Then
        ' Optimiser pour le tactile
        Call OptimiserInterfaceTactile
    End If
End Sub

' Optimiser pour interface tactile
Sub OptimiserInterfaceTactile()
    On Error Resume Next

    ' Boutons plus grands
    BUTTON_WIDTH = 200
    BUTTON_HEIGHT = 60

    ' Espacement plus important
    ELEMENT_SPACING = 30
    FORM_PADDING = 30

    ' Désactiver les tooltips (pas utiles sur tactile)
    preferencesUtilisateur("tooltips") = False

    ' Augmenter la taille des zones cliquables
    MIN_CLICKABLE_SIZE = 60
End Sub

' ========================================
' SAUVEGARDE DES PRÉFÉRENCES
' ========================================

' Sauvegarder les préférences utilisateur
Sub SauvegarderPreferences()
    On Error Resume Next

    ' En production, sauvegarder dans un fichier ou base de données
    ' Pour l'exemple, on utilise les propriétés du document
    ThisWorkbook.Sheets("Preferences").Range("A1") = "theme:" & preferencesUtilisateur("theme")
    ThisWorkbook.Sheets("Preferences").Range("A2") = "taillePolice:" & preferencesUtilisateur("taillePolice")
    ThisWorkbook.Sheets("Preferences").Range("A3") = "animations:" & preferencesUtilisateur("animations")
End Sub

' Charger les préférences utilisateur
Sub ChargerPreferences()
    On Error Resume Next

    Dim prefs As String
    Dim paires As Variant
    Dim paire As Variant
    Dim i As Integer

    ' Charger depuis le document (simulation)
    prefs = ThisWorkbook.Sheets("Preferences").Range("A1").Value & ";" & _
            ThisWorkbook.Sheets("Preferences").Range("A2").Value & ";" & _
            ThisWorkbook.Sheets("Preferences").Range("A3").Value

    paires = Split(prefs, ";")

    For i = LBound(paires) To UBound(paires)
        If Len(paires(i)) > 0 Then
            paire = Split(paires(i), ":")
            If UBound(paire) = 1 Then
                preferencesUtilisateur(paire(0)) = paire(1)
            End If
        End If
    Next i
End Sub

' ========================================
' FONCTIONS UTILITAIRES
' ========================================

' Obtenir la taille d'écran actuelle
Function ObtenirTailleEcran() As String
    ObtenirTailleEcran = tailleEcran
End Function

' Obtenir la résolution actuelle
Function ObtenirResolutionEcran() As String
    ObtenirResolutionEcran = resolutionEcran
End Function

' Vérifier si l'écran est petit
Function EstEcranPetit() As Boolean
    EstEcranPetit = (tailleEcran = "Petit" Or tailleEcran = "TresPetit")
End Function

' Vérifier si l'écran est grand
Function EstEcranGrand() As Boolean
    EstEcranGrand = (tailleEcran = "Grand" Or tailleEcran = "TresGrand")
End Function
