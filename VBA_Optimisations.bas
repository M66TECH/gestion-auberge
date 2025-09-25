Attribute VB_Name = "OptimisationsPerformance"
' ========================================
' MODULE OPTIMISATIONS PERFORMANCE & UX AVANCÉE
' ========================================
' Description: Optimisations pour des performances fluides et UX professionnelle

Option Explicit

' ========================================
' GESTION DU CACHE INTELLIGENT
' ========================================
Private dictCache As Object
Private tempsCache As Date

' Initialiser le système de cache
Sub InitialiserCache()
    On Error Resume Next
    
    If dictCache Is Nothing Then
        Set dictCache = CreateObject("Scripting.Dictionary")
        tempsCache = Now
    End If
End Sub

' Obtenir des données du cache
Function ObtenirDonneesCache(cle As String, fonctionChargement As String) As Variant
    On Error Resume Next
    
    Call InitialiserCache
    
    ' Vérifier si les données sont en cache et fraîches (< 5 minutes)
    If dictCache.Exists(cle) And DateDiff("n", tempsCache, Now) < 5 Then
        ObtenirDonneesCache = dictCache(cle)
    Else
        ' Charger les données et les mettre en cache
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
        Dim donnees As Variant
        donnees = Application.Run(fonctionChargement)
        
        dictCache(cle) = donnees
        tempsCache = Now
        
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        
        ObtenirDonneesCache = donnees
    End If
End Function

' Vider le cache si nécessaire
Sub ViderCache()
    On Error Resume Next
    
    If Not dictCache Is Nothing Then
        dictCache.RemoveAll
        tempsCache = Now
    End If
End Sub

' ========================================
' CHARGEMENT ASYNCHRONE OPTIMISÉ
' ========================================
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub ChargerDonneesAvecProgression(frm As Object, etapes As Collection)
    On Error Resume Next
    
    Dim progression As Double
    Dim etape As Variant
    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    ' Créer la barre de progression
    Call CreerBarreProgression(frm, 0, 100)
    
    For i = 1 To etapes.Count
        etape = etapes(i)
        
        ' Exécuter l'étape
        Application.Run etape("fonction"), etape("parametres")
        
        ' Mettre à jour la progression
        progression = (i / etapes.Count) * 100
        Call MettreAJourBarreProgression(frm, progression)
        
        ' Petite pause pour la fluidité
        Sleep 100
        DoEvents
    Next i
    
    Application.ScreenUpdating = True
    Call MasquerBarreProgression(frm)
    
    Call AfficherNotificationAvancee(frm, "Données chargées avec succès !", "success", 3)
End Sub

Sub MettreAJourBarreProgression(frm As Object, progression As Double)
    On Error Resume Next
    
    Dim barre As Object
    Set barre = frm.Controls("barreProgressionRemplie")
    
    If Not barre Is Nothing Then
        Dim container As Object
        Set container = frm.Controls("barreProgressionContainer")
        
        If Not container Is Nothing Then
            barre.Width = ((container.Width - 4) * progression) / 100
        End If
    End If
    
    Dim lblPourcentage As Object
    Set lblPourcentage = frm.Controls("lblPourcentageProgression")
    
    If Not lblPourcentage Is Nothing Then
        lblPourcentage.Caption = Format(progression, "0") & "%"
    End If
End Sub

' ========================================
' ANIMATIONS OPTIMISÉES
' ========================================

' Animation de fondu optimisée
Sub AnimerFonduOptimise(ctrl As Object, dureeMs As Integer)
    On Error Resume Next
    
    Dim pas As Integer
    Dim delai As Double
    Dim i As Integer
    
    pas = 20
    delai = dureeMs / pas
    
    ' Animation progressive
    For i = 0 To pas
        Dim opacite As Double
        opacite = i / pas
        
        ' Simuler l'opacité avec la couleur de fond
        If TypeName(ctrl) = "Label" Or TypeName(ctrl) = "CommandButton" Then
            Dim couleurOriginale As Long
            couleurOriginale = ctrl.BackColor
            ctrl.BackColor = MelangerCouleurs(COLOR_WHITE, couleurOriginale, opacite)
        End If
        
        DoEvents
        Sleep delai
    Next i
End Sub

' Animation de glissement optimisée
Sub AnimerGlissementOptimise(ctrl As Object, positionFinale As Integer, direction As String, dureeMs As Integer)
    On Error Resume Next
    
    Dim positionInitiale As Integer
    Dim distance As Integer
    Dim pas As Integer
    Dim increment As Double
    Dim i As Integer
    
    pas = 30
    positionInitiale = ctrl.Left
    distance = positionFinale - positionInitiale
    increment = distance / pas
    
    For i = 1 To pas
        ctrl.Left = ctrl.Left + increment
        DoEvents
        Sleep dureeMs / pas
    Next i
    
    ctrl.Left = positionFinale
End Sub

' ========================================
' GESTION DE MÉMOIRE OPTIMISÉE
' ========================================

' Nettoyer la mémoire
Sub NettoyerMemoire()
    On Error Resume Next
    
    ' Forcer le garbage collection VBA
    Dim obj As Object
    For Each obj In VBA.UserForms
        ' Vider les contrôles temporaires
        Call NettoyerControlesTemporaires(obj)
    Next obj
    
    ' Vider le cache si trop volumineux
    If Not dictCache Is Nothing Then
        If dictCache.Count > 50 Then
            Call ViderCache
        End If
    End If
    
    ' Forcer la libération de mémoire
    DoEvents
End Sub

' Nettoyer les contrôles temporaires
Sub NettoyerControlesTemporaires(frm As Object)
    On Error Resume Next
    
    Dim ctrl As Object
    Dim controlesASupprimer As Collection
    Set controlesASupprimer = New Collection
    
    ' Identifier les contrôles temporaires
    For Each ctrl In frm.Controls
        If InStr(ctrl.Name, "toast_") > 0 Or _
           InStr(ctrl.Name, "tooltip_") > 0 Or _
           InStr(ctrl.Name, "temp_") > 0 Or _
           InStr(ctrl.Name, "progression") > 0 Then
            controlesASupprimer.Add ctrl
        End If
    Next ctrl
    
    ' Supprimer les contrôles identifiés
    For Each ctrl In controlesASupprimer
        frm.Controls.Remove ctrl.Name
    Next ctrl
End Sub

' ========================================
' VALIDATION OPTIMISÉE
' ========================================

' Validation en arrière-plan
Sub ValiderEnArrierePlan(frm As Object, regles As Collection)
    On Error Resume Next
    
    Dim ctrl As Object
    Dim regle As Variant
    Dim erreurs As New Collection
    
    For Each ctrl In frm.Controls
        If regles.Exists(ctrl.Name) Then
            regle = regles(ctrl.Name)
            
            Dim estValide As Boolean
            estValide = True
            
            Select Case regle("type")
                Case "obligatoire"
                    estValide = Len(Trim(ctrl.Value)) > 0
                Case "email"
                    estValide = ValiderFormatEmailAvance(ctrl.Value)
                Case "telephone"
                    estValide = ValiderFormatTelephoneAvance(ctrl.Value)
                Case "date"
                    estValide = IsDate(ctrl.Value)
                Case "nombre"
                    estValide = IsNumeric(ctrl.Value)
                Case "longueur"
                    estValide = Len(ctrl.Value) >= regle("min") And Len(ctrl.Value) <= regle("max")
            End Select
            
            If Not estValide Then
                erreurs.Add Array(ctrl, regle("message"))
            End If
        End If
    Next ctrl
    
    ' Afficher les erreurs
    If erreurs.Count > 0 Then
        Dim erreur As Variant
        For Each erreur In erreurs
            Call AppliquerFocusVisuel(erreur(0), True)
            Call AfficherMessageErreur(frm, erreur(1))
        Next erreur
    End If
End Sub

' ========================================
' GESTION DES ÉTATS DE L'INTERFACE
' ========================================

' Mode chargement
Sub ActiverModeChargement(frm As Object, message As String)
    On Error Resume Next
    
    ' Désactiver tous les contrôles
    Dim ctrl As Object
    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "CommandButton" Then
            ctrl.Enabled = False
        End If
    Next ctrl
    
    ' Afficher l'indicateur de chargement
    Call CreerIndicateurChargement(frm, frm.Width / 2 - 15, frm.Height / 2 - 15)
    
    ' Afficher le message
    Dim lblMessage As Object
    Set lblMessage = frm.Controls.Add("Forms.Label.1", "lblChargementMessage")
    
    With lblMessage
        .Top = frm.Height / 2 + 20
        .Left = frm.Width / 2 - 100
        .Width = 200
        .Height = 20
        .Caption = message
        .TextAlign = fmTextAlignCenter
    End With
    
    Call StylerLabelNormal(lblMessage)
End Sub

' Désactiver le mode chargement
Sub DesactiverModeChargement(frm As Object)
    On Error Resume Next
    
    ' Réactiver tous les contrôles
    Dim ctrl As Object
    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "CommandButton" Then
            ctrl.Enabled = True
        End If
    Next ctrl
    
    ' Supprimer l'indicateur de chargement
    Call NettoyerControlesTemporaires(frm)
End Sub

' ========================================
' ANALYSE DES PERFORMANCES
' ========================================

' Mesurer le temps d'exécution
Function MesurerTempsExecution(fonction As String, ParamArray parametres() As Variant) As Double
    On Error Resume Next
    
    Dim tempsDebut As Double
    Dim tempsFin As Double
    Dim resultat As Variant
    
    tempsDebut = Timer
    
    ' Exécuter la fonction
    Select Case UBound(parametres)
        Case -1
            resultat = Application.Run(fonction)
        Case 0
            resultat = Application.Run(fonction, parametres(0))
        Case 1
            resultat = Application.Run(fonction, parametres(0), parametres(1))
        Case 2
            resultat = Application.Run(fonction, parametres(0), parametres(1), parametres(2))
    End Select
    
    tempsFin = Timer
    MesurerTempsExecution = tempsFin - tempsDebut
    
    ' Log si trop lent (> 1 seconde)
    If MesurerTempsExecution > 1 Then
        Debug.Print "Performance: " & fonction & " a pris " & Format(MesurerTempsExecution, "0.00") & " secondes"
    End If
End Function

' ========================================
' INITIALISATION OPTIMISÉE
' ========================================

' Initialiser l'application avec optimisations
Sub InitialiserApplicationOptimisee()
    On Error Resume Next
    
    ' Optimisations Excel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Initialiser le cache
    Call InitialiserCache
    
    ' Nettoyer la mémoire
    Call NettoyerMemoire
    
    ' Restaurer les paramètres
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Debug.Print "Application initialisée avec optimisations"
End Sub

' ========================================
' GESTION D'ERREURS OPTIMISÉE
' ========================================

' Gestionnaire d'erreurs centralisé
Sub GererErreurOptimisee(frm As Object, numeroErreur As Long, description As String, fonction As String)
    On Error Resume Next
    
    Dim messageErreur As String
    
    ' Créer un message d'erreur adapté
    Select Case numeroErreur
        Case 13 ' Type mismatch
            messageErreur = "Type de données incorrect dans " & fonction
        Case 9 ' Subscript out of range
            messageErreur = "Données manquantes ou incorrectes"
        Case 1004 ' Application-defined error
            messageErreur = "Erreur dans l'application : " & description
        Case Else
            messageErreur = "Erreur " & numeroErreur & " : " & description
    End Select
    
    ' Afficher l'erreur de manière élégante
    Call AfficherNotificationAvancee(frm, messageErreur, "error", 5)
    
    ' Log pour debugging
    Debug.Print "Erreur dans " & fonction & " : " & numeroErreur & " - " & description
    
    ' Tenter une récupération
    Call RecupererApresErreur(frm)
End Sub

' Récupération après erreur
Sub RecupererApresErreur(frm As Object)
    On Error Resume Next
    
    ' Réactiver les contrôles
    Call DesactiverModeChargement(frm)
    
    ' Vider le cache en cas d'erreur
    Call ViderCache
    
    ' Réinitialiser l'interface
    Call NettoyerControlesTemporaires(frm)
End Sub
