Attribute VB_Name = "DonneesTest"
' ========================================
' MODULE GENERATION DONNEES DE TEST
' ========================================
' Description: Génération de données fictives pour tester l'application

Option Explicit

' ========================================
' GENERER TOUTES LES DONNEES DE TEST
' ========================================
Sub GenererDonneesTest()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    ' Vider les données existantes (sauf en-têtes)
    Call ViderDonnees
    
    ' Générer les données
    Call GenererChambresTest
    Call GenererClientsTest
    Call GenererReservationsTest
    Call GenererPaiementsTest
    
    ' Actualiser le dashboard
    Call ActualiserDashboard
    
    Application.ScreenUpdating = True
    
    MsgBox "Données de test générées avec succès !" & vbCrLf & _
           "- 10 chambres" & vbCrLf & _
           "- 15 clients" & vbCrLf & _
           "- 20 réservations" & vbCrLf & _
           "- Paiements associés", vbInformation, APP_NAME
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de la génération des données : " & Err.Description, vbCritical, APP_NAME
End Sub

' ========================================
' VIDER LES DONNEES EXISTANTES
' ========================================
Sub ViderDonnees()
    Dim ws As Worksheet
    
    ' Vider Chambres (garder en-têtes)
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CHAMBRES)
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then
        ws.Range("A2:F" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row).Clear
    End If
    
    ' Vider Clients
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CLIENTS)
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then
        ws.Range("A2:G" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row).Clear
    End If
    
    ' Vider Réservations
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESERVATIONS)
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then
        ws.Range("A2:J" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row).Clear
    End If
    
    ' Vider Paiements
    Set ws = ThisWorkbook.Worksheets(FEUILLE_PAIEMENTS)
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then
        ws.Range("A2:G" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row).Clear
    End If
End Sub

' ========================================
' GENERER CHAMBRES DE TEST
' ========================================
Sub GenererChambresTest()
    ' Chambres simples
    Call AjouterChambre("101", "Simple", 65, "Chambre simple avec vue jardin", "TV, WiFi, Salle de bain")
    Call AjouterChambre("102", "Simple", 65, "Chambre simple standard", "TV, WiFi, Salle de bain")
    Call AjouterChambre("103", "Simple", 70, "Chambre simple avec balcon", "TV, WiFi, Balcon, Salle de bain")
    
    ' Chambres doubles
    Call AjouterChambre("201", "Double", 85, "Chambre double avec vue mer", "TV, WiFi, Balcon, Salle de bain")
    Call AjouterChambre("202", "Double", 85, "Chambre double standard", "TV, WiFi, Salle de bain")
    Call AjouterChambre("203", "Double", 90, "Chambre double supérieure", "TV, WiFi, Minibar, Salle de bain")
    Call AjouterChambre("204", "Double", 85, "Chambre double avec terrasse", "TV, WiFi, Terrasse, Salle de bain")
    
    ' Suites
    Call AjouterChambre("301", "Suite", 120, "Suite familiale avec salon", "TV, WiFi, Salon, Balcon, Salle de bain")
    Call AjouterChambre("302", "Suite", 140, "Suite de luxe avec jacuzzi", "TV, WiFi, Salon, Jacuzzi, Balcon")
    Call AjouterChambre("303", "Suite", 130, "Suite junior avec vue panoramique", "TV, WiFi, Salon, Vue panoramique")
End Sub

' ========================================
' GENERER CLIENTS DE TEST
' ========================================
Sub GenererClientsTest()
    Call AjouterClient("Dupont", "Jean", "0123456789", "jean.dupont@email.com", "123 Rue de la Paix, 75001 Paris")
    Call AjouterClient("Martin", "Marie", "0234567890", "marie.martin@email.com", "456 Avenue des Champs, 69000 Lyon")
    Call AjouterClient("Bernard", "Pierre", "0345678901", "pierre.bernard@email.com", "789 Boulevard Victor Hugo, 13000 Marseille")
    Call AjouterClient("Dubois", "Sophie", "0456789012", "sophie.dubois@email.com", "321 Rue de la République, 31000 Toulouse")
    Call AjouterClient("Moreau", "Michel", "0567890123", "michel.moreau@email.com", "654 Place Bellecour, 69002 Lyon")
    Call AjouterClient("Laurent", "Catherine", "0678901234", "catherine.laurent@email.com", "987 Rue de Rivoli, 75001 Paris")
    Call AjouterClient("Simon", "David", "0789012345", "david.simon@email.com", "147 Avenue de la Liberté, 06000 Nice")
    Call AjouterClient("Michel", "Isabelle", "0890123456", "isabelle.michel@email.com", "258 Cours Mirabeau, 13100 Aix-en-Provence")
    Call AjouterClient("Garcia", "Carlos", "0901234567", "carlos.garcia@email.com", "369 Rue Saint-Antoine, 33000 Bordeaux")
    Call AjouterClient("Rodriguez", "Ana", "0123456780", "ana.rodriguez@email.com", "741 Boulevard de la Croisette, 06400 Cannes")
    Call AjouterClient("Leroy", "François", "0234567801", "francois.leroy@email.com", "852 Rue du Faubourg Saint-Honoré, 75008 Paris")
    Call AjouterClient("Roux", "Nathalie", "0345678012", "nathalie.roux@email.com", "963 Avenue Montaigne, 75008 Paris")
    Call AjouterClient("Fournier", "Alain", "0456780123", "alain.fournier@email.com", "159 Quai des Grands Augustins, 75006 Paris")
    Call AjouterClient("Girard", "Sylvie", "0567801234", "sylvie.girard@email.com", "357 Rue de la Pompe, 75016 Paris")
    Call AjouterClient("Bonnet", "Thierry", "0678012345", "thierry.bonnet@email.com", "468 Avenue Kléber, 75016 Paris")
End Sub

' ========================================
' GENERER RESERVATIONS DE TEST
' ========================================
Sub GenererReservationsTest()
    Dim i As Integer
    Dim dateBase As Date
    Dim idRes As Long
    
    dateBase = Date - 30 ' Commencer il y a 30 jours
    
    ' Réservations passées (confirmées et soldées)
    idRes = CreerReservation(1, "101", dateBase, dateBase + 3, "Séjour d'affaires")
    Call ConfirmerReservation(idRes)
    Call EnregistrerPaiement(idRes, 195, "Carte bancaire", "Total")
    
    idRes = CreerReservation(2, "201", dateBase + 5, dateBase + 7, "Week-end romantique")
    Call ConfirmerReservation(idRes)
    Call EnregistrerPaiement(idRes, 170, "Espèces", "Total")
    
    idRes = CreerReservation(3, "301", dateBase + 10, dateBase + 14, "Vacances en famille")
    Call ConfirmerReservation(idRes)
    Call EnregistrerPaiement(idRes, 240, "Acompte", "Acompte")
    Call EnregistrerPaiement(idRes, 240, "Carte bancaire", "Solde")
    
    ' Réservations actuelles (en cours)
    dateBase = Date - 2
    idRes = CreerReservation(4, "102", dateBase, Date + 1, "Conférence professionnelle")
    Call ConfirmerReservation(idRes)
    Call EffectuerCheckIn(idRes)
    Call EnregistrerPaiement(idRes, 130, "Carte bancaire", "Acompte")
    
    idRes = CreerReservation(5, "202", dateBase, Date + 2, "Séjour touristique")
    Call ConfirmerReservation(idRes)
    Call EffectuerCheckIn(idRes)
    Call EnregistrerPaiement(idRes, 255, "Espèces", "Total")
    
    ' Réservations futures confirmées
    dateBase = Date + 1
    idRes = CreerReservation(6, "103", dateBase, dateBase + 2, "Arrivée demain")
    Call ConfirmerReservation(idRes)
    Call EnregistrerPaiement(idRes, 70, "Virement", "Acompte")
    
    idRes = CreerReservation(7, "203", Date + 3, Date + 6, "Séjour de détente")
    Call ConfirmerReservation(idRes)
    
    idRes = CreerReservation(8, "302", Date + 5, Date + 8, "Lune de miel")
    Call ConfirmerReservation(idRes)
    Call EnregistrerPaiement(idRes, 210, "Carte bancaire", "Acompte")
    
    ' Réservations en attente
    idRes = CreerReservation(9, "204", Date + 7, Date + 10, "À confirmer")
    idRes = CreerReservation(10, "303", Date + 10, Date + 13, "Demande de groupe")
    
    ' Réservations futures variées
    idRes = CreerReservation(11, "101", Date + 15, Date + 18, "Voyage d'affaires")
    Call ConfirmerReservation(idRes)
    
    idRes = CreerReservation(12, "201", Date + 20, Date + 22, "Week-end")
    Call ConfirmerReservation(idRes)
    Call EnregistrerPaiement(idRes, 85, "Chèque", "Acompte")
    
    ' Quelques réservations annulées pour l'historique
    idRes = CreerReservation(13, "301", Date + 25, Date + 28, "Annulation test")
    Call ConfirmerReservation(idRes)
    Call AnnulerReservation(idRes, "Changement de programme client")
    
    idRes = CreerReservation(14, "102", Date + 30, Date + 33, "Réservation longue")
    Call ConfirmerReservation(idRes)
    
    idRes = CreerReservation(15, "302", Date + 35, Date + 37, "Suite de luxe")
    Call ConfirmerReservation(idRes)
    Call EnregistrerPaiement(idRes, 140, "Carte bancaire", "Acompte")
End Sub

' ========================================
' GENERER PAIEMENTS COMPLEMENTAIRES
' ========================================
Sub GenererPaiementsTest()
    ' Les paiements sont générés avec les réservations
    ' Cette fonction peut ajouter des paiements supplémentaires si nécessaire
    
    ' Exemple : paiement de solde pour une réservation
    ' Call EnregistrerPaiement(7, 180, "Carte bancaire", "Solde")
End Sub

' ========================================
' GENERER DONNEES POUR DEMONSTRATION
' ========================================
Sub GenererDonneesDemo()
    Call GenererDonneesTest
    
    ' Ajouter quelques statistiques intéressantes
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RAPPORTS)
    
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "📊 DONNÉES DE DÉMONSTRATION GÉNÉRÉES"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    
    ws.Cells(3, 1).Value = "Contenu généré :"
    ws.Cells(4, 1).Value = "• 10 chambres (Simple, Double, Suite)"
    ws.Cells(5, 1).Value = "• 15 clients avec coordonnées complètes"
    ws.Cells(6, 1).Value = "• 15 réservations sur 2 mois"
    ws.Cells(7, 1).Value = "• Paiements partiels et complets"
    ws.Cells(8, 1).Value = "• Réservations passées, actuelles et futures"
    ws.Cells(9, 1).Value = "• Différents statuts (confirmée, en attente, annulée)"
    
    ws.Cells(11, 1).Value = "Utilisez le Dashboard pour naviguer dans l'application !"
    ws.Cells(11, 1).Font.Bold = True
    ws.Cells(11, 1).Font.Color = RGB(0, 128, 0)
    
    ws.Columns("A:B").AutoFit
    
    MsgBox "Démonstration prête ! Consultez le Dashboard pour explorer toutes les fonctionnalités.", _
           vbInformation, APP_NAME
End Sub

' ========================================
' REINITIALISER L'APPLICATION
' ========================================
Sub ReinitialiserApplication()
    Dim reponse As VbMsgBoxResult
    
    reponse = MsgBox("Cette action va supprimer toutes les données et réinitialiser l'application." & vbCrLf & _
                     "Êtes-vous sûr de vouloir continuer ?", vbYesNo + vbExclamation, APP_NAME)
    
    If reponse = vbYes Then
        Call ViderDonnees
        Call InitialiserDonneesBase
        Call ActualiserDashboard
        MsgBox "Application réinitialisée avec succès !", vbInformation, APP_NAME
    End If
End Sub
