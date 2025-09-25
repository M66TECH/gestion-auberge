# Manuel Utilisateur - Application Gestion Auberge

## Table des matières
1. [Introduction](#introduction)
2. [Installation et Configuration](#installation-et-configuration)
3. [Interface Principale](#interface-principale)
4. [Gestion des Chambres](#gestion-des-chambres)
5. [Gestion des Clients](#gestion-des-clients)
6. [Système de Réservations](#système-de-réservations)
7. [Facturation et Paiements](#facturation-et-paiements)
8. [Rapports et Statistiques](#rapports-et-statistiques)
9. [Dépannage](#dépannage)

## Introduction

L'application **Gestion Auberge** est un outil complet développé en Excel VBA pour gérer efficacement les opérations quotidiennes d'une auberge. Elle permet de :

- 🛏️ Gérer les chambres et leurs disponibilités
- 👤 Maintenir une base de données clients
- 📅 Créer et suivre les réservations
- 💳 Gérer la facturation et les paiements
- 📊 Générer des rapports et statistiques

## Installation et Configuration

### Prérequis
- Microsoft Excel 2016 ou version supérieure
- Windows (recommandé)
- Macros VBA activées

### Installation
1. Ouvrez le fichier `GestionAuberge.xlsm`
2. Si Excel demande d'activer les macros, cliquez sur **Activer le contenu**
3. Exécutez la macro `InitialiserApplication` pour configurer l'application

### Première utilisation
Lors du premier lancement, l'application :
- Crée automatiquement toutes les feuilles nécessaires
- Configure les en-têtes et le formatage
- Initialise des données d'exemple
- Affiche le tableau de bord principal

## Interface Principale

### Tableau de bord
Le **Dashboard** est votre point de départ. Il affiche :
- 📅 Les réservations du jour (arrivées et départs)
- 📊 Statistiques en temps réel
- 🔧 Boutons d'accès rapide aux fonctions principales

### Navigation
L'application utilise plusieurs feuilles Excel :
- **Dashboard** : Tableau de bord principal
- **Chambres** : Liste et gestion des chambres
- **Clients** : Base de données clients
- **Reservations** : Suivi des réservations
- **Paiements** : Historique des paiements
- **Rapports** : Zone de génération des rapports
- **Parametres** : Configuration de l'auberge

## Gestion des Chambres

### Ajouter une chambre
```vba
' Exemple d'utilisation
Call AjouterChambre("103", "Double", 85, "Chambre double avec vue mer", "TV, WiFi, Balcon")
```

### Modifier une chambre
```vba
Call ModifierChambre("103", "Suite", 120, "Libre", "Suite familiale", "TV, WiFi, Salon, Balcon")
```

### Changer le statut d'une chambre
Les statuts disponibles sont :
- **Libre** : Chambre disponible à la réservation
- **Occupée** : Chambre actuellement occupée
- **Maintenance** : Chambre indisponible pour maintenance

```vba
Call ChangerStatutChambre("103", "Maintenance")
```

### Fonctions utiles
- `ChambreExiste("103")` : Vérifier si une chambre existe
- `ObtenirChambresLibres()` : Liste des chambres disponibles
- `ObtenirTarifChambre("103")` : Obtenir le tarif d'une chambre

## Gestion des Clients

### Ajouter un client
```vba
Dim idClient As Long
idClient = AjouterClient("Dupont", "Jean", "0123456789", "jean.dupont@email.com", "123 Rue de la Paix, Paris")
```

### Rechercher un client
```vba
' Par ID
Dim clientInfo As Variant
clientInfo = RechercherClientParID(1)

' Par nom
Dim clients As Variant
clients = RechercherClientsParNom("Dupont")
```

### Modifier un client
```vba
Call ModifierClient(1, "Dupont", "Jean", "0987654321", "nouveau.email@email.com", "Nouvelle adresse")
```

### Historique client
```vba
Dim historique As Variant
historique = ObtenirHistoriqueClient(1)
```

## Système de Réservations

### Créer une réservation
```vba
Dim idReservation As Long
idReservation = CreerReservation(1, "103", #01/15/2024#, #01/18/2024#, "Séjour d'affaires")
```

### Confirmer une réservation
```vba
Call ConfirmerReservation(1)
```

### Annuler une réservation
```vba
Call AnnulerReservation(1, "Annulation client")
```

### Check-in et Check-out
```vba
' Arrivée du client
Call EffectuerCheckIn(1)

' Départ du client
Call EffectuerCheckOut(1)
```

### Recherches
```vba
' Réservations d'un client
Dim reservations As Variant
reservations = RechercherReservationsParClient(1)

' Réservations d'une date
reservations = RechercherReservationsParDate(#01/15/2024#)

' Réservations du jour
reservations = ObtenirReservationsDuJour()
```

## Facturation et Paiements

### Enregistrer un paiement
```vba
Dim idPaiement As Long
idPaiement = EnregistrerPaiement(1, 255.0, "Carte bancaire", "Acompte")
```

**Modes de paiement disponibles :**
- Espèces
- Carte bancaire
- Chèque
- Virement

**Types de paiement :**
- Acompte
- Solde
- Total

### Générer une facture
```vba
Call GenererFacture(1)
```
Cette fonction crée automatiquement une nouvelle feuille avec la facture formatée.

### Suivi des paiements
```vba
' Montant déjà payé
Dim montantPaye As Double
montantPaye = MontantDejaPayé(1)

' Montant restant à payer
Dim montantRestant As Double
montantRestant = MontantRestantAPayer(1)

' Vérifier si soldé
Dim estSolde As Boolean
estSolde = ReservationSoldee(1)
```

### Réservations non soldées
```vba
Dim nonSoldees As Variant
nonSoldees = ObtenirReservationsNonSoldees()
```

## Rapports et Statistiques

### Rapport mensuel
```vba
Call GenererRapportMensuel(12, 2024) ' Décembre 2024
```

### Statistiques personnalisées
```vba
' Chiffre d'affaires sur une période
Dim ca As Double
ca = CalculerChiffreAffaires(#01/01/2024#, #01/31/2024#)

' Taux d'occupation
Dim tauxOccupation As Double
tauxOccupation = CalculerTauxOccupation(#01/01/2024#, #01/31/2024#)

' Nombre de réservations
Dim nbReservations As Long
nbReservations = CompterReservationsPeriode(#01/01/2024#, #01/31/2024#)
```

## Fonctions d'Administration

### Actualiser le tableau de bord
```vba
Call ActualiserDashboard()
```

### Sauvegarder les données
Excel sauvegarde automatiquement, mais vous pouvez forcer la sauvegarde :
```vba
ThisWorkbook.Save
```

### Paramètres de l'auberge
Modifiez les paramètres dans la feuille **Parametres** :
- Nom de l'auberge
- Adresse
- Téléphone
- Email
- Taux de TVA

## Dépannage

### Problèmes courants

**Les macros ne fonctionnent pas**
- Vérifiez que les macros sont activées dans Excel
- Allez dans Fichier > Options > Centre de gestion de la confidentialité > Paramètres des macros

**Erreur "Feuille non trouvée"**
- Exécutez `InitialiserApplication()` pour recréer les feuilles manquantes

**Données corrompues**
- Utilisez une sauvegarde récente
- Vérifiez l'intégrité des données dans chaque feuille

**Performance lente**
- Fermez les autres applications
- Réduisez le nombre de formules complexes
- Nettoyez les données anciennes

### Support technique
Pour obtenir de l'aide :
1. Consultez ce manuel
2. Vérifiez les commentaires dans le code VBA
3. Contactez l'administrateur système

### Bonnes pratiques
- Sauvegardez régulièrement vos données
- Ne modifiez pas directement les en-têtes des colonnes
- Utilisez les fonctions VBA plutôt que la saisie manuelle
- Vérifiez les données avant validation
- Maintenez les paramètres à jour

---

**Version :** 1.0  
**Date :** 2024  
**Auteur :** Application VBA Gestion Auberge
