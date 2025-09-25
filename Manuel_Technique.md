# Manuel Technique - Application Gestion Auberge

## Architecture du Projet

### Structure des Modules VBA

#### ModulePrincipal.bas
- **Rôle** : Module central contenant l'initialisation et les constantes
- **Fonctions clés** :
  - `InitialiserApplication()` : Point d'entrée principal
  - `VerifierStructureFeuilles()` : Création automatique des feuilles
  - `ConfigurerFeuille()` : Configuration des en-têtes et formatage

#### ModuleChambres.bas
- **Rôle** : Gestion complète des chambres
- **Fonctions principales** :
  - `AjouterChambre()` : Création de nouvelles chambres
  - `ModifierChambre()` : Modification des propriétés
  - `ChangerStatutChambre()` : Gestion des statuts (Libre/Occupée/Maintenance)
  - `ChambreDisponible()` : Vérification de disponibilité avec gestion des conflits

#### ModuleClients.bas
- **Rôle** : Base de données clients
- **Fonctions principales** :
  - `AjouterClient()` : Création avec ID auto-incrémenté
  - `RechercherClientParID()` / `RechercherClientsParNom()` : Recherches multiples
  - `ValiderEmail()` / `ValiderTelephone()` : Validation des données

#### ModuleReservations.bas
- **Rôle** : Système de réservations complet
- **Fonctions principales** :
  - `CreerReservation()` : Création avec validation des dates et disponibilité
  - `ConfirmerReservation()` / `AnnulerReservation()` : Gestion des statuts
  - `EffectuerCheckIn()` / `EffectuerCheckOut()` : Processus d'arrivée/départ

#### ModulePaiements.bas
- **Rôle** : Facturation et gestion financière
- **Fonctions principales** :
  - `EnregistrerPaiement()` : Saisie avec validation des montants
  - `GenererFacture()` : Création automatique de factures formatées
  - `CalculerChiffreAffaires()` : Statistiques financières

#### ModuleRapports.bas
- **Rôle** : Génération de rapports et statistiques
- **Fonctions principales** :
  - `GenererRapportMensuel()` : Rapport complet avec métriques
  - `CalculerTauxOccupation()` : Calculs de performance

#### Dashboard.bas
- **Rôle** : Interface utilisateur principale
- **Fonctions principales** :
  - `ConfigurerFeuilleDashboard()` : Mise en page dynamique
  - `ActualiserDashboard()` : Rafraîchissement des données

### Base de Données Excel

#### Structure des Feuilles

**Feuille "Chambres"**
```
A: NumChambre (Texte) - Clé primaire
B: TypeChambre (Texte) - Simple/Double/Suite
C: TarifNuit (Numérique) - Prix par nuit
D: Statut (Texte) - Libre/Occupée/Maintenance
E: Description (Texte) - Description de la chambre
F: Equipements (Texte) - Liste des équipements
```

**Feuille "Clients"**
```
A: IDClient (Numérique) - Clé primaire auto-incrémentée
B: Nom (Texte) - Nom du client
C: Prenom (Texte) - Prénom du client
D: Telephone (Texte) - Numéro de téléphone
E: Email (Texte) - Adresse email
F: Adresse (Texte) - Adresse complète
G: DateCreation (Date) - Date de création du profil
```

**Feuille "Reservations"**
```
A: IDReservation (Numérique) - Clé primaire auto-incrémentée
B: IDClient (Numérique) - Clé étrangère vers Clients
C: NumChambre (Texte) - Clé étrangère vers Chambres
D: DateArrivee (Date) - Date d'arrivée
E: DateDepart (Date) - Date de départ
F: NbNuits (Numérique) - Calculé automatiquement
G: MontantTotal (Numérique) - Calculé automatiquement
H: Statut (Texte) - Confirmée/En attente/Annulée
I: DateReservation (Date) - Date de création
J: Commentaires (Texte) - Notes particulières
```

**Feuille "Paiements"**
```
A: IDPaiement (Numérique) - Clé primaire auto-incrémentée
B: IDReservation (Numérique) - Clé étrangère vers Reservations
C: Montant (Numérique) - Montant payé
D: ModePaiement (Texte) - Espèces/CB/Chèque/Virement
E: DatePaiement (Date) - Date du paiement
F: TypePaiement (Texte) - Acompte/Solde/Total
G: Statut (Texte) - Validé/En attente/Refusé
```

### Algorithmes Clés

#### Vérification de Disponibilité
```vba
Function ChambreDisponible(numChambre As String, dateArrivee As Date, dateDepart As Date) As Boolean
    ' 1. Vérifier existence de la chambre
    ' 2. Vérifier statut "Libre"
    ' 3. Parcourir toutes les réservations confirmées
    ' 4. Détecter les conflits de dates : NOT (dateDepart <= dateArrRes OR dateArrivee >= dateDepartRes)
End Function
```

#### Calcul du Taux d'Occupation
```vba
Function CalculerTauxOccupation(dateDebut As Date, dateFin As Date) As Double
    ' Formule : (Nuits occupées / (Nombre de chambres × Nombre de jours)) × 100
    ' Prend en compte uniquement les réservations confirmées
End Function
```

#### Génération d'ID Auto-incrémenté
```vba
Function ObtenirProchainID() As Long
    ' Parcourt toute la colonne ID
    ' Trouve la valeur maximale
    ' Retourne max + 1
End Function
```

### Gestion des Erreurs

#### Stratégie Globale
- Utilisation systématique de `On Error GoTo ErrHandler`
- Messages d'erreur informatifs pour l'utilisateur
- Logging des erreurs critiques
- Restauration de l'état d'Excel (ScreenUpdating, DisplayAlerts)

#### Validations de Données
- **Dates** : Vérification cohérence arrivée/départ, pas de dates passées
- **Montants** : Valeurs positives, pas de dépassement du montant total
- **Références** : Existence des clients, chambres, réservations
- **Formats** : Email, téléphone avec expressions régulières simplifiées

### Performance et Optimisation

#### Bonnes Pratiques Implémentées
```vba
Application.ScreenUpdating = False  ' Désactiver rafraîchissement écran
Application.DisplayAlerts = False   ' Désactiver alertes Excel
' ... code de traitement ...
Application.ScreenUpdating = True   ' Réactiver à la fin
Application.DisplayAlerts = True
```

#### Gestion Mémoire
- Libération explicite des objets Worksheet
- Utilisation de variables locales
- Éviter les boucles infinies avec compteurs de sécurité

### Sécurité

#### Protection des Données
- Feuilles de données protégées en écriture
- Accès uniquement via les fonctions VBA
- Validation systématique des entrées utilisateur

#### Sauvegarde
- Sauvegarde automatique lors des modifications importantes
- Recommandation de sauvegardes externes régulières

### Installation et Déploiement

#### Prérequis Techniques
- Excel 2016+ (utilisation de fonctions récentes)
- Macros VBA activées
- Droits d'écriture sur le répertoire de travail

#### Procédure d'Installation
1. Copier le fichier .xlsm
2. Ouvrir et activer les macros
3. Exécuter `InitialiserApplication()`
4. Configurer les paramètres dans la feuille Parametres

### Maintenance

#### Tâches Régulières
- Nettoyage des données anciennes (réservations > 2 ans)
- Vérification de l'intégrité des liens entre tables
- Mise à jour des tarifs et paramètres

#### Évolutions Possibles
- **Base de données externe** : Migration vers Access ou SQL Server
- **Interface web** : Développement d'une interface HTML/JavaScript
- **API emails** : Intégration avec Outlook pour confirmations automatiques
- **Graphiques dynamiques** : Tableaux de bord avec charts Excel

### Dépannage Technique

#### Erreurs Courantes
- **Erreur 1004** : Feuille non trouvée → Réexécuter InitialiserApplication()
- **Erreur 13** : Type incompatible → Vérifier format des dates/nombres
- **Erreur 9** : Index hors limites → Vérifier existence des données

#### Outils de Debug
- Utiliser F8 pour exécution pas à pas
- Points d'arrêt sur les lignes critiques
- Fenêtre Immédiat pour tester les fonctions

### Code de Maintenance

#### Nettoyage des Données
```vba
Sub NettoyerDonneesAnciennes()
    ' Supprimer réservations > 2 ans
    ' Archiver paiements anciens
    ' Compacter les ID
End Sub
```

#### Vérification d'Intégrité
```vba
Sub VerifierIntegriteDonnees()
    ' Vérifier cohérence Client-Réservation
    ' Vérifier cohérence Réservation-Paiement
    ' Signaler les incohérences
End Sub
```

---

**Développé avec Excel VBA**  
**Compatible Excel 2016+**  
**Architecture modulaire pour faciliter la maintenance**
