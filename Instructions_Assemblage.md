# Instructions d'Assemblage - Application Gestion Auberge

## Étapes pour créer le fichier Excel complet

### 1. Créer le fichier Excel de base
1. Ouvrir Microsoft Excel
2. Créer un nouveau classeur
3. Enregistrer sous le nom `GestionAuberge.xlsm` (format Macro Excel)
4. Activer l'onglet Développeur si nécessaire (Fichier > Options > Personnaliser le ruban)

### 2. Importer les modules VBA

#### Accéder à l'éditeur VBA
- Appuyer sur `Alt + F11` ou aller dans Développeur > Visual Basic

#### Importer chaque module
Pour chaque fichier .bas, suivre ces étapes :
1. Dans l'éditeur VBA : Fichier > Importer un fichier
2. Sélectionner le fichier .bas correspondant
3. Le module apparaît dans l'explorateur de projet

#### Liste des modules à importer
- `VBA_ModulePrincipal.bas` → Module "ModulePrincipal"
- `VBA_ModuleChambres.bas` → Module "ModuleChambres"  
- `VBA_ModuleClients.bas` → Module "ModuleClients"
- `VBA_ModuleReservations.bas` → Module "ModuleReservations"
- `VBA_ModulePaiements.bas` → Module "ModulePaiements"
- `VBA_ModuleRapports.bas` → Module "ModuleRapports"
- `VBA_Dashboard.bas` → Module "Dashboard"
- `VBA_DonneesTest.bas` → Module "DonneesTest"

### 3. Configuration initiale

#### Exécuter l'initialisation
1. Dans l'éditeur VBA, aller dans le module "ModulePrincipal"
2. Placer le curseur dans la fonction `InitialiserApplication`
3. Appuyer sur F5 pour exécuter
4. Ou depuis Excel : Développeur > Macros > InitialiserApplication > Exécuter

#### Vérifier la création des feuilles
Après initialisation, le classeur doit contenir :
- Dashboard (feuille principale)
- Chambres
- Clients  
- Reservations
- Paiements
- Parametres
- Rapports

### 4. Générer les données de test (optionnel)

#### Pour une démonstration complète
1. Exécuter la macro `GenererDonneesDemo` depuis le module "DonneesTest"
2. Cela créera automatiquement :
   - 10 chambres d'exemple
   - 15 clients fictifs
   - 15 réservations variées
   - Paiements associés

### 5. Configuration des paramètres

#### Personnaliser l'auberge
Dans la feuille "Parametres", modifier :
- NomAuberge : Nom de votre établissement
- AdresseAuberge : Adresse complète
- TelephoneAuberge : Numéro de téléphone
- EmailAuberge : Adresse email
- TauxTVA : Taux de TVA applicable

### 6. Protection et sécurité

#### Protéger les feuilles de données
1. Sélectionner chaque feuille de données (Chambres, Clients, etc.)
2. Clic droit > Protéger la feuille
3. Laisser cochées les options de base
4. Définir un mot de passe si souhaité

#### Masquer les feuilles techniques
1. Clic droit sur les onglets "Parametres" et "Rapports"
2. Masquer (les données restent accessibles via VBA)

### 7. Interface utilisateur

#### Créer des boutons (optionnel)
Pour une interface plus conviviale :
1. Dans la feuille Dashboard
2. Développeur > Insérer > Bouton (Contrôles de formulaire)
3. Dessiner le bouton et associer une macro

#### Exemple de boutons utiles
- "Actualiser Dashboard" → Macro `ActualiserDashboard`
- "Nouvelle Réservation" → Ouvrir la feuille Reservations
- "Générer Rapport" → Macro `GenererRapportMensuel`

### 8. Tests et validation

#### Tester les fonctionnalités principales
1. **Chambres** : Ajouter/modifier une chambre
2. **Clients** : Créer un nouveau client
3. **Réservations** : Créer une réservation complète
4. **Paiements** : Enregistrer un paiement
5. **Facturation** : Générer une facture
6. **Rapports** : Créer un rapport mensuel

#### Vérifier les validations
- Dates cohérentes pour les réservations
- Disponibilité des chambres
- Montants des paiements
- Intégrité des données

### 9. Finalisation

#### Sauvegarder le projet
1. Enregistrer le fichier .xlsm
2. Créer une copie de sauvegarde
3. Tester l'ouverture et l'activation des macros

#### Documentation
- Placer les fichiers de documentation dans le même dossier
- `Manuel_Utilisateur.md` : Guide d'utilisation
- `Manuel_Technique.md` : Documentation technique
- `README.md` : Vue d'ensemble du projet

### 10. Déploiement

#### Préparer la distribution
1. Créer un dossier "GestionAuberge_v1.0"
2. Inclure :
   - `GestionAuberge.xlsm` (fichier principal)
   - `Manuel_Utilisateur.pdf` (converti depuis .md)
   - `README.txt` (instructions rapides)

#### Instructions pour l'utilisateur final
```
1. Ouvrir GestionAuberge.xlsm
2. Cliquer "Activer le contenu" si demandé
3. L'application s'initialise automatiquement
4. Utiliser le Dashboard pour naviguer
```

## Dépannage de l'assemblage

### Erreurs courantes

**"Erreur de compilation"**
- Vérifier que tous les modules sont importés
- Contrôler la syntaxe VBA
- S'assurer de la compatibilité Excel

**"Macro introuvable"**
- Vérifier les noms des fonctions
- Contrôler l'importation des modules
- Redémarrer Excel si nécessaire

**"Feuille non trouvée"**
- Exécuter `InitialiserApplication`
- Vérifier les constantes de noms de feuilles
- Contrôler les références

### Optimisations avancées

#### Performance
- Désactiver le calcul automatique pour les gros volumes
- Utiliser des plages nommées pour les références fréquentes
- Optimiser les boucles VBA

#### Fonctionnalités supplémentaires
- Ajouter des graphiques dans le Dashboard
- Créer des UserForms pour la saisie
- Intégrer l'export PDF automatique

---

**Une fois assemblé, le fichier GestionAuberge.xlsm sera prêt à l'utilisation !**
