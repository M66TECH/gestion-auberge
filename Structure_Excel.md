# Structure des feuilles Excel - Gestion Auberge

## Feuilles de données

### 1. Feuille "Chambres"
| Colonne | Nom | Type | Description |
|---------|-----|------|-------------|
| A | NumChambre | Texte | Numéro de la chambre |
| B | TypeChambre | Texte | Simple/Double/Suite |
| C | TarifNuit | Numérique | Prix par nuit |
| D | Statut | Texte | Libre/Occupée/Maintenance |
| E | Description | Texte | Description de la chambre |
| F | Equipements | Texte | Liste des équipements |

### 2. Feuille "Clients"
| Colonne | Nom | Type | Description |
|---------|-----|------|-------------|
| A | IDClient | Numérique | Identifiant unique |
| B | Nom | Texte | Nom du client |
| C | Prenom | Texte | Prénom du client |
| D | Telephone | Texte | Numéro de téléphone |
| E | Email | Texte | Adresse email |
| F | Adresse | Texte | Adresse complète |
| G | DateCreation | Date | Date de création du profil |

### 3. Feuille "Reservations"
| Colonne | Nom | Type | Description |
|---------|-----|------|-------------|
| A | IDReservation | Numérique | Identifiant unique |
| B | IDClient | Numérique | Référence client |
| C | NumChambre | Texte | Numéro de chambre |
| D | DateArrivee | Date | Date d'arrivée |
| E | DateDepart | Date | Date de départ |
| F | NbNuits | Numérique | Nombre de nuits |
| G | MontantTotal | Numérique | Montant total |
| H | Statut | Texte | Confirmée/En attente/Annulée |
| I | DateReservation | Date | Date de la réservation |
| J | Commentaires | Texte | Notes particulières |

### 4. Feuille "Paiements"
| Colonne | Nom | Type | Description |
|---------|-----|------|-------------|
| A | IDPaiement | Numérique | Identifiant unique |
| B | IDReservation | Numérique | Référence réservation |
| C | Montant | Numérique | Montant payé |
| D | ModePaiement | Texte | Espèces/CB/Chèque/Virement |
| E | DatePaiement | Date | Date du paiement |
| F | TypePaiement | Texte | Acompte/Solde/Total |
| G | Statut | Texte | Validé/En attente/Refusé |

### 5. Feuille "Parametres"
| Colonne | Nom | Type | Description |
|---------|-----|------|-------------|
| A | Parametre | Texte | Nom du paramètre |
| B | Valeur | Texte | Valeur du paramètre |
| C | Description | Texte | Description |

### 6. Feuille "Dashboard"
Interface principale avec boutons et résumés

### 7. Feuille "Rapports"
Zone de génération des rapports dynamiques
