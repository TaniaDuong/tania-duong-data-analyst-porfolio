# Automatisation du process de reporting

## Contexte
Le service **Financial Performance Control** d’une banque d’investissement à Paris a pour mission de fournir à la **Direction Financière** des éléments financiers fiables pour piloter les activités des différentes **Business Lines (BL)**.  
Chaque mois, les contrôleurs de gestion de chaque BL suivent les mouvements de leurs portefeuilles et transmettent des fichiers Excel au service transverse pour consolidation.  
Le processus actuel est **manuel, lourd, non homogène et sujet à erreurs**, ce qui justifie l’automatisation.

## Objectif du projet
L’objectif de ce projet est de **fiabiliser et automatiser le processus de consolidation et de reporting financier**, afin de :  
- Réduire les tâches manuelles et les risques d’erreurs,  
- Fournir des indicateurs financiers homogènes et fiables,  
- Générer automatiquement un reporting prêt à l’usage pour le management.

## Processus
### Ancien processus
- Collecte manuelle des fichiers Excel des BL  
- Consolidation Excel manuelle (copier-coller, formules)  
- Calcul des indicateurs ligne par ligne  
- Contrôles partiels et manuels  
- Reporting final exposé aux erreurs et chronophage  

### Nouveau processus
- Import automatisé des fichiers Excel dans une **base Access** via **VBA**  
- Consolidation centralisée et contrôles de cohérence automatiques  
- Calcul standardisé des indicateurs clés : NBI, Charges, Résultat, ROI, CIR, Productivité  
- Génération automatique d’un reporting Excel avec tableaux et graphiques  
- Comparaison avec M-1 et N-1 pour suivre l’évolution  

## Technologies utilisées
- **Excel VBA** : automatisation, calculs, génération de reporting et graphiques  
- **Access** : centralisation et fiabilité des données  
- **CSV** : format standard pour l’import des fichiers Business Lines  

## Données (dataset)
- `NBI_BLX.csv` : Revenus par produit, desk et région  
- `Payroll_BLX.csv` : Masse salariale et effectifs par département et rôle  
- `FraisGeneraux_BLX.csv` : Charges opérationnelles par type et centre de coût

👉 Certains fichiers de test contiennent volontairement des erreurs de structure ou de contenu (colonnes manquantes, valeurs incohérentes, périodes invalides) afin d’illustrer le fonctionnement des **contrôles automatiques** et la génération du rapport d’erreurs.

## Solution technique

### 1. Standardisation et contrôle des fichiers Business Line

Afin de fiabiliser la collecte des données et d’éviter les différences de structure entre les fichiers envoyés par les Business Lines, un **fichier Excel standardisé (BL_InputTemplate.xlsm)** a été créé. Une **feuille d’accueil** permet de sélectionner la Business Line et la période, et un bouton déclenche la macro qui **nettoie automatiquement** des feuilles export, applique **tous les contrôles de structure et de contenu** (validation des produits, catégories, devises, chiffres non nuls, formats corrects) et **génère** les fichiers CSV conformes au format attendu pour l’import dans Access. 


## Instructions d’utilisation
1. Placer tous les fichiers CSV dans le dossier `data/`  
2. Ouvrir `ReportingTemplate.xlsm` dans Excel  
3. Activer les macros et lancer la macro principale  
4. Les données sont automatiquement importées dans Access et consolidées  
5. Le reporting Excel est généré avec tableaux et graphiques mis à jour  

## Résultats attendus
- Reporting consolidé fiable et rapide  
- Visualisation des indicateurs clés (NBI, Charges, Résultat, ROI, CIR)  
- Comparaison automatique avec le mois précédent et l’année précédente  
- Gain de temps significatif pour le service transverse  

## Captures d’écran
- Diagramme du processus actuel : `images/Process_Ancien.png`  
- Diagramme du nouveau processus : `images/Process_Nouveau.png`  
- Exemple de reporting Excel : `images/Graphiques_Exemple.png`  

## Auteurs
- **Tania DUONG**


