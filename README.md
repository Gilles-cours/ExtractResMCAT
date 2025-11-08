# Extracteur de données Excel MCAT

Script JavaScript pour extraire les valeurs "EquiDose" et "Incertitude" des colonnes F5, F15, F25, F35 d'un fichier Excel MCAT.

## Installation

1. Installer les dépendances :
```bash
npm install
```

## Utilisation

### Méthode 1 : Utiliser le fichier par défaut

Lancez le script et appuyez sur Entrée pour utiliser le fichier `report_standard_S112Part1.xlsx` :

```bash
npm start
```

ou

```bash
node extractExcel.js
```

### Méthode 2 : Spécifier un fichier

Lancez le script et entrez le chemin vers votre fichier Excel quand demandé :

```bash
npm start
```

Puis entrez le chemin du fichier, par exemple :
```
/chemin/vers/votre/fichier.xlsx
```

## Format de sortie

Le script génère un fichier `resultats_extraction.txt` contenant :

1. **Détails par ligne** : Pour chaque ligne du fichier Excel, affiche :
   - Valeur de F5 : EquiDose et Incertitude
   - Valeur de F15 : EquiDose et Incertitude
   - Valeur de F25 : EquiDose et Incertitude
   - Valeur de F35 : EquiDose et Incertitude

2. **Tableau récapitulatif** : Un tableau structuré avec toutes les valeurs organisées

## Exemple de sortie

```
────────────────────────────────────────
Ligne 2 - IRN: S112Part1
────────────────────────────────────────

Valeur de F5:
  EquiDose:     2.983e-9
  Incertitude:  0.0361

Valeur de F15:
  EquiDose:     1.318e-9
  Incertitude:  0.0078

...

═══════════════════════════════════════════════════════════════
TABLEAU RÉCAPITULATIF
═══════════════════════════════════════════════════════════════

IRN             | Valeur   | F5              | F15             | F25             | F35
────────────────┼──────────┼─────────────────┼─────────────────┼─────────────────┼────────────────
S112Part1       | EquiDose | 2.983e-9        | 1.318e-9        | 5.718e-10       | 2.338e-10
                | Incert.  | 0.0361          | 0.0078          | 0.0051          | 0.0041
...
```

## Structure du fichier Excel attendue

Le script attend un fichier Excel avec :
- Une première ligne contenant les en-têtes
- Des colonnes nommées :
  - `F5 Gamma EquiDose` et `F5 Gamma Uncertainty`
  - `F15 Gamma EquiDose` et `F15 Gamma Uncertainty`
  - `F25 Gamma EquiDose` et `F25 Gamma Uncertainty`
  - `F35 Gamma EquiDose` et `F35 Gamma Uncertainty`

## Fichiers

- `extractExcel.js` : Script principal
- `package.json` : Configuration et dépendances
- `resultats_extraction.txt` : Fichier de sortie généré
- `report_standard_S112Part1.xlsx` : Fichier Excel exemple

## Dépendances

- `xlsx` (SheetJS) : Librairie pour lire les fichiers Excel

## Licence

Libre d'utilisation
