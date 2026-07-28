# Uniformiser

Outil VBA pour MicroStation V8i SS3 : recopie niveaux, styles de texte et
palette de couleurs depuis une source (reference attachee ou fichier DGN
quelconque) vers le fichier actif, avec previsualisation et selection ligne
par ligne avant application.

## Fonctionnement

1. Choisir une source : une reference deja attachee au modele actif, ou un
   fichier DGN parcouru sur disque (sans toucher au fichier actif tant que
   rien n'est confirme).
2. Cocher les categories a recuperer (Niveaux, Styles de texte, Palette de
   couleurs).
3. **Previsualiser** : analyse la source, affiche chaque element candidat
   avec son etat (Nouveau / Identique / Conflit) et une case a cocher
   individuelle.
4. Decocher au besoin les elements a exclure, notamment les conflits (nom
   deja present dans le fichier actif avec des proprietes differentes) qui
   ne sont jamais ecrases automatiquement.
5. **Confirmer** : applique les elements coches au fichier actif.

## Perimetre v1 (limitations connues)

- **Styles de ligne personnalises** : exclus. Aucune API VBA documentee ne
  permet de copier une definition de style de ligne (stroke/point) entre
  fichiers (`LineStyles` est une collection lecture seule).
- **Styles de cotation (DimensionStyles)** : exclus, meme constat.
- **Conflits** : jamais ecrases automatiquement, toujours presentes dans la
  previsualisation pour arbitrage manuel.

## Installation

1. Copier le dossier du projet sur le poste cible.
2. Double-cliquer `install.cmd` (lance `install.ps1`) : copie le projet VBA
   dans le workspace MicroStation et configure le chargement automatique.
   Le script est relancable sans risque.
3. (Re)demarrer MicroStation.

Si `Uniformiser.mvba` n'est pas fourni pre-compile, importer les fichiers de
`src/` dans un nouveau projet VBA MicroStation (Macro > Project Manager),
puis l'enregistrer sous ce nom.

## Utilisation

Key-in MicroStation :

```
vba run [UniformiserV1]Uniformiser
```

Pour un lancement rapide, affecter une touche de fonction (Utilitaires >
Touches de fonction).

## Architecture

Archetype "outil batch pilote par formulaire" : pas de machine a etats
`IPrimitiveCommandEvents`, tout part des boutons du formulaire modeless.
Un metier = une classe.

```
Uniformiser()                            ' .bas : entree + globals
      |
      +--- CSettings (g_oSettings)       ' source retenue, categories cochees
      |
      +--- CMoteurUniformisation         ' resolution source + analyse +
      |    (g_oMoteur)                   ' application (API MicroStation)
      |         +-- CItemUniformisation  ' 1 ligne de previsualisation
      |             (une par element)    ' (categorie, nom, etat, inclus)
      |
      +--- MFileDialog.bas               ' boite "Parcourir" (Win32, zero API MicroStation)
      |
      +--- frmUniformiser                ' modeless, controles construits au runtime
```

## Fichiers

```
src/
  UniformiserV1.bas          - Module principal (entree + globals)
  CSettings.cls               - Parametres (source, categories cochees)
  CItemUniformisation.cls     - 1 ligne de previsualisation
  CMoteurUniformisation.cls   - Resolution source + analyse + application
  MFileDialog.bas             - Boite de dialogue "Parcourir" (Win32 pur)
  frmUniformiser.frm          - Formulaire (code)
  frmUniformiser.frx          - Blob designer (binaire, formulaire vide au repos)
scripts/
  export_frm_via_excel.ps1    - Regeneration du couple .frm/.frx si les
                                 proprietes du designer changent
install.cmd / install.ps1     - Installation automatique dans le workspace
                                 MicroStation
```

## Prerequis

- MicroStation V8i SS3, VBA active.
- Fichier `Default.mvba` non modifie : Uniformiser vit dans son propre
  projet `Uniformiser.mvba` pour ne jamais entrer en conflit.
