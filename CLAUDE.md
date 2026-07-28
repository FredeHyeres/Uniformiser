# Uniformiser - MicroStation V8i VBA

## Encodage obligatoire des fichiers VBA

Apres chaque creation ou modification de fichier `.cls`, `.bas` ou `.frm`,
normaliser en **CRLF + ANSI (Windows-1252)** avec :

    $t = [IO.File]::ReadAllText($path)
    $t = $t.Replace("`r`n", "`n").Replace("`n", "`r`n")
    [IO.File]::WriteAllText($path, $t, [Text.Encoding]::GetEncoding(1252))

Raison : le Write tool produit du LF/UTF-8. MicroStation importe les `.cls`
en LF comme des modules standard au lieu de classes, ce qui casse
`Implements` et `New`.

## Formulaire (FRM/FRX)

- Le formulaire est construit entierement au runtime dans
  `ConstruireControles` : le `.frx` est un blob binaire du formulaire vide.
- Si seul le **code** du `.frm` change : le `.frx` reste valide.
- Si les **proprietes du designer** changent (en-tete `Begin...End`) :
  regenerer le couple via `scripts\export_frm_via_excel.ps1`.

## Architecture

- Archetype "outil batch pilote par formulaire" (pas de machine a etats
  IPrimitiveCommandEvents) : tout part des boutons du formulaire modeless.
- Un metier = une classe (CMoteurUniformisation, CItemUniformisation).
- CSettings : parametres sans dependance API MicroStation.
- Le module .bas ne contient que le point d'entree et les globals.

```
Uniformiser()                            ' .bas : entree + globals
      |
      +--- CSettings (g_oSettings)       ' source retenue, categories cochees
      |
      +--- CMoteurUniformisation         ' resolution source + analyse +
      |    (g_oMoteur)                   ' application (API MST)
      |         +-- CItemUniformisation  ' 1 ligne de previsualisation
      |             (une par element)    ' (categorie, nom, etat, inclus)
      |
      +--- MFileDialog.bas               ' boite "Parcourir" (Win32, zero API MST)
      |
      +--- frmUniformiser                ' modeless, controles au runtime
```

## Perimetre v1 (limitations connues)

- **Styles de ligne personnalises** : EXCLUS. Aucune API VBA documentee ne
  permet de copier une definition de style de ligne (stroke/point) entre
  fichiers (`LineStyles` est une collection lecture seule, pas de
  `Add`/`Clone`). Voir KB `12_TODO.md`.
- **Styles de cotation (DimensionStyles)** : EXCLUS, meme constat (pas de
  methode `Add` documentee).
- **Source** : reference attachee (liste deroulante) OU fichier DGN
  quelconque parcouru (`Application.OpenDesignFileForProgram`, sans toucher
  au fichier actif de l'utilisateur).
- **Conflits** (element de meme nom deja present avec des proprietes
  differentes) : jamais ecrases automatiquement, toujours presentes dans la
  previsualisation avec une case a cocher individuelle.

## Lancement

Key-in MicroStation : `vba run [UniformiserV1]Uniformiser`

## Fichiers

```
src/
  UniformiserV1.bas          - Module principal (entree + globals)
  CSettings.cls               - Parametres (source, categories cochees)
  CItemUniformisation.cls     - 1 ligne de previsualisation
  CMoteurUniformisation.cls   - Resolution source + analyse + application
  MFileDialog.bas             - Boite de dialogue "Parcourir" (Win32 pur)
  frmUniformiser.frm          - Formulaire (code)
  frmUniformiser.frx          - Blob designer (binaire)
scripts/
  export_frm_via_excel.ps1    - Generation du couple .frm/.frx
```
