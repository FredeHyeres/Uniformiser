VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CItemUniformisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' CItemUniformisation - Une ligne de la previsualisation
' (1 niveau, 1 style de texte ou la palette de couleurs)
' =============================================================================

Public Categorie      As String                  ' "Niveaux" / "Styles de texte" / "Palette"
Public Nom            As String
Public Etat           As EEtatUniformisation      ' eUniNouveau / eUniIdentique / eUniConflit
Public Inclus         As Boolean
Public DetailConflit  As String
Public ObjetSource    As Object                   ' Level / TextStyle / ColorTable source
