VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' CSettings - Parametres d'Uniformiser
' Aucune dependance a l'API MicroStation
' =============================================================================

Public bSourceFichier      As Boolean   ' False = reference attachee, True = fichier parcouru
Public nIndexAttachement   As Long      ' index 1-based dans Attachments, 0 = aucun
Public sCheminFichier      As String    ' chemin du fichier parcouru

Public bInclureNiveaux     As Boolean
Public bInclureStylesTexte As Boolean
Public bInclurePalette     As Boolean

Sub Init()
    bSourceFichier = False
    nIndexAttachement = 0
    sCheminFichier = ""

    bInclureNiveaux = True
    bInclureStylesTexte = True
    bInclurePalette = True
End Sub
