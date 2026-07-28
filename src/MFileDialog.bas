Attribute VB_Name = "MFileDialog"
Option Explicit

' =============================================================================
' MFileDialog - Boite de dialogue Windows standard "Ouvrir un fichier"
' API Win32 pure (comdlg32.dll), aucune dependance a l'API MicroStation
' =============================================================================

Private Type OPENFILENAME
    lStructSize       As Long
    hwndOwner         As Long
    hInstance         As Long
    lpstrFilter       As String
    lpstrCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As String
    nMaxFile          As Long
    lpstrFileTitle    As String
    nMaxFileTitle     As Long
    lpstrInitialDir   As String
    lpstrTitle        As String
    flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_HIDEREADONLY  As Long = &H4

' Retourne le chemin complet choisi, ou "" si l'utilisateur annule.
Public Function AfficherOuvrirFichier(Optional ByVal sTitre As String = "Choisir un fichier DGN") As String
    Dim ofn As OPENFILENAME
    Dim sFiltre As String
    Dim sFichier As String

    sFiltre = "Fichiers DGN (*.dgn)" & Chr$(0) & "*.dgn" & Chr$(0) & _
              "Tous les fichiers (*.*)" & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)

    sFichier = String$(260, 0)

    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = 0
        .lpstrFilter = sFiltre
        .nFilterIndex = 1
        .lpstrFile = sFichier
        .nMaxFile = Len(sFichier)
        .lpstrFileTitle = String$(260, 0)
        .nMaxFileTitle = 260
        .lpstrTitle = sTitre
        .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY
    End With

    If GetOpenFileName(ofn) <> 0 Then
        AfficherOuvrirFichier = Left$(ofn.lpstrFile, InStr(ofn.lpstrFile, Chr$(0)) - 1)
    Else
        AfficherOuvrirFichier = ""
    End If
End Function
