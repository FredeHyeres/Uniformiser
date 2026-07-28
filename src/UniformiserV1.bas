Attribute VB_Name = "UniformiserV1"
Option Explicit

' =============================================================================
' UniformiserV1 - Uniformisation de style DGN (niveaux, styles de texte, palette)
' MicroStation V8i SS3 - VBA 6.x
' =============================================================================

Public Enum EEtatUniformisation
    eUniNouveau = 0
    eUniIdentique = 1
    eUniConflit = 2
End Enum

' --- Globals partages ---
Public g_oSettings      As CSettings
Public g_oMoteur        As CMoteurUniformisation

'----------------------------------------------------------------------
Sub Uniformiser()
    Dim oDgn As DesignFile
    On Error Resume Next
    Set oDgn = ActiveDesignFile
    On Error GoTo 0
    If oDgn Is Nothing Then
        MsgBox "Ouvrez d'abord un fichier DGN.", vbExclamation, "Uniformiser"
        Exit Sub
    End If

    If g_oSettings Is Nothing Then
        Set g_oSettings = New CSettings
        g_oSettings.Init
    End If
    Set g_oMoteur = New CMoteurUniformisation

    frmUniformiser.Initialiser g_oSettings
    frmUniformiser.StartUpPosition = 0
    frmUniformiser.Left = Application.Width * 0.4
    frmUniformiser.Top = Application.Height * 0.05
    frmUniformiser.Show vbModeless
End Sub

'----------------------------------------------------------------------
' Appelee par le bouton "Previsualiser"
Sub LancerAnalyse()
    If g_oMoteur Is Nothing Then Exit Sub

    ' --- DIAGNOSTIC TEMPORAIRE ------------------------------------------
    ' HasActiveModelReference = True mais ActiveModelReference plante quand
    ' meme : on isole ici precisement laquelle des deux expressions echoue,
    ' avec le vrai message d'erreur (pas juste le code OLE generique).
    ' A retirer une fois la cause identifiee.
    Dim sDiag As String
    Dim bHasModel As Boolean
    Dim oDF As DesignFile
    Dim oMR As ModelReference

    On Error Resume Next
    Err.Clear
    bHasModel = Application.HasActiveModelReference
    sDiag = sDiag & "HasActiveModelReference = " & bHasModel & vbCrLf & _
                     "  Err " & Err.Number & " : " & Err.Description & vbCrLf

    Err.Clear
    Set oDF = ActiveDesignFile
    sDiag = sDiag & "ActiveDesignFile : Nothing=" & (oDF Is Nothing) & vbCrLf & _
                     "  Err " & Err.Number & " : " & Err.Description & vbCrLf

    Err.Clear
    Set oMR = ActiveModelReference
    sDiag = sDiag & "ActiveModelReference : Nothing=" & (oMR Is Nothing) & vbCrLf & _
                     "  Err " & Err.Number & " : " & Err.Description
    On Error GoTo 0

    If Not bHasModel Or oDF Is Nothing Or oMR Is Nothing Then
        MsgBox sDiag, vbExclamation, "Uniformiser - contexte invalide"
        Exit Sub
    End If

    ' Etape 2 : appel reel, mais avec les variables locales DEJA capturees et
    ' validees ci-dessus (pas de reevaluation inline de ActiveModelReference/
    ' ActiveDesignFile sur la ligne d'appel). On capture aussi le vrai
    ' Err.Description au lieu de se fier au texte de la boite de dialogue.
    On Error Resume Next
    Err.Clear
    g_oMoteur.Analyser g_oSettings, oMR, oDF
    sDiag = sDiag & vbCrLf & vbCrLf & "Analyser() : Err " & Err.Number & " : " & Err.Description
    On Error GoTo 0

    MsgBox sDiag, vbInformation, "Uniformiser - diagnostic"
    ' --- FIN DIAGNOSTIC TEMPORAIRE ---------------------------------------

    frmUniformiser.AfficherResultat g_oMoteur
End Sub

'----------------------------------------------------------------------
' Appelee par le bouton "Confirmer" de la previsualisation
Sub LancerApplication()
    If g_oMoteur Is Nothing Then Exit Sub
    g_oMoteur.Appliquer ActiveDesignFile
    frmUniformiser.AfficherApplication
End Sub

'----------------------------------------------------------------------
Sub NettoyerAvantFermeture()
    If g_oMoteur Is Nothing Then Exit Sub
    g_oMoteur.FermerSource
End Sub
