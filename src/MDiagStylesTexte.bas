Attribute VB_Name = "MDiagStylesTexte"
Option Explicit

' =============================================================================
' MDiagStylesTexte - Diagnostic temporaire : dimensions des styles de texte
'
' Ecrit <dossier du fichier actif>\diag_styles_texte.txt avec, pour le
' fichier ACTIF et le fichier SOURCE deja resolu par Uniformiser :
'   - SubUnitsPerMasterUnit, UORsPerSubUnit, UORsPerMasterUnit,
'     MasterUnit/SubUnit (Label, Numerateur, Denominateur, System, Base)
' puis pour chaque style de texte source :
'   - Height/Width bruts (TextStyle lu via le DesignFile source)
'   - Height/Width "corrige" = brut / SubUnitsPerMasterUnit du modele source
'     (correctif actuel, cf KB 04_Geometrie.md #6.1b)
'   - si un style de meme nom existe deja dans le fichier actif : ses
'     Height/Width (valeur fiable) et le ratio empirique Actif/SourceBrut,
'     a comparer au ratio theorique 1/SubUnitsPerMasterUnit pour verifier
'     si le correctif s'applique ou non sur CE fichier source.
'
' A lancer APRES "Previsualiser" dans Uniformiser (la source doit etre
' resolue par g_oMoteur.Analyser). Key-in : vba run [UniformiserV1]DiagStylesTexte
' =============================================================================

Sub DiagStylesTexte()
    Dim dfActif As DesignFile
    Dim oModelActif As ModelReference
    Dim dfSource As DesignFile
    Dim oModelSrc As ModelReference
    Dim sOut As String
    Dim sPath As String
    Dim nFile As Integer

    On Error Resume Next
    Set dfActif = ActiveDesignFile
    Set oModelActif = ActiveModelReference
    On Error GoTo 0

    If dfActif Is Nothing Then
        MsgBox "Ouvrez un fichier DGN.", vbExclamation, "Diag styles de texte"
        Exit Sub
    End If

    If g_oMoteur Is Nothing Then
        MsgBox "Lancez d'abord 'Previsualiser' dans Uniformiser (la source doit etre resolue).", vbExclamation, "Diag styles de texte"
        Exit Sub
    End If

    Set dfSource = g_oMoteur.SourceDesignFile
    If dfSource Is Nothing Then
        MsgBox "Source non resolue. Lancez d'abord 'Previsualiser' dans Uniformiser.", vbExclamation, "Diag styles de texte"
        Exit Sub
    End If

    On Error Resume Next
    Set oModelSrc = dfSource.DefaultModelReference
    On Error GoTo 0

    sOut = "=== Diag styles de texte - " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ===" & vbCrLf & vbCrLf

    sOut = sOut & "--- Fichier ACTIF ---" & vbCrLf
    sOut = sOut & LigneUnitesFichier(dfActif, oModelActif)
    sOut = sOut & vbCrLf

    sOut = sOut & "--- Fichier SOURCE ---" & vbCrLf
    If Not g_oSettings Is Nothing Then
        If g_oSettings.bSourceFichier Then
            sOut = sOut & "  Origine = fichier parcouru (" & g_oSettings.sCheminFichier & ")" & vbCrLf
        Else
            sOut = sOut & "  Origine = reference attachee #" & g_oSettings.nIndexAttachement & vbCrLf
        End If
    End If
    sOut = sOut & LigneUnitesFichier(dfSource, oModelSrc)
    sOut = sOut & vbCrLf

    sOut = sOut & "--- Styles de texte (brut source vs corrige vs actif) ---" & vbCrLf
    sOut = sOut & LignesStylesTexte(dfSource, dfActif, oModelSrc)

    sPath = DossierDe(dfActif.FullName) & "diag_styles_texte.txt"
    nFile = FreeFile
    Open sPath For Output As #nFile
    Print #nFile, sOut
    Close #nFile

    MsgBox "Diagnostic ecrit dans :" & vbCrLf & sPath, vbInformation, "Diag styles de texte"
End Sub

Private Function DossierDe(ByVal sCheminFichier As String) As String
    DossierDe = Left$(sCheminFichier, InStrRev(sCheminFichier, "\"))
End Function

Private Function LigneUnitesFichier(ByVal df As DesignFile, ByVal oModel As ModelReference) As String
    Dim s As String

    On Error Resume Next
    s = s & "  Fichier               = " & df.FullName & vbCrLf
    s = s & "  SubUnitsPerMasterUnit = " & oModel.SubUnitsPerMasterUnit & vbCrLf
    s = s & "  UORsPerSubUnit        = " & oModel.UORsPerSubUnit & vbCrLf
    s = s & "  UORsPerMasterUnit     = " & oModel.UORsPerMasterUnit & vbCrLf
    s = s & "  MasterUnit.Label      = " & oModel.MasterUnit.Label & vbCrLf
    s = s & "  MasterUnit.Num/Denom  = " & oModel.MasterUnit.UnitsPerBaseNumerator & " / " & oModel.MasterUnit.UnitsPerBaseDenominator & vbCrLf
    s = s & "  SubUnit.Label         = " & oModel.SubUnit.Label & vbCrLf
    s = s & "  SubUnit.Num/Denom     = " & oModel.SubUnit.UnitsPerBaseNumerator & " / " & oModel.SubUnit.UnitsPerBaseDenominator & vbCrLf
    On Error GoTo 0

    LigneUnitesFichier = s
End Function

Private Function LignesStylesTexte(ByVal dfSource As DesignFile, ByVal dfActif As DesignFile, ByVal oModelSrc As ModelReference) As String
    Dim s As String
    Dim oStySrc As TextStyle
    Dim oStyDst As TextStyle
    Dim dSubPerMu As Double
    Dim dHautCorrigee As Double
    Dim dLargCorrigee As Double

    On Error Resume Next
    dSubPerMu = oModelSrc.SubUnitsPerMasterUnit
    On Error GoTo 0

    For Each oStySrc In dfSource.TextStyles
        s = s & "  [" & oStySrc.Name & "]" & vbCrLf

        On Error Resume Next
        s = s & "    Source brut    : Height=" & Format$(oStySrc.Height, "0.000000") & "  Width=" & Format$(oStySrc.Width, "0.000000") & vbCrLf
        On Error GoTo 0

        If dSubPerMu <> 0 Then
            On Error Resume Next
            dHautCorrigee = oStySrc.Height / dSubPerMu
            dLargCorrigee = oStySrc.Width / dSubPerMu
            On Error GoTo 0
            s = s & "    Corrige actuel : Height=" & Format$(dHautCorrigee, "0.000000") & "  Width=" & Format$(dLargCorrigee, "0.000000") & "  (brut / SubUnitsPerMasterUnit=" & dSubPerMu & ")" & vbCrLf
        Else
            s = s & "    Corrige actuel : (SubUnitsPerMasterUnit indisponible, pas de correction appliquee)" & vbCrLf
        End If

        Set oStyDst = Nothing
        On Error Resume Next
        Set oStyDst = dfActif.TextStyles(oStySrc.Name)
        On Error GoTo 0

        If Not oStyDst Is Nothing Then
            On Error Resume Next
            s = s & "    Actif (fiable) : Height=" & Format$(oStyDst.Height, "0.000000") & "  Width=" & Format$(oStyDst.Width, "0.000000") & vbCrLf
            If oStySrc.Height <> 0 Then
                s = s & "    Ratio empirique Actif/SourceBrut (Height) = " & Format$(oStyDst.Height / oStySrc.Height, "0.000000") & vbCrLf
            End If
            On Error GoTo 0
        Else
            s = s & "    Actif           : (pas de style de meme nom pour comparaison)" & vbCrLf
        End If

        s = s & vbCrLf
    Next oStySrc

    LignesStylesTexte = s
End Function
