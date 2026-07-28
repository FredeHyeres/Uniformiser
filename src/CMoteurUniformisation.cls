VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMoteurUniformisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' CMoteurUniformisation - Resolution source + analyse + application
' Toute l'API MicroStation du projet vit ici ("un metier = une classe")
' =============================================================================

Private m_dfSource  As DesignFile
Private m_colItems  As Collection

' =============================================================================
' Resolution de la source
' =============================================================================

' Iteration par INDEX protegee (KB 07_References.md #8.1 : For Each fragile).
' Chaque position renvoyee correspond exactement a Attachments(i), meme si un
' attachement est illisible (libelle de repli), pour garder l'alignement avec
' ObtenirDesignFileSourceParAttachement.
Public Function ListerAttachments(ByVal oModelActif As ModelReference) As Collection
    Dim colLibelles As New Collection
    Dim oAtts As Object
    Dim nCount As Long
    Dim i As Long
    Dim oAtt As Object
    Dim sLabel As String

    On Error Resume Next
    Set oAtts = oModelActif.Attachments
    On Error GoTo 0

    On Error Resume Next
    nCount = oAtts.Count
    On Error GoTo 0

    For i = 1 To nCount
        Set oAtt = Nothing
        On Error Resume Next
        Set oAtt = oAtts(i)
        On Error GoTo 0

        sLabel = ""
        If Not oAtt Is Nothing Then
            On Error Resume Next
            sLabel = oAtt.LogicalName
            If Len(sLabel) = 0 Then sLabel = oAtt.AttachName
            On Error GoTo 0
        End If
        If Len(sLabel) = 0 Then sLabel = "(reference " & i & ")"

        colLibelles.Add sLabel
    Next i

    Set ListerAttachments = colLibelles
End Function

Private Function ObtenirDesignFileSourceParAttachement(ByVal oModelActif As ModelReference, ByVal nIndex As Long) As DesignFile
    Dim oAtts As Object
    Dim oAtt As Object
    Dim dfRes As DesignFile

    On Error Resume Next
    Set oAtts = oModelActif.Attachments
    Set oAtt = oAtts(nIndex)
    Set dfRes = oAtt.DesignFile
    On Error GoTo 0

    Set ObtenirDesignFileSourceParAttachement = dfRes
End Function

Private Function ObtenirDesignFileSourceParFichier(ByVal sChemin As String) As DesignFile
    Dim dfRes As DesignFile

    If Len(sChemin) = 0 Then Exit Function

    On Error Resume Next
    Set dfRes = Application.OpenDesignFileForProgram(sChemin, True)
    On Error GoTo 0

    Set ObtenirDesignFileSourceParFichier = dfRes
End Function

' =============================================================================
' Orchestration : analyse
' =============================================================================

Public Sub Analyser(ByVal oSettings As CSettings, ByVal oModelActif As ModelReference, ByVal dfActif As DesignFile)
    Dim dfSource As DesignFile

    Set m_colItems = New Collection
    Set m_dfSource = Nothing

    If oSettings.bSourceFichier Then
        Set dfSource = ObtenirDesignFileSourceParFichier(oSettings.sCheminFichier)
    Else
        If oSettings.nIndexAttachement > 0 Then
            Set dfSource = ObtenirDesignFileSourceParAttachement(oModelActif, oSettings.nIndexAttachement)
        End If
    End If

    If dfSource Is Nothing Then Exit Sub
    Set m_dfSource = dfSource

    If oSettings.bInclureNiveaux Then AjouterTous AnalyserNiveaux(dfSource, dfActif)
    If oSettings.bInclureStylesTexte Then AjouterTous AnalyserStylesTexte(dfSource, dfActif)
    If oSettings.bInclurePalette Then AjouterTous AnalyserPalette(dfSource, dfActif)
End Sub

Private Sub AjouterTous(ByVal colSrc As Collection)
    Dim oItem As CItemUniformisation
    For Each oItem In colSrc
        m_colItems.Add oItem
    Next oItem
End Sub

' --- Niveaux ---------------------------------------------------------------

Private Function AnalyserNiveaux(ByVal dfSource As DesignFile, ByVal dfActif As DesignFile) As Collection
    Dim colRes As New Collection
    Dim oModelSrc As ModelReference
    Dim oLvlSrc As Level
    Dim oLvlDst As Level
    Dim oItem As CItemUniformisation
    Dim sDetail As String

    ' [OK confirme 2026-07-28] DesignFile.Levels s'appuie sur le "default
    ' model reference" du fichier (CHM, doc Levels Property) : pas fiable
    ' pour un DesignFile obtenu via OpenDesignFileForProgram ou
    ' Attachment.DesignFile -> erreur COM "Bad modelRef" (-2147467259).
    ' Passer explicitement par DefaultModelReference.Levels le contourne.
    Set oModelSrc = dfSource.DefaultModelReference

    For Each oLvlSrc In oModelSrc.Levels
        Set oLvlDst = TrouverNiveauCorrespondant(dfActif, oLvlSrc)

        Set oItem = New CItemUniformisation
        oItem.Categorie = "Niveaux"
        oItem.Nom = LibelleNiveau(oLvlSrc)
        Set oItem.ObjetSource = oLvlSrc

        If oLvlDst Is Nothing Then
            oItem.Etat = eUniNouveau
            oItem.Inclus = True
        Else
            sDetail = ComparerNiveaux(oLvlSrc, oLvlDst)
            If Len(sDetail) = 0 Then
                oItem.Etat = eUniIdentique
                oItem.Inclus = False
            Else
                oItem.Etat = eUniConflit
                oItem.DetailConflit = sDetail
                oItem.Inclus = True
            End If
        End If

        colRes.Add oItem
    Next oLvlSrc

    Set AnalyserNiveaux = colRes
End Function

Private Function ComparerNiveaux(ByVal oSrc As Level, ByVal oDst As Level) As String
    Dim sDiffs As String
    Dim sStyleSrc As String
    Dim sStyleDst As String

    If oSrc.ElementColor <> oDst.ElementColor Then
        sDiffs = sDiffs & "Couleur " & oDst.ElementColor & " -> " & oSrc.ElementColor & "; "
    End If

    If oSrc.ElementLineWeight <> oDst.ElementLineWeight Then
        sDiffs = sDiffs & "Epaisseur " & oDst.ElementLineWeight & " -> " & oSrc.ElementLineWeight & "; "
    End If

    sStyleSrc = "": sStyleDst = ""
    On Error Resume Next
    sStyleSrc = oSrc.ElementLineStyle.Name
    sStyleDst = oDst.ElementLineStyle.Name
    On Error GoTo 0
    If sStyleSrc <> sStyleDst Then
        sDiffs = sDiffs & "Style " & sStyleDst & " -> " & sStyleSrc & "; "
    End If

    If oSrc.Description <> oDst.Description Then
        sDiffs = sDiffs & "Description modifiee; "
    End If

    ComparerNiveaux = sDiffs
End Function

' [OK confirme] Les 63 niveaux "legacy" numerotes d'un DGN V8 existent par
' defaut dans tout fichier et n'ont le plus souvent pas de Name (chaine
' vide) : seul leur Number les identifie de facon fiable. Chercher par Name
' seul renvoie systematiquement Nothing pour ces niveaux (tout apparait en
' "Nouveau" avec un nom vide) -- d'ou le repli sur FindByNumber.
Private Function LibelleNiveau(ByVal oLvl As Level) As String
    If Len(oLvl.Name) > 0 Then
        LibelleNiveau = oLvl.Name
    Else
        LibelleNiveau = "(niveau " & oLvl.Number & ")"
    End If
End Function

Private Function TrouverNiveauCorrespondant(ByVal dfActif As DesignFile, ByVal oLvlSrc As Level) As Level
    Dim oRes As Level

    On Error Resume Next
    If Len(oLvlSrc.Name) > 0 Then
        Set oRes = dfActif.Levels(oLvlSrc.Name)
    End If
    If oRes Is Nothing Then
        Set oRes = dfActif.Levels.FindByNumber(oLvlSrc.Number)
    End If
    On Error GoTo 0

    Set TrouverNiveauCorrespondant = oRes
End Function

' --- Styles de texte --------------------------------------------------------

Private Function AnalyserStylesTexte(ByVal dfSource As DesignFile, ByVal dfActif As DesignFile) As Collection
    Dim colRes As New Collection
    Dim oModelSrc As ModelReference
    Dim oStySrc As TextStyle
    Dim oStyDst As TextStyle
    Dim oItem As CItemUniformisation
    Dim sDetail As String

    Set oModelSrc = dfSource.DefaultModelReference

    For Each oStySrc In dfSource.TextStyles
        Set oStyDst = Nothing
        On Error Resume Next
        Set oStyDst = dfActif.TextStyles(oStySrc.Name)
        On Error GoTo 0

        Set oItem = New CItemUniformisation
        oItem.Categorie = "Styles de texte"
        oItem.Nom = oStySrc.Name
        Set oItem.ObjetSource = oStySrc

        If oStyDst Is Nothing Then
            oItem.Etat = eUniNouveau
            oItem.Inclus = True
        Else
            sDetail = ComparerStylesTexte(oStySrc, oStyDst, oModelSrc)
            If Len(sDetail) = 0 Then
                oItem.Etat = eUniIdentique
                oItem.Inclus = False
            Else
                oItem.Etat = eUniConflit
                oItem.DetailConflit = sDetail
                oItem.Inclus = True
            End If
        End If

        colRes.Add oItem
    Next oStySrc

    Set AnalyserStylesTexte = colRes
End Function

Private Function ComparerStylesTexte(ByVal oSrc As TextStyle, ByVal oDst As TextStyle, ByVal oModelSrc As ModelReference) As String
    Dim sDiffs As String
    Dim sFontSrc As String
    Dim sFontDst As String
    Dim dHauteurSrc As Double
    Dim dLargeurSrc As Double

    sFontSrc = "": sFontDst = ""
    On Error Resume Next
    sFontSrc = oSrc.Font.Name
    sFontDst = oDst.Font.Name
    On Error GoTo 0
    If sFontSrc <> sFontDst Then
        sDiffs = sDiffs & "Police " & sFontDst & " -> " & sFontSrc & "; "
    End If

    dHauteurSrc = TailleReelleSource(oSrc.Height, oModelSrc)
    dLargeurSrc = TailleReelleSource(oSrc.Width, oModelSrc)

    If Abs(dHauteurSrc - oDst.Height) > 0.0001 Then
        sDiffs = sDiffs & "Hauteur " & Format$(oDst.Height, "0.####") & " -> " & Format$(dHauteurSrc, "0.####") & "; "
    End If
    If Abs(dLargeurSrc - oDst.Width) > 0.0001 Then
        sDiffs = sDiffs & "Largeur " & Format$(oDst.Width, "0.####") & " -> " & Format$(dLargeurSrc, "0.####") & "; "
    End If
    If oSrc.Justification <> oDst.Justification Then
        sDiffs = sDiffs & "Justification modifiee; "
    End If
    If oSrc.Color <> oDst.Color Then
        sDiffs = sDiffs & "Couleur " & oDst.Color & " -> " & oSrc.Color & "; "
    End If
    If oSrc.IsBold <> oDst.IsBold Then sDiffs = sDiffs & "Gras modifie; "
    If oSrc.IsItalic <> oDst.IsItalic Then sDiffs = sDiffs & "Italique modifie; "
    If oSrc.IsUnderlined <> oDst.IsUnderlined Then sDiffs = sDiffs & "Souligne modifie; "

    ComparerStylesTexte = sDiffs
End Function

' --- Palette de couleurs ----------------------------------------------------

Private Function AnalyserPalette(ByVal dfSource As DesignFile, ByVal dfActif As DesignFile) As Collection
    Dim colRes As New Collection
    Dim ctSrc As ColorTable
    Dim ctDst As ColorTable
    Dim oItem As CItemUniformisation
    Dim aSrc() As Long
    Dim aDst() As Long
    Dim i As Long
    Dim bDiff As Boolean

    Set ctSrc = dfSource.ExtractColorTable
    Set ctDst = dfActif.ExtractColorTable
    aSrc = ctSrc.GetColors
    aDst = ctDst.GetColors

    bDiff = False
    For i = LBound(aSrc) To UBound(aSrc)
        If aSrc(i) <> aDst(i) Then
            bDiff = True
            Exit For
        End If
    Next i

    Set oItem = New CItemUniformisation
    oItem.Categorie = "Palette"
    oItem.Nom = "Palette de couleurs"
    Set oItem.ObjetSource = ctSrc

    If bDiff Then
        oItem.Etat = eUniConflit
        oItem.DetailConflit = "La palette differe du fichier actif"
        oItem.Inclus = True
    Else
        oItem.Etat = eUniIdentique
        oItem.Inclus = False
    End If

    colRes.Add oItem
    Set AnalyserPalette = colRes
End Function

' =============================================================================
' Orchestration : application
' =============================================================================

Public Sub Appliquer(ByVal dfActif As DesignFile)
    If m_colItems Is Nothing Then Exit Sub
    AppliquerNiveaux dfActif
    AppliquerStylesTexte dfActif
    AppliquerPalette dfActif
End Sub

Private Sub AppliquerNiveaux(ByVal dfActif As DesignFile)
    Dim oItem As CItemUniformisation
    Dim oLvlSrc As Level
    Dim oLvlDst As Level
    Dim sNomStyle As String
    Dim bModifie As Boolean

    For Each oItem In m_colItems
        If oItem.Categorie = "Niveaux" And oItem.Inclus Then
            Set oLvlSrc = oItem.ObjetSource
            Set oLvlDst = Nothing

            Select Case oItem.Etat
                Case eUniNouveau
                    ' AddNewLevel est une methode de DesignFile, PAS de la
                    ' collection Levels (confirme CHM : "Applies To" =
                    ' DesignFile Object). dfActif.Levels.AddNewLevel leve
                    ' "membre de methode introuvable".
                    On Error Resume Next
                    Set oLvlDst = dfActif.AddNewLevel(oLvlSrc.Name)
                    On Error GoTo 0
                Case eUniConflit
                    Set oLvlDst = TrouverNiveauCorrespondant(dfActif, oLvlSrc)
            End Select

            If Not oLvlDst Is Nothing Then
                On Error Resume Next
                oLvlDst.ElementColor = oLvlSrc.ElementColor
                oLvlDst.ElementLineWeight = oLvlSrc.ElementLineWeight
                oLvlDst.Description = oLvlSrc.Description

                sNomStyle = ""
                sNomStyle = oLvlSrc.ElementLineStyle.Name
                If Len(sNomStyle) > 0 Then
                    Set oLvlDst.ElementLineStyle = dfActif.LineStyles(sNomStyle)
                End If
                On Error GoTo 0

                bModifie = True
            End If
        End If
    Next oItem

    If bModifie Then dfActif.Levels.Rewrite
End Sub

Private Sub AppliquerStylesTexte(ByVal dfActif As DesignFile)
    Dim oItem As CItemUniformisation
    Dim oStySrc As TextStyle
    Dim oStyDst As TextStyle
    Dim oModelSrc As ModelReference

    If m_dfSource Is Nothing Then Exit Sub
    Set oModelSrc = m_dfSource.DefaultModelReference

    For Each oItem In m_colItems
        If oItem.Categorie = "Styles de texte" And oItem.Inclus Then
            Set oStySrc = oItem.ObjetSource
            Set oStyDst = Nothing

            Select Case oItem.Etat
                Case eUniNouveau
                    On Error Resume Next
                    Set oStyDst = dfActif.TextStyles.Add(Nothing, oStySrc.Name)
                    On Error GoTo 0
                Case eUniConflit
                    On Error Resume Next
                    Set oStyDst = dfActif.TextStyles(oStySrc.Name)
                    On Error GoTo 0
            End Select

            If Not oStyDst Is Nothing Then
                CopierProprietesStyleTexte oStySrc, oStyDst, dfActif, oModelSrc
            End If
        End If
    Next oItem
End Sub

' [OK confirme 2026-07-28] Assignation DIRECTE oDst.Font = oSrc.Font, meme
' motif que InterpolationTopo (CMstSettings.SyncStyleTexteModele, en
' production). L'ancienne approche (re-resoudre la police par NOM dans
' dfActif.Fonts) echouait TOUJOURS silencieusement (Err=0, mais Nothing) :
' dfActif.Fonts ne liste que les polices DEJA utilisees dans ce fichier
' precis, pas toutes les polices systeme -- "Arial" n'y est pas trouvable
' par nom tant qu'aucun element du fichier actif ne l'utilise deja, meme
' si elle existe reellement. L'assignation directe laisse MicroStation
' gerer l'import/la resolution cross-fichier lui-meme (meme esprit que
' CopyElement pour la symbologie).
Private Sub CopierProprietesStyleTexte(ByVal oSrc As TextStyle, ByVal oDst As TextStyle, ByVal dfActif As DesignFile, ByVal oModelSrc As ModelReference)
    On Error Resume Next

    Set oDst.Font = oSrc.Font

    oDst.Height = TailleReelleSource(oSrc.Height, oModelSrc)
    oDst.Width = TailleReelleSource(oSrc.Width, oModelSrc)
    oDst.Justification = oSrc.Justification
    oDst.Color = oSrc.Color
    oDst.IsBold = oSrc.IsBold
    oDst.IsItalic = oSrc.IsItalic
    oDst.IsUnderlined = oSrc.IsUnderlined

    On Error GoTo 0
End Sub

' [OK confirme 2026-07-28] TextStyle.Height/.Width lus via un DesignFile
' SECONDAIRE (Attachment.DesignFile ou OpenDesignFileForProgram) sont en
' SOUS-UNITES du modele, pas en unites maitres -- contrairement a Point3d
' (KB 6.1) et contrairement a une lecture sur le fichier ACTIF (fiable).
' Verifie runtime Uniformiser : oSrc.Height=45 alors que MicroStation
' affiche 0.45 (direct ET via reference, donc pas un artefact d'echelle de
' reference) ; SubUnitsPerMasterUnit=100 des deux cotes -> 45/100=0.45
' exact. Diviser par SubUnitsPerMasterUnit DU MODELE SOURCE (jamais celui
' du fichier actif, qui n'a pas ce biais).
Private Function TailleReelleSource(ByVal dVal As Double, ByVal oModelSrc As ModelReference) As Double
    Dim dSubPerMu As Double

    On Error Resume Next
    dSubPerMu = oModelSrc.SubUnitsPerMasterUnit
    On Error GoTo 0

    If dSubPerMu = 0 Then
        TailleReelleSource = dVal
    Else
        TailleReelleSource = dVal / dSubPerMu
    End If
End Function

Private Sub AppliquerPalette(ByVal dfActif As DesignFile)
    Dim oItem As CItemUniformisation

    For Each oItem In m_colItems
        If oItem.Categorie = "Palette" And oItem.Inclus Then
            On Error Resume Next
            dfActif.AttachColorTable oItem.ObjetSource
            On Error GoTo 0
        End If
    Next oItem
End Sub

' =============================================================================
' Fermeture
' =============================================================================

Public Sub FermerSource()
    Set m_dfSource = Nothing
End Sub

Private Sub Class_Terminate()
    Set m_dfSource = Nothing
End Sub

' =============================================================================
' Etat expose au formulaire
' =============================================================================

Public Property Get Items() As Collection
    Set Items = m_colItems
End Property

Public Property Get SourceValide() As Boolean
    SourceValide = Not (m_dfSource Is Nothing)
End Property

Public Property Get NbItems() As Long
    If m_colItems Is Nothing Then Exit Property
    NbItems = m_colItems.Count
End Property

Public Property Get NbInclus() As Long
    Dim oItem As CItemUniformisation
    Dim n As Long
    If m_colItems Is Nothing Then Exit Property
    For Each oItem In m_colItems
        If oItem.Inclus Then n = n + 1
    Next oItem
    NbInclus = n
End Property

Public Property Get NbNouveaux() As Long
    Dim oItem As CItemUniformisation
    Dim n As Long
    If m_colItems Is Nothing Then Exit Property
    For Each oItem In m_colItems
        If oItem.Etat = eUniNouveau Then n = n + 1
    Next oItem
    NbNouveaux = n
End Property

Public Property Get NbConflits() As Long
    Dim oItem As CItemUniformisation
    Dim n As Long
    If m_colItems Is Nothing Then Exit Property
    For Each oItem In m_colItems
        If oItem.Etat = eUniConflit Then n = n + 1
    Next oItem
    NbConflits = n
End Property
