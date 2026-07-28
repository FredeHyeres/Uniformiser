VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUniformiser
   Caption         =   "Uniformiser"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4590
   OleObjectBlob   =   "frmUniformiser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUniformiser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' frmUniformiser - Formulaire modeless d'Uniformiser
' Controles crees au runtime dans ConstruireControles
' =============================================================================

Private m_bConstruit    As Boolean
Private m_bInit         As Boolean
Private m_oSettings     As CSettings
Private m_bClicDejaTraite As Boolean   ' voir lstApercu_Click / lstApercu_MouseUp

' --- Controles declares WithEvents pour recevoir les evenements ---
Private WithEvents optAttachement   As MSForms.OptionButton
Attribute optAttachement.VB_VarHelpID = -1
Private WithEvents optFichier       As MSForms.OptionButton
Attribute optFichier.VB_VarHelpID = -1
Private WithEvents cmbAttachements  As MSForms.ComboBox
Attribute cmbAttachements.VB_VarHelpID = -1
Private WithEvents btnParcourir     As MSForms.CommandButton
Attribute btnParcourir.VB_VarHelpID = -1

Private WithEvents chkNiveaux       As MSForms.CheckBox
Attribute chkNiveaux.VB_VarHelpID = -1
Private WithEvents chkStylesTexte   As MSForms.CheckBox
Attribute chkStylesTexte.VB_VarHelpID = -1
Private WithEvents chkPalette       As MSForms.CheckBox
Attribute chkPalette.VB_VarHelpID = -1

Private WithEvents lstApercu        As MSForms.ListBox
Attribute lstApercu.VB_VarHelpID = -1

Private WithEvents btnPrevisualiser As MSForms.CommandButton
Attribute btnPrevisualiser.VB_VarHelpID = -1
Private WithEvents btnConfirmer     As MSForms.CommandButton
Attribute btnConfirmer.VB_VarHelpID = -1
Private WithEvents btnFermer        As MSForms.CommandButton
Attribute btnFermer.VB_VarHelpID = -1

Private txtFichier                  As MSForms.TextBox   ' lecture seule, pas d'evenement
Private lblResultat                 As MSForms.Label
Private lblStatus                   As MSForms.Label

' =============================================================================
Private Sub UserForm_Initialize()
    If Not m_bConstruit Then ConstruireControles
End Sub

' =============================================================================
Public Sub Initialiser(ByVal oSettings As CSettings)
    m_bInit = True
    Set m_oSettings = oSettings

    If Not m_bConstruit Then ConstruireControles

    RemplirAttachements

    optAttachement.Value = Not oSettings.bSourceFichier
    optFichier.Value = oSettings.bSourceFichier
    txtFichier.Text = oSettings.sCheminFichier
    MettreAJourEtatSource

    chkNiveaux.Value = oSettings.bInclureNiveaux
    chkStylesTexte.Value = oSettings.bInclureStylesTexte
    chkPalette.Value = oSettings.bInclurePalette

    lstApercu.Clear
    m_bClicDejaTraite = False
    lblResultat.Caption = ""
    lblStatus.Caption = "Pret."
    btnConfirmer.Enabled = False

    m_bInit = False
End Sub

Private Sub RemplirAttachements()
    Dim colLib As Collection
    Dim v As Variant

    cmbAttachements.Clear

    ' [PIEGE] ActiveModelReference leve "bad model ref" s'il n'y a pas de
    ' modele actif : HasActiveModelReference est le garde documente (CHM).
    If Not Application.HasActiveModelReference Then Exit Sub

    Set colLib = g_oMoteur.ListerAttachments(ActiveModelReference)
    For Each v In colLib
        cmbAttachements.AddItem CStr(v)
    Next v

    If m_oSettings.nIndexAttachement >= 1 And m_oSettings.nIndexAttachement <= cmbAttachements.ListCount Then
        cmbAttachements.ListIndex = m_oSettings.nIndexAttachement - 1
    ElseIf cmbAttachements.ListCount > 0 Then
        cmbAttachements.ListIndex = 0
    End If

    ' cmbAttachements_Change ne fait rien tant que m_bInit est actif (garde
    ' anti-ecrasement), donc la selection ci-dessus n'est jamais reportee
    ' dans les settings : resynchroniser explicitement, sinon Analyser voit
    ' nIndexAttachement = 0 et n'essaie meme pas de resoudre la source.
    m_oSettings.nIndexAttachement = cmbAttachements.ListIndex + 1
End Sub

Private Sub MettreAJourEtatSource()
    Dim bFichier As Boolean
    bFichier = CBool(optFichier.Value)

    cmbAttachements.Enabled = Not bFichier
    txtFichier.Enabled = bFichier
    btnParcourir.Enabled = bFichier
End Sub

' =============================================================================
Public Sub AfficherResultat(ByVal oMoteur As CMoteurUniformisation)
    If Not oMoteur.SourceValide Then
        lstApercu.Clear
        lblResultat.Caption = ""
        lblStatus.Caption = "Source introuvable ou non definie."
        btnConfirmer.Enabled = False
        Exit Sub
    End If

    RemplirApercu oMoteur

    lblResultat.Caption = "Elements : " & oMoteur.NbItems & vbCrLf & _
                          "Nouveaux : " & oMoteur.NbNouveaux & vbCrLf & _
                          "Conflits : " & oMoteur.NbConflits

    If oMoteur.NbItems > 0 Then
        btnConfirmer.Enabled = True
        lblStatus.Caption = "Previsualisation prete."
    Else
        btnConfirmer.Enabled = False
        lblStatus.Caption = "Aucune difference trouvee."
    End If
End Sub

' =============================================================================
Public Sub AfficherApplication()
    lblStatus.Caption = "Application terminee (" & g_oMoteur.NbInclus & " element(s))."
    RemplirApercu g_oMoteur
    btnConfirmer.Enabled = False
End Sub

Private Sub RemplirApercu(ByVal oMoteur As CMoteurUniformisation)
    Dim oItem As CItemUniformisation
    lstApercu.Clear
    m_bClicDejaTraite = False
    For Each oItem In oMoteur.Items
        lstApercu.AddItem LigneApercu(oItem)
    Next oItem
End Sub

Private Function LigneApercu(ByVal oItem As CItemUniformisation) As String
    Dim sCoche As String
    Dim sEtat As String

    sCoche = IIf(oItem.Inclus, "[X]", "[ ]")
    Select Case oItem.Etat
        Case eUniNouveau:   sEtat = "Nouveau"
        Case eUniIdentique: sEtat = "Identique"
        Case eUniConflit:   sEtat = "Conflit"
    End Select

    LigneApercu = sCoche & " " & oItem.Categorie & " | " & oItem.Nom & " | " & sEtat
    If Len(oItem.DetailConflit) > 0 Then
        LigneApercu = LigneApercu & " (" & oItem.DetailConflit & ")"
    End If
End Function

' =============================================================================
Private Sub ConstruireControles()
    Dim fraSource As MSForms.Frame
    Dim fraCategories As MSForms.Frame

    Me.Width = 340
    Me.Height = 486

    ' --- Frame : Source ---
    Set fraSource = Me.Controls.Add("Forms.Frame.1", "fraSource")
    fraSource.Caption = "Source"
    fraSource.Left = 6: fraSource.Top = 6: fraSource.Width = 322: fraSource.Height = 92

    Set optAttachement = fraSource.Controls.Add("Forms.OptionButton.1", "optAttachement")
    optAttachement.Caption = "Reference attachee"
    optAttachement.Left = 8: optAttachement.Top = 14: optAttachement.Width = 150: optAttachement.Height = 14

    Set optFichier = fraSource.Controls.Add("Forms.OptionButton.1", "optFichier")
    optFichier.Caption = "Fichier parcouru"
    optFichier.Left = 164: optFichier.Top = 14: optFichier.Width = 150: optFichier.Height = 14

    Set cmbAttachements = fraSource.Controls.Add("Forms.ComboBox.1", "cmbAttachements")
    cmbAttachements.Left = 8: cmbAttachements.Top = 32: cmbAttachements.Width = 306: cmbAttachements.Height = 18
    cmbAttachements.Style = 2   ' fmStyleDropDownList

    Set txtFichier = fraSource.Controls.Add("Forms.TextBox.1", "txtFichier")
    txtFichier.Left = 8: txtFichier.Top = 58: txtFichier.Width = 246: txtFichier.Height = 18
    txtFichier.Locked = True

    Set btnParcourir = fraSource.Controls.Add("Forms.CommandButton.1", "btnParcourir")
    btnParcourir.Caption = "..."
    btnParcourir.Left = 258: btnParcourir.Top = 58: btnParcourir.Width = 56: btnParcourir.Height = 18

    ' --- Frame : Categories ---
    Set fraCategories = Me.Controls.Add("Forms.Frame.1", "fraCategories")
    fraCategories.Caption = "Categories a recuperer"
    fraCategories.Left = 6: fraCategories.Top = 102: fraCategories.Width = 322: fraCategories.Height = 40

    Set chkNiveaux = fraCategories.Controls.Add("Forms.CheckBox.1", "chkNiveaux")
    chkNiveaux.Caption = "Niveaux"
    chkNiveaux.Left = 8: chkNiveaux.Top = 16: chkNiveaux.Width = 100: chkNiveaux.Height = 14

    Set chkStylesTexte = fraCategories.Controls.Add("Forms.CheckBox.1", "chkStylesTexte")
    chkStylesTexte.Caption = "Styles de texte"
    chkStylesTexte.Left = 112: chkStylesTexte.Top = 16: chkStylesTexte.Width = 110: chkStylesTexte.Height = 14

    Set chkPalette = fraCategories.Controls.Add("Forms.CheckBox.1", "chkPalette")
    chkPalette.Caption = "Palette de couleurs"
    chkPalette.Left = 226: chkPalette.Top = 16: chkPalette.Width = 96: chkPalette.Height = 14

    ' --- Actions : previsualiser / confirmer ---
    Set btnPrevisualiser = Me.Controls.Add("Forms.CommandButton.1", "btnPrevisualiser")
    btnPrevisualiser.Caption = "Previsualiser"
    btnPrevisualiser.Left = 6: btnPrevisualiser.Top = 148: btnPrevisualiser.Width = 104: btnPrevisualiser.Height = 22

    Set btnConfirmer = Me.Controls.Add("Forms.CommandButton.1", "btnConfirmer")
    btnConfirmer.Caption = "Confirmer"
    btnConfirmer.Left = 116: btnConfirmer.Top = 148: btnConfirmer.Width = 104: btnConfirmer.Height = 22
    btnConfirmer.Enabled = False

    ' --- Liste de previsualisation ---
    Set lstApercu = Me.Controls.Add("Forms.ListBox.1", "lstApercu")
    lstApercu.Left = 6: lstApercu.Top = 176: lstApercu.Width = 322: lstApercu.Height = 210
    lstApercu.ColumnCount = 1
    ' IntegralHeight=True (defaut) arrondit la hauteur visible a un nombre
    ' entier de lignes et rend la derniere ligne inaccessible au scroll
    ' quand Height n'est pas un multiple exact de la hauteur de ligne
    ' (bug MSForms connu). False garantit que toute la liste reste
    ' atteignable, y compris les listes longues.
    lstApercu.IntegralHeight = False

    ' --- Resultat + statut + fermer ---
    Set lblResultat = Me.Controls.Add("Forms.Label.1", "lblResultat")
    lblResultat.Caption = ""
    lblResultat.Left = 6: lblResultat.Top = 390: lblResultat.Width = 322: lblResultat.Height = 40
    lblResultat.WordWrap = True

    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus")
    lblStatus.Caption = "Pret."
    lblStatus.Left = 6: lblStatus.Top = 434: lblStatus.Width = 200: lblStatus.Height = 14

    Set btnFermer = Me.Controls.Add("Forms.CommandButton.1", "btnFermer")
    btnFermer.Caption = "Fermer"
    btnFermer.Left = 250: btnFermer.Top = 432: btnFermer.Width = 78: btnFermer.Height = 20

    m_bConstruit = True
End Sub

' =============================================================================
' Evenements - Source
' =============================================================================
Private Sub optAttachement_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bSourceFichier = False
    MettreAJourEtatSource
End Sub

Private Sub optFichier_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bSourceFichier = True
    MettreAJourEtatSource
End Sub

Private Sub cmbAttachements_Change()
    If m_bInit Then Exit Sub
    m_oSettings.nIndexAttachement = cmbAttachements.ListIndex + 1
End Sub

Private Sub btnParcourir_Click()
    Dim sChemin As String
    sChemin = MFileDialog.AfficherOuvrirFichier("Choisir le fichier DGN source")
    If Len(sChemin) > 0 Then
        txtFichier.Text = sChemin
        m_oSettings.sCheminFichier = sChemin
    End If
End Sub

' =============================================================================
' Evenements - Categories
' =============================================================================
Private Sub chkNiveaux_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bInclureNiveaux = chkNiveaux.Value
End Sub

Private Sub chkStylesTexte_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bInclureStylesTexte = chkStylesTexte.Value
End Sub

Private Sub chkPalette_Click()
    If m_bInit Then Exit Sub
    m_oSettings.bInclurePalette = chkPalette.Value
End Sub

' =============================================================================
' Evenement - Previsualisation : case a cocher par ligne
' [OK confirme] Click seul ne se declenche PAS quand on reclique une ligne
' deja selectionnee (limitation MSForms.ListBox : Click ne suit que les
' changements de ListIndex), ce qui empechait de decocher un element sur
' lequel le focus etait deja positionne. MouseUp se declenche a CHAQUE clic
' gauche, y compris sur la ligne deja selectionnee ; le flag
' m_bClicDejaTraite evite de basculer deux fois (Click puis MouseUp) le cas
' normal ou la ligne cliquee change.
' =============================================================================
Private Sub lstApercu_Click()
    m_bClicDejaTraite = True
    BasculerLigneCourante
End Sub

Private Sub lstApercu_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Then Exit Sub
    If m_bClicDejaTraite Then
        m_bClicDejaTraite = False
        Exit Sub
    End If
    BasculerLigneCourante
End Sub

Private Sub BasculerLigneCourante()
    Dim oItem As CItemUniformisation
    Dim nSel As Long

    If g_oMoteur Is Nothing Then Exit Sub
    nSel = lstApercu.ListIndex
    If nSel < 0 Then Exit Sub

    Set oItem = g_oMoteur.Items(nSel + 1)
    oItem.Inclus = Not oItem.Inclus

    lstApercu.List(nSel) = LigneApercu(oItem)
    lstApercu.ListIndex = nSel
End Sub

' =============================================================================
' Boutons d'action
' =============================================================================
Private Sub btnPrevisualiser_Click()
    btnPrevisualiser.Enabled = False
    lblStatus.Caption = "Analyse en cours..."
    lblResultat.Caption = ""
    btnConfirmer.Enabled = False
    DoEvents
    LancerAnalyse
    btnPrevisualiser.Enabled = True
End Sub

Private Sub btnConfirmer_Click()
    Dim rep As VbMsgBoxResult
    rep = MsgBox("Confirmer l'application de " & g_oMoteur.NbInclus & " element(s) coche(s) ?", _
                 vbYesNo + vbQuestion, "Uniformiser")
    If rep = vbYes Then
        lblStatus.Caption = "Application en cours..."
        DoEvents
        LancerApplication
    End If
End Sub

Private Sub btnFermer_Click()
    NettoyerAvantFermeture
    Me.Hide
    CommandState.StartDefaultCommand
End Sub

' =============================================================================
' Fermeture par la croix
' =============================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = 1
        NettoyerAvantFermeture
        Me.Hide
        CommandState.StartDefaultCommand
    End If
End Sub

