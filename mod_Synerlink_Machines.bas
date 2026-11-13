Attribute VB_Name = "mod_Synerlink_Machines"
Option Compare Database
Option Explicit

'===============================================================================
' Gestion de T_Synerlink_Machines avant import
'
'   F_Synerlink_Machines      : fenêtre (mode Dialogue) ouverte par le bouton
'                               Mise à jour de F_Synerlink, avec 2 boutons :
'                               « Lancer la mise à jour » -> gSynerlinkLancer = True
'                               « Fermer »                -> sans import
'   F_Synerlink_Machines_SF   : sous-formulaire en feuille de données sur
'                               T_Synerlink_Machines (ajout, suppression, Actif)
'
' CreerFormulaireMachines : à lancer UNE fois (fenêtre Exécution) pour créer les
' deux formulaires. Peut être relancée : elle remplace les formulaires existants.
'===============================================================================

Public gSynerlinkLancer As Boolean   ' Vrai si l'utilisateur a demandé l'import

Private Const F_MAIN As String = "F_Synerlink_Machines"
Private Const F_SF As String = "F_Synerlink_Machines_SF"
Private Const CM As Long = 567       ' 1 cm en twips

Public Sub CreerFormulaireMachines()
    On Error Resume Next
    DoCmd.Close acForm, F_MAIN, acSaveNo
    DoCmd.Close acForm, F_SF, acSaveNo
    DoCmd.DeleteObject acForm, F_MAIN
    DoCmd.DeleteObject acForm, F_SF
    On Error GoTo 0

    CreerSousFormulaire
    CreerFormulairePrincipal
    MsgBox "Formulaires " & F_MAIN & " et " & F_SF & " créés.", vbInformation
End Sub

'-------------------------------------------------------------------------------
' Sous-formulaire : feuille de données sur T_Synerlink_Machines
'-------------------------------------------------------------------------------
Private Sub CreerSousFormulaire()
    Dim frm As Form, ctl As Control
    Dim nomTemp As String, code As String, y As Long

    Set frm = CreateForm()
    nomTemp = frm.Name
    With frm
        .RecordSource = "SELECT * FROM T_Synerlink_Machines ORDER BY Ordre"
        .DefaultView = 2                        ' feuille de données
        .AllowFormView = False
        .AllowDatasheetView = True
        .AllowAdditions = True
        .AllowDeletions = True
        .AllowEdits = True
        .BeforeUpdate = "[Event Procedure]"
    End With

    ' En feuille de données, le titre de colonne = nom du contrôle
    y = 0.2 * CM
    Set ctl = CreateControl(nomTemp, acTextBox, acDetail, , "Ordre", 0.2 * CM, y, 2 * CM, 0.5 * CM)
    ctl.Name = "Ordre"
    ctl.DefaultValue = "=Nz(DMax(""Ordre"",""T_Synerlink_Machines""),0)+1"
    LargeurColonne ctl, 1.5 * CM

    y = y + 0.6 * CM
    Set ctl = CreateControl(nomTemp, acComboBox, acDetail, , "Niveau", 0.2 * CM, y, 2.5 * CM, 0.5 * CM)
    ctl.Name = "Type"
    ctl.RowSourceType = "Value List"
    ctl.RowSource = "1;""Ligne"";2;""Machine"""
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0;" & 2 * CM
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    ctl.DefaultValue = "2"
    LargeurColonne ctl, 2.2 * CM

    y = y + 0.6 * CM
    Set ctl = CreateControl(nomTemp, acTextBox, acDetail, , "Reference", 0.2 * CM, y, 3 * CM, 0.5 * CM)
    ctl.Name = "Référence"
    LargeurColonne ctl, 2.6 * CM

    y = y + 0.6 * CM
    Set ctl = CreateControl(nomTemp, acTextBox, acDetail, , "Libelle", 0.2 * CM, y, 6 * CM, 0.5 * CM)
    ctl.Name = "Libellé"
    LargeurColonne ctl, 7 * CM

    y = y + 0.6 * CM
    Set ctl = CreateControl(nomTemp, acCheckBox, acDetail, , "Actif", 0.2 * CM, y)
    ctl.Name = "Actif"
    ctl.DefaultValue = "True"
    LargeurColonne ctl, 1.2 * CM

    ' Contrôle avant enregistrement d'une ligne
    code = ""
    code = code & "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf
    code = code & "    If IsNull(Me!Ordre) Or IsNull(Me!Niveau) Or Len(Nz(Me!Reference, """")) = 0 Then" & vbCrLf
    code = code & "        MsgBox ""Renseignez l'ordre, le type et la référence."", vbExclamation, ""Synerlink""" & vbCrLf
    code = code & "        Cancel = True" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Sub"
    frm.HasModule = True
    frm.Module.AddFromString code

    DoCmd.Close acForm, nomTemp, acSaveYes
    DoCmd.Rename F_SF, acForm, nomTemp
End Sub

' Largeur de colonne en feuille de données (ignorée si Access la refuse en mode Création)
Private Sub LargeurColonne(ctl As Control, ByVal largeur As Long)
    On Error Resume Next
    ctl.ColumnWidth = largeur
End Sub

'-------------------------------------------------------------------------------
' Formulaire principal : sous-formulaire + boutons
'-------------------------------------------------------------------------------
Private Sub CreerFormulairePrincipal()
    Dim frm As Form, ctl As Control
    Dim nomTemp As String, code As String

    Set frm = CreateForm()
    nomTemp = frm.Name
    With frm
        .Caption = "Machines Synerlink"
        .RecordSelectors = False
        .NavigationButtons = False
        .DividingLines = False
        .ScrollBars = 0
        .PopUp = True
        .Modal = True
        .AutoCenter = True
        .AutoResize = True
        .Width = 16.6 * CM
        .Section(acDetail).Height = 12.6 * CM
    End With

    Set ctl = CreateControl(nomTemp, acSubform, acDetail, , , 0.3 * CM, 0.3 * CM, 16 * CM, 11 * CM)
    ctl.Name = "SF_Machines"
    ctl.SourceObject = F_SF

    Set ctl = CreateControl(nomTemp, acLabel, acDetail, , , 0.3 * CM, 11.6 * CM, 8.5 * CM, 0.6 * CM)
    ctl.Name = "lblAide"
    ctl.Caption = "Pour supprimer : sélectionnez la ligne (case grise à gauche) puis Suppr."
    ctl.FontItalic = True

    Set ctl = CreateControl(nomTemp, acCommandButton, acDetail, , , 9 * CM, 11.5 * CM, 4.4 * CM, 0.75 * CM)
    ctl.Name = "BP_Lancer"
    ctl.Caption = "Lancer la mise à jour"
    ctl.OnClick = "[Event Procedure]"

    Set ctl = CreateControl(nomTemp, acCommandButton, acDetail, , , 13.6 * CM, 11.5 * CM, 2.7 * CM, 0.75 * CM)
    ctl.Name = "BP_Fermer"
    ctl.Caption = "Fermer"
    ctl.OnClick = "[Event Procedure]"

    code = ""
    code = code & "' Enregistre la ligne en cours de la feuille de données ; Faux si refusé" & vbCrLf
    code = code & "Private Function Enregistrer() As Boolean" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    With Me.SF_Machines.Form" & vbCrLf
    code = code & "        If .Dirty Then .Dirty = False" & vbCrLf
    code = code & "        Enregistrer = Not .Dirty" & vbCrLf
    code = code & "    End With" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf
    code = code & "Private Sub BP_Lancer_Click()" & vbCrLf
    code = code & "    Dim premier As Variant" & vbCrLf
    code = code & "    If Not Enregistrer() Then Exit Sub" & vbCrLf
    code = code & "    ' La première ligne active doit être une ligne de production (niveau 1)" & vbCrLf
    code = code & "    premier = DMin(""Ordre"", ""T_Synerlink_Machines"", ""Actif"")" & vbCrLf
    code = code & "    If IsNull(premier) Then" & vbCrLf
    code = code & "        MsgBox ""Aucune ligne active."", vbExclamation, ""Synerlink"": Exit Sub" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    If DLookup(""Niveau"", ""T_Synerlink_Machines"", ""Ordre="" & premier) <> 1 Then" & vbCrLf
    code = code & "        MsgBox ""La première ligne active doit être une ligne de production (type Ligne)."", vbExclamation, ""Synerlink"": Exit Sub" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    gSynerlinkLancer = True" & vbCrLf
    code = code & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    code = code & "Private Sub BP_Fermer_Click()" & vbCrLf
    code = code & "    If Not Enregistrer() Then Exit Sub" & vbCrLf
    code = code & "    gSynerlinkLancer = False" & vbCrLf
    code = code & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    code = code & "End Sub"
    frm.HasModule = True
    frm.Module.AddFromString code

    DoCmd.Close acForm, nomTemp, acSaveYes
    DoCmd.Rename F_MAIN, acForm, nomTemp
End Sub
