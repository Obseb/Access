Attribute VB_Name = "mod_Synerlink_Machines"
Option Compare Database
Option Explicit

'===============================================================================
' Formulaire F_Synerlink_Machines : gestion de T_Synerlink_Machines avant import
'
'   - ouvert en mode Dialogue par le bouton Mise à jour de F_Synerlink
'   - liste des lignes (niveau 1) et machines (niveau 2) dans l'ordre :
'     modifier, ajouter, supprimer, activer / désactiver
'   - colonne « Fichier » : indique si le fichier Excel de la machine existe
'   - bouton « Lancer la mise à jour » -> gSynerlinkLancer = True puis fermeture
'   - bouton « Fermer »                -> enregistre et ferme sans import
'
' CreerFormulaireMachines : à lancer UNE fois (fenêtre Exécution) pour créer le
' formulaire. Peut être relancée : elle remplace le formulaire existant.
'===============================================================================

Public gSynerlinkLancer As Boolean   ' Vrai si l'utilisateur a demandé l'import

' Pour la colonne « Fichier » du formulaire : "OK", "Absent", ou "" pour une ligne
Public Function SynerlinkFichierEtat(ByVal niveau As Variant, ByVal ref As Variant) As String
    Static dossier As String
    If Nz(niveau, 0) <> 2 Or Len(Nz(ref, "")) = 0 Then Exit Function
    If Len(dossier) = 0 Then dossier = DossierBOM()
    If Len(dossier) = 0 Then SynerlinkFichierEtat = "?": Exit Function
    If Len(Dir(dossier & ref & ".xlsx")) > 0 Then
        SynerlinkFichierEtat = "OK"
    Else
        SynerlinkFichierEtat = "Absent"
    End If
End Function

'-------------------------------------------------------------------------------
' Création du formulaire
'-------------------------------------------------------------------------------
Public Sub CreerFormulaireMachines()
    Const NOM As String = "F_Synerlink_Machines"
    Const CM As Long = 567                      ' 1 cm en twips
    Dim frm As Form, ctl As Control, fc As FormatCondition
    Dim nomTemp As String, x As Long
    Dim mdl As Module, code As String

    ' Supprime l'ancien formulaire s'il existe
    On Error Resume Next
    DoCmd.Close acForm, NOM, acSaveNo
    DoCmd.DeleteObject acForm, NOM
    On Error GoTo 0

    Set frm = CreateForm()
    nomTemp = frm.Name
    DoCmd.RunCommand acCmdFormHdrFtr           ' ajoute en-tête et pied de formulaire

    With frm
        .Caption = "Machines Synerlink"
        .RecordSource = "SELECT * FROM T_Synerlink_Machines ORDER BY Ordre"
        .DefaultView = 1                        ' formulaires continus
        .AllowAdditions = True
        .AllowDeletions = True
        .AllowEdits = True
        .RecordSelectors = True                 ' pour sélectionner / supprimer une ligne
        .NavigationButtons = False
        .ScrollBars = 2                         ' verticale
        .PopUp = True
        .Modal = True
        .AutoCenter = True
        .BeforeUpdate = "[Event Procedure]"
        .Width = 18 * CM
        .Section(acHeader).Height = 0.8 * CM
        .Section(acDetail).Height = 0.6 * CM
        .Section(acFooter).Height = 1.2 * CM
    End With

    ' Colonnes : en-tête (étiquette) + contrôle dans le détail
    x = 0.2 * CM
    AjouterColonne nomTemp, "Ordre", "Ordre", acTextBox, x, 1.6 * CM
    AjouterColonne nomTemp, "Niveau", "Type", acComboBox, x, 2.4 * CM
    AjouterColonne nomTemp, "Reference", "Référence", acTextBox, x, 2.8 * CM
    AjouterColonne nomTemp, "Libelle", "Libellé", acTextBox, x, 6.5 * CM
    AjouterColonne nomTemp, "Actif", "Actif", acCheckBox, x, 1.2 * CM
    AjouterColonne nomTemp, "", "Fichier", acTextBox, x, 1.8 * CM, "txtFichier"

    ' Réglages particuliers
    With frm.Controls("Ordre")
        .DefaultValue = "=Nz(DMax(""Ordre"",""T_Synerlink_Machines""),0)+1"
    End With
    With frm.Controls("Niveau")
        .RowSourceType = "Value List"
        .RowSource = "1;""Ligne"";2;""Machine"""
        .ColumnCount = 2
        .ColumnWidths = "0;" & 2 * CM      ' 1re colonne (n°) masquée
        .BoundColumn = 1
        .LimitToList = True
        .DefaultValue = "2"
    End With
    With frm.Controls("Actif")
        .DefaultValue = "True"
    End With
    With frm.Controls("txtFichier")
        .ControlSource = "=SynerlinkFichierEtat([Niveau],[Reference])"
        .Locked = True
        .Enabled = False
        .TextAlign = 2                          ' centré
        Set fc = .FormatConditions.Add(acFieldValue, acEqual, """Absent""")
        fc.ForeColor = vbRed
        fc.FontBold = True
    End With
    ' Lignes de production (niveau 1) en gras
    Set fc = frm.Controls("Reference").FormatConditions.Add(acExpression, , "[Niveau]=1")
    fc.FontBold = True
    Set fc = frm.Controls("Libelle").FormatConditions.Add(acExpression, , "[Niveau]=1")
    fc.FontBold = True

    ' Boutons du pied de formulaire
    Set ctl = CreateControl(nomTemp, acCommandButton, acFooter, , , 9.5 * CM, 0.25 * CM, 4.6 * CM, 0.7 * CM)
    ctl.Name = "BP_Lancer"
    ctl.Caption = "Lancer la mise à jour"
    ctl.OnClick = "[Event Procedure]"
    Set ctl = CreateControl(nomTemp, acCommandButton, acFooter, , , 14.3 * CM, 0.25 * CM, 3 * CM, 0.7 * CM)
    ctl.Name = "BP_Fermer"
    ctl.Caption = "Fermer"
    ctl.OnClick = "[Event Procedure]"
    Set ctl = CreateControl(nomTemp, acLabel, acFooter, , , 0.2 * CM, 0.3 * CM, 9 * CM, 0.6 * CM)
    ctl.Name = "lblAide"
    ctl.Caption = "Sélectionnez une ligne (case de gauche) puis Suppr pour la supprimer."
    ctl.FontItalic = True

    ' Code du formulaire
    code = "Option Explicit" & vbCrLf & vbCrLf & _
        "' Contrôle avant enregistrement d'une ligne" & vbCrLf & _
        "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf & _
        "    If IsNull(Me.Ordre) Or IsNull(Me.Niveau) Or Len(Nz(Me.Reference, """")) = 0 Then" & vbCrLf & _
        "        MsgBox ""Renseignez l'ordre, le type et la référence."", vbExclamation, ""Synerlink""" & vbCrLf & _
        "        Cancel = True" & vbCrLf & _
        "    End If" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "' Enregistre la ligne en cours ; renvoie Faux si l'enregistrement est refusé" & vbCrLf & _
        "Private Function Enregistrer() As Boolean" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    If Me.Dirty Then Me.Dirty = False" & vbCrLf & _
        "    Enregistrer = Not Me.Dirty" & vbCrLf & _
        "End Function" & vbCrLf & vbCrLf & _
        "Private Sub BP_Lancer_Click()" & vbCrLf & _
        "    Dim premier As Variant" & vbCrLf & _
        "    If Not Enregistrer() Then Exit Sub" & vbCrLf & _
        "    ' La première ligne active doit être une ligne de production (niveau 1)" & vbCrLf & _
        "    premier = DMin(""Ordre"", ""T_Synerlink_Machines"", ""Actif"")" & vbCrLf & _
        "    If IsNull(premier) Then" & vbCrLf & _
        "        MsgBox ""Aucune ligne active."", vbExclamation, ""Synerlink"": Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    If DLookup(""Niveau"", ""T_Synerlink_Machines"", ""Ordre="" & premier) <> 1 Then" & vbCrLf & _
        "        MsgBox ""La première ligne active doit être une ligne de production (type Ligne)."", vbExclamation, ""Synerlink"": Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    gSynerlinkLancer = True" & vbCrLf & _
        "    DoCmd.Close acForm, Me.Name" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub BP_Fermer_Click()" & vbCrLf & _
        "    If Not Enregistrer() Then Exit Sub" & vbCrLf & _
        "    gSynerlinkLancer = False" & vbCrLf & _
        "    DoCmd.Close acForm, Me.Name" & vbCrLf & _
        "End Sub"
    frm.HasModule = True
    Set mdl = frm.Module
    mdl.AddFromString code

    DoCmd.Close acForm, nomTemp, acSaveYes
    DoCmd.Rename NOM, acForm, nomTemp
    MsgBox "Formulaire " & NOM & " créé.", vbInformation
End Sub

' Ajoute une colonne : étiquette dans l'en-tête + contrôle dans le détail
Private Sub AjouterColonne(ByVal nomForm As String, ByVal champ As String, ByVal titre As String, _
                           ByVal typeCtl As AcControlType, ByRef x As Long, ByVal largeur As Long, _
                           Optional ByVal nomCtl As String = "")
    Const CM As Long = 567
    Dim ctl As Control, lbl As Control
    Dim l As Long

    l = largeur
    If typeCtl = acCheckBox Then
        Set ctl = CreateControl(nomForm, typeCtl, acDetail, , champ, x + l / 2 - 0.2 * CM, 0.15 * CM)
    Else
        Set ctl = CreateControl(nomForm, typeCtl, acDetail, , champ, x, 0.05 * CM, l, 0.5 * CM)
    End If
    If Len(nomCtl) > 0 Then ctl.Name = nomCtl Else ctl.Name = champ

    Set lbl = CreateControl(nomForm, acLabel, acHeader, , , x, 0.2 * CM, l, 0.5 * CM)
    lbl.Name = "lbl_" & ctl.Name
    lbl.Caption = titre
    lbl.FontBold = True
    If typeCtl = acCheckBox Then lbl.TextAlign = 2

    x = x + l + 0.1 * CM
End Sub
