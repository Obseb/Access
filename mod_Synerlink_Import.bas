Attribute VB_Name = "mod_Synerlink_Import"
Option Compare Database
Option Explicit

'===============================================================================
' Réactualisation de T_Synerlink_Arbo à partir des fichiers <machine>.xlsx
' (ex. C:\GMAO\Hierarchie Synerlink Access\6155DO1.xlsx)
'
' Points d'entrée :
'   SynerlinkImport()            -> lit les Excel et remplace T_Synerlink_Arbo
'   SynerlinkCompacter()         -> compacte la base qui contient les tables
'                                   (si personne d'autre ne l'utilise)
'
' Les tables T_Synerlink_* peuvent être locales ou liées (base séparée, ex.
' DB_GMAO_Machine.accdb) : le module écrit directement dans la base qui les
' contient, sauvegarde comprise.
'
' Format Excel attendu (les colonnes sont repérées par leur titre en ligne 1) :
'   Onglet « Global »            : ignoré
'   Un onglet par sous-ensemble  : t Global/Structure | Qté | Reference |
'                                  Désignation (Fr) | Plans
'   (l'ancien format avec Revision / Plan/Drawing / Designation reste accepté)
'   Onglets ignorés : Global, Feuil*, Plans_Drawings*, M000
'
' Construction de la table :
'   Niveau 1 = ligne, niveau 2 = machine (libellés/ordre : T_Synerlink_Machines)
'   Niveau 3 = chaque onglet, dans l'ordre du classeur
'   Niveaux suivants = lignes de l'onglet, profondeur lue en colonne A
'   Plan = nom du fichier seul ; Treeview = Vrai si la ligne a des enfants
'   ID = 1, 2, 3... dans l'ordre de l'arborescence (champ Entier long)
'===============================================================================

Private Const DOSSIER_BOM As String = "C:\GMAO\Hierarchie Synerlink Access\"
Private Const SUFFIXE_FICHIER As String = ".xlsx"
Private Const TABLE_ARBO As String = "T_Synerlink_Arbo"
Private Const TABLE_SAUV As String = "T_Synerlink_Arbo_Sauv"
Private Const TABLE_PARAM As String = "T_Synerlink_Machines"

' Une ligne de la future table
Private Type tLigne
    Niveau As Integer
    Arbo As String
    Qte As Variant
    Ref As Variant
    Rev As Variant
    Descr As Variant
    Plan As Variant
End Type

Private mLignes() As tLigne
Private mNb As Long
Private mAvertissements As String

'===============================================================================
' IMPORT
'===============================================================================
Public Function SynerlinkImport() As Boolean
    Dim t0 As Single

    On Error GoTo Erreur
    t0 = Timer
    mNb = 0
    mAvertissements = ""
    ReDim mLignes(1 To 100000)

    ' 1. Table de paramètres (créée au premier lancement si elle n'existe pas)
    If Not TableExiste(TABLE_PARAM) Then CreerTableParametres

    ' 2. Contrôle de présence des fichiers
    Dim manquants As String, nouveaux As String
    ControlerFichiers manquants, nouveaux
    If Len(manquants) > 0 Then
        MsgBox "Import annulé : fichiers introuvables dans" & vbCrLf & DOSSIER_BOM & vbCrLf & vbCrLf & _
               manquants & vbCrLf & _
               "Ajoutez ces fichiers, ou décochez « Actif » pour ces machines dans " & TABLE_PARAM & ".", _
               vbExclamation, "Synerlink"
        Exit Function
    End If
    If Len(nouveaux) > 0 Then
        If MsgBox("Ces fichiers sont dans le dossier mais pas dans " & TABLE_PARAM & _
                  " (ils seront ignorés) :" & vbCrLf & vbCrLf & nouveaux & vbCrLf & _
                  "Continuer quand même ?", vbQuestion + vbYesNo, "Synerlink") = vbNo Then Exit Function
    End If

    If MsgBox("Réactualiser " & TABLE_ARBO & " à partir des fichiers Excel ?" & vbCrLf & _
              "(la table actuelle sera remplacée)", vbQuestion + vbOKCancel, "Synerlink") = vbCancel Then Exit Function

    ' 3. Lecture de tous les Excel en mémoire (aucune modification de la base)
    DoCmd.Hourglass True
    LireTousLesFichiers
    If mNb = 0 Then Err.Raise vbObjectError + 1, , "Aucune ligne lue dans les fichiers Excel."

    ' 4. Écriture dans la base qui contient la table
    EcrireTable

    DoCmd.Hourglass False
    SysCmd acSysCmdClearStatus
    MsgBox "Import terminé : " & Format(mNb, "#,##0") & " lignes en " & _
           Format(Timer - t0, "0") & " s." & _
           IIf(Len(mAvertissements) > 0, vbCrLf & vbCrLf & "Remarques :" & vbCrLf & mAvertissements, ""), _
           vbInformation, "Synerlink"
    SynerlinkImport = True
    Exit Function

Erreur:
    DoCmd.Hourglass False
    SysCmd acSysCmdClearStatus
    MsgBox "Erreur pendant l'import :" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "La table n'a pas été modifiée (ou a été restaurée depuis " & TABLE_SAUV & ").", _
           vbCritical, "Synerlink"
End Function

'-------------------------------------------------------------------------------
' Lecture des classeurs
'-------------------------------------------------------------------------------
Private Sub LireTousLesFichiers()
    Dim xl As Object, wb As Object, ws As Object
    Dim rs As DAO.Recordset
    Dim nLigne As Long, nMachine As Long, iOnglet As Long
    Dim fichier As String, nbFich As Long, iFich As Long

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & TABLE_PARAM & " WHERE Actif ORDER BY Ordre", dbOpenSnapshot)
    If Not rs.EOF Then rs.MoveLast: nbFich = rs.RecordCount: rs.MoveFirst

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False

    On Error GoTo Erreur
    Do Until rs.EOF
        iFich = iFich + 1
        If rs!Niveau = 1 Then
            nLigne = nLigne + 1
            nMachine = 0
            Ajouter 1, CStr(nLigne), 0, rs!Reference, Null, rs!Libelle, Null
        Else
            nMachine = nMachine + 1
            Ajouter 2, nLigne & "." & nMachine, 0, rs!Reference, Null, rs!Libelle, Null
            fichier = DOSSIER_BOM & rs!Reference & SUFFIXE_FICHIER
            SysCmd acSysCmdSetStatus, "Lecture " & rs!Reference & " (" & iFich & "/" & nbFich & ")..."
            Set wb = xl.Workbooks.Open(fichier, 0, True)    ' UpdateLinks:=0, ReadOnly:=True
            iOnglet = 0
            For Each ws In wb.Worksheets
                If Not OngletIgnore(ws.Name) Then
                    If LireOnglet(ws, nLigne & "." & nMachine, iOnglet + 1) Then iOnglet = iOnglet + 1
                End If
            Next ws
            wb.Close False
            Set wb = Nothing
        End If
        rs.MoveNext
    Loop
    rs.Close
    xl.Quit
    Set xl = Nothing
    Exit Sub

Erreur:
    Dim msg As String
    msg = Err.Description & IIf(Len(fichier) > 0, " (fichier " & fichier & ")", "")
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    xl.Quit
    Set xl = Nothing
    On Error GoTo 0
    Err.Raise vbObjectError + 2, , msg
End Sub

Private Function OngletIgnore(ByVal nom As String) As Boolean
    nom = LCase$(Trim$(nom))
    OngletIgnore = (nom = "global") Or (nom Like "feuil*") Or (nom Like "plans_drawings*") Or (nom = "m000")
End Function

' Lit un onglet ; renvoie Faux s'il est vide ou n'a pas le bon format
Private Function LireOnglet(ws As Object, ByVal prefixe As String, ByVal numOnglet As Long) As Boolean
    Dim v As Variant, derniere As Long, r As Long
    Dim compteurs(1 To 30) As Long
    Dim prof As Integer, niv As Integer, k As Integer
    Dim arbo As String
    Dim c As Long, titre As String
    Dim cQte As Long, cRef As Long, cRev As Long, cDes As Long, cPlan As Long
    Dim rev As Variant, plan As Variant

    derniere = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If derniere < 3 Then Exit Function

    v = ws.Range(ws.Cells(1, 1), ws.Cells(derniere, 10)).Value2

    ' Repérage des colonnes par leur titre (ligne 1)
    For c = 2 To 10
        titre = LCase$(Trim$(Replace(Replace(Chaine(v(1, c)), vbLf, " "), vbCr, "")))
        Select Case True
            Case titre Like "qt*":              If cQte = 0 Then cQte = c
            Case titre = "reference":           cRef = c
            Case titre = "revision":            cRev = c
            Case titre = "désignation (fr)":    cDes = c
            Case titre = "plans":               cPlan = c
            Case titre = "designation":         If cPlan = 0 Then cPlan = c   ' ancien format
        End Select
    Next c

    If cRef = 0 Or cQte = 0 Or cDes = 0 Then
        mAvertissements = mAvertissements & "- onglet « " & ws.Name & " » de " & ws.Parent.Name & _
                          " ignoré (en-tête inattendu)" & vbCrLf
        Exit Function
    End If

    compteurs(3) = numOnglet - 1
    For r = 3 To derniere
        If Not LigneVide(v, r) Then
            prof = Profondeur(v(r, 1))
            niv = prof + 2
            If niv > 30 Then niv = 30
            ' Numérotation : on incrémente le niveau courant et on remet à zéro les suivants
            compteurs(niv) = compteurs(niv) + 1
            For k = niv + 1 To 30: compteurs(k) = 0: Next k
            arbo = prefixe
            For k = 3 To niv
                arbo = arbo & "." & compteurs(k)
            Next k

            If cRev > 0 Then rev = Texte(v(r, cRev)) Else rev = Null
            If cPlan > 0 Then plan = NomPlan(v(r, cPlan)) Else plan = Null
            Ajouter niv, arbo, Quantite(v(r, cQte)), Texte(v(r, cRef)), rev, Texte(v(r, cDes)), plan
            LireOnglet = True
        End If
    Next r
End Function

' Valeur de cellule -> texte (cellule vide ou en erreur #N/A -> "")
Private Function Chaine(ByVal x As Variant) As String
    If IsError(x) Or IsNull(x) Or IsEmpty(x) Then Exit Function
    Chaine = CStr(x)
End Function

Private Function LigneVide(v As Variant, ByVal r As Long) As Boolean
    Dim c As Long
    For c = 1 To UBound(v, 2)
        If Len(Trim$(Chaine(v(r, c)))) > 0 Then Exit Function
    Next c
    LigneVide = True
End Function

' « 1.1.1 » -> 3 ; accepte aussi la virgule si Excel a converti en nombre
Private Function Profondeur(ByVal x As Variant) As Integer
    Dim s As String
    s = Trim$(Chaine(x))
    Profondeur = 1 + (Len(s) - Len(Replace(Replace(s, ".", ""), ",", "")))
End Function

Private Function Quantite(ByVal x As Variant) As Variant
    Dim s As String
    s = Trim$(Chaine(x))
    If Len(s) = 0 Then
        Quantite = Null
    Else
        Quantite = CSng(Val(Replace(s, ",", ".")))
    End If
End Function

Private Function Texte(ByVal x As Variant) As Variant
    If IsError(x) Or IsNull(x) Or IsEmpty(x) Then Texte = Null: Exit Function
    If Len(CStr(x)) = 0 Then Texte = Null: Exit Function
    Texte = Left$(CStr(x), 255)
End Function

' Garde uniquement le nom du fichier : « PLANS_DRAWINGS/GP00_xxx.pdf » -> « GP00_xxx.pdf »
Private Function NomPlan(ByVal x As Variant) As Variant
    Dim s As String, p As Long
    If IsError(x) Or IsNull(x) Or IsEmpty(x) Then NomPlan = Null: Exit Function
    s = Replace(CStr(x), "\", "/")
    p = InStrRev(s, "/")
    If p > 0 Then s = Mid$(s, p + 1)
    If Len(s) = 0 Then NomPlan = Null Else NomPlan = Left$(s, 255)
End Function

Private Sub Ajouter(ByVal niv As Integer, ByVal arbo As String, ByVal qte As Variant, _
                    ByVal ref As Variant, ByVal rev As Variant, ByVal descr As Variant, ByVal plan As Variant)
    mNb = mNb + 1
    If mNb > UBound(mLignes) Then ReDim Preserve mLignes(1 To UBound(mLignes) + 50000)
    With mLignes(mNb)
        .Niveau = niv
        .Arbo = arbo
        .Qte = qte
        .Ref = ref
        .Rev = rev
        .Descr = descr
        .Plan = plan
    End With
End Sub

'-------------------------------------------------------------------------------
' Écriture (directement dans la base qui contient la table, avec sauvegarde)
'-------------------------------------------------------------------------------
Private Sub EcrireTable()
    Dim dbT As DAO.Database
    Dim rs As DAO.Recordset
    Dim i As Long
    Dim etape As Integer
    Dim chemin As String

    On Error GoTo Erreur
    DBEngine.SetOption dbMaxLocksPerFile, 500000

    chemin = CheminBaseSynerlink()
    If Len(chemin) = 0 Then
        Set dbT = CurrentDb                         ' table locale
    Else
        Set dbT = DBEngine.OpenDatabase(chemin)     ' table liée : on écrit dans la base de données
    End If

    ' Sauvegarde de la table actuelle (dans la même base)
    SysCmd acSysCmdSetStatus, "Sauvegarde de " & TABLE_ARBO & "..."
    If TableExisteDans(dbT, TABLE_SAUV) Then
        dbT.Execute "DELETE * FROM " & TABLE_SAUV, dbFailOnError
        dbT.Execute "INSERT INTO " & TABLE_SAUV & " SELECT * FROM " & TABLE_ARBO, dbFailOnError
    Else
        dbT.Execute "SELECT * INTO " & TABLE_SAUV & " FROM " & TABLE_ARBO, dbFailOnError
        If Len(chemin) > 0 And Not TableExiste(TABLE_SAUV) Then
            DoCmd.TransferDatabase acLink, "Microsoft Access", chemin, acTable, TABLE_SAUV, TABLE_SAUV
        End If
    End If
    etape = 1

    ' Vidage
    dbT.Execute "DELETE * FROM " & TABLE_ARBO, dbFailOnError
    etape = 2

    ' Insertion
    Set rs = dbT.OpenRecordset(TABLE_ARBO, dbOpenTable)
    For i = 1 To mNb
        If i Mod 2000 = 0 Then SysCmd acSysCmdSetStatus, "Écriture " & Format(i, "#,##0") & " / " & Format(mNb, "#,##0")
        With mLignes(i)
            rs.AddNew
            rs!ID = i                ' ID = 1, 2, 3... dans l'ordre de l'arborescence
            If i < mNb Then
                rs!Treeview = (mLignes(i + 1).Niveau > .Niveau)
            Else
                rs!Treeview = False
            End If
            rs!Niveau = .Niveau
            rs!Arborescence = .Arbo
            rs!Quantite = .Qte
            rs!Reference = .Ref
            rs!Revision = .Rev
            rs!Description = .Descr
            rs!Plan = .Plan
            rs.Update
        End With
    Next i
    rs.Close
    Set rs = Nothing
    If Len(chemin) > 0 Then dbT.Close
    Set dbT = Nothing
    Exit Sub

Erreur:
    Dim msg As String
    msg = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If etape = 2 Then
        ' Restauration depuis la sauvegarde
        dbT.Execute "DELETE * FROM " & TABLE_ARBO
        dbT.Execute "INSERT INTO " & TABLE_ARBO & " SELECT * FROM " & TABLE_SAUV
    End If
    If Len(chemin) > 0 Then dbT.Close
    On Error GoTo 0
    Err.Raise vbObjectError + 3, , msg
End Sub

'===============================================================================
' COMPACTAGE de la base qui contient les tables Synerlink
'===============================================================================
' Renvoie Vrai si la base a été compactée, Faux si elle est utilisée par
' quelqu'un (on réessaiera plus tard) ou si la table est locale.
' ATTENTION : les formulaires ouverts sur ces tables doivent être libérés avant.
Public Function SynerlinkCompacter() As Boolean
    Dim chemin As String
    chemin = CheminBaseSynerlink()
    If Len(chemin) = 0 Then Exit Function
    SysCmd acSysCmdSetStatus, "Compactage de " & Mid$(chemin, InStrRev(chemin, "\") + 1) & "..."
    SynerlinkCompacter = CompacterBase(chemin)
    SysCmd acSysCmdClearStatus
End Function

Public Function CompacterBase(ByVal chemin As String) As Boolean
    Dim db As DAO.Database
    Dim racine As String, tmp As String, sauv As String

    racine = Left$(chemin, InStrRev(chemin, ".") - 1)
    tmp = racine & "_tmp.accdb"
    sauv = racine & "_avant_compactage.accdb"

    ' Test d'accès exclusif : échoue si quelqu'un a la base ouverte
    On Error Resume Next
    Set db = DBEngine.OpenDatabase(chemin, True)
    If Err.Number <> 0 Then Exit Function
    db.Close: Set db = Nothing
    On Error GoTo Erreur

    If Len(Dir(tmp)) > 0 Then Kill tmp
    DBEngine.CompactDatabase chemin, tmp        ' compacte vers un nouveau fichier
    If Len(Dir(sauv)) > 0 Then Kill sauv
    Name chemin As sauv                         ' garde la version précédente
    Name tmp As chemin                          ' met la version compactée à la place
    CompacterBase = True
    Exit Function

Erreur:
    On Error Resume Next
    If Len(Dir(chemin)) = 0 And Len(Dir(sauv)) > 0 Then Name sauv As chemin
    If Len(Dir(tmp)) > 0 Then Kill tmp
End Function

'===============================================================================
' Outils
'===============================================================================
' Chemin de la base qui contient T_Synerlink_Arbo ("" si la table est locale)
Public Function CheminBaseSynerlink() As String
    Dim cnx As String, p As Long
    cnx = CurrentDb.TableDefs(TABLE_ARBO).Connect
    p = InStr(1, cnx, "DATABASE=", vbTextCompare)
    If p > 0 Then
        cnx = Mid$(cnx, p + 9)
        If InStr(cnx, ";") > 0 Then cnx = Left$(cnx, InStr(cnx, ";") - 1)
        CheminBaseSynerlink = cnx
    End If
End Function

' Table de paramètres : lignes (niveau 1) et machines (niveau 2), dans l'ordre
Private Sub CreerTableParametres()
    Dim dbT As DAO.Database, chemin As String
    chemin = CheminBaseSynerlink()
    If Len(chemin) = 0 Then Set dbT = CurrentDb Else Set dbT = DBEngine.OpenDatabase(chemin)

    dbT.Execute "CREATE TABLE " & TABLE_PARAM & " (Ordre LONG CONSTRAINT PK PRIMARY KEY, " & _
                "Niveau SHORT, Reference TEXT(50), Libelle TEXT(255), Actif YESNO)", dbFailOnError
    ' Initialisation avec les lignes/machines présentes aujourd'hui dans T_Synerlink_Arbo
    dbT.Execute "INSERT INTO " & TABLE_PARAM & " (Ordre, Niveau, Reference, Libelle, Actif) " & _
                "SELECT ID, Niveau, Reference, Description, True FROM " & TABLE_ARBO & _
                " WHERE Niveau <= 2", dbFailOnError
    If Len(chemin) > 0 Then
        dbT.Close
        DoCmd.TransferDatabase acLink, "Microsoft Access", chemin, acTable, TABLE_PARAM, TABLE_PARAM
    End If
    Application.RefreshDatabaseWindow
End Sub

Private Sub ControlerFichiers(ByRef manquants As String, ByRef nouveaux As String)
    Dim rs As DAO.Recordset
    Dim f As String, connus As String

    Set rs = CurrentDb.OpenRecordset("SELECT Reference, Actif FROM " & TABLE_PARAM & " WHERE Niveau = 2", dbOpenSnapshot)
    Do Until rs.EOF
        connus = connus & "|" & UCase$(rs!Reference) & "|"
        If rs!Actif Then
            If Len(Dir(DOSSIER_BOM & rs!Reference & SUFFIXE_FICHIER)) = 0 Then
                manquants = manquants & "- " & rs!Reference & SUFFIXE_FICHIER & vbCrLf
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close

    f = Dir(DOSSIER_BOM & "*" & SUFFIXE_FICHIER)
    Do While Len(f) > 0
        If Left$(f, 2) <> "~$" Then     ' fichiers temporaires d'Excel ouverts
            If InStr(connus, "|" & UCase$(Left$(f, Len(f) - Len(SUFFIXE_FICHIER))) & "|") = 0 Then
                nouveaux = nouveaux & "- " & f & vbCrLf
            End If
        End If
        f = Dir()
    Loop
End Sub

Private Function TableExiste(ByVal nom As String) As Boolean
    TableExiste = TableExisteDans(CurrentDb, nom)
End Function

Private Function TableExisteDans(db As DAO.Database, ByVal nom As String) As Boolean
    Dim t As DAO.TableDef
    For Each t In db.TableDefs
        If t.Name = nom Then TableExisteDans = True: Exit Function
    Next t
End Function
