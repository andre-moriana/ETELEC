VERSION 4.00
Begin VB.Form Finterv 
   AutoRedraw      =   -1  'True
   Caption         =   "Fichier interventions"
   ClientHeight    =   6240
   ClientLeft      =   300
   ClientTop       =   735
   ClientWidth     =   8940
   FillColor       =   &H00C00000&
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   0
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Height          =   6930
   Left            =   240
   LinkTopic       =   "Feuille1"
   LockControls    =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   8940
   Top             =   105
   Width           =   9060
   WindowState     =   2  'Maximized
   Begin VB.TextBox Texte11 
      Height          =   1065
      Left            =   1155
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   4620
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.Frame Frame1 
      Caption         =   "Planning du "
      Height          =   2850
      Left            =   5040
      TabIndex        =   14
      Top             =   105
      Width           =   4530
      Begin VB.Data Donnée4 
         Caption         =   "Data4"
         Connect         =   "dBASE IV;"
         DatabaseName    =   "C:\GESTION"
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2625
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "INTERV"
         Top             =   0
         Visible         =   0   'False
         Width           =   1590
      End
      Begin MSGrid.Grid Grid1 
         Height          =   2640
         Left            =   0
         TabIndex        =   15
         Top             =   210
         Width           =   4530
         _Version        =   65536
         _ExtentX        =   7990
         _ExtentY        =   4657
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Rows            =   11
         Cols            =   6
         ScrollBars      =   0
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   2850
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   4530
         _Version        =   458752
         _ExtentX        =   7990
         _ExtentY        =   5027
         _StockProps     =   1
         BackColor       =   12632256
         Year            =   1997
         Month           =   1
         Day             =   11
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         DayFontColor    =   -2147483641
         FirstDay        =   2
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         GridFontColor   =   0
         GridLinesColor  =   0
      End
   End
   Begin VB.Data Donnée1 
      Appearance      =   0  'Flat
      Caption         =   "Donnée1"
      Connect         =   "dbase iv"
      DatabaseName    =   "c:\gestion"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   3360
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CLIENT"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Donnée2 
      Appearance      =   0  'Flat
      Connect         =   "dbase iv"
      DatabaseName    =   "c:\gestion"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5670
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "copro"
      Top             =   0
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox Texte7 
      DataField       =   "NCOPRO"
      DataSource      =   "baseint"
      Height          =   300
      Left            =   1470
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   840
      Width           =   2430
   End
   Begin VB.TextBox Texte6 
      DataField       =   "CCLIENT"
      DataSource      =   "baseint"
      Height          =   300
      Left            =   1470
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   420
      Width           =   2445
   End
   Begin VB.TextBox Texte4 
      DataField       =   "HEUREINT"
      DataSource      =   "baseint"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4200
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1260
      Width           =   750
   End
   Begin VB.TextBox Texte3 
      BackColor       =   &H00FFFFFF&
      DataField       =   "DATEINT"
      DataSource      =   "baseint"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2730
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1230
   End
   Begin VB.TextBox Texte2 
      DataField       =   "typeint"
      DataSource      =   "baseint"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1995
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   105
      Width           =   750
   End
   Begin VB.TextBox Texte1 
      DataField       =   "descro"
      DataSource      =   "baseint"
      Height          =   1275
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1680
      Width           =   4845
   End
   Begin VB.CommandButton Commande2 
      Caption         =   "COPRO."
      Height          =   330
      Left            =   3990
      TabIndex        =   1
      Top             =   840
      Width           =   915
   End
   Begin VB.CommandButton Commande1 
      Caption         =   "CLIENT"
      Height          =   345
      Left            =   3990
      TabIndex        =   0
      Top             =   420
      Width           =   915
   End
   Begin VB.Data baseint 
      Connect         =   "dbase iv"
      DatabaseName    =   "C:\gestion"
      Exclusive       =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4515
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "INTERV"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   105
      Top             =   105
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Copies          =   1
      DefaultExt      =   ".dbf"
      DialogTitle     =   "Ouvrir une base de donnée"
      FileName        =   "interv*.dbf"
      Filter          =   "dbase (*.dbf)|*.dbf"
      FilterIndex     =   1
      InitDir         =   "c:\gestion"
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "INTERV.frx":0000
      Height          =   3480
      Left            =   105
      OleObjectBlob   =   "INTERV.frx":0010
      TabIndex        =   13
      Top             =   3045
      Width           =   9465
   End
   Begin VB.Label Etiquette6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nom Copro."
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   11
      Top             =   840
      Width           =   1410
   End
   Begin VB.Label Etiquette5 
      BackStyle       =   0  'Transparent
      Caption         =   "Code CLient"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   105
      TabIndex        =   9
      Top             =   420
      Width           =   1365
   End
   Begin VB.Label Etiquette4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date et heure d'intervention"
      Height          =   240
      Left            =   105
      TabIndex        =   6
      Top             =   1260
      Width           =   2685
   End
   Begin VB.Label Etiquette3 
      BackStyle       =   0  'Transparent
      Caption         =   "Type d'intervention"
      Height          =   255
      Left            =   105
      TabIndex        =   4
      Top             =   105
      Width           =   1815
   End
   Begin VB.Label Etiquette2 
      BackStyle       =   0  'Transparent
      Caption         =   "Descriptif de l'intervention"
      Height          =   255
      Left            =   105
      TabIndex        =   3
      Top             =   1470
      Width           =   3015
   End
   Begin VB.Menu m_fichier 
      Caption         =   "&Fichiers"
      Begin VB.Menu M_F_OUVRIR 
         Caption         =   "&Ouvrir"
      End
      Begin VB.Menu M_CONFIMP 
         Caption         =   "&Configuration imprimante"
      End
      Begin VB.Menu m4 
         Caption         =   "-"
      End
      Begin VB.Menu M_Quit 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu M_RECH 
      Caption         =   "&Recherche"
      Begin VB.Menu M_R_NUMBA 
         Caption         =   "par NUMERO de &B.A."
      End
      Begin VB.Menu M_R_NUM 
         Caption         =   "par NUMERO de &FICHE"
      End
   End
   Begin VB.Menu M_TRI 
      Caption         =   "&Tri"
      Begin VB.Menu M_T_NUM 
         Caption         =   "&Numero de fiches"
         Checked         =   -1  'True
      End
      Begin VB.Menu M_T_CCLIENT 
         Caption         =   "&Code Client"
      End
      Begin VB.Menu M_T_TYPE 
         Caption         =   "&Type d'Intervention"
      End
   End
   Begin VB.Menu M_SELECTION 
      Caption         =   "&Selection"
      Begin VB.Menu M_S_CCLIENT 
         Caption         =   "d'un CODE &CLIENT"
      End
      Begin VB.Menu M_S_NCOPRO 
         Caption         =   "d'une C&OPRO"
      End
      Begin VB.Menu M_S_DATEAP 
         Caption         =   "d'une &DATE d'appel"
      End
   End
   Begin VB.Menu M_P 
      Caption         =   "&Planning"
      Begin VB.Menu M_P_date 
         Caption         =   "&date"
      End
      Begin VB.Menu M_IMP_PLANING 
         Caption         =   "Imprimer"
      End
   End
End
Attribute VB_Name = "Finterv"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim nomf$, trif$
Dim VTtech(12) As String
Public vdateint$
Function EFFgrilleP()
For I = 1 To 5
VTtech(I) = ""
Next
For vcol = 1 To Grid1.Cols - 1
  For vlign = 1 To Grid1.Rows - 1
    Grid1.Row = vlign
    Grid1.Col = vcol
    Grid1.Text = ""
  Next vlign
Next vcol
End Function

Function MAJgrilleP()
   Grid1.Row = 0
   For I = 1 To Grid1.Cols - 1
     Grid1.Col = I
     Grid1.Text = "Tech." & Str(I)
   Next I
   'ecriture de la grille par technicien
   I = 1
   
 If Not Donnée4.Recordset.EOF Then
   Donnée4.Recordset.MoveFirst
   Do
     If Donnée4.Recordset("technicien") <> "" Then
     VTtech(I) = Donnée4.Recordset("technicien")
     Grid1.Col = I
     Grid1.Row = 0
     Grid1.Text = VTtech(I)
     vligne = 1
     Do
        Grid1.Col = I
    If Donnée4.Recordset("ordre") <> 0 Then
        Grid1.Row = Donnée4.Recordset("ordre")
        Grid1.Text = Donnée4.Recordset("numero")
    End If
   '     vligne = vligne + 1
        Donnée4.Recordset.MoveNext
        If Donnée4.Recordset.EOF Then Exit Do
     Loop While Donnée4.Recordset("technicien") = VTtech(I)
   I = I + 1
   Else
   Donnée4.Recordset.MoveNext
   End If
   Loop While Not Donnée4.Recordset.EOF
 End If
End Function

Function MAJinterv(maj%, vartech$, vardateint$)

Select Case maj%

Case 1
  If vartech$ = "" Then
    vartech$ = UCase(Left(Trim(InputBox("Entrez les initiales du technicien '3 caractères maxi.'", "Mise à jour rendez vous")), 3))
  End If
  baseint.ReadOnly = False
  baseint.Recordset.Edit
  baseint.Recordset("technicien") = vartech$
  baseint.Recordset("dateint") = Format(CDate(vardateint$), "mm-dd-yyyy")
  baseint.Recordset("ordre") = Grid1.Row
  baseint.Recordset.Update
  baseint.ReadOnly = True
  
Case 0
  rechnum = "NUMERO = " & Grid1.Text
  baseint.Recordset.FindFirst rechnum
  If Not baseint.Recordset.NoMatch Then
  baseint.ReadOnly = False
  baseint.Recordset.Edit
  baseint.Recordset("technicien") = ""
  baseint.Recordset("dateint") = Null
  baseint.Recordset("ordre") = Null
  baseint.Recordset.Update
  baseint.ReadOnly = True
  End If
  
End Select

End Function

Private Sub Calendar1_AfterUpdate()
    Me.MousePointer = 11
    Sélection = Calendar1.Value
    Texte$ = "SELECT INTERV.* FROM INTERV"
    Texte$ = Texte$ + " WHERE DATEINT = #" & Format(Sélection, "mm-dd-yyyy") & "# And numba = Null Or DATEINT = #" & Format(Sélection, "mm-dd-yyyy") & "# And numba = 0 ORDER BY TECHNICIEN"
    Donnée4.RecordSource = Texte$
    Donnée4.Refresh
    vdateint$ = Format(Sélection, "mm-dd-yyyy")
    EFFgrilleP
    MAJgrilleP
    Me.MousePointer = 1
    Calendar1.Visible = False
    Grid1.Visible = True
    Frame1.Caption = "Planing du " & Format(Sélection, "dddd d mmm yyyy")

End Sub

Private Sub Calendar1_DblClick()
Calendar1_AfterUpdate
End Sub

Private Sub Commande1_Click()
If baseint.Recordset.RecordCount > 0 Then
  If baseint.Recordset("cclient") <> "" Then
    varcclient$ = baseint.Recordset("cclient")
    Fclient2.Show 1
  End If
Else
 MsgBox "AUCUN CLIENT SELECTIONNE", 48, ""
End If
End Sub

Private Sub Commande2_Click()
  If baseint.Recordset.RecordCount > 0 Then
   If Trim(baseint.Recordset("ncopro")) <> "" Then
     varcclient$ = texte6.Text
     varncopro$ = texte7.Text
     fcopro.Show 1
   End If
  Else
     MsgBox "AUCUNE COPRO SELECTIONNEE", 48, ""
  End If
 
End Sub




Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TDescro As Object
If Button = 1 Then
   Select Case Shift
   Case 4
    On Error GoTo errorimp
    Printer.ScaleMode = 2
    taillep = 12
    Printer.FontSize = taillep
    Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print " "
    Printer.Print "NUMERO " & baseint.Recordset("numero") & "       date d'appel :" & baseint.Recordset("date_ap")
    Printer.FontSize = 20
    Printer.Print " "
    Printer.Print "                    TYPE D'INTERVENTION  : " + baseint.Recordset("typeint")
    Printer.FontSize = taillep
    Printer.Print " "
    Printer.Print "CODE CLIENT : " + baseint.Recordset("cclient")
    Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print " "
    Donnée1.Recordset.FindFirst "cclient= '" & baseint.Recordset("cclient") & "'"
    Printer.FontSize = 16
    Printer.FontBold = True
    Printer.Print Donnée1.Recordset("nclient");
    If Donnée1.Recordset("facc") Then
      Printer.Print "                                  CE"
    End If
    Printer.Print " "
    Printer.FontSize = taillep
    Printer.FontBold = False
    If Donnée1.Recordset("adrc1") <> "" Then Printer.Print Donnée1.Recordset("adrc1")
    If Donnée1.Recordset("adrc2") <> "" Then Printer.Print Donnée1.Recordset("adrc2")
    If Donnée1.Recordset("codepost") <> "" Then Printer.Print Donnée1.Recordset("codepost"), ;
    If Donnée1.Recordset("villec") <> "" Then Printer.Print Donnée1.Recordset("villec")
    Printer.Print "Contact " & Donnée1.Recordset("contact"), "Téléphone  :" & Donnée1.Recordset("ntelc")
    Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print " "
    If baseint.Recordset("ncopro") <> "" Then
      Printer.FontSize = 16
      Printer.FontBold = True
      Printer.Print baseint.Recordset("ncopro")
      Printer.FontSize = taillep
      Printer.FontBold = False
      donnée2.RecordSource = "select copro.* from copro where cclient = '" & baseint.Recordset("cclient") & "'"
      donnée2.Refresh
      donnée2.Recordset.FindFirst "ncopro like '" & baseint.Recordset("ncopro") & "*'"
      If donnée2.Recordset("adrs1") <> "" Then Printer.Print donnée2.Recordset("adrs1")
      If donnée2.Recordset("adrs2") <> "" Then Printer.Print donnée2.Recordset("adrs2")
      If donnée2.Recordset("cposts") <> "" Then Printer.Print donnée2.Recordset("cposts"), ;
      If donnée2.Recordset("villes") <> "" Then Printer.Print donnée2.Recordset("villes")
      Printer.Print "Contact " & donnée2.Recordset("conts"), "Téléphone  :" & donnée2.Recordset("ntels")
  '    printer.Print " "
      Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    End If
    Printer.Print " "
    longt = 1
    Set TDescro = texte11
    TDescro.Text = baseint.Recordset("descro")
 ' suppression des espaces en trod
    Do
    longt = InStr(TDescro, "  ")
    If longt > 0 Then
      TDescro.SelStart = longt
      TDescro.SelLength = 1
      TDescro.SelText = ""
    Else
      Exit Do
    End If
    Loop
    
    ltext = Len(TDescro.Text)
    post = 1
    ncar = InStr(post, TDescro.Text, " ")
    NcarC = ncar - post
    Lcar = NcarC * Printer.Font.Size
    Do
    If Printer.CurrentX < Printer.ScaleWidth - Lcar Then
      Printer.Print Mid(TDescro.Text, post, NcarC);
      post = post + NcarC
      ncar = InStr(post, TDescro.Text, " ")
      If ncar <> 0 Then
        NcarC = ncar - post + 1
        Lcar = NcarC * Printer.Font.Size
      Else
        If post < ltext Then Printer.Print Mid(TDescro.Text, post, ltext)
        Exit Do
      End If
    Else
      Printer.Print " "
    End If
    Loop
     Printer.Print " "
    Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print " "
    Printer.Print "Date d'intervention :";
    If baseint.Recordset("dateint") <> "" Then
      Printer.Print baseint.Recordset("dateint")
    Else
      Printer.Print " "
    End If
    Printer.Print "Heure :";
    If baseint.Recordset("heureint") <> "" Then
      Printer.Print baseint.Recordset("heureint")
    Else
      Printer.Print " "
    End If
    Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print " "
    Printer.EndDoc

   Case 2
     If nomf$ = "INTERV" Then
       varcclient$ = baseint.Recordset("cclient")
       varnumero% = baseint.Recordset("numero")
       FMAJ_INT.Show 1
       baseint.Refresh
     End If
   Case 1
   
   End Select
 End If
 Exit Sub
errorimp:
    msg = Error(Err)
    MsgBox msg
    Printer.KillDoc
    Exit Sub

End Sub


Private Sub Form_Load()


'ouverture du fichier intervention
   nomf$ = "INTERV"
   trif$ = "numero"
   baseint.RecordSource = "SELECT " & nomf$ & ".* FROM " & nomf$ & " WHERE numba = null or numba = 0  ORDER BY numero desc"
   baseint.Refresh
   
   Donnée4.RecordSource = "SELECT interv.* FROM INTERV WHERE dateint = #" & Format(Date, "mm-dd-yyyy") & "# and numba = null OR dateint = #" & Format(Date, "mm-dd-yyyy") & "# and numba = 0 order by technicien,ordre"
   Donnée4.Refresh
   Frame1.Caption = "Planing du " & Format(Now, "dddd d mmm yyyy")
' mise a jour de la grille
   Grid1.ColWidth(0) = 400
   For I = 1 To Grid1.Cols - 1
     Grid1.ColWidth(I) = 750
   Next I
   Grid1.Col = 0
   For I = 1 To Grid1.Rows - 1
     Grid1.Row = I
     Grid1.Text = Str(I)
   Next I
   vdateint$ = Format(Now, "mm-dd-yyyy")
   MAJgrilleP
End Sub

Private Sub Grid1_Click()
If Grid1.Text <> "" Then
  baseint.Recordset.FindFirst "numero = " & Grid1.Text
  DBGrid1.ReBind
End If
End Sub

Private Sub Grid1_DblClick()
If Grid1.Text = "" Then
  If baseint.Recordset("dateint") <> "" Then
    msg = "Il existe une date de rendez vous pour cette intervention." & Chr(10) & Chr(13)
    msg = "La date de rendez vous actuel est " & baseint.Recordset("dateint") & Chr(10) & Chr(13) & " ."
    msg = msg + "Confirmez vous la suppression de l'ancien rendez vous ?"
    rep = MsgBox(msg, 292, "Confirmer")
    If rep = 6 Then
      MAJinterv 1, VTtech(Grid1.Col), vdateint$
      Donnée4.Refresh
      EFFgrilleP
      MAJgrilleP
 '     DBGrid1.ReBind
    End If
  Else
    MAJinterv 1, VTtech(Grid1.Col), vdateint$
    Donnée4.Refresh
    EFFgrilleP
    MAJgrilleP
'    DBGrid1.ReBind
  End If
Else
  msg = "ATTENTION." & Chr(10) & Chr(13)
  msg = msg + "Confirmez vous la suppression du rendez vous ?"
  rep = MsgBox(msg, 292, "Confirmer")
  If rep = 6 Then
    MAJinterv 0, "", ""
    Donnée4.Refresh
    EFFgrilleP
    MAJgrilleP
'    DBGrid1.ReBind
  End If

End If
DBGrid1.ReBind
End Sub

Private Sub M_ARCINT94_Click()
Me.MousePointer = 11
  Me.Caption = "Fichier ARCHIVE INTERVENTION 94"
  M_finterv.Checked = False
  M_ARCINT94.Checked = True
  M_ARCINT95.Checked = False
  commande5.Enabled = False
  nomf$ = "ARCINT94"
  etiquette17.Caption = "Numero de Fiche"
  baseint.RecordSource = "SELECT " & nomf$ & ".* FROM " & nomf$ & " ORDER BY numero desc"
  baseint.Refresh
  baseint.Caption = baseint.Recordset.RecordCount & " fiches"
Me.MousePointer = 1
End Sub

Private Sub M_ARCINT95_Click()
Me.MousePointer = 11
  Me.Caption = "Fichier ARCHIVE INTERVENTION 95"
  M_finterv.Checked = False
  M_ARCINT95.Checked = True
  M_ARCINT94.Checked = False
  commande5.Enabled = False
  nomf$ = "ARCINT95"
  etiquette17.Caption = "Numero de Fiche"
  baseint.RecordSource = "SELECT " & nomf$ & ".* FROM " & nomf$ & " ORDER BY numero desc "
  baseint.Refresh
  baseint.Caption = baseint.Recordset.RecordCount & " fiches"
Me.MousePointer = 1
End Sub


Private Sub M_CONFIMP_Click()
  CommonDialog1.Flags = &H40
  CommonDialog1.Action = 5

End Sub

Private Sub M_F_OUVRIR_Click()
Me.MousePointer = 11
  CommonDialog1.ShowOpen
  nomf$ = CommonDialog1.FileTitle
  If nomf$ <> "" Then
    nomf$ = Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4)
    baseint.RecordSource = "SELECT " & nomf$ & ".* FROM " & nomf$
    baseint.Refresh
    Me.Caption = "Fichier INTERVENTION '" & nomf$ & "'"
  End If
  
Me.MousePointer = 1

End Sub


Private Sub M_IMP_PLANING_Click()
Dim TDescro As Object
On Error GoTo errorimp
Printer.ScaleMode = 2
Donnée4.Recordset.MoveFirst
taillep = Printer.FontSize
For I = 1 To Donnée4.Recordset.RecordCount

If Donnée4.Recordset("technicien") <> "" Then
    
    Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    
    Printer.Print " "
    Printer.Print "NUMERO " & Donnée4.Recordset("numero") & "       date d'appel :" & Donnée4.Recordset("date_ap") & "                 ";
    Printer.FontSize = 20
    Printer.Print Donnée4.Recordset("ordre") & "   " & Donnée4.Recordset("technicien")
    Printer.Print " "
    Printer.Print "                    TYPE D'INTERVENTION  : " + baseint.Recordset("typeint")
    Printer.FontSize = taillep
    Printer.Print " "
    Printer.Print "CODE CLIENT : " + Donnée4.Recordset("cclient")
    Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print " "
    Donnée1.Recordset.MoveFirst
    Donnée1.Recordset.FindFirst "CCLIENT = '" & Donnée4.Recordset("cclient") & "'"
    Printer.FontSize = 16
    Printer.FontBold = True
    Printer.Print Donnée1.Recordset("nclient");
    If Donnée1.Recordset("facc") Then
      Printer.Print "                                  CE"
    End If
    Printer.Print " "
    Printer.FontBold = False
    Printer.FontSize = taillep
    If Donnée1.Recordset("adrc1") <> "" Then Printer.Print Donnée1.Recordset("adrc1")
    If Donnée1.Recordset("adrc2") <> "" Then Printer.Print Donnée1.Recordset("adrc2")
    If Donnée1.Recordset("codepost") <> "" Then Printer.Print Donnée1.Recordset("codepost"), ;
    If Donnée1.Recordset("villec") <> "" Then Printer.Print Donnée1.Recordset("villec")
    Printer.Print "Contact " & Donnée1.Recordset("contact"), "Téléphone  :" & Donnée1.Recordset("ntelc")
    Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print " "
    If Donnée4.Recordset("ncopro") <> "" Then
      Printer.FontBold = True
      Printer.FontSize = 16
      Printer.Print Donnée4.Recordset("ncopro")
      Printer.FontBold = False
      Printer.FontSize = taillep
      donnée2.RecordSource = "select copro.* from copro where cclient = '" & Donnée4.Recordset("cclient") & "'"
      donnée2.Refresh
      donnée2.Recordset.FindFirst "ncopro like '" & Donnée4.Recordset("ncopro") & "*'"
      If donnée2.Recordset("adrs1") <> "" Then Printer.Print donnée2.Recordset("adrs1")
      If donnée2.Recordset("adrs2") <> "" Then Printer.Print donnée2.Recordset("adrs2")
      If donnée2.Recordset("cposts") <> "" Then Printer.Print donnée2.Recordset("cposts"), ;
      If donnée2.Recordset("villes") <> "" Then Printer.Print donnée2.Recordset("villes")
      Printer.Print "Contact " & donnée2.Recordset("conts"), "Téléphone  :" & donnée2.Recordset("ntels")
      Printer.Print " "
      Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    End If
    Printer.Print " "
    longt = 1
    Set TDescro = texte11
    TDescro.Text = Donnée4.Recordset("descro")
 ' suppression des espaces en trod
    Do
    longt = InStr(TDescro, "  ")
    If longt > 0 Then
      TDescro.SelStart = longt
      TDescro.SelLength = 1
      TDescro.SelText = ""
    Else
      Exit Do
    End If
    Loop
    
    ltext = Len(TDescro.Text)
    post = 1
    ncar = InStr(post, TDescro.Text, " ")
    NcarC = ncar - post
    Lcar = NcarC * Printer.Font.Size
    Do
    If Printer.CurrentX < Printer.ScaleWidth - Lcar Then
      Printer.Print Mid(TDescro.Text, post, NcarC);
      post = post + NcarC
      ncar = InStr(post, TDescro.Text, " ")
      If ncar <> 0 Then
        NcarC = ncar - post + 1
        Lcar = NcarC * Printer.Font.Size
      Else
        If post < ltext Then Printer.Print Mid(TDescro.Text, post, ltext)
        Exit Do
      End If
    Else
      Printer.Print " "
    End If
    Loop
    Printer.Print " "
    Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print " "
    Printer.Print "Date d'intervention :";
    If Donnée4.Recordset("dateint") <> "" Then
      Printer.Print Donnée4.Recordset("dateint")
    Else
      Printer.Print " "
    End If
    Printer.Print "Heure :";
    If Donnée4.Recordset("heureint") <> "" Then
      Printer.Print Donnée4.Recordset("heureint")
    Else
      Printer.Print " "
    End If
    Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print " "
    Printer.NewPage

End If
Donnée4.Recordset.MoveNext
Next
Printer.EndDoc
 Exit Sub
errorimp:
    msg = Error(Err)
    MsgBox msg
    Printer.KillDoc
    Exit Sub

End Sub

Private Sub M_P_date_Click()
Grid1.Visible = False
Calendar1.Value = Date
Calendar1.Visible = True
End Sub

Private Sub M_PURGER_Click()
End Sub

Private Sub M_quit_Click()
  Unload Finterv
End Sub

Private Sub M_R_NUM_Click()
  Me.MousePointer = 11
  baseint.RecordSource = "Select " & nomf$ & ".* from " & nomf$ & " order by numero"
  baseint.Refresh
  baseint.Caption = baseint.Recordset.RecordCount
  Me.MousePointer = 1
  rep$ = InputBox("Entrer le numero de la fiche ")
  If Trim(rep$) <> "" Then
    Me.MousePointer = 11
    baseint.Recordset.FindFirst "NUMERO = " & rep$
    If baseint.Recordset.NoMatch Then MsgBox ("FICHE INEXISTANTE")
    Me.MousePointer = 1
  End If

End Sub

Private Sub M_R_NUMBA_Click()
  Me.MousePointer = 11
  baseint.RecordSource = "Select " & nomf$ & ".* from " & nomf$ & " order by numba"
  baseint.Refresh
  baseint.Caption = baseint.Recordset.RecordCount
  Me.MousePointer = 1
  rep$ = InputBox("Entrer le numero du B.A. ")
  If Trim(rep$) <> "" Then
    Me.MousePointer = 11
    baseint.Recordset.FindFirst "NUMBA = " & rep$
    If baseint.Recordset.NoMatch Then MsgBox ("FICHE INEXISTANTE")
    Me.MousePointer = 1
  End If
End Sub

Private Sub M_S_CCLIENT_Click()
  
  VCclient$ = ""
  If baseint.Recordset.RecordCount > 0 Then
    VCclient$ = baseint.Recordset("CCLIENT")
  End If
  Sélection$ = InputBox("Entrer le code client ", , VCclient$)
 If Sélection$ <> "" Then
  Me.MousePointer = 11
  baseint.RecordSource = "SELECT " & nomf$ & ".* FROM " & nomf$ & " WHERE NUMBA = null and CCLIENT = '" & Sélection$ & "' or NUMBA = 0 and CCLIENT = '" & Sélection$ & "' ORDER BY " & trif$
  baseint.Refresh
  Me.MousePointer = 1
 End If
End Sub

Private Sub M_S_DATEAP_Click()
  Sélection = InputBox("Entrer la DATE de sélection ", , Format$(Now, "dd-mm-yyyy"))
  If Sélection <> "" Then
  If IsDate(Sélection) Then
    Me.MousePointer = 11
    Sélection = Mid(Sélection, 4, 3) + Left(Sélection, 3) + Right(Sélection, 4)
    Texte$ = "SELECT " & nomf$ & ".* FROM " & nomf$
    Texte$ = Texte$ + " WHERE  NUMBA = null and DATE_AP = #" & Sélection & "# or NUMBA = 0 and DATE_AP = #" & Sélection & "# ORDER BY " & trif$
 '   etiquette18.Caption = Sélection
    baseint.RecordSource = Texte$
    baseint.Refresh
    baseint.Caption = baseint.Recordset.RecordCount & " fiches"
    Me.MousePointer = 1
  Else
    MsgBox ("Erreur dans le format de DATE")
  End If
  End If
End Sub

Private Sub M_S_NCOPRO_Click()
  VCclient$ = ""
  If baseint.Recordset.RecordCount > 0 Then
    If baseint.Recordset("ncopro") <> "" Then VCclient$ = baseint.Recordset("ncopro")
  End If
  Sélection$ = UCase(InputBox("Entrer le nom de la COPRO ", , VCclient$))
  pos% = InStr(Sélection$, "'")
  If pos% <> 0 Then Mid$(Sélection$, pos%, 1) = "?"
 If Sélection$ <> "" Then
  Me.MousePointer = 11
 ' etiquette18.Caption = Sélection$
  baseint.RecordSource = "SELECT " & nomf$ & ".* FROM " & nomf$ & " WHERE NUMBA = null and NCOPRO LIKE'" & Sélection$ & "' or NUMBA = 0 and NCOPRO LIKE'" & Sélection$ & "'  order by " & trif$
  baseint.Refresh
  baseint.Caption = baseint.Recordset.RecordCount & " fiches"
  Me.MousePointer = 1
 End If
End Sub

Private Sub M_T_CCLIENT_Click()
 If M_T_CCLIENT.Checked = False Then
  Me.MousePointer = 11
  M_T_CCLIENT.Checked = True
  M_T_NUM.Checked = False
  M_T_TYPE.Checked = False
  trif$ = "CCLIENT"
  baseint.RecordSource = "SELECT " & nomf$ & ".* FROM " & nomf$ & " where NUMBA = null or NUMBA = 0 ORDER BY " & trif$
  baseint.Refresh
  Me.MousePointer = 1
End If

End Sub

Private Sub M_T_NUM_Click()
  Me.MousePointer = 11
  M_T_CCLIENT.Checked = False
  M_T_NUM.Checked = True
  M_T_TYPE.Checked = False
  trif$ = "NUMERO DESC"
  baseint.RecordSource = "SELECT " & nomf$ & ".* FROM " & nomf$ & " where NUMBA = null or NUMBA = 0 ORDER BY " & trif$
  baseint.Refresh
  Me.MousePointer = 1

End Sub

Private Sub M_T_TYPE_Click()
If M_T_TYPE.Checked = False Then
  Me.MousePointer = 11
  M_T_CCLIENT.Checked = False
  M_T_NUM.Checked = False
  M_T_TYPE.Checked = True
  trif$ = "TYPEINT"
  baseint.RecordSource = "SELECT " & nomf$ & ".* FROM " & nomf$ & " where NUMBA = null or NUMBA = 0 ORDER BY " & trif$
  baseint.Refresh
  Me.MousePointer = 1
End If

End Sub


