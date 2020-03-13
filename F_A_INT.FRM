VERSION 4.00
Begin VB.Form F_A_INT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AJOUT FICHE INTERVENTION "
   ClientHeight    =   4245
   ClientLeft      =   180
   ClientTop       =   1560
   ClientWidth     =   7725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   0
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   4650
   Left            =   120
   LinkTopic       =   "Feuille1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7725
   Top             =   1215
   Width           =   7845
   Begin VB.CommandButton Command1 
      Caption         =   "Modification copropriété"
      Height          =   435
      Left            =   5880
      TabIndex        =   14
      Top             =   315
      Width           =   1485
   End
   Begin VB.TextBox Texte3 
      Height          =   330
      Left            =   2625
      TabIndex        =   4
      Top             =   3255
      Width           =   960
   End
   Begin VB.TextBox Texte2 
      Height          =   330
      Left            =   2625
      TabIndex        =   3
      Top             =   2835
      Width           =   1590
   End
   Begin VB.Data Donnée2 
      Appearance      =   0  'Flat
      Caption         =   "Donnée2"
      Connect         =   "dbase iv"
      DatabaseName    =   "c:\gestion"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   5460
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "INTERV"
      Top             =   3960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Donnée1 
      Appearance      =   0  'Flat
      Caption         =   "Donnée1"
      Connect         =   "dbase iv"
      DatabaseName    =   "c:\gestion"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   5460
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "COPRO"
      Top             =   3675
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.CommandButton Commande1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "RETOUR"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5880
      TabIndex        =   6
      Top             =   3000
      Width           =   1365
   End
   Begin VB.TextBox Texte4 
      Alignment       =   2  'Center
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2625
      MaxLength       =   3
      TabIndex        =   5
      Top             =   3675
      Width           =   675
   End
   Begin VB.ComboBox Liste1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "F_A_INT.frx":0000
      Left            =   2310
      List            =   "F_A_INT.frx":0022
      TabIndex        =   0
      Top             =   105
      Width           =   3180
   End
   Begin VB.TextBox Texte1 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   135
      MaxLength       =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1260
      Width           =   7440
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "F_A_INT.frx":00F4
      Height          =   315
      Left            =   2310
      TabIndex        =   1
      Top             =   525
      Width           =   3165
      _Version        =   65536
      _ExtentX        =   5583
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      ListField       =   "NCOPRO"
      BoundColumn     =   "NCOPRO"
   End
   Begin VB.Label Etiquette7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caractère  0/240"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5670
      TabIndex        =   13
      Top             =   945
      Width           =   1815
   End
   Begin VB.Label Etiquette6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Technicien"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   840
      TabIndex        =   12
      Top             =   3675
      Width           =   1650
   End
   Begin VB.Label Etiquette5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Heure de rendez-vous"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   105
      TabIndex        =   11
      Top             =   3255
      Width           =   2385
   End
   Begin VB.Label Etiquette4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date de rendez-vous"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   105
      TabIndex        =   10
      Top             =   2835
      Width           =   2355
   End
   Begin VB.Label Etiquette3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nature de l'appel"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   105
      TabIndex        =   9
      Top             =   945
      Width           =   2160
   End
   Begin VB.Label Etiquette2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copropriété"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   840
      TabIndex        =   8
      Top             =   525
      Width           =   1380
   End
   Begin VB.Label Etiquette1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type d'intervention"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -210
      TabIndex        =   7
      Top             =   105
      Width           =   2385
   End
End
Attribute VB_Name = "F_A_INT"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
If DBCombo1.MatchedWithList Then
  varncopro$ = DBCombo1.Text
  fcopro.Show 1

Else
  'Creation de la copro si ncopro est pas dans la liste, ni vide
  If DBCombo1.Text <> "" Then
    varncopro$ = LTrim(UCase(DBCombo1.Text))
    fcopro.Show 1
    Donnée1.Refresh
  End If
End If

End Sub

Private Sub Commande1_Click()
    rep = MsgBox("Voulez vous enregistrer ? ", 4)
    If rep = 6 Then
      numfile% = FreeFile
      Open "c:\gestion\numero.mem" For Binary As numfile%
      Get numfile%, 33, numero%
      numero% = numero% + 1
      Put numfile%, 33, numero%
      Close numfile%
      donnée2.Recordset.AddNew
      donnée2.Recordset("numero") = numero%
      donnée2.Recordset("cclient") = varcclient$  'Fclient!Donnée1.Recordset("cclient")
      donnée2.Recordset("DATE_AP") = Format$(Now, "dd/mm/yyyy")
      donnée2.Recordset("typeint") = Left$(liste1.Text, 5)
      donnée2.Recordset("ncopro") = DBCombo1.Text
      donnée2.Recordset("descro") = texte1.Text
      If texte2.Text <> "" Then donnée2.Recordset("dateint") = texte2.Text
      donnée2.Recordset("heureint") = texte3.Text
      donnée2.Recordset("technicien") = texte4.Text
      donnée2.Recordset.Update
    End If
Unload F_A_INT
End Sub

Private Sub DBCombo1_LostFocus()
If Not DBCombo1.MatchedWithList And DBCombo1.Text <> "" Then
  varncopro$ = LTrim(UCase(DBCombo1.Text))
  fcopro.Show 1
  Donnée1.Refresh
    If varncopro$ <> "" Then
      DBCombo1.Text = varncopro$
    Else
      DBCombo1.SelStart = 0
      DBCombo1.SelLength = Len(DBCombo1.Text)
      DBCombo1.SetFocus
    End If
End If
End Sub


Private Sub Form_Load()
   
   Donnée1.RecordSource = "SELECT copro.* FROM copro WHERE cclient = '" & varcclient$ & "' ORDER BY ncopro"
   Donnée1.Refresh

End Sub

Private Sub Liste2_KeyPress(keyascii As Integer)
If keyascii = 13 Then
  varncopro$ = liste2.Text
  fcopro.Show 1
  Donnée1.Refresh
  liste2.Clear
  Do Until Donnée1.Recordset.EOF
    liste2.AddItem Donnée1.Recordset("ncopro")
    Donnée1.Recordset.MoveNext
  Loop
End If
End Sub

Private Sub liste2_LostFocus()
 liste2.Text = UCase(liste2.Text)
End Sub

Private Sub texte1_Change()
etiquette7.Caption = "Caractère  " & Len(texte1.Text) & "/240"
End Sub

Private Sub Texte1_LostFocus()
  texte1.Text = UCase(texte1.Text)
End Sub

Private Sub texte2_LostFocus()
If texte2.Text <> "" Then
  If Not (IsDate(texte2.Text)) Then
    texte2.SelStart = 0
    texte2.SelLength = Len(texte2.Text)
    texte2.SetFocus
  End If
End If

End Sub

Private Sub Texte3_LostFocus()
  texte3.Text = UCase(texte3.Text)
End Sub

Private Sub Texte4_LostFocus()
  texte4.Text = UCase(texte4.Text)
End Sub

