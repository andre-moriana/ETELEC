VERSION 4.00
Begin VB.Form LCopro 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6390
   ClientLeft      =   705
   ClientTop       =   705
   ClientWidth     =   8400
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   0
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   Height          =   7080
   Left            =   645
   LinkTopic       =   "Feuille1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   Top             =   75
   Width           =   8520
   Begin VB.TextBox Text2 
      DataField       =   "MEMO"
      DataSource      =   "Donnée1"
      Enabled         =   0   'False
      Height          =   960
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1995
      Width           =   3795
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Client"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   105
      TabIndex        =   16
      Top             =   0
      Width           =   3795
      Begin VB.CommandButton Commande1 
         Caption         =   "CLIENT"
         Height          =   255
         Left            =   105
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         DataField       =   "CCLIENT"
         DataSource      =   "Donnée1"
         Height          =   225
         Left            =   1050
         TabIndex        =   19
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "NCLIENT"
         DataSource      =   "Donnée2"
         Height          =   540
         Left            =   1260
         TabIndex        =   18
         Top             =   315
         Width           =   2325
      End
   End
   Begin VB.TextBox Text1 
      DataField       =   "NCOPRO"
      DataSource      =   "Donnée1"
      Enabled         =   0   'False
      Height          =   540
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1260
      Width           =   3795
   End
   Begin VB.Data Donnée2 
      Connect         =   "dBASE IV;"
      DatabaseName    =   "C:\GESTION"
      Exclusive       =   0   'False
      Height          =   330
      Left            =   7140
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CLIENT"
      Top             =   5565
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Cadre1 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   2745
      Left            =   3990
      TabIndex        =   9
      Top             =   105
      Width           =   4320
      Begin VB.TextBox Texte8 
         Appearance      =   0  'Flat
         DataField       =   "NTELS"
         DataSource      =   "Donnée1"
         Height          =   285
         Left            =   1365
         MaxLength       =   16
         TabIndex        =   6
         Top             =   2310
         Width           =   2430
      End
      Begin VB.TextBox Texte7 
         Appearance      =   0  'Flat
         DataField       =   "CONTS"
         DataSource      =   "Donnée1"
         Height          =   285
         Left            =   1365
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1995
         Width           =   2430
      End
      Begin VB.TextBox Texte6 
         Appearance      =   0  'Flat
         DataField       =   "CPOSTS"
         DataSource      =   "Donnée1"
         Height          =   285
         Left            =   1365
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1575
         Width           =   750
      End
      Begin VB.TextBox Texte5 
         Appearance      =   0  'Flat
         DataField       =   "VILLES"
         DataSource      =   "Donnée1"
         Height          =   285
         Left            =   840
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1155
         Width           =   3375
      End
      Begin VB.TextBox Texte4 
         Appearance      =   0  'Flat
         DataField       =   "ADRS2"
         DataSource      =   "Donnée1"
         Height          =   285
         Left            =   210
         MaxLength       =   40
         TabIndex        =   2
         Top             =   735
         Width           =   4005
      End
      Begin VB.TextBox Texte3 
         Appearance      =   0  'Flat
         DataField       =   "ADRS1"
         DataSource      =   "Donnée1"
         Height          =   285
         Left            =   210
         MaxLength       =   40
         TabIndex        =   1
         Top             =   420
         Width           =   4005
      End
      Begin VB.Label Etiquette3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Adresse"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   210
         Width           =   1380
      End
      Begin VB.Label Etiquette7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero tel."
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   210
         TabIndex        =   10
         Top             =   2310
         Width           =   1380
      End
      Begin VB.Label Etiquette6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   210
         TabIndex        =   11
         Top             =   1995
         Width           =   1065
      End
      Begin VB.Label Etiquette5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Code postal"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   210
         TabIndex        =   12
         Top             =   1575
         Width           =   1485
      End
      Begin VB.Label Etiquette4 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ville"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   210
         TabIndex        =   13
         Top             =   1155
         Width           =   750
      End
   End
   Begin VB.Data Donnée1 
      Connect         =   "dbase iv"
      DatabaseName    =   "c:\gestion"
      Enabled         =   0   'False
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
      Left            =   135
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "COPRO"
      Top             =   5460
      Visible         =   0   'False
      Width           =   4350
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "LCOPRO.frx":0000
      Height          =   3270
      Left            =   105
      OleObjectBlob   =   "LCOPRO.frx":0010
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3045
      Width           =   8205
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarque :"
      Height          =   225
      Left            =   105
      TabIndex        =   20
      Top             =   1785
      Width           =   1485
   End
   Begin VB.Label Etiquette2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom copropriété"
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
      Left            =   105
      TabIndex        =   14
      Top             =   945
      Width           =   1935
   End
   Begin VB.Menu M1 
      Caption         =   "&Fichier Copro"
      Begin VB.Menu M_ajout 
         Caption         =   "&Ajouter"
         Shortcut        =   ^A
      End
      Begin VB.Menu M_modif 
         Caption         =   "&Modifier"
         Shortcut        =   ^M
      End
      Begin VB.Menu M_supprim 
         Caption         =   "&Supprimer"
         Shortcut        =   ^S
      End
      Begin VB.Menu M_Enreg 
         Caption         =   "&Enregistrer"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu M_RECH 
         Caption         =   "&Recherche"
         Shortcut        =   ^R
      End
      Begin VB.Menu m2 
         Caption         =   "-"
      End
      Begin VB.Menu M_quit 
         Caption         =   "&Quitter"
      End
   End
End
Attribute VB_Name = "LCopro"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim ajout%


Private Sub Commande1_Click()
  varcclient$ = Donnée1.Recordset("cclient")
  Fclient2.Show 1
End Sub


Private Sub Donnée1_Reposition()
If Not ajout% Then
 varcclient = Donnée1.Recordset.Fields("cclient")
 donnée2.Recordset.FindFirst "CCLIENT = '" & varcclient & "'"
End If
End Sub


Private Sub Donnée1_Validate(Action As Integer, Save As Integer)
If Save Then
  rep = MsgBox("Voulez vous enregistrer ? ", 4)
  If rep = 6 Then
    Donnée1.Recordset("cclient") = varcclient
    Donnée1.Recordset("ncopro") = Text1.Text
    Donnée1.UpdateRecord
  Else
    Donnée1.Recordset.CancelUpdate
    Donnée1.UpdateControls
  End If
End If
End Sub




Private Sub Form_Load()
  ajout% = False
  Me.MousePointer = 11
  donnée2.Refresh
  Donnée1.RecordSource = "SELECT copro.* FROM copro order by ncopro "
  Donnée1.Refresh
  varcclient = Donnée1.Recordset.Fields("cclient")
  donnée2.Recordset.FindFirst "CCLIENT = '" & varcclient & "'"
  Me.MousePointer = 1
  
End Sub



Private Sub M_ajout_Click()
  ajout% = True
  commande1.Enabled = False
  cadre1.Enabled = True
  M_Ajout.Enabled = False
  M_modif.Enabled = False
  M_Supprim.Enabled = False
  M_Enreg.Enabled = True
  Text1.Enabled = True
  Donnée1.Recordset.AddNew
  Label2.Caption = varcclient

End Sub

Private Sub M_Enreg_Click()
  Me.MousePointer = 11
  Donnée1.Recordset.Update
  Donnée1.Refresh
  ajout% = False
  commande1.Enabled = False
  cadre1.Enabled = False
  M_Ajout.Enabled = True
  M_modif.Enabled = True
  M_Supprim.Enabled = True
  M_Enreg.Enabled = False
  M_RECH.Enabled = True
  Text1.Enabled = False
  Text2.Enabled = False
  commande1.Enabled = True
  Me.MousePointer = 1
  DBGrid1.ReBind
End Sub

Private Sub M_modif_Click()
  cadre1.Enabled = True
  Text2.Enabled = True
  M_modif.Enabled = False
  M_Ajout.Enabled = False
  M_Supprim.Enabled = False
  M_Enreg.Enabled = True
  M_RECH.Enabled = False
  Text1.Enabled = True
'  liste1.Enabled = False
'  donnée1.Enabled = False
  Donnée1.Recordset.Edit
End Sub

Private Sub M_quit_Click()
  Unload LCopro
End Sub

Private Sub M_RECH_Click()
    Donnée1.UpdateControls
    RNcopro$ = ""
    repere = Donnée1.Recordset.Bookmark

    RNcopro$ = UCase(InputBox("Entrez le NOM de la cpro.", "RECHERCHE"))
  If RNcopro$ <> "" Then
    Fclient.MousePointer = 11
    recherche$ = "NCOPRO like '*" & RNcopro$ & "*'"
    Donnée1.Recordset.FindFirst recherche$
    If Donnée1.Recordset.NoMatch Then Donnée1.Recordset.Bookmark = repere
  End If

End Sub

Private Sub M_supprim_Click()
  reponse = MsgBox("Etes vous sur de vouloir éffacer la fiche COPRO ?", 36 + 256, "ATTENTION")
  If reponse = 6 Then
    Me.MousePointer = 11
    Donnée1.Recordset.Delete
    Donnée1.Refresh
    Me.MousePointer = 1
  End If

End Sub

Private Sub Texte3_LostFocus()
  texte3.Text = UCase(texte3.Text)

End Sub

Private Sub Texte4_LostFocus()
  texte4.Text = UCase(texte4.Text)

End Sub

Private Sub Texte5_LostFocus()
  texte5.Text = UCase(texte5.Text)

End Sub

Private Sub Texte6_LostFocus()
  texte6.Text = UCase(texte6.Text)

End Sub

Private Sub Texte7_LostFocus()
  texte7.Text = UCase(texte7.Text)

End Sub

Private Sub Texte8_LostFocus()
  texte8.Text = UCase(texte8.Text)

End Sub


