VERSION 4.00
Begin VB.Form Modifint 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MODIFICATION FICHE INTERVENTION"
   ClientHeight    =   4170
   ClientLeft      =   2670
   ClientTop       =   1875
   ClientWidth     =   6330
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
   Height          =   4575
   Left            =   2610
   LinkTopic       =   "Feuille1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6330
   Top             =   1530
   Width           =   6450
   Begin VB.CommandButton Command1 
      Caption         =   " Modification copro."
      Height          =   330
      Left            =   2520
      TabIndex        =   13
      Top             =   945
      Width           =   3270
   End
   Begin VB.TextBox Texte3 
      Height          =   330
      Left            =   2730
      MaxLength       =   5
      TabIndex        =   4
      Top             =   3045
      Width           =   750
   End
   Begin VB.TextBox Texte2 
      Height          =   330
      Left            =   2730
      TabIndex        =   3
      Top             =   2625
      Width           =   1275
   End
   Begin VB.Data Donnée2 
      Appearance      =   0  'Flat
      Caption         =   "Donnée2"
      Connect         =   "dbase iv"
      DatabaseName    =   "c:\gestion"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   2655
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "interv"
      Top             =   3795
      Visible         =   0   'False
      Width           =   2055
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
      Height          =   900
      Left            =   210
      MaxLength       =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1530
      Width           =   5970
   End
   Begin VB.ComboBox Liste1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "MODIFINT.frx":0000
      Left            =   2625
      List            =   "MODIFINT.frx":0022
      TabIndex        =   0
      Top             =   105
      Width           =   3165
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
      Left            =   1575
      MaxLength       =   3
      TabIndex        =   5
      Top             =   3360
      Width           =   570
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
      Left            =   4830
      TabIndex        =   6
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Data Donnée1 
      Appearance      =   0  'Flat
      Caption         =   "Donnée1"
      Connect         =   "dbase iv"
      DatabaseName    =   "c:\gestion"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   270
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "COPRO"
      Top             =   3810
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Label Etiquette7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caractère  0/240"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4515
      TabIndex        =   14
      Top             =   1260
      Width           =   1815
   End
   Begin MSDBCtls.DBCombo Liste2 
      Bindings        =   "MODIFINT.frx":00FA
      DataField       =   "NCOPRO"
      DataSource      =   "Donnée2"
      Height          =   315
      Left            =   2625
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
      Height          =   435
      Left            =   0
      TabIndex        =   12
      Top             =   105
      Width           =   2430
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
      Height          =   435
      Left            =   945
      TabIndex        =   11
      Top             =   525
      Width           =   1485
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
      TabIndex        =   10
      Top             =   1260
      Width           =   2160
   End
   Begin VB.Label Etiquette4 
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
      Left            =   210
      TabIndex        =   9
      Top             =   2655
      Width           =   2355
   End
   Begin VB.Label Etiquette5 
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
      Left            =   180
      TabIndex        =   8
      Top             =   3045
      Width           =   2385
   End
   Begin VB.Label Etiquette6 
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
      Left            =   150
      TabIndex        =   7
      Top             =   3405
      Width           =   1650
   End
End
Attribute VB_Name = "Modifint"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()

If liste2.MatchedWithList Then  ' Modification si ncopro est dans la liste
  varncopro$ = liste2.Text
  fcopro.Show 1
Else
  'Creation de la copro si ncopro est pas dans la liste, ni vide
  If liste2.Text <> "" Then
    varncopro$ = LTrim(UCase(liste2.Text))
    fcopro.Show 1
    Donnée1.Refresh
'    liste2.Refresh
  End If
End If

End Sub


Private Sub Commande1_Click()
    rep = MsgBox("Voulez vous enregistrer ? ", 4)
    If rep = 6 Then
      donnée2.Recordset.Edit
'      donnée2.Recordset("DATE_AP") = Format$(Now, "dd/mm/yyyy")
      donnée2.Recordset("typeint") = Left$(liste1.Text, 5)
      donnée2.Recordset("ncopro") = liste2.Text
      donnée2.Recordset("descro") = texte1.Text
      If texte2.Text <> "" Then donnée2.Recordset("dateint") = texte2.Text
      donnée2.Recordset("heureint") = texte3.Text
      donnée2.Recordset("technicien") = texte4.Text
      donnée2.Recordset.Update
'    Else
'      Action = 0
'      donnée2.UpdateControls
    End If
 Unload Modifint
End Sub

Private Sub Form_Load()
  
   Donnée1.RecordSource = "SELECT copro.* FROM copro WHERE cclient = '" & varcclient$ & "' ORDER BY ncopro"
   Donnée1.Refresh
   
   donnée2.RecordSource = "select interv.* from interv"
   donnée2.Refresh
   donnée2.Recordset.FindFirst "numero = " & varnumero%
   If donnée2.Recordset("typeint") <> "" Then liste1.Text = donnée2.Recordset("typeint")
   If donnée2.Recordset("descro") <> "" Then texte1.Text = donnée2.Recordset("descro")
   etiquette7.Caption = "Caractère  " & Len(texte1.Text) & "/240"
   If donnée2.Recordset("dateint") <> "" Then texte2.Text = donnée2.Recordset("dateint")
   If donnée2.Recordset("heureint") <> "  H" And donnée2.Recordset("heureint") <> "" Then texte3.Text = donnée2.Recordset("heureint")
   If donnée2.Recordset("technicien") <> "" Then texte4.Text = donnée2.Recordset("technicien")
   Me.MousePointer = 1
End Sub

Private Sub liste2_LostFocus()
If Not liste2.MatchedWithList And liste2.Text <> "" Then
  varncopro$ = LTrim(UCase(liste2.Text))
  fcopro.Show 1
  Donnée1.Refresh
  If varncopro$ <> "" Then
    liste2.Text = varncopro$
  Else
    liste2.SelStart = 0
    liste2.SelLength = Len(liste2.Text)
    liste2.SetFocus
  End If
End If
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

Private Sub Texte4_LostFocus()
  texte4.Text = UCase(texte4.Text)
End Sub

