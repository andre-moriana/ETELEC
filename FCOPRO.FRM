VERSION 4.00
Begin VB.Form fcopro 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fichier Copropriétées"
   ClientHeight    =   5895
   ClientLeft      =   2640
   ClientTop       =   735
   ClientWidth     =   4275
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
   Height          =   6300
   Left            =   2580
   LinkTopic       =   "Feuille1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   4275
   Top             =   390
   Width           =   4395
   Begin VB.TextBox Text1 
      DataField       =   "MEMO"
      DataSource      =   "Donnée1"
      Height          =   1170
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   4620
      Width           =   4005
   End
   Begin VB.CommandButton Commande1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "RETOUR"
      Height          =   330
      Left            =   2895
      TabIndex        =   16
      Top             =   420
      Width           =   1185
   End
   Begin VB.TextBox Texte2 
      Appearance      =   0  'Flat
      DataField       =   "NCOPRO"
      DataSource      =   "Donnée1"
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      MaxLength       =   40
      TabIndex        =   1
      Top             =   945
      Width           =   3945
   End
   Begin VB.TextBox Texte8 
      Appearance      =   0  'Flat
      DataField       =   "NTELS"
      DataSource      =   "Donnée1"
      Height          =   360
      Left            =   1575
      MaxLength       =   16
      TabIndex        =   7
      Top             =   3885
      Width           =   2385
   End
   Begin VB.TextBox Texte7 
      Appearance      =   0  'Flat
      DataField       =   "CONTS"
      DataSource      =   "Donnée1"
      Height          =   360
      Left            =   1575
      MaxLength       =   40
      TabIndex        =   6
      Top             =   3465
      Width           =   2385
   End
   Begin VB.TextBox Texte6 
      Appearance      =   0  'Flat
      DataField       =   "CPOSTS"
      DataSource      =   "Donnée1"
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
      MaxLength       =   5
      TabIndex        =   5
      Top             =   2940
      Width           =   825
   End
   Begin VB.TextBox Texte5 
      Appearance      =   0  'Flat
      DataField       =   "VILLES"
      DataSource      =   "Donnée1"
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
      Left            =   705
      MaxLength       =   40
      TabIndex        =   4
      Top             =   2520
      Width           =   3360
   End
   Begin VB.TextBox Texte4 
      Appearance      =   0  'Flat
      DataField       =   "ADRS2"
      DataSource      =   "Donnée1"
      Height          =   360
      Left            =   120
      MaxLength       =   40
      TabIndex        =   3
      Top             =   2100
      Width           =   3960
   End
   Begin VB.TextBox Texte3 
      Appearance      =   0  'Flat
      DataField       =   "ADRS1"
      DataSource      =   "Donnée1"
      Height          =   360
      Left            =   135
      MaxLength       =   40
      TabIndex        =   2
      Top             =   1680
      Width           =   3945
   End
   Begin VB.Data Donnée1 
      Appearance      =   0  'Flat
      Caption         =   "Fichier copro"
      Connect         =   "dbase iv"
      DatabaseName    =   "c:\gestion"
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
      Height          =   315
      Left            =   315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "COPRO"
      Top             =   0
      Visible         =   0   'False
      Width           =   4050
   End
   Begin VB.TextBox Texte1 
      Appearance      =   0  'Flat
      DataField       =   "CCLIENT"
      DataSource      =   "Donnée1"
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   885
      MaxLength       =   11
      TabIndex        =   0
      Top             =   180
      Width           =   1710
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarque :"
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
      Height          =   300
      Left            =   105
      TabIndex        =   17
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Label Etiquette7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero tel."
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
      Height          =   300
      Left            =   315
      TabIndex        =   14
      Top             =   3885
      Width           =   1335
   End
   Begin VB.Label Etiquette6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
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
      Height          =   255
      Left            =   630
      TabIndex        =   13
      Top             =   3465
      Width           =   1065
   End
   Begin VB.Label Etiquette5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Code postal"
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
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   2940
      Width           =   1485
   End
   Begin VB.Label Etiquette4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ville"
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
      Height          =   300
      Left            =   105
      TabIndex        =   11
      Top             =   2520
      Width           =   705
   End
   Begin VB.Label Etiquette3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Adresse"
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
      Left            =   105
      TabIndex        =   10
      Top             =   1365
      Width           =   1305
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
      Height          =   375
      Left            =   105
      TabIndex        =   9
      Top             =   675
      Width           =   1935
   End
   Begin VB.Label Etiquette1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   120
      TabIndex        =   8
      Top             =   195
      Width           =   1020
   End
End
Attribute VB_Name = "fcopro"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim nouveau%

Private Sub Commande1_Click()
  Unload fcopro
End Sub

Private Sub Donnée1_Validate(Action As Integer, Save As Integer)
If Save Then
  rep = MsgBox("Voulez vous enregistrer ? ", 4)
  If rep = 6 Then
    Donnée1.Recordset("dateints") = Format$(Now, "dd/mm/yyyy")
  Else
    varncopro$ = ""
    Donnée1.UpdateControls
  End If
End If
End Sub

Private Sub Form_Load()
  Me.MousePointer = 11
  texte1.Text = varcclient$
  Donnée1.RecordSource = "SELECT copro.* FROM copro WHERE cclient = '" & varcclient$ & "'"
  Donnée1.Refresh
  If varncopro$ <> "" Then
  '
  Do
  pos% = InStr(varncopro$, "'")
  If pos% <> 0 Then Mid$(varncopro$, pos%, 1) = "?"
  Loop While pos% <> 0
   rech$ = "ncopro like '" & varncopro$ & "'"
   Donnée1.Recordset.FindFirst rech$
   If Not (Donnée1.Recordset.NoMatch) Then
      Me.MousePointer = 0
      Donnée1.Recordset.Edit
   Else
      Me.MousePointer = 0
      Donnée1.Recordset.AddNew
      texte1.Text = varcclient$
      texte2.Text = varncopro$
   End If
  Else
   Me.MousePointer = 0
   Donnée1.Recordset.AddNew
   texte1.Text = varcclient$
   texte2.Text = varncopro$
  End If
End Sub

Private Sub texte2_LostFocus()
  texte2.Text = UCase(texte2.Text)
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

