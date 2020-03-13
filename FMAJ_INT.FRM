VERSION 4.00
Begin VB.Form FMAJ_INT 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mise a jour INTERVENTION"
   ClientHeight    =   4020
   ClientLeft      =   885
   ClientTop       =   1875
   ClientWidth     =   5955
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
   Height          =   4425
   Left            =   825
   LinkMode        =   1  'Source
   LinkTopic       =   "FMajInt"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   0
   Top             =   1530
   Width           =   6075
   Begin VB.CommandButton Commande2 
      Caption         =   "FACTURE"
      Height          =   360
      Left            =   3675
      MouseIcon       =   "FMAJ_INT.frx":0000
      TabIndex        =   15
      Top             =   3465
      Width           =   1515
   End
   Begin VB.TextBox Texte5 
      Height          =   330
      Left            =   2835
      TabIndex        =   10
      Top             =   2940
      Width           =   1695
   End
   Begin VB.Data Donnée1 
      Appearance      =   0  'Flat
      Caption         =   "Donnée1"
      Connect         =   "dbase iv"
      DatabaseName    =   "c:\gestion"
      Exclusive       =   0   'False
      Height          =   330
      Left            =   3255
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "INTERV"
      Top             =   3465
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.CommandButton Commande1 
      Caption         =   "RETOUR"
      Height          =   360
      Left            =   720
      MouseIcon       =   "FMAJ_INT.frx":0442
      TabIndex        =   12
      Top             =   3465
      Width           =   1515
   End
   Begin VB.TextBox Texte4 
      Appearance      =   0  'Flat
      DataField       =   "Tempsdep"
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
      Left            =   2835
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2295
      Width           =   600
   End
   Begin VB.TextBox Texte3 
      Appearance      =   0  'Flat
      DataField       =   "Tempsmo"
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
      Left            =   2835
      MaxLength       =   2
      TabIndex        =   8
      Top             =   1830
      Width           =   600
   End
   Begin VB.TextBox Texte2 
      Appearance      =   0  'Flat
      DataField       =   "Numba"
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
      Left            =   2835
      MaxLength       =   5
      TabIndex        =   7
      Top             =   1365
      Width           =   1095
   End
   Begin VB.TextBox Texte1 
      Appearance      =   0  'Flat
      DataField       =   "technicien"
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
      Left            =   2835
      MaxLength       =   3
      TabIndex        =   6
      Top             =   930
      Width           =   615
   End
   Begin VB.ComboBox Liste1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FMAJ_INT.frx":0884
      Left            =   2835
      List            =   "FMAJ_INT.frx":089D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   525
      Width           =   3030
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "NUMERO"
      DataSource      =   "Donnée1"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   2835
      TabIndex        =   14
      Top             =   105
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NUMERO DE FICHE"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   210
      TabIndex        =   13
      Top             =   105
      Width           =   2220
   End
   Begin VB.Label Etiquette6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date d'intervention"
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
      Left            =   600
      TabIndex        =   11
      Top             =   2985
      Width           =   2070
   End
   Begin VB.Label Etiquette5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Temps déplacement"
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
      Left            =   330
      TabIndex        =   5
      Top             =   2355
      Width           =   2355
   End
   Begin VB.Label Etiquette4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Temps main d'oeuvre"
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
      Left            =   360
      TabIndex        =   4
      Top             =   1860
      Width           =   2295
   End
   Begin VB.Label Etiquette3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero du B.A."
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
      Left            =   495
      TabIndex        =   3
      Top             =   1380
      Width           =   2145
   End
   Begin VB.Label Etiquette2 
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
      Left            =   945
      TabIndex        =   2
      Top             =   975
      Width           =   1710
   End
   Begin VB.Label Etiquette1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FACTURATION"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1155
      TabIndex        =   1
      Top             =   525
      Width           =   1500
   End
End
Attribute VB_Name = "FMAJ_INT"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()

End Sub


Private Sub Commande1_Click()
  Unload FMAJ_INT
End Sub

Private Sub Commande2_Click()
z = Shell("c:\msoffice\excel\excel.exe c:\gestion\Facture~1", 3)

End Sub

Private Sub Donnée1_Validate(Action As Integer, Save As Integer)
If Save Then
    rep = MsgBox("Voulez vous enregistrer ? ", 4)
    If rep = 6 Then
      
      If liste1.Text <> "" Then Donnée1.Recordset("typfac") = Left(liste1.Text, 1)
      If Trim(texte5.Text) <> "" Then Donnée1.Recordset("dateint") = CDate(texte5.Text)
      Donnée1.UpdateRecord
    Else
      Donnée1.Recordset.CancelUpdate
    End If
End If
End Sub



Private Sub Form_Load()

    Donnée1.RecordSource = "SELECT interv.* FROM interv WHERE cclient = '" & varcclient$ & "'"
    Donnée1.Refresh
    rech$ = "numero = " & varnumero%
    Donnée1.Recordset.FindFirst rech$
    If Donnée1.Recordset("typfac") <> "" Then liste1.ListIndex = Val(Donnée1.Recordset("typfac"))
    If Donnée1.Recordset("dateint") <> "" Then texte5.Text = Donnée1.Recordset("dateint")
    Donnée1.Recordset.Edit
End Sub


Private Sub liste1_LostFocus()
If VarType(Donnée1.Recordset("typfac")) = 1 Then
  vartypfac = 0
Else
  vartypfac = Val(Donnée1.Recordset("typfac"))
End If
If liste1.ListIndex <> vartypfac Then texte1.datachanged = True
End Sub

Private Sub Texte1_LostFocus()
  texte1.Text = UCase(texte1.Text)
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
If texte5.Text <> "" Then
  If Not (IsDate(texte5.Text)) Then
    texte5.SelStart = 0
    texte5.SelLength = Len(texte5.Text)
    texte5.SetFocus
  Else
    texte1.datachanged = True
  End If

End If

End Sub

