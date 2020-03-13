VERSION 4.00
Begin VB.Form Fclient2 
   Appearance      =   0  'Flat
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FICHE CLIENT"
   ClientHeight    =   4485
   ClientLeft      =   315
   ClientTop       =   1425
   ClientWidth     =   5700
   ClipControls    =   0   'False
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
   Height          =   4890
   Left            =   255
   LinkTopic       =   "Feuille1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5700
   Top             =   1080
   Width           =   5820
   Begin VB.CommandButton Commande1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "RETOUR"
      Height          =   435
      Left            =   3780
      TabIndex        =   0
      Top             =   3990
      Width           =   1710
   End
   Begin VB.Data Donnée1 
      Appearance      =   0  'Flat
      Caption         =   "Donnée1"
      Connect         =   "dbase iv"
      DatabaseName    =   "c:\gestion"
      Exclusive       =   0   'False
      Height          =   315
      Left            =   60
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CLIENT"
      Top             =   4140
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Texte1 
      Appearance      =   0  'Flat
      DataField       =   "NCLIENT"
      DataSource      =   "Donnée1"
      Height          =   480
      Index           =   0
      Left            =   1260
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   315
      Width           =   4230
   End
   Begin VB.TextBox Texte1 
      Appearance      =   0  'Flat
      DataField       =   "ADRC1"
      DataSource      =   "Donnée1"
      Height          =   285
      Index           =   1
      Left            =   1260
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   945
      Width           =   4215
   End
   Begin VB.TextBox Texte1 
      Appearance      =   0  'Flat
      DataField       =   "ADRC2"
      DataSource      =   "Donnée1"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   1260
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1305
      Width           =   4215
   End
   Begin VB.TextBox Texte1 
      Appearance      =   0  'Flat
      DataField       =   "VILLEC"
      DataSource      =   "Donnée1"
      Height          =   300
      Index           =   3
      Left            =   1260
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2670
   End
   Begin VB.TextBox Texte1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "CODEPOST"
      DataSource      =   "Donnée1"
      Height          =   285
      Index           =   4
      Left            =   4620
      MaxLength       =   5
      MultiLine       =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Texte1 
      Appearance      =   0  'Flat
      DataField       =   "REM"
      DataSource      =   "Donnée1"
      Height          =   780
      Index           =   7
      Left            =   1125
      MaxLength       =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2970
      Width           =   4335
   End
   Begin VB.TextBox Texte1 
      Appearance      =   0  'Flat
      DataField       =   "CONTACT"
      DataSource      =   "Donnée1"
      Height          =   285
      Index           =   5
      Left            =   1260
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2100
      Width           =   3585
   End
   Begin VB.TextBox Texte1 
      Appearance      =   0  'Flat
      DataField       =   "NTELC"
      DataSource      =   "Donnée1"
      Height          =   285
      Index           =   6
      Left            =   2730
      MaxLength       =   16
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2115
   End
   Begin VB.Label Etiquette1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom :"
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
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   330
      Width           =   975
   End
   Begin VB.Label Etiquette1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Adresse :"
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
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   945
      Width           =   975
   End
   Begin VB.Label Etiquette1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Ville :"
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
      Index           =   2
      Left            =   525
      TabIndex        =   13
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Etiquette1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "C P :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   4095
      TabIndex        =   12
      Top             =   1785
      Width           =   450
   End
   Begin VB.Label Etiquette2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact :"
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
      Index           =   0
      Left            =   315
      TabIndex        =   11
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Etiquette2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de téléhone :"
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
      Index           =   1
      Left            =   315
      TabIndex        =   10
      Top             =   2595
      Width           =   2310
   End
   Begin VB.Label Etiquette2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Memo :"
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
      Index           =   2
      Left            =   315
      TabIndex        =   9
      Top             =   2940
      Width           =   855
   End
End
Attribute VB_Name = "Fclient2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Commande1_Click()
  Unload Fclient2
End Sub

Private Sub Form_Activate()
  rech$ = "CCLIENT = '" & varcclient$ & "'"
  Donnée1.Refresh
  Donnée1.Recordset.FindFirst rech$
  If Donnée1.Recordset.NoMatch Then
    MsgBox ("FICHE CLIENT INEXISTANTE")
    Unload Fclient2
  End If
End Sub

