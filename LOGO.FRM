VERSION 4.00
Begin VB.Form logo 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3060
   ClientLeft      =   1965
   ClientTop       =   2295
   ClientWidth     =   6165
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   FillColor       =   &H00808080&
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   0
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF8080&
   Height          =   3465
   Left            =   1905
   LinkTopic       =   "Feuille1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "LOGO.frx":0000
   ScaleHeight     =   3060
   ScaleWidth      =   6165
   Top             =   1950
   Width           =   6285
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   4200
      TabIndex        =   2
      Top             =   1995
      Width           =   1275
   End
   Begin VB.Timer Minuterie1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   5565
      Top             =   2520
   End
   Begin VB.Label Etiquette2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   $"LOGO.frx":4EE6
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1995
      Left            =   3570
      TabIndex        =   1
      Top             =   210
      Width           =   2445
   End
   Begin VB.Label Etiquette1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Droits de copies réservés à MORIANA André"
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
      Left            =   525
      TabIndex        =   0
      Top             =   2625
      Width           =   4950
   End
End
Attribute VB_Name = "logo"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Fclient.Show
  Unload logo

End Sub

Private Sub Form_Load()
  Load Serveur
End Sub

Private Sub Minuterie1_Timer()
  Minuterie1.Enabled = False
  Fclient.Show
  Unload logo
End Sub

