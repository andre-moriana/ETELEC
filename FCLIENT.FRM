VERSION 4.00
Begin VB.Form Fclient 
   AutoRedraw      =   -1  'True
   Caption         =   "SUIVI DES INTERVENTIONS CLIENTS"
   ClientHeight    =   6510
   ClientLeft      =   15
   ClientTop       =   645
   ClientWidth     =   9570
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
   Height          =   7200
   Left            =   -45
   LinkMode        =   1  'Source
   LinkTopic       =   "Fclient"
   ScaleHeight     =   6510
   ScaleWidth      =   9570
   Top             =   15
   Width           =   9690
   WindowState     =   2  'Maximized
   Begin VB.Data Donnée1 
      Connect         =   "dBASE IV;"
      DatabaseName    =   "C:\GESTION"
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CLIENT"
      Top             =   2835
      Width           =   5370
   End
   Begin VB.CommandButton Command 
      Caption         =   "Enregistrer"
      Height          =   330
      Index           =   5
      Left            =   2310
      TabIndex        =   6
      Top             =   6195
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00000000&
      Caption         =   "Ajout"
      Height          =   330
      Index           =   3
      Left            =   6615
      TabIndex        =   4
      Top             =   6195
      Width           =   1485
   End
   Begin VB.CommandButton Command 
      Caption         =   "Modification"
      Height          =   330
      Index           =   4
      Left            =   8085
      TabIndex        =   5
      Top             =   6195
      Width           =   1485
   End
   Begin VB.CommandButton Command 
      Caption         =   "Ajout "
      Height          =   330
      Index           =   1
      Left            =   2310
      TabIndex        =   2
      Top             =   6195
      Width           =   1485
   End
   Begin VB.CommandButton Command 
      Caption         =   "Modification"
      Height          =   330
      Index           =   2
      Left            =   3780
      TabIndex        =   3
      Top             =   6195
      Width           =   1485
   End
   Begin VB.CommandButton Command 
      Caption         =   "Recherche "
      Height          =   330
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   6195
      Width           =   1485
   End
   Begin VB.Frame cadre1 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   2850
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   5370
      Begin VB.TextBox Texte1 
         DataField       =   "VILLEC"
         DataSource      =   "Donnée1"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   10
         Top             =   1575
         Width           =   4005
      End
      Begin VB.TextBox Texte1 
         Alignment       =   2  'Center
         DataField       =   "CODEPOST"
         DataSource      =   "Donnée1"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1260
         MaxLength       =   5
         TabIndex        =   11
         Top             =   1890
         Width           =   855
      End
      Begin VB.TextBox Texte1 
         DataField       =   "CONTACT"
         DataSource      =   "Donnée1"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   12
         Top             =   2205
         Width           =   4005
      End
      Begin VB.TextBox Texte1 
         DataField       =   "NTELC"
         DataSource      =   "Donnée1"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   2625
         MaxLength       =   16
         TabIndex        =   13
         Top             =   2520
         Width           =   1800
      End
      Begin VB.TextBox Texte1 
         DataField       =   "NCLIENT"
         DataSource      =   "Donnée1"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   0
         Left            =   1260
         LinkItem        =   "L1C1"
         LinkTopic       =   "Excel|[Class1]Feuil1"
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "FCLIENT.frx":0000
         Top             =   315
         Width           =   4005
      End
      Begin VB.TextBox Texte1 
         DataField       =   "ADRC1"
         DataSource      =   "Donnée1"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   8
         Top             =   945
         Width           =   4005
      End
      Begin VB.TextBox Texte1 
         DataField       =   "ADRC2"
         DataSource      =   "Donnée1"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   9
         Top             =   1260
         Width           =   4005
      End
      Begin VB.Label Etiquette1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nom :"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   37
         Top             =   315
         Width           =   975
      End
      Begin VB.Label Etiquette1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Adresse :"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   36
         Top             =   945
         Width           =   1080
      End
      Begin VB.Label Etiquette1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Ville :"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   420
         TabIndex        =   35
         Top             =   1575
         Width           =   735
      End
      Begin VB.Label Etiquette1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "C P :"
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
         Height          =   270
         Index           =   4
         Left            =   735
         TabIndex        =   34
         Top             =   1890
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Etiquette2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact :"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   33
         Top             =   2205
         Width           =   855
      End
      Begin VB.Label Etiquette2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de téléphone :"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   32
         Top             =   2520
         Width           =   2310
      End
   End
   Begin VB.Frame Cadre3 
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   400
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   5460
      TabIndex        =   18
      Top             =   0
      Width           =   4110
      Begin VB.TextBox Texte1 
         DataField       =   "REM"
         DataSource      =   "Donnée1"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   7
         Left            =   105
         MaxLength       =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   1995
         Width           =   3810
      End
      Begin VB.TextBox Texte1 
         Alignment       =   2  'Center
         DataField       =   "MO1"
         DataSource      =   "Donnée1"
         BeginProperty Font 
            name            =   "CCLarge"
            charset         =   0
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   2625
         MaxLength       =   3
         TabIndex        =   21
         Top             =   1050
         Width           =   540
      End
      Begin VB.TextBox Texte1 
         Alignment       =   2  'Center
         DataField       =   "MO2"
         DataSource      =   "Donnée1"
         BeginProperty Font 
            name            =   "CCLarge"
            charset         =   0
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   2625
         TabIndex        =   23
         Top             =   1365
         Width           =   540
      End
      Begin VB.TextBox Texte1 
         Alignment       =   2  'Center
         DataField       =   "DEPL"
         DataSource      =   "Donnée1"
         BeginProperty Font 
            name            =   "CCLarge"
            charset         =   0
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   2625
         MaxLength       =   3
         TabIndex        =   25
         Top             =   1680
         Width           =   540
      End
      Begin VB.TextBox Texte1 
         Alignment       =   2  'Center
         DataField       =   "NJ"
         DataSource      =   "Donnée1"
         BeginProperty Font 
            name            =   "CCLarge"
            charset         =   0
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   17
         Top             =   840
         Width           =   540
      End
      Begin VB.TextBox Texte1 
         Alignment       =   2  'Center
         DataField       =   "MD"
         DataSource      =   "Donnée1"
         BeginProperty Font 
            name            =   "CCLarge"
            charset         =   0
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   9
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   15
         Top             =   525
         Width           =   540
      End
      Begin VB.TextBox Texte1 
         Alignment       =   2  'Center
         DataField       =   "NOCPT"
         DataSource      =   "Donnée1"
         BeginProperty Font 
            name            =   "CCLarge"
            charset         =   0
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   14
         Top             =   210
         Width           =   1905
      End
      Begin VB.CheckBox Coche1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contrat d'entretien"
         DataField       =   "FACC"
         DataSource      =   "Donnée1"
         BeginProperty Font 
            name            =   "CCLarge"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   2310
         TabIndex        =   19
         Top             =   525
         Width           =   1710
      End
      Begin VB.Label Etiquette2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Memo :"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   30
         Top             =   1785
         Width           =   855
      End
      Begin VB.Label Etiquette8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Taux deplacement :"
         BeginProperty Font 
            name            =   "CCLarge"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   315
         TabIndex        =   16
         Top             =   1575
         Width           =   2115
      End
      Begin VB.Label Etiquette7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Taux main d'oeuvre autres heures :"
         BeginProperty Font 
            name            =   "CCLarge"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   0
         TabIndex        =   28
         Top             =   1365
         Width           =   2430
      End
      Begin VB.Label Etiquette6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Taux main d'oeuvre 1ere heure :"
         BeginProperty Font 
            name            =   "CCLarge"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   210
         TabIndex        =   27
         Top             =   1155
         Width           =   2220
      End
      Begin VB.Label Etiquette5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de jour :"
         BeginProperty Font 
            name            =   "CCLarge"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   105
         TabIndex        =   24
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label Etiquette4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mode de paiement :"
         BeginProperty Font 
            name            =   "CCLarge"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   105
         TabIndex        =   22
         Top             =   525
         Width           =   1485
      End
      Begin VB.Label Etiquette3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de compte :"
         BeginProperty Font 
            name            =   "CCLarge"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   105
         TabIndex        =   20
         Top             =   210
         Width           =   1485
      End
   End
   Begin VB.Data Donnée3 
      Appearance      =   0  'Flat
      Caption         =   "Donnée3"
      Connect         =   "dbase iv"
      DatabaseName    =   "C:\gestion"
      Exclusive       =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CLIENT"
      Top             =   5460
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Frame Cadre4 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   0
      TabIndex        =   29
      Top             =   3150
      Width           =   9570
      Begin VB.TextBox Texte11 
         Height          =   960
         Left            =   5670
         TabIndex        =   40
         Top             =   1680
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Data donnée2 
         Caption         =   "donnée2"
         Connect         =   "dBASE IV;"
         DatabaseName    =   "C:\GESTION"
         Exclusive       =   0   'False
         Height          =   345
         Left            =   840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Copro"
         Top             =   1995
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.Data Baseint 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Connect         =   "dBASE IV;"
         DatabaseName    =   "c:\gestion"
         Exclusive       =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   840
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "INTERV"
         Top             =   1680
         Visible         =   0   'False
         Width           =   3690
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FCLIENT.frx":0003
         Height          =   2745
         Left            =   105
         OleObjectBlob   =   "FCLIENT.frx":0013
         TabIndex        =   0
         Top             =   210
         Width           =   9360
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Interventions"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5355
      TabIndex        =   39
      Top             =   6195
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clients"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   105
      TabIndex        =   38
      Top             =   6195
      Width           =   750
   End
   Begin MSComDlg.CommonDialog Dialogue1 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Menu M_client 
      Caption         =   "&Client"
      Begin VB.Menu M_Ajout 
         Caption         =   "&Ajouter"
         Shortcut        =   ^A
      End
      Begin VB.Menu M_Modifcli 
         Caption         =   "&Modification"
         Shortcut        =   ^M
      End
      Begin VB.Menu M_Sup 
         Caption         =   "&Supprimer"
         Shortcut        =   ^S
      End
      Begin VB.Menu M_enreg 
         Caption         =   "&Enregistrer"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu m1 
         Caption         =   "-"
      End
      Begin VB.Menu m_config 
         Caption         =   "Configuration imprimante"
      End
      Begin VB.Menu M_Impr 
         Caption         =   "&Impression"
      End
      Begin VB.Menu m2 
         Caption         =   "-"
      End
      Begin VB.Menu M_R_cclient 
         Caption         =   "a&ller au..."
         Begin VB.Menu M_suivant 
            Caption         =   "&Suivant"
         End
         Begin VB.Menu M_Precedent 
            Caption         =   "&Précédent"
         End
         Begin VB.Menu M_Premier 
            Caption         =   "P&remier"
         End
         Begin VB.Menu M_Dernier 
            Caption         =   "&Dernier"
         End
      End
      Begin VB.Menu M_R_Nclient 
         Caption         =   "&Rechercher"
         Shortcut        =   ^R
      End
      Begin VB.Menu M_R_Poursuivre 
         Caption         =   "&Poursuivre"
         Enabled         =   0   'False
      End
      Begin VB.Menu m4 
         Caption         =   "-"
      End
      Begin VB.Menu M_QUIT 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu m_int 
      Caption         =   "&Intervention"
      Begin VB.Menu M_A_Interv 
         Caption         =   "&Ajouter"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu Modif_int 
         Caption         =   "M&odifier"
         Enabled         =   0   'False
      End
      Begin VB.Menu maj_int 
         Caption         =   "&Mise a jour"
         Enabled         =   0   'False
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu m6 
         Caption         =   "-"
      End
      Begin VB.Menu M_Imp_Interv 
         Caption         =   "&Imprimer"
         Enabled         =   0   'False
      End
      Begin VB.Menu M_saut 
         Caption         =   "&Saut de page"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mfenetre 
      Caption         =   "&Fenêtre"
      Begin VB.Menu m_liste_int 
         Caption         =   "&Liste des interventions"
      End
      Begin VB.Menu Liste_copro 
         Caption         =   "Li&ste des copro."
      End
   End
   Begin VB.Menu M_aide 
      Caption         =   "?"
      Begin VB.Menu M_LOGO 
         Caption         =   "A propos de"
      End
   End
End
Attribute VB_Name = "Fclient"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim RCclient$, RNclient$, codecli$, recherche$
Dim FCModif%, datachanged%, rech%, ajout%

Private Function clival%(VCclient$)
    clival% = False
    donnée3.RecordSource = "SELECT client.* FROM client ORDER BY Cclient "
    donnée3.Refresh
    Fclient.MousePointer = 11
    recherche$ = "CCLIENT like '" & VCclient$ & "'"
    donnée3.Recordset.FindFirst recherche$
    clival% = donnée3.Recordset.NoMatch
    Fclient.MousePointer = 0
    donnée3.Recordset.Close
End Function









Private Sub Command_Click(Index As Integer)
Select Case Index
 Case 0
 M_R_NCLIENT_Click
 Case 1
 M_ajout_Click
 Case 2
 M_Modifcli_Click
 Case 3
 M_A_Interv_Click
 Case 4
 Modif_int_Click
 Case 5
 M_Enreg_Click
End Select
End Sub


Private Sub Donnée1_Reposition()
If Not (datachanged%) And Not (Donnée1.Recordset.EOF) Then
  Fclient.MousePointer = 11
  If Donnée1.Recordset.Fields("cclient") <> "" Then
    Donnée1.Caption = Donnée1.Recordset.Fields("cclient")
  Else
    Donnée1.Caption = Space(11)
  End If
  varcclient$ = Donnée1.Recordset.Fields("cclient")
  baseint.RecordSource = "SELECT interv.* FROM interv WHERE cclient = '" & varcclient$ & "' ORDER BY NUMERO DESC"
  baseint.Refresh
  If baseint.Recordset.RecordCount > 0 Then
    Maj_int.Enabled = True
    M_imp_interv.Enabled = True
    modif_int.Enabled = True
  Else
    Maj_int.Enabled = False
    M_imp_interv.Enabled = False
    modif_int.Enabled = False
  End If
  baseint.Caption = baseint.Recordset.RecordCount & " fiches"
 
  texte1(0).SelStart = 0
  texte1(0).SelLength = Len(texte1(0).Text)
  Fclient.MousePointer = 0
End If
End Sub

Private Sub Donnée1_Validate(Action As Integer, Save As Integer)
  If Save Or Action = 6 Then
    rep = MsgBox("La fiche a été modifiée,voulez vous l'enregistrer ? ", 4)
    Fclient.MousePointer = 11
    If rep = 6 Then
      If Not (ajout%) Then
        Donnée1.UpdateRecord
      Else
        codecli$ = Left(texte1(0).Text, 2) & texte1(4).Text & Right(Trim(texte1(6).Text), 4)
        
        If codecli$ <> "" And clival%(codecli$) Then
         Donnée1.Recordset("cclient") = codecli$
         Donnée1.UpdateRecord
        Else
         MsgBox "CODE CLIENT DEJA EXISTANT"
         Donnée1.Recordset.CancelUpdate
         Donnée1.UpdateControls
        End If
      End If
    Else
      Donnée1.Recordset.CancelUpdate
      Donnée1.UpdateControls
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Const KEY_ESCAPE = &H1B
    If KeyCode = KEY_ESCAPE Then
        Donnée1.UpdateControls
    End If

End Sub

Private Sub Form_Load()
  ajout% = False
  datachanged% = False
  FCModif% = False
  Donnée1.RecordSource = "SELECT CLIENT.* FROM CLIENT ORDER BY Nclient "
  Donnée1.Refresh
  If Donnée1.Recordset.Fields("cclient") <> "" Then varcclient$ = Donnée1.Recordset.Fields("cclient")
  baseint.RecordSource = "SELECT interv.* FROM interv WHERE cclient = '" & varcclient$ & "' ORDER BY NUMERO DESC"
  baseint.Refresh
  If baseint.Recordset.RecordCount > 0 Then
    Maj_int.Enabled = True
    M_imp_interv.Enabled = True
    modif_int.Enabled = True
  Else
    Maj_int.Enabled = False
    M_imp_interv.Enabled = False
    modif_int.Enabled = False
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Serveur
End Sub

Private Sub Liste_copro_Click()
 LCopro.Show 1
End Sub

Private Sub M_A_Interv_Click()
    varcclient$ = Donnée1.Recordset.Fields("cclient")
    F_A_INT.Show 1
    Fclient.MousePointer = 11
    baseint.RecordSource = "SELECT interv.* FROM interv WHERE cclient = '" & varcclient$ & "' ORDER BY NUMERO DESC"
    baseint.Refresh
    If baseint.Recordset.RecordCount > 0 Then
      Maj_int.Enabled = True
      modif_int.Enabled = True
      M_imp_interv.Enabled = True
    End If
    baseint.Caption = baseint.Recordset.RecordCount & " fiches"
    Fclient.MousePointer = 1
End Sub

Private Sub M_ajout_Click()
  
  ajout% = True
  datachanged% = True
  M_R_cclient.Enabled = False
  M_R_nclient.Enabled = False
  M_Ajout.Enabled = False
  M_impr.Enabled = False
  M_sup.Enabled = False
  M_Enreg.Enabled = True
  m_int.Enabled = False
  mfenetre.Enabled = False
  I = 0
  For I = 0 To 4
    Command(I).Visible = False
  Next
  Command(5).Visible = True
  baseint.Visible = False
  Donnée1.Visible = False
  M_Modifcli.Enabled = False

  cadre1.Enabled = True
  cadre3.Enabled = True
  
  Donnée1.Recordset.AddNew
End Sub


Private Sub m_config_Click()
  dialogue1.Flags = &H40
  dialogue1.Action = 5
End Sub

Private Sub M_Dernier_Click()
Donnée1.Recordset.MoveLast
End Sub

Private Sub M_Enreg_Click()

  Donnée1.Recordset.Update
  Fclient.MousePointer = 11
  datachanged% = False
  If ajout% Then
    Donnée1.Refresh
    Donnée1.Recordset.FindFirst "cclient = '" & codecli$ & "'"
  End If
  ajout% = False
  M_Enreg.Enabled = False
  cadre1.Enabled = False
  cadre3.Enabled = False
  M_Ajout.Enabled = True
  M_Modifcli.Enabled = True
  M_sup.Enabled = True
  M_R_cclient.Enabled = True
  M_R_nclient.Enabled = True
  M_R_Poursuivre.Enabled = False
  m_int.Enabled = True
  mfenetre.Enabled = True
  Donnée1.Visible = True
  Fclient.MousePointer = 0
  I = 0
  For I = 0 To 4
    Command(I).Visible = True
  Next
  Command(5).Visible = False

End Sub

Private Sub M_Imp_Interv_Click()
Dim TDescro As Object
Dim msg
    On Error GoTo ErrorHandler2
    Printer.ScaleMode = 2
    taillep = Printer.FontSize
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
      donnée2RecordSource = "select copro.* from copro where cclient = '" & baseint.Recordset("cclient") & "'"
      donnée2.Refresh
      donnée2.Recordset.FindFirst "ncopro like '" & baseint.Recordset("ncopro") & "*'"
      If donnée2.Recordset("adrs1") <> "" Then Printer.Print donnée2.Recordset("adrs1")
      If donnée2.Recordset("adrs2") <> "" Then Printer.Print donnée2.Recordset("adrs2")
      If donnée2.Recordset("cposts") <> "" Then Printer.Print donnée2.Recordset("cposts"), ;
      If donnée2.Recordset("villes") <> "" Then Printer.Print donnée2.Recordset("villes")
      Printer.Print "Contact " & donnée2.Recordset("conts"), "Téléphone  :" & donnée2.Recordset("ntels")
  '    Printer.Print " "
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
 Exit Sub
ErrorHandler2:
msg = Error(Err)
MsgBox msg
Exit Sub
End Sub

Private Sub M_Impr_Click()
    Dim msg
    On Error GoTo ErrorHandler
    PrintForm
    Exit Sub
ErrorHandler:
    msg = "Impossible d'imprimer la feuille."
    MsgBox msg
    Resume Next


End Sub


Private Sub m_liste_int_Click()
  Finterv.Show 1
End Sub

Private Sub M_LOGO_Click()
logo.Show
End Sub

Private Sub M_Modifcli_Click()
  datachanged% = True
  codecli$ = Left(texte1(0).Text, 2) & texte1(4).Text & Right(Trim(texte1(6).Text), 4)
  M_R_cclient.Enabled = False
  M_R_nclient.Enabled = False
  M_Ajout.Enabled = False
  M_impr.Enabled = False
  M_sup.Enabled = False
  M_Enreg.Enabled = True
'  M_liste_int.Enabled = False
  M_Modifcli.Enabled = False
  m_int.Enabled = False
  mfenetre.Enabled = False
  I = 0
  For I = 0 To 4
    Command(I).Visible = False
  Next
  Command(5).Visible = True

  cadre1.Enabled = True
  cadre3.Enabled = True
'  M_A_Interv.Enabled = False

  Donnée1.Visible = False
  baseint.Visible = False
  
  Donnée1.Recordset.Edit
End Sub

Private Sub M_Precedent_Click()
Donnée1.Recordset.MovePrevious
If Donnée1.Recordset.BOF Then Donnée1.Recordset.MoveNext
End Sub

Private Sub M_Premier_Click()
Donnée1.Recordset.MoveFirst
End Sub

Private Sub M_quit_Click()
  Unload Fclient
End Sub

Private Sub M_R_NCLIENT_Click()
    Donnée1.UpdateControls
    RCclient$ = ""
    repere = Donnée1.Recordset.Bookmark

    RNclient$ = UCase(InputBox("Entrez le NOM du client", "RECHERCHE"))
  If RNclient$ <> "" Then
    Fclient.MousePointer = 11
    recherche$ = "NCLIENT like '" & RNclient$ & "*'"
    Donnée1.Recordset.FindFirst recherche$
    If Donnée1.Recordset.NoMatch Then
      Donnée1.Recordset.Bookmark = repere
      M_R_Poursuivre.Enabled = False
    Else
      M_R_Poursuivre.Enabled = True
    End If
    Fclient.MousePointer = 1
  End If
End Sub

Private Sub M_R_Poursuivre_Click()
  
'  If RCclient$ <> "" Then recherche$ = "CCLIENT LIKE '" & RCclient$ & "*'"
'  If RNclient$ <> "" Then recherche$ = "NCLIENT LIKE '" & RNclient$ & "*'"
  Fclient.MousePointer = 11
    
  Donnée1.Recordset.FindNext recherche$
  
  If Donnée1.Recordset.NoMatch Then M_R_Poursuivre.Enabled = False
  Fclient.MousePointer = 0
  
End Sub


Private Sub M_suivant_Click()
Donnée1.Recordset.MoveNext
If Donnée1.Recordset.EOF Then Donnée1.Recordset.MovePrevious
End Sub

Private Sub M_Sup_Click()
  reponse = MsgBox("Etes vous sur de vouloir éffacer la fiche client ?", 36 + 256, "ATTENTION")
  If reponse = 6 Then
    Donnée1.Recordset.Delete
    Donnée1.Refresh
  End If

End Sub

Private Sub MAJ_INT_Click()
  varcclient$ = Donnée1.Recordset("cclient")
  varnumero% = baseint.Recordset("numero")
  FMAJ_INT.Show 1
  baseint.Refresh
  baseint.Caption = baseint.Recordset.RecordCount & " fiches"
  If baseint.Recordset.RecordCount > 0 Then
    modif_int.Enabled = True
    Maj_int.Enabled = True
  Else
    modif_int.Enabled = False
    Maj_int.Enabled = False
  End If

End Sub

Private Sub Modif_int_Click()
    Me.MousePointer = 11
    varnumero% = baseint.Recordset("numero")
    Modifint.Show 1
    Fclient.MousePointer = 11
    varcclient$ = Donnée1.Recordset.Fields("cclient")
    baseint.RecordSource = "SELECT interv.* FROM interv WHERE cclient = '" & varcclient$ & "' ORDER BY NUMERO DESC"
    baseint.Refresh
    If baseint.Recordset.RecordCount > 0 Then
      Maj_int.Enabled = True
      modif_int.Enabled = True
    End If
    baseint.Caption = baseint.Recordset.RecordCount & " fiches"
    Fclient.MousePointer = 1
End Sub


Private Sub Texte1_LostFocus(Index As Integer)
  texte1(Index).Text = UCase(texte1(Index).Text)
End Sub


