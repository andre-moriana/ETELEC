VERSION 4.00
Begin VB.Form Serveur 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "MSComm Terminal "
   ClientHeight    =   3105
   ClientLeft      =   105
   ClientTop       =   1620
   ClientWidth     =   7470
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
   Height          =   3510
   Icon            =   "SERVEUR.frx":0000
   Left            =   45
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   7470
   Top             =   1275
   Visible         =   0   'False
   Width           =   7590
   Begin VB.Timer Minuterie1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   750
      Top             =   375
   End
   Begin VB.Data Donnée3 
      Appearance      =   0  'Flat
      Caption         =   "Donnée3"
      Connect         =   "dbase iv;"
      DatabaseName    =   "c:\gestion"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   5670
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "COPRO"
      Top             =   105
      Width           =   1155
   End
   Begin VB.Data Donnée2 
      Appearance      =   0  'Flat
      Caption         =   "Donnée2"
      Connect         =   "dbase iv;"
      DatabaseName    =   "c:\gestion"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   5670
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "INTERV"
      Top             =   420
      Width           =   1155
   End
   Begin VB.Data Donnée1 
      Appearance      =   0  'Flat
      Caption         =   "Donnée1"
      Connect         =   "dbase iv;"
      DatabaseName    =   "C:\GESTION"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   5670
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Client"
      Top             =   735
      Width           =   1155
   End
   Begin VB.TextBox Texte1 
      Appearance      =   0  'Flat
      Height          =   1700
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1320
      Width           =   7455
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   225
      Top             =   360
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      CDTimeout       =   0
      CommPort        =   2
      CTSTimeout      =   0
      DSRTimeout      =   0
      DTREnable       =   -1  'True
      Handshaking     =   0
      InBufferSize    =   2048
      InputLen        =   1
      Interval        =   1000
      NullDiscard     =   0   'False
      OutBufferSize   =   2048
      ParityReplace   =   "?"
      RThreshold      =   1
      RTSEnable       =   -1  'True
      Settings        =   "1200,e,7,1"
      SThreshold      =   0
   End
End
Attribute VB_Name = "Serveur"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z

Dim numfich%
Dim Ret                 'Scratch integer
Dim Temp$               'Scratch string
Dim hlogfile            'Handle of open log file
Dim ESC$, RC$, Pro1$, Pro2$, Pro3$, connect$, deconnect$, msg$, dta$, codecli$, ident$
Dim Max%, niveau%, page%, annul%, Mise_a_jour%, champ%, choix%, lecture%


Dim Mvar1$, Mvar2$, Mvar3$, Mvar4$, Mvar5$, Mvar6$
Dim Vmes1$, VMes2$, VMes3$, VMes4$, VMes5$, VMes6$
Dim REP_STAT_TERMINAL%
Dim REP_STAT_FONCTIONNEMENT%
Dim REP_STAT_VITESSE%
Dim REP_STAT_PROTOCOLE%
Dim CONNECT_EM_ECRAN%
Dim CONNECT_EM_CLAVIER%
Dim CONNECT_EM_MODEM%
Dim CONNECT_EM_PRISE%
Dim CONNECT_REP_ECRAN%
Dim CONNECT_REP_CLAVIER%
Dim CONNECT_REP_MODEM%
Dim CONNECT_REP_PRISE%

Dim FicheMes As DBMes

Dim NumCourMes&

Dim NumfMes%
Dim NbreMes&
Dim LongEntMes%
Dim LongEnrMes%

Private Sub AffDep()
   On Error GoTo msgereur
   donnée2.Recordset.FindFirst "NUMERO = " & texte1.Text
   If Not (donnée2.Recordset.NoMatch) Then
      Donnée1.Recordset.FindFirst "CCLIENT = '" & donnée2.Recordset("CCLIENT") & "'"
      MsComm1.Output = Chr(20) + Chr(12)
      MsComm1.Output = Chr(20) + Chr(27) + Chr(&H44) + Chr(12)
      MsComm1.Output = Chr(27) + Chr(69) + Chr(27) + Chr(93)
      If donnée2.Recordset("ncopro") <> "" Then
        donnée3.RecordSource = "select copro.* from copro where CCLIENT = '" & Donnée1.Recordset("cclient") & "'"
        donnée3.Refresh
        
        MsComm1.Output = Chr(31) & Chr(64) & Chr(65) & Chr(24) + Donnée1.Recordset("Nclient")
        MsComm1.Output = Chr(31) & Chr(65) & Chr(65)
        donnée3.Recordset.FindFirst "ncopro = '" & donnée2.Recordset("ncopro") & "'"
        MsComm1.Output = Chr(10)
        MsComm1.Output = Chr(27) & Chr(93) & Chr(24) & Chr(13)
        MsComm1.Output = donnée3.Recordset("NCopro") & Chr(10) & Chr(13)
        MsComm1.Output = Chr(27) & Chr(93) & Chr(24) & Chr(13)
        If donnée3.Recordset("ADRS1") <> "" Then MsComm1.Output = donnée3.Recordset("ADRS1") & Chr(10) & Chr(13)
        MsComm1.Output = Chr(27) & Chr(93) & Chr(24) & Chr(13)
        If donnée3.Recordset("ADRS2") <> "" Then MsComm1.Output = donnée3.Recordset("ADRS2") & Chr(10) & Chr(13)
        MsComm1.Output = Chr(27) & Chr(93) & Chr(24) & Chr(13)
        If donnée3.Recordset("CPOSTS") <> "" Then MsComm1.Output = donnée3.Recordset("CPOSTS") & "  "
        If donnée3.Recordset("VILLES") <> "" Then MsComm1.Output = donnée3.Recordset("Villes") & Chr(10) & Chr(13)
        MsComm1.Output = Chr(13) + Chr(10) + Chr(13)
      Else
        MsComm1.Output = Chr(10)
        MsComm1.Output = Chr(27) + Chr(93) + Chr(24) + Chr(13)
        MsComm1.Output = Donnée1.Recordset("Nclient") & Chr(10) & Chr(13)
        MsComm1.Output = Chr(27) + Chr(93) + Chr(24) + Chr(13)
        If Donnée1.Recordset("ADRC1") <> "" Then MsComm1.Output = Donnée1.Recordset("ADRC1") & Chr(10) & Chr(13)
        MsComm1.Output = Chr(27) + Chr(93) + Chr(24) + Chr(13)
        If Donnée1.Recordset("ADRC2") <> "" Then MsComm1.Output = Donnée1.Recordset("ADRC2") & Chr(10) & Chr(13)
        MsComm1.Output = Chr(27) + Chr(93) + Chr(24) + Chr(13)
        If Donnée1.Recordset("codepost") <> "" Then MsComm1.Output = Donnée1.Recordset("CodePost") & "  "
        If Donnée1.Recordset("VILLEC") <> "" Then MsComm1.Output = Donnée1.Recordset("Villec") & Chr(10) & Chr(13)
        MsComm1.Output = Chr(27) + Chr(93) + Chr(24) + Chr(13)
        MsComm1.Output = Chr(13) & "Numero de Telephone : "
        If Donnée1.Recordset("NTELC") <> "" Then MsComm1.Output = Donnée1.Recordset("NTELC") & "  "
      End If
          
      MsComm1.Output = Chr(27) + Chr(92) + Chr(27) + Chr(71) + Chr(13) + Chr(10) + Chr(96) + Chr(18) + Chr(103)
      MsComm1.Output = Chr(27) + Chr(65) + Chr(27) + Chr(93)
      If donnée2.Recordset("DESCRO") <> "" Then
        longdescro% = Len(donnée2.Recordset("descro"))
        descro$ = donnée2.Recordset("descro") & Space(240 - longdescro%)
        MsComm1.Output = descro$ & Chr(10) & Chr(13)
      
      End If
      MsComm1.Output = Chr(27) + Chr(92) + Chr(27) + Chr(71) + Chr(13) + Chr(96) + Chr(18) + Chr(103) + Chr(10) + Chr(10)
      MsComm1.Output = "Mise a jour pressez         " + Chr(27) + Chr(93) + " ANNULATION " + Chr(27) + Chr(92) + Chr(10) + Chr(13)
      MsComm1.Output = "Retour a la liste pressez   " + Chr(27) + Chr(93) + "  SOMMAIRE  " + Chr(27) + Chr(92) + Chr(10) + Chr(13)
      niveau% = niveau% + 1
      annul% = False
   Else
      MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + Chr(24) + "FICHE INEXISTANTE" + Chr(7)
      MsComm1.Output = Chr(31) + Chr(88) + Chr(65) + Chr(24)
      MsComm1.Output = "Detail fiche tapez son numero " + Chr(17)
   End If
   Exit Sub
msgereur:
    nd% = Len(texte1.Text)
    texte1.SelStart = 0
    texte1.SelLength = nd%
    texte1.SelText = ""
    Exit Sub

End Sub

Private Sub AffMAJ()
   MsComm1.Output = Chr(11) + Chr(11) + Chr(11) + Chr(11) + Chr(24) + Chr(10) + Chr(24) + Chr(10) + Chr(24)
   MsComm1.Output = Chr(31) + Chr(88) + Chr(65) + Chr(27) + Chr(58) + Chr(105) + Chr(67) + Chr(10) + Chr(10) + Chr(10) + Chr(10) + Chr(10) + Chr(10)
   
   nd% = Len(texte1.Text)
   texte1.SelStart = 0
   texte1.SelLength = nd%
   texte1.SelText = ""
' Numero de ba
   MsComm1.Output = Chr(31) + Chr(74) + Chr(65) + Chr(27) + Chr(67) + Chr(27) + Chr(93) + Chr(24)
   MsComm1.Output = "     Numero du B.A.    "
   If donnée2.Recordset("numba") <> "" Then
     Mvar1$ = donnée2.Recordset("numba")
     MsComm1.Output = Mvar1$
   End If
' Temps Main d Oeuvre
   MsComm1.Output = Chr(31) + Chr(75) + Chr(65) + Chr(27) + Chr(66) + Chr(27) + Chr(93) + Chr(24)
   MsComm1.Output = " Temps main d'" + Chr(25) + Chr(&H7A) + "uvre    "
   If donnée2.Recordset("Tempsmo") <> "" Then
     Mvar2$ = donnée2.Recordset("Tempsmo")
     MsComm1.Output = Mvar2$
   End If
' Temps déplacement
   MsComm1.Output = Chr(31) + Chr(76) + Chr(65) + Chr(27) + Chr(67) + Chr(27) + Chr(93) + Chr(24)
   MsComm1.Output = "  Temps deplacement    "
   If donnée2.Recordset("Tempsdep") <> "" Then
     Mvar3$ = donnée2.Recordset("Tempsdep")
     MsComm1.Output = Mvar3$
   End If
' Date d'intervention
   MsComm1.Output = Chr(31) + Chr(77) + Chr(65) + Chr(27) + Chr(66) + Chr(27) + Chr(93) + Chr(24)
   MsComm1.Output = "Date d'intervention    "
   If donnée2.Recordset("dateint") <> "" Then
     Mvar4$ = donnée2.Recordset("dateint")
     MsComm1.Output = Mvar4$
   End If
' Technicien
   MsComm1.Output = Chr(31) + Chr(78) + Chr(65) + Chr(27) + Chr(67) + Chr(27) + Chr(93) + Chr(24)
   MsComm1.Output = "         Technicien    "
   If donnée2.Recordset("technicien") <> "" Then
     Mvar5$ = donnée2.Recordset("Technicien")
     MsComm1.Output = Mvar5$
   End If
' Descriptif de l'intervention
   MsComm1.Output = Chr(31) + Chr(80) + Chr(65) + "Descriptif de l'intervention" + Chr(10)
   MsComm1.Output = Chr(27) + Chr(92) + Chr(27) + Chr(71) + Chr(13) + Chr(96) + Chr(18) + Chr(103)
   MsComm1.Output = Chr(31) + Chr(82) + Chr(65) + Chr(27) + Chr(71) + Chr(27) + Chr(93) + Chr(24)
   MsComm1.Output = Chr(31) + Chr(83) + Chr(65) + Chr(27) + Chr(71) + Chr(27) + Chr(93) + Chr(24)
   MsComm1.Output = Chr(31) + Chr(84) + Chr(65) + Chr(27) + Chr(71) + Chr(27) + Chr(93) + Chr(24)
   MsComm1.Output = Chr(31) + Chr(85) + Chr(65) + Chr(27) + Chr(71) + Chr(27) + Chr(93) + Chr(24)
   MsComm1.Output = Chr(31) + Chr(86) + Chr(65) + Chr(27) + Chr(71) + Chr(27) + Chr(93) + Chr(24)
   MsComm1.Output = Chr(31) + Chr(87) + Chr(65) + Chr(27) + Chr(71) + Chr(27) + Chr(93) + Chr(24)
   MsComm1.Output = Chr(10) + Chr(27) + Chr(92) + Chr(27) + Chr(71) + Chr(13) + Chr(96) + Chr(18) + Chr(103)
   MsComm1.Output = Chr(31) + Chr(78) + Chr(65) + Chr(27) + Chr(67) + Chr(27) + Chr(93) + Chr(24)

End Sub

Private Function AttDTA$()
   While MsComm1.InBufferCount = 0 And MsComm1.CDHolding
   DoEvents
   Wend
   Char$ = MsComm1.Input
   If Len(Char$) > 0 Then
     AttDTA$ = Char$
   Else
     AttDTA$ = Chr(0)
   End If
End Function

Private Function Confirm%(code$)
   Select Case code$
       Case "*4222", "*0120", "*6742", "*9174"
          Confirm% = True
       Case Else
          Confirm% = False
   End Select
End Function

Private Sub Form_Load()
    ESC$ = Chr(27)
    Pro1$ = ESC$ + Chr(57)
    Pro3$ = ESC$ + Chr(58)
    Pro3$ = ESC$ + Chr(59)
    RC$ = Chr(13) + Chr(10)
    
    niveau% = 1
    MsComm1.CommPort = 2
    MsComm1.PortOpen = True
    MsComm1.DTREnable = True


End Sub

Private Sub ListeDep()
   For I = 1 To 10
      If couleur$ = Chr(67) Then
        couleur$ = Chr(70)
      Else
        couleur$ = Chr(67)
      End If
      MsComm1.Output = Chr(27) & couleur$ & Chr(27) & Chr(93) & Chr(24) & Chr(13)
      If Not (donnée2.Recordset.EOF) Then
          If donnée2.Recordset("Dateint") <> "" Then
            datei$ = donnée2.Recordset("Dateint")
          Else
            datei$ = "          "
          End If
          If donnée2.Recordset("Date_ap") <> "" Then
            datea$ = donnée2.Recordset("date_ap")
          Else
            datea$ = "          "
          End If
          MsComm1.Output = Format(donnée2.Recordset("numero"), "#0000") & " " & datea$ & " " & datei$ & " " & donnée2.Recordset("heureint") & " "
          If donnée2.Recordset("technicien") <> "" Then
            MsComm1.Output = donnée2.Recordset("technicien")
          Else
            MsComm1.Output = "   "
          End If
          donnée2.Recordset.MoveNext
      End If
      MsComm1.Output = Chr(13) & Chr(10)
   Next
End Sub

Private Sub maj()
   Echo% = True
   Select Case Mise_a_jour%
      Case 1
      Max% = 4
        MsComm1.Output = Chr(31) + Chr(73) + Chr(88) + Chr(17)
        If Len(Trim(texte1.Text)) = 4 Then
          Mvar1$ = Trim(texte1.Text)
          Mise_a_jour% = 2
          nd% = Len(texte1.Text)
          texte1.SelStart = 0
          texte1.SelLength = nd%
          texte1.SelText = ""
          Prog_P 1
        End If
              
      Case 2
        Max% = 2
        MsComm1.Output = Chr(31) + Chr(74) + Chr(88) + Chr(17)
        If Len(Trim(texte1.Text)) > 0 Then
          Mvar2$ = Trim(texte1.Text)
          Mise_a_jour% = 3
          nd% = Len(texte1.Text)
          texte1.SelStart = 0
          texte1.SelLength = nd%
          texte1.SelText = ""
          Prog_P 1
        End If

      Case 3
        Max% = 2
        MsComm1.Output = Chr(31) + Chr(75) + Chr(88) + Chr(17)
        If Len(Trim(texte1.Text)) > 0 Then
          Mvar3$ = Trim(texte1.Text)
          Mise_a_jour% = 4
          nd% = Len(texte1.Text)
          texte1.SelStart = 0
          texte1.SelLength = nd%
          texte1.SelText = ""
          Prog_P 1
        End If
             
      Case 4
        Max% = 10
        MsComm1.Output = Chr(31) + Chr(76) + Chr(88) + Chr(17)
        If Len(Trim(texte1.Text)) = 10 Then
          If IsDate(texte1.Text) Then
            Mvar4$ = Trim(texte1.Text)
            Mise_a_jour% = 5
          End If
          nd% = Len(texte1.Text)
          texte1.SelStart = 0
          texte1.SelLength = nd%
          texte1.SelText = ""
          Prog_P 1
        End If
        
      Case 5
        Max% = 3
        MsComm1.Output = Chr(31) + Chr(77) + Chr(88) + Chr(17)
        If Len(Trim(texte1.Text)) >= 1 Then
          Mvar5$ = Trim(texte1.Text)
          Mise_a_jour% = 6
          nd% = Len(texte1.Text)
          texte1.SelStart = 0
          texte1.SelLength = nd%
          texte1.SelText = ""
          Prog_P 1
        End If

      Case 6
        Max% = 240
        MsComm1.Output = Chr(31) + Chr(81) + Chr(65) + Chr(17)
        If Len(Trim(texte1.Text)) <> 0 Then
          Mvar6$ = Trim(texte1.Text)
          Mise_a_jour% = 7
          nd% = Len(texte1.Text)
          texte1.SelStart = 0
          texte1.SelLength = nd%
          texte1.SelText = ""
          MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + Chr(24) + "CONFIRMATION ?" + Chr(7)
          Prog_P 1
        End If

      Case 7
        Max% = 1
        If Len(texte1.Text) = 1 Then
          If texte1.Text = "O" Then
            '**************************
            '* Mise a jour du fichier *
            '**************************
            donnée2.Recordset.Edit
            donnée2.Recordset("Numba") = UCase(Mvar1$)
            donnée2.Recordset("Tempsmo") = UCase(Mvar2$)
            donnée2.Recordset("Tempsdep") = UCase(Mvar3$)
            donnée2.Recordset("Dateint") = UCase(Mvar4$)
            donnée2.Recordset("technicien") = UCase(Mvar5$)
            donnée2.UpdateRecord
            MsComm1.Output = Chr(12)
          End If
          Echo% = False
          annul% = False
          Mise_a_jour = 0
          Prog_P 6
        End If
   End Select
   nd% = Len(texte1.Text)
   texte1.SelStart = 0
   texte1.SelLength = nd%
   texte1.SelText = ""

End Sub

Private Sub Minuterie1_Timer()
'  If Format$(Now, "HH:MM") = "21H00" Then
'    If Not (MsComm1.CDHolding) Then
     ' Appel automatique
'     MsComm1.Output = "ATT0,0491931638;"
'     tempo = Timer
'     While Timer < tempo + 5: Wend
'     MsComm1.Output = "ATDRH0"
'    End If
'  End If
End Sub

Private Static Sub MSComm1_OnComm()
 Select Case MsComm1.CommEvent
   Case MSCOMM_EV_RING
         Echo% = False
         MsComm1.InputLen = 0
         MsComm1.Output = "ATA" + Chr(13)
           
   Case MSCOMM_EV_CD
         If MsComm1.CDHolding Then
           

             dta$ = " "
             Echo% = False
             numfich% = FreeFile
             Page_Acceuil$ = "c:\gestion\logo.txt"
             
             start& = Timer                           ' Attente de 5 secondes
             While start& + 5 > Timer: Wend

             Open Page_Acceuil$ For Binary As numfich%
             MsComm1.Output = Chr(12)
             Do While Not EOF(numfich%)
               Get numfich%, , dta$
               MsComm1.Output = dta$
             Loop
             
             Close numfich%
             MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + Chr(27) + Chr(93) + Chr(24) + "Pressez SOMMAIRE"
             MsComm1.Output = Chr(31) + Chr(88) + Chr(65) + "Copyright MORIANA Andre ETELEC          "
             niveau% = 1
             Max% = 12
             MsComm1.Output = Pro1$ + Chr$(&H70)
             nd% = Len(texte1.Text)
             texte1.SelStart = 0
             texte1.SelLength = nd%
             texte1.SelText = ""
          End If

   Case MSCOMM_EV_RECEIVE
     dta$ = MsComm1.Input                 'LECTURE DU CARATERE RECU SUR LE PORT SERIE

'**************************************************************************************
'*********                                                                   **********
'*********             RECEPTION DES CODES SPECIAUX MINITEL                  **********
'*********                                                                   **********
'**************************************************************************************

 If dta$ = Chr(&H13) Then
   
     dta$ = AttDTA$()
   
   If VarType(dta$) = 8 Then
     Select Case Asc(dta$)
        Case 108                             'TELEPHONE SONNE
        
        Case 65                              'TOUCHE ENVOI
               Prog_P 1

        Case 66                              'TOUCHE RETOUR
            
               Prog_P 2
        Case 67                              'TOUCHE REPETITION
               
               Prog_P 3
        Case 68                              'TOUCHE GUIDE
 
               Prog_P 4
        Case 69                              'TOUCHE ANNULATION
                                             
               Prog_P 5
        Case 70                              'TOUCHE SOMMAIRE
               Prog_P 6
        Case 71                              'TOUCHE CORRECTION

               nd% = Len(texte1.Text)
               If nd% > 0 Then
                 texte1.SelStart = nd% - 1
                 texte1.SelLength = 1
                 texte1.SelText = ""
                 MsComm1.Output = Chr(8) + Chr(24)
               End If
        Case 72                              'TOUCHE SUITE
          
               Prog_P 8
        Case 73                              'TOUCHE CONNEXION FIN
             Echo% = False
             numfich% = FreeFile
             Page_Acceuil$ = "c:\gestion\logo.txt"
             
             start& = Timer                           ' Attente de 5 secondes
             While start& + 5 > Timer: Wend

             Open Page_Acceuil$ For Binary As numfich%
             MsComm1.Output = Chr(12)
             Do While Not EOF(numfich%)
               Get numfich%, , dta$
               MsComm1.Output = dta$
             Loop
             
             Close numfich%
             MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + Chr(27) + Chr(93) + Chr(24) + "Pressez SOMMAIRE"
             MsComm1.Output = Chr(31) + Chr(88) + Chr(65) + "Copyright MORIANA Andre ETELEC          "
             niveau% = 1
             Max% = 12
             MsComm1.Output = Pro1$ + Chr$(&H70)
             nd% = Len(texte1.Text)
             texte1.SelStart = 0
             texte1.SelLength = nd%
             texte1.SelText = ""
'                Prog_P 9

        Case 82                            'CONNEXION DECONNEXION DU TELEPHONE

        Case 83                            'CONNEXION DECONNEXION DU MODEM
                                           'CHANGEMENT D ETAT FIL DP (détection porteuse)
        Case 84                            'CHANGEMENT D ETAT FIL PT
        Case 86                            'CHANGEMENT D ETAT DU STATUT MODE DE FONCTIONNEMENT
        Case 87                            'ACQUITEMENT DE LA MISE EN TRANSPARENCE
        Case 88                            'DEBUT ET FIN DE RETOURNEMENT
        Case 89                            'PHASE DE DECCONEXION


     End Select
   End If
'**************************************************************************************
'***********                                                                 **********
'***********               SEQUENCES DU LANGAGE PROTOCOLE                    **********
'***********                                                                 **********
'**************************************************************************************

 ElseIf dta$ = Chr(27) Then
     dta$ = AttDTA$()
     If VarType(dta$) = 8 Then
     If Asc(dta$) = 58 Then                     'PROTOCOLE 2 SUR 2 OCTETS
                                                '************************
     dta$ = AttDTA$()
        If dta$ = Chr(&H71) Then                            'REPONSE STATUT TERMINAL
                                                            '-----------------------
           dta$ = AttDTA$()
           REP_STAT_TERMINAL% = Asc(dta$)
        
        ElseIf dta$ = Chr(&H73) Then                        'REPONSE STATUT FONCTIONNEMENT
                                                            '-----------------------------
           dta$ = AttDTA$()
           REP_STAT_FONCTIONNEMENT% = Asc(dta$)

        ElseIf Asc(dta$) = &H75 Then                        'REPONSE STATUT VITESSE
                                                            '----------------------
           dta$ = AttDTA$()
           REP_STAT_VITESSE% = Asc(dta$)

        ElseIf Asc(dta$) = &H77 Then                        'REPONSE STATUT PROTOCOLE
                                                            '------------------------
           dta$ = AttDTA$()
           REP_STAT_PROTOCOLE% = Asc(dta$)

        End If
     ElseIf Asc(dta$) = 59 Then                 'PROTOCOLE 3 SUR 3 OCTETS
                                                '************************
           dta$ = AttDTA$()
        
        If Asc(dta$) = &H63 Then                            '&63h = FROM
                                                            '-----------
           dta$ = AttDTA$()

           ETAT$ = AttDTA$()
           
           Select Case Asc(dta$)
                                        'MODULE EMMISSION
                                        '*****************
               Case &H50                           ' ECRAN
                   CONNECT_EM_ECRAN% = Asc(ETAT$)
               Case &H51                           ' CLAVIER
                   CONNECT_EM_CLAVIER% = Asc(ETAT$)
               Case &H52                           ' MODEM
                   CONNECT_EM_MODEM% = Asc(ETAT$)
               Case &H53                           ' PRISE
                   CONNECT_EM_PRISE% = Asc(ETAT$)

                                        'MODULE RECEPTION
                                        '****************
               Case &H58                           ' ECRAN
                   CONNECT_REC_ECRAN% = Asc(ETAT$)
               Case &H59                           ' CLAVIER
                   CONNECT_REC_CLAVIER% = Asc(ETAT$)
               Case &H5A                           ' MODEM
                   CONNECT_REC_MODEM% = Asc(ETAT$)
               Case &H5B                           ' PRISE
                   CONNECT_REC_PRISE% = Asc(ETAT$)
           End Select
         End If
       End If
     End If
 Else
     If VarType(dta$) = 8 And Len(dta$) > 0 Then
       If 31 > Asc(dta$) < 128 Then

          nd% = Len(texte1.Text)
          texte1.SelStart = nd%
          texte1.SelText = dta$
          nd% = Len(texte1.Text)
          If Echo% Then MsComm1.Output = dta$
          If nd% = Max% Then Prog_P 0
       End If
     End If
 End If

End Select
End Sub

Private Sub Prog_P(commande%)
 On Error GoTo msgerreur
 Select Case niveau%
'**********************************************************************************
'**                                      NIVEAU 1                                **
'**********************************************************************************
    Case 1
      Select Case commande%
'******************************************* RAPPEL AUTOMATIQUE ** ENVOI***********
         Case 0, 1

            If Len(texte1.Text) <> 0 Then
              num$ = texte1.Text
              MsComm1.DTREnable = False
              While MsComm1.CDHolding: Wend
              start& = Timer
              While start& + 5 > Timer: Wend
               
              MsComm1.DTREnable = True
              start& = Timer
              While start& + 5 > Timer: Wend
              MsComm1.Output = "ATD0," & num$ & ";" & Chr(13)
              start& = Timer
              While start& + 20 > Timer: Wend
              MsComm1.Output = "ATDRH0" & Chr(13)
              
              niveau% = 1
            End If
'******************************************* SOMMAIRE ****************************
         Case 6
            
            niveau% = niveau% + 1
            Echo% = False
            Max% = 1
            nd% = Len(texte1.Text)
            texte1.SelStart = 0
            texte1.SelLength = nd%
            texte1.SelText = ""
           ' IDENTIFICATION
           '  MsComm1.InputLen = 0
           '  MsComm1.InBufferCount = 0
           '  MsComm1.Output = Chr(27) & Chr(57) & Chr(122)
           '  start& = Timer
           '  While start& + 5 > Timer: Wend
           '  ident$ = MsComm1.Input
           ' fin ident
            MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + Chr(27) + Chr(92) + Chr(24)
            MsComm1.Output = Chr(12) + Chr(10) + Chr(10) + Chr(10) + Chr(10) + Chr(10) + Chr(27) + Chr(&H4F) + "  SOMMAIRE"
            MsComm1.Output = Chr(10) + Chr(10) + Chr(10) + Chr(10) + Chr(13) + " 1->" + Chr(27) + Chr(&H4C) + " Fiches d'interventions"
            MsComm1.Output = Chr(10) + Chr(10) + Chr(13) + Chr(27) + Chr(&H4F) + " 2->" + Chr(27) + Chr(&H4C) + " Messagerie "


      End Select

'**********************************************************************************
'**                                      NIVEAU 2                                **
'**********************************************************************************
    Case 2
      Select Case commande%
'******************************************* ENVOI ********************************
        Case 1, 0
           choix% = Val(texte1.Text)
           Select Case choix%
               Case 1
                 niveau% = niveau% + 1
                 Prog_P 8
               Case 2
                 niveau% = 10
                 Prog_P 8
           End Select
           nd% = Len(texte1.Text)
           texte1.SelStart = 0
           texte1.SelLength = nd%
           texte1.SelText = ""
      End Select
'**********************************************************************************
'**                                      NIVEAU 3                                **
'**********************************************************************************

 Case 3
    Select Case commande%
'******************************************* SOMAIRE ******************************
 Case 6
    MsComm1.Output = Chr(20) + Chr(31) + Chr(64) + Chr(65) + " " + Chr(18) + Chr(103) + Chr(12)
    niveau% = niveau% - 2
    Prog_P 6

'******************************************* SUITE    *****************************
 Case 8
        Echo% = True
        MsComm1.Output = Chr(12) + Chr(&H1F) + Chr(65) + Chr(65) + Chr(24) + "CODE CLIENT " + Chr(17)
        nd% = Len(texte1.Text)
        texte1.SelStart = 0
        texte1.SelLength = nd%
        texte1.SelText = ""
        Max% = 11
        
'******************************************* ENVOI ********************************
 Case 0, 1
 MsComm1.Output = Chr(20)
 codecli$ = texte1.Text
 Donnée1.Recordset.FindFirst "cclient like '" & codecli$ & "'"
 If Not (Donnée1.Recordset.NoMatch) Then
' chargement du code ident$
   If codecli$ <> "*ET*" Then
    If Mid$(ident$, 2, 11) <> codecli$ Then
      MsgBox "Tentative de connection du code " & codecli$ & " non initialisé ", 4122, "ATTENTION!"
    End If
    MsComm1.Output = Pro1$ & Chr(&H79) & Chr(1) & codecli$ & Chr(4)
   End If
'
   MsComm1.Output = Chr(27) & Chr(&H44) & Chr(12)
   MsComm1.Output = Chr(10) & Donnée1.Recordset("Nclient") & Chr(10) & Chr(13)
   If Donnée1.Recordset("ADRC1") <> "" Then MsComm1.Output = Donnée1.Recordset("ADRC1")
   MsComm1.Output = Chr(10) & Chr(13)
   If Donnée1.Recordset("ADRC2") <> "" Then MsComm1.Output = Donnée1.Recordset("ADRC2")
   MsComm1.Output = Chr(10) & Chr(13)
   If Donnée1.Recordset("CODEPOST") <> "" Then MsComm1.Output = Donnée1.Recordset("Codepost") & "   "
   If Donnée1.Recordset("VILLEC") <> "" Then MsComm1.Output = Donnée1.Recordset("Villec")
   MsComm1.Output = Chr(10) & Chr(13)
   MsComm1.Output = Chr(13) & "Numero de Telephone : " & Donnée1.Recordset("Ntelc")
   MsComm1.Output = Chr(10) & Chr(13)
   MsComm1.Output = Chr(27) & Chr(71) & Chr(13) & Chr(10) & Chr(96) & Chr(18) & Chr(103)
   If Donnée1.Recordset("CCLIENT") Like "ET132970120" Then
     selection$ = ""
   Else
     selection$ = "WHERE CCLIENT = '" & Donnée1.Recordset("CCLIENT") & "'"
   End If
   donnée2.RecordSource = "select interv.* from interv " & selection$
   donnée2.Refresh
   niveau% = niveau% + 1
   Echo% = True
   Prog_P 8
 Else
   Prog_P 6
   nd% = Len(texte1.Text)
   texte1.SelStart = 0
   texte1.SelLength = nd%
   texte1.SelText = ""
 End If
End Select
'**********************************************************************************
'**                                      NIVEAU 4                                **
'**********************************************************************************
    Case 4
      Select Case commande%
'******************************************* SOMMAIRE ****************************
          Case 6
            MsComm1.Output = Chr(12)
            niveau% = niveau% - 3
            Prog_P 6
'******************************************* SUITE *******************************
          Case 8
             couleur$ = Chr(67)
             MsComm1.Output = Chr(10) & Chr(20) & Chr(31) & Chr(73) & Chr(65)

             ListeDep

             MsComm1.Output = Chr(27) & Chr(92) & Chr(27) & Chr(71) & Chr(13) & Chr(96) & Chr(18) & Chr(103) & Chr(10)
             MsComm1.Output = "Premiere page pressez        " & Chr(27) & Chr(93) & "  RETOUR  " & Chr(27) & Chr(92)
             MsComm1.Output = Chr(10) & Chr(13)
             MsComm1.Output = "Page suivante pressez        " & Chr(27) & Chr(93) & "  SUITE   " & Chr(27) & Chr(92)
             MsComm1.Output = Chr(13) & Chr(10) & "Creation d'une fiche presser " & Chr(27) & Chr(93) & "REPETITION" & Chr(27) & Chr(92)
             MsComm1.Output = Chr(13) & Chr(10) & "Detail fiche tapez son numero " & Chr(17)
             Echo% = True
             nd% = Len(texte1.Text)
             texte1.SelStart = 0
             texte1.SelLength = nd%
             texte1.SelText = ""
             Max% = 4
        Case 2
'******************************************* RETOUR ******************************
          If donnée2.Recordset.RecordCount > 0 Then donnée2.Recordset.MoveFirst
          Prog_P 8
'******************************************* ENVOI *******************************
        Case 0, 1
'          echo% = False
          If Trim(texte1.Text) <> "" Then AffDep
          nd% = Len(texte1.Text)
          texte1.SelStart = 0
          texte1.SelLength = nd%
          texte1.SelText = ""
        End Select
'**********************************************************************************
'**                               NIVEAU 5                                       **
'**********************************************************************************
  
  Case 5
    Select Case commande%
'******************************************* SOMMAIRE *****************************
      Case 6
         MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + Chr(24)
         MsComm1.Output = Chr(20) + Chr(27) + Chr(&H44) + Chr(12) + Chr(10)
         Donnée1.Recordset.FindFirst "cclient like '" & codecli$ & "'"
         MsComm1.Output = Donnée1.Recordset("Nclient") & Chr(10) & Chr(13)
         If Donnée1.Recordset("ADRC1") <> "" Then MsComm1.Output = Donnée1.Recordset("ADRC1")
         MsComm1.Output = Chr(10) & Chr(13)
         If Donnée1.Recordset("ADRC2") <> "" Then MsComm1.Output = Donnée1.Recordset("ADRC2")
         MsComm1.Output = Chr(10) & Chr(13)
         If Donnée1.Recordset("CODEPOST") <> "" Then MsComm1.Output = Donnée1.Recordset("Codepost") & "   "
         If Donnée1.Recordset("VILLEC") <> "" Then MsComm1.Output = Donnée1.Recordset("Villec")
         MsComm1.Output = Chr(10) & Chr(13)
         MsComm1.Output = Chr(13) & "Numero de Telephone : " & Donnée1.Recordset("Ntelc")
         MsComm1.Output = Chr(10) & Chr(13)
         MsComm1.Output = Chr(27) + Chr(71) + Chr(13) + Chr(10) + Chr(96) + Chr(18) + Chr(103)
         niveau% = niveau% - 1
         nd% = Len(texte1.Text)
         texte1.SelStart = 0
         texte1.SelLength = nd%
         texte1.SelText = ""
         Max% = 4
         Prog_P 8
'******************************************* ENVOI *******************************
      Case 0, 1
         If annul% Then
           maj
         End If
'******************************************* ANNULATION **************************
      Case 5
       If Not (annul%) Then
         Max% = 6
         Echo% = False
         Mise_a_jour% = 1
         AffMAJ
         annul% = True
         Prog_P 1
       End If

      End Select
'****************************************************************************
'****                                                                    ****
'****                          MESSAGERIE                                ****
'****                                                                    ****
'****************************************************************************
   Case 10
     choix% = 0
     Echo% = False
     lecture = True
     Max% = 1
     MsComm1.Output = Chr(12) + Chr(10) + Chr(10) + Chr(10) + Chr(10) + Chr(10) + Chr(27) + Chr(&H4F) + "  MESSAGERIE"
     MsComm1.Output = Chr(10) + Chr(10) + Chr(10) + Chr(10) + Chr(13) + " 1->" + Chr(27) + Chr(&H4C) + " Depot de message"
     MsComm1.Output = Chr(10) + Chr(10) + Chr(13) + Chr(27) + Chr(&H4F) + " 2->" + Chr(27) + Chr(&H4C) + " Lecture des messages "
     MsComm1.Output = Chr(10) + Chr(10) + Chr(13) + Chr(27) + Chr(&H4F) + " 3->" + Chr(27) + Chr(&H4C) + " Retour au sommaire "
     nd% = Len(texte1.Text)
     texte1.SelStart = 0
     texte1.SelLength = nd%
     texte1.SelText = ""
    
     niveau% = niveau% + 1
   Case 11
     Select Case commande%
       Case 0, 1
         If choix% = 0 Then choix% = Val(texte1.Text)
         Select Case choix%
             Case 1
             If lecture% Then
               lecture% = False
               niveau% = niveau% + 1
               MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + " Date depot : " + Date$ + "  " + Time$
               
               MsComm1.Output = Chr(12) + Chr(10) + Chr(27) + Chr(&H4F) + Chr(27) + Chr(66) + Chr(27) + Chr(93) + Chr(24)
               MsComm1.Output = "  DEPOT DE MESSAGE"
               MsComm1.Output = Chr(13) + Chr(10) + Chr(10) + Chr(27) + Chr(&H4C) + Chr(27) + Chr(66) + Chr(27) + Chr(92) + Chr(96) + Chr(18) + Chr(103)
               MsComm1.Output = Chr(13) + Chr(11) + Chr(27) + Chr(71) + "Destinatere"
               MsComm1.Output = Chr(13) + Chr(10) + Chr(27) + Chr(71) + Chr(27) + Chr(93) + Chr(24)
               MsComm1.Output = "      Nom:"
               MsComm1.Output = Chr(13) + Chr(10) + Chr(24)
               MsComm1.Output = "  Service:"
               MsComm1.Output = Chr(13) + Chr(10) + Chr(10) + Chr(27) + Chr(&H4C) + Chr(27) + Chr(66) + Chr(27) + Chr(92) + Chr(96) + Chr(18) + Chr(103)
               MsComm1.Output = Chr(13) + Chr(11) + Chr(27) + Chr(71) + "Expediteur"
               MsComm1.Output = Chr(13) + Chr(10) + Chr(27) + Chr(71) + Chr(27) + Chr(93) + Chr(24)
               MsComm1.Output = "      Nom:"
               MsComm1.Output = Chr(13) + Chr(10) + Chr(24)
               MsComm1.Output = "  Societe:"
               MsComm1.Output = Chr(13) + Chr(10) + Chr(24)
               MsComm1.Output = "Telephone:"
               MsComm1.Output = Chr(13) + Chr(10) + Chr(10) + Chr(27) + Chr(66) + Chr(27) + Chr(92) + Chr(96) + Chr(18) + Chr(103)
               MsComm1.Output = Chr(13) + Chr(11) + Chr(27) + Chr(71) + "Message"
               I = 1
               While I < 7
                 MsComm1.Output = Chr(13) + Chr(10) + Chr(27) + Chr(66) + Chr(27) + Chr(93) + Chr(24)
                 I = I + 1
               Wend
               MsComm1.Output = Chr(13) + Chr(10) + Chr(27) + Chr(66) + Chr(27) + Chr(92) + Chr(96) + Chr(18) + Chr(103)
               champ% = 1
               MsComm1.Output = "champ suivant presser la touche SUITE"
               MsComm1.Output = Chr(31) + Chr(69) + Chr(75) + Chr(17) + Chr(27) + Chr(71) + Chr(27) + Chr(93)
               Max% = 30
               Echo% = True
             End If
             Case 2
             If lecture% Then
               lecture% = False
               NumfMes% = FreeFile
               NumCourMes& = 0
               Open "c:\message.mes" For Random As NumfMes% Len = Len(FicheMes)
               NbreMes& = LOF(NumfMes%) / Len(FicheMes)
               
               Do
                 NumCourMes& = NumCourMes& + 1
                 If NumCourMes& <= NbreMes Then Get NumfMes%, NumCourMes&, FicheMes
               Loop While NumCourMes& <= NbreMes& And FicheMes.Eff = "*"
             If NumCourMes& <= NbreMes& Then
               MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + " Date depot : " + Mid$(FicheMes.DateAp, 7, 2) + "/" + Mid$(FicheMes.DateAp, 5, 2) + "/" + Mid$(FicheMes.DateAp, 3, 2) + "  " + FicheMes.HeureAp

               MsComm1.Output = Chr(12) + Chr(20) + Chr(10) + Chr(27) + Chr(&H4F) + Chr(27) + Chr(66) + Chr(27) + Chr(93) + Chr(24)
               MsComm1.Output = " LECTURE MESSAGES"
               MsComm1.Output = Chr(13) + Chr(10) + Chr(10) + Chr(27) + Chr(&H4C) + Chr(27) + Chr(66) + Chr(27) + Chr(92) + Chr(96) + Chr(18) + Chr(103)
               MsComm1.Output = Chr(13) + Chr(11) + Chr(27) + Chr(71) + "Destinatere"
               MsComm1.Output = Chr(13) + Chr(10) + Chr(27) + Chr(71) + Chr(27) + Chr(93) + Chr(24)
               MsComm1.Output = "      Nom:" + FicheMes.NomD + Chr(11)
               MsComm1.Output = Chr(13) + Chr(10) + Chr(24)
               MsComm1.Output = "  Service:" + FicheMes.Service + Chr(11)
               MsComm1.Output = Chr(13) + Chr(10) + Chr(10) + Chr(27) + Chr(&H4C) + Chr(27) + Chr(66) + Chr(27) + Chr(92) + Chr(96) + Chr(18) + Chr(103)
               MsComm1.Output = Chr(13) + Chr(11) + Chr(27) + Chr(71) + "Expediteur"
               MsComm1.Output = Chr(13) + Chr(10) + Chr(27) + Chr(71) + Chr(27) + Chr(93) + Chr(24)
               MsComm1.Output = "      Nom:" + FicheMes.NomE + Chr(11)
               MsComm1.Output = Chr(13) + Chr(10) + Chr(24)
               MsComm1.Output = "  Societe:" + FicheMes.Societe + Chr(11)
               MsComm1.Output = Chr(13) + Chr(10) + Chr(24)
               MsComm1.Output = "Telephone:" + FicheMes.Tel
               MsComm1.Output = Chr(13) + Chr(10) + Chr(10) + Chr(27) + Chr(66) + Chr(27) + Chr(92) + Chr(96) + Chr(18) + Chr(103)
               MsComm1.Output = Chr(13) + Chr(11) + Chr(27) + Chr(71) + "Message"
               I = 1
               While I < 7
                 MsComm1.Output = Chr(13) + Chr(10) + Chr(27) + Chr(66) + Chr(27) + Chr(93) + Chr(24)
                 I = I + 1
               Wend
               MsComm1.Output = Chr(31) + Chr(78) + Chr(65) + Chr(27) + Chr(66) + Chr(27) + Chr(93) + FicheMes.Message
               MsComm1.Output = Chr(27) + Chr(66) + Chr(27) + Chr(92) + Chr(96) + Chr(18) + Chr(103)
               MsComm1.Output = Chr(27) + Chr(71) + " PAGE SUIVANTE............." + Chr(27) + Chr(93) + "   SUITE  " + Chr(27) + Chr(92) + Chr(10) + Chr(13)
               MsComm1.Output = Chr(27) + Chr(71) + " PAGE PRECEDANTE..........." + Chr(27) + Chr(93) + "  RETOUR  " + Chr(27) + Chr(92) + Chr(10) + Chr(13)
               MsComm1.Output = Chr(27) + Chr(71) + " MISE A JOUR..............." + Chr(27) + Chr(93) + "ANNULATION" + Chr(27) + Chr(92) + Chr(10) + Chr(13)
               MsComm1.Output = Chr(27) + Chr(71) + " MENU GENERAL.............." + Chr(27) + Chr(93) + " SOMMAIRE " + Chr(27) + Chr(92)
             Else
               MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + "Plus de message" + Chr(7)
               niveau% = 10
               Prog_P 0


             End If
             End If
             
             Case 3
               MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + Space(40)
               lecture% = True
               choix% = 0
               niveau% = 1
               Prog_P 6
             

           End Select
         nd% = Len(texte1.Text)
         texte1.SelStart = 0
         texte1.SelLength = nd%
         texte1.SelText = ""
     Case 8
        If choix% = 2 And NumCourMes& <= NbreMes& Then
          
          Do
            NumCourMes& = NumCourMes& + 1
            If NumCourMes& <= NbreMes& Then Get NumfMes%, NumCourMes&, FicheMes
          Loop While NumCourMes& < NbreMes& + 1 And FicheMes.Eff = "*"
          If NumCourMes& <= NbreMes& Then
            MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + " Date depot : " + Mid$(FicheMes.DateAp, 7, 2) + "/" + Mid$(FicheMes.DateAp, 5, 2) + "/" + Mid$(FicheMes.DateAp, 3, 2) + "  " + FicheMes.HeureAp
            MsComm1.Output = Chr(31) + Chr(69) + Chr(75) + Chr(27) + Chr(93) + FicheMes.NomD + Chr(11)
            MsComm1.Output = Chr(31) + Chr(70) + Chr(75) + Chr(27) + Chr(93) + FicheMes.Service + Chr(11)
            MsComm1.Output = Chr(31) + Chr(73) + Chr(75) + Chr(27) + Chr(93) + FicheMes.NomE + Chr(11)
            MsComm1.Output = Chr(31) + Chr(74) + Chr(75) + Chr(27) + Chr(93) + FicheMes.Societe + Chr(11)
            MsComm1.Output = Chr(31) + Chr(75) + Chr(75) + Chr(27) + Chr(93) + FicheMes.Tel
            MsComm1.Output = Chr(31) + Chr(78) + Chr(65) + Chr(27) + Chr(66) + Chr(27) + Chr(93) + FicheMes.Message
          Else
            MsComm1.Output = Chr(31) + Chr(64) + Chr(98) + "FIN  " + Chr(7)
            Prog_P 2
          End If
        End If
     Case 2
        If choix% = 2 And NumCourMes& > 1 Then
          Do
            NumCourMes& = NumCourMes& - 1
            If NumCourMes& > 1 Then Get NumfMes%, NumCourMes&, FicheMes
          Loop While NumCourMes& > 0 And FicheMes.Eff = "*"
          If NumCourMes& > 0 Then
            MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + " Date depot : " + Mid$(FicheMes.DateAp, 7, 2) + "/" + Mid$(FicheMes.DateAp, 5, 2) + "/" + Mid$(FicheMes.DateAp, 3, 2) + "  " + FicheMes.HeureAp
            MsComm1.Output = Chr(31) + Chr(69) + Chr(75) + Chr(27) + Chr(93) + FicheMes.NomD + Chr(11)
            MsComm1.Output = Chr(31) + Chr(70) + Chr(75) + Chr(27) + Chr(93) + FicheMes.Service + Chr(11)
            MsComm1.Output = Chr(31) + Chr(73) + Chr(75) + Chr(27) + Chr(93) + FicheMes.NomE + Chr(11)
            MsComm1.Output = Chr(31) + Chr(74) + Chr(75) + Chr(27) + Chr(93) + FicheMes.Societe + Chr(11)
            MsComm1.Output = Chr(31) + Chr(75) + Chr(75) + Chr(27) + Chr(93) + FicheMes.Tel
            MsComm1.Output = Chr(31) + Chr(78) + Chr(65) + Chr(27) + Chr(66) + Chr(27) + Chr(93) + FicheMes.Message
          Else
            MsComm1.Output = Chr(31) + Chr(64) + Chr(98) + "DEBUT" + Chr(7)
            Prog_P 8
          End If
        End If
      Case 5
          MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + Chr(24) + "VOTRE CODE " + Chr(7)
          Max% = 5
          niveau% = 15
      Case 6
         MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + Space(40)
         lecture% = True
         choix% = 0
         niveau% = 10
         Prog_P 0
      
      End Select
     
   Case 12
     Select Case commande%
       
       Case 8, 0
          Select Case champ%
            Case 1
               Vmes1$ = texte1.Text
               champ% = champ% + 1
               MsComm1.Output = Chr(31) + Chr(70) + Chr(75) + Chr(17) + Chr(27) + Chr(71) + Chr(27) + Chr(93)
            Case 2
               VMes2$ = texte1.Text
               champ% = champ% + 1
               MsComm1.Output = Chr(31) + Chr(73) + Chr(75) + Chr(17) + Chr(27) + Chr(71) + Chr(27) + Chr(93)
            Case 3
               VMes3$ = texte1.Text
               champ% = champ% + 1
               MsComm1.Output = Chr(31) + Chr(74) + Chr(75) + Chr(17) + Chr(27) + Chr(71) + Chr(27) + Chr(93)
            Case 4
               VMes4$ = texte1.Text
               champ% = champ% + 1
               MsComm1.Output = Chr(31) + Chr(75) + Chr(75) + Chr(17) + Chr(27) + Chr(71) + Chr(27) + Chr(93)
          
            Case 5
               VMes5$ = texte1.Text
               champ% = champ% + 1
               MsComm1.Output = Chr(31) + Chr(78) + Chr(65) + Chr(17) + Chr(27) + Chr(66) + Chr(27) + Chr(93)
               Max% = 240
             
            Case 6
               VMes6$ = texte1.Text
               champ% = champ% + 1
               MsComm1.Output = Chr(31) + Chr(88) + Chr(65) + Chr(17) + Chr(27) + Chr(71) + Chr(27) + Chr(92)
               MsComm1.Output = "Confirmer pour enregistrement " + Chr(7)
               Max% = 1
             
            Case 7
              Echo% = False
              If texte1.Text = "O" Then
                 NumfMes% = FreeFile
                 Open "c:\message.mes" For Random As NumfMes% Len = Len(FicheMes)
                 NbreMes& = LOF(NumfMes%) / Len(FicheMes)
                 
                 FicheMes.Eff = " "
                 FicheMes.NomD = Vmes1$
                 FicheMes.Service = VMes2$
                 FicheMes.NomE = VMes3$
                 FicheMes.Societe = VMes4$
                 FicheMes.Tel = VMes5$
                 FicheMes.Message = VMes6$
                 FicheMes.DateAp = Mid$(Date$, 7, 4) + Mid$(Date$, 1, 2) + Mid$(Date$, 4, 2)
                 FicheMes.HeureAp = Left(Time$, 5)
                 Put NumfMes, NbreMes& + 1, FicheMes
                 
                 Close NumfMes%
               End If
               MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + Space(40) + Chr(20)
               niveau% = 10
               choix% = 0
               Prog_P 0

          End Select
         nd% = Len(texte1.Text)
         texte1.SelStart = 0
         texte1.SelLength = nd%
         texte1.SelText = ""
     
     End Select
   Case 15
'*******************************************************************************
'**                       ANNULATION MESSAGES                                 **
'*******************************************************************************
    Select Case commande%
       Case 0, 1
          MsComm1.Output = Chr(20)
          conf$ = texte1.Text
          If Confirm%(conf$) Then
            FicheMes.HeureMaj = texte1.Text
            FicheMes.DateMaj = Mid$(Date$, 7, 4) + Mid$(Date$, 1, 2) + Mid$(Date$, 4, 2)
            FicheMes.Eff = "*"
            Put NumfMes%, NumCourMes&, FicheMes.DateMaj
            Put NumfMes%, NumCourMes&, FicheMes.HeureMaj
            Put NumfMes%, NumCourMes&, FicheMes.Eff
          End If
          nd% = Len(texte1.Text)
          texte1.SelStart = 0
          texte1.SelLength = nd%
          texte1.SelText = ""
          MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + Space(40)
          niveau% = 10
          Prog_P 0
        End Select
'**************************************************************************************
'****                             AJOUT DEPANNAGE                                  ****
'**************************************************************************************
 Case 20
   Select Case commande%
     Case 0, 1 'envoi
       Select Case Ligne%
         Case 1
            'type intervention
         
         Case 2
            'liste des copro

         Case 3
            'descriptif de l intervention

         Case 4
            'validation

       End Select
     Case 2   'retour
       If Ligne% > 1 Then Ligne% = Ligne% - 1
     Case 6   'sommaire

     Case 8   'suite
       If Ligne% < 5 Then Ligne% = Ligne% + 1

   End Select
 End Select
Exit Sub
msgerreur:
      MsComm1.Output = Chr(31) + Chr(64) + Chr(65) + Chr(24) + "Ereur de saisie" + Chr(7)
      MsComm1.Output = Chr(31) + Chr(88) + Chr(65) + Chr(24)
      MsComm1.Output = "Detail fiche tapez son numero " + Chr(17)
      nd% = Len(texte1.Text)
      texte1.SelStart = 0
      texte1.SelLength = nd%
      texte1.SelText = ""
    Exit Sub

End Sub

Private Sub Texte1_KeyPress(keyascii As Integer)
  MsComm1.Output = Chr(keyascii)

End Sub

