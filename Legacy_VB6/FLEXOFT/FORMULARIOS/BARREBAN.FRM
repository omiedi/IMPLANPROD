VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form BARREBAN 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Revisa bandeja de Entrada"
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   2355
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   1920
      Top             =   1200
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   120
      Pattern         =   "*.tmp"
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1890
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   1365
      Top             =   105
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   735
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1365
      TabIndex        =   2
      Top             =   630
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   1
      Top             =   630
      Width           =   855
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   735
      TabIndex        =   0
      Top             =   210
      Width           =   1275
   End
End
Attribute VB_Name = "BARREBAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ESTABARRIENDO%
Dim TINIBA

Private Sub Form_Load()
   If App.PrevInstance = True Then
      Close: End
   End If
   Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
   Static VECES%
   If (VECES% Mod 15) = 0 Then
      VECES% = 1
      Call REVIBANDE
   End If
   VECES% = VECES% + 1
End Sub

Sub REVIBANDE()
   '
   Static CANCLI&, LANIDS$()
   '
   If CANCLI& < 1 Then
      TXTXX$ = LOADFILE$("F:\COTELEF\BADACLIE.DAT")
      CANCLI& = 1 + Len(TXTXX$) / 512
      ReDim LANIDS$(CANCLI&)
      For U& = 1 To CANCLI&
        TTYY$ = Mid$(TXTXX$, 1 + 512 * (U& - 1), 512)
        LANIDS$(U&) = Trim$(Mid$(TTYY$, 449, 16))
      Next U&
   End If
   '
   If ESTABARRIENDO% > 0 Then
      Exit Sub
   End If
   '
   ESTABARRIENDO% = 1
   TINIBA = Timer
   Label1 = "Barriendo ....."
   OPENFORMS = DoEvents
   '
   Dim DIRMA$()
   Screen.MousePointer = 11
   '
   MAPISession1.DownLoadMail = False
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   MAPIMessages1.FetchUnreadOnly = False
   MAPIMessages1.Fetch
   CAMENSA& = MAPIMessages1.MsgCount
   Screen.MousePointer = 11
   ESTABARRIENDO = 2
   '
   LSTUSER$ = LOADFILE$("F:\COTELEF\LSTUSERS.LST")
   HUBOCA% = 0
   Label3 = CAMENSA&
   For U& = 1 To CAMENSA      ' CAMENSA - 1000 To CAMENSA&
10   If U& > MAPIMessages1.MsgCount Then GoTo 99
     OPENFORMS = DoEvents
     Label2 = U&
     MAPIMessages1.MsgIndex = U& - 1
     ORENVI$ = Trim$(MAPIMessages1.MsgOrigDisplayName)
     ASUNTO$ = Trim$(LCase$(MAPIMessages1.MsgSubject))
     '
     CABORRA% = 0
     BBORRA$ = "VIAGRA": GoSub BORRAMEN
     BBORRA$ = "ROLEX": GoSub BORRAMEN
     BBORRA$ = "R0LEX": GoSub BORRAMEN
     BBORRA$ = "FACEBOOK": GoSub BORRAMEN
     BBORRA$ = "DISPENSER": GoSub BORRAMEN
     BBORRA$ = "CHEAP": GoSub BORRAMEN
     BBORRA$ = "CANADIAN": GoSub BORRAMEN
     BBORRA$ = "PHARMACY": GoSub BORRAMEN
     BBORRA$ = "SEX": GoSub BORRAMEN
     BBORRA$ = "VICODIN": GoSub BORRAMEN
     BBORRA$ = "ORÇAMENTO": GoSub BORRAMEN
     BBORRA$ = "SAUDE": GoSub BORRAMEN
     BBORRA$ = "RUSSIA": GoSub BORRAMEN
     BBORRA$ = "MEDS SHOP": GoSub BORRAMEN
     BBORRA$ = "PILLSTORE": GoSub BORRAMEN
     BBORRA$ = "PFIZER": GoSub BORRAMEN
     BBORRA$ = "CARNAVAL": GoSub BORRAMEN
     BBORRA$ = "WATCH": GoSub BORRAMEN
     BBORRA$ = "WATHES": GoSub BORRAMEN
     BBORRA$ = "PUBLICIDAD": GoSub BORRAMEN
     BBORRA$ = "AUTOBRONCEANTE": GoSub BORRAMEN
     BBORRA$ = "AMAZON": GoSub BORRAMEN
     BBORRA$ = "AGROPECUARIA": GoSub BORRAMEN
     BBORRA$ = "ACREDITADO": GoSub BORRAMEN
     BBORRA$ = "CHAMPION": GoSub BORRAMEN
     BBORRA$ = "SHOUT": GoSub BORRAMEN
     BBORRA$ = "LENGTH": GoSub BORRAMEN
     BBORRA$ = "CONFLICKER": GoSub BORRAMEN
     BBORRA$ = "DEGREE": GoSub BORRAMEN
     BBORRA$ = "ROCKET": GoSub BORRAMEN
     BBORRA$ = "DIPLOMA": GoSub BORRAMEN
     
     If CABORRA% > 0 Then
        U& = U& - CABORRA% - 1
        If U& < 1 Then U& = 1
     End If
     If U& > MAPIMessages1.MsgCount Then GoTo 99
     '
     If InStr(UCase$(ASUNTO$), "FLEXOFT WEB SERVICES") > 0 Then
        LANID$ = ""
        USID$ = ""
        EMPREX$ = ""
        USEX$ = ""
        TTXX$ = MAPIMessages1.MsgNoteText
        Do While TTXX$ <> ""
           PPXX% = InStr(TTXX$, vbCrLf)
           If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
           TTYY$ = Trim$(Left$(TTXX$, PPXX% - 1))
           TTXX$ = Mid$(TTXX$, PPXX% + 2)
           If Left$(UCase$(TTYY$), 8) = "EMPRESA:" Then EMPREX$ = Trim$(Mid$(TTYY$, 9))
           If Left$(UCase$(TTYY$), 7) = "LAN-ID:" Then LANID$ = Trim$(Mid$(TTYY$, 8))
           If Left$(UCase$(TTYY$), 8) = "USER-ID:" Then USID$ = Trim$(Mid$(TTYY$, 9))
           If Left$(UCase$(TTYY$), 9) = "USERNAME:" Then USEX$ = Trim$(Mid$(TTYY$, 10))
       Loop
       '
       BUSQUE$ = Left$(UCase$(LANID$ + Space$(16)), 16) + Left$(UCase$(USID$ + Space$(16)), 16)
       TTXX$ = BUSQUE$ + Left$(EMPREX$ + Space$(48), 48) + Left$(USEX$ + Space$(32), 32)
       '
       For KU& = 1 To Len(LSTUSER$) Step 128
          TTYY$ = Mid$(LSTUSER$, KU&, 128)
          If Left$(TTYY$, 32) = BUSQUE$ Then
             If TTYY$ <> TTXX$ Then
                Mid$(LSTUSER$, KU&, Len(TTXX$)) = TTXX$
                HUBOCA% = 1
             End If
             GoTo 89
          End If
       Next KU&
       TTXX$ = AJUSTI$(TTXX$, 128)
       LSTUSER$ = LSTUSER$ + TTXX$
       HUBOCA% = 1
       '
       If LANID$ <> "" Then
          For KU& = 1 To CANCLI&
             If LANIDS$(KU&) = LANID$ Then GoTo 89
          Next KU&
          MsgBox "Debe dar de Alta Empresa '" + EMPREX$ + "' con LAN_ID = '" + LANID$ + "'.", vbSystemModal
       End If
     End If
89 Next U&
   '
   If HUBOCA% > 0 Then
      Call SAVEFILE("F:\COTELEF\LSTUSERS.LST", LSTUSER$)
' SUBIR A LA PAGINA EL MAIL DE RECONOCIMIENTO
   End If
   '
99 MAPISession1.Action = 2  ' !!!
   ESTABARRIENDO% = 0
   Label1 = "En Pausa - Esperando"
   Label2 = ""
   Label3 = ""
   Screen.MousePointer = 1
Exit Sub
'
BORRAMEN:
   If InStr(UCase$(ASUNTO$), BBORRA$) > 0 Or InStr(UCase$(ORENVI$), BBORRA$) > 0 Then
      MAPIMessages1.Delete
      CABORRA% = CABORRA% + 1
   End If
Return
'
End Sub

Private Sub Timer2_Timer()
   If ESTABARRIENDO < 1 Then
       Label1 = Time$
     ElseIf ESTABARRIENDO = 1 Then
       Label1 = "Barriendo " & Int(Timer - TINIBA) & " ..."
   End If
   OPENFORMS = DoEvents
End Sub

Private Sub Timer3_Timer()
    '
    If ESTABARRIENDO% = 1 Then
       Exit Sub
    End If
    '
    LUCOMAI$ = "F:\FLEXOFT\MAILCTRL\"
    If EXISTE%(LUCOMAI$ + ".") > 0 Then
       File1.Path = LUCOMAI$
       File1.Refresh
       '
       For U& = 1 To File1.ListCount
         XX$ = LOADFILE$(LUCOMAI$ + File1.List(U& - 1))
         SMX$ = BUSCAVALOR(XX$, "REPEAT=")
         If Left$(SMX$, 4) = "FREC" Then
            CATAN = Val(Mid$(SMX$, 5, 3))
            TERVA$ = Mid$(SMX$, 8, 1)
            FEFIN = CDate(Mid$(SMX$, 13))
            '
            If CDate(Int(Now)) > FEFIN Then
               ' AQUI AGREGAR BORRADO DEL TMX
               On Error Resume Next
               Kill LUCOMAI + File1.List(U& - 1)
               On Error GoTo 0
               Exit Sub
            End If
            '
            FULTIEN = FileDateTime(LUCOMAI$ + File1.List(U& - 1))
            Select Case TERVA$
               Case "D"     ' DIAS
                 TRANSCU = (CDbl(Now) - CDbl(FULTIEN))
               Case "H"
                 TRANSCU = 60 * (CDbl(Now) - CDbl(FULTIEN))
               Case "M"
                 TRANSCU = 24 * 60 * (CDbl(Now) - CDbl(FULTIEN))
               Case Else
                 TRANSCU = 0
            End Select
            '
            If TRANSCU >= CATAN Then
               ' RE-ENVÍA EL MAIL
               PPXX% = InStr(File1.List(U& - 1), ".")
               If PPXX% < 1 Then PPXX% = 1 + Len(File1.List(U& - 1))
               CLA1$ = Left$(File1.List(U& - 1), PPXX% - 1)
               CLAX$ = CLA1$ + CONTRACLA$(CLA1$)
               CLAX$ = AJUSTI$(CLAX$, 13) + LUCOMAI$ + "/AUTO"
               On Error Resume Next
               ' EN \FLEXOFT\NOVEDAD estará seguramente la ultima version
               Shell "F:\FLEXOFT\NOVEDAD\FLXMAIL " + CLAX$, MODEJ%                    ' EJECUTA MINIMIZADO
               On Error GoTo 0
            End If
         End If
      Next U&
    End If
End Sub
'
