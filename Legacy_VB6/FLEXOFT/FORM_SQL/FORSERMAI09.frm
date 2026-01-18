VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form FORSERMAI09 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Revisa bandeja de Entrada"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6300
      Picture         =   "FORSERMAI09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cerrar y Salir"
      Top             =   105
      Width           =   600
   End
   Begin VB.CommandButton Command2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5670
      Picture         =   "FORSERMAI09.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Configuracion"
      Top             =   105
      Width           =   600
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5040
      Picture         =   "FORSERMAI09.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Re-Lanzar Ejecucion"
      Top             =   105
      Width           =   600
   End
   Begin VB.CommandButton Command18 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4410
      Picture         =   "FORSERMAI09.frx":09CE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Detener Ejecucion"
      Top             =   105
      Width           =   600
   End
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   3990
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3570
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   3150
      Top             =   0
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   600
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
      Top             =   420
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
      Left            =   210
      TabIndex        =   1
      Top             =   420
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
      Top             =   105
      Width           =   1275
   End
End
Attribute VB_Name = "FORSERMAI09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'agregar referencia microsoft outlook 11.0 object library

'Dim olApp As Outlook.Application
'Dim Inbox As Outlook.MAPIFolder
'Dim oOMail As Outlook.MailItem
'Dim InboxItems As Outlook.Items
Dim Mailobject As Object 'Atenuado como objeto Mailobject
Dim Subject As String 'Sin perjuicio de As String Dim
Dim Tomails As String

Dim ESTABARRIENDO%
Dim TINIBA


Private Sub Command1_Click()
   Command1.Enabled = False
   Command18.Enabled = True
   Timer1.Enabled = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   Timer1.Enabled = False
   ESTABARRIENDO% = 0
   Command1.Enabled = True
   Command2.Enabled = True
   Screen.MousePointer = 1
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   TTXX$ = "Criterios de Borrado"
   OPX% = ALTERNA%(TTXX$)
   If OPX% = 1 Then
       CRITBORR09.Show 1
   End If
   Command2.Enabled = True
End Sub

Private Sub Command3_Click()
   Close
   Unload Me
   End
End Sub

Private Sub Form_Load()
   If App.PrevInstance = True Then
      Close: End
   End If
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   Screen.MousePointer = 11
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   If EXISTE%(App.Path + "\CORREO_MO.EXE") < 1 Then
      TTXX$ = "Falta Aplicación 'CORREO_MO.EXE' en Carpeta\'" + REPLACAR$(App.Path, "\", "/") + "'\   \El Acceso a Microsoft Outlook\Deberá Autorizarse en Forma Manual.\   \Solicite e Instale el Componente Referido\para Automatizar el Proceso."
      Call MENSERR(24, TTXX$)
   End If
   '
   Call ABRECONEXION
   Timer1.Enabled = True
'   Call REVIBANDE
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Cancel = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   End
End Sub

Private Sub Timer1_Timer()
   Static VECES%
   If ESTABARRIENDO > 0 Then Exit Sub
   '
   If (VECES% Mod 2) = 0 Then
      VECES% = 1
      Call REVIBANDE
   End If
   VECES% = VECES% + 1
End Sub

Sub REVIBANDE()
   '
   Static YACRE%, CANCLI&, LANIDS$()
   '
   POROUTLOOK% = 1   '  1
   If ESTABARRIENDO% = 1 Then Exit Sub
   '
   If YACRE% < 1 Then
       Call CREACAMPO("", "BANPENPRO", "NUME_MAIL", 4, 4, 1)
       Call CREACAMPO("", "BANPENPRO", "IDEN_MAIL", 10, 128)
       Call CREACAMPO("", "BANPENPRO", "ORIG_MAIL", 10, 256)
       Call CREACAMPO("", "BANPENPRO", "DREL_MAIL", 10, 256)
       Call CREACAMPO("", "BANPENPRO", "ASUN_MAIL", 10, 256)
       Call CREACAMPO("", "BANPENPRO", "PARA_MAIL", 12, 0)
       Call CREACAMPO("", "BANPENPRO", "TEXT_MAIL", 12, 0)
       Call CREACAMPO("", "BANPENPRO", "HTMT_MAIL", 12, 0)
       Call CREACAMPO("", "BANPENPRO", "FHOR_RECEP", 8, 0, 1)       ' FECHA Y HORA QUE FUE PREPARADO
       Call CREACAMPO("", "BANPENPRO", "PROS_DEST", 10, 32, 1)
       Call CREACAMPO("", "BANPENPRO", "BORR_MAIL", 3, 2, 1)
       '
       Call CREACAMPO("", "ADJPENPRO", "NUME_MAIL", 4, 4, 1)
       Call CREACAMPO("", "ADJPENPRO", "NORD_MAIL", 3, 2, 1)
       Call CREACAMPO("", "ADJPENPRO", "FILE_NAME", 10, 256)
       '
       Call CREACAMPO("", "BANDENTRA", "NUME_MAIL", 4, 4, 1)
       Call CREACAMPO("", "BANDENTRA", "IDEN_MAIL", 10, 128)
       Call CREACAMPO("", "BANDENTRA", "Codi_Cli", 10, 12, 1)
       Call CREACAMPO("", "BANDENTRA", "ORIG_MAIL", 10, 256)
       Call CREACAMPO("", "BANDENTRA", "DREL_MAIL", 10, 256, 1)
       Call CREACAMPO("", "BANDENTRA", "ASUN_MAIL", 10, 256)
       Call CREACAMPO("", "BANDENTRA", "PARA_MAIL", 12, 0)
       Call CREACAMPO("", "BANDENTRA", "TEXT_MAIL", 12, 0)
       Call CREACAMPO("", "BANDENTRA", "HTMT_MAIL", 12, 0)
       Call CREACAMPO("", "BANDENTRA", "SOLI_CONF", 3, 2)
       Call CREACAMPO("", "BANDENTRA", "CONF_RECEP", 8, 0, 1)       ' FECHA Y HORA CONFIRMACION DE RECEPCION
       Call CREACAMPO("", "BANDENTRA", "FHOR_RECEP", 8, 0, 1)       ' FECHA Y HORA QUE FUE PREPARADO
       Call CREACAMPO("", "BANDENTRA", "USUA_DESTI", 3, 2, 1)       ' USUARIO QUE LO PREPARÓ
       Call CREACAMPO("", "BANDENTRA", "NUMAIRESPU", 4, 4, 1)       ' NUMERO DE MAIL DE RESPUESTA
       '
       Call CREACAMPO("", "ADJUENTRA", "NUME_MAIL", 4, 4, 1)
       Call CREACAMPO("", "ADJUENTRA", "NORD_MAIL", 3, 2, 1)
       Call CREACAMPO("", "ADJUENTRA", "FILE_NAME", 10, 256, 0)
       Call CREACAMPO("", "ADJUENTRA", "FILE_DATA", 12, 0)
       YACRE% = 1
   End If
   '
   If POROUTLOOK% > 0 Then
      Call REVIBANOUTL       ' VA POR EL MICROSOFT OUTLOOK
      Exit Sub
   End If
   '
   '
   ' AQUI SIGUE CON EL CONTROL MAPI
   '
   Screen.MousePointer = 11
   ESTABARRIENDO% = 1
   '
   MAPISession1.DownLoadMail = False     '  True
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   MAPIMessages1.FetchSorted = True
   MAPIMessages1.FetchUnreadOnly = True                              ' revisar este criterio
   MAPIMessages1.Fetch
   CAMENSA& = MAPIMessages1.MsgCount
   Screen.MousePointer = 11
   '
   OFXX = 120
   Label3 = CAMENSA&
6  For U& = 1 To OFXX     ' CAMENSA& - OFXX To CAMENSA&                                          ' deberia empezar de 1
      If ESTABARRIENDO% < 1 Then GoTo 99                                ' INTERRUMPIDO
      If U& < 1 Then GoTo 89
10    If U& > OFXX - 1 Then GoTo 90
      OPENFORMS = DoEvents
      Label2 = U&
      Me.Label3 = MAPIMessages1.MsgCount
12    MAPIMessages1.MsgIndex = MAPIMessages1.MsgCount - OFXX + U& '    U& - 1
      ORENVI$ = TRIM$(MAPIMessages1.MsgOrigDisplayName)
      DIELREM$ = TRIM$(MAPIMessages1.MsgOrigAddress)
         If Left$(DIELREM$, 5) = "SMTP:" Then DIELREM$ = Mid$(DIELREM$, 6)
      ASUNTO$ = TRIM$(LCase$(MAPIMessages1.MsgSubject))
      MESSTEXT$ = MAPIMessages1.MsgNoteText
      FRECEP = MAPIMessages1.MsgDateReceived
      SOLCONF% = MAPIMessages1.MsgReceiptRequested
      Tipo = MAPIMessages1.MsgType
      '
      IDENXX$ = REPLACAR$(MAPIMessages1.MsgID, "'", "-")    ' PARA EVITAR ERROR DE SINTAXIS
      IDENMA$ = REPLACAR$(IDENXX$, Chr$(34), "*")           ' PARA EVITAR ERROR DE SINTAXIS
          ' lo busca en los pendientes de procesar
          QTR& = CantRegistros("BANPENPRO", "IDEN_MAIL = '" & IDENMA$ & "'")
          YAENBANPEN% = 0
          If QTR& > 0 Then
             ' PROCESAR PARA GUARDAR O BORRAR
             Call PROCESAMAIL(IDENMA$, BORRAX%)
             If BORRAX% > 0 Then GoTo 10
             YAENBANPEN% = 1
            ' GoTo 89                         ' PASA AL SIGUIENTE
          End If
          '
          ' lo busca en los procesados
          CONDI$ = "ORIG_MAIL = ' " & ORENVI$ & "' "
          CONDI$ = CONDI$ + " AND DREL_MAIL = '" & DIELREM$ & "' "
          CONDI$ = CONDI$ + " AND ASUN_MAIL = '" & ASUNTO$ & "' "
          CONDI$ = CONDI$ + " AND TEXT_MAIL = '" & MESSTEXT$ & "'"
          QTR& = CantRegistros("BANDENTRA", CONDI$)
          If QTR& > 0 Then
             MAPIMessages1.Delete
             GoTo 10
          End If
          '
          PRINMAI09.Label4 = IDENMA$
          PRINMAI09.Label14 = ORENVI$
          PRINMAI09.Label16 = DIELREM$
          PRINMAI09.Label15 = ASUNTO$
      DIRELDEST$ = TRIM$(MAPIMessages1.RecipAddress)
      If Left$(DIRELDEST$, 5) = "SMTP:" Then DIRELDEST$ = Mid$(DIRELDEST$, 6)
      '
      CARECIP% = MAPIMessages1.RecipCount
      If CARECIP% > 1 Then
         DIRELDEST$ = ""
         For IRECIP% = 1 To CARECIP%
            MAPIMessages1.RecipIndex = IRECIP% - 1
            DIRDE$ = TRIM$(MAPIMessages1.RecipAddress)
            If Right$(DIRDE$, 1) = ";" Then DIRDE$ = Left$(DIRDE$, Len(DIRDE$) - 1)
            If DIRDE$ <> "" Then
               If DIRELDEST$ <> "" Then DIRELDEST$ = DIRELDEST$ + ";"
               DIRELDEST$ = DIRELDEST$ + DIRDE$
            End If
         Next IRECIP%
      End If
          PRINMAI09.Label17 = DIRELDEST$
          '
          PRINMAI09.Label2 = Format(FRECEP, "DD/MM/YY   hh:mm")
          PRINMAI09.Check1.Value = 0
          If SOLCONF% = True Then PRINMAI09.Check1.Value = 1
          PRINMAI09.Text1 = MESSTEXT
          '
      CANAD% = MAPIMessages1.AttachmentCount
          PRINMAI09.List3.Clear
      For KKI% = 1 To CANAD%
         MAPIMessages1.AttachmentIndex = KKI% - 1
         AA1$ = MAPIMessages1.AttachmentName
         aa2$ = MAPIMessages1.AttachmentPathName
         NOMPELADO$ = AA1$
            While InStr(NOMPELADO$, "\") > 0
               NOMPELADO$ = Mid$(NOMPELADO$, 2)
            Wend
            PRINMAI09.List3.AddItem AJUSTI$(NOMPELADO$, 128) + aa2$
         aa3 = MAPIMessages1.AttachmentPosition
         AA4 = MAPIMessages1.AttachmentType
      Next KKI%
      '
      ' aqui ya tengo cargados todos los datos del mensaje
      '
      If DIELREM$ = "" Then GoTo 75    ' QUE BORRE LOS QUE NO TIENEN REMITENTE

65    Call PROCESASEGUROS(OKX%, GUAMAI, GUADJU)
      If OKX% > 0 Then
          MAPIMessages1.Delete
          NUMAIPX& = XVALO(VALOBADA("BANPENPRO", "NUME_MAIL", "IDEN_MAIL = '" & IDENMA$ & "'", "NOMESS"))
          If NUMAIPIX& > 0 Then
             CONDI$ = "NUME_MAIL = " & NUMAIPX&
             Call BORRAREGISTROS("BANPENPRO", CONDI$, RAFE&, "NOMESS")
             Call BORRAREGISTROS("ADJPENPRO", CONDI$, RAFE&, "NOMESS")
          End If
          GoTo 10                ' NO INCREMENTA U& CON LO QUE CAE EN EL SIGUIENTE
      End If
      '
75    ' VUELVE A LEER TODOS LOS DATOS
      ORENVI$ = PRINMAI09.Label14
      DIELREM$ = PRINMAI09.Label16
      ASUNTO$ = PRINMAI09.Label15
      TXTX$ = PRINMAI09.Text1
      '
      OKX% = BORRAPORREGLAS%(ORENVI$, DIELREM$, ASUNTO$, TXTX$)
      If OKX% > 0 Then
          NUMAIPX& = XVALO(VALOBADA("BANPENPRO", "NUME_MAIL", "IDEN_MAIL = '" & IDENMA$ & "'", "NOMESS"))
          CONDI$ = "NUME_MAIL = " & NUMAIPX&
          Call BORRAREGISTROS("BANPENPRO", CONDI$, RAFE&, "NOMESS")
          Call BORRAREGISTROS("ADJPENPRO", CONDI$, RAFE&, "NOMESS")
          On Error Resume Next
          MAPIMessages1.Delete
          On Error GoTo 0
          GoTo 10              ' NO INCREMENTA U& CON LO QUE CAE EN EL SIGUIENTE
      End If
      '
      If YAENBANPEN% < 1 Then
         Call GUARDADATPRIN               ' PARA MOSTRAR LOS DATOS Y DEFINIR CRITERIO
      End If
      '
      Screen.MousePointer = 1
      '
89 Next U&
   '
90 OFXX = OFXX + 60
   GoTo 6
   '
99 MAPISession1.Action = 2  ' !!!
   ESTABARRIENDO% = 0
   Label1 = "En Pausa - Esperando"
   Label2 = ""
   Label3 = ""
   Screen.MousePointer = 1
   '
End Sub
'
Sub REVIBANOUTL()
   '
   Screen.MousePointer = 11
   ESTABARRIENDO% = 1
   '
   On Error Resume Next
   FORSERMAI09.Visible = True
   On Error GoTo 0
   ' Call LEVANTABLOQOUTL
   '
   Set olApp = CreateObject("Outlook.Application")
   Set Inbox = olApp.GetNamespace("Mapi").GetDefaultFolder(olFolderInbox) ' Establecer Bandeja de entrada = OlApp.GetNamespace ("Mapi"). GetDefaultFolder (olFolderInbox)
   Set InboxItems = Inbox.Items
   CAMENSA& = InboxItems.Count
   '
   ' PRIMERO BORRA DIRECTO POR REGLAS DE BORRADO
   U& = 0
   For Each Mailobject In InboxItems 'Para cada Mailobject En InboxItems
       '
       ' Call LEVANTABLOQOUTL
       '
       U& = U& + 1: Label2 = U&
       On Error Resume Next
       FORSERMAI09.Visible = True
       On Error GoTo 0
       OPENFORMS = DoEvents
       Me.Label3 = CAMENSA&
       If ESTABARRIENDO% < 1 Then GoTo 99
       '
       ' lee los datos del mensaje
       ORENVI$ = "": ASUNTO$ = "": DIELREM$ = "": TXTX$ = "": FRECEP = NUNCA
       '
       On Error Resume Next
       ASUNTO$ = TRIM$(Mailobject.Subject)
       ORENVI$ = TRIM$(Mailobject.SenderName)
       DIELREM$ = TRIM$(Mailobject.SenderEmailAddress)
       TXTX$ = Mailobject.Body
       FRECEP = Mailobject.ReceivedTime
       On Error GoTo 0
       '
       If ORENVI$ + DIELREM$ + TRIM$(TXTX$) = "" Then GoTo 29 ' ES PROBABLE ERROR DE LECTURA
       '
       BORRAX% = BORRAPORREGLAS%(ORENVI$, DIELREM$, ASUNTO$, TXTX$)
       If BORRAX% > 0 Then
          On Error Resume Next
          FRECEP = Mailobject.ReceivedTime
          On Error GoTo 0
          FRECEP1$ = Format$(FRECEP, "dd/mm/yy  hh:mm")
          FRECEP2$ = FETEXCOR1$(FECHANUM(Left$(FRECEP1$, 8))) + "  " + Mid$(FRECEP1$, 11)
          GoSub BUSCAENBANPENPRO
          '
          Mailobject.Delete
          If NUMAI& > 0 Then
             CONDIX$ = "NUME_MAIL = " & NUMAI& & ""
             Call BORRAREGISTROS("BANPENPRO", CONDIX$, RAFE&)
             Call BORRAREGISTROS("ADJPENPRO", CONDIX$, RAFE&)
          End If
       End If
29 Next
'   Set olApp = Nothing 'Establecer olApp = Nada
'   Set Inbox = Nothing 'Establecer Bandeja de entrada = Nothing
'   Set InboxItems = Nothing 'Establecer InboxItems Nada =
'   Set Mailobject = Nothing 'Establecer Mailobject = Nada
'   Set TempRst = Nothing 'Establecer TempRst = Nada
'   '
'   '
'   Set olApp = CreateObject("Outlook.Application")
'   Set Inbox = olApp.GetNamespace("Mapi").GetDefaultFolder(olFolderInbox) ' Establecer Bandeja de entrada = OlApp.GetNamespace ("Mapi"). GetDefaultFolder (olFolderInbox)
   Set InboxItems = Inbox.Items
   CAMENSA& = InboxItems.Count
   '
   U& = 0
   For Each Mailobject In InboxItems 'Para cada Mailobject En InboxItems
       U& = U& + 1: Label2 = U&
       Me.Label3 = CAMENSA&
       If ESTABARRIENDO% < 1 Then GoTo 99
       '
       ' lee los datos del mensaje
       GoSub LEETODOSLOSDATOS
       '
       If ORENVI$ + DIELREM$ + TRIM$(MESSTEXT$) + TRIM$(MESSHTML$) = "" Then GoTo 94 ' ES PROBABLE ERROR DE LECTURA       '
       ' Mailobject.UnRead = False
       ' TERMINO DE LEER LOS DATOS
       '
       ' BORRA LOS QUE NO TIENEN REMITENTE
       If DIELREM$ = "" Then
            Mailobject.Delete
            GoTo 89
       End If
       '
' If InStr(UCase$(DIELREM$), "ABADIE") > 0 Then
' AAA = BBB
' End If
       ' lo busca en BANDENTRA
       GoSub BUSCAENBANDEJADENTRADA
       If YAESTABAENBANDEJADENTRADA% = 1 Then
            Mailobject.Delete
            GoTo 89
       End If
       '
       ' lo busca en BANPENPRO
       GoSub BUSCAENBANPENPRO
       NUMAIPENPRO& = NUMAI&
       '
       If ESTABAPEN% > 0 Then
           If FIDESTI$ = "" And BORRAX% < 1 Then
              ' LEE TODOS LOS DATOS
              ORENVI$ = PRINMAI09.Label14
              DIELREM$ = PRINMAI09.Label16
              ASUNTO$ = PRINMAI09.Label15
              TXTX$ = PRINMAI09.Text1
              '
              BORRAX% = BORRAPORREGLAS%(ORENVI$, DIELREM$, ASUNTO$, TXTX$)
           End If
           '
           If FIDESTI$ = "" And BORRAX% > 0 Then
              Mailobject.Delete
              CONDIX$ = "NUME_MAIL = " & NUMAI& & ""
              Call BORRAREGISTROS("BANPENPRO", CONDIX$, RAFE&)
              Call BORRAREGISTROS("ADJPENPRO", CONDIX$, RAFE&)
              GoTo 89
           End If
           '
           If FIDESTI$ <> "" Then
               GoSub PROCESASEGUROS
               ' lo busca en BANDENTRA
               GoSub BUSCAENBANDEJADENTRADA
               If YAESTABAENBANDEJADENTRADA% = 1 Then
                    ' OJO CON LA CONDICION DEL NUMERO DE MAIL
                    If NUMAIPENPRO& > 0 Then
                       CONDIX$ = "NUME_MAIL = " & NUMAIPENPRO& & ""
                       Call BORRAREGISTROS("BANPENPRO", CONDIX$, RAFE&)
                       Call BORRAREGISTROS("ADJPENPRO", CONDIX$, RAFE&)
                       Mailobject.Delete
                       GoTo 89
                    End If
               End If
           End If
           '
       End If
       '
       If ESTABAPEN% < 1 Then
           GoSub GUARDAENPENPRO               ' PARA MOSTRAR LOS DATOS Y DEFINIR CRITERIO
       End If
'PRINMAI09.Show 1
       '
89     On Error Resume Next
       Me.Visible = True
       On Error GoTo 0
       OPENFORMS = DoEvents
94 Next
    
99 Set olApp = Nothing 'Establecer olApp = Nada
   Set Inbox = Nothing 'Establecer Bandeja de entrada = Nothing
   Set InboxItems = Nothing 'Establecer InboxItems Nada =
   Set Mailobject = Nothing 'Establecer Mailobject = Nada
   Set TempRst = Nothing 'Establecer TempRst = Nada
   ESTABARRIENDO% = 0
   Screen.MousePointer = 1
Exit Sub
   '
' ******************************************************************************************************
   '
LEETODOSLOSDATOS:
       '
       ORENVI$ = "": ASUNTO$ = "": DIELREM$ = "": FRECEP = NUNCA
       IDENZZ$ = "": SOLCONF% = 0: Tipo = "": MESSTEXT$ = "": MESSHTML$ = ""
       DIRELDEST$ = "": CARECIP% = 0
       '
       On Error Resume Next
       ORENVI$ = TRIM$(Mailobject.SenderName)
       ASUNTO$ = Mailobject.Subject
       DIELREM$ = Mailobject.SenderEmailAddress
       FRECEP = Mailobject.ReceivedTime
       FRECEP1$ = Format$(FRECEP, "dd/mm/yy  hh:mm")
       FRECEP2$ = FETEXCOR1$(FECHANUM(Left$(FRECEP1$, 8))) + "  " + Mid$(FRECEP1$, 11)
       IDENZZ$ = Mailobject.EntryID
         IDENXX$ = REPLACAR$(IDENZZ$, "'", "-")                ' PARA EVITAR ERROR DE SINTAXIS
         IDENMA$ = REPLACAR$(IDENXX$, Chr$(34), "*")           ' PARA EVITAR ERROR DE SINTAXIS
       SOLCONF% = Mailobject.ReadReceiptRequested
       Tipo = Mailobject.SenderEmailType
       MESSTEXT$ = Mailobject.Body
       MESSHTML$ = Mailobject.HTMLBody
       '
       ' AQUI HAY QUE IDENTIFICAR EL DESTINATARIO
       DIRELDEST$ = Mailobject.ReceivedByName
       If Left$(DIRELDEST$, 5) = "SMTP:" Then DIRELDEST$ = Mid$(DIRELDEST$, 6)
       CARECIP% = Mailobject.Recipients.Count
       On Error GoTo 0
       '
       If CARECIP% > 0 Then
         For IRECIP% = 1 To CARECIP%
            NORECIP$ = Mailobject.Recipients(IRECIP%).Name
            DIRDE$ = Mailobject.Recipients(IRECIP%).Address ' s.Index = IRECIP% - 1
            If Right$(DIRDE$, 1) = ";" Then DIRDE$ = Left$(DIRDE$, Len(DIRDE$) - 1)
            If Left$(DIRDE$, 5) = "SMTP:" Then DIRDE$ = Mid$(DIRDE$, 6)
            If DIRDE$ <> "" Then
               If DIRELDEST$ <> "" Then DIRELDEST$ = DIRELDEST$ + ";"
               DIRELDEST$ = DIRELDEST$ + DIRDE$
            End If
         Next IRECIP%
       End If
       '
       ' REGISTRA LOS DATOS EN PRINMAI09
30     PRINMAI09.Label4 = IDENMA$
       PRINMAI09.Label14 = ORENVI$
       PRINMAI09.Label16 = DIELREM$
       PRINMAI09.Label15 = ASUNTO$
       '
       PRINMAI09.Label2 = FRECEP1$
       PRINMAI09.Check1.Value = 0
          If SOLCONF% = True Then PRINMAI09.Check1.Value = 1
       PRINMAI09.Text1 = MESSTEXT$
       PRINMAI09.Text2 = MESSHTML$
       PRINMAI09.Label17 = DIRELDEST$
       '
       CANAD% = Mailobject.Attachments.Count
       PRINMAI09.List3.Clear
       For KKI% = 1 To CANAD%
          AA1$ = Mailobject.Attachments(KKI%).FileName
          ' Mailobject.Attachments(KKI%).SaveAsFile ("C:\FLEXOFT\" & AA1$)
          NOMPELADO$ = AA1$
          While InStr(NOMPELADO$, "\") > 0
             NOMPELADO$ = Mid$(NOMPELADO$, 2)
          Wend
          PRINMAI09.List3.AddItem AJUSTI$(NOMPELADO$, 128) + aa2$
       Next KKI%
       '
Return ' **********************************************************************************************
'
BUSCAENBANDEJADENTRADA:
       '
       YAESTABAENBANDEJADENTRADA% = 0
       CONDI$ = "ORIG_MAIL = '" & ORENVI$ & "' "
       CONDI$ = CONDI$ + " AND DREL_MAIL = '" & DIELREM$ & "' "
       CONDI$ = CONDI$ + " AND ASUN_MAIL = '" & ASUNTO$ & "' "
       CONDI$ = CONDI$ + " AND FHOR_RECEP = '" & FRECEP2$ & "' "
       QTR& = CantRegistros("BANDENTRA", CONDI$)
       If QTR& > 0 Then
          Call ACTUREGISTRO("BANDENTRA", "IDEN_MAIL", CONDI$, IDENMA$, RAFE&)
          Call ACTUREGISTRO("BANDENTRA", "HTMT_MAIL", CONDI$, MESSHTML$, RAFE&)
          '
          NUMAI& = XVALO(VALOBADA("BANDENTRA", "NUME_MAIL", CONDI$))
          For KKI% = 1 To PRINMAI09.List3.ListCount
             NADJUN$ = TRIM$(PRINMAI09.List3.List(KKI% - 1))
             CONDIX$ = "NUME_MAIL = " & NUMAI& & " AND FILE_NAME = '" & NADJUN$ & "'"
             If TRIM$(VALOBADA("ADJUENTRA", "FILE_DATA", CONDIX$)) = "" Then GoTo 20
          Next KKI%
          YAESTABAENBANDEJADENTRADA% = 1
       End If
       '
20 Return ' *********************************************************************************************

BUSCAENBANPENPRO:
       '
       ESTABAPEN% = 0: FIDESTI$ = "": BORRAX% = 0: NUMAI& = 0
       '
       '  busca en pendientes por caracterisicas
       CONDI$ = "ORIG_MAIL = '" & ORENVI$ & "' "
       CONDI$ = CONDI$ + " AND DREL_MAIL = '" & DIELREM$ & "' "
       CONDI$ = CONDI$ + " AND ASUN_MAIL = '" & ASUNTO$ & "' "
       CONDI$ = CONDI$ + " AND FHOR_RECEP = '" & FRECEP2$ & "' "
       QTR& = CantRegistros("BANPENPRO", CONDI$)
       If QTR& > 0 Then
          Call ACTUREGISTRO("BANPENPRO", "IDEN_MAIL", CONDI$, IDENMA$, RAFE&)
          Call ACTUREGISTRO("BANPENPRO", "HTMT_MAIL", CONDI$, MESSHTML$, RAFE&)
          '
          ESTABAPEN% = 1
          FIDESTI$ = TRIM$(VALOBADA("BANPENPRO", "PROS_DEST", CONDI$))
          BORRAX% = XVALO(VALOBADA("BANPENPRO", "BORR_MAIL", CONDI$))
          NUMAI& = XVALO(VALOBADA("BANPENPRO", "NUME_MAIL", CONDI$))
          '
          If FIDESTI$ = "" Then
             If BORRAX% < 1 Then
                FIDESTI$ = CODPROSPE$(DIELREM$)
                Call ACTUREGISTRO("BANPENPRO", "PROS_DEST", CONDI$, FIDESTI$, RAFE&)
             End If
          End If
       End If

Return ' *********************************************************************************************

PROCESASEGUROS:
    '
    ' LEE TODOS LOS DATOS
    ORENVI$ = PRINMAI09.Label14
    DIELREM$ = PRINMAI09.Label16
    DIELDES$ = PRINMAI09.Label17
    ASUNTO$ = PRINMAI09.Label15
    IDMA$ = PRINMAI09.Label4
    FRECEP = PRINMAI09.Label2
       FEXI% = FECHANUM(Left$(FRECEP, 8))
       FANOTA% = FEXI%
       FRECEP = FETEXCOR1$(FEXI%) + " " + Mid$(FRECEP, 10)
    SOLCONF% = PRINMAI09.Check1.Value
    TXTX$ = PRINMAI09.Text1
    TYTY$ = PRINMAI09.Text2
    '
    USDESTI% = 1
    If InStr(UCase$(DIELDES$), "COMERCIAL") > 0 Then USDESTI% = 2      ' JORGE REAL
    If InStr(UCase$(DIELDES$), "ADMINISTRACION") > 0 Then USDESTI% = 3 ' LORENA
    If InStr(UCase$(DIELDES$), "ASISTEC") > 0 Then USDESTI% = 5        ' OMAR
    If InStr(UCase$(DIELDES$), "FGRASSO") > 0 Then USDESTI% = 4        ' FLAVIO
    '
    ' OBTIENE CODIGO DE PROSPECTO REMITENTE
    ORIMESS$ = FIDESTI$
    '
    NUPROMAI& = 1 + XVALO(VALOBADA("BANDENTRA", "MAX(NUME_MAIL)", "NUME_MAIL > 0", "NOMESS"))
    CONDI$ = "NUME_MAIL = " & NUPROMAI&
    Call CreaRegistro("BANDENTRA", "NUME_MAIL", NUPROMAI&, "NOMESS")
    '
    Call ACTUREGISTRO("BANDENTRA", "NUME_MAIL", CONDI$, NUPROMAI&, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "IDEN_MAIL", CONDI$, IDMA$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "Codi_Cli", CONDI$, ORIMESS$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "ORIG_MAIL", CONDI$, Left$(ORENVI$, 256), RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "DREL_MAIL", CONDI$, Left$(DIELREM$, 256), RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "ASUN_MAIL", CONDI$, Left$(ASUNTO$, 256), RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "PARA_MAIL", CONDI$, DIELDES$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "TEXT_MAIL", CONDI$, TXTX$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "HTMT_MAIL", CONDI$, TYTY$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "SOLI_CONF", CONDI$, SOLCONF%, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "FHOR_RECEP", CONDI$, FRECEP, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "USUA_DESTI", CONDI$, USDESTI%, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "CONF_RECEP", CONDI$, NUNCA$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "NUMAIRESPU", CONDI$, 0, RAFE&, "NOMESS")
    '
    ' GUARDA LOS ADJUNTOS
    Call BORRAREGISTROS("ADJUENTRA", CONDI$, RAFE&)
    CANAD% = Mailobject.Attachments.Count
    For KKI% = 1 To CANAD%
         FIFUNA$ = Mailobject.Attachments(KKI%).FileName
         If FIFUNA$ <> "" Then
7           On Error Resume Next
            Mailobject.Attachments(KKI%).SaveAsFile ("C:\FLEXOFT\XXX.TMP")
            If Err Then
               On Error GoTo 0
               TXTXT$ = "Imposible Guardar Archivo Adjunto '" + FIFUNA$ + "'"
               TXTXT$ = TXTXT$ + vbCrLf + "como 'C:\FLEXOFT\XXX.TMP'"
               TXTXT$ = TXTXT$ + vbCrLf + vbCrLf + "Puede que el Archivo Destino Sea de Solo Lectura"
               TXTXT$ = TXTXT$ + vbCrLf + "o Esté Bloqueado en Uso por Otro Proceso."
               AAX% = MsgBox(TXTXT$, vbAbortRetryIgnore + vbDefaultButton2, "Mensaje Critico")
               If AAX% = 4 Then GoTo 7
               If AAX% = 5 Then GoTo 19
               Close: End
            End If
            On Error GoTo 0
            SetAttr "C:\FLEXOFT\XXX.TMP", vbNormal     ' PARA QUE DESPUES SE PUEDA BORRAR
            XXX$ = LOADFILE$("C:\FLEXOFT\XXX.TMP")
            NOMPELADO$ = FIFUNA$: While InStr(NOMPELADO$, "\") > 0: NOMPELADO$ = Mid$(NOMPELADO$, 2): Wend
            '
            If UCase$(Left$(XXX$, 14)) = "NO VIRUS FOUND" Then GoTo 19
            If UCase$(Left$(XXX$, 23)) = "INTERNAL VIRUS DATABASE" Then GoTo 19
            If XXX$ <> "" Then
10             NUMAPRO& = XVALO(RANDSTRING$(-6))
               If CantRegistros("ADJUENTRA", "NUME_MAIL = " & NUMAPRO&, "NOMESS") > 0 Then GoTo 10
               Call CreaRegistro("ADJUENTRA", "NUME_MAIL", NUMAPRO&, "NOMESS")
               '
               CONDI$ = "NUME_MAIL = " & NUPROMAI&
               NORDE& = 1 + XVALO(VALOBADA("ADJUENTRA", "MAX(NORD_MAIL)", CONDI$, "NOMESS"))
               Call ACTUREGISTRO("ADJUENTRA", "NORD_MAIL", "NUME_MAIL = " & NUMAPRO&, NORDE&, RAFE&, "NOMESS")
               Call ACTUREGISTRO("ADJUENTRA", "FILE_NAME", "NUME_MAIL = " & NUMAPRO&, FIFUNA$, RAFE, "NOMESS")
               Call ACTUREGISTRO("ADJUENTRA", "FILE_DATA", "NUME_MAIL = " & NUMAPRO&, XXX$, RAFE&, "NOMESS")
               Call ACTUREGISTRO("ADJUENTRA", "NUME_MAIL", "NUME_MAIL = " & NUMAPRO&, NUPROMAI&, RAFE&, "NOMESS")
            End If
         End If
19  Next KKI%
    '
    ' ANOTA EN SEGUIMIENTO COMERCIAL
    TTTT$ = "Recepcion de Mail ($" & TRIM$(Str$(NUPROMAI&)) & "$):" + CRLF$ + CRLF$
    TTTT$ = TTTT$ + " - De: " + ORENVI$ + CRLF$
    TTTT$ = TTTT$ + " - Para: " + DIELDES$ + CRLF$
    TTTT$ = TTTT$ + " - Asunto: " + ASUNTO$ + CRLF$ + CRLF$
    Call FISECOM09.ANOTAHISTORIAL(TTTT$, ORIMESS$, "", FANOTA%)
    '
    ' PONE AVISO EMERGENTE AL DESTINATARIO
    '
Return ' *********************************************************************************************

GUARDAENPENPRO:
    '
    ' LEE TODOS LOS DATOS
    ORENVI$ = PRINMAI09.Label14
    DIELREM$ = PRINMAI09.Label16
    DIELDES$ = PRINMAI09.Label17
    ASUNTO$ = PRINMAI09.Label15
    IDMA$ = PRINMAI09.Label4
    FRECEP = PRINMAI09.Label2
       FEXI% = FECHANUM(Left$(FRECEP, 8))
       FANOTA% = FEXI%
       FRECEP = FETEXCOR1$(FEXI%) + " " + Mid$(FRECEP, 10)
    SOLCONF% = PRINMAI09.Check1.Value
    TXTX$ = PRINMAI09.Text1
    TYTY$ = PRINMAI09.Text2
    '
    ' GUARDA MAIL EN BANDENTRA
    NUPROMAI& = 1 + XVALO(VALOBADA("BANPENPRO", "MAX(NUME_MAIL)", "NUME_MAIL > 0", "NOMESS"))
    CONDI$ = "NUME_MAIL = " & NUPROMAI&
    Call CreaRegistro("BANPENPRO", "NUME_MAIL", NUPROMAI&, "NOMESS")
    '
    Call ACTUREGISTRO("BANPENPRO", "IDEN_MAIL", CONDI$, IDENMA$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "ORIG_MAIL", CONDI$, Left$(ORENVI$, 256), RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "DREL_MAIL", CONDI$, Left$(DIELREM$, 256), RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "ASUN_MAIL", CONDI$, Left$(ASUNTO$, 256), RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "PARA_MAIL", CONDI$, DIELDES$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "TEXT_MAIL", CONDI$, TXTX$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "HTMT_MAIL", CONDI$, TYTY$, RAFE&)
    Call ACTUREGISTRO("BANPENPRO", "FHOR_RECEP", CONDI$, FRECEP, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "PROS_DEST", CONDI$, "", RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "BORR_MAIL", CONDI$, 0, RAFE&, "NOMESS")
    '
    ' GUARDA LOS NOMBRES DE LOS ADJUNTOS
    CANADJ% = PRINMAI09.List3.ListCount
    For KKII% = 1 To CANADJ%
       FIFUNA$ = TRIM$(Left$(PRINMAI09.List3.List(KKII% - 1), 128))
11     NUMAPRO& = XVALO(RANDSTRING$(-6))
       If CantRegistros("ADJPENPRO", "NUME_MAIL = '" & NUMAPRO& & "'", "NOMESS") > 0 Then GoTo 11
       Call CreaRegistro("ADJPENPRO", "NUME_MAIL", NUMAPRO&, "NOMESS")
       '
       CONDI$ = "NUME_MAIL = " & NUPROMAI&
       NORDE& = 1 + XVALO(VALOBADA("ADJPENPRO", "MAX(NORD_MAIL)", CONDI$, "NOMESS"))
       Call ACTUREGISTRO("ADJPENPRO", "NORD_MAIL", "NUME_MAIL = " & NUMAPRO&, NORDE&, RAFE&, "NOMESS")
       Call ACTUREGISTRO("ADJPENPRO", "FILE_NAME", "NUME_MAIL = " & NUMAPRO&, FIFUNA$, RAFE, "NOMESS")
       Call ACTUREGISTRO("ADJPENPRO", "NUME_MAIL", "NUME_MAIL = " & NUMAPRO&, NUPROMAI&, RAFE&, "NOMESS")
    Next KKII%
    '
Return

End Sub
'
Sub LEVANTABLOQOUTL()
    '
    Static EXINOB%, TULAN
    '
Exit Sub
    If EXINOB% = 0 Then
       EXINOB% = (-1)
       If EXISTE%(App.Path + "\CORREO_MO.EXE") > 0 Then
          EXINOB% = 1
       End If
    End If
    '
    ' LANZA LIBERADOR DE BLOQUEOS DE OUTLOOK
    If EXINOB% = 1 Then
       If Timer - TULAN > 3 Then    ' CADA 3 SEGUNDOS
          TULAN = Timer
          Shell App.Path + "\CORREO_MO.EXE"
       End If
    End If
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
 Exit Sub
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
Sub PROCESASEGUROS(OKX%, Optional GUARMAIL, Optional GUARADJUN)
    '
    Static YACRE%
    OKX% = 0                      ' DEVUELVE 1 SI LO PROCESO COMO MAIL SEGURO
    '
    ' LEE TODOS LOS DATOS
    ORENVI$ = PRINMAI09.Label14
    DIELREM$ = PRINMAI09.Label16
    DIELDES$ = PRINMAI09.Label17
    ASUNTO$ = PRINMAI09.Label15
    IDMA$ = PRINMAI09.Label4
    FRECEP = PRINMAI09.Label2
       FEXI% = FECHANUM(Left$(FRECEP, 8))
       FANOTA% = FEXI%
       FRECEP = FETEXCOR1$(FEXI%) + " " + Mid$(FRECEP, 10)
    SOLCONF% = PRINMAI09.Check1.Value
    TXTX$ = PRINMAI09.Text1
    TYTY$ = PRINMAI09.Text2
    '
    USDESTI% = 1
    If InStr(UCase$(DIELDES$), "COMERCIAL") > 0 Then USDESTI% = 2      ' JORGE REAL
    If InStr(UCase$(DIELDES$), "ADMINISTRACION") > 0 Then USDESTI% = 3 ' LORENA
    If InStr(UCase$(DIELDES$), "ASISTEC") > 0 Then USDESTI% = 5        ' OMAR
    If InStr(UCase$(DIELDES$), "FGRASSO") > 0 Then USDESTI% = 4        ' FLAVIO
    
    ' OBTIENE CODIGO DE PROSPECTO REMITENTE
    ORIMESS$ = CODPROSPE$(DIELREM$)
    If ORIMESS$ = "" Then
        OKX% = 0
        Exit Sub
    End If
    '
    If YACRE% < 1 Then
       ' GUARDA MAIL EN BANDENTRA
       Call CREACAMPO("", "BANDENTRA", "NUME_MAIL", 4, 4, 1)
       Call CREACAMPO("", "BANDENTRA", "IDEN_MAIL", 10, 128)
       Call CREACAMPO("", "BANDENTRA", "Codi_Cli", 10, 12, 1)
       Call CREACAMPO("", "BANDENTRA", "ORIG_MAIL", 10, 256)
       Call CREACAMPO("", "BANDENTRA", "DREL_MAIL", 10, 256, 1)
       Call CREACAMPO("", "BANDENTRA", "ASUN_MAIL", 10, 256)
       Call CREACAMPO("", "BANDENTRA", "PARA_MAIL", 12, 0)
       Call CREACAMPO("", "BANDENTRA", "TEXT_MAIL", 12, 0)
       Call CREACAMPO("", "BANDENTRA", "HTMT_MAIL", 12, 0)
       Call CREACAMPO("", "BANDENTRA", "SOLI_CONF", 3, 2)
       Call CREACAMPO("", "BANDENTRA", "CONF_RECEP", 8, 0, 1)       ' FECHA Y HORA CONFIRMACION DE RECEPCION
       Call CREACAMPO("", "BANDENTRA", "FHOR_RECEP", 8, 0, 1)       ' FECHA Y HORA QUE FUE PREPARADO
       Call CREACAMPO("", "BANDENTRA", "USUA_DESTI", 3, 2, 1)       ' USUARIO QUE LO PREPARÓ
       Call CREACAMPO("", "BANDENTRA", "NUMAIRESPU", 4, 4, 1)       ' NUMERO DE MAIL DE RESPUESTA
       '
       Call CREACAMPO("", "ADJUENTRA", "NUME_MAIL", 4, 4, 1)
       Call CREACAMPO("", "ADJUENTRA", "NORD_MAIL", 3, 2, 1)
       Call CREACAMPO("", "ADJUENTRA", "FILE_NAME", 10, 256, 0)
       Call CREACAMPO("", "ADJUENTRA", "FILE_DATA", 12, 0)
       YACRE% = 1
    End If
    '
    If GUARMAIL < 0 Then Exit Sub
    '
    NUPROMAI& = 1 + XVALO(VALOBADA("BANDENTRA", "MAX(NUME_MAIL)", "NUME_MAIL > 0", "NOMESS"))
    CONDI$ = "NUME_MAIL = " & NUPROMAI&
    Call CreaRegistro("BANDENTRA", "NUME_MAIL", NUPROMAI&, "NOMESS")
    '
    Call ACTUREGISTRO("BANDENTRA", "NUME_MAIL", CONDI$, NUPROMAI&, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "IDEN_MAIL", CONDI$, IDMA$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "Codi_Cli", CONDI$, ORIMESS$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "ORIG_MAIL", CONDI$, Left$(ORENVI$, 256), RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "DREL_MAIL", CONDI$, Left$(DIELREM$, 256), RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "ASUN_MAIL", CONDI$, Left$(ASUNTO$, 256), RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "PARA_MAIL", CONDI$, DIELDES$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "TEXT_MAIL", CONDI$, TXTX$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "HTMT_MAIL", CONDI$, TYTY$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "SOLI_CONF", CONDI$, SOLCONF%, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "FHOR_RECEP", CONDI$, FRECEP, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "USUA_DESTI", CONDI$, USDESTI%, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "CONF_RECEP", CONDI$, NUNCA, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANDENTRA", "NUMAIRESPU", CONDI$, 0, RAFE&, "NOMESS")
    '
    ' ANOTA EN SEGUIMIENTO COMERCIAL
    TTTT$ = "Recepcion de Mail ($" & TRIM$(Str$(NUPROMAI&)) & "$):" + CRLF$ + CRLF$
    TTTT$ = TTTT$ + " - De: " + ORENVI$ + CRLF$
    TTTT$ = TTTT$ + " - Para: " + DIELDES$ + CRLF$
    TTTT$ = TTTT$ + " - Asunto: " + ASUNTO$ + CRLF$ + CRLF$
    Call FISECOM09.ANOTAHISTORIAL(TTTT$, ORIMESS$, "", FANOTA%)
    '
'    If GUARADJUN < 0 Then GoTo 30
'    '
    ' GUARDA LOS ADJUNTOS
    CANADJ% = PRINMAI09.List3.ListCount
    For KKII% = 1 To CANADJ%
       AFUNA$ = Mid$(PRINMAI09.List3.List(KKII% - 1), 129)
       FIFUNA$ = AFUNA$
       While InStr(FIFUNA$, "\") > 0
           FIFUNA$ = Mid$(FIFUNA$, 2)
       Wend
       XXX$ = LOADFILE$(AFUNA$)
       If UCase$(Left$(XXX$, 14)) = "NO VIRUS FOUND" Then GoTo 19
       If UCase$(Left$(XXX$, 23)) = "INTERNAL VIRUS DATABASE" Then GoTo 19
       If XXX$ <> "" Then
10        NUMAPRO& = XVALO(RANDSTRING$(-6))
          If CantRegistros("ADJUENTRA", "NUME_MAIL = " & NUMAPRO&, "NOMESS") > 0 Then GoTo 10
          Call CreaRegistro("ADJUENTRA", "NUME_MAIL", NUMAPRO&, "NOMESS")
          '
          CONDI$ = "NUME_MAIL = " & NUPROMAI&
          NORDE& = 1 + XVALO(VALOBADA("ADJUENTRA", "MAX(NORD_MAIL)", CONDI$, "NOMESS"))
          Call ACTUREGISTRO("ADJUENTRA", "NORD_MAIL", "NUME_MAIL = " & NUMAPRO&, NORDE&, RAFE&, "NOMESS")
          Call ACTUREGISTRO("ADJUENTRA", "FILE_NAME", "NUME_MAIL = " & NUMAPRO&, FIFUNA$, RAFE, "NOMESS")
          Call ACTUREGISTRO("ADJUENTRA", "FILE_DATA", "NUME_MAIL = " & NUMAPRO&, XXX$, RAFE&, "NOMESS")
          Call ACTUREGISTRO("ADJUENTRA", "NUME_MAIL", "NUME_MAIL = " & NUMAPRO&, NUPROMAI&, RAFE&, "NOMESS")
       End If
19  Next KKII%
    '
    ' PONE AVISO EMERGENTE AL DESTINATARIO
    
    ' BORRA EL MAIL YA GUARDADO EN LA BASE DE DATOS
30  TTYY$ = MAPIMessages1.MsgNoteText
    TTZZ$ = VALOBADA("BANDENTRA", "TEXT_MAIL", CONDI$, "NOMESS")
    If TTYY$ = TTZZ$ Then
       OKX% = 1
    End If
    '
End Sub

Function BORRAPORREGLAS%(ORENVI$, DIELREM$, ASUNTO$, TXTX$)
    '
    Static YACARLI%
    Static CANCRI%, STRI1$(), UBI1$(), STRI2$(), UBI2$()
    '
    Call CREACAMPO("", "CRITBORR", "NCRITBO", 3, 2, 1)
    Call CREACAMPO("", "CRITBORR", "STRING1", 10, 128)
    Call CREACAMPO("", "CRITBORR", "UBICA1", 10, 32)
    Call CREACAMPO("", "CRITBORR", "STRING2", 10, 128)
    Call CREACAMPO("", "CRITBORR", "UBICA2", 10, 32)
    '
    If CantRegistros&("CRITBORR", "NCRITBO > 0", "NOMESS") <> CANCRI% Then
       YACARLI% = 0
    End If
    '
    If YACARLI% < 1 Then
       Dim CRIBO As ADODB.Recordset
       Set CRIBO = New ADODB.Recordset
       Set CRIBO = FLEXCONN.Execute("SELECT * FROM CRITBORR WHERE NCRITBO > 0 ORDER BY NCRITBO")
       CANCRI% = 0
       With CRIBO
          Do While Not .EOF
            CANCRI% = CANCRI% + 1
            ReDim Preserve STRI1$(CANCRI%), UBI1$(CANCRI%), STRI2$(CANCRI%), UBI2$(CANCRI%)
            '
            STRI1$(CANCRI%) = TRIM$(SAFETEXT$(!STRING1))
            UBI1$(CANCRI%) = UCase$(TRIM$(SAFETEXT$(!UBICA1)))
            STRI2$(CANCRI%) = "": UBI2$(CANCRI%) = ""
            If TRIM$(SAFETEXT$(!STRING2)) <> "" Then
               If TRIM$(SAFETEXT$(!UBICA2)) <> "" Then
                  STRI2$(CANCRI%) = TRIM$(SAFETEXT$(!STRING2))
                  UBI2$(CANCRI%) = UCase$(TRIM$(SAFETEXT$(!UBICA2)))
               End If
            End If
          .MoveNext
          Loop
       End With
       CRIBO.Close
       Set CRIBO = Nothing
       YACARLI% = 1
    End If
    '
    OKX% = 0                              ' DEVUELVE 1 SI LO PROCESO COMO MAIL SEGURO
    '
    If TRIM$(DIELREM$) = "" Then    ' SI NO TIENE REMITENTE LO BORRA
        OKX% = 1: GoTo 99
    End If
    '
    For KKI% = 1 To CANCRI%
       TX1$ = "": TX2$ = ""
       If UBI1$(KKI%) = "ASUNTO" Then
            TX1$ = ASUNTO$
          ElseIf UBI1$(KKI%) = "REMITENTE" Then
            TX1$ = ORENVI$ + "<" + DIELREM$ + ">"
          ElseIf UBI1$(KKI%) = "TEXTO" Then
            TX1$ = TXTX$
       End If
       If UBI2$(KKI%) = "ASUNTO" Then
            TX2$ = ASUNTO$
          ElseIf UBI2$(KKI%) = "REMITENTE" Then
            TX2$ = ORENVI$ + "<" + DIELREM$ + ">"
          ElseIf UBI2$(KKI%) = "TEXTO" Then
            TX2$ = TXTX$
       End If
       '
       If InStr(UCase$(TX1$), UCase$(STRI1$(KKI%))) > 0 Then
          If STRI2$(KKI%) = "" Or InStr(UCase$(TX2$), UCase$(STRI2$(KKI%))) > 0 Then
             OKX% = 1: GoTo 99
          End If
       End If
       '
    Next KKI%
    '
99  BORRAPORREGLAS% = OKX%
    '
End Function
'
Sub GUARDADATPRIN()
    '
    ' LEE TODOS LOS DATOS
    IDENMA$ = PRINMAI09.Label4
    ORENVI$ = PRINMAI09.Label14
    DIELREM$ = PRINMAI09.Label16
    DIELDES$ = PRINMAI09.Label17
    ASUNTO$ = PRINMAI09.Label15
    FRECEP = FETEXCOR1$(FECHANUM(Left$(PRINMAI09.Label2, 8))) + "  " + Mid$(PRINMAI09.Label2, 11)
    TXTX$ = PRINMAI09.Text1
    TYTY$ = PRINMAI09.Text2
    SOLCONF% = PRINMAI09.Check1.Value
    '
    ' GUARDA MAIL EN BANDENTRA
    NUPROMAI& = 1 + XVALO(VALOBADA("BANPENPRO", "MAX(NUME_MAIL)", "NUME_MAIL > 0", "NOMESS"))
    CONDI$ = "NUME_MAIL = " & NUPROMAI&
    Call CreaRegistro("BANPENPRO", "NUME_MAIL", NUPROMAI&, "NOMESS")
    '
    Call ACTUREGISTRO("BANPENPRO", "IDEN_MAIL", CONDI$, IDENMA$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "ORIG_MAIL", CONDI$, Left$(ORENVI$, 256), RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "DREL_MAIL", CONDI$, Left$(DIELREM$, 256), RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "ASUN_MAIL", CONDI$, Left$(ASUNTO$, 256), RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "PARA_MAIL", CONDI$, DIELDES$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "TEXT_MAIL", CONDI$, TXTX$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "HTMT_MAIL", CONDI$, TYTY$, RAFE&)
    Call ACTUREGISTRO("BANPENPRO", "FHOR_RECEP", CONDI$, FRECEP, RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "PROS_DEST", CONDI$, "", RAFE&, "NOMESS")
    Call ACTUREGISTRO("BANPENPRO", "BORR_MAIL", CONDI$, 0, RAFE&, "NOMESS")
    '
    ' GUARDA LOS NOMBRES DE LOS ADJUNTOS
    CANADJ% = PRINMAI09.List3.ListCount
    For KKII% = 1 To CANADJ%
       FIFUNA$ = TRIM$(Left$(PRINMAI09.List3.List(KKII% - 1), 128))
10     NUMAPRO& = XVALO(RANDSTRING$(-6))
       If CantRegistros("ADJPENPRO", "NUME_MAIL = '" & NUMAPRO& & "'", "NOMESS") > 0 Then GoTo 10
       Call CreaRegistro("ADJPENPRO", "NUME_MAIL", NUMAPRO&, "NOMESS")
       '
       CONDI$ = "NUME_MAIL = " & NUPROMAI&
       NORDE& = 1 + XVALO(VALOBADA("ADJPENPRO", "MAX(NORD_MAIL)", CONDI$, "NOMESS"))
       Call ACTUREGISTRO("ADJPENPRO", "NORD_MAIL", "NUME_MAIL = " & NUMAPRO&, NORDE&, RAFE&, "NOMESS")
       Call ACTUREGISTRO("ADJPENPRO", "FILE_NAME", "NUME_MAIL = " & NUMAPRO&, FIFUNA$, RAFE, "NOMESS")
       Call ACTUREGISTRO("ADJPENPRO", "NUME_MAIL", "NUME_MAIL = " & NUMAPRO&, NUPROMAI&, RAFE&, "NOMESS")
19  Next KKII%
    '
End Sub
'
Sub PROCESAMAIL(IDENMA$, BORRAX%)
    '
    NUMAIPX& = XVALO(VALOBADA("BANPENPRO", "NUME_MAIL", "IDEN_MAIL = '" & IDENMA$ & "'", "NOMESS"))
    FIDESTI$ = TRIM$(VALOBADA("BANPENPRO", "PROS_DEST", "IDEN_MAIL = '" & IDENMA$ & "'", "NOMESS"))
    DEBORRA% = XVALO(VALOBADA("BANPENPRO", "BORR_MAIL", "IDEN_MAIL = '" & IDENMA$ & "'", "NOMESS"))
    '
    BORRAX% = 0
    If FIDESTI$ = "DELETE" Then
       MAPIMessages1.Delete
       BORRAX% = 1
       CONDI$ = "NUME_MAIL = " & NUMAIPX&
       Call BORRAREGISTROS("BANPENPRO", CONDI$, RAFE&, "NOMESS")
       Call BORRAREGISTROS("ADJPENPRO", CONDI$, RAFE&, "NOMESS")
    End If
    '
End Sub

