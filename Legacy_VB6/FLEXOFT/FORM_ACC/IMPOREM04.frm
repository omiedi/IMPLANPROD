VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form IMPOREM04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesos de Importacion de Datos"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   1485
      Left            =   11025
      ScaleHeight     =   1425
      ScaleWidth      =   1005
      TabIndex        =   27
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox Picture16 
         Height          =   1500
         Left            =   0
         ScaleHeight     =   1440
         ScaleWidth      =   1020
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   1400
            Left            =   0
            Picture         =   "IMPOREM04.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
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
         Left            =   120
         Picture         =   "IMPOREM04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   210
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "IMPOREM04.frx":24E2
         Top             =   1050
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   210
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Resultado de Pre-Importación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Importación Definitiva"
      Top             =   1560
      Width           =   11655
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   10800
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   29
         Top             =   2260
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   30
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         Height          =   1965
         Left            =   150
         ScaleHeight     =   1965
         ScaleWidth      =   10215
         TabIndex        =   16
         Top             =   420
         Width           =   10215
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   1965
            Left            =   -20
            ScaleHeight     =   1965
            ScaleWidth      =   10215
            TabIndex        =   18
            Top             =   0
            Width           =   10215
            Begin VB.PictureBox Picture3 
               Height          =   1950
               Left            =   -20
               ScaleHeight     =   1890
               ScaleWidth      =   0
               TabIndex        =   23
               Top             =   0
               Width           =   60
            End
            Begin VB.PictureBox Picture7 
               Height          =   270
               Left            =   0
               ScaleHeight     =   210
               ScaleWidth      =   1.19940e5
               TabIndex        =   20
               Top             =   0
               Width           =   1.20000e5
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  BackColor       =   &H008080FF&
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   40
                  TabIndex        =   22
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label Label82 
                  AutoSize        =   -1  'True
                  BackColor       =   &H008080FF&
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   105
                  TabIndex        =   21
                  Top             =   0
                  Width           =   120
               End
            End
            Begin VB.ListBox List4 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1700
               IntegralHeight  =   0   'False
               ItemData        =   "IMPOREM04.frx":4404
               Left            =   0
               List            =   "IMPOREM04.frx":4406
               TabIndex        =   19
               Top             =   250
               Width           =   1.30000e5
            End
         End
         Begin VB.PictureBox Picture10 
            Height          =   1950
            Left            =   -20
            ScaleHeight     =   1890
            ScaleWidth      =   0
            TabIndex        =   17
            Top             =   0
            Width           =   60
         End
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         IntegralHeight  =   0   'False
         ItemData        =   "IMPOREM04.frx":4408
         Left            =   120
         List            =   "IMPOREM04.frx":440A
         TabIndex        =   11
         Top             =   2910
         Width           =   10400
      End
      Begin VB.CommandButton Command3 
         Caption         =   "List"
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
         Left            =   10800
         Picture         =   "IMPOREM04.frx":440C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Listado por Impresora"
         Top             =   3400
         Width           =   600
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
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
         Left            =   10800
         Picture         =   "IMPOREM04.frx":4716
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Inicia Importación Definitiva"
         Top             =   1365
         Width           =   600
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1725
         Left            =   10360
         TabIndex        =   5
         Top             =   670
         Width           =   225
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         Left            =   120
         Max             =   10000
         TabIndex        =   4
         Top             =   2340
         Width           =   10250
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10620
         TabIndex        =   25
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10800
         TabIndex        =   24
         Top             =   2910
         Width           =   600
      End
      Begin VB.Line Line6 
         X1              =   135
         X2              =   135
         Y1              =   420
         Y2              =   2520
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Registros no Procesables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2670
         Width           =   2775
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10800
         TabIndex        =   9
         Top             =   690
         Width           =   600
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10620
         TabIndex        =   8
         Top             =   400
         Width           =   615
      End
      Begin VB.Label TEPRUE 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3720
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Origen de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   10785
      Begin VB.CommandButton Command4 
         Caption         =   "O.K."
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
         Left            =   9975
         Picture         =   "IMPOREM04.frx":4A20
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Inicia Pre-Importación"
         Top             =   420
         Width           =   600
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   2520
         ScaleHeight     =   255
         ScaleWidth      =   5715
         TabIndex        =   13
         Top             =   600
         Width           =   5775
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            TabIndex        =   14
            ToolTipText     =   "Dbl-Click para Seleccionar Archivo"
            Top             =   0
            Width           =   5760
         End
      End
      Begin VB.CommandButton Command22 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8280
         TabIndex        =   1
         ToolTipText     =   "Click para Seleccionar Archivo"
         Top             =   600
         Width           =   220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta / Archivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   945
         TabIndex        =   2
         Top             =   600
         Width           =   1470
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   600
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Left            =   4620
      TabIndex        =   15
      Top             =   1050
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "IMPOREM04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Unload Me
End Sub
'
Private Sub Command2_Click()
   Command2.Enabled = False
   Screen.MousePointer = 11
   '
   'IMPORTACIÓN DEFINITIVA
   '
   If List4.ListCount > 0 Then
   NOI1% = 0
   U& = 0
   URECORD& = List4.ListCount
     For JK% = 0 To List4.ListCount - 1
       A$ = List4.List(JK%)
       FEC1$ = Mid$(A$, 1, 8)
        FE1% = FECHANUM(FEC1$)
       CE1$ = Mid$(A$, 10, 12)
       NFAC1$ = Mid$(A$, 23, 8)
       NC1% = Val(Mid$(A$, 32, 5))
       CA1$ = Mid$(A$, 59, 11)
        CI1% = CODINT%(CA1$)
       'CANTIDAD NEGATIVA XQ ES MOVIMIENTO DE SALIDA.
       CAN1 = Val(Mid$(A$, 71, 8)) * (-1)
       PUNI1# = Val(Mid$(A$, 80, 10))
       NPED1& = Val(Mid$(A$, 91, 6))
       'CODGIO DE REMITO
       CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
       If CMO$ = "" Then CMO$ = "RT"
       'LOTE
       L1$ = ""
       'REFERENCIA
       REFE1$ = RASOCLI$(NC1%)
       'NUMERO DE ITEM
       NOI1% = NOI1% + 1
       'DEPOSITO
       DEPX% = DEPOPRE%(CI1%)
       'ACA GRABO EN MOVISTO
       Call MOVISTOK(FE1%, CI1%, L1$, CMO$, NFAC1$, NFAC1$, NFAC1$, NFAC1$, REFE1$, NOI1%, PUNI1#, 1, NC1%, DEPX%, CAN1)
       U& = U& + 1
       Call TRACE(20, U&, URECORD&)
       '
     Next JK%
     Call COMUNI("Se ha Completado la Importacion Definitiva")
     Command2.Enabled = False
   End If
   '
99 Screen.MousePointer = 1
   DoEvents
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   '
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   Label1 = UCase$(CommonDialog1.FileName)
   '
   On Error GoTo 0
   '
99 Screen.MousePointer = 1
   Command22.Enabled = True
   On Error GoTo 0
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If List1.ListCount > 0 Then
     Screen.MousePointer = 11
     Printer.Print "Registros no Procesables"
     Printer.Print " "
     For U& = 1 To List1.ListCount
       Printer.Print List1.List(U& - 1)
     Next U&
     Printer.EndDoc
     Screen.MousePointer = 1
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   Frame5.Visible = True
   List1.Clear
   List4.Clear
   '
   On Error Resume Next
   HScroll1.Value = 0
   VScroll2.Value = 0
   On Error GoTo 0
   '
   TEPRUE = ""
   '
   If EXISTE%(Label1) < 1 Then
     Call MENSERR(24, "Establezca Archivo Origen de Importación")
     GoTo 99
   End If
   '
   RUTA1$ = Label1
   NX% = FILEOPEN%(RUTA1$, 1, 128)
   If NX% <= 0 Then
        Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     GoTo 99
   End If
   '
   U& = 0
   '
   Do While Not EOF(NX%)
      Line Input #NX%, A$
      '
      L1$ = Space(128)
      '
      FEC1$ = Mid$(A$, 1, 6)
      CE1$ = Mid$(A$, 7, 12)
      NFAC1$ = Mid$(A$, 19, 8)
      CUI1$ = Mid$(A$, 27, 13)
      CA1$ = Mid$(A$, 40, 11)
      CAN1$ = Mid$(A$, 51, 7)
      PUNI1$ = Mid$(A$, 58, 8)
      '
      U& = U& + 1
      'VALIDACIONES
      C1% = CODINT%(CA1$)
      If C1% <= 0 Then
        MERROX$ = "Registro" + Str$(U&) + " - Codigo de Producto Inexistente o No Valido"
        List1.AddItem MERROX$
        GoTo 19
      End If
      '
      CAN2 = Val(CAN1$)
      If CAN2 < 0.005 Then
        MERROX$ = "Registro" + Str$(U&) + " - Cantidad No Valida"
        List1.AddItem MERROX$
        GoTo 19
      End If
      '
      If TRIM(CUI1$) = "" Then
        MERROX$ = "Registro" + Str$(U&) + " - C.U.I.T. No Valido"
        List1.AddItem MERROX$
        GoTo 19
      End If
      'VALIDO QUE EL CUIT PERTENEZCA A ALGUN CLIENTE EXISTENTE
      NC1% = 0
      NC1% = NUCICU%(CUI1$)
      If NC1% <= 0 Then
        MERROX$ = "Registro" + Str$(U&) + " - Cliente Inexistente o No Valido"
        List1.AddItem MERROX$
        GoTo 19
      End If
      'AGREGO AL LIST DE PREIMPORTACION
      'OBTENGO EL NUMERO DE PEDIDO
      NPED1& = 0
      Call ABREPEDCLI
      With PedEnca
        .FindFirst " NroOcom = '" & CE1$ & "' "
        If .NoMatch = False Then
          NPED1& = !NroPed
        End If
      End With
      'CONVIERTO FECHA A FORMATO FLEX
      F1% = FECHANUM(FEC1$)
      FEC2$ = FECHATEX(F1%)
      F3# = CDbl(CVDAT(F1%))
      'VALIDAR QUE NO EXISTA EL COMPROBANTE (FECHA+NRO.DOC)
      'en MOVISTO
      Call ABRESTOCKS
      With Movisto
        SQLX$ = " DoSecLar = '" & NFAC1$ & "' AND CDBL(FechMov) = " & F3 & " "
        .FindFirst SQLX$
        If .NoMatch = False Then
          MERROX$ = "Registro" + Str$(U&) + " - Movimiento ya Registrado"
          List1.AddItem MERROX$
          GoTo 19
        End If
      End With
      '
      'FORMATEO CANTIDAD A FORMATO FLEX
      CAN2 = Val(CAN1$)
      CAN3$ = FORMATNUM$(CAN2, "F8.1")
      'FORMATEO P.UNIT A FORMATO FLEX
      PU2 = Val(PUNI1$)
      PU3$ = FORMATNUM$(PU2, "F10.4")
      'FORMATEO EL NUMERO DE CLIENTE
      NC2$ = FORMATNUM$(NC1%, "F5.0")
      RASO2$ = RASOCLI$(NC1%)
      'FORMATEO NUMERO DE PEDIDO
      NPED2$ = ""
      If NPED1& > 0 Then
        NPED2$ = FORMATNUM$(NPED1&, "F6.0")
      End If
      '
      Call REPLA(L1$, FEC2$, 1, 8)
      Call REPLA(L1$, CE1$, 10, 12)
      Call REPLA(L1$, NFAC1$, 23, 8)
      Call REPLA(L1$, NC2$, 32, 5)
      Call REPLA(L1$, RASO2$, 38, 20)
      Call REPLA(L1$, CA1$, 59, 11)
      Call REPLA(L1$, CAN3$, 71, 8)
      Call REPLA(L1$, PU3$, 80, 10)
      Call REPLA(L1$, NPED2$, 91, 6)
      '
17    List4.AddItem L1$
      '
      If Len(RTrim$(L1$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim$(L1$ + "*")
      End If
      On Error Resume Next
        List4.TopIndex = List4.ListCount - 7
      On Error GoTo 0
      DoEvents
      '
19    Label6 = TRIM$(Str$(List4.ListCount))
      Label3 = TRIM$(Str$(List1.ListCount))
      Label6.Refresh: Label3.Refresh
      '
   Loop
   '
   Screen.MousePointer = 11
   '
   Call COMUNI("Se Ha Completado la Pre-Importación.")
   '
   Command4.Enabled = True
   MXX% = List4.ListCount - 5
   If MXX% < 0 Then MXX% = 0
   VScroll2.Value = 0
   VScroll2.Max = MXX%
   '
   HScroll1.Value = 0
   HScroll1.Max = 1
   If TEPRUE.Width > HScroll1.Width Then
     HScroll1.Max = (TEPRUE.Width - HScroll1.Width) / 10
   End If
   '
99 Screen.MousePointer = 1
   Command4.Enabled = True
End Sub
Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      
      ' &H8000000F
    End If

   Call CENTRAFORM(Me)
   '
   ENCA$ = Space(128)
   Call REPLA(ENCA$, "Fecha", 1, 8)
   Call REPLA(ENCA$, "O/C", 10, 12)
   Call REPLA(ENCA$, "Factura", 23, 8)
   Call REPLA(ENCA$, " Cta.", 32, 5)
   Call REPLA(ENCA$, "R.Social", 38, 20)
   Call REPLA(ENCA$, "Código", 59, 11)
   Call REPLA(ENCA$, "Cantidad", 71, 8)
   Call REPLA(ENCA$, " Pr. Unit.", 80, 10)
   Call REPLA(ENCA$, "Pedido", 91, 6)
   Label8 = ENCA$
   Label82 = Space(300)
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
End Sub
'
Private Sub HScroll1_Change()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  Picture7.Left = (-1) * 10 * HScroll1.Value
  Picture2.Left = -20
  Picture2.Width = HScroll1.Width
  Ab1 = HScroll1.Value
  Ab1 = 10 * Ab1 + 40
  Ab2 = Abs(Picture7.Left)
  If Ab1 > Ab2 + 80 Then
    Picture2.Left = (-1) * (Ab1 - Ab2)
    Picture2.Width = HScroll1.Width - Picture2.Left
  End If
  On Error GoTo 0
End Sub
Private Sub HScroll1_Scroll()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  Picture7.Left = (-1) * 10 * HScroll1.Value
  Picture2.Left = -20
  Picture2.Width = HScroll1.Width
  Ab1 = HScroll1.Value
  Ab1 = 10 * Ab1 + 40
  Ab2 = Abs(Picture7.Left)
  If Ab1 > Ab2 + 80 Then
    Picture2.Left = (-1) * (Ab1 - Ab2)
    Picture2.Width = HScroll1.Width - Picture2.Left
  End If
  On Error GoTo 0
End Sub

Private Sub Label1_DBLClick()
    Call Command22_Click
End Sub


Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub VScroll2_Change()
   On Error Resume Next
   List4.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub
Private Sub VScroll2_Scroll()
   On Error Resume Next
   List4.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub

Function NUCICU%(CUI1$) 'DEVUELVE EL NUMERO DE CLIENTE
  '
  Call APERBASDAT("CLIEN")
  With Deudor12
    .FindFirst " Cuit_Cli = '" & CUI1$ & "' "
    If .NoMatch = False Then
      NUCICU% = !Nume_Cli
      Exit Function
    End If
  End With
  NUCICU% = 0
  '
End Function
