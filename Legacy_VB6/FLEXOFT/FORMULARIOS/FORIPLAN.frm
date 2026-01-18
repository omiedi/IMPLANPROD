VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FORIPLAN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importacion de DEUDA PROVEEDORES"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
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
      Height          =   750
      Left            =   11025
      Picture         =   "FORIPLAN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Inicia Importación"
      Top             =   945
      Width           =   750
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   "F:\CLIENTES\VDR\TRAIDOS\Plan de cuentas.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   9030
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Hoja1$"
      Top             =   105
      Width           =   1905
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
      Height          =   750
      Left            =   11025
      Picture         =   "FORIPLAN.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   105
      Width           =   750
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FORIPLAN.frx":0454
      Height          =   7470
      Left            =   0
      OleObjectBlob   =   "FORIPLAN.frx":0464
      TabIndex        =   1
      Top             =   315
      Width           =   10830
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aro Plastico"
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   7
      Top             =   1575
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resorte"
      Height          =   255
      Index           =   3
      Left            =   7560
      TabIndex        =   6
      Top             =   1260
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Carcaza"
      Height          =   255
      Index           =   2
      Left            =   7560
      TabIndex        =   5
      Top             =   945
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Acelerante"
      Height          =   255
      Index           =   1
      Left            =   7560
      TabIndex        =   4
      Top             =   630
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mezcla"
      Height          =   255
      Index           =   0
      Left            =   7560
      TabIndex        =   3
      Top             =   315
      Width           =   1380
   End
End
Attribute VB_Name = "FORIPLAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Close
   Hide
End Sub
Private Sub Command4_Click()
Screen.MousePointer = 11
   URECORD& = Data1.Recordset.Recordcount
   DBGrid1.Height = (URECORD& + 50) * DBGrid1.RowHeight
   '
   LU$ = LUDAT$
   '
   JJ& = 0
   For U& = 1 To URECORD&
     On Error Resume Next
     DBGrid1.Row = U&
     If Err Then GoTo 19
     On Error GoTo 0
     DBGrid1.Col = 0
     TTXX$ = TRIM$(DBGrid1.TEXT)
     PPXX% = InStr(TTXX$, "-")
     If PPXX% < 2 Then GoTo 19
     COCUE$ = TRIM$(Left$(TTXX$, PPXX% - 1))
     DECUE$ = TRIM$(Mid$(TTXX$, PPXX% + 1))
     '
     XX$ = REGBLAN$("CUENTAS")
     Call REPLA(XX$, COCUE$, 1, 12)
     Call REPLA(XX$, "D", 13, 1)
     YY$ = Space$(128)
     Call REPLA(YY$, DECUE$, 1, 100)
     Call REPLA(YY$, DECUE$, 101, 28)
     JJ& = JJ& + 1: JI% = JJ&
     ZZ$ = REGBLAN$("INVECUE")
     Call REPLA(ZZ$, COCUE$, 1, 12)
     Call REPLA(ZZ$, MKI$(JI%), 13, 2)
     QQ$ = REGBLAN$("PLAGRAL")
     Call REPLA(QQ$, MKI$(JI%), 1, 2)
     Call REPLA(QQ$, MKI$(3), 3, 2)
     Call GRAREG("CUENTAS", XX$, JJ&)
     Call GRAREG("DENOCUE", YY$, JJ&)
     Call GRAREG("INVECUE", ZZ$, JJ&)
     Call GRAREG("PLAGRAL", QQ$, JJ&)
19 Next U&
   '
   Call SETULTREG("CUENTAS", JJ&)
   Call SETULTREG("DENOCUE", JJ&)
   Call SETULTREG("INVECUE", JJ&)
   Call SETULTREG("PLAGRAL", JJ&)
   Call CIERRARCH("CUENTAS")
   Call CIERRARCH("DENOCUE")
   Call CIERRARCH("PLAGRAL")
   Call FILESORT("INVECUE", "", 1, 1)
   Call CIERRARCH("*.*")
   '
Screen.MousePointer = 1

   Call COMUNI("La Importación del Plan de Cuentas\se ha Completado Exitosamente.")
   '
Exit Sub

End Sub

