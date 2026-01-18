VERSION 5.00
Begin VB.Form FACOCOMP 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación de Proveedores - Opciones"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
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
      Height          =   855
      Left            =   4080
      Picture         =   "FACOCOMP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Listado de Facturación - Formato Bonficación"
      Top             =   2920
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4200
      ScaleHeight     =   75
      ScaleWidth      =   585
      TabIndex        =   20
      Top             =   3840
      Width           =   645
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      Picture         =   "FACOCOMP.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Control de Integridad de Acreedores"
      Top             =   2060
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Height          =   855
      Left            =   4095
      Picture         =   "FACOCOMP.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4095
      Picture         =   "FACOCOMP.frx":075E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1180
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tablas de Validación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Index           =   1
      Left            =   210
      TabIndex        =   7
      Top             =   3240
      Width           =   3690
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   8
         Top             =   525
         Width           =   2430
         Begin VB.CommandButton Command1 
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
            Left            =   0
            Picture         =   "FACOCOMP.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Acceso Directo a Tabla de Imputaciones Contables de Compras"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Plan de Cuen- tas Contables"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   945
            TabIndex        =   10
            ToolTipText     =   "Acceso a la Actualización del Plan de Cuentas Contables"
            Top             =   105
            Width           =   1425
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Operaciones Especiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   3690
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   16
         Top             =   1995
         Width           =   2430
         Begin VB.CommandButton Command4 
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
            Left            =   0
            Picture         =   "FACOCOMP.frx":0BB2
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Ajuste Directo de Saldos Acreedores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Ajuste de Sal- dos Acreedores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   18
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   80
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   13
         Top             =   1470
         Width           =   2430
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
            Left            =   0
            Picture         =   "FACOCOMP.frx":0CFC
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Imputación de Créditos y Anticipos de Proveedores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Imputación de Creditos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   15
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   80
            Width           =   1110
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   1
         Top             =   945
         Width           =   2430
         Begin VB.CommandButton Command23 
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
            Left            =   0
            Picture         =   "FACOCOMP.frx":0E46
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Edicion de Documentos de Compras"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta por Núme- ro de Documento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   630
            TabIndex        =   3
            Top             =   75
            Width           =   1785
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   4
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command25 
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
            Left            =   0
            Picture         =   "FACOCOMP.frx":1150
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Anulación de Facturas, Débitos o Créditos de Proveedores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de Documentos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   6
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   75
            Width           =   1530
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3555
      Picture         =   "FACOCOMP.frx":145A
      Top             =   4080
      Width           =   2010
   End
End
Attribute VB_Name = "FACOCOMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Command1.Enabled = False
    '
    If (USNBR% Mod 100) <> 1 Then
       Call COMUNI("La Opci¢n Elegida Requiere\Privilegios de Supervisor.")
       Command1.Enabled = True
       Exit Sub
    End If
    '
    Call CARCUEINGAS
    Command1.Enabled = True
End Sub

Private Sub Command11_Click()
    '\\\OT-05-0412-WAR-21/06/05///
Dim NUFA$()
    Command11.Enabled = False
    FIN& = ULTREG&("CUECORR")
    ReDim NUFA$(FIN&)
    '
    For I& = 1 To FIN&
        Y$ = REGLEIDO("CUECORR", I&)
        NUFARCH1$ = TRIM$(Mid$(Y$, 7, 18))
        NUFA$(I&) = NUFARCH1$
        Call TRACE(20, I&, FIN&)
    Next I&
    '
    Screen.MousePointer = 11
    For I& = 1 To FIN&
      Y$ = REGLEIDO("CUECORR", I&)
      NUFARCH1$ = TRIM$(Mid$(Y$, 7, 18))
      Call TRACE(20, I&, FIN&)
      If Left$(NUFARCH1$, 2) <> "OP" And Left$(NUFARCH1$, 2) <> "AS" Then
        For B& = 1 To I& - 1
          If NUFA$(B&) = NUFARCH1$ Then
            'DATOS DEL DOCUMENTO DEL VECTOR
            X$ = REGLEIDO("CUECORR", B&)
            NC% = CVI(Mid$(X$, 5, 2))
            RASOP1$ = RASOCLI(NC% * -1)
            FE1 = CVI(Mid$(X$, 1, 2))
            IMP1 = CVD(Mid$(X$, 123, 8))
            If IMP1 >= 0.005 Then
              'DATOS DEL DOCUMENTO DEL ARCHIVO
              NCARCH% = CVI(Mid$(Y$, 5, 2))
              RASOPARCH$ = RASOCLI$(NCARCH% * -1)
              FE2 = CVI(Mid$(Y$, 1, 2))
              IMP2 = CVD(Mid$(Y$, 123, 8))
              If (Abs(IMP1 - IMP2)) < 0.05 Then
                MSJ1$ = "El Documento: " & NUFA$(I&)
                MSJ1$ = MSJ1$ & " - " & RASOP1$ & " - " & FECHATEX(FE1)
                MSJ1$ = MSJ1$ & " - $" & FORMATNUM(IMP1, "F14.2")
                MSJ1$ = MSJ1$ & "\ se ha Encontrado Duplicado en: "
                MSJ1$ = MSJ1$ & " - " & RASOPARCH$ & " - " & FECHATEX(FE2)
                MSJ1$ = MSJ1$ & " - $" & FORMATNUM(IMP2, "F14.2")
                Call COMUNI(MSJ1$)
              'MsgBox "ACA"
              End If
            End If
          End If
        Next B&
      End If
    Next I&
    '
    Screen.MousePointer = 1
    Call COMUNI("Fin de la Revisión")
    Command11.Enabled = True
    '\\\OT-05-0412-WAR-FIN///
End Sub

Private Sub Command2_Click()
    '
    Hide
    '
End Sub

Private Sub Command23_Click()
    '
    Command23.Enabled = False
    Call OPNOIN("Consulta por Número de Documento de Compra")
    'Call REPIORPA
    Command23.Enabled = True
    '
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    Call ANUFACO
    Command25.Enabled = True
End Sub

Private Sub Command3_Click()
    Hide
    Call IMPUPRO1
    On Error Resume Next
    FIMPUPRO.Hide
    On Error GoTo 0
    FIMPUPRO.Show 1
    Call IMPUPRO2
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call AJUSALP        ' Ajuste de Saldos de Proveedores
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
  '
    Command5.Enabled = False
    Screen.MousePointer = 11
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If DES% < 1 Then GoTo 99
    XY$ = REGBLAN("FACBOCOM")
    JJ& = 0
    For J& = 1 To ULTREG("FACBONCO")
      ZY$ = REGLEIDO("FACBONCO", J&)
      F1% = CVI(Mid$(ZY$, 3, 2))
      If F1% >= DES% Then
        If F1% <= HAS% Then
          JJ& = JJ& + 1
          XY$ = ZY$
          Call GRAREG("FACBOCOM", XY$, JJ&)
        End If
      End If
    Next J&
    Call SETULTREG("FACBOCOM", JJ&)
    '
    Call FILESORT("FACBOCOM", "FACBOCOM", 5, 5, "ASC")
    Call FILESORT("FACBOCOM", "FACBOCOM", 2, 2, "ASC") '4
    Call FINAL("*LIFABOCO")
    '
99  Screen.MousePointer = 1
    Command5.Enabled = True
  '
End Sub

Private Sub Command7_Click()
    Call HELPONLINE("FACOMP02")
End Sub

