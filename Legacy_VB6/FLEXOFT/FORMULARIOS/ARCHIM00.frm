VERSION 5.00
Begin VB.Form ARCHIM00 
   BackColor       =   &H00D0F0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Archivos Maestros - Actualización y Consulta"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   6195
      Picture         =   "ARCHIM00.frx":0000
      TabIndex        =   40
      ToolTipText     =   "Importación de Archivos Maestros desde Excel"
      Top             =   210
      Width           =   770
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0FF&
      Caption         =   "Actualización Bases de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   4425
      Begin VB.Frame Frame4 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Materiales, Fórmulas y Rutas"
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
         Left            =   210
         TabIndex        =   12
         Top             =   420
         Width           =   4005
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   22
            Top             =   525
            Width           =   2325
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
               Picture         =   "ARCHIM00.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Maestro de Artí- culos de Stock"
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
               TabIndex        =   24
               Top             =   105
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   5
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command4 
            Caption         =   "Inventarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   11
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Negativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   10
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Consignación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   9
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Críticos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   8
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Estadísticas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   7
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Análisis A-B-C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   6
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Clientes / Proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   210
         TabIndex        =   4
         Top             =   1995
         Width           =   4005
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   28
            Top             =   840
            Width           =   2325
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
               Left            =   -20
               Picture         =   "ARCHIM00.frx":0614
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Clientes en Cuenta Corriente"
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
               Left            =   735
               TabIndex        =   30
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   25
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command19 
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
               Left            =   -20
               Picture         =   "ARCHIM00.frx":0A56
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Proveedores y Subcontratistas"
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
               Left            =   735
               TabIndex        =   27
               Top             =   0
               Width           =   1590
            End
         End
      End
   End
   Begin VB.CommandButton Command24 
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
      Height          =   770
      Left            =   8715
      Picture         =   "ARCHIM00.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   210
      Width           =   770
   End
   Begin VB.CommandButton Command29 
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
      Height          =   770
      Left            =   7875
      Picture         =   "ARCHIM00.frx":0EAA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   210
      Width           =   770
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   7035
      Picture         =   "ARCHIM00.frx":11B4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   770
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0F0FF&
      Caption         =   "Tablas de Validación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   4830
      TabIndex        =   13
      Top             =   1155
      Width           =   4635
      Begin VB.CommandButton Command7 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   3570
         Picture         =   "ARCHIM00.frx":15F6
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   525
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1050
         Width           =   3225
      End
      Begin VB.CommandButton Command6 
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
         Height          =   360
         Left            =   3255
         TabIndex        =   32
         Top             =   525
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   31
         Text            =   " "
         Top             =   525
         Width           =   1695
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   15
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command20 
            Caption         =   "Análisis A-B-C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   21
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Estadísticas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   20
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Críticos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   19
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Consignación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   18
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Negativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   17
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Inventarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   16
            Top             =   420
            Width           =   1590
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Productos, Fórmulas y Rutas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   14
         Top             =   1575
         Width           =   4215
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Secciones Productivas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   2310
            TabIndex        =   39
            Top             =   735
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Depósitos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   2310
            TabIndex        =   38
            Top             =   315
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Unidades de Medida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   315
            TabIndex        =   37
            Top             =   735
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Tipos de Material"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   315
            TabIndex        =   36
            Top             =   315
            Width           =   1800
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -735
         TabIndex        =   35
         Top             =   630
         Width           =   2325
      End
   End
End
Attribute VB_Name = "ARCHIM00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$
'

Private Sub Command1_Click()
   Command1.Enabled = False
   If EXISTE%("AMASTO02.EXE") > 0 Then
     Call CONECRUN("AMASTO02", "Maestro de Items de Stock")
   Else
     AMASTO00.Show 1
   End If
   Command1.Enabled = True
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
    If App.EXEName = "FLGEST00" Then
        Call CONECRUN("AMACLI02", "Padron Maestro de Clientes")
      Else
        AMACLI00.Show 1
    End If
    Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Call CONECRUN("AMAPRO", "Actualización de Base de Datos de Proveedores")
    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   IMPORGEN.Show 1
99 Command2.Enabled = True
End Sub

Private Sub COMMAND21_CLICK()
   Command21.Enabled = False
   Call CONECRUN("INOCON01", "Informes de No-Conformidad")
   Command21.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   Call CONECRUN("DOCING01", "Documentos de Ingenieria")
   Command22.Enabled = True
End Sub

Private Sub Command23_Click()
    Call OPNOIN("")
End Sub

Private Sub Command24_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Command25_Click()
Call BLOQARCH("MASTER")
End Sub

Private Sub Command26_Click()
Call BLOQARCH("INVERT")
End Sub

Private Sub Command27_Click()
Call BLOQARCH("INVENUE")
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("ARCHIMA")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    '
    If EXISTE%("AMASTO01.EXE") > 0 Then
        Call CONECRUN("AMASTO01", "Maestro de Items de Stock")
        GoTo 99
      Else
        Call CONECRUN("HORUTAS", "Hojas de Ruta y Proceso")
    End If
    '
99  Command3.Enabled = True
End Sub


Private Sub Command7_Click()
    Command7.Enabled = False
'    If MODUCONE = "TADEPOSI" Then
'         Call CARTADEPOSI
'       ElseIf MODUCONE = "TABTIMA" Then
'         Call CARTATIMA
'       ElseIf MODUCONE = "UMEDIDA" Then
'         Call CARUMEDIDA
'       Else
'         If TRIM$(MODUCONE$) <> "" Then
'            Call CONECRUN(MODUCONE$, Text2.TEXT)
'         End If
'    End If
   If TRIM$(MODUCONE$) <> "" Then
        If MODUCONE = "TADEPOSI" Then
            Call CARTADEPOSI
          ElseIf MODUCONE = "TABTIMA" Then
            Call CARTATIMA
          ElseIf MODUCONE = "UMEDIDA" Then
            Call CARUMEDIDA
          ElseIf MODUCONE = "SECPROD" Then
            Call CARSECPROD
          Else
            Call CONECRUN(MODUCONE$, Text2.TEXT)
        End If
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call CONECRUN("INSTRU01", "B.D.Instrumentos de medición y Ensayo")
   Command8.Enabled = True
End Sub

Private Sub Form_Load()
   '
   'Call VERJOBID
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   '
   TABLA$(0) = "TABTIMA"       ' tabla de tipos de material
   TABLA$(1) = "UMEDIDA"
   TABLA$(2) = "TADEPOSI"
   TABLA$(3) = "SECPROD"
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
End Sub

Private Sub Option1_Click(Index As Integer)
    Option1(Index).Value = True
    Text2.TEXT = REPLACAR$(Option1(Index).Caption, "- ", "")
    MODUCONE$ = TABLA$(Index)
    TETABLA$ = MODUCONE$
    If Left$(TETABLA$, 1) = "#" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(#)"
      ElseIf Left$(TETABLA$, 1) = "*" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(*)"
    End If
    Text1.TEXT = TETABLA$
    '
    For K% = 0 To 3
       If K% <> Index Then
          Option1(K%).Value = False
       End If
    Next K%
    Command7.SetFocus
End Sub

Private Sub Option1_DblClick(Index As Integer)
    If TRIM$(MODUCONE$) <> "" Then
        If MODUCONE = "TADEPOSI" Then
            Call CARTADEPOSI
          ElseIf MODUCONE = "TABTIMA" Then
            Call CARTATIMA
          ElseIf MODUCONE = "UMEDIDA" Then
            Call CARUMEDIDA
          ElseIf MODUCONE = "SECPROD" Then
            Call CARSECPROD
          Else
            Call CONECRUN(MODUCONE$, Text2.TEXT)
        End If
    End If
End Sub

Sub CARUMEDIDA()
   '
   Call TRALOCAL("UNIMED", "")
10 VTB% = VENTABU%("TUMEDIDA/NC/TITUPAN=Unidades de Medida")
   If VTB% >= 0 Then
      JJ& = 0
      For U& = 1 To ULTREG&("!UNIMED")
         X$ = REGLEIDO$("!UNIMED", U&)
         If TRIM$(Left$(X$, 4)) <> "" Then
           If Len(TRIM$(Left$(X$, 4))) > 2 Then
             Call MENSERR(24, "Codigo de Unidad no Válido.\Máximo 2 Posiciones")
             GoTo 10
           End If
           Y$ = Left$(X$, 2) + Mid$(X$, 5, 28)
           JJ& = JJ& + 1
           Call GRAREG("UNIMED", X$, JJ&)
           Call GRAREG("UMEDIDA", Y$, JJ&)
         End If
      Next U&
      Call SETULTREG("UNIMED", JJ&)
      Call SETULTREG("UMEDIDA", JJ&)
   End If
   Call CIERRARCH("UNIMED")
   Call CIERRARCH("UMEDIDA")
   '
End Sub


