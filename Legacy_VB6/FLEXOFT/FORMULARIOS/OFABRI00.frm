VERSION 5.00
Begin VB.Form OFABRI00 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0F8&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "O/F's - Ordenes de Fabricación"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5745
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "OFABRI00.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "News"
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
      Left            =   4800
      Picture         =   "OFABRI00.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Informes de Producción"
      Top             =   2875
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Reserv"
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
      Left            =   4800
      Picture         =   "OFABRI00.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Control y Gestión de Reservas"
      Top             =   3715
      Width           =   750
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4800
      ScaleHeight     =   75
      ScaleWidth      =   705
      TabIndex        =   35
      Top             =   4620
      Width           =   765
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command14 
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
      Height          =   750
      Left            =   4800
      Picture         =   "OFABRI00.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Configuración de Procesos y Formularios"
      Top             =   2035
      Width           =   750
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Programación de Carga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   210
      TabIndex        =   11
      Top             =   4095
      Width           =   4425
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2310
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   29
         Top             =   525
         Width           =   1800
         Begin VB.CommandButton Command6 
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
            Picture         =   "OFABRI00.frx":0E98
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Consulta de Tiempo Total Programado por Equipo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Carga Total por Equipo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   630
            TabIndex        =   31
            Top             =   105
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   26
         Top             =   525
         Width           =   1800
         Begin VB.CommandButton Command7 
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
            Picture         =   "OFABRI00.frx":11A2
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Consulta y Actualización de Operaciones Proyectadas por O.F."
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Operaciones Programadas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   630
            TabIndex        =   28
            Top             =   105
            Width           =   1320
         End
      End
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
      Height          =   750
      Left            =   4800
      Picture         =   "OFABRI00.frx":12EC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1195
      Width           =   750
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
      Left            =   4800
      Picture         =   "OFABRI00.frx":15F6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   355
      Width           =   750
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Base de Datos de O-F's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   210
      TabIndex        =   1
      Top             =   210
      Width           =   4425
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1200
         ScaleHeight     =   585
         ScaleWidth      =   2100
         TabIndex        =   13
         Top             =   520
         Width           =   2160
         Begin VB.CommandButton Command5 
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
            Left            =   1575
            Picture         =   "OFABRI00.frx":1740
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Ver y Re-Imprimir Copia de Orden de Fabricación Emitida"
            Top             =   0
            Width           =   530
         End
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
            Left            =   3675
            Picture         =   "OFABRI00.frx":1A4A
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Orden de Disposicion de Tintorería"
            Top             =   0
            Width           =   530
         End
         Begin VB.CommandButton Command41 
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
            Picture         =   "OFABRI00.frx":1B94
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Generar Orden de Fabricación"
            Top             =   0
            Width           =   530
         End
         Begin VB.CommandButton Command46 
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
            Left            =   525
            Picture         =   "OFABRI00.frx":1E9E
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Cierre de Orden de Fabricación"
            Top             =   0
            Width           =   530
         End
         Begin VB.CommandButton Command45 
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
            Left            =   1050
            Picture         =   "OFABRI00.frx":21A8
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Anulación de Orden de Fabricación"
            Top             =   0
            Width           =   530
         End
      End
      Begin VB.Frame Frame11 
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
         TabIndex        =   2
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command23 
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
            TabIndex        =   8
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
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
            TabIndex        =   7
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command21 
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
            TabIndex        =   6
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command20 
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
            TabIndex        =   5
            Top             =   1155
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
            TabIndex        =   4
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command16 
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
            TabIndex        =   3
            Top             =   1890
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   210
      TabIndex        =   0
      Top             =   1740
      Width           =   4425
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0F8&
         Caption         =   "Dispositivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   315
         TabIndex        =   23
         Top             =   1470
         Width           =   3795
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0F8&
            Caption         =   "Impresión"
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
            Left            =   2205
            TabIndex        =   25
            Top             =   270
            Width           =   1170
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0F8&
            Caption         =   "Pantalla"
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
            Left            =   525
            TabIndex        =   24
            Top             =   270
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   3735
         TabIndex        =   17
         Top             =   630
         Width           =   3795
         Begin VB.CommandButton Command10 
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
            Left            =   1365
            Picture         =   "OFABRI00.frx":24B2
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Pendientes por Sector Productivo"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command63 
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
            Left            =   1950
            Picture         =   "OFABRI00.frx":27BC
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Pendientes por Número de Orden de Fabricación"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command61 
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
            Picture         =   "OFABRI00.frx":2AC6
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Consulta General de la Base de Datos"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command62 
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
            Left            =   2550
            Picture         =   "OFABRI00.frx":2DD0
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Consulta / Listado Selectivo por Estado"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command64 
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
            Left            =   3150
            Picture         =   "OFABRI00.frx":2F1A
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Consulta / Listado Pendientes por Código de Producto"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "B-D de O/F's "
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
            TabIndex        =   21
            Top             =   105
            Width           =   750
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4240
      Picture         =   "OFABRI00.frx":3224
      Top             =   4935
      Width           =   2010
   End
End
Attribute VB_Name = "OFABRI00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    If Option1.Value = True Then
           Call CONECRUN("*ORF-C05", "Consulta O/F's por Sector")
        Else
           ZZ$ = REGSEL$("SECPROD")
           If Len(ZZ$) <= 4 Then GoTo 99
           SESEL$ = Left$(ZZ$, 4)
           J& = 0
           Screen.MousePointer = 11
           For U& = 1 To ULTREG&("ORDEFABR")
               X$ = REGLEIDO$("ORDEFABR", U&)
               STAT% = Asc(Mid$(X$, 108, 1))
               If STAT% > 0 Then
                  If STAT% < 3 Then
                     CATOT# = CVD(Mid$(X$, 89, 8))
                     CAPRA# = CVD(Mid$(X$, 119, 8))
                     CAPEN# = CATOT# - CAPRA#
                     If CAPEN# < 0 Then CAPEN# = 0
                     If CAPEN# > 0 Then
10                      SPX$ = Mid$(X$, 113, 4)
                        If ECOARCH$("SECPROD", SPX$) = "" Then
                          CIXI% = CVI(Mid$(X$, 83, 2))
                          OBX$ = OBJPLA$("ENTRASECPROD", MKI$(CIXI%) + String$(4, 0))
                          If OBX$ = "" Then GoTo 99
                          SPX$ = Mid$(OBX$, 3, 4)
                          Call REPLA(X$, SPX$, 113, 4)
                          Call GRAREG("ORDEFABR", X$, U&)
                          GoTo 10
                        End If
15                      If SPX$ = SESEL$ Then
                           Call REPLA(X$, MKD$(CAPEN#), 159, 8)
                           Call REPLA(X$, MKD$(0), 167, 8)
                           J& = J& + 1
                           Call GRAREG("ORFASEL", X$, J&)
                        End If
                     End If
                  End If
               End If
           Next U&
           Call SETULTREG("ORFASEL", J&)
           Call CIERRARCH("ORFASEL")
           Call CIERRARCH("*.*")
           Call HEADERS("ORFASEL", "")
           Call HEADERS("ORFASEL", "Sector: " + ECOARCH$("SECPROD", SESEL$))
           Screen.MousePointer = 1
           '
           Call FINAL("*SORDEFA")
    End If
99  Command10.Enabled = True
    Screen.MousePointer = 1
    '
End Sub

Private Sub Command14_Click()
    'OPFABR00.Show 1
    Call CONECRUN("FLSETUP/OPFABR00", "Configuración de Funcionamiento")
    Call FINAL("")
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call CONECRUN("IFABRI02", "Informes de Producción")
   Command2.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    DISPOTINT.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("ORDEFABR")
End Sub

Private Sub Command41_Click()
    Command41.Enabled = False
    Call ORFNUE(1)
    Command41.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   OPX% = COMALTER%("Opciones Disponibles:\\Ver y Re-Imprimir\Re-Generar")
   If OPX% = 1 Then
        Call VERORFA
      ElseIf OPX% = 2 Then
        Call REPRINORFA
   End If
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call TOTEQUIP
'    Call FINAL("*TOTIEPRO")
    Command6.Enabled = True
End Sub

Private Sub Command64_Click()
    Command64.Enabled = False
    If Option1.Value = True Then
           Call CONECRUN("*ORF-C02", "Consulta O/F's por Producto")
        Else
           J& = 0
           Screen.MousePointer = 11
           For U& = 1 To ULTREG&("ORDEFABR")
               X$ = REGLEIDO$("ORDEFABR", U&)
               STAT% = Asc(Mid$(X$, 108, 1))
               If STAT% > 0 Then
                  If STAT% < 3 Then
                     CATOT# = CVD(Mid$(X$, 89, 8))
                     CAPRA# = CVD(Mid$(X$, 119, 8))
                     CAPEN# = CATOT# - CAPRA#
                     If CAPEN# < 0 Then CAPEN# = 0
                     Call REPLA(X$, MKD$(CAPEN#), 159, 8)
                     Call REPLA(X$, MKD$(0), 167, 8)
                     J& = J& + 1
                     Call GRAREG("ORFASEL", X$, J&)
                  End If
               End If
           Next U&
           Call SETULTREG("ORFASEL", J&)
           Call CIERRARCH("*.*")
           Screen.MousePointer = 1
           '
           Call FINAL("*SORDEFA")
    End If
    Command64.Enabled = True
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call ACTECOR(1)
'    Call FINAL("*OPERFAB")
    Command7.Enabled = True
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    OPRESERVA.Show 1
    Command9.Enabled = True
End Sub

Private Sub Command45_Click()
    Call ORFANU
End Sub

Private Sub Command46_Click()
    Call ORFCIE
End Sub

Private Sub Command61_Click()
    Command61.Enabled = False
    Call REVIORFLAN
    If Option1.Value = True Then
           Call CONECRUN("*CONORFA", "Consulta Base de datos de O.F's")
        Else
           Call OPNOIN("")
    End If
    Command61.Enabled = True
End Sub

Private Sub Command62_Click()
    Command62.Enabled = False
    Call GRAREG("ORFASTA", Chr$(0) + "Proyectadas", 1)
    Call GRAREG("ORFASTA", Chr$(1) + "Lanzadas", 2)
    Call GRAREG("ORFASTA", Chr$(2) + "Suspendidas", 3)
    Call GRAREG("ORFASTA", Chr$(3) + "Cerradas", 4)
    Call GRAREG("ORFASTA", Chr$(9) + "Anuladas", 5)
    Call CIERRARCH("ORFASTA")
    If Option1.Value = True Then
           Call CONECRUN("*CONORFS", "Consulta O.F's por Estado")
        Else
           OPX$ = OBJPLA$("SELORFSTA", "")
           If OPX$ <> "" Then
              If ECOARCH$("ORFASTA", OPX$) <> "" Then
                 J& = 0
                 Screen.MousePointer = 11
                 Call HEADERS("ORFASEL", "")
                 Call HEADERS("ORFASEL", "Estado: " + ECOARCH$("ORFASTA", OPX$))
                 For U& = 1 To ULTREG&("ORDEFABR")
                    X$ = REGLEIDO$("ORDEFABR", U&)
                    If Mid$(X$, 108, 1) = OPX$ Then
                       J& = J& + 1
                       Call GRAREG("ORFASEL", X$, J&)
                    End If
                Next U&
                Call SETULTREG("ORFASEL", J&)
                Call CIERRARCH("*.*")
                Screen.MousePointer = 1
                '
                Call FINAL("*LIORFST")
                '
             End If
          End If
    End If
    Command62.Enabled = True
End Sub

Private Sub Command63_Click()
    Command63.Enabled = False
    If Option1.Value = True Then
           Call CONECRUN("*ORF-C01", "Consulta por Número de O.F.")
        Else
           Call OPNOIN("")
    End If
    Command63.Enabled = True
End Sub

Private Sub Form_Activate()
   WindowState = 0
   OPENFORMS = DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    If EXISTE%("IFABRI02.EXE") > 0 Then
      Command2.Visible = True
    End If
    '
    Call GRATITFOR(Caption)
    '
    Refresh
End Sub
'

Sub VERORFA()
    '
3   Call SELECHO("ECORDEF/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NORFA$ = TRIM$(VALACT1$("ECORDEF"))
    If NORFA$ = "" Then GoTo 99
    '
    TIDOCUM$ = "Orden de Fabricación " + NORFA$
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 99
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Cancel = 0
   Hide
End Sub
