VERSION 5.00
Begin VB.Form CONFICRP 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de Máquinas (CRP) - Configuración"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command23 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8925
      Picture         =   "CONFICRP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   2045
      Width           =   855
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Planilla de Carga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   4410
      TabIndex        =   44
      Top             =   2940
      Width           =   5370
      Begin VB.CommandButton Command22 
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
         Left            =   4200
         Picture         =   "CONFICRP.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   420
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
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
         Width           =   3750
      End
      Begin VB.CommandButton Command21 
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
         Left            =   3780
         TabIndex        =   46
         Top             =   420
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
         Left            =   2415
         TabIndex        =   45
         Text            =   " "
         Top             =   420
         Width           =   1380
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   49
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.CommandButton Command19 
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
      Height          =   800
      Left            =   8925
      Picture         =   "CONFICRP.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Terminar - Finalizar la Aplicación"
      Top             =   315
      Width           =   855
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
      Height          =   800
      Left            =   8925
      Picture         =   "CONFICRP.frx":075E
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1180
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Disponibilidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4320
      Left            =   210
      TabIndex        =   8
      Top             =   210
      Width           =   4005
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "General de Planta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1750
         Left            =   210
         TabIndex        =   17
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   32
            Top             =   900
            Width           =   2430
            Begin VB.CommandButton Command17 
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
               Picture         =   "CONFICRP.frx":0A68
               Style           =   1  'Graphical
               TabIndex        =   33
               ToolTipText     =   " "
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Calendario de Dias Feriados"
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
               TabIndex        =   34
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   35
            Top             =   420
            Width           =   2430
            Begin VB.CommandButton Command20 
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
               Picture         =   "CONFICRP.frx":0D72
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Turnos Diarios de Producción"
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
               TabIndex        =   37
               Top             =   0
               Width           =   1485
            End
         End
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
         TabIndex        =   10
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command16 
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
         Begin VB.CommandButton Command15 
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
            TabIndex        =   15
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command14 
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
            TabIndex        =   14
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command9 
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
            TabIndex        =   13
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command8 
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
            TabIndex        =   12
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command7 
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
            TabIndex        =   11
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Máquinas y Equipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1750
         Left            =   210
         TabIndex        =   9
         Top             =   2310
         Width           =   3585
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   38
            Top             =   900
            Width           =   2430
            Begin VB.CommandButton Command99 
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
               Picture         =   "CONFICRP.frx":107C
               Style           =   1  'Graphical
               TabIndex        =   39
               ToolTipText     =   " "
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Paradas de Equi- po Programadas"
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
               TabIndex        =   40
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   41
            Top             =   420
            Width           =   2430
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
               Picture         =   "CONFICRP.frx":1386
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Horarios Especia- les por Equipo"
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
               TabIndex        =   43
               Top             =   0
               Width           =   1695
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Recursos Generales"
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
      Left            =   4410
      TabIndex        =   0
      Top             =   210
      Width           =   4320
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1260
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   20
         Top             =   1785
         Width           =   2535
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
            Left            =   0
            Picture         =   "CONFICRP.frx":1690
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Definicion de Depósitos y Sucursales"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Depósitos y Almacenes"
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
            TabIndex        =   21
            Top             =   105
            Width           =   1605
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   25
         Top             =   1365
         Width           =   2535
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
            Picture         =   "CONFICRP.frx":199A
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Definición de  Máquinas / Equipo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Máquinas y Equi- pos de Producción"
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
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   22
         Top             =   945
         Width           =   2535
         Begin VB.CommandButton Command2 
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
            Picture         =   "CONFICRP.frx":1CA4
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Definicion de Secciones Productivas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Secciones y Cen- tros de Producción"
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
            TabIndex        =   24
            Top             =   0
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   28
         Top             =   525
         Width           =   2535
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
            Picture         =   "CONFICRP.frx":1FAE
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Definición de Plantas y Unidades Productivas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Plantas y Uni- dades Productivas"
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
            Top             =   0
            Width           =   1695
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
         TabIndex        =   1
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
            TabIndex        =   7
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
            TabIndex        =   6
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
            TabIndex        =   5
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
            TabIndex        =   4
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
            TabIndex        =   3
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
            TabIndex        =   2
            Top             =   1890
            Width           =   1590
         End
      End
   End
End
Attribute VB_Name = "CONFICRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call OPNOIN("")
End Sub

Private Sub Command17_Click()
    Command17.Enabled = False
    Call FINAL("*CALFERI")
    Command17.Enabled = True
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
    Call CARTADEPOSI
    Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call CARSECPROD
    Command2.Enabled = True
End Sub

Private Sub Command20_Click()
    Command20.Enabled = False
    XX$ = Chr$(1) + "LUNES": Call GRAREG("DIASEM", XX$, 1)
    XX$ = Chr$(2) + "MARTES": Call GRAREG("DIASEM", XX$, 2)
    XX$ = Chr$(3) + "MIERCOLES": Call GRAREG("DIASEM", XX$, 3)
    XX$ = Chr$(4) + "JUEVES": Call GRAREG("DIASEM", XX$, 4)
    XX$ = Chr$(5) + "VIERNES": Call GRAREG("DIASEM", XX$, 5)
    XX$ = Chr$(6) + "SABADO": Call GRAREG("DIASEM", XX$, 6)
    XX$ = Chr$(7) + "DOMINGO": Call GRAREG("DIASEM", XX$, 7)
    Call CIERRARCH("DIASEM")
    Call FINAL("*TURGEN")
    Command20.Enabled = True
End Sub

Private Sub Command21_Click()
    Command21.Enabled = False
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
    Command21.Enabled = True
End Sub
'
Private Sub Command22_Click()
    Command22.Enabled = False
    ACONEC$ = "FLFORMS"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
    Command22.Enabled = True
End Sub

Private Sub Command23_Click()
    '
    Call GRACONTROL("", "PLACARMA", Text1.TEXT)
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONFCARG")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    OPXX% = ALTERNA%("Actualizar Parque de Máquinas\Colores Gráfica de GANTT")
    If OPXX% = 1 Then
         Call FINAL("*PADMAQ")
       ElseIf OPXX% = 2 Then
         COLOMAQ.Show 1
    End If
    Command3.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    ALTX% = ALTERNA%("Generales por Equipo\Especiales Diarios por Equipo")
    If ALTX% = 1 Then
         Call CONECRUN("*TURMAQ", "Horarios Generales por Equipo")
       ElseIf ALTX% = 2 Then
         Call CONECRUN("*TURESP", "Horarios Especiales por Dia y Equipo")
    End If
    Command6.Enabled = True
End Sub

Private Sub Command99_Click()
   Command99.Enabled = False
   Call FINAL("*PARAPROG")
   Command99.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Text1.TEXT = CONTROL$("", "PLACARMA")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    '
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub


