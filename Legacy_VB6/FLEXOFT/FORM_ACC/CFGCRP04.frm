VERSION 5.00
Begin VB.Form CFGCRP04 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de Máquinas (CRP) - Configuración"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4830
      TabIndex        =   45
      Top             =   5670
      Width           =   4005
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
         TabIndex        =   51
         Top             =   1890
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
         TabIndex        =   50
         Top             =   1890
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
         TabIndex        =   49
         Top             =   1155
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
         TabIndex        =   48
         Top             =   1155
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
         TabIndex        =   47
         Top             =   420
         Width           =   1590
      End
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
         TabIndex        =   46
         Top             =   420
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00004000&
      Height          =   5000
      Left            =   8400
      ScaleHeight     =   4935
      ScaleWidth      =   1005
      TabIndex        =   40
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command24 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "CFGCRP04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Disponibilidad de Mano de Obra"
         Top             =   2835
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "CFGCRP04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   920
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "CFGCRP04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Terminar - Finalizar la Aplicación"
         Top             =   280
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "CFGCRP04.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1680
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CFGCRP04.frx":0A68
         Top             =   3885
         Width           =   1515
      End
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
      Height          =   1485
      Left            =   4200
      TabIndex        =   35
      Top             =   2835
      Width           =   4005
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   260
         TabIndex        =   54
         Text            =   " "
         Top             =   680
         Width           =   2850
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2100
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1050
         Width           =   495
      End
      Begin VB.CommandButton Command22 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   3200
         Picture         =   "CFGCRP04.frx":298A
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   315
         Width           =   550
      End
      Begin VB.CommandButton Command21 
         Caption         =   "."
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
         Left            =   2910
         TabIndex        =   37
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1550
         TabIndex        =   36
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proyección en Días:"
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
         Left            =   -840
         TabIndex        =   52
         Top             =   1155
         Width           =   2850
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -840
         TabIndex        =   39
         Top             =   410
         Width           =   2325
      End
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
      Height          =   4110
      Left            =   210
      TabIndex        =   8
      Top             =   210
      Width           =   3795
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "General de Planta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1750
         Left            =   210
         TabIndex        =   10
         Top             =   315
         Width           =   3375
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   23
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
               Picture         =   "CFGCRP04.frx":2C94
               Style           =   1  'Graphical
               TabIndex        =   24
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
               TabIndex        =   25
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   26
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
               Picture         =   "CFGCRP04.frx":2F9E
               Style           =   1  'Graphical
               TabIndex        =   27
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
               TabIndex        =   28
               Top             =   0
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Máquinas y Equipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1750
         Left            =   210
         TabIndex        =   9
         Top             =   2205
         Width           =   3375
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   29
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
               Picture         =   "CFGCRP04.frx":32A8
               Style           =   1  'Graphical
               TabIndex        =   30
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
               TabIndex        =   31
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   32
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
               Picture         =   "CFGCRP04.frx":35B2
               Style           =   1  'Graphical
               TabIndex        =   33
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
               TabIndex        =   34
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
      Height          =   2535
      Left            =   4200
      TabIndex        =   0
      Top             =   210
      Width           =   4005
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   11
         Top             =   1680
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
            Picture         =   "CFGCRP04.frx":38BC
            Style           =   1  'Graphical
            TabIndex        =   22
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
            TabIndex        =   12
            Top             =   105
            Width           =   1605
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   16
         Top             =   1260
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
            Picture         =   "CFGCRP04.frx":3BC6
            Style           =   1  'Graphical
            TabIndex        =   17
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
            TabIndex        =   18
            Top             =   0
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   13
         Top             =   840
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
            Picture         =   "CFGCRP04.frx":3ED0
            Style           =   1  'Graphical
            TabIndex        =   14
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
            TabIndex        =   15
            Top             =   0
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   19
         Top             =   420
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
            Picture         =   "CFGCRP04.frx":41DA
            Style           =   1  'Graphical
            TabIndex        =   20
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
            TabIndex        =   21
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
Attribute VB_Name = "CFGCRP04"
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
    Call FINAL("*TURDIAR")
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
    'AGREGUE ACA
    PERIOC$ = TRIM$(FORMATNUM$(Text3.TEXT, "F4.0"))
    Call GRACONTROL("", "PERDIA", PERIOC$)
    
    Call GRACONTROL("", "PLACARMA", Text1.TEXT)
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command24_Click()
   Command24.Enabled = False
   OPXX% = 1 'OPXX% = COMALTER%("Puede Ajustar la Disponibilidad de M.O.D.\\General Por Sección\Por Fecha y Sección")
   Select Case OPXX%
     Case 1: Call FINAL("*DOTASEG")
     Case 2: Call FINAL("*DOTASEF")
   End Select
   Command24.Enabled = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONFCARG")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    OPXX% = ALTERNA%("Parque de Máquinas y Equipos\Definir Lineas Transferenciales")
    If OPXX% = 1 Then
         Call FINAL("*PADMAQ")
       ElseIf OPXX% = 2 Then
         Call FINAL("*LINTRAN")
    End If
'    OPXX% = ALTERNA%("Actualizar Parque de Máquinas\Colores Gráfica de GANTT")
'    If OPXX% = 1 Then
'         Call FINAL("*PADMAQ")
'       ElseIf OPXX% = 2 Then
'         COLOMAQ.Show 1
'    End If
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
    'AGREGADO
    Text3.TEXT = CONTROL$("", "PERDIA")
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


Private Sub Text3_LostFocus()
'
Text3.TEXT = CSTRING$(MKS$(XVALO(Text3.TEXT)), 3, 2, 0, 2)
'
End Sub
