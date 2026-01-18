VERSION 5.00
Begin VB.Form INVEIMPO 
   BackColor       =   &H00DDDDFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventario Materiales de Importacion"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   650
      Left            =   7140
      Picture         =   "INVEIMPO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Control Contra Maestro de Stock"
      Top             =   1700
      Width           =   750
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7140
      ScaleHeight     =   75
      ScaleWidth      =   705
      TabIndex        =   48
      Top             =   2415
      Width           =   760
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00DDDDFF&
      Caption         =   "Reportes y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4530
      Left            =   105
      TabIndex        =   16
      Top             =   2730
      Width           =   7785
      Begin VB.Frame Frame7 
         BackColor       =   &H00DDDDFF&
         Caption         =   "Período Mensual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   4515
         TabIndex        =   44
         Top             =   420
         Width           =   3060
         Begin VB.CommandButton Command3 
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
            Height          =   300
            Left            =   2625
            TabIndex        =   45
            Top             =   525
            Width           =   175
         End
         Begin VB.Label Label27 
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
            Height          =   300
            Left            =   2205
            TabIndex        =   50
            Top             =   525
            Width           =   420
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mes y Año:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   105
            TabIndex        =   47
            Top             =   585
            Width           =   2010
         End
         Begin VB.Label Label25 
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
            Height          =   300
            Left            =   210
            TabIndex        =   46
            Top             =   945
            Width           =   2590
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00DDDDFF&
         Caption         =   "Generar Reporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   4515
         TabIndex        =   37
         Top             =   2100
         Width           =   3060
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2235
            TabIndex        =   38
            Top             =   1365
            Width           =   2295
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
               Picture         =   "INVEIMPO.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   39
               ToolTipText     =   "Listado Resumen de Ajustes Productos Importados"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado Resu- men General"
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
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2235
            TabIndex        =   41
            Top             =   525
            Width           =   2295
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
               Picture         =   "INVEIMPO.frx":0614
               Style           =   1  'Graphical
               TabIndex        =   42
               ToolTipText     =   "Análisis de Ajustes por Grupo de Productos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado Analitico por Grupo"
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
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDFF&
         Caption         =   "Grupo de Materiales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3900
         Left            =   210
         TabIndex        =   18
         Top             =   420
         Width           =   4215
         Begin VB.CommandButton Command5 
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
            Height          =   300
            Left            =   1470
            TabIndex        =   19
            Top             =   525
            Width           =   175
         End
         Begin VB.Label Label28 
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
            Height          =   300
            Left            =   1050
            TabIndex        =   51
            Top             =   525
            Width           =   420
         End
         Begin VB.Label Label24 
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
            Height          =   300
            Left            =   210
            TabIndex        =   36
            Top             =   3360
            Width           =   3795
         End
         Begin VB.Label Label23 
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
            Height          =   300
            Left            =   2730
            TabIndex        =   35
            Top             =   3045
            Width           =   1275
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contrapartida:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   525
            TabIndex        =   34
            Top             =   3105
            Width           =   2115
         End
         Begin VB.Label Label21 
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
            Height          =   300
            Left            =   210
            TabIndex        =   33
            Top             =   2625
            Width           =   3795
         End
         Begin VB.Label Label20 
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
            Height          =   300
            Left            =   2730
            TabIndex        =   32
            Top             =   2310
            Width           =   1275
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            TabIndex        =   31
            Top             =   2370
            Width           =   960
         End
         Begin VB.Label Label18 
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
            Height          =   300
            Left            =   2730
            TabIndex        =   30
            Top             =   1740
            Width           =   1275
         End
         Begin VB.Label Label17 
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
            Height          =   300
            Left            =   2730
            TabIndex        =   29
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cot Fin:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            TabIndex        =   28
            Top             =   1740
            Width           =   960
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cot Ini:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            TabIndex        =   27
            Top             =   1425
            Width           =   960
         End
         Begin VB.Label Label10 
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
            Height          =   300
            Left            =   1365
            TabIndex        =   26
            Top             =   1365
            Width           =   435
         End
         Begin VB.Label Label9 
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
            Height          =   300
            Left            =   210
            TabIndex        =   25
            Top             =   1740
            Width           =   1590
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   315
            TabIndex        =   24
            Top             =   1425
            Width           =   960
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "% Nac:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1890
            TabIndex        =   23
            Top             =   585
            Width           =   960
         End
         Begin VB.Label Label4 
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
            Height          =   300
            Left            =   2940
            TabIndex        =   22
            Top             =   525
            Width           =   1065
         End
         Begin VB.Label Label2 
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
            Height          =   300
            Left            =   210
            TabIndex        =   21
            Top             =   900
            Width           =   3795
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            TabIndex        =   20
            Top             =   585
            Width           =   960
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DDDDFF&
      Caption         =   "Configuración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   105
      TabIndex        =   2
      Top             =   210
      Width           =   6840
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDFF&
         Caption         =   "Indices y Monedas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   210
         TabIndex        =   7
         Top             =   420
         Width           =   3165
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   8
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
               Picture         =   "INVEIMPO.frx":091E
               Style           =   1  'Graphical
               TabIndex        =   9
               ToolTipText     =   "Actualizacion Indices y Cotizaciones Mensuales"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Actualización Valores Mensuales"
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
               TabIndex        =   10
               Top             =   105
               Width           =   1650
            End
         End
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   11
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
               Picture         =   "INVEIMPO.frx":0A68
               Style           =   1  'Graphical
               TabIndex        =   12
               ToolTipText     =   "Actualizar Tabla de Indices y Monedas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Tabla de Indi- ces y Monedas"
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
               TabIndex        =   13
               Top             =   0
               Width           =   1500
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DDDDFF&
         Caption         =   "Base de Datos de Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   3465
         TabIndex        =   3
         Top             =   420
         Width           =   3165
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   4
            Top             =   840
            Width           =   2325
            Begin VB.CommandButton Command8 
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
               Picture         =   "INVEIMPO.frx":0BB2
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "Actualización Valores F.O.B. por Grupo de Productos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Base de Datos de Valores FOB"
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
               TabIndex        =   6
               Top             =   105
               Width           =   1500
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   14
            Top             =   420
            Width           =   2325
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
               Height          =   600
               Left            =   0
               Picture         =   "INVEIMPO.frx":0CFC
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "Actualización Tabla de Clasificación"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Tabla de Clasificación "
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
               Width           =   1500
            End
         End
      End
   End
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
      Height          =   650
      Left            =   7140
      Picture         =   "INVEIMPO.frx":0E46
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   315
      Width           =   750
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
      Height          =   650
      Left            =   7140
      Picture         =   "INVEIMPO.frx":0F90
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1000
      Width           =   750
   End
End
Attribute VB_Name = "INVEIMPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EXTE$
Dim MESAX$(13), INDIME#(13)
Dim INDI$(256)
Dim INDAJ#(256, 13), YACAIN%
Dim PORNA(256), IVARIA#(256), IAJUX%(256)
Dim BULEIN As String * 132
Dim SALINI(32767), SALFIN(32767)
Dim MESANT$

Private Sub Command1_Click()
   Command1.Enabled = False
   Dim ORI%(32767)
   FIN& = NUMAS%
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("MASTER", U& + 1)
     If Asc(Mid$(X$, 72, 1)) = 2 Then
       ORI%(U&) = 1
     End If
   Next U&
   '
   FIN& = ULTREG&("FOBMAIMP")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("FOBMAIMP", U&)
     CIXI% = CVI(Left$(X$, 2))
     If CIXI% > 0 Then
       If CIXI% <= NUMAS% Then
         If ORI%(CIXI%) = 1 Then
           ORI%(CIXI%) = 2
         End If
       End If
     End If
   Next U&
   '
   JJ& = 0
   FIN& = NUINV%
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("INVERT", U& + 1)
     If Asc(Mid$(X$, 16, 1)) <> 96 Then
       CIXI% = CVI(Mid$(X$, 17, 2))
       If CIXI% > 0 Then
         If CIXI% <= NUMAS% Then
           If ORI%(CIXI%) = 1 Then ' no está en FOBMAIMP
             Z$ = REGBLAN$("IMPNOCLA")
             Call REPLA(Z$, MKI$(CIXI%), 1, 2)
             Call REPLA(Z$, MKS$(STOCKACT(CIXI%)), 3, 4)
             JJ& = JJ& + 1
             Call GRAREG("IMPNOCLA", Z$, JJ&)
           End If
         End If
       End If
     End If
   Next U&
   Call SETULTREG("IMPNOCLA", JJ&)
   Call CIERRARCH("*.*")
   Call CONECRUN("*LINOCLA", "Listado Importados no Clasificados")
   '
   Command1.Enabled = True
End Sub

Private Sub Command18_Click()
   '
   Command18.Enabled = False
   If ULTREG&("TAMONED") < 1 Then
      Call MENSERR(24, "Falta Definir Tabla\de Indices y Monedas de Ajuste.")
      GoTo 99
   End If
   '
   MESANT$ = ""
10 REMO$ = REGSEL$("TAMONED")
   If Len(REMO$) <= 4 Then GoTo 99
   '
   INDIAJ% = Asc(Left$(REMO$, 1))
   DENOIN$ = TRIM$(Mid$(REMO$, 2, 20))
   '
   Call CARTAMESES
   Call LEEINDIME(INDIAJ%)
   '
   For U& = 0 To 12
      TTXX$ = String$(24, 0)
      Call REPLA(TTXX$, Chr$(U&), 1, 1)
      Call REPLA(TTXX$, MESAX$(U&), 2, 15)
      Call REPLA(TTXX$, MKD$(INDIME#(U&)), 17, 8)
      Call GRAREG("!TAINDAJ", TTXX$, U& + 1)
   Next U&
   Call SETULTREG("!TAINDAJ", 13)
   Call CIERRARCH("!TAINDAJ")
   '
   VTB% = VENTABU%("TAINDAJ/NC/TITUPAN=Indice de Ajuste: " + DENOIN$)
   If VTB% >= 0 Then
     EXTE$ = "DAT": If EJER% < 0 Then EXTE$ = "OLD"
     N1% = FILEOPEN%(LUDAT$ + "INDAJU." + EXTE$, 0, 132)
     CAREIN& = LOF(N1%) / 132
     If INDIAJ% <= CAREIN& Then
       Get #N1%, INDIAJ%, BULEIN$
       TTXX$ = BULEIN$
       For U& = 0 To 12
         XX$ = REGLEIDO$("!TAINDAJ", U& + 1)
         Call REPLA(TTXX$, Mid$(XX$, 17, 8), 21 + 8 * U&, 8)
       Next U&
       BULEIN$ = TTXX$
       Put #N1%, INDIAJ%, BULEIN$
       Close #N1%: N1% = 0
       Call CIERRARCH("*.*")
     End If
     GoTo 10
     '
   End If
   '
99 YACAIN% = 0
   Command18.Enabled = True
   '
End Sub

Private Sub Command19_Click()
   '
   Command19.Enabled = False
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAINDICE/NC/TITUPAN=Indices y Monedas")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", X$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
   Command19.Enabled = True
   '
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call LIANASTO
   Command2.Enabled = True
End Sub

Private Sub Command22_Click()
   '
   Command22.Enabled = False
   Call TRALOCAL("GRUMAIMP", "")
   VTB% = VENTABU%("GRUMAIMP/NC/TITUPAN=Tabla de Clasificación - Productos Importados")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!GRUMAIMP")
         X$ = REGLEIDO$("!GRUMAIMP", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("GRUMAIMP", X$, U&)
      Next U&
      FIN& = ULTREG&("!GRUMAIMP")
      Call SETULTREG("GRUMAIMP", FIN&)
   End If
   Call CIERRARCH("GRUMAIMP")
   Call FRESHECHO("GRUMAIMP")
   MESANT$ = ""
   Command22.Enabled = True
   '
End Sub

Private Sub Command24_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command29_Click()
   Call HELPONLINE("AJUINFLA")
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call CARTAMESES
   '
   JJ& = 0
   For U& = 1 To 12
      TTXX$ = String$(24, 0)
      Call REPLA(TTXX$, Chr$(U&), 1, 1)
      Call REPLA(TTXX$, MESAX$(U&), 2, 15)
      JJ& = JJ& + 1
      Call GRAREG("!LIMESEJ", TTXX$, JJ&)
   Next U&
   Call SETULTREG("!LIMESEJ", 12)
   Call CIERRARCH("!LIMESEJ")
   '
   XX$ = REGSEL$("!LIMESEJ")
   If Len(XX$) > 4 Then
     Label27 = Asc(Left$(XX$, 15))
     Label25 = Mid$(XX$, 2, 15)
     If Val(Label27) > 0 Then
       Call CARMAINDI
       INDIAJ% = Val(Label10)
       Label18 = FORMATNUM(INDAJ#(INDIAJ%, Val(Label27)), "F10.4")
       Label17 = FORMATNUM(INDAJ#(INDIAJ%, Val(Label27) - 1), "F10.4")
     End If
     '
   End If
   '
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call REANASTO
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   XX$ = REGSEL$("GRUMAIMP")
   If Len(XX$) > 4 Then
     Label28 = Asc(Left$(XX$, 15))
     Label2 = Mid$(XX$, 2, 31)
     Label10 = Asc(Mid$(XX$, 33, 1))
     Label9 = ECOARCH$("TAMONED", Chr$(Val(Label10)))
     Label4 = FORMATNUM(CVS(Mid$(XX$, 35, 4)), "F7.2")
     Label20 = Mid$(XX$, 39, 12)
     Label21 = ECOARCH$("ECUECON", Label20)
     Label23 = Mid$(XX$, 51, 12)
     Label24 = ECOARCH$("ECUECON", Label23)
     '
     If Val(Label27) > 0 Then
       Call CARMAINDI
       INDIAJ% = Val(Label10)
       Label18 = FORMATNUM(INDAJ#(INDIAJ%, Val(Label27)), "F10.4")
       Label17 = FORMATNUM(INDAJ#(INDIAJ%, Val(Label27) - 1), "F10.4")
     End If
     '
   End If
   '

End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   If ULTREG&("GRUMAIMP") < 1 Then
      Call MENSERR(24, "Falta Definir Grupos\de Productos de Importacion.")
      GoTo 99
   End If
   '
   MESANT$ = ""
10 REMO$ = REGSEL$("GRUMAIMP")
   If Len(REMO$) <= 4 Then GoTo 99
   '
   Screen.MousePointer = 11
   GRUMA% = Asc(Left$(REMO$, 1))
   DENOGRU$ = TRIM$(Mid$(REMO$, 2, 31))
   MONEFO% = Asc(Mid$(REMO$, 33, 1))
   DEMONE$ = ECOARCH$("TAMONED", Chr$(MONEFO%))
   '
   RFF$ = RECFILT$("FOBMAIMP", 3, MKI$(GRUMA%))
   JJ& = 0
   For U& = 1 To Len(RFF$) Step 4
     REG& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("FOBMAIMP", REG&)
     JJ& = JJ& + 1
     Call GRAREG("!FOBMAIMP", XX$, JJ&)
   Next U&
   '
   Call SETULTREG("!FOBMAIMP", JJ&)
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
   VTB% = VENTABU%("FOBMAIMP/TITUPAN=Valores FOB (" + TRIM$(DEMONE$) + ") - " + DENOGRU$)
   '
   If VTB% >= 0 Then
     Screen.MousePointer = 11
     KK& = -3
     For U& = 1 To ULTREG&("!FOBMAIMP")
       XX$ = REGLEIDO$("!FOBMAIMP", U&)
       Call REPLA(XX$, MKI$(GRUMA%), 3, 2)
       CI% = CVI(Left$(XX$, 2))
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           RGG$ = RECFILT$("FOBMAIMP", 1, MKI$(CI%))
           If Len(RGG$) >= 4 Then
             REG& = CVS(Left$(RGG$, 4))
             YY$ = REGLEIDO$("FOBMAIMP", REG&)
             If CVI(Left$(YY$, 2)) = CI% Then
               Call GRAREG("FOBMAIMP", String$(128, 0), REG&)
             End If
           End If
           '
           KK& = KK& + 4
           If KK& <= Len(RFF$) Then
               REG& = CVS(Mid$(RFF$, KK&, 4))
             Else
               REG& = 1 + ULTREG&("FOBMAIMP")
           End If
           Call GRAREG("FOBMAIMP", XX$, REG&)
           '
         End If
       End If
     Next U&
     '
     For KK1& = KK& + 4 To Len(RFF$)
       REG& = CVS(Mid$(RFF$, KK1&, 4))
       Call GRAREG("FOBMAIMP", String$(128, 0), REG&)
     Next KK1&
     Call CIERRARCH("FOBMAIMP")
     Call CIERRARCH("*.*")
     Screen.MousePointer = 1
   End If
   GoTo 10
   '
99 Command8.Enabled = True
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + " - " + EJACT$ + "  -  " + EPAC$
    End If
    '
    EXTE$ = "DAT": If EJER% < 0 Then EXTE$ = "OLD"
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Sub CARTAMESES()
   '
   MESAX$(0) = "APERTURA"
   '
   FECHINVE$ = CONTROL$("INVESABO", "FECHINVE")
   INIANO$ = "01/01" + Mid$(FECHINVE$, 6)
   FEX = FECHANUM(INIANO$)  ' DIANTE(FECHANUM(INIANO$))
   '
   'FEX = COMEJ%
   MESAX$(1) = METEXTO$(Int(FEX))
   JJ% = 1
10 MEBA% = Val(Mid$(FECHATEX$(FEX), 4, 2))
12 FEX = DIPOST(FEX)
   If Val(Mid$(FECHATEX$(FEX), 4, 2)) <> MEBA% Then
     JJ% = JJ% + 1
     MESAX$(JJ%) = METEXTO$(Int(FEX))
     If JJ% < 12 Then GoTo 10
     GoTo 20
   End If
   GoTo 12
   '
20 JJ% = JJ% + 1
   MESAX$(JJ%) = "CIERRE"
   '
End Sub

Sub LEEINDIME(NUINDI%)
    '
    N1% = FILEOPEN%(LUDAT$ + "INDAJU." + EXTE$, 0, 132)
    CAREIN& = LOF(N1%) / 132
    For U& = 1 To ULTREG&("TAMONED")
      DENOIN$ = Mid$(REGLEIDO$("TAMONED", U&), 2, 20)
      If U& <= CAREIN& Then
           Get #N1%, U&, BULEIN$
           TTXX$ = BULEIN$
         Else
           TTXX$ = Space$(20) + String$(112, 0)
      End If
      INDI$(U&) = DENOIN$
      Call REPLA(TTXX$, DENOIN$, 1, 20)
      BULEIN$ = TTXX$
      Put #N1%, U&, BULEIN$
      '
      If U& = NUINDI% Then
        For KKI% = 0 To 13
          INDIME#(KKI%) = CVD(Mid$(BULEIN$, 21 + 8 * KKI%, 8))
        Next KKI%
      End If
      '
    Next U&
    Close #N1%
    '
End Sub

Sub CARMAINDI()
    '
    EXTE$ = "DAT": If EJER% < 0 Then EXTE$ = "OLD"
    N1% = FILEOPEN%(LUDAT$ + "INDAJU." + EXTE$, 0, 132)
    CAREIN& = LOF(N1%) / 132
    For U& = 1 To CAREIN&
      Get #N1%, U&, BULEIN$
      TTXX$ = BULEIN$
      INDA1# = 0
      For KKI% = 0 To 12
        INDA2# = CVD(Mid$(BULEIN$, 21 + 8 * KKI%, 8))
        If INDA2# > 0.0001 Then INDA1# = INDA2#
        INDAJ#(U&, KKI%) = INDA1#
      Next KKI%
    Next U&
    Close #N1%
    '
    For U& = 1 To ULTREG&("GRUMAIMP")
      XX$ = REGLEIDO$("GRUMAIMP", U&)
      If U& <= 256 Then
        PORNA(U&) = CVS(Mid$(XX$, 35, 4))
        IAJUX%(U&) = Asc(Mid$(XX$, 33, 1))
      End If
    Next U&
    Call CIERRARCH("GRUMAIMP")
    YACAIN% = 1
    '
End Sub

Sub LIANASTO()
    '
    MESASEL$ = TRIM$(Label25)
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Exit Sub
    End If
    '
    If Val(Label28) < 1 Then
      Call MENSERR(24, "Seleccione Grupo de Productos.")
      Exit Sub
    End If
    '
    Call CALCUREMO
    '
    GRUMA% = Label28
    RFF$ = RECFILT$("FOBMAIMP", 3, MKI$(GRUMA%))
    JJ& = 0
    For U& = 1 To Len(RFF$) Step 4
      REG& = CVS(Mid$(RFF$, U&, 4))
      XX$ = REGLEIDO$("FOBMAIMP", REG&)
      CI% = CVI(Left$(XX$, 2))
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          JJ& = JJ& + 1
          Call GRAREG("LIANAFOB", XX$, JJ&)
        End If
      End If
    Next U&
    '
    Call SETULTREG("LIANAFOB", JJ&)
    Call CIERRARCH("LIANAFOB")
    '
    Call HEADERS("LIANAFOB", "")
    Call HEADERS("LIANAFOB", "Corresponde a: " + TRIM$(Label2) + " - " + TRIM$(Label25))
    '
    Call FINAL("*LIANAFOB")
    Screen.MousePointer = 1
    '
End Sub

Sub REANASTO()
    '
    MESASEL$ = TRIM$(Label25)
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Exit Sub
    End If
    '
    If Val(Label28) < 1 Then
      Call MENSERR(24, "Seleccione Grupo de Productos.")
      Exit Sub
    End If
    '
    Call CALCUREMO
    '
    For KKI% = 0 To 256
      IVARIA#(KKI%) = 0
    Next KKI%
    '
    For REG& = 1 To ULTREG&("FOBMAIMP")
      XX$ = REGLEIDO$("FOBMAIMP", REG&)
      CI% = CVI(Left$(XX$, 2))
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          GRUMAI% = CVI(Mid$(XX$, 3, 2))
          If GRUMAI% > 0 Then
            If GRUMAI% <= 256 Then
              IVARIA#(GRUMAI%) = IVARIA#(GRUMAI%) + CVD(Mid$(XX$, 105, 8))
            End If
          End If
        End If
      End If
    Next REG&
    '
    For U& = 1 To ULTREG&("GRUMAIMP")
      XX$ = REGLEIDO$("GRUMAIMP", U&)
      Call REPLA(XX$, MKD$(IVARIA#(U&)), 65, 8)
      Call GRAREG("GRUMAIMP", XX$, U&)
    Next U&
    Call CIERRARCH("GRUMAIMP")
    '
    Call HEADERS("REANAFOB", "")
    Call HEADERS("REANAFOB", "Corresponde a: " + TRIM$(Label25))
    '
    Call FINAL("*REANAFOB")
    Screen.MousePointer = 1
    '
End Sub
'
Sub CALCUREMO()
    '
    MESASEL$ = TRIM$(Label25)
    If MESANT$ <> "" Then
      If MESASEL$ = MESANT$ Then Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Call MESANO$(MESASEL$, DES%, HAS%)
700 FEX = DES%: DESX$ = FECHATEX$(FEX)
710 FEX = HAS%: HASX$ = FECHATEX$(FEX)
    '
    Call CARMAINDI
    '
    For KKI% = 0 To NUMAS%
      SALINI(KKI%) = 0: SALFIN(KKI%) = 0
    Next KKI%
    '
    LU$ = LUDAT$
    CAMO& = ULTREG&("MOVISTO")
    Dim GG As String * 128
    AMOVI% = FILEOPEN%(LU$ + "MOVISTO.DAT", 0, 128)
    For I& = 1 To CAMO&
       Call TRACE(20, I&, CAMO&)
       Get #AMOVI%, I&, GG$
       FECHA% = CVI(Left$(GG$, 2))
       If FECHA% <= HAS% Then
         CI% = CVI(Mid$(GG$, 3, 2))
         If CI% < 1 Or CI% > NUMAS% Then GoTo 799
         '
         CAN# = CVD(Mid$(GG$, 88, 8)) - CVD(Mid$(GG$, 96, 8))
         If FECHA% < DES% Then SALINI(CI%) = SALINI(CI%) + CAN#
         If FECHA% < HAS% Then SALFIN(CI%) = SALFIN(CI%) + CAN#
         '
       End If
799 Next I&
    '
    IMES% = Val(Label27)
    For REG& = 1 To ULTREG&("FOBMAIMP")
      XX$ = REGLEIDO$("FOBMAIMP", REG&)
      CI% = CVI(Left$(XX$, 2))
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          GRUMAI% = CVI(Mid$(XX$, 3, 2))
          VAFOB# = CVD(Mid$(XX$, 5, 8))
          Call REPLA(XX$, MKS$(SALINI(CI%)), 65, 4)
          VUINI# = VAFOB# * INDAJ#(IAJUX%(GRUMAI%), IMES% - 1) * (1 + PORNA(GRUMAI%) / 100)
          Call REPLA(XX$, MKD$(VUINI#), 69, 8)
          IMINI# = VUINI# * SALINI(CI%)
          Call REPLA(XX$, MKD$(IMINI#), 77, 8)
          Call REPLA(XX$, MKS$(SALFIN(CI%)), 85, 4)
          VUFIN# = VAFOB# * INDAJ#(IAJUX%(GRUMAI%), IMES%) * (1 + PORNA(GRUMAI%) / 100)
          Call REPLA(XX$, MKD$(VUFIN#), 89, 8)
          IMFIN# = VUFIN# * SALFIN(CI%)
          Call REPLA(XX$, MKD$(IMFIN#), 97, 8)
          Call REPLA(XX$, MKD$(IMFIN# - IMINI#), 105, 8)
          Call GRAREG("FOBMAIMP", XX$, REG&)
        End If
      End If
    Next REG&
    '
    Call CIERRARCH("FOBMAIMP")
    Screen.MousePointer = 1
    '
End Sub
