VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FiNVEN07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventarios de Materiales - Generación y Consulta"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   57
      Top             =   360
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   8600
      ScaleHeight     =   6705
      ScaleWidth      =   1980
      TabIndex        =   51
      Top             =   0
      Width           =   2010
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   59
         Top             =   3840
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   60
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Mast."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "FINVEN07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Alta de Nuevo Item de Stock"
         Top             =   1730
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "FINVEN07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   900
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "FINVEN07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Direct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "FINVEN07.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2520
         Width           =   650
      End
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   7665
      ScaleHeight     =   585
      ScaleWidth      =   1635
      TabIndex        =   26
      Top             =   7455
      Width           =   1695
      Begin VB.CommandButton Command16 
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
         Picture         =   "FINVEN07.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Costos Informados"
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
         TabIndex        =   28
         Top             =   105
         Width           =   960
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Fábrica"
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
      Left            =   4515
      TabIndex        =   19
      Top             =   7770
      Width           =   3585
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   23
         Top             =   420
         Width           =   2535
         Begin VB.CommandButton Command9 
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
            Picture         =   "FINVEN07.frx":09F2
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Totales por Material"
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
            TabIndex        =   25
            Top             =   0
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   20
         Top             =   840
         Width           =   2535
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
            Left            =   -20
            Picture         =   "FINVEN07.frx":0B3C
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Número de Orden de Fabricación"
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
            Left            =   840
            TabIndex        =   22
            Top             =   0
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Proveedores (Consignación)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   210
      TabIndex        =   9
      Top             =   7665
      Width           =   3585
      Begin VB.PictureBox Picture10 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   16
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command12 
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
            Picture         =   "FINVEN07.frx":0E46
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Clasificado     por Proveedor"
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
            Top             =   0
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   13
         Top             =   840
         Width           =   2325
         Begin VB.CommandButton Command11 
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
            Picture         =   "FINVEN07.frx":1150
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Material"
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
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   10
         Top             =   1260
         Width           =   2325
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
            Left            =   -20
            Picture         =   "FINVEN07.frx":145A
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Materiales Reservados     a Consignar"
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
            Left            =   840
            TabIndex        =   12
            Top             =   0
            Width           =   1380
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "REPORTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   4440
      TabIndex        =   2
      Top             =   105
      Width           =   4005
      Begin VB.Frame Frame9 
         BackColor       =   &H00CDDADA&
         Caption         =   "VALORIZACIÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   2050
         TabIndex        =   42
         Top             =   2625
         Width           =   1750
         Begin VB.OptionButton Option5 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Importes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   315
            TabIndex        =   44
            Top             =   720
            Width           =   1065
         End
         Begin VB.OptionButton Option6 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Cantidades"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   315
            TabIndex        =   43
            Top             =   470
            Value           =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00CDDADA&
         Caption         =   "DISPOSITIVO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   210
         TabIndex        =   39
         Top             =   2625
         Width           =   1750
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Pantalla"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   315
            TabIndex        =   41
            Top             =   470
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Impresión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   315
            TabIndex        =   40
            Top             =   735
            Width           =   1170
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00CDDADA&
         Caption         =   "CONSULTAS Y LISTADOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   210
         TabIndex        =   29
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   945
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   30
            Top             =   1260
            Width           =   2325
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
               Height          =   645
               Left            =   -20
               Picture         =   "FINVEN07.frx":15A4
               Style           =   1  'Graphical
               TabIndex        =   31
               ToolTipText     =   "Inventario General Discriminado en Parciales por Depósito"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "General Analítico    por Depósito"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   750
               Left            =   700
               TabIndex        =   32
               Top             =   100
               Width           =   1620
            End
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   600
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   33
            Top             =   840
            Width           =   2325
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
               Height          =   645
               Left            =   -20
               Picture         =   "FINVEN07.frx":18AE
               Style           =   1  'Graphical
               TabIndex        =   34
               ToolTipText     =   "Existencias en un Unico Depósito"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Parcial por Depósito"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   720
               TabIndex        =   35
               Top             =   100
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   315
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   36
            Top             =   420
            Width           =   2325
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
               Height          =   645
               Left            =   0
               Picture         =   "FINVEN07.frx":1BB8
               Style           =   1  'Graphical
               TabIndex        =   37
               ToolTipText     =   "Inventario General (Totales por Item en Todos los Depósitos)"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Inventario General"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   720
               TabIndex        =   38
               Top             =   100
               Width           =   1485
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "GENERACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   4215
      Begin VB.Frame Frame7 
         BackColor       =   &H00CDDADA&
         Caption         =   "COMPLEMENTOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   45
         Top             =   2625
         Width           =   3795
         Begin VB.PictureBox Picture12 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   1920
            ScaleHeight     =   615
            ScaleWidth      =   1665
            TabIndex        =   49
            Top             =   350
            Width           =   1695
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
               Height          =   645
               Left            =   0
               Picture         =   "FINVEN07.frx":1EC2
               Style           =   1  'Graphical
               TabIndex        =   56
               ToolTipText     =   "Ajuste de Existencias"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Ajustes de Stocks"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   50
               Top             =   105
               Width           =   960
            End
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   210
            ScaleHeight     =   615
            ScaleWidth      =   1575
            TabIndex        =   46
            Top             =   350
            Width           =   1600
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
               Height          =   645
               Left            =   -20
               Picture         =   "FINVEN07.frx":200C
               Style           =   1  'Graphical
               TabIndex        =   47
               ToolTipText     =   "Listado de Captura de Datos para Inventario"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Tablas de Captura"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   48
               Top             =   105
               Width           =   960
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         Caption         =   "INVENTARIO CONTINUO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   210
         TabIndex        =   1
         Top             =   420
         Width           =   3795
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   315
            TabIndex        =   8
            Text            =   " "
            Top             =   735
            Width           =   2955
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2205
            TabIndex        =   6
            Text            =   " "
            Top             =   315
            Width           =   1065
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   525
            ScaleHeight     =   615
            ScaleWidth      =   2505
            TabIndex        =   3
            Top             =   1260
            Width           =   2535
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
               Height          =   645
               Left            =   0
               Picture         =   "FINVEN07.frx":2316
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Re-Calcular Existencias a Fecha"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   720
               TabIndex        =   5
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Inventario"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   210
            TabIndex        =   7
            Top             =   450
            Width           =   1905
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3480
      OleObjectBlob   =   "FINVEN07.frx":2620
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FINVEN07.frx":2854
      TabIndex        =   58
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "FiNVEN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
    '
    If Option1.Value = True Then
        Call FINAL("*COMATCO")
      ElseIf Option2.Value = True Then
        Call FINAL("*LIMATCO")
    End If
    '
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    ACONEC$ = "ARCHIG07/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command13.Enabled = True
End Sub

Private Sub Command16_Click()
    Call CONECRUN("AMASTO02", "Maestro de Items de Stock")
End Sub

Private Sub Command17_Click()
    PREPA% = 0
    FORMAJU% = 0
    AJUSTK07.Show
End Sub

Private Sub Command2_Click()
    Command13.Enabled = False
    OPX% = ALTERNA%("Tabla General de Captura\Material no Trazable\Trazable por Lotes\Por Numero de Serie")
    If OPX% = 1 Then
         ACONEC$ = "ARCHIG07/LIMASTER"
         Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
       ElseIf OPX% = 2 Then
         Call FINAL("?LINVENOTRA")
       ElseIf OPX% = 3 Then
         Call FINAL("?LISLOTEINV")
       ElseIf OPX% = 4 Then
         Call FINAL("?LISINVNUSER")
    End If
    Command13.Enabled = True
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    If DERACCE%("AJUSTOCK", "Ajustes de Stock") >= 2 Then
       PREPA% = 0
       FORMAJU% = 0
       AJUSTK07.Show 1
    End If
    Command25.Enabled = True
End Sub

Private Sub Command3_Click()
    If Option1.Value = True Then
        If Option6.Value = True Then
            Call FINAL("*COINGEN1")
          ElseIf Option5.Value = True Then
            If DERACCE%("INVEGEN", "Inventario General Valorizado") > 1 Then
                Call FINAL("*COINGEN2")
            End If
        End If
      ElseIf Option2.Value = True Then
        If Option5.Value = True Then
            If DERACCE%("INVEGEN", "Inventario General Valorizado") < 2 Then GoTo 99
        End If
        Call GENINV3            ' filtrar y listar
    End If
99  Exit Sub
End Sub

Private Sub Command4_Click()
    '
    If Option1.Value = True Then
        If Option6.Value = True Then
            Call FINAL("*COINDEP1")
          ElseIf Option5.Value = True Then
            If DERACCE%("INVPARDE", "Inventario Parcial por Deposito Valorizado") > 1 Then
                Call FINAL("*COINDEP2")
            End If
        End If
      ElseIf Option2.Value = True Then
        If Option5.Value = True Then
            If DERACCE%("INVPARDE", "Inventario Parcial por Deposito Valorizado") < 2 Then GoTo 99
        End If
        Call GENINV4            ' filtrar y listar
    End If
    '
99  Exit Sub
End Sub

Private Sub Command5_Click()
    '
    If Option1.Value = True Then
        If Option6.Value = True Then
            Call FINAL("*COINANA1")
          ElseIf Option5.Value = True Then
            If DERACCE%("INVANALI", "Inventario Analitico Valorizado") > 1 Then
                Call FINAL("*COINANA2")
            End If
        End If
      ElseIf Option2.Value = True Then
        If Option5.Value = True Then
            If DERACCE%("INVANALI", "Inventario Analitico Valorizado") < 2 Then GoTo 99
        End If
        Call GENINV5            ' filtrar y listar
    End If
    '
99  Exit Sub
End Sub

Private Sub Command7_Click()
    GENINV07.Show 1
End Sub

Private Sub Command8_Click()
    Call ACCDIR("RECEPMA")
End Sub

Private Sub Form_Load()
    '
    VERANTER$ = "INVENTAR"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = EMPREAC$
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture13)
    '
    FECHAINV$ = TRIM$(Control$("INVENTAR", "FECHAINV"))
    If FECHANUM(FECHAINV$) < 1 Then
        FECHAINV$ = ""
    End If
    Text1.TEXT = FECHAINV$
    Text2.TEXT = TRIM$(Control$("INVENTAR", "FEGENINV"))
    Show
    Refresh
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Option5_LostFocus()
  '\\\OT-06-0172-WAR-10/05/06///
  If DERACCE%("INVEVAL", "Inventario Valorizado") < 2 Then
    Option6.Value = 1
    Option5.Value = 0
  End If
  '\\\OT-06-0172-WAR-FIN///
End Sub

Private Sub Text1_GotFocus()
    Command7.SetFocus
End Sub

Private Sub Text2_GotFocus()
    Command7.SetFocus
End Sub

Sub GENINV3()
    '
    LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
    If LONCODI% < 16 Or LONCODI% > 24 Then LONCODI% = 16

10  CI1% = PRICOD%
    CI2% = ULTCOD%
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0) + MKI$(0)
    '
    OBX$ = OBJPLA$("DESHASCOD3", VDEF$)
    If Len(OBX$) < 5 Then
        Exit Sub
    End If
    COD1$ = CODEXT$(CVI(Left$(OBX$, 2)))
    COD2$ = CODEXT$(CVI(Mid$(OBX$, 3, 2)))
    TMT% = Asc(Mid$(OBX$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(OBX$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación por Tipo"
    End If
    '
    GCV% = CVI(Mid$(OBX$, 6, 2))
    GRUCOS$ = ECOARCH$("GRUCOVEN", MKI$(GCV%))
    If GRUCOS$ = "" Then
        GCV% = 0
        Call REPLA(OBX$, MKI$(0), 6, 2)
        GRUCOS$ = "Sin Clasificación Contable"
    End If
    '
    If TRIM$(COD1$) = "" Or TRIM$(COD2$) = "" Or COD2$ < COD1$ Then
        GoTo 10
    End If
    '
    ModoInventario% = XVALO(Control("INVENTAR", "MODOINVE"))
    '
    Call COPYSTRU("INVEGEN", "INVEGENF")
    Call BLANARCH("INVEGENF")
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("INVEGEN")
    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
    '
    For i& = 1 To FIN&
        Call TRACE(20, i&, FIN&)
        X$ = REGLEIDO$("INVEGEN", i&)
        CI% = CVI(Left$(X$, 2))
        If CI% > 0 Then
           If CI% <= NUMAS% Then
              COD$ = CODEXT$(CI%)
              If COD$ >= COD1$ Then
                 If COD$ <= COD2$ Then
                   If TMT% = 0 Or TIMATE%(CI%) = TMT% Then
                     'If GCV% = 0 Or GRUCOVE%(CI%) = GCV% Then
                       Call REGAPP("INVEGENF", X$)
                     'End If
                   End If
                 End If
              End If
           End If
        End If
    Next i&
    '
    Call HEADERS("INVEGENF", "")
    Call HEADERS("INVEGENF", "Material: " + TRIM$(TIMAT$) + " - " + TRIM$(GRUCOS$))
    '
    Call CIERRARCH("INVEGENF")
    Screen.MousePointer = 1
'    '
    If Option6.Value = True Then
         Call FINAL("*LIINGEN1")
       ElseIf Option5.Value = True Then
         '\\\OT-07-0018-RG-23/01/07///
         MON$ = Control$("INVENTAR", "MONEDINV")
         MONDESC$ = ECOARCH$("TAMONED", Chr(Val(MON$)))
         Call HEADERS("INVEGENF", "Valorización: " + TRIM$(MON$) + " - " + TRIM$(MONDESC$))
         If Control("INVENTAR", "ITMONELE") = "SI" Then Call HEADERS("INVEGENF", "Condicion: Solo Articulos Valorizados Con la Moneda Elegida")
         '\\\OT-07-0018-RG-FIN///
         If ModoInventario% <> 7 Then
            Call FINAL("*LIINGEN2")
         Else
            Call FINAL("*LIINGEN3")
         End If
    End If
    '
    '
'    If FiNVEN04.Option6.Value = True Then
'         Call FINAL("*LIINANA1")
'       ElseIf FiNVEN04.Option5.Value = True Then
'         Call FINAL("*LIINANA2")
'    End If
    '
End Sub
'

Sub GENINV4()
    
    '
    
    LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
    If LONCODI% < 16 Or LONCODI% > 24 Then LONCODI% = 16

    CI1% = PRICOD%
    CI2% = ULTCOD%
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0)
    '
10  OBX$ = OBJPLA$("DESHASCODEP", VDEF$)
    If Len(OBX$) < 5 Then
        Exit Sub
    End If
    COD1$ = CODEXT$(CVI(Left$(OBX$, 2)))
    COD2$ = CODEXT$(CVI(Mid$(OBX$, 3, 2)))
    If TRIM$(COD1$) = "" Or TRIM$(COD2$) = "" Or COD2$ < COD1$ Then
        GoTo 10
    End If
    SUCUI% = Asc(Mid$(OBX$, 5, 1))
    If ECOARCH("ECODEP", Chr$(SUCUI%)) = "" Then
        GoTo 10
    End If
    '
    Call COPYSTRU("INVEDEP", "INVEDEQ")
    Call BLANARCH("INVEDEQ")
    If InStr(1, EMPREAC$, "DOSIVAC") > 0 Then Call BLANARCH("INVEDED")
    If InStr(1, EMPREAC$, "MEDITEA") > 0 Then Call BLANARCH("INVEDEUC") 'AGREGA FECHA DE ULTIMA COMPRA
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("INVEDEP")
    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
    '
    For i& = 1 To FIN&
        Call TRACE(20, i&, FIN&)
        X$ = REGLEIDO$("INVEDEP", i&)
        If Asc(Mid$(X$, 3, 1)) = SUCUI% Then
           CI% = CVI(Left$(X$, 2))
           If CI% > 0 Then
              If CI% <= NUMAS% Then
                 COD$ = CODEXT$(CI%)
                 If COD$ >= COD1$ Then
                    If COD$ <= COD2$ Then
                       Call REGAPP("INVEDEQ", X$)
                       'SI ES OPCION CON FECHA DE ULTIMA COMPRA
                       If InStr(1, EMPREAC$, "MEDITEA") > 0 Then
                          CONDI$ = "Cod_Inte = " & CI% & " ORDER BY Femi_Ocom DESC"
                          FECHULTCOMPRA% = CVINT(VALOBADA("OCOMDETA", "Femi_Ocom", CONDI$, "NOMESS"))
                          Call REPLA(X$, MKI$(FECHULTCOMPRA%), 21, 2)
                          Call REGAPP("INVEDEUC", X$)
                       End If
                        '
                       ' GENERA REPORTE PARA DOSIVAC, AGREGANDO AL INVEDEQ LAS COLUMNAS DE FACTURADO Y EL PRODUCTO DE ESTE CON EL VALOR UNITARIO
                       If InStr(1, EMPREAC$, "DOSIVAC") > 0 Then
                            Y$ = REGBLAN$("INVEDED")
                            
                            cantidadsinfacturar# = 0
                            SQLX$ = "select nupedcli, sum(cantingre) as caingre, sum(cantsalid) as casalid from movisto with(nolock) where codimovi='pp' and cod_inte=" & Str$(CI%) & " and depomovi=" & Str$(SUCUI%) & " and fechmov>='2013-01-01' group by nupedcli"
                            Set RST = READSET(SQLX$)
                            With RST
                                Do While Not .EOF
                                    cantidadsinfacturar# = cantidadsinfacturar# + XVALO(!CAINGRE) - XVALO(!casalid)
                                    If cantidadsinfacturar# > 0.05 Then
                                        NPED& = XVALO(!NUPEDCLI)
                                        CONDI$ = "codimovi='rt' and cod_inte=" & Str$(CI%) & " and depomovi=" & SUCUI% & " and nupedcli=" & Str$(NPED&) & " and (doseclar<>'' and doseclar is not null)"
                                        cantidadremitida# = XVALO(VALOBADA("movisto", "sum(cantsalid)", CONDI$))
                                        ' se agregan las siguientes lineas por que se encontraron movimientos sin remito y solo con despacho
                                        If cantidadremitida# < 0.05 Then
                                            CONDI$ = "codimovi='dm' and cod_inte=" & Str$(CI%) & " and depomovi=" & SUCUI% & " and nupedcli=" & NPED&
                                            cantidadremitida# = XVALO(VALOBADA("movisto", "sum(cantsalid)", CONDI$))
                                        End If
                                        If cantidadremitida# > 0.05 Then
                                            cantidadsinfacturar# = cantidadsinfacturar# - cantidadremitida#
                                        End If
                                    End If
                                    .MoveNext
                                Loop
                            End With
                            On Error Resume Next
                            RST.Close
                            Set RST = Nothing
                            On Error GoTo 0

                            SALDO# = 0
                            If cantidadsinfacturar# > 0.05 Then
                                SALDO# = cantidadsinfacturar#
                            End If
                            
                            Call REPLA(Y$, X$, 1, 20)
                            COSTO# = CVS(Mid$(X$, 9, 4))
                            ImpoTot# = SALDO# * COSTO#
                            Call REPLA(Y$, MKS$(SALDO#), 21, 4)
                            Call REPLA(Y$, MKD$(ImpoTot#), 25, 8)
                            Call REGAPP("INVEDED", Y$)
                        End If
                    End If
                 End If
              End If
           End If
        End If
    Next i&
    '
    Call CIERRARCH("INVEDEQ")
    If OPXX% = 2 Then Call CIERRARCH("INVEDEUC")
    Screen.MousePointer = 1
    '
'    If FORMINVEN.Option6.Value = True Then
'         Call FINAL("*LIINDEP1")
'       ElseIf FORMINVEN.Option5.Value = True Then
'         Call FINAL("*LIINDEP2")
'    End If
    '
    '\\\OT-08-00779-OD-04/09/08///
    Call HEADERS("INVEDEQ", "")
    If InStr(1, EMPREAC$, "MEDITEA") > 0 Then Call HEADERS("INVEDEUC", "")
    '\\\OT-08-00779-OD-04/09/08///
    '
    '\\\OT-05-0819-WAR-18/11/05///
    Call HEADERS("INVEDEQ", "Depósito: " + ECOARCH("ECODEP", Chr$(SUCUI%)))
    If InStr(1, EMPREAC$, "MEDITEA") > 0 Then Call HEADERS("INVEDEUC", "Depósito: " + ECOARCH("ECODEP", Chr$(SUCUI%)))
    If FiNVEN07.Option6.Value = True Then
         Call FINAL("*LIINDEP1")
       ElseIf FiNVEN07.Option5.Value = True Then
         '\\\OT-07-0018-RG-23/01/07///
         MON$ = Control$("INVENTAR", "MONEDINV")
         MONDESC$ = ECOARCH$("TAMONED", Chr(Val(MON$)))
         'Call HEADERS("INVEDEQ", "")
         If InStr(1, EMPREAC$, "MEDITEA") > 0 Then
            Call HEADERS("INVEDEUC", "Valorización: " + TRIM$(MON$) + " - " + TRIM$(MONDESC$))
            Call HEADERS("INVEDEUC", "Condicion: Solo Articulos Valorizados Con la Moneda Elegida")
         End If
         If InStr(1, EMPREAC$, "DOSIVAC") > 0 Then
            Call HEADERS("INVEDED", "Valorización: " + TRIM$(MON$) + " - " + TRIM$(MONDESC$))
            Call HEADERS("INVEDED", "Condicion: Solo Articulos Valorizados Con la Moneda Elegida")
         End If
         '\\\OT-07-0018-RG-FIN///
         '
         If InStr(1, EMPREAC$, "DOSIVAC") > 0 Then
            Call FINAL("*LIINDEP3")
         ElseIf InStr(1, EMPREAC$, "MEDITEA") > 0 Then
            Call FINAL("*LIINDEP4") ' CON FECHA DE ULTIMA COMPRA
         Else
            Call FINAL("*LIINDEP2")
            
         End If
    End If
    '\\\OT-05-0819-WAR-FIN///
    'CODIGO VIEJO
'    If FiNVEN04.Option6.Value = True Then
'         Call FINAL("*LIINANA1")
'       ElseIf FiNVEN04.Option5.Value = True Then
'         Call FINAL("*LIINANA2")
'    End If
    '
End Sub
'

Sub GENINV5()
    '
    CI1% = PRICOD%
    CI2% = ULTCOD%
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0)
    '
10  OBX$ = OBJPLA$("DESHASCOD1", VDEF$)
    If Len(OBX$) < 5 Then
        Exit Sub
    End If
    COD1$ = CODEXT$(CVI(Left$(OBX$, 2)))
    COD2$ = CODEXT$(CVI(Mid$(OBX$, 3, 2)))
    TMT% = Asc(Mid$(OBX$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(OBX$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    '
    If TRIM$(COD1$) = "" Or TRIM$(COD2$) = "" Or COD2$ < COD1$ Then
        GoTo 10
    End If
    '
    Call COPYSTRU("INVANAL", "INVANALF")
    Call BLANARCH("INVANALF")
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("INVANAL")
    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
    '
    For i& = 1 To FIN&
        Call TRACE(20, i&, FIN&)
        X$ = REGLEIDO$("INVANAL", i&)
        CI% = CVI(Left$(X$, 2))
        If CI% > 0 Then
           If CI% <= NUMAS% Then
              COD$ = CODEXT$(CI%)
              If COD$ >= COD1$ Then
                 If COD$ <= COD2$ Then
                    If TMT% = 0 Or TIMATE%(CI%) = TMT% Then
                        Call REGAPP("INVANALF", X$)
                    End If
                 End If
              End If
           End If
        End If
    Next i&
    '
    Call CIERRARCH("INVANALF")
    Screen.MousePointer = 1
    '/////////******FORMINVEN no existe en este proyecto****/////
'     If FORMINVEN.Option6.Value = True Then
'         Call FINAL("*LIINANA1")
'       ElseIf FORMINVEN.Option5.Value = True Then
'         Call FINAL("*LIINANA2")
'    End If
    '
    If FiNVEN07.Option6.Value = True Then
         Call FINAL("*LIINANA1")
       ElseIf FiNVEN07.Option5.Value = True Then
         '\\\OT-07-0018-RG-23/01/07///
         MON$ = Control$("INVENTAR", "MONEDINV")
         MONDESC$ = ECOARCH$("TAMONED", Chr(Val(MON$)))
         'Call HEADERS("INVANALF", "")
         Call HEADERS("INVANALF", "Valorización: " + TRIM$(MON$) + " - " + TRIM$(MONDESC$))
         Call HEADERS("INVANALF", "Condicion: Solo Articulos Valorizados Con la Moneda Elegida")
         '\\\OT-07-0018-RG-FIN///
         '
         Call FINAL("*LIINANA2")
    End If
    '
End Sub

