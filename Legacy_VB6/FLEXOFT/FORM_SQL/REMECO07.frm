VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form REMECO07 
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Gestión - Resúmenes Mensuales de Compras"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   60
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   525
      ScaleHeight     =   585
      ScaleWidth      =   2370
      TabIndex        =   48
      Top             =   6300
      Width           =   2430
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
         Picture         =   "REMECO07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle por Com- pras de Contado"
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
         TabIndex        =   50
         Top             =   105
         Width           =   1800
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C7D9CC&
      Caption         =   "LIBROS DE COMPRA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   210
      TabIndex        =   41
      Top             =   3720
      Width           =   4635
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   840
         ScaleHeight     =   615
         ScaleWidth      =   2970
         TabIndex        =   42
         Top             =   525
         Width           =   3000
         Begin VB.CommandButton Command16 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   0
            Picture         =   "REMECO07.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Emite Subdiario de Compras por Impresora"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command21 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   1840
            Picture         =   "REMECO07.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Emite Subdiario de I.V.A.- Compras por Impresora"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command18 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   640
            Left            =   0
            Picture         =   "REMECO07.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Emite Subdiario de Compras por Impresora"
            Top             =   640
            Width           =   650
         End
         Begin VB.CommandButton Command13 
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
            Left            =   2420
            Picture         =   "REMECO07.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Resumen Mensual I.V.A. Compras"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Emisión de Subdiarios"
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
            Left            =   705
            TabIndex        =   44
            Top             =   105
            Width           =   990
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C7D9CC&
      Caption         =   "GESTION DE COMPRAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   210
      TabIndex        =   31
      Top             =   1200
      Width           =   4635
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1470
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   32
         Top             =   1365
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
            Height          =   630
            Left            =   0
            Picture         =   "REMECO07.frx":0D72
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Facturación de Compras"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   700
            TabIndex        =   34
            Top             =   105
            Width           =   1200
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1150
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   35
         Top             =   945
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
            Height          =   630
            Left            =   -20
            Picture         =   "REMECO07.frx":107C
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Notas de Débito"
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
            Left            =   700
            TabIndex        =   37
            Top             =   105
            Width           =   1410
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   840
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   38
         Top             =   525
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
            Height          =   630
            Left            =   0
            Picture         =   "REMECO07.frx":1386
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Notas de Crédito"
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
            TabIndex        =   40
            Top             =   105
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C7D9CC&
      Caption         =   "PERÍODO MENSUAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   210
      TabIndex        =   27
      Top             =   105
      Width           =   4635
      Begin VB.PictureBox Picture11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   3500
         ScaleHeight     =   510
         ScaleWidth      =   975
         TabIndex        =   62
         Top             =   240
         Width           =   1000
         Begin VB.Image Image2 
            Height          =   480
            Left            =   460
            Picture         =   "REMECO07.frx":1690
            Top             =   20
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   0
            Picture         =   "REMECO07.frx":199A
            Top             =   20
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.CommandButton Command12 
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
         Height          =   285
         Left            =   3200
         TabIndex        =   29
         ToolTipText     =   "Click Para Seleccionar Período Mensual"
         Top             =   410
         Width           =   175
      End
      Begin VB.TextBox MESAN 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   28
         ToolTipText     =   "Doble-Click Para Rango de Fechas"
         Top             =   390
         Width           =   2115
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mes y año"
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
         Index           =   10
         Left            =   -1200
         TabIndex        =   30
         Top             =   450
         Width           =   2220
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H0050765A&
      ForeColor       =   &H80000008&
      Height          =   7000
      Left            =   9660
      ScaleHeight     =   6975
      ScaleWidth      =   1980
      TabIndex        =   21
      Top             =   -105
      Width           =   2010
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H000C93DC&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   64
         Top             =   5880
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0076C9F5&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   65
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "REMECO07.frx":1CA4
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   3950
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command11 
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
         Picture         =   "REMECO07.frx":1FAE
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Archivos de Transferencia Ingresos Brutos / C.I.T.I. Compras"
         Top             =   2570
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "AcDir"
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
         Picture         =   "REMECO07.frx":22B8
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3240
         Visible         =   0   'False
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
         Picture         =   "REMECO07.frx":25C2
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   990
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Setup"
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
         Picture         =   "REMECO07.frx":28CC
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   4560
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command15 
         Caption         =   "A-B-M"
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
         Picture         =   "REMECO07.frx":2D0E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Alrchivo Maestro de Proveedores"
         Top             =   1880
         Visible         =   0   'False
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
         Picture         =   "REMECO07.frx":3018
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   300
         Width           =   650
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C7D9CC&
      Caption         =   "DISPOSITIVO DE SALIDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   210
      TabIndex        =   8
      Top             =   5250
      Width           =   4635
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C7D9CC&
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
         Left            =   2520
         TabIndex        =   10
         Top             =   420
         Width           =   1170
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C7D9CC&
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
         Left            =   945
         TabIndex        =   9
         Top             =   420
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C7D9CC&
      Caption         =   "OTROS REPORTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6030
      Left            =   5040
      TabIndex        =   0
      Top             =   105
      Width           =   4425
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3960
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C7D9CC&
         Caption         =   "ESTADÍSTICAS DE COMPRAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1225
         Left            =   210
         TabIndex        =   52
         Top             =   4620
         Width           =   4005
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   420
            ScaleHeight     =   615
            ScaleWidth      =   3165
            TabIndex        =   53
            Top             =   420
            Width           =   3200
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
               Height          =   630
               Left            =   0
               Picture         =   "REMECO07.frx":3162
               Style           =   1  'Graphical
               TabIndex        =   56
               ToolTipText     =   "Listado Estadístico de Facturación de Compras"
               Top             =   0
               Width           =   560
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
               Left            =   2620
               Picture         =   "REMECO07.frx":346C
               Style           =   1  'Graphical
               TabIndex        =   55
               ToolTipText     =   "Ranking de Compras por Articulo por Rango de Fechas"
               Top             =   0
               Width           =   560
            End
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
               Height          =   630
               Left            =   2050
               Picture         =   "REMECO07.frx":3776
               Style           =   1  'Graphical
               TabIndex        =   54
               ToolTipText     =   "Ranking de Compras por Proveedor por Rango de Fechas"
               Top             =   0
               Width           =   555
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Resúmenes Estadísticos"
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
               Left            =   700
               TabIndex        =   57
               Top             =   90
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C7D9CC&
         Caption         =   "RESUMEN OPERATIVO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1225
         Left            =   210
         TabIndex        =   16
         Top             =   2020
         Width           =   4005
         Begin VB.ListBox List1 
            Height          =   840
            Left            =   0
            Sorted          =   -1  'True
            TabIndex        =   51
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   945
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   17
            Top             =   350
            Width           =   2325
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
               Height          =   630
               Left            =   0
               Picture         =   "REMECO07.frx":3BB8
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Resumen Mensual           por Cuenta"
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
               Left            =   700
               TabIndex        =   19
               Top             =   105
               Width           =   1800
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C7D9CC&
         Caption         =   "PROCESOS CONTABLES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   11
         Top             =   3320
         Width           =   4005
         Begin VB.PictureBox Picture15 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   420
            ScaleHeight     =   615
            ScaleWidth      =   3180
            TabIndex        =   12
            Top             =   350
            Width           =   3210
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
               Height          =   630
               Left            =   2620
               Picture         =   "REMECO07.frx":3FFA
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "Re-Imputación de Comprobantes de Compras"
               Top             =   0
               Width           =   560
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
               Height          =   630
               Left            =   0
               Picture         =   "REMECO07.frx":4144
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Consulta / Listado por Cuenta Contable"
               Top             =   0
               Width           =   560
            End
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
               Height          =   630
               Left            =   2050
               Picture         =   "REMECO07.frx":444E
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Consulta de Asiento Contable por Pantalla"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Contabilización de Compras"
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
               TabIndex        =   15
               Top             =   105
               Width           =   1365
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C7D9CC&
         Caption         =   "ORDENES DE PAGO Y AJUSTES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   210
         TabIndex        =   1
         Top             =   315
         Width           =   4005
         Begin VB.CommandButton Command23 
            Caption         =   "Command23"
            Height          =   255
            Left            =   3240
            TabIndex        =   63
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   840
            ScaleHeight     =   615
            ScaleWidth      =   2400
            TabIndex        =   2
            Top             =   840
            Width           =   2430
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
               Height          =   630
               Left            =   -20
               Picture         =   "REMECO07.frx":4758
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Detalle por Ordenes  de Pago"
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
               Left            =   700
               TabIndex        =   4
               Top             =   105
               Width           =   1695
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   525
            ScaleHeight     =   615
            ScaleWidth      =   2400
            TabIndex        =   5
            Top             =   420
            Width           =   2430
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
               Height          =   630
               Left            =   0
               Picture         =   "REMECO07.frx":4A62
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Ajustes y Otros Movimientos"
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
               Index           =   0
               Left            =   700
               TabIndex        =   7
               Top             =   0
               Width           =   1590
            End
         End
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "REMECO07.frx":4D6C
      TabIndex        =   61
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "REMECO07.frx":4DE6
      Top             =   1200
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu MnuCerrarPeriodoActivo 
         Caption         =   "Cerrar Período Activo"
      End
      Begin VB.Menu nnuAbrirPeriodoActivo 
         Caption         =   "Abrir Período Activo"
      End
   End
   Begin VB.Menu mnuConsultasylistados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuCuentaCorriente 
         Caption         =   "Resumen de Cuenta Corriente"
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEstadoDeCuenta 
         Caption         =   "Estado de Cuenta"
         Begin VB.Menu mnuCtasxPagar 
            Caption         =   "Cuentas por Pagar"
         End
         Begin VB.Menu mnuSaldosydeudasproveedores 
            Caption         =   "Saldos y Deudas de Proveedores"
         End
      End
      Begin VB.Menu mnuGestionCompras 
         Caption         =   "Gestión de Compras"
         Begin VB.Menu mnuFacturaCompra 
            Caption         =   "Facturas de Compras"
         End
         Begin VB.Menu mnuNotaDebito 
            Caption         =   "Notas de Débito de Compras"
         End
         Begin VB.Menu mnuNotaCredito 
            Caption         =   "Notas de Crédito de Compras"
         End
      End
      Begin VB.Menu mnuLibrosCompra 
         Caption         =   "Libros de Compra"
         Begin VB.Menu mnuSubdiarioCompras 
            Caption         =   "Subdiario de Compras por Impresora"
         End
         Begin VB.Menu mnuSubdirioivacompras 
            Caption         =   "Subdiario de I.V.A.- Compras por Impresora"
         End
         Begin VB.Menu mnuResumenMensualivacomp 
            Caption         =   "Resumen Mensual I.V.A. Compras"
         End
      End
      Begin VB.Menu mnuOtros 
         Caption         =   "Otros"
         Begin VB.Menu mnuDetalleOrdenPago 
            Caption         =   "Detalle de Ordenes de Pago"
         End
         Begin VB.Menu mnuAjustesOtrosMovimientos 
            Caption         =   "Ajustes y Otros Movimientos"
         End
         Begin VB.Menu mnuResumenMensualCta 
            Caption         =   "Resumen Mensual por Cuenta"
         End
      End
      Begin VB.Menu mnuProcesosContables 
         Caption         =   "Procesos Contables"
         Begin VB.Menu mnulistadoCtaCble 
            Caption         =   "Listado por Cuenta Contable"
         End
         Begin VB.Menu mnuconsultaAsientocontable 
            Caption         =   "Consulta de Asiento Contable por Pantalla"
         End
      End
      Begin VB.Menu mnuEstadisticas 
         Caption         =   "Estadisticas"
         Begin VB.Menu mnuEstadisticoCompras 
            Caption         =   "Listado Estadístico de Facturación de Compras"
         End
         Begin VB.Menu RankingComprasProvxFecha 
            Caption         =   "Ranking de Compras por Proveedor por Rango de Fechas"
         End
         Begin VB.Menu mnuRankingArticuloxFecha 
            Caption         =   "Ranking de Compras por Articulo por Rango de Fechas"
         End
      End
      Begin VB.Menu mnuGastosPorProvincia 
         Caption         =   "Gastos por Provincia"
      End
      Begin VB.Menu mnuReportesVarios 
         Caption         =   "Reportes Varios"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuReimputacionComprobanteCompras 
         Caption         =   "Re-Imputación de Comprobantes de Compras"
      End
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu mnuAsociarReporte 
         Caption         =   "Vincular Reportes al Formulario"
      End
      Begin VB.Menu MnuConfigFormIva 
         Caption         =   "Configuración de Libro de IVA"
      End
   End
   Begin VB.Menu mnuAccesoDirectos 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuMaestroProved 
         Caption         =   "Maestro de Proveedores"
      End
   End
   Begin VB.Menu Impuestos 
      Caption         =   "Impuestos"
      Begin VB.Menu IIBB 
         Caption         =   "Ingresos Brutos"
         Begin VB.Menu IIBB_BuenosAires 
            Caption         =   "Buenos Aires"
            Begin VB.Menu IIBB_BsAsPercepciones 
               Caption         =   "Percepciones"
               Begin VB.Menu IIBB_BsAsListaPer 
                  Caption         =   "Listado de Percepciones"
               End
               Begin VB.Menu IIBB_BsAsExpoPer 
                  Caption         =   "Archivo de Transferencia"
               End
            End
            Begin VB.Menu IIBB_BsAsRetenciones 
               Caption         =   "Retenciones (Propias)"
               Begin VB.Menu IIBB_BsAsListaRet 
                  Caption         =   "Listado de Retenciones"
               End
               Begin VB.Menu IIBB_BsAsExpoRet 
                  Caption         =   "Archivo Transferencia"
               End
            End
         End
         Begin VB.Menu IIBB_CABA 
            Caption         =   "CABA"
            Begin VB.Menu IIBBPercepRetenCABAExpo 
               Caption         =   "Percepciones y Retenciones - Archivo de Transferencia"
            End
         End
         Begin VB.Menu IIBB_ConvenioMulti 
            Caption         =   "Convenio Multilateral"
            Begin VB.Menu IIBB_CMPercepciones 
               Caption         =   "Percepciones"
               Begin VB.Menu IIBB_CMListaPer 
                  Caption         =   "Listado de Percepciones "
                  Index           =   0
               End
               Begin VB.Menu IIBB_CMExpoPer 
                  Caption         =   "Archivo de Transferencia"
                  Index           =   1
               End
            End
            Begin VB.Menu IIBB_CMRecaudaciones 
               Caption         =   "Recaudaciones Bancarias (SIRCREB)"
               Begin VB.Menu IIBB_CMSircrebListaRet 
                  Caption         =   "Listado de Retenciones"
               End
               Begin VB.Menu IIBB_CMSircrebCoeficientes 
                  Caption         =   "Coeficientes de Distribución"
               End
               Begin VB.Menu IIBB_CMSircrebExpoRet 
                  Caption         =   "Archivo de Transferencia"
               End
            End
         End
      End
      Begin VB.Menu Iva 
         Caption         =   "I.V.A."
         Begin VB.Menu IvaPercepciones 
            Caption         =   "Percepciones"
            Begin VB.Menu IvaExpoPer 
               Caption         =   "Archivo de Transferencia"
            End
         End
         Begin VB.Menu DespachosCourier 
            Caption         =   "Despachos y Courier"
            Begin VB.Menu DespachosCourierTransf 
               Caption         =   "Archivo de Transferencia"
            End
         End
      End
      Begin VB.Menu Ganancias 
         Caption         =   "Ganancias"
         Begin VB.Menu GananciasRetenciones 
            Caption         =   "Retenciones (Propias)"
            Begin VB.Menu GananciasListaRet 
               Caption         =   "Listado de Retenciones"
            End
            Begin VB.Menu GananciasExpoRet 
               Caption         =   "Archivo de Transferencia"
            End
         End
      End
      Begin VB.Menu RegimenCompras 
         Caption         =   "Regimen Informativo de Compras"
      End
      Begin VB.Menu mnuLibroIvaDigital 
         Caption         =   "Libro IVA Digital"
         Begin VB.Menu mnuLibroIvaDigitaGenerarArchivo 
            Caption         =   "Generar Archivo de Transferencia"
         End
         Begin VB.Menu mnuLibroIvaDigitalCompararMisComprobantes 
            Caption         =   "Comparar con Portal IVA"
         End
      End
   End
End
Attribute VB_Name = "REMECO07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Desde%, Hasta%, CORRESP$
'jandy SQL
'Dim DESDAT As Double
'Dim HASDAT As Double
Dim DESDAT As String
Dim HASDAT As String


Private Sub CargaCoeficientesSIRCREB()

    ' *** CREA / ACTUALIZA TABLA DE JURISDICCIONES
    For FILA& = 1 To ULTREG&("PROVINB")
        Registro0$ = REGLEIDO$("PROVINB", FILA&)
        CodigoJur% = CVI(Mid$(Registro0$, 37, 2))
        Provincia$ = Mid$(Registro0$, 2, 25)
        
        Call REPLA(Registro1$, MKI$(CodigoJur%), 1, 2)
        Call REPLA(Registro1$, Provincia$, 3, 25)
        
        Call GRAREG("JURIDIB", Registro1$, FILA&)
        
    Next FILA&
    
    
    Call SELMESA(Des%, Has%, PERIODO$)
    InicioPeriodo$ = FECHATEXLAR$(Des%)
    PERIODO$ = Mid$(InicioPeriodo$, 7, 4) & Mid$(InicioPeriodo$, 4, 2)
    ConsultaSql$ = "SELECT * FROM SIRCREB WITH(NOLOCK) WHERE Periodo=" & PERIODO$
    
    Dim Rst_SIRCREB As ADODB.Recordset
    Set Rst_SIRCREB = READSET(ConsultaSql$)
    
    With Rst_SIRCREB
    
        NroRegistro& = 0
        Do While .EOF = False
            
            CodigoJurisdiccion% = !CodigoJurisdiccion
            Coeficiente# = !Coeficiente
        
            RegistroD$ = REGBLAN("COEFSIRC")
            Call REPLA(RegistroD$, MKI$(CodigoJurisdiccion%), 1, 2)
            Call REPLA(RegistroD$, MKD$(Coeficiente#), 3, 8)
            
            NroRegistros& = NroRegistros& + 1
            Call GRAREG("!COEFSIRC", RegistroD$, NroRegistros&)
            
            .MoveNext
        Loop
    
    End With
    
    Rst_SIRCREB.Close
    Set Rst_SIRCREB = Nothing
    Call SETULTREG("!COEFSIRC", NroRegistros&)
    Call CIERRARCH("*.*")
    
    
    ControlOk% = VENTABU%("COEFSIRC")
    If ControlOk% = 0 Then
        CondicionSql$ = "PERIODO=" & PERIODO$
        Call ACTUREGISTRO("SIRCREB", "MARCABORRADO", CondicionSql$, 1, REGAF&)
        
        Dim RstUp_SIRCREB As ADODB.Recordset
        Set RstUp_SIRCREB = New ADODB.Recordset
        CondicionSql$ = "SELECT * FROM SIRCREB WHERE PERIODO=" & PERIODO$
        RstUp_SIRCREB.Open FILTSQL$(CondicionSql$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RstUp_SIRCREB
             For FILA& = 1 To ULTREG&("!COEFSIRC")
                .AddNew
                
                RegistroD$ = REGLEIDO$("!COEFSIRC", FILA&)
                CodigoJurisdiccion% = CVI(Mid$(RegistroD$, 1, 2))
                RegistroProvincia$ = FILTERGETRECORD("PROVINB", 9, MKI$(CodigoJurisdiccion%))
                !PERIODO = CLng(PERIODO$)
                !CodigoJurisdiccion = CodigoJurisdiccion%
                !CodigoProvincia = Mid$(RegistroProvincia$, 1, 1)
                !Provincia = Mid$(RegistroProvincia$, 2, 25)
                !CodigoContable = CVI(Mid$(RegistroProvincia$, 29, 2))
                !CuentaContable = CODCUE$(!CodigoContable)
                !Coeficiente = XVALO(CVD(Mid$(RegistroD$, 3, 8)))
                !FechaUltimaModificacion = Now
                !MarcaBorrado = 0
                
                .Update
             Next FILA&
        End With
        RstUp_SIRCREB.Close
        Set RstUp_SIRCREB = Nothing
        
        CondicionSql$ = "PERIODO=" & PERIODO$ & " AND MARCABORRADO=1"
        Call BORRAREGISTROS("SIRCREB", CondicionSql$, REGAF&)
        
        Call COMUNI("Coeficientes Actualizados Correctamente.")
        
     End If
    
End Sub


Sub GASPORCUECON()
   '
   Screen.MousePointer = 11
   '
   Call CREACAMPO("", "GASCUECON", "CueContaI", 3, 0, 1)
   Call CREACAMPO("", "GASCUECON", "NuProve", 3, 0)
   Call CREACAMPO("", "GASCUECON", "CodiCom_Lar", 10, 24, 1)
   Call CREACAMPO("", "GASCUECON", "FechEmisi", 8, 0, 1)
   Call CREACAMPO("", "GASCUECON", "Importe", 7, 0)
   Call CREACAMPO("", "GASCUECON", "CODEXCUE", 10, 16)
   '
   
   'BORRO LOS DATOS DE LA TABLA
   CONDI$ = "CUECONTAI < 32769 OR CUECONTAI IS NULL"
   Call BORRAREGISTROS("GASCUECON", CONDI$, REGAF&, "NOMESS")
   '
   FIN& = ULTREG&("ANAICOM")
   For i& = 1 To FIN&
       'LEO EL DAT
       Call TRACE(20, i&, FIN&)
       XX$ = REGLEIDO$("ANAICOM", i&)
       CUECONTABLE% = CVI(Mid$(XX$, 1, 2))
       CODIGOEXTERNOCTA$ = CODCUE$(CUECONTABLE%)
       NUPROV% = CVI(Mid$(XX$, 3, 2))
       DOCUMEN$ = TRIM$(Mid$(XX$, 5, 18))
       FE% = CVI(Mid$(XX$, 23, 2))
       fech = FETEXCOR1$(FE%)
       IMPORTE# = CVD(Mid$(XX$, 25, 8))
       '
       'GRABO EN TABLA SQL
10     RNDX$ = RANDSTRING$(-8)
       CONDI1$ = "Codicom_lar ='" & RNDX$ & "'"
       If CantRegistros("GASCUECON", CONDI1$, "NOMESS") > 0 Then GoTo 10
       Call CreaRegistro("GASCUECON", "Codicom_lar", RNDX$, "NOMESS")
       Call ACTUREGISTRO("GASCUECON", "CUECONTAI", CONDI1$, CUECONTABLE%, RAFE&, "NOMESS")
       Call ACTUREGISTRO("GASCUECON", "NUPROVE", CONDI1$, NUPROV%, RAFE&, "NOMESS")
       Call ACTUREGISTRO("GASCUECON", "FECHEMISI", CONDI1$, fech, RAFE&, "NOMESS")
       Call ACTUREGISTRO("GASCUECON", "IMPORTE", CONDI1$, FORMATNUM$(IMPORTE#, "F16.5"), RAFE&, "NOMESS")
       Call ACTUREGISTRO("GASCUECON", "CODEXCUE", CONDI1$, CODIGOEXTERNOCTA$, RAFE&, "NOMESS")
       '
       Call ACTUREGISTRO("GASCUECON", "Codicom_lar", CONDI1$, DOCUMEN$, RAFE&, "NOMESS")
       '
   Next i&
   '
   'VALIDO SI HAY DATOS
   If i& < 1 Then
      Screen.MousePointer = 1
      Call COMUNI("Archivo de Trabajo Vacio")
      Exit Sub
   End If

   'OPCIONES DE LISTADOS
15  OPX% = ALTERNA%("Gastos Mensuales Acumulados\Gastos Mensuales Detallados\Gastos Por Proveedor")

   If OPX% < 1 Or OPX% > 3 Then Exit Sub
   '
   Select Case OPX%
     Case 1
        Call FINAL("?GASCUMCUECO")
     Case 2
       GACUECON.Show 1
     Case 3
        Call FINAL("?GAPROCUECON")
   End Select

   Screen.MousePointer = 1
   GoTo 15
End Sub


Private Sub ExportaRecaudacionesBancarias()

    CuentaSircreb& = XVALO(CONTROL$("FACOMPRA", "SIRCREB"))
    If CuentaSircreb& = 0 Then
        Call MENSERR(24, "Imposible Continuar.\Falta Cuenta Contable Correspondiente\a Recaudaciones Bancarias (SIRCREB)\ \Configurar en Setup de Facturación de Proveedores.")
        Exit Sub
    End If

    Campos$ = "Jurisdiccion/A12;Cuit/A13;CBU/A24;Tipo Cta/A7;Moneda/A6;Importe/F12.2;Observaciones/A48"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , ANTOT)
    FRMZELECHO.Width = 12000
    FRMZELECHO.Caption = "Informe de Errores"
            
    ' *** CREA EL ARCHIVO DE RETENCIONES
    RutaDelArchivo$ = SELECCION_Archivo$("TXT", CommonDialog1)
    LongitudRegistro% = 58
    If TRIM$(RutaDelArchivo$) = "" Then Exit Sub
    NumeroArchivo% = FILEOPEN%(RutaDelArchivo$, 2, LongitudRegistro%) ' ABRE EL ARCHIVO DE ESCRITURA

    Call SELMESA(Des%, Has%, CORRESP$)
    If Des% < 1 Then Exit Sub
    InicioPeriodo$ = FECHATEXLAR$(Des%)
    PeriodoRetenido$ = Mid$(InicioPeriodo$, 7, 4) & Mid$(InicioPeriodo$, 4, 2)
    DESDE0$ = FETEXCOR1$(Des%)
    HASTA0$ = FETEXCOR1$(Has%)

    Screen.MousePointer = 11
    Dim CCBanco(1024) As Integer
    Dim ImporteRetenidoPorBanco(1024) As Single
    Dim IndiceBanco As Integer
    JX& = 0
    
    SQLX$ = "SELECT sum(IDEBECOMP) AS IDEBE, SUM(IHABECOMP) AS IHABE, NUPROVE FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE FECHCONTA >= '" & DESDE0$ & "'"
    SQLX$ = SQLX$ + " AND FECHCONTA <= '" & HASTA0$ & "'"
    SQLX$ = SQLX$ + " AND NUPROVE > 0"
    SQLX$ = SQLX$ + " AND CUECONTAI=" & CUEINT%(CStr(CuentaSircreb&))
    SQLX$ = SQLX$ + " GROUP BY NUPROVE"
    Dim RstSIRCREB As ADODB.Recordset
    Set RstSIRCREB = READSET(SQLX$)

    With RstSIRCREB
            Do While Not .EOF
                RetencionFallida% = 0
                TxtCuentaContableInvalida$ = ""
                TxtCuitInvalidoe$ = ""
                
                NumeroProveedor% = XVALO(!NUPROVE)
                CuitBanco$ = CuitCli$((-1) * NumeroProveedor%)
                CuentaContableBanco% = CUEINT%(CUEMADRE$((-1) * NumeroProveedor%))
                ImpoRetenido# = XVALO(!IDEBE) - XVALO(!IHABE)

                ' VERIFICA QUE LA CUENTA CONTABLE SEA VALIDA
                If CuentaContableBanco% = 0 Then
                    RetencionFallida% = 1: TxtCuentaContableInvalida$ = "Falta Cuenta Contable en Proveedor (" & NumeroProveedor% & ")"
                End If
                
                If VALICUIT%(CuitBanco$) = 0 Then
                    RetencionFallida% = 1: TxtCuitInvalidoe$ = "Cuit Invalido para el Proveedor (" & NumeroProveedor% & ")"
                End If
                
                If RetencionFallida% = 1 Then
                    FalloGeneración% = 1
                    If TxtCuentaContableInvalida$ <> "" Then
                        JX& = JX& + 1
                        FRMZELECHO.msf2.Rows = JX& + 1
                        FRMZELECHO.msf2.TextMatrix(JX&, 1) = ""                              ' COMPROBANTE
                        FRMZELECHO.msf2.TextMatrix(JX&, 2) = ""                              ' JURISDICCION
                        FRMZELECHO.msf2.TextMatrix(JX&, 3) = CuitBanco$                                         ' CUIT
                        FRMZELECHO.msf2.TextMatrix(JX&, 4) = ""                                               ' CBU
                        FRMZELECHO.msf2.TextMatrix(JX&, 5) = ""                              ' TIPO DE CUENTA
                        FRMZELECHO.msf2.TextMatrix(JX&, 6) = ""                                     ' TIPO DE MONEDA
                        FRMZELECHO.msf2.TextMatrix(JX&, 7) = TRIM$(FORMATNUM$(ImpoRetenido#, "F10.2"))   ' IMPORTE RETENIDO
                        FRMZELECHO.msf2.TextMatrix(JX&, 8) = TxtCuentaContableInvalida$                                       ' OBSERVACIONES
                    End If
                    If TxtCuitInvalidoe$ <> "" Then
                        JX& = JX& + 1
                        FRMZELECHO.msf2.Rows = JX& + 1
                        FRMZELECHO.msf2.TextMatrix(JX&, 1) = ""                              ' COMPROBANTE
                        FRMZELECHO.msf2.TextMatrix(JX&, 2) = ""                              ' JURISDICCION
                        FRMZELECHO.msf2.TextMatrix(JX&, 3) = CuitBanco$                                         ' CUIT
                        FRMZELECHO.msf2.TextMatrix(JX&, 4) = ""                                               ' CBU
                        FRMZELECHO.msf2.TextMatrix(JX&, 5) = ""                              ' TIPO DE CUENTA
                        FRMZELECHO.msf2.TextMatrix(JX&, 6) = ""                                     ' TIPO DE MONEDA
                        FRMZELECHO.msf2.TextMatrix(JX&, 7) = TRIM$(FORMATNUM$(ImpoRetenido#, "F10.2"))   ' IMPORTE RETENIDO
                        FRMZELECHO.msf2.TextMatrix(JX&, 8) = TxtCuitInvalidoe$                                     ' OBSERVACIONES
                    End If
                Else
                    For iBanco = 1 To 1024
                        If CCBanco(iBanco) = CuentaContableBanco% Then
                            ImporteRetenidoPorBanco(iBanco) = ImporteRetenidoPorBanco(iBanco) + ImpoRetenido#
                            GoTo 85
                        End If
                    Next iBanco
                    
                    ' EL BANCO NO EXISTE EN EL VECTOR CCBANCO, POR LO TANTO LO AGREGA
                    IndiceBanco = IndiceBanco + 1
                    CCBanco(IndiceBanco) = CuentaContableBanco%
                    ImporteRetenidoPorBanco(IndiceBanco) = ImporteRetenidoPorBanco(IndiceBanco) + ImpoRetenido#
                
                    If IndiceBanco = 1024 Then
                        Call COMUNI("Se Ha Superado el Máximo de Cuentas Bancarias.\Las Retenciones Serán Exportadas Parcialmente.\ \Consulte a su Soporte de Sistemas.")
                        GoTo 98
                    End If
                End If
                
                
'                Dim RstCoeficientes_SIRCREB As ADODB.Recordset
'                ConsultaSql$ = "SELECT * FROM SIRCREB WITH(NOLOCK) WHERE PERIODO=" & PeriodoRetenido$
'                Set RstCoeficientes_SIRCREB = READSET(ConsultaSql$)
'                With RstCoeficientes_SIRCREB
'                    If (.EOF And .BOF) Then
'                        Call COMUNI("Imposible Exportar Recaudaciones.\ \Falta Definir Coeficientes de Distribución\Para el Período '" & PeriodoRetenido$ & "'")
'                        Screen.MousePointer = 1
'                        GoTo 90
'                    End If
'                    FalloEnGeneracion% = 0: RetencionFallida% = 0: TextoError$ = ""
'                    Do While Not .EOF
'                        RetencionFallida% = 0
'
'                            CodigoDeJurisdiccion$ = CStr(!CodigoJurisdiccion)
'                        CodigoDeJurisdiccion$ = Right("000" & CodigoDeJurisdiccion$, 3)
'                        CuitBanco$ = VALOBADA("PROVED12", "CUIT_CLI", "NUME_CLI=" & CStr(CuentaProveedor%))
'                            CuentaContableBanco% = XVALO(VALOBADA("PROVED12", "CUMA_CLI", "NUME_CLI=" & CStr(CuentaProveedor%)))
'                        RegistroBanco$ = FILTERGETRECORD$("CUEBANC", 1, CODCUE$(CuentaContableBanco%))
'                        CBU$ = TRIM$(Mid$(RegistroBanco$, 49, 25))
'                        TipoDeCuentaBancaria$ = Mid$(RegistroBanco$, 74, 2)
'                        TipoDeMoneda$ = TRIM$(Mid$(RegistroBanco$, 76, 2))
'                        Coeficiente# = !Coeficiente
'                        ImporteJurisdiccion# = ImporteRetenidoBanco# * Coeficiente#
'
'                        ' VALIDACION DE LOS DATOS
'                        TextoError$ = ""
'                        If CodigoDeJurisdiccion$ = "" Or XVALO(CodigoDeJurisdiccion$) = 0 Then
'                            TextoError$ = "Jurisdicción Inválida": RetencionFallida% = 1: FalloEnGeneracion% = 1
'                        End If
'                        If CuitBanco$ = "" Or VALICUIT%(CuitBanco$) = 0 Then
'                            If TextoError$ = "" Then
'                                TextoError$ = "Cuit Inválido"
'                            Else
'                                TextoError$ = TextoError$ & ", Cuit Inválido"
'                            End If
'                            RetencionFallida% = 1: FalloEnGeneracion% = 1
'                        End If
'                        If TRIM$(CBU$) = "" Or TRIM$(TipoDeCuentaBancaria$) = "" Or TRIM$(TipoDeMoneda$) = "" Then
'                            If TextoError$ = "" Then
'                                TextoError$ = "Faltan Datos Bancarios. Revisar CBU, Tipo de Cuenta y de Moneda."
'                            Else
'                                TextoError$ = TextoError$ & ", Faltan Datos Bancarios. Revisar CBU, Tipo de Cuenta y de Moneda"
'                            End If
'                            RetencionFallida% = 1: FalloEnGeneracion% = 1
'                        End If
'
'
'                        ' SI ALGUN DATO NO PASO LA VALIDACION, CARGA LA RETENCION EN FRMZELECHO PARA QUE LUEGO EL USUARIO
'                        ' PUEDA VER QUE SUCEDIO
'                        If RetencionFallida% = 1 Then
'                            JX& = JX& + 1
'                            FRMZELECHO.MSF2.Rows = JX& + 1
'                            FRMZELECHO.MSF2.TextMatrix(JX&, 1) = CodigoDeJurisdiccion$                              ' JURISDICCION
'                            FRMZELECHO.MSF2.TextMatrix(JX&, 2) = CuitBanco$                                         ' CUIT
'                            FRMZELECHO.MSF2.TextMatrix(JX&, 3) = CBU$                                               ' CBU
'                            FRMZELECHO.MSF2.TextMatrix(JX&, 4) = TipoDeCuentaBancaria$                              ' TIPO DE CUENTA
'                            FRMZELECHO.MSF2.TextMatrix(JX&, 5) = TipoDeMoneda$                                      ' TIPO DE MONEDA
'                            FRMZELECHO.MSF2.TextMatrix(JX&, 6) = TRIM$(FORMATNUM$(ImporteJurisdiccion#, "F10.2"))   ' IMPORTE RETENIDO
'                            FRMZELECHO.MSF2.TextMatrix(JX&, 7) = TextoError$                                        ' OBSERVACIONES
'                        Else
'                        ' TODAS LAS RETENCIONES SUPERARON LAS VALIDACIONES, CON LO CUAL SE GUARDA EL REGISTRO
'                            ImporteRetenido$ = TRIM$(FORMATNUM$(ImporteJurisdiccion#, "F10.2"))
'                            ImporteRetenido$ = Right("0000000000" & ImporteRetenido$, 10)
'                            ImporteRetenido$ = REPLACAR$(ImporteRetenido$, ".", ",")
'                            PERIODO$ = Mid$(PeriodoRetenido$, 1, 4) & "/" & Mid$(PeriodoRetenido$, 5, 2)
'                            RegistroRetencion$ = CodigoDeJurisdiccion$ & CuitBanco$ & PERIODO$ & CBU$ & TipoDeCuentaBancaria$ & TipoDeMoneda$ & ImporteRetenido$
'                            Print #NumeroArchivo%, RegistroRetencion$
'
'                            TotalDeRetenciones% = TotalDeRetenciones% + 1
'                        End If
'
'                        .MoveNext
'                    Loop
'            End With
'            RstCoeficientes_SIRCREB.Close
'            Set RstCoeficientes_SIRCREB = Nothing
            
            .MoveNext
85      Loop
    End With
    
90  RstSIRCREB.Close
    Set RstSIRCREB = Nothing
    
    Dim UltimoIdAsiento As String
    
    Dim RstAsientos_SIRCREB As ADODB.Recordset
    ConsultaSql$ = "SELECT * FROM DETAASIEN WITH(NOLOCK) WHERE FECHASIEN >='" & DESDE0$ & "' AND FechAsien <='" & HASTA0$ & "'"
    ConsultaSql$ = ConsultaSql$ & " AND StatPase = 1 AND CueCoInt=" & CUEINT%(CStr(CuentaSircreb&))
    Set RstAsientos_SIRCREB = READSET(ConsultaSql$)
    
    With RstAsientos_SIRCREB
        Do While Not .EOF
            RetencionFallida% = 0
            TxtContrapartidaInvalida$ = ""
            TxtCuentaContableInvalida$ = ""
            TxtCuitInvalidoe$ = ""
            
            IdAsiento$ = SAFETEXT$(!nuidasien)
            ImpoRetenido# = XVALO(!idebepase) - XVALO(!ihabepase)

            ' BUSCA LA CUENTA CONTABLE DEL BANCO EN LA CONTRAPARTIDA
            CondicionSql$ = " AND IDEBEPASE > 0.005"
            If ImpoRetenido# > 0.005 Then CondicionSql$ = " AND IHabePase > 0.005"
            If CantRegistros("DETAASIEN", "NUIDASIEN='" & IdAsiento$ & "' and StatPase = 1" & CondicionSql$) > 1 Then
                RetencionFallida% = 1: TxtContrapartidaInvalida$ = "Imposible Detectar Banco. Existe más de una Cuenta de Contrapartida."
            End If
            
            ' BUSCA EL NUMERO DE CUENTA DEL PROVEEDOR 'BANCO' A PARTIR DE CONSIDERAR LA CUENTA CONTABLE
            ' ANTERIOR COMO SU CUENTA MADRE EN EL MAESTRO
            CuentaContableBanco% = XVALO(VALOBADA("DETAASIEN", "CueCoInt", "NUIDASIEN='" & IdAsiento$ & "'" & CondicionSql$))
            NumeroProveedor% = XVALO(VALOBADA("PROVED12", "NUME_CLI", "CUMA_CLI=" & CStr(CuentaContableBanco%)))
            CuitBanco$ = CuitCli$((-1) * NumeroProveedor%)

            ' VERIFICA QUE LA CUENTA CONTABLE SEA VALIDA
            If DENOCUE$(CuentaContableBanco%) = "" Then
                RetencionFallida% = 1: TxtCuentaContableInvalida$ = "Falta Cuenta Contable en Proveedor (" & NumeroProveedor% & ")"
            End If
            
            If VALICUIT%(CuitBanco$) = 0 Then
                RetencionFallida% = 1: TxtCuitInvalidoe$ = "Cuit Invalido para el Proveedor (" & NumeroProveedor% & ")"
            End If
                
            If RetencionFallida% = 1 Then
                FalloGeneración% = 1
                If TxtContrapartidaInvalida$ <> "" Then
                    JX& = JX& + 1
                    FRMZELECHO.msf2.Rows = JX& + 1
                    FRMZELECHO.msf2.TextMatrix(JX&, 1) = IdAsiento$                              ' COMPROBANTE
                    FRMZELECHO.msf2.TextMatrix(JX&, 2) = ""                              ' JURISDICCION
                    FRMZELECHO.msf2.TextMatrix(JX&, 3) = CuitBanco$                                         ' CUIT
                    FRMZELECHO.msf2.TextMatrix(JX&, 4) = ""                                               ' CBU
                    FRMZELECHO.msf2.TextMatrix(JX&, 5) = ""                              ' TIPO DE CUENTA
                    FRMZELECHO.msf2.TextMatrix(JX&, 6) = ""                                     ' TIPO DE MONEDA
                    FRMZELECHO.msf2.TextMatrix(JX&, 7) = TRIM$(FORMATNUM$(ImpoRetenido#, "F10.2"))   ' IMPORTE RETENIDO
                    FRMZELECHO.msf2.TextMatrix(JX&, 8) = TxtContrapartidaInvalida$                                       ' OBSERVACIONES
                End If
                If TxtCuentaContableInvalida$ <> "" Then
                    JX& = JX& + 1
                    FRMZELECHO.msf2.Rows = JX& + 1
                    FRMZELECHO.msf2.TextMatrix(JX&, 1) = IdAsiento$                              ' COMPROBANTE
                    FRMZELECHO.msf2.TextMatrix(JX&, 2) = ""                              ' JURISDICCION
                    FRMZELECHO.msf2.TextMatrix(JX&, 3) = CuitBanco$                                         ' CUIT
                    FRMZELECHO.msf2.TextMatrix(JX&, 4) = ""                                               ' CBU
                    FRMZELECHO.msf2.TextMatrix(JX&, 5) = ""                              ' TIPO DE CUENTA
                    FRMZELECHO.msf2.TextMatrix(JX&, 6) = ""                                     ' TIPO DE MONEDA
                    FRMZELECHO.msf2.TextMatrix(JX&, 7) = TRIM$(FORMATNUM$(ImpoRetenido#, "F10.2"))   ' IMPORTE RETENIDO
                    FRMZELECHO.msf2.TextMatrix(JX&, 8) = TxtCuentaContableInvalida$                                       ' OBSERVACIONES
                End If
                If TxtCuitInvalidoe$ <> "" Then
                    JX& = JX& + 1
                    FRMZELECHO.msf2.Rows = JX& + 1
                    FRMZELECHO.msf2.TextMatrix(JX&, 1) = IdAsiento$                              ' COMPROBANTE
                    FRMZELECHO.msf2.TextMatrix(JX&, 2) = ""                              ' JURISDICCION
                    FRMZELECHO.msf2.TextMatrix(JX&, 3) = CuitBanco$                                         ' CUIT
                    FRMZELECHO.msf2.TextMatrix(JX&, 4) = ""                                               ' CBU
                    FRMZELECHO.msf2.TextMatrix(JX&, 5) = ""                              ' TIPO DE CUENTA
                    FRMZELECHO.msf2.TextMatrix(JX&, 6) = ""                                     ' TIPO DE MONEDA
                    FRMZELECHO.msf2.TextMatrix(JX&, 7) = TRIM$(FORMATNUM$(ImpoRetenido#, "F10.2"))   ' IMPORTE RETENIDO
                    FRMZELECHO.msf2.TextMatrix(JX&, 8) = TxtCuitInvalidoe$                                     ' OBSERVACIONES
                End If
            Else
                For iBanco = 1 To 1024
                    If CCBanco(iBanco) = CuentaContableBanco% Then
                        ImporteRetenidoPorBanco(iBanco) = ImporteRetenidoPorBanco(iBanco) + ImpoRetenido#
                        GoTo 95
                    End If
                Next iBanco
                
                ' EL BANCO NO EXISTE EN EL VECTOR CCBANCO, POR LO TANTO LO AGREGA
                IndiceBanco = IndiceBanco + 1
                CCBanco(IndiceBanco) = CuentaContableBanco%
                ImporteRetenidoPorBanco(IndiceBanco) = ImporteRetenidoPorBanco(IndiceBanco) + ImpoRetenido#
            
                If IndiceBanco = 1024 Then
                    Call COMUNI("Se Ha Superado el Máximo de Cuentas Bancarias.\Las Retenciones Serán Exportadas Parcialmente.\ \Consulte a su Soporte de Sistemas.")
                    GoTo 98
                End If
            End If
            
            
'            RetencionFallida% = 0
'            IdAsiento$ = SAFETEXT$(!NUIDASIEN)
'            ImpoRetenido# = XVALO(!IDebePase) - XVALO(!IHabePase)
'
'            ' BUSCA LA CUENTA CONTABLE DEL BANCO EN LA CONTRAPARTIDA
'            TextoError$ = ""
'            CondicionSql$ = " AND IDEBEPASE > 0.005"
'            If ImpoRetenido# > 0.005 Then CondicionSql$ = " AND IHabePase > 0.005"
'            If CantRegistros("DETAASIEN", "NUIDASIEN='" & IdAsiento$ & "' and StatPase = 1" & CondicionSql$) > 1 Then
'                RetencionFallida% = 1: FalloEnGeneracion% = 1
'                TextoError$ = "Imposible Detectar Banco. Existe más de una Cuenta de Contrapartida."
'            End If
'
'            ' BUSCA EL NUMERO DE CUENTA DEL PROVEEDOR 'BANCO' A PARTIR DE CONSIDERAR LA CUENTA CONTABLE
'            ' ANTERIOR COMO SU CUENTA MADRE EN EL MAESTRO
'            CuentaContableBanco% = XVALO(VALOBADA("DETAASIEN", "CueCoInt", "NUIDASIEN='" & IdAsiento$ & "'" & CondicionSql$))
'            NumeroProveedor% = XVALO(VALOBADA("PROVED12", "NUME_CLI", "CUMA_CLI=" & CStr(CuentaContableBanco%)))
'            If NumeroProveedor% = 0 Then
'                RetencionFallida% = 1: FalloEnGeneracion% = 1
'                If TextoError$ = "" Then
'                    TextoError$ = "La Contrapartida No es Cta. Contable Madre de Algún Proveedor."
'                Else
'                    TextoError$ = TextoError$ & ", la Contrapartida No es Cta. Contable Madre de Algún Proveedor."
'                End If
'            End If
'
'            RegistroBanco$ = FILTERGETRECORD$("CUEBANC", 1, CODCUE$(CuentaContableBanco%))
'            If TRIM$(RegistroBanco$) = "" Then
'                RetencionFallida% = 1: FalloEnGeneracion% = 1
'                If TextoError$ = "" Then
'                    TextoError$ = "Falta Cta. Contable de Banco en Tabla de Cuentas Bancarias"
'                Else
'                    TextoError$ = TextoError$ & ", falta Cta. Contable de Banco en Tabla de Cuentas Bancarias"
'                End If
'            End If
'
'
'            ' SI ALGUN DATO NO PASO LA VALIDACION, CARGA LA RETENCION EN FRMZELECHO PARA QUE LUEGO EL USUARIO
'            ' PUEDA VER QUE SUCEDIO
'            If RetencionFallida% = 1 And UltimoIdAsiento$ <> IdAsiento$ Then
'                JX& = JX& + 1
'                FRMZELECHO.MSF2.Rows = JX& + 1
'                FRMZELECHO.MSF2.TextMatrix(JX&, 1) = CodigoDeJurisdiccion$                              ' JURISDICCION
'                FRMZELECHO.MSF2.TextMatrix(JX&, 2) = CuitBanco$                                         ' CUIT
'                FRMZELECHO.MSF2.TextMatrix(JX&, 3) = CBU$                                               ' CBU
'                FRMZELECHO.MSF2.TextMatrix(JX&, 4) = TipoDeCuentaBancaria$                              ' TIPO DE CUENTA
'                FRMZELECHO.MSF2.TextMatrix(JX&, 5) = TipoDeMoneda$                                      ' TIPO DE MONEDA
'                FRMZELECHO.MSF2.TextMatrix(JX&, 6) = TRIM$(FORMATNUM$(ImporteJurisdiccion#, "F10.2"))   ' IMPORTE RETENIDO
'                FRMZELECHO.MSF2.TextMatrix(JX&, 7) = TextoError$                                        ' OBSERVACIONES
'                UltimoIdAsiento$ = IdAsiento$
'            Else
'            ' TODAS LAS RETENCIONES SUPERARON LAS VALIDACIONES, CON LO CUAL SE GUARDA EL REGISTRO
'                For iBanco = 1 To 1024
'                    If CCBanco(iBanco) = CuentaContableBanco% Then
'                        ImporteRetenidoPorBanco(iBanco) = ImporteRetenidoPorBanco(iBanco) + ImpoRetenido#
'                        GoTo 95
'                    End If
'                Next iBanco
'
'                ' EL BANCO NO EXISTE EN EL VECTOR CCBANCO, POR LO TANTO LO AGREGA
'                IndiceBanco = IndiceBanco + 1
'                CCBanco(IndiceBanco) = CuentaContableBanco%
'                ImporteRetenidoPorBanco(IndiceBanco) = ImpoRetenido#
'
'                If IndiceBanco = 1024 Then
'                    Call COMUNI("Se Ha Superado el Máximo de Cuentas Bancarias.\Las Retenciones Serán Exportadas Parcialmente.\ \Consulte a su Soporte de Sistemas.")
'                    Exit Do
'                End If
'            End If
            
95          .MoveNext
        Loop
    End With
    
    RstAsientos_SIRCREB.Close
    Set RstAsientos_SIRCREB = Nothing
    
98  For iBanco = 1 To IndiceBanco
    
        If ImporteRetenidoPorBanco(iBanco) > 0.005 Then
                ConsultaSql$ = "SELECT * FROM SIRCREB WITH(NOLOCK) WHERE PERIODO=" & PeriodoRetenido$
                Set RstCoeficientes_SIRCREB = READSET(ConsultaSql$)
                With RstCoeficientes_SIRCREB
                    If (.EOF And .BOF) Then
                        Call COMUNI("Imposible Exportar Recaudaciones.\ \Falta Definir Coeficientes de Distribución\Para el Período '" & PeriodoRetenido$ & "'")
                        Screen.MousePointer = 1
                        GoTo 97
                    End If
                    Do While Not .EOF
                        RetencionFallida% = 0
                        TxtJurisdiccionInvalida$ = ""
                        TxtFaltaCBU = ""
                        TipoDeCuentaBancaria$ = ""
                        TipoDeMoneda$ = ""
                        
                            CodigoDeJurisdiccion$ = CStr(!CodigoJurisdiccion)
                        CodigoDeJurisdiccion$ = Right("000" & CodigoDeJurisdiccion$, 3)
                        
                        CuentaContableBanco% = CCBanco(iBanco)
                        
                        ' A PARTIR DE LA CUENTA DEL BANCO OBTIENE EL CUIT DEL PROVEEDOR, CONSIDERANDO A DICHA CUENTA
                        ' COMO LA CUENTA MADRE.
                        CuitBanco$ = VALOBADA("PROVED12", "CUIT_CLI", "CUMA_CLI=" & CStr(CuentaContableBanco%))
                        
                        ' OBTIENE DATOS DEL BANCO Y A SU VEZ SE VERIFICA QUE LA CUENTA CONTABLE SEA UNA
                        ' DE LAS CONFIGURADAS EN LA TABLA DE CUENTAS BANCARIAS
                        RegistroBanco$ = FILTERGETRECORD$("CUEBANC", 1, CODCUE$(CuentaContableBanco%))
                        CBU$ = TRIM$(Mid$(RegistroBanco$, 49, 25))
                        TipoDeCuentaBancaria$ = Mid$(RegistroBanco$, 74, 2)
                        TipoDeMoneda$ = TRIM$(Mid$(RegistroBanco$, 76, 2))
                        Coeficiente# = !Coeficiente
                        ImporteJurisdiccion# = ImporteRetenidoPorBanco(iBanco) * Coeficiente#
                        
                        ' VALIDACION DE LOS DATOS
                        TextoError$ = ""
                        If CodigoDeJurisdiccion$ = "" Or XVALO(CodigoDeJurisdiccion$) = 0 Then
                            RetencionFallida% = 1: TxtJurisdiccionInvalida$ = "Jurisdicción Inválida"
                        End If
                        
                        If TRIM$(CBU$) = "" Then
                            RetencionFallida% = 1: TxtFaltaCBU = "Falta CBU para " & DENOCUE$(CuentaContableBanco%)
                        End If
                        
                        If TRIM$(TipoDeCuentaBancaria$) = "" Then
                            RetencionFallida% = 1: TxtFaltaTipoCuenta$ = "Falta Tipo de Cuenta para " & DENOCUE$(CuentaContableBanco%)
                        End If
                        
                        If TRIM$(TipoDeMoneda$) = "" Then
                            RetencionFallida% = 1: TxtFaltaTipoMoneda$ = "Falta Tipo de Moneda para " & DENOCUE$(CuentaContableBanco%)
                        End If
            
                        ' SI ALGUN DATO NO PASO LA VALIDACION, CARGA LA RETENCION EN FRMZELECHO PARA QUE LUEGO EL USUARIO
                        ' PUEDA VER QUE SUCEDIO
                        If RetencionFallida% = 1 Then
                            FalloGeneracion% = 1
                            If TxtJurisdiccionInvalida$ <> "" Then
                                JX& = JX& + 1
                                FRMZELECHO.msf2.Rows = JX& + 1
                                FRMZELECHO.msf2.TextMatrix(JX&, 1) = CodigoDeJurisdiccion$                              ' JURISDICCION
                                FRMZELECHO.msf2.TextMatrix(JX&, 2) = CuitBanco$                                         ' CUIT
                                FRMZELECHO.msf2.TextMatrix(JX&, 3) = CBU$                                               ' CBU
                                FRMZELECHO.msf2.TextMatrix(JX&, 4) = TipoDeCuentaBancaria$                              ' TIPO DE CUENTA
                                FRMZELECHO.msf2.TextMatrix(JX&, 5) = TipoDeMoneda$                                      ' TIPO DE MONEDA
                                FRMZELECHO.msf2.TextMatrix(JX&, 6) = TRIM$(FORMATNUM$(ImporteJurisdiccion#, "F10.2"))   ' IMPORTE RETENIDO
                                FRMZELECHO.msf2.TextMatrix(JX&, 7) = TxtJurisdiccionInvalida$                                        ' OBSERVACIONES
                            End If
                            If TxtFaltaCBU <> "" Then
                                JX& = JX& + 1
                                FRMZELECHO.msf2.Rows = JX& + 1
                                FRMZELECHO.msf2.TextMatrix(JX&, 1) = CodigoDeJurisdiccion$                              ' JURISDICCION
                                FRMZELECHO.msf2.TextMatrix(JX&, 2) = CuitBanco$                                         ' CUIT
                                FRMZELECHO.msf2.TextMatrix(JX&, 3) = CBU$                                               ' CBU
                                FRMZELECHO.msf2.TextMatrix(JX&, 4) = TipoDeCuentaBancaria$                              ' TIPO DE CUENTA
                                FRMZELECHO.msf2.TextMatrix(JX&, 5) = TipoDeMoneda$                                      ' TIPO DE MONEDA
                                FRMZELECHO.msf2.TextMatrix(JX&, 6) = TRIM$(FORMATNUM$(ImporteJurisdiccion#, "F10.2"))   ' IMPORTE RETENIDO
                                FRMZELECHO.msf2.TextMatrix(JX&, 7) = TxtFaltaCBU                                        ' OBSERVACIONES
                            End If
                            If TipoDeCuentaBancaria$ <> "" Then
                                JX& = JX& + 1
                                FRMZELECHO.msf2.Rows = JX& + 1
                                FRMZELECHO.msf2.TextMatrix(JX&, 1) = CodigoDeJurisdiccion$                              ' JURISDICCION
                                FRMZELECHO.msf2.TextMatrix(JX&, 2) = CuitBanco$                                         ' CUIT
                                FRMZELECHO.msf2.TextMatrix(JX&, 3) = CBU$                                               ' CBU
                                FRMZELECHO.msf2.TextMatrix(JX&, 4) = TipoDeCuentaBancaria$                              ' TIPO DE CUENTA
                                FRMZELECHO.msf2.TextMatrix(JX&, 5) = TipoDeMoneda$                                      ' TIPO DE MONEDA
                                FRMZELECHO.msf2.TextMatrix(JX&, 6) = TRIM$(FORMATNUM$(ImporteJurisdiccion#, "F10.2"))   ' IMPORTE RETENIDO
                                FRMZELECHO.msf2.TextMatrix(JX&, 7) = TipoDeCuentaBancaria$                                        ' OBSERVACIONES
                            End If
                            If TipoDeMoneda$ <> "" Then
                                JX& = JX& + 1
                                FRMZELECHO.msf2.Rows = JX& + 1
                                FRMZELECHO.msf2.TextMatrix(JX&, 1) = CodigoDeJurisdiccion$                              ' JURISDICCION
                                FRMZELECHO.msf2.TextMatrix(JX&, 2) = CuitBanco$                                         ' CUIT
                                FRMZELECHO.msf2.TextMatrix(JX&, 3) = CBU$                                               ' CBU
                                FRMZELECHO.msf2.TextMatrix(JX&, 4) = TipoDeCuentaBancaria$                              ' TIPO DE CUENTA
                                FRMZELECHO.msf2.TextMatrix(JX&, 5) = TipoDeMoneda$                                      ' TIPO DE MONEDA
                                FRMZELECHO.msf2.TextMatrix(JX&, 6) = TRIM$(FORMATNUM$(ImporteJurisdiccion#, "F10.2"))   ' IMPORTE RETENIDO
                                FRMZELECHO.msf2.TextMatrix(JX&, 7) = TipoDeMoneda$                                        ' OBSERVACIONES
                            End If
                        Else
                        ' TODAS LAS RETENCIONES SUPERARON LAS VALIDACIONES, CON LO CUAL SE GUARDA EL REGISTRO
                            ImporteRetenido$ = TRIM$(FORMATNUM$(ImporteJurisdiccion#, "F10.2"))
                            ImporteRetenido$ = Right("0000000000" & ImporteRetenido$, 10)
                            ImporteRetenido$ = REPLACAR$(ImporteRetenido$, ".", ",")
                            PERIODO$ = Mid$(PeriodoRetenido$, 1, 4) & "/" & Mid$(PeriodoRetenido$, 5, 2)
                            RegistroRetencion$ = CodigoDeJurisdiccion$ & CuitBanco$ & PERIODO$ & CBU$ & TipoDeCuentaBancaria$ & TipoDeMoneda$ & ImporteRetenido$
                            Print #NumeroArchivo%, RegistroRetencion$
            
                            TotalDeRetenciones% = TotalDeRetenciones% + 1
                        End If
            
                        .MoveNext
                    Loop
                End With
                RstCoeficientes_SIRCREB.Close
                Set RstCoeficientes_SIRCREB = Nothing
        End If
    Next iBanco

97  On Error Resume Next
    RstCoeficientes_SIRCREB.Close
    Set RstCoeficientes_SIRCREB = Nothing
    Close #NumeroArchivo%
    On Error GoTo 0
    
    Screen.MousePointer = 1

    If FalloEnGeneracion% = 1 Then
        Call MENSERR(24, "Se Han Detectado Errores en Algunas de las Retenciones.\ \Acepte este Mensaje Para Visualizar\los Errores Detectados.")
        FRMZELECHO.Show 1
        Call FILEKILL(RutaDelArchivo$)
    ElseIf TotalDeRetenciones% > 0 Then
        Call COMUNI("Se Han Exportado Correctamente " & CStr(TotalDeRetenciones%) & " Retenciones")
    Else
        Call COMUNI("No Se Han Encontrado Retenciones en el Período Elegido.")
    End If

End Sub

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
    '
    Command10.Enabled = False
    '
    If DERACCE%("REIMPUCO", "Re-Imputación de Comprobantes de Compras") < 2 Then
      GoTo 99
    End If
    '
    If VALIMESAN% < 1 Then GoTo 99
    Des% = Desde%: DESDEX = CDbl(CVDAT(Des%))
    Has% = Hasta%: HASTAX = CDbl(CVDAT(Has%))
    
    DESDAT = FETEXCOR1$(Des%)
    HASDAT = FETEXCOR1$(Has%)

10  Call SELECHO("PROVED1")
    NC% = XVALO(VALACT1$("PROVED1"))
    If NC% < 1 Then GoTo 99
    '
    Call APERBASDAT("CABAN")
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FCOM' OR TIPOMOVIM = 'AJUP') "
    SQLX$ = SQLX$ + "AND NuProve = " & NC% & " "
'    SQLX$ = SQLX$ + "AND FechConta >= " & DESDEX & " "
'    SQLX$ = SQLX$ + "AND FechConta <= " & HASTAX & " "
    SQLX$ = SQLX$ + "AND FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "Order by FechConta ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
25  On Error Resume Next

    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
  
    '
    JJ& = 0
    '
    NIVCONXXX% = NIVCONFI%
    '
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          '\\\OT-08-1032-FL-12/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
          If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70
          '\\\OT-08-1032-FL-FIN
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!CODICOM_LAR, 11, 8))
            y$ = MKS$(JJ& + 1) + AJUSTI$(!CODICOM_LAR, 18) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUERE", y$, JJ&)
          End If
70      .MoveNext
        Loop
      End If
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    Call SETULTREG("!SELCUERE", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no Tiene\Comprobantes de Compras\que Puedan Re-Imputarse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
16  Call SELECHO("!SELCUERE")
    If VALACT1$("!SELCUERE") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = TRIM$(Left$(VALACT2$("!SELCUERE"), 18))
    '
20  ATRI$ = DOCULA$ + " - $ " + TRIM$(Mid$(VALACT2$("!SELCUERE"), 19, 14)) + " - " + TRIM$(rasocli$(NC%))
    '
    Call BLANARCH("!REIMPUTA")
    'JANDY SQL MODO PESIMISTA CON UN UPDATE VER PERFORMANCE!!
    
    Dim CABASEL As ADODB.Recordset
    Set CABASEL = New ADODB.Recordset
  
    SQLX$ = "SELECT * FROM CAJABANC "
    'SQLX$ = SQLX$ + "WHERE INT(CDBL(FechConta)) >= " & DESDEX & " "
    'SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASTAX & " "
    SQLX$ = SQLX$ + "WHERE FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & DOCULA$ & "' "
    SQLX$ = SQLX$ + "AND NuProve = " & NC% & " "
    '
    'Set CABASEL = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
26  On Error Resume Next
    CABASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 26
    End If
    On Error GoTo 0
   
  
    SUIMPU# = 0
    '
    JJR& = 0
    With CABASEL
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call MENSERR(24, "Imposible Editar Imputación Contable\del Comprobante Elegido")
        GoTo 10
      End If
      '
      Do While Not .EOF
        CUE% = !CUECONTAI
        If CUE% > 0 Then
           IPA# = !Idebecomp - !IHABECOMP
           y$ = REGBLAN$("!REIMPUTA")
           Call REPLA(y$, CODCUE$(CUE%), 1, 12)
           Call REPLA(y$, MKD$(IPA#), 13, 8)
           '\\\OT-06-0124-WAR-03/04/06///
           CC1$ = "": If IsNull(!CenCosto) = False Then CC1$ = !CenCosto
           Call REPLA(y$, CC1$, 21, 8)
           '\\\OT-06-0124-WAR-FIN///
           JJR& = JJR& + 1
           Call GRAREG("!REIMPUTA", y$, JJR&)
           SUIMPU# = SUIMPU# + IPA#
        End If
      .MoveNext
      Loop
    End With
    
    Call CIERRARCH("!REIMPUTA")
    '
    If Abs(SUIMPU#) >= 0.005 Then
      Call MENSERR(24, "Imposible Editar Imputación Contable.\Comprobante No Balanceado.")
      GoTo 10
    End If
    '
30  ATRI$ = "TITUPAN=Re-Imputación Contable: " + ATRI$
    'VTB% = VENTABU%("REIMPUTA/" + ATRI$)
    VTB% = VENTABU%("REIMPUCO/" + ATRI$) '\\\OT-06-0124-WAR-03/04/06///
    If VTB% < 0 Then GoTo 10
    '
    REIMPU% = 0
    CABASEL.MoveFirst
    For U& = 1 To ULTREG&("!REIMPUTA")
      If U& <= JJR& Then
        y$ = REGLEIDO$("!REIMPUTA", U&)
        CUENUE% = CUEINT%(Left$(y$, 12))
        CC1$ = Mid$(y$, 21, 8)
        '
        If CUENUE% > 0 Then
          With CABASEL
          '.Edit
          !CUECONTAI = CUENUE%
          !CenCosto = CC1$
          If !TipoMovim = "COMPRA" Then
             ' SOLICITADO POR Claudia Matsusaka 2l 23/2/07 que
             ' cambie la descripción acorde con la nueva cuenta
             ' contable
             DEMOVX$ = DENOCUE$(CUENUE%)
             NC1% = (-1) * Abs(NC%)
             !DESCRIMOV = Left$(TRIM$(DEMOVX$) + " - " + rasocli$(NC1%), 48)
          End If
          .Update
          REIMPU% = 1
          GoTo 35
          End With
        End If
35      If Not CABASEL.EOF Then
          CABASEL.MoveNext
        End If
      End If
      '
   Next U&
   '
   CABASEL.Close
   Set CABASEL = Nothing
    
   If REIMPU% > 0 Then
        Call COMUNI("Re-Imputación Completada")
      Else
        Call COMUNI("No se Pudo Completar la Re-Imputación")
   End If
   '
   GoTo 16
   '
99 Command10.Enabled = True
   '
End Sub

Private Sub Command11_Click()

   Call COMUNI("Opciones Trasladadas al Menú Superior denominado 'Impuestos'")
   Exit Sub
   
   Command11.Enabled = False
   '\\\OT-05-0275-WAR-11/05/05///
'2  JJ0% = COMALTER%("Opciones de Trabajo:\\Ingresos Brutos\C.I.T.I. Compras\Ganancias\Regimen Informativo de Compras\IVA")
2  JJ0% = COMALTER%("Opciones de Trabajo:\\Ingresos Brutos\I.V.A.\Ganancias\Régimen Informativo de Compras")
   If JJ0% < 1 Or JJ0% > 4 Then GoTo 99
   If JJ0% = 1 Then
3    JJ00% = COMALTER%("Jurisdicción:\\Buenos Aires\Convenio Multilateral")
     If JJ00% < 1 Or JJ00% > 2 Then GoTo 2
     
     ' *** BUENOS AIRES
     If JJ00% = 1 Then
     
5        JJ1% = COMALTER%("Opciones de Trabajo:\\Percepciones\Retenciones")
         If JJ1% < 1 Or JJ1% > 2 Then GoTo 3
         '\\\OT-05-0275-WAR-FIN///
         If JJ1% = 2 Then
           'RETENCIONES
           OPX% = COMALTER%("Opciones de Trabajo:\\Subdiario I.Brutos\Archivo Transferencia")
           If OPX% < 1 Or OPX% > 2 Then GoTo 5
           '
           Call DESHASFECHA(Des%, Has%, PERIO$)
           If Des% < 1 Then GoTo 99
           '
           If InStr(1, EMPREAC$, "FERVI") > 0 Then
                FF1% = FECHANUM("01/01/80")
                FIN& = ULTREG&("SURETIB")
                For U& = 1 To FIN&
                    XX$ = REGLEIDO$("SURETIB", U&)
                    NOP& = CVS(Mid$(XX$, 7, 4))
                    If NOP& = 75157 Or NOP& = 75158 Then
                          Call REPLA(XX$, MKI(FF1%), 1, 2)
                          Call GRAREG("SURETIB", XX$, U&)
                    End If
                Next U&
           End If
           '
           If OPX% = 1 Then
             CORRESP$ = PERIO$
             Call HEADERS("SUBRETI", "")
             Call HEADERS("SUBRETI", "Corresponde a: " + CORRESP$)
             j& = 0: FIN& = ULTREG("SURETIB")
             For U& = 1 To FIN&
               Call TRACE(20, U&, FIN&)
               XX$ = REGLEIDO$("SURETIB", U&)
               FEMIX% = CVI(Left$(XX$, 2))
               If FEMIX% >= Des% Then
                 If FEMIX% <= Has% Then
                   j& = j& + 1
                   Call GRAREG("SUBRETI", XX$, j&)
                 End If
               End If
             Next U&
             Call SETULTREG("SUBRETI", j&)
             Call FINAL("*SUBRETI")
           ElseIf OPX% = 2 Then
             CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
             CommonDialog1.DialogTitle = "Archivo Origen de Datos"
             '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
             CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
             CommonDialog1.FilterIndex = 1
             On Error GoTo 99
             CommonDialog1.ShowOpen
             On Error GoTo 0
             NOMBASALID$ = UCase$(CommonDialog1.FileName)
             ASALID$ = TRIM$(NOMBASALID$)
             If ASALID$ = "" Then GoTo 5
             
             NX% = FILEOPEN%(ASALID$, 2, 46)
             ASALID1$ = LUCOM$ + "DDJJRDET.TXT"
             NX1% = FILEOPEN%(ASALID1$, 2, 128)
             '
10           cuitagen$ = CONTROL$("CUITAGEN", "CUITAGEN")
             cuitagen$ = InputBox$("Ingrese C.U.I.T. Agente de Retención", "C.U.I.T. Agente Retencion", cuitagen$)
             If VALICUIT%(cuitagen$) <> 1 Then
               Call MENSERR(24, "C.U.I.T. no Válido")
               GoTo 10
             End If
             '
             REBLA$ = Space$(219)
             '\\\OT-05-0285-WAR-12/05/05///
             Call REPLA(REBLA$, cuitagen$, 1, 13)
             'PL1$ = FECHATEX(CVINT(CVDAT(HAS%)))
             PL11 = CVDAT(Has%)
             PL1$ = Format(PL11, "YYYYMM")
             Call REPLA(REBLA$, PL1$, 14, 6)
             Call REPLA(REBLA$, "S", 73, 1)
             PL22 = CVDAT(Has%)
             PL2$ = Format(PL22, "YYYYMMDD")
             Call REPLA(REBLA$, PL2$, 144, 8)
             Call REPLA(REBLA$, "OTA0001", 152, 7)
             Call REPLA(REBLA$, "PSO", 167, 3)
             Call REPLA(REBLA$, "2.5000", 188, 6)
             Call REPLA(REBLA$, "S", 210, 1)
             '\\\OT-05-0285-WAR-FIN///
             CAREGI% = 0
             For U& = 1 To ULTREG&("SURETIB")
20             XX$ = REGLEIDO$("SURETIB", U&)
               FEMIX% = CVI(Left$(XX$, 2))
               If FEMIX% >= Des% Then
                 If FEMIX% <= Has% Then
                   TTXX$ = Space$(45)
                     CUITTX$ = TRIM$(Mid$(XX$, 51, 14))
                     If TRIM$(CUITTX$) = "" Or VALICUIT%(CUITTX$) <> 1 Then
                       CUITTX$ = InputBox$("Ingrese Numero de CUIT", TRIM$(Mid$(XX$, 21, 30)))
                       Call REPLA(XX$, CUITTX$, 51, 14)
                       Call GRAREG("SURETIB", XX$, U&)
                       GoTo 20
                     End If
                   Call REPLA(TTXX$, CUITTX$, 1, 13)
                     FEX = CVI(Left$(XX$, 2))
                     FEXU$ = FECHATEX$(FEX)
                     FEXU$ = Left$(FEXU$, 6) + "20" + Mid$(FEXU$, 7, 2)
                   Call REPLA(TTXX$, FEXU$, 14, 10)
                     SUCUX$ = "0001"
                   Call REPLA(TTXX$, SUCUX$, 24, 4)
                     NUCOX$ = TRIM$(Str$(CVS(Mid$(XX$, 3, 4))))
                     While Len(NUCOX$) < 8: NUCOX$ = "0" + NUCOX$: Wend
                   Call REPLA(TTXX$, NUCOX$, 28, 8)
                     IMPOX$ = REPLACAR$(TRIM$(Format$(CVD(Mid$(XX$, 87, 8)), "#########0.00")), ".", ",")
                     While Len(IMPOX$) < 11: IMPOX$ = "0" + IMPOX$: Wend
                     'While Len(IMPOX$) < ANCHIMPOR%: IMPOX$ = "0" + IMPOX$: Wend
                   'Call REPLA(TTXX$, IMPOX$, 36, ANCHIMPOR%)
                   Call REPLA(TTXX$, IMPOX$, 36, 11)
                   Call REPLA(TTXX$, "A", 47, 1) ' SE AGREGA SEGUN NUEVA NORMATIVA DE ARBA SETIEMBRE 2011
                   If XVALO(IMPOX$) > 0.005 Then ' SE AGREGA PORQUE EL SIAP PROTESTA SI ES CERO
                        Print #NX%, TTXX$
                   End If
                   '\\\OT-05-0285-WAR-12/05/05///
                   TTYY$ = REBLA$
                   Call REPLA(TTYY$, CUITTX$, 20, 13)
                   NC% = CVI(Mid$(XX$, 95, 2))
                   RS1$ = rasocli(NC% * (-1))
                   Call REPLA(TTYY$, AJUSTI(RS1$, 40), 33, 40)
                   DO1$ = TRIM$(DOMICLI(NC% * (-1)))
                   LO1$ = TRIM$(LOCACLI(NC% * (-1)))
                   DP1$ = AJUSTI(DO1$ + " " + LO1$, 70)
                   Call REPLA(TTYY$, DP1$, 74, 70)
                   'NUMERO COMPROBANTE
                     NUCOX$ = TRIM$(Str$(CVS(Mid$(XX$, 3, 4))))
                     While Len(NUCOX$) < 8: NUCOX$ = "0" + NUCOX$: Wend
                   Call REPLA(TTYY$, NUCOX$, 159, 8)
                   'BASE IMPONIBLE
                   IMPOX$ = FORMATNUM$(CVD(Mid$(XX$, 79, 8)), "F16.2")
                   Call REPLA(TTYY$, AJUSTD(IMPOX$, 16), 170, 16)
                 ' IMPORTE RETENIDO
                   If CVD(Mid$(XX$, 87, 8)) < 0.005 Then GoTo 201 ' SALTEA REGISTROS SIN RETENCION
                   IMPOX2$ = FORMATNUM$(CVD(Mid$(XX$, 87, 8)), "F16.2")
                   Call REPLA(TTYY$, AJUSTD(IMPOX2$, 16), 194, 16)
                   ' tasa retenida - MODIFICADO 3/8/2005 - ANTES ERA FIJO =2.5000
                   TARET = 0
                   On Error Resume Next
                   TARET = CVD(Mid$(XX$, 87, 8)) / CVD(Mid$(XX$, 79, 8))
                   On Error GoTo 0
                   Call REPLA(TTYY$, FORMATNUM$(TARET, "F6.4"), 188, 6)
                   'NUMERO DE ORDEN
                   CAREGI% = CAREGI% + 1
                   NUO$ = AJUSTD(CStr(CAREGI%), 9)
                   Call REPLA(TTYY$, NUO$, 211, 9)
                   Print #NX1%, TTYY$
                   '\\\OT-05-0285-WAR-FIN///
                 End If
               End If
201          Next U&
             ' PONE MARCA DE FIN DE ARCHIVO SEGúN MUESTRA RECIBIDA
             Print #NX1%, Chr$(26)
             ' FIN MARCA DE FIN DE ARCHIVO
             Close #NX%
             Close #NX1%
             ASALID3$ = REPLACAR(ASALID$, "\", "/")
             ASALID4$ = REPLACAR(ASALID1$, "\", "/")
             Call COMUNI("Archivos de Transferencia\ '" + ASALID3$ + "' y '" + ASALID4$ + "' \Generados con " + TRIM$(Str$(CAREGI%)) + " Registros.")
            
           End If
           GoTo 5
         End If
          '\\\OT-05-0275-WAR-11/05/05///
         If JJ1% = 1 Then
            'PERCEPCIONES
202         OP% = COMALTER%("Opciones de Trabajo:\\Subdiario de I.Brutos\Archivo Transferencia")
            If OP% < 1 Or OP% > 2 Then GoTo 5
            
21          Call DESHASFECHA(Des%, Has%, PERIO$)
            If Des% < 1 Then GoTo 99
            'JANDY SQL
            DES1$ = FETEXCOR1$(Des%)
            HAS1$ = FETEXCOR1$(Has%)
            Call APERBASDAT("CABAN")
            Dim PERLISTMP As ADODB.Recordset
            Set PERLISTMP = New ADODB.Recordset
            PERLISTMP.CursorLocation = adUseClient
            SQLX$ = "SELECT * FROM CAJABANC "
            SQLX$ = SQLX$ + "WHERE FECHCONTA >= '" & DES1$ & "'"
            SQLX$ = SQLX$ + "AND FECHCONTA <= '" & HAS1$ & "' "
            SQLX$ = SQLX$ + "AND TIPOMOVIM = 'IBRU' "
            SQLX$ = SQLX$ + "AND VARIMOVIM = 'PERC' "
            SQLX$ = SQLX$ + "AND NUPROVE > 0 "
            SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC "
25          On Error Resume Next
            PERLISTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
       
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 25
            End If
            On Error GoTo 0
            '
            If OP% = 1 Then ' SUBDIARIO DE INGRESOS BRUTOS
                With PERLISTMP
                  '
                  On Error Resume Next
                  .MoveFirst
                  If Err Then
                    Call COMUNI("No existen Percepciones en el Rango de Fechas Elegido")
                    On Error GoTo 0
                    GoTo 21
                  End If
                  On Error GoTo 0
                  JJ& = 0
                  ZZ$ = REGBLAN("DEPERIB")
                  NIVCONXXX% = NIVCONFI%
                  Do While .EOF = False
                    '\\\OT-08-1032-FL-12/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
                    If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70
                    '\\\OT-08-1032-FL-FIN
                    Call REPLA(ZZ$, MKI(CVINT(!fechconta)), 1, 2)
                    Call REPLA(ZZ$, !CODICOM_LAR, 3, 22)
                    NC% = !NUPROVE
                    CTP$ = CODICLI$((-1) * NC%)
                    Call REPLA(ZZ$, CTP$, 25, 12)
                    IMPER# = !Idebecomp - !IHABECOMP
                    Call REPLA(ZZ$, MKD(IMPER#), 37, 8)
                    CPR$ = TRIM$(Mid$(!OpciMovim, 4))
                    CPR1$ = ECOARCH$("PROVINC", CPR$)
                    Call REPLA(ZZ$, CPR$, 45, 1)
                    CCTT$ = CuitCli$((-1) * NC%)
                    Call REPLA(ZZ$, CCTT$, 70, 13)
                    JJ& = JJ& + 1
                    Call GRAREG("DEPERIB", ZZ$, JJ&)
70                .MoveNext
                  Loop
                  Call SETULTREG("DEPERIB", JJ&)
                  Call FILESORT("DEPERIB", "DEPERIB", 1, 1, "ASC") ' POR FECHA
                  Call FILESORT("DEPERIB", "DEPERIB", 6, 6, "ASC") ' POR JURISDICCIóN
                  Call FINAL("*DEPERIB")
                  '
                End With
                PERLISTMP.Close
                Set PERLISTMP = Nothing
            End If
            If OP% = 2 Then ' ARCHIVO DE TRANSFERENCIA DE PERECEPCIONES
                 Screen.MousePointer = 11
                 'ASALID1$ = LUCOM$ + "DDJJPDET.TXT"
                 'SE AGREGA PARA QUE EL USUARIO INGRESE EL NOMBRE DEL ARCHIVO DE TRANSFERENCIA SEGUN NUEVA NORMATIVA A PARTIR DE SETIEMBRE 2011
                 'NOMBASALID$ = TRIM$(InputBox("Ingrese Nombre del Archivo", "Archivo IIBB", ""))
                 'If NOMBASALID$ = "" Then GoTo 99
                 'ASALID1$ = "C:\FLEXOFT\" & NOMBASALID$ & ".TXT"
                 CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
                 CommonDialog1.DialogTitle = "Archivo Origen de Datos"
                 '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
                 CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
                 CommonDialog1.FilterIndex = 1
                 On Error GoTo 99
                 CommonDialog1.ShowOpen
                 On Error GoTo 0
              
                 DoEvents
                 NOMBASALID$ = UCase$(CommonDialog1.FileName)
                 ASALID1$ = TRIM$(NOMBASALID$)
                 If ASALID1$ = "" Then
                    Screen.MousePointer = 1: GoTo 202
                 End If
                 
                 NX1% = FILEOPEN%(ASALID1$, 2, 128)
                 CAREGI% = 0
                  With PERLISTMP
                  '
                    On Error Resume Next
                    .MoveFirst
                    If Err Then
                      Screen.MousePointer = 1
                      Call COMUNI("No existen Percepciones en el Rango de Fechas Elegido")
                      On Error GoTo 0
                      GoTo 21
                    End If
                    On Error GoTo 0
                    
                    Call SELECHO("PROVINC")
                    PROV$ = VALACT1$("PROVINC")
                    
                    Do While .EOF = False
                      NC% = !NUPROVE
                      
                      PROVIPRO$ = PROVCLI$((-1) * NC%)
                      If PROVIPRO$ <> PROV$ Then GoTo 80
                      
                      CCTT$ = CuitCli$((-1) * NC%)
                      FECPERC$ = CVDAT((CVINT(!fechconta)))
                      CODIKOMLAR$ = SAFETEXT$(!CODICOM_LAR)
                      TICOMP$ = "F"
                      If UCase$(Left$(CODIKOMLAR$, 3)) = "NC-" Then TICOMP$ = "C"
                      If UCase$(Left$(CODIKOMLAR$, 3)) = "RB-" Then TICOMP$ = "R"
                      If UCase$(Left$(CODIKOMLAR$, 3)) = "ND-" Then TICOMP$ = "D"
                      If UCase$(Left$(CODIKOMLAR$, 3)) = "DP-" Then TICOMP$ = "C"
                      If UCase$(Left$(CODIKOMLAR$, 3)) = "CP-" Then TICOMP$ = "D"
                      '
                      LECOMP$ = Mid$(CODIKOMLAR$, 4, 1)
                      NUMSUCURSAL$ = Mid$(CODIKOMLAR$, 6, 4)
                      NUEMISION$ = Right$(CODIKOMLAR$, 8)
                      '
                      'SE AGREGA EL MONTO IMPONIBLE PARA LA IMPORTACION AL SIAP
                      'ES POR CT-102926-FL
                      FEMMI = CVINT(!fechconta)
                      BASEIMPON# = INEFAC#(CODIKOMLAR$, FEMMI, NC%)
                      ' Verifica si es un credito
                      SIGNO% = 0: If BASEIMPON# < 0 Then SIGNO% = 1
                      BASEIMPON# = Abs(BASEIMPON#)
                      '
                      BASEIMPO$ = FORMATNUM(BASEIMPON#, "F12.2")
                      BASEIMPO$ = TRIM$(REPLACAR$(BASEIMPO$, ".", ","))
                      While Len(TRIM$(BASEIMPO$)) < 12: BASEIMPO$ = "0" + BASEIMPO$: Wend
                      ' Agrega el signo cuando se trata de un credito
                      If SIGNO% = 1 Then Call REPLA(BASEIMPO$, "-", 1, 1)
                      '
                      IMPOXRE# = XVALO(!Idebecomp) - XVALO(!IHABECOMP)
                      IMPOXRE# = Abs(IMPOXRE#)
                      '
                      IMPOPERC$ = FORMATNUM(IMPOXRE#, "F11.2")
                      '
                      IMPOPERC$ = TRIM$(REPLACAR$(IMPOPERC$, ".", ","))
                      While Len(TRIM$(IMPOPERC$)) < 11: IMPOPERC$ = "0" + IMPOPERC$: Wend
                      ' Agrega el signo cuando se trata de un credito
                      If SIGNO% = 1 Then Call REPLA(IMPOPERC$, "-", 1, 1)
                      '
                      Call REPLA(TX$, CCTT$, 1, 13) ' CUIT
                      Call REPLA(TX$, FECPERC$, 14, 10) ' FECHA PERCEPCION
                      Call REPLA(TX$, TICOMP$, 24, 1) ' TIPO DE COMPROBANTE
                      Call REPLA(TX$, LECOMP$, 25, 1) ' LETRA COMPROBANTE
                      Call REPLA(TX$, NUMSUCURSAL$, 26, 4) ' NUMERO SUCURSAL
                      Call REPLA(TX$, NUEMISION$, 30, 8) ' NUMERO DE EMISION
                      'Call REPLA(TX$, IMPOPERC$, 38, 11) ' IMPORTE DE PERCEPCION
                      'SE AGREGA EL MONTO IMPONIBLE PARA EL ARCHIVO DE TRASNFERENCIA DE LAS PERCEPCIONES AL SIAP
                      'POR CONSULTA DE DOSIVAC CT-102926-FL - OT-10-0154
                      Call REPLA(TX$, BASEIMPO$, 38, 12) ' MONTO IMPONIBLE DE PERCEPCION
                      Call REPLA(TX$, IMPOPERC$, 50, 11) ' IMPORTE DE PERCEPCION
                      Call REPLA(TX$, "A", 61, 1) ' SE AGREGA SEGUN NUEVA NORMATIVA DE ARBA SETIEMBRE 2011
                      
                      CAREGI% = CAREGI% + 1
                      Print #NX1%, TX$
80                    .MoveNext
                    Loop
                  End With
                  Close #NX1%
                  ASALID2$ = REPLACAR(ASALID1$, "\", "/")
                  Screen.MousePointer = 1
                  Call COMUNI("Archivos de Transferencia\ '" + ASALID2$ + "' \Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.")
                End If
         End If
         GoTo 5
         ' *** IIBB CONVENIO MULTILATERAL
     Else
85          JJ1% = COMALTER%("Opciones de Trabajo:\\Percepciones\Retenciones\Recaudaciones Bancarias (SIRCREB)")
            If JJ1% < 1 Or JJ1% > 3 Then GoTo 3
            If JJ1% = 1 Then
                 'PERCEPCIONES
851              OP% = COMALTER%("Opciones de Trabajo:\\Subdiario de I.Brutos\Archivo Transferencia")
                 If OP% < 1 Or OP% > 2 Then GoTo 85
                 
86               Call DESHASFECHA(Des%, Has%, PERIO$)
                 If Des% < 1 Then GoTo 99
                 DES1$ = FETEXCOR1$(Des%)
                 HAS1$ = FETEXCOR1$(Has%)
                 Call APERBASDAT("CABAN")
                 Dim PERLISTMP0 As ADODB.Recordset
                 Set PERLISTMP0 = New ADODB.Recordset
                 PERLISTMP0.CursorLocation = adUseClient
                 SQLX$ = "SELECT * FROM CAJABANC "
                 SQLX$ = SQLX$ + "WHERE FECHCONTA >= '" & DES1$ & "'"
                 SQLX$ = SQLX$ + "AND FECHCONTA <= '" & HAS1$ & "' "
                 SQLX$ = SQLX$ + "AND TIPOMOVIM = 'IBRU' "
                 SQLX$ = SQLX$ + "AND VARIMOVIM = 'PERC' "
                 SQLX$ = SQLX$ + "AND NUPROVE > 0 "
                 SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC "
                 'Set PERLISTMP = CueCorri.OpenRecordset(SQLX$, 4)
87               On Error Resume Next
                 PERLISTMP0.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
            
                 If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                    On Error GoTo 0
                    Call CapturaErrorOpen
                    GoTo 87
                 End If
                 On Error GoTo 0
                 '
                 If OP% = 1 Then ' SUBDIARIO DE INGRESOS BRUTOS
                     With PERLISTMP0
                       '
                       On Error Resume Next
                       .MoveFirst
                       If Err Then
                         Call COMUNI("No existen Percepciones en el Rango de Fechas Elegido")
                         On Error GoTo 0
                         GoTo 86
                       End If
                       On Error GoTo 0
                       JJ& = 0
                       ZZ$ = REGBLAN("DEPERIB")
                       NIVCONXXX% = NIVCONFI%
                       Do While .EOF = False
                         '\\\OT-08-1032-FL-12/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
                         If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 88
                         '\\\OT-08-1032-FL-FIN
                         Call REPLA(ZZ$, MKI(CVINT(!fechconta)), 1, 2)
                         Call REPLA(ZZ$, !CODICOM_LAR, 3, 22)
                         NC% = !NUPROVE
                         CTP$ = CODICLI$((-1) * NC%)
                         Call REPLA(ZZ$, CTP$, 25, 12)
                         IMPER# = !Idebecomp - !IHABECOMP
                         Call REPLA(ZZ$, MKD(IMPER#), 37, 8)
                         CPR$ = TRIM$(Mid$(!OpciMovim, 4))
                         CPR1$ = ECOARCH$("PROVINC", CPR$)
                         Call REPLA(ZZ$, CPR$, 45, 1)
                         CCTT$ = CuitCli$((-1) * NC%)
                         Call REPLA(ZZ$, CCTT$, 70, 13)
                         JJ& = JJ& + 1
                         Call GRAREG("DEPERIB", ZZ$, JJ&)
88                       .MoveNext
                       Loop
                       Call SETULTREG("DEPERIB", JJ&)
                       Call FILESORT("DEPERIB", "DEPERIB", 1, 1, "ASC") ' POR FECHA
                       Call FILESORT("DEPERIB", "DEPERIB", 6, 6, "ASC") ' POR JURISDICCIóN
                       Call FINAL("*DEPERIB")
                       '
                     End With
                     PERLISTMP0.Close
                     Set PERLISTMP0 = Nothing
                 End If
                 If OP% = 2 Then ' ARCHIVO DE TRANSFERENCIA DE PERECEPCIONES
                      
                      Screen.MousePointer = 11
                      CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
                      CommonDialog1.DialogTitle = "Archivo Origen de Datos"
                      CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
                      CommonDialog1.FilterIndex = 1
                      On Error GoTo 99
                      CommonDialog1.ShowOpen
                      On Error GoTo 0
                   
                      DoEvents
                      NOMBASALID$ = UCase$(CommonDialog1.FileName)
                      ASALID1$ = TRIM$(NOMBASALID$)
                      If ASALID1$ = "" Then
                        Screen.MousePointer = 1: GoTo 851
                      End If
                      ASALID2$ = Mid$(NOMBASALID$, 1, Len(NOMBASALID$) - 4) & " - ADUANERAS.TXT"
                      
                      NX1% = FILEOPEN%(ASALID1$, 2, 51)
                      NX2% = FILEOPEN%(ASALID2$, 2, 56)
                      CAREGI% = 0
                       With PERLISTMP0
                       '
                         On Error Resume Next
                         .MoveFirst
                         If Err Then
                           Screen.MousePointer = 1
                           Call COMUNI("No Existen Percepciones en el Rango de Fechas Elegido")
                           On Error GoTo 0
                           GoTo 21
                         End If
                         On Error GoTo 0
                         
                         Do While .EOF = False
                           NC% = !NUPROVE
                           CCTT$ = CuitCli$((-1) * NC%)
                           FECPERC$ = CVDAT((CVINT(!fechconta)))
                           CODIKOMLAR$ = SAFETEXT$(!CODICOM_LAR)
                           TICOMP$ = "F"
                           If UCase$(Left$(CODIKOMLAR$, 3)) = "NC-" Then TICOMP$ = "C"
                           If UCase$(Left$(CODIKOMLAR$, 3)) = "RB-" Then TICOMP$ = "R"
                           If UCase$(Left$(CODIKOMLAR$, 3)) = "ND-" Then TICOMP$ = "D"
                           If UCase$(Left$(CODIKOMLAR$, 3)) = "DP-" Then TICOMP$ = "C"
                           If UCase$(Left$(CODIKOMLAR$, 3)) = "CP-" Then TICOMP$ = "D"
                           '
                           LECOMP$ = Mid$(CODIKOMLAR$, 4, 1)
                                NUMSUCURSAL$ = REPLACAR$(Mid$(CODIKOMLAR$, 6, 4), "/", "")
                                NUMSUCURSAL$ = REPLACAR$(NUMSUCURSAL$, "-", "")
                                NUSUC& = XVALO(NUMSUCURSAL$)
                                If NUSUC& = 0 Then NUSUC& = 1
                                NUMSUCURSAL$ = CStr(NUSUC&)
                           Do While Len(NUMSUCURSAL$) < 4: NUMSUCURSAL$ = "0" & NUMSUCURSAL$: Loop
                           
                                NUEMISION$ = REPLACAR$(Right$(CODIKOMLAR$, 8), "/", "")
                                NUEMISION$ = REPLACAR$(NUEMISION$, "-", "")
                                NUEMISION$ = REPLACAR$(NUEMISION$, ".", "")
                                NUCBTE& = XVALO(NUEMISION$)
                                NUEMISION$ = CStr(NUCBTE&)
                           Do While Len(NUEMISION$) < 8: NUEMISION$ = "0" & NUEMISION$: Loop
                           
                           'NUEMISION$ = Right$(CODIKOMLAR$, 8)
                           '
                           FEMMI = CVINT(!fechconta)
                           BASEIMPON# = INEFAC#(CODIKOMLAR$, FEMMI, NC%)
                           ' Verifica si es un credito
                           SIGNO% = 0: If BASEIMPON# < 0 Then SIGNO% = 1
                           ' Agrega el signo cuando se trata de un credito
                           If SIGNO% = 1 Then Call REPLA(BASEIMPO$, "-", 1, 1)
                           '
                           IMPOXRE# = XVALO(!Idebecomp) - XVALO(!IHABECOMP)
                           IMPOXRE# = Abs(IMPOXRE#)
                           '
                           IMPOPERC$ = FORMATNUM(IMPOXRE#, "F10.2")
                           '
                           IMPOPERC$ = TRIM$(REPLACAR$(IMPOPERC$, ".", ","))
                           While Len(TRIM$(IMPOPERC$)) < 11: IMPOPERC$ = "0" + IMPOPERC$: Wend
                           ' Agrega el signo cuando se trata de un credito
                           If SIGNO% = 1 Then Call REPLA(IMPOPERC$, "-", 1, 1)
                           '
                           CUECONTABLE% = XVALO(!CUECONTAI)
                           encontro% = 0: YY$ = ""
                           For IIK& = 1 To ULTREG&("PROVINB")
                                XX$ = REGLEIDO$("PROVINB", IIK&)
                                CC& = CVI(Mid$(XX$, 27, 2))
                                If CC& = CUECONTABLE% Then
                                    encontro% = 1: YY$ = XX$
                                    Exit For
                                End If
                           Next IIK&
                           If encontro% = 1 Then
                                CODPROV_CM$ = CStr(CVI(Mid$(YY$, 37, 2)))
                                If TRIM$(CODPROV_CM$) = "" Then
                                    Call COMUNI("No Se Ha Encontrado la Provincia\para la Percepción del Comprobante '" & CODIKOMLAR$ & "'.\" _
                                                & "Dicha Percepción No Se Incluye en el Archivo de Transferencia.")
                                    GoTo 89
                                End If
                           Else
                                Call COMUNI("No Se Ha Encontrado la Jurisdicción\para la Percepción del Comprobante '" & CODIKOMLAR$ & "'.\" _
                                            & "del Proveedor '" & TRIM$(AJUSTI$(rasocli$((-1) * NC%), 32)) & "'.\ \" _
                                            & "Dicha Percepción No Se Incluye en el Archivo de Transferencia.")
                                GoTo 89
                           End If
                           '
                           DespachoImportacion$ = TRIM$(SAFETEXT$(!DespaImpo))
                           '
                           TX$ = ""
                           If DespachoImportacion = "" Then
                                Call REPLA(TX$, CODPROV_CM$, 1, 3) ' CODIGO DE PROVINCIA S/ CONVENIO MULTILATERAL
                                Call REPLA(TX$, CCTT$, 4, 13) ' CUIT
                                Call REPLA(TX$, FECPERC$, 17, 10) ' FECHA PERCEPCION
                                Call REPLA(TX$, NUMSUCURSAL$, 27, 4) ' NUMERO SUCURSAL
                                Call REPLA(TX$, NUEMISION$, 31, 8) ' NUMERO DE EMISION
                                Call REPLA(TX$, TICOMP$, 39, 1) ' TIPO DE COMPROBANTE
                                Call REPLA(TX$, LECOMP$, 40, 1) ' LETRA COMPROBANTE
                                Call REPLA(TX$, IMPOPERC$, 41, 11) ' IMPORTE DE PERCEPCION
                                CAREGI% = CAREGI% + 1
                                Print #NX1%, TX$
                           Else
                                Call REPLA(TX$, CODPROV_CM$, 1, 3) ' CODIGO DE PROVINCIA S/ CONVENIO MULTILATERAL
                                Call REPLA(TX$, CCTT$, 4, 13) ' CUIT
                                Call REPLA(TX$, FECPERC$, 17, 10) ' FECHA PERCEPCION
                                Call REPLA(TX$, DespachoImportacion$, 27, 20) ' IMPORTE DE PERCEPCION
                                Call REPLA(TX$, AJUSTD$(TRIM$(IMPOPERC$), 10), 47, 10) ' IMPORTE DE PERCEPCION
                                CAREGIADUA% = CAREGIADUA% + 1
                                Print #NX2%, TX$
                           End If

89                         .MoveNext
                         Loop
                       End With
                       Close #NX1%
                       Close #NX2%
                       ASALID11$ = REPLACAR(ASALID1$, "\", "/")
                       ASALID22$ = REPLACAR(ASALID2$, "\", "/")
                       Screen.MousePointer = 1
                       MensajeFinProceso$ = "Archivos de Transferencia\ '" + ASALID11$ + "' \Generado con " + TRIM$(Str$(CAREGI%)) + " Percepciones."
                       If CAREGIADUA% > 0 Then MensajeFinProceso = MensajeFinProceso + "\ \Además, Se Generó el Archivo '" + ASALID22$ + "'\con " + TRIM$(Str$(CAREGIADUA%)) + " Percepciones Aduaneros"
                       Call COMUNI(MensajeFinProceso$)
                 End If
            ElseIf JJ1% = 2 Then
                
                Call COMUNI("Opcional No Disponible")
            
            ElseIf JJ1% = 3 Then
         
92              OpcionDevuelta% = COMALTER%("Opciones\\Cargar de Coeficientes\Exportación SIFERE\Listado de Retenciones")
                If OpcionDevuelta% < 1 Or OpcionDevuelta% > 3 Then GoTo 85
                If OpcionDevuelta% = 1 Then
                    ' COEFICIONTES POR PERIODO
                    Call CargaCoeficientesSIRCREB
                ElseIf OpcionDevuelta% = 2 Then
                    ' EXPORTACION SIFERE
                    Call ExportaRecaudacionesBancarias
                Else
                    ' LISTADO DE RETENCIONES
                End If
            End If
            GoTo 85
        End If
     End If
   '\\\OT-05-0275-WAR-FIN///
   '\\\OT-06-0271-WAR-07/07/06///
'   If JJ0% = 2 Then 'C.I.T.I. Compras
'     Call CITICOMP
'   End If
   '\\\OT-06-0271-WAR-FIN///
   
    If JJ0% = 2 Then
        If VALIMESAN% < 1 Then GoTo 99
100     OpcionIVA% = COMALTER%("Opciones\\Percepciones\Retenciones")
        If OpcionIVA% < 1 Or OpcionIVA > 2 Then GoTo 2
        If OpcionIVA% = 1 Then
            Call GeneraArchTransfPercepIVA(Desde%, Hasta%): GoTo 100
        Else
            Call COMUNI("Opción No Disponible"): GoTo 100
        End If
    End If
   
' *************************************************************************************************

   If JJ0% = 3 Then ' RETENCION / PERCEPCIONES DE GANANCIAS
130     OP% = COMALTER%("Opciones de Trabajo:\\Subdiario de Ganancias\Archivo Transferencia")
        If OP% < 1 Or OP% > 2 Then GoTo 2

131     Call DESHASFECHA(Des%, Has%, PERIO$)
        If Des% < 1 Then GoTo 99
        DES1$ = FETEXCOR1$(Des%)
        HAS1$ = FETEXCOR1$(Has%)
        Call APERBASDAT("CABAN")

        SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
        SQLX$ = SQLX$ + "WHERE FECHCONTA >= '" & DES1$ & "'"
        SQLX$ = SQLX$ + "AND FECHCONTA <= '" & HAS1$ & "' "
        SQLX$ = SQLX$ + "AND TIPOMOVIM = 'RTGAN' "
        SQLX$ = SQLX$ + "AND NUPROVE > 0 "
        SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC "
135     On Error Resume Next
        Set PERLISTMP = READSET(SQLX$)

        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 135
        End If
        On Error GoTo 0
        '
        If OP% = 1 Then ' SUBDIARIO DE RETENCIONES GANANCIAS
            With PERLISTMP
              '
              On Error Resume Next
              .MoveFirst
              If Err Then
                Call COMUNI("No existen Retenciones en el Rango de Fechas Elegido")
                On Error GoTo 0
                GoTo 131
              End If
              On Error GoTo 0
              JJ& = 0
              ZZ$ = REGBLAN("DERETGAN")
              NIVCONXXX% = NIVCONFI%
              Do While .EOF = False
                If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 170
                Call REPLA(ZZ$, MKI(CVINT(!fechconta)), 1, 2)
                Call REPLA(ZZ$, SAFETEXT$(!CODICOM_LAR), 3, 22)
                NC% = XVALO(!NUPROVE)
                CTP$ = rasocli$((-1) * NC%)
                Call REPLA(ZZ$, CTP$, 25, 32)
                IMPER# = XVALO(!IHABECOMP) - XVALO(!Idebecomp)
                Call REPLA(ZZ$, MKD(IMPER#), 70, 8)
                CPR$ = TRIM$(Mid$(SAFETEXT$(!OpciMovim), 4))
                CCTT$ = CuitCli$((-1) * NC%)
                Call REPLA(ZZ$, CCTT$, 57, 13)
                JJ& = JJ& + 1
                Call GRAREG("DERETGAN", ZZ$, JJ&)
170           .MoveNext
              Loop
              Call SETULTREG("DERETGAN", JJ&)
              Call FILESORT("DERETGAN", "DERETGAN", 1, 1, "ASC") ' POR FECHA
              Call FINAL("*DERETGAN")
              '
            End With
            PERLISTMP.Close
            Set PERLISTMP = Nothing
        End If
        If OP% = 2 Then ' ARCHIVO DE TRANSFERENCIA DE PERECEPCIONES

             With PERLISTMP
                '
                On Error Resume Next
                .MoveFirst
                If Err Then
                  Screen.MousePointer = 1
                  Call COMUNI("No Existen Retenciones en el Rango de Fechas Elegido")
                  On Error GoTo 0
                  GoTo 131
                End If
                On Error GoTo 0

                CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
                CommonDialog1.DialogTitle = "Archivo Origen de Datos"
                CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
                CommonDialog1.FilterIndex = 1
                On Error GoTo 99
                CommonDialog1.ShowOpen
                On Error GoTo 0
   
                DoEvents
                NOMBASALID$ = UCase$(CommonDialog1.FileName)
                ASALID1$ = TRIM$(NOMBASALID$)
                If ASALID1$ = "" Then GoTo 130
   
                Screen.MousePointer = 11
  
                NX1% = FILEOPEN%(ASALID1$, 2, 128)
                CAREGI% = 0
                Do While .EOF = False
                  NC% = XVALO(!NUPROVE)
                  COMPROBANTELARGO$ = SAFETEXT$(!CODICOM_LAR)
                  CODICOMPRO$ = "05"
                  If UCase$(Left$(COMPROBANTELARGO$, 3)) = "NC-" Then CODICOMPRO$ = "03"
                  If UCase$(Left$(COMPROBANTELARGO$, 3)) = "ND-" Then CODICOMPRO$ = "04"
                  If UCase$(Left$(COMPROBANTELARGO$, 3)) = "FC-" Then CODICOMPRO$ = "01"
                  If UCase$(Left$(COMPROBANTELARGO$, 3)) = "OP-" Then CODICOMPRO$ = "06"
                  If UCase$(Left$(COMPROBANTELARGO$, 3)) = "RB-" Then CODICOMPRO$ = "02"
                  '
                  FECHA% = CVINT(!fechconta)
                  FECHARET$ = FECHATEXLAR$(FECHA%)
                  '
                  NUMECOMPRO$ = Left(SAFETEXT$(!NumeComp), 16)
                  '
                  CONDI$ = "NUPROVE=" & NC% & " AND CODICOM_LAR='" & COMPROBANTELARGO$ & "' AND TIPOMOVIM='OPAG'"
                  IMPORTECOMPROBANTE# = XVALO(VALOBADA("CAJABANC", "VALABSCOMP", CONDI$))
                  '
                  CODIGOIMP& = XVALO(!CodigoImpuesto)
                  CODIGOREG& = XVALO(!CodigoRegimen)
                  If CODIGOIMP& = 0 Or CODIGOREG& = 0 Then
                    RECORD$ = FILTERGETRECORD$("DARGAPRO", 1, MKI$(NC%))
                    CATIB1& = Asc(Mid$(RECORD$, 3, 1))
                    REGISTROCATE$ = REGLEIDO$("TACRETG", CATIB1&)
                    CODIGOIMP& = CVI(Mid$(REGISTROCATE$, 51, 2))
                    CODIGOREG& = CVI(Mid$(REGISTROCATE$, 53, 2))
                  End If
                  '
                  CodigoOperacion = 1
                  '
                  BASECALCULO# = XVALO(VALOBADA("CAJABANC", "BASEIMPO", CONDI$))
                  '
                  FechaRetencion$ = FECHARET$
                  '
                  POSIV% = PIVACLI((-1) * NC%)
                  If POSIV% = 1 Then
                    CODIGOCONDICION% = 1
                  ElseIf POSIV% = 2 Then
                    CODIGOCONDICION% = 2
                  Else
                    CODIGOCONDICION% = 3
                  End If
                  '
                  SUJETOSUSPENDIDO$ = Space(1)
                  '
                  ImporteRetencion# = XVALO(!IHABECOMP) - XVALO(!Idebecomp)
                  '
                  PORCENTAJEEXCLUSION$ = Space(6)
                  '
                  FECHABOLETIN$ = Space(10)
                  '
                  TIPODOCUMENTO$ = "80"
                  '
                  Cuit$ = CuitCli$((-1) * NC%)
                  Cuit$ = REPLACAR$(Cuit$, "-", "")
                  '
                  CERTIFICADOORIGINAL$ = Space(14)
                  '
                  DENOMINACIONORDENANTE$ = Space(30)
                  '
                  ACRECENTAMIENTO$ = Space(1)
                  '
                  CUITPAISRETENIDO$ = Space(11)
                  '
                  CUITDELORDENANTE$ = Space(11)
                  '
                  TX$ = Space(197)
                  Call REPLA(TX$, CODICOMPRO$, 1, 2)
                  Call REPLA(TX$, FECHARET$, 3, 10)
                  Call REPLA(TX$, NUMECOMPRO$, 13, 16)
                  Call REPLA(TX$, REPLACAR$(FORMATNUM$(IMPORTECOMPROBANTE#, "F16.2"), ".", ","), 29, 16)
                  Call REPLA(TX$, FORMATNUM$(CODIGOIMP&, "I3"), 45, 3)
                  Call REPLA(TX$, FORMATNUM$(CODIGOREG&, "I3"), 48, 3)
                  Call REPLA(TX$, FORMATNUM$(CodigoOperacion, "I1"), 51, 1)
                  Call REPLA(TX$, REPLACAR$(FORMATNUM$(BASECALCULO#, "F14.2"), ".", ","), 52, 14)
                  Call REPLA(TX$, FechaRetencion$, 66, 10)
                  Call REPLA(TX$, FORMATNUM$(CODIGOCONDICION%, "I2"), 76, 2)
                  Call REPLA(TX$, SUJETOSUSPENDIDO$, 78, 1)
                  Call REPLA(TX$, REPLACAR$(FORMATNUM$(ImporteRetencion#, "F14.2"), ".", ","), 79, 14)
                  Call REPLA(TX$, PORCENTAJEEXCLUSION$, 93, 6)
                  Call REPLA(TX$, FECHABOLETIN$, 99, 10)
                  Call REPLA(TX$, TIPODOCUMENTO$, 109, 2)
                  Call REPLA(TX$, Cuit$, 111, 20)
                  Call REPLA(TX$, CERTIFICADOORIGINAL$, 131, 14)
                  Call REPLA(TX$, DENOMINACIONORDENANTE$, 145, 30)
                  Call REPLA(TX$, ACRECENTAMIENTO$, 175, 1)
                  Call REPLA(TX$, CUITPAISRETENIDO$, 176, 11)
                  Call REPLA(TX$, CUITDELORDENANTE$, 187, 11)

                  CAREGI% = CAREGI% + 1
                  Print #NX1%, TX$
                  .MoveNext
                Loop
             End With
             Close #NX1%
             ASALID2$ = REPLACAR(ASALID1$, "\", "/")
             Screen.MousePointer = 1
             Call COMUNI("Archivos de Transferencia\ '" + ASALID2$ + "' \Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.")
        End If
        GoTo 2
   End If
   
   If JJ0% = 4 Then
        If VALIMESAN% < 1 Then GoTo 99
        Call GENIVACOM
        Call RegimenInformativoCompras
        GoTo 2
   End If
  
99 Command11.Enabled = True
   On Error GoTo 0

End Sub

Sub GeneraArchTransfPercepIVA(Des%, Has%)

            'SINTESIS
            'SE RECORRE EL CAJABANC SEGUN RANGO SELECCIONADO
            'Y QUE LAS CUENTAS QUE SE CONFIGURAN EN EL SETUP DE FACTURACION DE PROVEEDORES '
            'VALIDA QUE ESTE CONFIGURADA LA CUENTA CONTABLE, Y EL CODIGO DE REGIEMEN
            'RECORRE EL CAJABANC Y TOMA LOS IMPORTES DE PERCEPCION DE IVA (VARIMOVI = 'Per-IVA')
            'ESTOS SON LOS QUE SE CARGAN MANUALMENTE EN EL CASILLERO PERCEPCION DE IVA
            '
            CUENTAPERIVA$ = CONTROL$("", "CUEIVA-" + "5")
            CUE% = CUECOINT%(CUENTAPERIVA$)
            If CUE% < 1 Then
               Call COMUNI("No Se Encuentra Configurada La Cuenta Contable de Percepción de I.V.A.\Configurar desde el Setup de Facturación de Proveedores.")
               Exit Sub
            End If
            'VALIDO SI EN EL DAT PROVINB ESTAN CARGADAS LAS CUENTAS CONTABLES
            CODREG& = XVALO(CONTROL$("COREPER", "COREPER"))
            If CODREG& < 1 Then
               Call COMUNI("No Se Ha Configurado el Código de Régimen Para las Percepciones de I.V.A..\Configurar desde el Setup de Facturación de Proveedores.")
               Exit Sub
            End If
            '
            Campos$ = "Documento/A20;Prov-Cte./A26;Cuit/a13;Fecha/d8;Informe/A64"
            Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
            Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , ANTOT)
            FRMZELECHO.Width = 12000
            FRMZELECHO.Caption = "Informe de Errores"

            DES1$ = FETEXCOR1$(Des%)
            HAS1$ = FETEXCOR1$(Has%)
            ARCHIX$ = SELECCION_Archivo$("TXT", CommonDialog1)
            If TRIM$(ARCHIX$) = "" Then Exit Sub
            
            Screen.MousePointer = 11
            NARCHI2% = FILEOPEN%(ARCHIX$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
            GENERO_ARCHIVO% = 0: I1& = 0
            '
            SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
            SQLX$ = SQLX$ + "WHERE FECHCONTA >= '" & DES1$ & "'"
            SQLX$ = SQLX$ + "AND FECHCONTA <= '" & HAS1$ & "' "
            SQLX$ = SQLX$ + "AND ((TIPOMOVIM = 'IVA' AND VARIMOVIM = 'Per-IVA' AND NUPROVE>0)"
            SQLX$ = SQLX$ + "OR (TIPOMOVIM='GASTC'))"
            SQLX$ = SQLX$ + " AND CUECONTAI = " & CUE%
            SQLX$ = SQLX$ + " ORDER BY FECHEMISI ASC "
            Set RS = READSET(SQLX$)
            With RS
              FILAS& = 0: j& = 0
              Do While Not .EOF
                 FACTURAX$ = TRIM$(SAFETEXT$(!CODICOM_LAR))
                 NUMEFAC = XVALO(!NumeComp): NUMEFAX$ = SAFETEXT$(NUMEFAC)
                 PREFIJOX$ = Mid$(FACTURAX$, 6, 4)
                 NC = XVALO(!NUPROVE)
                 If NC = 0 Then
                    CuentaContable% = XVALO(VALOBADA("cajabanc", "cuecontai", "codicom_lar='" & FACTURAX$ & "' and tipomovim='CONCITC' and idebecomp>0"))
                    NC = XVALO(VALOBADA("proved12", "nume_cli", "cuma_cli=" & CStr(CuentaContable%)))
                    PREFIJOX$ = "0000": NUMEFAC = XVALO(!CODICOM_COR)
                 End If
                 
                 CUITX$ = TRIM$(CuitCli(-1 * NC))
                 FE% = CVINT(!FECHEMISI)
                 FEMIS$ = Format(CVDAT(FE%), "DD/MM/YYYY")
                 MONTOPERCEP# = XVALO(!Idebecomp) - XVALO(!IHABECOMP)
                 
                 'validaciones
                 TXTERROR$ = ""
                 Faltax$ = "Falta: "
                 If CUITX$ = "" Then TXTERROR$ = Faltax$ + "C.U.I.T., ": Faltax$ = ""
                 If FACTURAX$ = "" Then TXTERROR$ = TXTERROR$ + Faltax$ + "N° de Factura, ": Faltax$ = ""
                 If PREFIJOX$ = "" Then TXTERROR$ = TXTERROR$ + Faltax$ + "Prefijo de Factura, ": Faltax$ = ""
                 NC = NC * -1
                 RACLI$ = NC & " - " & rasocli$(NC)
                 If TXTERROR$ <> "" Then
                   j& = j& + 1
                   FRMZELECHO.msf2.Rows = j& + 1
                   FRMZELECHO.msf2.TextMatrix(j&, 1) = DOCUMENTO$
                   FRMZELECHO.msf2.TextMatrix(j&, 2) = RACLI$
                   FRMZELECHO.msf2.TextMatrix(j&, 3) = CUITX$
                   FRMZELECHO.msf2.TextMatrix(j&, 4) = FECHATEX$(FE%)
                   TXTERROR$ = TRIM$(TXTERROR$)
                   POS1% = InStrRev(TXTERROR$, ",")
                   If POS1% > 0 Then
                       TXTERROR$ = Left$(TXTERROR$, POS1% - 1)
                   End If
                   FRMZELECHO.msf2.TextMatrix(j&, 5) = TXTERROR$
                 End If
                 'fin validacion

                    GENERO_ARCHIVO% = 1
                    CUECR$ = CODREG&
                    ' *** VERIFICAR SI EL PROVEEDOR ESTA CONFIGURADO COMO PROVEEDOR DE IMPORTACION (ADUANA AFIP O DEL TIPO FEDEX)
                    CondicionSql$ = "campo2='" & CStr(Abs(NC)) & "' and lista='PERCIVAPROIMPO'"
                    If CantRegistros("datasql", CondicionSql$) > 0 Then CUECR$ = "267"
                    '
                     If Len(CUECR$) < 3 Then CUECR$ = String$(3, "0") & CUECR$: CUECR$ = Right(CUECR$, 3)
                    Call REPLA(TX$, CUECR$, 1, 3)
                     If Len(CUITX$) < 13 Then CUITX$ = String$(13, " ") & CUITX$: CUITX$ = Right(CUITX$, 13)
                    Call REPLA(TX$, CUITX$, 4, 13)
                     If Len(FEMIS$) < 10 Then FEMIS$ = String$(10, "0") & FEMIS$: FEMIS$ = Right(FEMIS$, 10)
                    Call REPLA(TX$, FEMIS$, 17, 10)
                     If Len(PREFIJOX$) < 8 Then PREFIJOX$ = String$(8, "0") & PREFIJOX$: PREFIJOX$ = Right(PREFIJOX$, 8)
                    Call REPLA(TX$, PREFIJOX$, 27, 8)
                      If Len(NUMEFAX$) < 8 Then NUMEFAX$ = String$(8, "0") & NUMEFAX$: NUMEFAX$ = Right(NUMEFAX$, 8)
                    Call REPLA(TX$, NUMEFAX$, 35, 8)
                   
                    MONTOPERCEPX$ = TRIM$(FORMATNUM$(MONTOPERCEP#, "F16.2"))
                    'FORMATEO SI TIENE DECIMALES HAY QUE QUITARLE EL PUNTO
                    If InStr(MONTOPERCEPX$, ".") > 0 Then
                      MONTOPERCEPX$ = REPLACAR(MONTOPERCEPX$, ".", ",")
                    Else
                      MONTOPERCEPX$ = MONTOPERCEPX$ & ",00"
                    End If
                    If Len(MONTOPERCEPX$) < 16 Then MONTOPERCEPX$ = String$(16, "0") & MONTOPERCEPX$: MONTOPERCEPX$ = Right(MONTOPERCEPX$, 16)
                    Call REPLA(TX$, MONTOPERCEPX$, 43, 16)
                    FILAS& = FILAS& + 1
                    Print #NARCHI2%, TX$
                 
                 .MoveNext
              Loop
            End With
            'Close #NARCHI2%
            '
            If j& > 0 Then
                 Call COMUNI("Imposible Procesar Archivo de Transferencia\Al Aceptar Este Mensaje\Se Visualizará el Detalle de Los Mismos")
                 Screen.MousePointer = 1
                 FRMZELECHO.Show 1
                 Close #NARCHI2%
                 GoTo 40
            End If

'            'AHORA CARGO LOS ASIENTOS
'             Call CargaAsientosPERCEPCIONIVA(DES%, HAS%, ARCHIX$, OK%, RET_TX$, RET_FILAS&) 'RET_TX$ ES UN PARAMETRO DE RETORNO
'             If OK% < 1 Then
'                Close #NARCHI2%
'                GoTo 40
'             Else
'               Print #NARCHI2%, RET_TX$
'            End If

30          If GENERO_ARCHIVO% > 0 Or OK% >= 0 Then 'JX& SI ES > 0 ES POR QUE REGISTRO ALGUN ERROR
               'AHORA AGREGA INFORMACION DE ASIENTOS, SI VIENE CON OK = 1 TODO BIEN, SI NO TRAE -1 Y BORRA EL ARCVHIVO
'                If OK% < 1 Then
'                  GoTo 40
'               Else
                 'Print #NARCHI2%, RET_TX$
'               End If
               ARCHIX$ = REPLACAR$(ARCHIX$, "\", "/")
               Screen.MousePointer = 1
               If RET_FILAS& + FILAS& > 0 Then
                 Call COMUNI("Archivo Generado Correctamente en " & ARCHIX$ & " Con " & RET_FILAS& + FILAS& & " Registros")
               End If
                 Close #NARCHI2%
            Else
                Call COMUNI("Sin Registros Que Cumplan Con la Condición")
                Close #NARCHI2%
40              On Error Resume Next
                Call FILEKILL(ARCHIX$)
                If Err > 0 Then
                 Call COMUNI("No Se Ha Podido Eliminar el Archivo:\" & ARCHIX$ & "\Realice Las Correcciones Informadas en el Detalle y Vuelva a Intentar")
                End If
            End If
99          Screen.MousePointer = 1


'            'SINTESIS
'            'SE RECORRE EL CAJABANC SEGUN RANGO SELECCIONADO
'            'Y QUE LAS CUENTAS QUE SE CONFIGURAN EN EL SETUP DE FACTURACION DE PROVEEDORES '
'            'VALIDA QUE ESTE CONFIGURADA LA CUENTA CONTABLE, Y EL CODIGO DE REGIEMEN
'            'RECORRE EL CAJABANC Y TOMA LOS IMPORTES DE PERCEPCION DE IVA (VARIMOVI = 'Per-IVA')
'            'ESTOS SON LOS QUE SE CARGAN MANUALMENTE EN EL CASILLERO PERCEPCION DE IVA
'            '
'            CUENTAPERIVA$ = CONTROL$("", "CUEIVA-" + "5")
'            CUE% = CUECOINT%(CUENTAPERIVA$)
'            If CUE% < 1 Then
'               Call COMUNI("No Se Encuentra Configurada La Cuenta Contable de Percepción de I.V.A.\Configurar desde el Setup de Facturación de Proveedores.")
'               Exit Sub
'            End If
'            'VALIDO SI EN EL DAT PROVINB ESTAN CARGADAS LAS CUENTAS CONTABLES
'            CODREG& = XVALO(CONTROL$("COREPER", "COREPER"))
'            If CODREG& < 1 Then
'               Call COMUNI("No Se Ha Configurado el Código de Régimen Para las Percepciones de I.V.A..\Configurar desde el Setup de Facturación de Proveedores.")
'               Exit Sub
'            End If
'            '
'            Campos$ = "Documento/A20;Prov-Cte./A26;Cuit/a13;Fecha/d8;Informe/A64"
'            Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
'            Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , ANTOT)
'            FRMZELECHO.Width = 12000
'            FRMZELECHO.Caption = "Informe de Errores"
'
'            DES1$ = FETEXCOR1$(Des%)
'            HAS1$ = FETEXCOR1$(Has%)
'            ARCHIX$ = SELECCION_Archivo$("TXT", CommonDialog1)
'            If TRIM$(ARCHIX$) = "" Then Exit Sub
'
'            Screen.MousePointer = 11
'            NARCHI2% = FILEOPEN%(ARCHIX$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
'            GENERO_ARCHIVO% = 0: I1& = 0
'            '
'            SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
'            SQLX$ = SQLX$ + "WHERE FECHEMISI >= '" & DES1$ & "'"
'            SQLX$ = SQLX$ + "AND FECHEMISI <= '" & HAS1$ & "' "
'            SQLX$ = SQLX$ + "AND TIPOMOVIM = 'IVA' "
'            SQLX$ = SQLX$ + "AND VARIMOVIM = 'Per-IVA' "
'            SQLX$ = SQLX$ + " AND CUECONTAI = " & CUE%
'            SQLX$ = SQLX$ + " ORDER BY FECHEMISI ASC "
'            Set rs = READSET(SQLX$)
'            With rs
'              FILAS& = 0: J& = 0
'              Do While Not .EOF
'                 FACTURAX$ = TRIM$(SAFETEXT$(!CODICOM_LAR))
'                 NUMEFAC = XVALO(!NumeComp): NUMEFAX$ = SAFETEXT$(NUMEFAC)
'                 PREFIJOX$ = Mid$(FACTURAX$, 6, 4)
'                 NC = XVALO(!NUPROVE)
'                 If NC = 0 Then D
'
'                 CUITX$ = TRIM$(CuitCli(-1 * NC))
'                 FE% = CVINT(!FECHEMISI)
'                 FEMIS$ = Format(CVDAT(FE%), "DD/MM/YYYY")
'                 MONTOPERCEP# = XVALO(!Idebecomp) - XVALO(!IHABECOMP)
'
'                 'validaciones
'                 TXTERROR$ = ""
'                 Faltax$ = "Falta: "
'                 If CUITX$ = "" Then TXTERROR$ = Faltax$ + "C.U.I.T., ": Faltax$ = ""
'                 If FACTURAX$ = "" Then TXTERROR$ = TXTERROR$ + Faltax$ + "N° de Factura, ": Faltax$ = ""
'                 If PREFIJOX$ = "" Then TXTERROR$ = TXTERROR$ + Faltax$ + "Prefijo de Factura, ": Faltax$ = ""
'                 NC = NC * -1
'                 RACLI$ = NC & " - " & rasocli$(NC)
'                 If TXTERROR$ <> "" Then
'                   J& = J& + 1
'                   FRMZELECHO.msf2.Rows = J& + 1
'                   FRMZELECHO.msf2.TextMatrix(J&, 1) = Documento$
'                   FRMZELECHO.msf2.TextMatrix(J&, 2) = RACLI$
'                   FRMZELECHO.msf2.TextMatrix(J&, 3) = CUITX$
'                   FRMZELECHO.msf2.TextMatrix(J&, 4) = FECHATEX$(FE%)
'                   TXTERROR$ = TRIM$(TXTERROR$)
'                   POS1% = InStrRev(TXTERROR$, ",")
'                   If POS1% > 0 Then
'                       TXTERROR$ = Left$(TXTERROR$, POS1% - 1)
'                   End If
'                   FRMZELECHO.msf2.TextMatrix(J&, 5) = TXTERROR$
'                 End If
'                 'fin validacion
'
'                    GENERO_ARCHIVO% = 1
'                    CUECR$ = CODREG&
'                    ' *** VERIFICAR SI EL PROVEEDOR ESTA CONFIGURADO COMO PROVEEDOR DE IMPORTACION (ADUANA AFIP O DEL TIPO FEDEX)
'                    CondicionSql$ = "campo2='" & CStr(Abs(NC)) & "' and lista='PERCIVAPROIMPO'"
'                    If CantRegistros("datasql", CondicionSql$) > 0 Then CUECR$ = "267"
'                    '
'                     If Len(CUECR$) < 3 Then CUECR$ = String$(3, "0") & CUECR$: CUECR$ = Right(CUECR$, 3)
'                    Call REPLA(TX$, CUECR$, 1, 3)
'                     If Len(CUITX$) < 13 Then CUITX$ = String$(13, " ") & CUITX$: CUITX$ = Right(CUITX$, 13)
'                    Call REPLA(TX$, CUITX$, 4, 13)
'                     If Len(FEMIS$) < 10 Then FEMIS$ = String$(10, "0") & FEMIS$: FEMIS$ = Right(FEMIS$, 10)
'                    Call REPLA(TX$, FEMIS$, 17, 10)
'                     If Len(PREFIJOX$) < 8 Then PREFIJOX$ = String$(8, "0") & PREFIJOX$: PREFIJOX$ = Right(PREFIJOX$, 8)
'                    Call REPLA(TX$, PREFIJOX$, 27, 8)
'                      If Len(NUMEFAX$) < 8 Then NUMEFAX$ = String$(8, "0") & NUMEFAX$: NUMEFAX$ = Right(NUMEFAX$, 8)
'                    Call REPLA(TX$, NUMEFAX$, 35, 8)
'
'                    MONTOPERCEPX$ = TRIM$(FORMATNUM$(MONTOPERCEP#, "F16.2"))
'                    'FORMATEO SI TIENE DECIMALES HAY QUE QUITARLE EL PUNTO
'                    If InStr(MONTOPERCEPX$, ".") > 0 Then
'                      MONTOPERCEPX$ = REPLACAR(MONTOPERCEPX$, ".", ",")
'                    Else
'                      MONTOPERCEPX$ = MONTOPERCEPX$ & ",00"
'                    End If
'                    If Len(MONTOPERCEPX$) < 16 Then MONTOPERCEPX$ = String$(16, "0") & MONTOPERCEPX$: MONTOPERCEPX$ = Right(MONTOPERCEPX$, 16)
'                    Call REPLA(TX$, MONTOPERCEPX$, 43, 16)
'                    FILAS& = FILAS& + 1
'                    Print #NARCHI2%, TX$
'
'                 .MoveNext
'              Loop
'            End With
'            'Close #NARCHI2%
'            '
'            If J& > 0 Then
'                 Call COMUNI("Imposible Procesar Archivo de Transferencia\Al Aceptar Este Mensaje\Se Visualizará el Detalle de Los Mismos")
'                 Screen.MousePointer = 1
'                 FRMZELECHO.Show 1
'                 Close #NARCHI2%
'                 GoTo 40
'            End If
'
''            'AHORA CARGO LOS ASIENTOS
''             Call CargaAsientosPERCEPCIONIVA(DES%, HAS%, ARCHIX$, OK%, RET_TX$, RET_FILAS&) 'RET_TX$ ES UN PARAMETRO DE RETORNO
''             If OK% < 1 Then
''                Close #NARCHI2%
''                GoTo 40
''             Else
''               Print #NARCHI2%, RET_TX$
''            End If
'
'30          If GENERO_ARCHIVO% > 0 Or OK% >= 0 Then 'JX& SI ES > 0 ES POR QUE REGISTRO ALGUN ERROR
'               'AHORA AGREGA INFORMACION DE ASIENTOS, SI VIENE CON OK = 1 TODO BIEN, SI NO TRAE -1 Y BORRA EL ARCVHIVO
''                If OK% < 1 Then
''                  GoTo 40
''               Else
'                 'Print #NARCHI2%, RET_TX$
''               End If
'               ARCHIX$ = REPLACAR$(ARCHIX$, "\", "/")
'               Screen.MousePointer = 1
'               If RET_FILAS& + FILAS& > 0 Then
'                 Call COMUNI("Archivo Generado Correctamente en " & ARCHIX$ & " Con " & RET_FILAS& + FILAS& & " Registros")
'               End If
'                 Close #NARCHI2%
'            Else
'                Call COMUNI("Sin Registros Que Cumplan Con la Condición")
'                Close #NARCHI2%
'40              On Error Resume Next
'                Call FILEKILL(ARCHIX$)
'                If Err > 0 Then
'                 Call COMUNI("No Se Ha Podido Eliminar el Archivo:\" & ARCHIX$ & "\Realice Las Correcciones Informadas en el Detalle y Vuelva a Intentar")
'                End If
'            End If
'99          Screen.MousePointer = 1

End Sub


Sub CargaAsientosPERCEPCIONIVA(Des%, Has%, ARCHIX$, OK%, RET_TX$, RET_FILAS&)
    RET_TX$ = "": RET_FILAS& = 0
    DES1$ = FETEXCOR1$(Des%)
    HAS1$ = FETEXCOR1$(Has%)
'      1.-Cod.Jurisdicción
'      2.-Cuit
'      3.-F.Asiento
'      4.-ASIENTO
'      5.-PREFIJO
'      6.-NUFACTURA
'      7.-DEBE
'      8.-HABER
'      9.-NPROV
'      10.-DH
    OK% = -1
    Campos$ = "Cod.Jurisdicción/A0;Cuit/A0;F.Asiento/A0;Asiento/A20;Prefijo/A64;N°Factura/F0;Haber/F0;Prov/F0;Fallo/A0;DH/A0;Fallo/A0;Resultado/A0"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = 12000
    FRMZELECHO.Caption = "Asientos No Conformes Para Emisión de Archivo de Transferencia"
    
    CUECONTA$ = CONTROL$("", "CUEIVA-" + "5")
    CUE% = CUECOINT%(CUECONTA$)
    If CUE% < 1 Then
      Call COMUNI("Falta Configurar Cuenta Contable de Percepción de I.V.A.\Configurar la Misma Desde Setupde Facturas de Proveedores")
      OK% = -1
      GoTo 99
    End If
    
    CODREG& = XVALO(CONTROL$("COREPER", "COREPER"))
    If CODREG& < 1 Then
      OK% = -1
      Call COMUNI("Falta Configurar Código de Percepción de I.V.A.\Configurar la Misma Desde Setup de Facturas de Proveedores")
      GoTo 99
    End If
    '
'    CODREGXTC = XVALO(CONTROL("OPCHEQ07", "CRETIVAT"))
'    If CODREGXTC < 1 Then
'      OK% = -1
'      Call COMUNI("Falta Configurar Código de Retención de I.V.A. de Tarjetas de Crédito\Configurar la Misma Desde Setup de Gestión de Valores")
'      GoTo 99
'    End If

    JX& = 0
    'RECORRO EL DETAASIEN SEGUN CONDICION
    SQLX$ = "SELECT * FROM DETAASIEN WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE FECHASIEN >= '" & DES1$ & "'"
    SQLX$ = SQLX$ + "AND FECHASIEN <= '" & HAS1$ & "' "
    SQLX$ = SQLX$ + " AND CUECOINT = " & CUE%
    SQLX$ = SQLX$ + " AND STATPASE >=0 "
    SQLX$ = SQLX$ + " ORDER BY NUIDASIEN, FECHASIEN ASC "
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        NUMEAS$ = SAFETEXT$(!nuidasien)
        IMPODEBE1 = XVALO(!idebepase)
        IMPOHABE1 = XVALO(!ihabepase)
        FE% = CVINT(!FECHASIEN)
        CUE% = XVALO(!CUECOINT)
        ASIENTO$ = SAFETEXT$(!nuidasien)
        TTXX$ = SAFETEXT$(!DATRETPER)
        Call TEXTOLOTES(TTXX$, ";")
        PREFI$ = "": NUFACTURA$ = ""
        On Error Resume Next
        For i& = 1 To UBound(CADENATEX$)
          Select Case i&
          Case 2
            PREFI$ = CADENATEX$(i&)
          Case 3
            NUFACTURA$ = CADENATEX$(i&)
          End Select
        Next i&
        On Error GoTo 0

        JX& = JX& + 1
'        COMPROBAX$ = Mid$(TTXX$, 1, 1)
'        CERTIFICADOX$ = Mid$(TTXX$, 2, 20)
'        LETRAX$ = Mid$(TTXX$, 22, 10)
        'CARGO LA INFORMACION EN LA GRILLA
        FRMZELECHO.msf2.Rows = JX& + 1
        FRMZELECHO.msf2.TextMatrix(JX&, 1) = CODREG&
        FRMZELECHO.msf2.TextMatrix(JX&, 2) = "" ' CUIT
        FRMZELECHO.msf2.TextMatrix(JX&, 3) = FECHATEX$(FE%)
        FRMZELECHO.msf2.TextMatrix(JX&, 4) = ASIENTO$
        FRMZELECHO.msf2.TextMatrix(JX&, 5) = PREFI$
        FRMZELECHO.msf2.TextMatrix(JX&, 6) = NUFACTURA$
         If IMPODEBE1 > 0 Then DH$ = "D"
         If IMPOHABE1 > 0 Then DH$ = "H"
        FRMZELECHO.msf2.TextMatrix(JX&, 7) = FORMATNUM$(IMPODEBE1, "F12.2")
        FRMZELECHO.msf2.TextMatrix(JX&, 8) = FORMATNUM$(IMPOHABE1, "F12.2")
        FRMZELECHO.msf2.TextMatrix(JX&, 9) = "" 'NPROV
        FRMZELECHO.msf2.TextMatrix(JX&, 10) = DH$ 'SI EL IMPORTE ESTA DEL LADO DEL DEBE O DEL HABER
        .MoveNext
      Loop
      End With
      '
      'RECORRER LA GRILLA Y TOMAR LOS OTROS VALORES
      'AHORA RECORRO LA GRILLA
      JX& = 0: FALLO% = 0
      With FRMZELECHO
       For JX& = 1 To .msf2.Rows - 1
          ASIENTO$ = FRMZELECHO.msf2.TextMatrix(JX&, 4)
          DH$ = FRMZELECHO.msf2.TextMatrix(JX&, 10)
          If DH$ = "D" Then
             IMPORTEX# = XVALO(FRMZELECHO.msf2.TextMatrix(JX&, 7))
             CONDI$ = "NUIDASIEN = '" & ASIENTO$ & "' AND IHABEPASE > 0 AND STATPASE >=0 "
          Else
             IMPORTEX# = XVALO(FRMZELECHO.msf2.TextMatrix(JX&, 8))
             CONDI$ = "NUIDASIEN = '" & ASIENTO$ & "' AND IDEBEPASE > 0  AND STATPASE >=0  "
          End If
          'VALIDO QUE HAYA UNA SOLA CUENTA DE CONTRAPARTIDA
          If CantRegistros&("DETAASIEN", CONDI$, "NOMESS") = 1 Then
             CUE% = VALOBADA("DETAASIEN", "CUECOINT", CONDI$, "NOMESS")
             NPROV = PROV_SegCtaCble(CUE%)
                CUITX$ = CuitCli(-1 * NPROV)
                Valid% = VALICUIT%(CUITX$)
                If Valid% < 1 Then
                  FRMZELECHO.msf2.TextMatrix(JX&, 12) = "FALLO"
                   FALLO% = 1
                End If
                CUITX$ = FormatearCuitValido$(CUITX$) ' LE AGREGA FORMATO CON LOS SEPARADORES - EN CASO QUE NO LOS TENGA
                FRMZELECHO.msf2.TextMatrix(JX&, 2) = CUITX$
                FRMZELECHO.msf2.TextMatrix(JX&, 9) = NPROV
           Else
             FALLO% = 1
             FRMZELECHO.msf2.TextMatrix(JX&, 11) = "MAS DE 1"
             FRMZELECHO.msf2.TextMatrix(JX&, 12) = "FALLO"
           End If
           If TRIM$(PREFI$) = "" Then FRMZELECHO.msf2.TextMatrix(JX&, 12) = "FALLO"
           If TRIM$(NUFACTURA$) = "" Then FRMZELECHO.msf2.TextMatrix(JX&, 12) = "FALLO"
       Next JX&
      End With
      'VALIDA SI HAY FALLOS MUESTRA LOS FALLOS
      If FALLO% > 0 Then
        OK% = -1
        Dim AXIENTO$(), ERRORX$()
        FRMZELECHO.msf2.TextMatrix(0, 4) = "Asiento"
        FRMZELECHO.msf2.TextMatrix(0, 5) = "Informe": j& = 0
        With FRMZELECHO
          For JX1& = 1 To .msf2.Rows - 1
           
            If TRIM$(.msf2.TextMatrix(JX1&, 12)) <> "FALLO" And .msf2.TextMatrix(JX1&, 11) <> "MAS DE 1" Then
              Call cmdEliminaItem(.msf2, JX1&)
              JX1& = JX1& - 1
              If JX1& >= .msf2.Rows - 1 Then Exit For
            Else
                
                CUITX$ = TRIM$(FRMZELECHO.msf2.TextMatrix(JX1&, 2))
                ASIENTO$ = FRMZELECHO.msf2.TextMatrix(JX1&, 4)
                PREFI$ = TRIM$(FRMZELECHO.msf2.TextMatrix(JX1&, 5))
                NUFACTURA$ = TRIM$(FRMZELECHO.msf2.TextMatrix(JX1&, 6))
                Valid% = VALICUIT%(CUITX$)
                NPROV = FRMZELECHO.msf2.TextMatrix(JX1&, 9)
                CONTRAPARTIDA$ = FRMZELECHO.msf2.TextMatrix(JX1&, 11)
                'CARGO LA INFORMACION EN LA GRILLA
                Faltax$ = "Falta: ": TXTERROR$ = ""
                If PREFI$ = "" Then TXTERROR$ = Faltax$ + "Prefijo, ": Faltax$ = ""
                If NUFACTURA$ = "" Then TXTERROR$ = TXTERROR$ + Faltax$ + "N° de Factura, ": Faltax$ = ""
                If Valid% < 1 Then TXTERROR$ = TXTERROR$ + Faltax$ + "C.U.I.T , ": Faltax$ = ""
                If NPROV < 1 Then TXTERROR$ = TXTERROR$ + Faltax$ + "Cta. Proveedor , ": Faltax$ = ""
                'la siguiente fila debe estar ultima
                If CONTRAPARTIDA$ = "MAS DE 1" Then TXTERROR$ = TXTERROR$ + "Más de 1 Cuenta de Contrapartida": Faltax$ = ""
                POS1% = InStrRev(TXTERROR$, ",")
                If POS1% > 0 Then
                   TXTERROR$ = Left$(TXTERROR$, POS1% - 1)
                End If
                '
                j& = j& + 1
                ReDim Preserve AXIENTO$(j&), ERRORX$(j&)
                AXIENTO$(j&) = ASIENTO$
                ERRORX$(j&) = TXTERROR$
            End If
          Next JX1&
          On Error Resume Next
          'VUELVO SOLO LOS ERRORES  EN LA GRILLA.
          For i& = 1 To UBound(AXIENTO$)
             .msf2.TextMatrix(i&, 4) = AXIENTO$(i&)
             .msf2.TextMatrix(i&, 5) = ERRORX$(i&)
          Next i&
          Erase AXIENTO$, ERRORX$
          On Error GoTo 0
          
          Call COMUNI("Imposible Procesar Archivo de Transferencia\Al Aceptar Este Mensaje\Se Visualizará el Detalle de Los Mismos")
          .Show 1
          GoTo 99
        End With
      Else 'SI NO HAY FALLOS AGREGA AL ARCHIVO
        With FRMZELECHO
          For JX& = 1 To .msf2.Rows - 1
            
            If RET_FILAS = 0 Then OK% = 1
            CUPRO$ = .msf2.TextMatrix(JX&, 1)
            CUITX$ = .msf2.TextMatrix(JX&, 2)
            
            FE% = FECHANUM(FRMZELECHO.msf2.TextMatrix(JX&, 3))
              FEMIS$ = Format(CVDAT(FE%), "DD/MM/YYYY")
            PREFI$ = TRIM$(.msf2.TextMatrix(JX&, 5))
            NUFACTU$ = TRIM$(.msf2.TextMatrix(JX&, 6))
            
            IMPORTEAPLIC# = 0
            DH$ = FRMZELECHO.msf2.TextMatrix(JX&, 10)
            If DH$ = "D" Then IMPORTEAPLIC# = XVALO(.msf2.TextMatrix(JX&, 7))
            If DH$ = "H" Then IMPORTEAPLIC# = XVALO(.msf2.TextMatrix(JX&, 8))

            If Len(CUPRO$) < 3 Then CUPRO$ = String$(3, "0") & CUPRO$: CUPRO$ = Right(CUPRO$, 3)
            Call REPLA(TX$, CUPRO$, 1, 3)
             If Len(CUITX$) < 13 Then CUITX$ = String$(13, " ") & CUITX$: CUITX$ = Right(CUITX$, 13)
            Call REPLA(TX$, CUITX$, 4, 13)
             If Len(FEMIS$) < 10 Then FEMIS$ = String$(10, "0") & FEMIS$: FEMIS$ = Right(FEMIS$, 10)
            Call REPLA(TX$, FEMIS$, 17, 10)
              If Len(PREFI$) < 8 Then PREFI$ = String$(8, "0") & PREFI$: PREFI$ = Right(PREFI$, 8)
            Call REPLA(TX$, Left(PREFI$, 8), 27, 8)
              If Len(NUFACTU$) < 8 Then NUFACTU$ = String$(8, "0") & NUFACTU$: NUFACTU$ = Right(NUFACTU$, 8)
            Call REPLA(TX$, NUFACTU$, 35, 8)
            
            IMPORTEAPL$ = TRIM$(FORMATNUM$(IMPORTEAPLIC#, "F16.2"))
            'FORMATEO SI TIENE DECIMALES HAY QUE QUITARLE EL PUNTO
            If InStr(IMPORTEAPL$, ".") > 0 Then
              IMPORTEAPL$ = REPLACAR(IMPORTEAPL$, ".", ",")
            Else
              IMPORTEAPL$ = IMPORTEAPL$ & ",00"
            End If
            If Len(IMPORTEAPL$) < 16 Then IMPORTEAPL$ = String$(16, "0") & IMPORTEAPL$: IMPORTEAPL$ = Right(IMPORTEAPL$, 16)
            Call REPLA(TX$, IMPORTEAPL$, 43, 16)
            RET_TX$ = RET_TX$ + TX$ + vbCrLf 'CARGO EN UNA VARIABLE EL TEXTO QUE VA DE RETORNO A LA LLAMADA A ESTA RUTINA.
            RET_FILAS& = RET_FILAS& + 1
            'LE QUITO EL UTLIMO CARACTER (ENTER)
            A = Len(RET_TX$)
          Next JX&
          POS1% = InStrRev(RET_TX$, Chr(13))
          If POS1% > 0 Then
             RET_TX$ = Left$(RET_TX$, POS1% - 1)
          End If

        End With
      End If
      
   'HASTA ACA
'      1.-Cod.Jurisdicción
'      2.-Cuit
'      3.-F.Asiento
'      4.-ASIENTO
'      5.-PREFIJO
'      6.-NUFACTURA
'      7.-DEBE
'      8.-HABER
'      9.-NPROV
'      10.-DH
      
99
End Sub




Private Sub RegimenInformativoCompras()
    
    On Error Resume Next
    Call TRALOCAL("LIVACOM", "")
    On Error GoTo 0
    
    If Err > 0 Then
        Call MENSERR(24, "Imposible Copiar Archivo con Datos del Período")
        Exit Sub
    End If
    
'    If InStr(1, EMPREAC$, "ASEMBLI") > 0 Then
'        MODOASEMBLI% = 1
'    End If
    
    PERIODO$ = TRIM$(MESAN)
    ' borrar archivo RIV
    NX% = FILEOPEN%(LUCOM$ & "REGINFO_CV_COMPRAS_CBTE_" & PERIODO$ & ".TXT", 2, 266)
    NX1% = FILEOPEN%(LUCOM$ & "REGINFO_CV_COMPRAS_ALICUOTAS_" & PERIODO$ & ".TXT", 2, 62)
    NX2% = FILEOPEN%(LUCOM$ & "REGINFO_CV_COMPRAS_IMPORTACIONES_" & PERIODO$ & ".TXT", 2, 62)

    UltimoNumeroComprobante$ = ""
    UltimoCuitProveedor$ = ""
    
    Dim VGD_FechaComprobante$(3)
    Dim VGD_TipoCbte$(3)
        VGD_TipoCbte$(1) = TipoComprobante("FC-0000-00000000-B")
        VGD_TipoCbte$(2) = TipoComprobante("ND-0000-00000000-B")
        VGD_TipoCbte$(3) = TipoComprobante("NC-0000-00000000-B")
    Dim VGD_PuntoVta$(3)
    Dim VGD_NumeroCbtDesde$(3)
    Dim VGD_NumeroCbtHasta$(3)
    VGD_CodigoDocComprador$ = "99"
    VGD_NumeroIdentificacionComprador$ = RIVFormatoNumero$(0, 20)
    VGD_RazonSocial$ = AJUSTI$("VENTA GLOBAL DIARIA", 30)
    Dim VGD_ImporteTotal$(3)
    Dim VGD_ImporteNoGravado$(3)
    VGD_PercepcionNoCategorizados$ = RIVFormatoNumero$(0, 15)
    Dim VGD_ImporteOperacionExenta$(3)
    Dim VGD_ImportePercepcionesNacionales$(3)
    Dim VGD_ImportePercepcionIB$(3)
    VGD_ImportePercepcionMunicipales$ = RIVFormatoNumero$(0, 15)
    VGD_ImporteImpuestosInternos$ = RIVFormatoNumero$(0, 15)
    VGD_CodMoneda$ = CodigoMoneda(1)
    VGD_TipoCambio$ = RIVFormatoNumero$(1, 10, 6)
    Dim VGD_CantidadAlicuotas$(3)
    VGD_CodigoOperacion$ = " "
    VGD_OtrosTributos$ = RIVFormatoNumero$(0, 15)
    VGD_FechaVencimiento$ = "00000000"
    Dim VGD_RegistroRIV$(3)
    Dim INDICE%
    
    Dim RegistroRIV_Alicuotas$
    Dim RegistroRIV_Importaciones$
    Dim FlagError%
    Dim ListaDeErrores$
    Dim EsRegimenSimplificado%
    
    INDICE% = 1: GoSub InicializarCamposVGD
    INDICE% = 2: GoSub InicializarCamposVGD
    INDICE% = 3: GoSub InicializarCamposVGD
    
    FIN& = ULTREG&("!LIVACOM")
    If FIN& = 0 Then GoTo 99
    
    Screen.MousePointer = 11
    
    For i& = 1 To ULTREG&("!LIVACOM")
        Call TRACE(24, i&, ULTREG&("!LIVACOM"))
        
        REGISTRO$ = REGLEIDO$("!LIVACOM", i&)
        NumeroComprobante$ = Mid$(REGISTRO$, 11, 18)
        
        If Mid$(NumeroComprobante$, 1, 2) = "TC" Then
            NUMMOVITARJETA$ = Mid$(NumeroComprobante$, 4, 6) & Mid$(NumeroComprobante$, 11, 2)
            NumeroComprobante$ = "CO-A-0001-" & NUMMOVITARJETA$
        End If

        CUITPROVEEDOR$ = TRIM$(Mid$(REGISTRO$, 63, 16))
        If TRIM$(NumeroComprobante$) = "" Then
            NumeroComprobante$ = UltimoNumeroComprobante$
            CUITPROVEEDOR$ = UltimoCuitProveedor$
        End If
       
        If (NumeroComprobante$ <> UltimoNumeroComprobante$) Or ((NumeroComprobante$ = UltimoNumeroComprobante$) And (CUITPROVEEDOR$ <> UltimoCuitProveedor$)) Then
            '
            RegistroAnterior$ = REGISTRO$
            
            FechaComprobante$ = FECHATEXLAR$(CVI(Mid$(REGISTRO$, 3, 2)))
            FechaComprobante$ = Mid$(FechaComprobante$, 7, 4) & Mid$(FechaComprobante$, 4, 2) & Mid$(FechaComprobante$, 1, 2)

            If UltimoNumeroComprobante$ <> "" And FlagError% = 0 Then
                If TRIM$(RegistroRIV$) <> "" Then
                    Print #NX%, RegistroRIV$
                    RegistroRIV$ = ""
                    If RegistroRIV_IvaAduana$ <> "" Then
                        Print #NX1%, RegistroRIV_IvaAduana$
                        RegistroRIV_IvaAduana$ = ""
                    End If
                End If
            End If
            FlagError% = 0
            EsRegimenSimplificado% = 0

25          fechaCbte$ = FETEXCOR1$(CVI(Mid$(REGISTRO$, 3, 2)))
            NumeroProveedor% = CVI(Mid$(REGISTRO$, 31, 2)) * (-1)
            If NumeroProveedor% = 0 Then
                ListaDeErrores$ = ListaDeErrores$ & "No Se Pudo Encontrar Numero de Proveedor para el Comprobante '" & NumeroComprobante$ & "'"
                FlagError% = 1: GoTo 35
            End If

            TipoCbte$ = TipoComprobante$(NumeroComprobante$)
            
            PUNTOVTA$ = PuntoVenta$(NumeroComprobante$)
            ' Es obligatorio que el punto de venta sea > 0
            If PUNTOVTA$ = "00000" Then PUNTOVTA$ = "00001"

            NumeroCbt$ = NumeroCbte$(NumeroComprobante$)
                             
            ' Verifica si se trata de una factura de importación común o simplificada
                Despa$ = SAFETEXT$(VALOBADA("CAJABANC", "DespaImpo", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%), "NOMESS"))
            NumeroDespachoImportacion$ = UCase(AJUSTI$(Despa$, 16))
            
            ' *** 33693450239 ES EL NUMERO DE CUIT DE AFIP
            If InStr(1, NumeroDespachoImportacion$, "PART") > 0 Then NumeroDespachoImportacion$ = AJUSTI$("", 16)
            EsFacturaDespachante# = 0
            EsFacturaDespachante# = XVALO(VALOBADA("CAJABANC", "ValAbsComp", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%) & " AND Varimovim='Aduana'"))
            If (EsFacturaDespachante# > 0.005 And TRIM$(NumeroDespachoImportacion$) = "") Then
                Respuesta% = COMALTER%("Se Ha Detectado Comprobante '" & NumeroComprobante$ & "' Con Iva Importación y Sin Despacho.\" & _
                                       "¿Corresponde a una Importación Simplificada?\\No, Cancelar\Si, Continuar")
                If Respuesta% = 2 Then
                    EsRegimenSimplificado% = 1
                Else
                    ListaDeErrores$ = ListaDeErrores$ & "Comprobante c/Iva Aduana sin Despacho en '" & NumeroComprobante$ & "', Proveedor '" & CStr(Abs(NumeroProveedor%)) & "'" & vbCrLf
                    FlagError% = 1: GoTo 35
                End If
            ElseIf TRIM$(NumeroDespachoImportacion$) <> "" Then
                TipoCbte$ = "066": PUNTOVTA$ = "00000": NumeroCbt$ = "00000000000000000000"
            End If

            CodigoDocComprador$ = CodigoDocumento$(PIVACLI%(NumeroProveedor%))
                        
            'NumeroIdentificacionComprador$ = SAFETEXT$(VALOBADA("CAJABANC", "Cuit", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%), "NOMESS"))
            'If Len(NumeroIdentificacionComprador$) < 11 Or VALICUIT%(NumeroIdentificacionComprador$) = 0 Then NumeroIdentificacionComprador$ = TRIM$(Mid$(REGISTRO$, 63, 16))
            'If NumeroIdentificacionComprador$ = "" Or VALICUIT%(NumeroIdentificacionComprador$) = 0 Then
            '    NumeroIdentificacionComprador$ = CUITCLI$(NumeroProveedor%)
            'End If
            NumeroIdentificacionComprador$ = TRIM$(Mid$(REGISTRO$, 63, 16))
            If NumeroIdentificacionComprador$ = "" Then
                ListaDeErrores$ = ListaDeErrores$ & "CUIT Inválido en el Proveedor '" & CStr(Abs(NumeroProveedor%)) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                FlagError% = 1: GoTo 35
            End If
            If PIVACLI%(NumeroProveedor%) <> 5 And PIVACLI%(NumeroProveedor%) <> 0 Then
                If VALICUIT%(NumeroIdentificacionComprador$) = 0 Then
                    ListaDeErrores$ = ListaDeErrores$ & "CUIT Inválido en el Proveedor '" & CStr(Abs(NumeroProveedor%)) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                    FlagError% = 1: GoTo 35
                End If
            End If
            NumeroIdentificacionComprador$ = SOLONUM$(NumeroIdentificacionComprador$)
            NumeroIdentificacionComprador$ = AJUSTD$(NumeroIdentificacionComprador$, 20)
            NumeroIdentificacionComprador$ = REPLACAR(NumeroIdentificacionComprador$, " ", "0")

            ' Toma la razon social desde la base de datos. Prevé la posibilidad de que
            ' se trate de proveedores varios
            razonSocial$ = AJUSTI$(SAFETEXT$(VALOBADA("CAJABANC", "RefeCom", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%), "NOMESS")), 30)
            If TRIM$(razonSocial$) = "" Then razonSocial$ = AJUSTI$(rasocli$(NumeroProveedor%), 30)
            LargoRS% = InStr(1, razonSocial$, "//")
            If LargoRS% > 0 Then razonSocial$ = AJUSTI$(Mid(razonSocial$, 1, LargoRS% - 1), 30)
            If PIVACLI%(NumeroProveedor%) = 0 Then
                razonSocial$ = AJUSTI$("CONSUMIDOR FINAL", 30)
            End If
            
            ImporteTotal$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 91, 8)), 15)
            
            ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
            ValorIva$ = TRIM$(Mid$(REGISTRO$, 129, 8))
            LetraComprobante$ = Mid$(NumeroComprobante$, 4, 1)
            If InStr(1, UCase(ValorIva$), "NO GRAV") > 0 And (LetraComprobante$ <> "B" And LetraComprobante$ <> "C") Then
                ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
            End If
            If InStr(1, "011;012;013;006;007;008;019;020;021", TipoCbte$) > 0 Then ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
            
            PercepcionNoCategorizados$ = RIVFormatoNumero$(0, 15)

            ImporteOperacionExenta$ = RIVFormatoNumero$(0, 15)
            If PIVACLI%(NumeroProveedor%) = 5 Or PIVACLI%(NumeroProveedor%) = 7 Then
                ImporteOperacionExenta$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
            End If
            
            ImportePercepcionPagosACuenta$ = RIVFormatoNumero$(CVD(Mid$(REGISTRO$, 161, 8)), 15)
            
            ' Percepciones de iva + ganancias
                PercepcionesVarias# = CVD(Mid$(REGISTRO$, 169, 8))
            ImportePercepcionesNacionales$ = RIVFormatoNumero$(PercepcionesVarias#, 15)
            
            ImportePercepcionIB$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 193, 8)), 15)
            
            ImportePercepcionMunicipales$ = RIVFormatoNumero$(0, 15)
            
            ImporteImpuestosInternos$ = RIVFormatoNumero$(0, 15)
                
                MONE% = XVALO(VALOBADA("CAJABANC", "Mone_Emis", "CODICOM_LAR='" & NumeroComprobante$ & "' AND FechConta='" & fechaCbte$ & "'"))
            CODMONEDA$ = CodigoMoneda(MONE%)
            If CODMONEDA$ = "000" Then CODMONEDA$ = "PES"
            
                TCAMBIO# = XVALO(VALOBADA("CAJABANC", "Tipo_Cam", "CODICOM_LAR='" & NumeroComprobante$ & "' AND FechConta='" & fechaCbte$ & "'"))
            TIPOCAMBIO$ = RIVFormatoNumero$(TCAMBIO#, 10, 6)
            If CODMONEDA$ = "PES" And TCAMBIO# > 1 Then
                    ListaDeErrores$ = ListaDeErrores$ & "Error en Tipo de Cambio en el Proveedor '" & CStr(Abs(NumeroProveedor%)) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                    FlagError% = 1: GoTo 35
            End If
            If XVALO(TIPOCAMBIO$) = 0 And CODMONEDA$ = "PES" Then TIPOCAMBIO$ = RIVFormatoNumero(1, 10, 6)

                CANTALICUOTAS% = 1
            If InStr(1, "006;007;008;009;082;013;012;011;015;013;111", TipoCbte$) > 0 Then CANTALICUOTAS% = 0 ' Si es Comprobante "C" la cantidad debe ser cero
            CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
            
            Percepciones# = XVALO(ImportePercepcionesNacionales$) / 100 + XVALO(ImportePercepcionIB$) / 100 + XVALO(ImportePercepcionPagosACuenta$) / 100
            If PIVACLI%(NumeroProveedor%) = 7 Then
                CodigoOperacion$ = "Z"
            ElseIf PIVACLI%(NumeroProveedor%) = 5 Or (TRIM$(NumeroDespachoImportacion$) <> "" Or EsFacturaDespachante# > 0.005) Then
                CodigoOperacion$ = "X"
            ElseIf InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 Or (TRIM$(Mid$(REGISTRO$, 129, 8)) = "" And Percepciones# > 0.005) Then
                CodigoOperacion$ = "N"
            ElseIf PIVACLI%(NumeroProveedor%) = 4 Or InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "EXENTO") > 0 Then
                CodigoOperacion$ = "E"
            ElseIf CANTALICUOTAS% = 0 Then
                CodigoOperacion$ = "0"
            Else
                CodigoOperacion$ = " "
            End If
                
            ImporteIva$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
            
            OtrosTributos$ = RIVFormatoNumero$(0, 15)
            
            CuitEmisor$ = "00000000000"
            
            DenominacionEmisor$ = AJUSTI$(" ", 30)
            
            IvaComision$ = RIVFormatoNumero("0", 15)
            
            ' Arma el registro para el regimen informativo
            RegistroRIV$ = FechaComprobante$ & TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & NumeroDespachoImportacion$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & razonSocial$ _
                               & ImporteTotal$ & ImporteNoGravado$ & ImporteOperacionExenta$ & ImportePercepcionPagosACuenta$ & ImportePercepcionesNacionales$ _
                               & ImportePercepcionIB$ & ImportePercepcionMunicipales$ & ImporteImpuestosInternos$ & CODMONEDA$ & TIPOCAMBIO$ & CantidadAlicuotas$ _
                               & CodigoOperacion$ & ImporteIva$ & OtrosTributos$ & CuitEmisor$ & DenominacionEmisor$ & IvaComision$
            Debug.Print RegistroRIV$
            UltimoNumeroComprobante$ = NumeroComprobante$
            UltimoCuitProveedor$ = CUITPROVEEDOR$

            'If (InStr(1, UCase(ValorIva$), "NO GRAV") = 0 And XVALO(ImporteIva$) > 0) Then
            If XVALO(CantidadAlicuotas$) > 0 Then
            
                ' Tratamiento para los comprobantes con parte no gravada o exenta
                If TRIM$(CodigoOperacion$) <> "" Then
                    ImporteNeto$ = RIVFormatoNumero(0, 15)
                    ALICUOTA$ = "0003"
                    ImporteIva0$ = RIVFormatoNumero(0, 15)

                    If (CodigoOperacion$ = "N") Then
                        ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                        Percepciones# = XVALO(ImportePercepcionesNacionales$) / 100 + XVALO(ImportePercepcionIB$) / 100 + XVALO(ImportePercepcionPagosACuenta$) / 100
                        ImporteEx# = XVALO(VALOBADA("CAJABANC", "Sum(BaseImpo)", "VariMovim like 'Exento' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                        If ConIva# < 0.01 Or (Percepciones# > 0.005 And ConIva# < 0.01) Or (ImporteEx# > 0.005 And ConIva# < 0.01) Then
                            ' cuando el comprobante es totalmente no gravado, dicho importe se escribe en el archivo de alicuotas
                            ImporteNeto$ = ImporteTotal$
                            If Percepciones# > 0.005 Or ImporteEx# > 0.005 Then ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                            ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
                            Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                        Else
                            ' Cuando el comprobante tiene parte gravada y parte que no lo esta, el gravado se escribe
                            ' en el archivo de comprobantes y los importes gravados en el de alicuotas
                            CANTALICUOTAS% = 0
                            CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                            Call REPLA(RegistroRIV$, " ", 239, 1)
                            Call REPLA(RegistroRIV$, CantidadAlicuotas$, 238, 1)
                            GoTo 35
                        End If
                    End If
                    If CodigoOperacion$ = "E" Then
                        ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                        Percepciones# = XVALO(ImportePercepcionesNacionales$) / 100 + XVALO(ImportePercepcionIB$) / 100 + XVALO(ImportePercepcionPagosACuenta$) / 100
                        NoGravado# = XVALO(VALOBADA("CAJABANC", "Sum(BaseImpo)", "VariMovim like 'No Grav%' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                        If (ConIva# < 0.005 Or (Percepciones# > 0.005 And ConIva# < 0.01)) And NoGravado# < 0.005 Then
                            ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                            ImporteOperacionExenta$ = RIVFormatoNumero$(0, 15)
                            Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 135, 15)
                        ElseIf NoGravado# > 0.005 Then
                            ImporteOperacionExenta$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                            Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 135, 15)
                            CANTALICUOTAS% = 0
                            CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                            If Mid$(NumeroComprobante$, 1, 2) = "ND" Or Mid$(NumeroComprobante$, 1, 2) = "FC" Then
                                ImporteNeto$ = RIVFormatoNumero(0, 15)
                                ImporteIva0$ = RIVFormatoNumero(0, 15)
                                ALICUOTA$ = "0003"
                                CANTALICUOTAS% = 1
                                CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                            Else
                                GoTo 35
                            End If
                        Else
                            CANTALICUOTAS% = 0
                            CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                            Call REPLA(RegistroRIV$, " ", 239, 1)
                            Call REPLA(RegistroRIV$, CantidadAlicuotas$, 238, 1)
                            ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                            Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                            GoTo 35
                        End If
                    End If
                    If CodigoOperacion$ = "X" And (TRIM$(NumeroDespachoImportacion$) <> "" Or EsFacturaDespachante# > 0.005) Then
                        If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") > 0 Then
                            ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                            ImpoNoGra# = XVALO(VALOBADA("CAJABANC", "Sum(BaseImpo)", "VariMovim='Aduana' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                            ImpoAduana# = XVALO(ImporteIva$) / 100
                            OtrosTributos$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
                            ImporteIva$ = RIVFormatoNumero(0, 15)
                            Call REPLA(RegistroRIV$, ImporteIva$, 240, 15)
                            CantidadAlicuotas$ = "0"
                            If ImpoNoGra# > 0.005 Then
                                ImporteNGAnterior = XVALO(Mid$(RegistroRIV$, 120, 15)) / 100
                                ImpoNoGra# = ImpoNoGra# + ImporteNGAnterior
                                ImporteNoGravado$ = RIVFormatoNumero(ImpoNoGra#, 15, , 1)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                            End If
                            If (ConIva# < 0.005 And EsRegimenSimplificado% = 0) Then
                                Call REPLA(RegistroRIV$, RIVFormatoNumero(ImpoAduana#, 15), 150, 15)
                                ImporteNeto$ = RIVFormatoNumero(0, 15)
                                ImporteIva0$ = RIVFormatoNumero(0, 15)
                                ALICUOTA$ = "0003"
                                CantidadAlicuotas$ = "1"
                            Else
                                If MODOASEMBLI% > 0 And EsRegimenSimplificado% = 1 Then
                                    RegistroRIV_IvaAduana$ = String$(65, "0") & "0003" & RIVFormatoNumero(ImpoAduana#, 15)
                                End If
                                GoTo 35
                            End If
                        ElseIf InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "EXENTO") > 0 Then
                            'ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                            ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                            If ConIva# > 0.005 Then
                                ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                CANTALICUOTAS% = 0
                                CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                                GoTo 35
                            Else
                                ImporteNoGravado$ = RIVFormatoNumero(0, 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                            End If
                        ElseIf InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 Then
                            ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                            If ConIva# > 0.005 Then
                                ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                CANTALICUOTAS% = 0
                                CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                                GoTo 35
                            Else
                                ImporteNoGravado$ = RIVFormatoNumero(0, 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                            End If
                        ElseIf InStr(1, "0004;0005;0006;0008;0009", CodigoAlicuota(Mid$(REGISTRO$, 129, 8))) > 0 Then
                            ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                            ImporteIva0$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
                            ALICUOTA$ = CodigoAlicuota(ValorIva$)
                        End If
'                        If CodigoOperacion$ = "X" Then
'                            If CantRegistros("CAJABANC", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%) & " AND IDEBECOMP>0 AND TIPOMOVIM='COMPRA'") > 1 Then
'                                CANTALICUOTAS% = 0
'                                CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
'                                Call REPLA(RegistroRIV$, " ", 239, 1)
'                                Call REPLA(RegistroRIV$, CantidadAlicuotas$, 238, 1)
'                                GoTo 35
'                            Else
'                                ImporteOperacionExenta$ = RIVFormatoNumero(0, 15)
'                                Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 135, 15)
'                                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
'                            End If
'                        End If
                    ElseIf CodigoOperacion$ = "X" Then
'                        If CantRegistros("CAJABANC", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%) & " AND IDEBECOMP>0 AND TIPOMOVIM='COMPRA'") > 1 Then
'                                CANTALICUOTAS% = 0
'                                CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
'                                Call REPLA(RegistroRIV$, " ", 239, 1)
'                                Call REPLA(RegistroRIV$, CantidadAlicuotas$, 238, 1)
'                                GoTo 35
'                        Else
                                ImporteOperacionExenta$ = RIVFormatoNumero(0, 15)
                                Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 135, 15)
                                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
'                        End If
                    End If
                Else
                    ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                    ImporteIva0$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
                    ALICUOTA$ = CodigoAlicuota(ValorIva$)
                End If
                
                If (TRIM$(NumeroDespachoImportacion$) = "" And EsFacturaDespachante# < 0.005) Or EsRegimenSimplificado% = 1 Then
                
                    RegistroRIV_Alicuotas$ = TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & ImporteNeto$ & ALICUOTA$ & ImporteIva0$
                    If MODOASEMBLI% = 1 And EsRegimenSimplificado% = 1 Then
                        RegistroRIV_IvaAduana$ = RegistroRIV_Alicuotas$
                    Else
                        Print #NX1%, RegistroRIV_Alicuotas$
                    End If
                Else
                    
                    RegistroRIV_Importaciones$ = NumeroDespachoImportacion$ & ImporteNeto$ & ALICUOTA$ & ImporteIva0$
                    Print #NX2%, RegistroRIV_Importaciones$
                End If
            End If
        Else
            If FlagError% = 1 Then GoTo 35
            If TipoCbte$ = "006" Then GoTo 35
            
            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 Then
                ImporteNGAnterior = XVALO(Mid$(RegistroRIV$, 120, 15)) / 100
                ImporteNGActual = CVD(Mid$(REGISTRO$, 137, 8))
                'If Not (ImporteNGActual < 0 And Mid$(REGISTRO$, 11, 2) = "FC") Then
                If Mid$(REGISTRO$, 11, 2) <> "NC" Then
                    'If A <> 1 Then ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior + ImporteNGActual, 15, , 1)
                    ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior + ImporteNGActual, 15, , 1)
                Else
                    ConsideraSigno% = 0
                    If ImporteNGActual > 0 Then ' INCORPORADO PARA EL VALOR CIF
                        ImporteNGActual = (-1) * ImporteNGActual
                        ConsideraSigno% = 1
                    End If
                    ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior + ImporteNGActual, 15, , ConsideraSigno%)
                End If
                'End If
            End If
            
            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "EXENTO") > 0 Then
                ImporteExAnterior = XVALO(Mid$(RegistroRIV$, 135, 15)) / 100
                ImporteExActual = CVD(Mid$(REGISTRO$, 137, 8))
                ImporteOperacionExenta$ = RIVFormatoNumero(ImporteExAnterior + ImporteExActual, 15)
            End If
            
            If PIVACLI%(NumeroProveedor%) = 5 Or PIVACLI%(NumeroProveedor%) = 7 Then
                IMPO# = XVALO(Mid$(RegistroRIV$, 135, 15)) / 100
                ImporteOperacionExenta$ = RIVFormatoNumero(IMPO# + CVD(Mid$(REGISTRO$, 185, 8)), 15)
            End If
                                                                             
                IMPO# = XVALO(Mid$(RegistroRIV$, 150, 15)) / 100
            ImportePercepcionPagosACuenta$ = RIVFormatoNumero$(IMPO# + CVD(Mid$(REGISTRO$, 161, 8)), 15)
            
                IMPO# = XVALO(Mid$(RegistroRIV$, 165, 15)) / 100
            ImportePercepcionesNacionales$ = RIVFormatoNumero$(IMPO# + CVD(Mid$(REGISTRO$, 169, 8)), 15)

                IMPO# = XVALO(Mid$(RegistroRIV$, 180, 15)) / 100
            ImportePercepcionIB$ = RIVFormatoNumero(IMPO# + CVD(Mid$(REGISTRO$, 193, 8)), 15)
                                
            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") = 0 Then
                    IMPO# = XVALO(Mid$(RegistroRIV$, 240, 15)) / 100
                    ImporteIvaActual# = Abs(CVD(Mid$(REGISTRO$, 145, 8)))
                    If A = 0 Then
                        ImporteIva$ = RIVFormatoNumero(IMPO# + ImporteIvaActual#, 15)
                    Else
                        If ImporteIvaActual# > 0.005 Then
                            IMPO# = XVALO(Mid$(RegistroRIV$, 150, 15)) / 100
                            ImportePercepcionPagosACuenta$ = RIVFormatoNumero(IMPO# + ImporteIvaActual#, 15)
                        End If
                    End If
            End If

                CANTALICUOTAS% = XVALO(CantidadAlicuotas$)
            ImporteNeto0# = CVD(Mid$(REGISTRO$, 137, 8))
            ValorIva$ = TRIM$(REPLACAR$(Mid$(REGISTRO$, 129, 8), "%", ""))
            CantidadAlicuotaNuevoRegistro% = 0
            If (Not (ImporteNeto0# < 0 And Mid$(REGISTRO$, 11, 2) = "FC")) And XVALO(ImporteIvaActual#) > 0.005 And InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") = 0 Then CantidadAlicuotaNuevoRegistro% = 1
'            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 And (CodigoOperacion$ = "X" Or CodigoOperacion$ = "E") And _
'                CantRegistros("CAJABANC", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%) & " AND IDEBECOMP>0 AND (TIPOMOVIM='COMPRA' or TIPOMOVIM='REDCOM')") < 2 Then
'
'                CantidadAlicuotaNuevoRegistro% = 1
'                ImporteNGAnterior = XVALO(ImporteNoGravado$) / 100
'                ImporteNGActual = CVD(Mid$(REGISTRO$, 137, 8))
'                ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior - ImporteNGActual, 15)
'            End If
            CantidadAlicuotas$ = CStr(CANTALICUOTAS% + CantidadAlicuotaNuevoRegistro%)
            
            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") > 0 Then
'                IMPO# = XVALO(Mid$(RegistroRIV$, 255, 15)) / 100
'                OtAtributos# = CVD(Mid$(REGISTRO$, 145, 8))
'                OTROSTRIBUTOS$ = RIVFormatoNumero(IMPO# + OtAtributos#, 15)
'                IMPO# = XVALO(ImporteIva$) / 100
'                ImporteIva$ = RIVFormatoNumero(IMPO# - OtAtributos#, 15)
                IMPO# = XVALO(ImportePercepcionPagosACuenta$) / 100
                ImporteAduana# = CVD(Mid$(REGISTRO$, 145, 8))
                If MODOASEMBLI% = 0 Or (MODOASEMBLI% = 1 And EsRegimenSimplificado% = 0) Then
                    ImportePercepcionPagosACuenta$ = RIVFormatoNumero(IMPO# + ImporteAduana#, 15)
                End If
                ' SI NO TIENE ALICUOTA PREVIA, AGREGA UN REGISTRO EN EL ARCHIVO DE ALICUOTA. ES FICTICIO PORQUE
                ' TANTO EL NETO GRAVADO COMO EL IMPORTE DE IVA, ESTAN EN CERO.
                ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                If ConIva# < 0.005 And (MODOASEMBLI% = 0 Or (MODOASEMBLI% = 1 And EsRegimenSimplificado% = 0)) Then
                    CantidadAlicuotaNuevoRegistro% = 1
                    CantidadAlicuotas$ = CStr(CANTALICUOTAS% + CantidadAlicuotaNuevoRegistro%)
                End If
                
                If MODOASEMBLI% = 1 And EsRegimenSimplificado% = 1 And RegistroRIV_IvaAduana$ <> "" Then
                    CantidadAlicuotaNuevoRegistro% = 1
                End If
            End If
                        
            RegistroRIV$ = FechaComprobante$ & TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & NumeroDespachoImportacion$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & razonSocial$ _
                           & ImporteTotal$ & ImporteNoGravado$ & ImporteOperacionExenta$ & ImportePercepcionPagosACuenta$ & ImportePercepcionesNacionales$ _
                           & ImportePercepcionIB$ & ImportePercepcionMunicipales$ & ImporteImpuestosInternos$ & CODMONEDA$ & TIPOCAMBIO$ & CantidadAlicuotas$ _
                           & CodigoOperacion$ & ImporteIva$ & OtrosTributos$ & CuitEmisor$ & DenominacionEmisor$ & IvaComision$
            
            If (CantidadAlicuotaNuevoRegistro% > 0) Then
                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                ImporteIva0$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
                ALICUOTA$ = CodigoAlicuota(ValorIva$)
                If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 And (CodigoOperacion$ = "X" Or CodigoOperacion$ = "E") Then
                    ALICUOTA$ = "0003"
                End If
                If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") > 0 And (MODOASEMBLI% = 0 Or (MODOASEMBLI% = 1 And EsRegimenSimplificado% = 0)) Then
                    ImporteNeto$ = RIVFormatoNumero(0, 15)
                    ImporteIva0$ = RIVFormatoNumero(0, 15)
                End If
                If A = 1 Then
                    ALICUOTA$ = "0003"
                    ImporteNeto$ = RIVFormatoNumero(0, 15)
                    ImporteIva0$ = RIVFormatoNumero(0, 15)
                End If
                
                If (TRIM$(NumeroDespachoImportacion$) = "" And EsFacturaDespachante# < 0.005) Or EsRegimenSimplificado% = 1 Then
                
                    RegistroRIV_Alicuotas$ = TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & ImporteNeto$ & ALICUOTA$ & ImporteIva0$
                    If MODOASEMBLI% = 1 And EsRegimenSimplificado% = 1 Then
                        If RegistroRIV_IvaAduana$ <> "" Then        ' EL IMPORTE DEL IVA ADUANA SE AGREGA AL CREDITO FISCAL DE ALGUNA DE LAS ALICUOTAS DEL COMPROBANTE
                            If Mid$(RegistroRIV_IvaAduana$, 66, 4) = "0003" Then
                                ImpoAduana# = XVALO(Mid$(RegistroRIV_IvaAduana$, 70, 15)) / 100
                                IMPOIVA# = XVALO(Mid$(RegistroRIV_Alicuotas$, 70, 15)) / 100
                                IMPOIVA# = IMPOIVA# + ImpoAduana#
                                RegistroRIV_Alicuotas$ = Mid$(RegistroRIV_Alicuotas$, 1, 69) & RIVFormatoNumero(IMPOIVA#, 15)
                            ElseIf ALICUOTA$ = "0003" Then
                                ImpoAduana# = XVALO(Mid$(RegistroRIV_IvaAduana$, 70, 15)) / 100
                                IMPOIVA# = XVALO(Mid$(RegistroRIV_Alicuotas$, 70, 15)) / 100
                                IMPOIVA# = IMPOIVA# + ImpoAduana#
                                RegistroRIV_Alicuotas$ = Mid$(RegistroRIV_IvaAduana$, 1, 69) & RIVFormatoNumero(IMPOIVA#, 15)
                            End If
                            RegistroRIV_IvaAduana$ = ""
                        Else
                            If RegistroRIV_IvaAduana$ = "" Then
                                RegistroRIV_IvaAduana$ = RegistroRIV_Alicuotas$
                                GoTo 35
                            End If
                        End If
                    End If
                    Print #NX1%, RegistroRIV_Alicuotas$
                Else
                    
                    RegistroRIV_Importaciones$ = NumeroDespachoImportacion$ & ImporteNeto$ & ALICUOTA$ & ImporteIva0$
                    Print #NX2%, RegistroRIV_Importaciones$
                End If
            End If
        End If
35  Next i&
    
    If NumeroComprobante$ <> "" Then Print #NX%, RegistroRIV$
    
    Close #NX%
    Close #NX1%
    Close #NX2%
    
    Screen.MousePointer = 1
    
'    Dim XX1 As String * 266
'
'    ARCHIIBB$ = "REGINFO_CV_COMPRAS_CBTE_" & PERIODO$ & ".TXT"
'    NARCHI% = FILEOPEN%(LUCOM$ + ARCHIIBB$, 1, 266)
'    Do While Not EOF(NARCHI%)
'        Line Input #NARCHI%, XX$
'        SIGNO% = 1
'        If Mid$(XX$, 9, 3) = "003" Or Mid$(XX$, 9, 3) = "008" Or Mid$(XX$, 9, 3) = "021" Or Mid$(XX$, 9, 3) = "013" Or Mid$(XX$, 9, 3) = "053" Then SIGNO% = -1
'
'        IMPO1# = IMPO1# + SIGNO% * XVALO(Mid$(XX$, 105, 15)) / 100
'        impo2# = impo2# + SIGNO% * XVALO(Mid$(XX$, 120, 15)) / 100
'        impo3# = impo3# + SIGNO% * XVALO(Mid$(XX$, 150, 15)) / 100
'        impo4# = impo4# + SIGNO% * XVALO(Mid$(XX$, 135, 15)) / 100
'        impo5# = impo5# + SIGNO% * XVALO(Mid$(XX$, 165, 15)) / 100
'        impo6# = impo6# + SIGNO% * XVALO(Mid$(XX$, 180, 15)) / 100
'        impo7# = impo7# + SIGNO% * XVALO(Mid$(XX$, 195, 15)) / 100
'        impo8# = impo8# + SIGNO% * XVALO(Mid$(XX$, 210, 15)) / 100
'        impo9# = impo9# + SIGNO% * XVALO(Mid$(XX$, 255, 15)) / 100
'        impo10# = impo10# + SIGNO% * XVALO(Mid$(XX$, 240, 15)) / 100
'20  Loop
'
'    MsgBox ("Total: " & TRIM$(FORMATNUM$(IMPO1#, "F15.2")) & vbCrLf _
'        & "No gravado: " & TRIM$(FORMATNUM$(impo2#, "F15.2")) & vbCrLf _
'        & "Iva: " & TRIM$(FORMATNUM$(impo10#, "F15.2")) & vbCrLf _
'        & "Percep. Iva: " & TRIM$(FORMATNUM$(impo3#, "F15.2")) & vbCrLf _
'        & "Op.Exentas: " & TRIM$(FORMATNUM$(impo4#, "F15.2")) & vbCrLf _
'        & "Imp.Nacionales: " & TRIM$(FORMATNUM$(impo5#, "F15.2")) & vbCrLf _
'        & "Imp.IIBB: " & TRIM$(FORMATNUM$(impo6#, "F15.2")) & vbCrLf _
'        & "Imp.Municipales: " & TRIM$(FORMATNUM$(impo7#, "F15.2")) & vbCrLf _
'        & "Imp.Internos: " & TRIM$(FORMATNUM$(impo8#, "F15.2")) & vbCrLf _
'        & "Otro Atributos: " & TRIM$(FORMATNUM$(impo9#, "F15.2")) & vbCrLf)
'
'    Close #NARCHI%
'
'    ARCHIIBB$ = "REGINFO_CV_COMPRAS_ALICUOTAS_" & PERIODO$ & ".TXT"
'    NARCHI1% = FILEOPEN%(LUCOM$ + ARCHIIBB$, 1, 62)
'    Do While Not EOF(NARCHI1%)
'        Line Input #NARCHI1%, XX$
'        SIGNO% = 1
'        If Mid$(XX$, 1, 3) = "003" Or Mid$(XX$, 1, 3) = "008" Or Mid$(XX$, 1, 3) = "021" Then
'              SIGNO% = -1
'        End If
'        IMPO20# = IMPO20# + SIGNO% * XVALO(Mid$(XX$, 51, 15)) / 100
'        impo21# = impo21# + SIGNO% * XVALO(Mid$(XX$, 70, 15)) / 100
'    Loop
'
'    MsgBox ("Neto: " & TRIM$(FORMATNUM$(IMPO20#, "F15.2")) & vbCrLf _
'        & "Iva: " & TRIM$(FORMATNUM$(impo21#, "F15.2")))
'
'    Close #NARCHI1%
    
    If TRIM$(ListaDeErrores$) <> "" Then
        Call COMUNI("Se Han Detectado Errores en la Generación.\Pulse Aceptar para Ver el Detalle")
        MSJESREGINFO.Text1 = ListaDeErrores$
        MSJESREGINFO.Caption = "Régimen Informativo de Compras"
        Call MSJESREGINFO.Command1_Click
        MSJESREGINFO.Show 1
        Call MENSERR(24, "Tener Presente que los Datos Generados Son Incompletos.\Debe Corregir los Errores y Volver a Generar")
    Else
        Call COMUNI("Regimen Informativo de Compras: Generación Ok")
    End If
    
99  Exit Sub
    
    
InicializarCamposVGD:
    
    VGD_FechaComprobante$(INDICE%) = "00000000"
    VGD_PuntoVta$(INDICE%) = ""
    VGD_NumeroCbtDesde$(INDICE%) = "0"
    VGD_NumeroCbtHasta$(INDICE%) = "0"
    VGD_ImporteTotal$(INDICE%) = "0"
    VGD_ImporteNoGravado$(INDICE%) = "0"
    VGD_ImporteOperacionExenta$(INDICE%) = "0"
    VGD_ImportePercepcionesNacionales$(INDICE%) = "0"
    VGD_ImportePercepcionIB$(INDICE%) = "0"
    VGD_CantidadAlicuotas$(INDICE%) = "0"
    VGD_RegistroRIV$(INDICE%) = ""

    Return
    
End Sub

Private Sub LibroIvaDigitalComprasLocales()
    
    On Error Resume Next
    Call TRALOCAL("LIVACOM", "")
    On Error GoTo 0
    
    If Err > 0 Then
        Call MENSERR(24, "Imposible Copiar Archivo con Datos del Período")
        Exit Sub
    End If
    
'    If InStr(1, EMPREAC$, "ASEMBLI") > 0 Then
'        MODOASEMBLI% = 1
'    End If
    
    PERIODO$ = TRIM$(MESAN)
    ' borrar archivo RIV
    NX% = FILEOPEN%(LUCOM$ & "LIBRO_IVA_DIGITAL_COMPRAS_CBTE_" & PERIODO$ & ".TXT", 2, 266)
    NX1% = FILEOPEN%(LUCOM$ & "LIBRO_IVA_DIGITAL_COMPRAS_ALICUOTAS_" & PERIODO$ & ".TXT", 2, 62)

    UltimoNumeroComprobante$ = ""
    UltimoCuitProveedor$ = ""
    
    Dim VGD_FechaComprobante$(3)
    Dim VGD_TipoCbte$(3)
        VGD_TipoCbte$(1) = TipoComprobante("FC-0000-00000000-B")
        VGD_TipoCbte$(2) = TipoComprobante("ND-0000-00000000-B")
        VGD_TipoCbte$(3) = TipoComprobante("NC-0000-00000000-B")
    Dim VGD_PuntoVta$(3)
    Dim VGD_NumeroCbtDesde$(3)
    Dim VGD_NumeroCbtHasta$(3)
    VGD_CodigoDocComprador$ = "99"
    VGD_NumeroIdentificacionComprador$ = RIVFormatoNumero$(0, 20)
    VGD_RazonSocial$ = AJUSTI$("VENTA GLOBAL DIARIA", 30)
    Dim VGD_ImporteTotal$(3)
    Dim VGD_ImporteNoGravado$(3)
    VGD_PercepcionNoCategorizados$ = RIVFormatoNumero$(0, 15)
    Dim VGD_ImporteOperacionExenta$(3)
    Dim VGD_ImportePercepcionesNacionales$(3)
    Dim VGD_ImportePercepcionIB$(3)
    VGD_ImportePercepcionMunicipales$ = RIVFormatoNumero$(0, 15)
    VGD_ImporteImpuestosInternos$ = RIVFormatoNumero$(0, 15)
    VGD_CodMoneda$ = CodigoMoneda(1)
    VGD_TipoCambio$ = RIVFormatoNumero$(1, 10, 6)
    Dim VGD_CantidadAlicuotas$(3)
    VGD_CodigoOperacion$ = " "
    VGD_OtrosTributos$ = RIVFormatoNumero$(0, 15)
    VGD_FechaVencimiento$ = "00000000"
    Dim VGD_RegistroRIV$(3)
    Dim INDICE%
    
    Dim RegistroRIV_Alicuotas$
    Dim RegistroRIV_Importaciones$
    Dim FlagError%
    Dim ListaDeErrores$
    Dim EsRegimenSimplificado%
    
    INDICE% = 1: GoSub InicializarCamposVGD
    INDICE% = 2: GoSub InicializarCamposVGD
    INDICE% = 3: GoSub InicializarCamposVGD
    
    FIN& = ULTREG&("!LIVACOM")
    If FIN& = 0 Then GoTo 99
    
    Screen.MousePointer = 11
    
    For i& = 1 To ULTREG&("!LIVACOM")
        Call TRACE(24, i&, ULTREG&("!LIVACOM"))
        
        REGISTRO$ = REGLEIDO$("!LIVACOM", i&)
        NumeroComprobante$ = Mid$(REGISTRO$, 11, 18)
        
        If Mid$(NumeroComprobante$, 1, 2) = "TC" Then
            NUMMOVITARJETA$ = Mid$(NumeroComprobante$, 4, 6) & Mid$(NumeroComprobante$, 11, 2)
            NumeroComprobante$ = "CO-A-0001-" & NUMMOVITARJETA$
        End If

        CUITPROVEEDOR$ = TRIM$(Mid$(REGISTRO$, 63, 16))
        
        If TRIM$(NumeroComprobante$) = "" Then
            NumeroComprobante$ = UltimoNumeroComprobante$
            CUITPROVEEDOR$ = UltimoCuitProveedor$
            Call REPLA(REGISTRO$, MKI$(Abs(NumeroProveedor%)), 31, 2)
        End If
       
        ' VERIFICA SI SE TRATA DE UN DESPACHO DE IMPORTACION
        NumeroProveedor% = CVI(Mid$(REGISTRO$, 31, 2)) * (-1)
        Despa$ = SAFETEXT$(VALOBADA("CAJABANC", "DespaImpo", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%), "NOMESS"))
        NumeroDespachoImportacion$ = UCase(AJUSTI$(Despa$, 16))
        
        ' VERIFICA QUE NO SE TRATE DE UN DESPACHO DE IMPORTACION O UNA PARTICULAR CON CUIT DE AFIP. LAS PARTICULAS QUE NO TIENEN CUIT DE AFIP, SE PROCESAN COMO LOCALES.
        CuitAfip$ = "33693450239"
        If TRIM$(NumeroDespachoImportacion$) <> "" Then
            If InStr(1, NumeroDespachoImportacion$, "PART") < 1 Or _
                (InStr(1, NumeroDespachoImportacion$, "PART") > 0 And REPLACAR$(CUITPROVEEDOR$, "-", "") = CuitAfip$) Then
                UltimoNumeroComprobante$ = NumeroComprobante$
                UltimoCuitProveedor$ = CUITPROVEEDOR$
                GoTo 35
            End If
        End If
       
        If (NumeroComprobante$ <> UltimoNumeroComprobante$) Or ((NumeroComprobante$ = UltimoNumeroComprobante$) And (CUITPROVEEDOR$ <> UltimoCuitProveedor$)) Then
            '
            RegistroAnterior$ = REGISTRO$
            
            FechaComprobante$ = FECHATEXLAR$(CVI(Mid$(REGISTRO$, 3, 2)))
            FechaComprobante$ = Mid$(FechaComprobante$, 7, 4) & Mid$(FechaComprobante$, 4, 2) & Mid$(FechaComprobante$, 1, 2)

            If UltimoNumeroComprobante$ <> "" And FlagError% = 0 Then
                If TRIM$(RegistroRIV$) <> "" Then
                    Print #NX%, RegistroRIV$
                    RegistroRIV$ = ""
                    If RegistroRIV_IvaAduana$ <> "" Then
                        Print #NX1%, RegistroRIV_IvaAduana$
                        RegistroRIV_IvaAduana$ = ""
                    End If
                End If
            End If
            FlagError% = 0
            EsRegimenSimplificado% = 0

25          fechaCbte$ = FETEXCOR1$(CVI(Mid$(REGISTRO$, 3, 2)))
            NumeroProveedor% = CVI(Mid$(REGISTRO$, 31, 2)) * (-1)
            If NumeroProveedor% = 0 Then
                ListaDeErrores$ = ListaDeErrores$ & "No Se Pudo Encontrar Numero de Proveedor para el Comprobante '" & NumeroComprobante$ & "'"
                FlagError% = 1: GoTo 35
            End If

            TipoCbte$ = TipoComprobante$(NumeroComprobante$)
            
            PUNTOVTA$ = PuntoVenta$(NumeroComprobante$)
            ' Es obligatorio que el punto de venta sea > 0
            If PUNTOVTA$ = "00000" Then PUNTOVTA$ = "00001"

            NumeroCbt$ = NumeroCbte$(NumeroComprobante$)
                             
            CodigoDocComprador$ = CodigoDocumento$(PIVACLI%(NumeroProveedor%))
                        
            'NumeroIdentificacionComprador$ = SAFETEXT$(VALOBADA("CAJABANC", "Cuit", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%), "NOMESS"))
            'If Len(NumeroIdentificacionComprador$) < 11 Or VALICUIT%(NumeroIdentificacionComprador$) = 0 Then NumeroIdentificacionComprador$ = TRIM$(Mid$(REGISTRO$, 63, 16))
            'If NumeroIdentificacionComprador$ = "" Or VALICUIT%(NumeroIdentificacionComprador$) = 0 Then
            '    NumeroIdentificacionComprador$ = CUITCLI$(NumeroProveedor%)
            'End If
            NumeroIdentificacionComprador$ = TRIM$(Mid$(REGISTRO$, 63, 16))
            If NumeroIdentificacionComprador$ = "" Then
                ListaDeErrores$ = ListaDeErrores$ & "CUIT Inválido en el Proveedor '" & CStr(Abs(NumeroProveedor%)) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                FlagError% = 1: GoTo 35
            End If
            If PIVACLI%(NumeroProveedor%) <> 5 And PIVACLI%(NumeroProveedor%) <> 0 Then
                If VALICUIT%(NumeroIdentificacionComprador$) = 0 Then
                    ListaDeErrores$ = ListaDeErrores$ & "CUIT Inválido en el Proveedor '" & CStr(Abs(NumeroProveedor%)) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                    FlagError% = 1: GoTo 35
                End If
            End If
            NumeroIdentificacionComprador$ = SOLONUM$(NumeroIdentificacionComprador$)
            NumeroIdentificacionComprador$ = AJUSTD$(NumeroIdentificacionComprador$, 20)
            NumeroIdentificacionComprador$ = REPLACAR(NumeroIdentificacionComprador$, " ", "0")

            ' Toma la razon social desde la base de datos. Prevé la posibilidad de que
            ' se trate de proveedores varios
            razonSocial$ = AJUSTI$(SAFETEXT$(VALOBADA("CAJABANC", "RefeCom", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%), "NOMESS")), 30)
            If TRIM$(razonSocial$) = "" Then razonSocial$ = AJUSTI$(rasocli$(NumeroProveedor%), 30)
            LargoRS% = InStr(1, razonSocial$, "//")
            If LargoRS% > 0 Then razonSocial$ = AJUSTI$(Mid(razonSocial$, 1, LargoRS% - 1), 30)
            If PIVACLI%(NumeroProveedor%) = 0 Then
                razonSocial$ = AJUSTI$("CONSUMIDOR FINAL", 30)
            End If
            
                MONE% = XVALO(VALOBADA("CAJABANC", "Mone_Emis", "CODICOM_LAR='" & NumeroComprobante$ & "' AND FechConta='" & fechaCbte$ & "'"))
            CODMONEDA$ = CodigoMoneda(MONE%)
            If CODMONEDA$ = "000" Then CODMONEDA$ = "PES"
            
                TCAMBIO# = XVALO(VALOBADA("CAJABANC", "Tipo_Cam", "CODICOM_LAR='" & NumeroComprobante$ & "' AND FechConta='" & fechaCbte$ & "'"))
            TIPOCAMBIO$ = RIVFormatoNumero$(TCAMBIO#, 10, 6)
            If CODMONEDA$ = "PES" And TCAMBIO# > 1 Then
                    ListaDeErrores$ = ListaDeErrores$ & "Error en Tipo de Cambio en el Proveedor '" & CStr(Abs(NumeroProveedor%)) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                    FlagError% = 1: GoTo 35
            End If
            If XVALO(TIPOCAMBIO$) = 0 And CODMONEDA$ = "PES" Then TIPOCAMBIO$ = RIVFormatoNumero(1, 10, 6)
            
            TipoCambioCbte# = 1
            If MONE% > 1 Then TipoCambioCbte# = TCAMBIO#
            
            ImporteTotal$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 91, 8)) / TipoCambioCbte#, 15)
            
            ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
            ValorIva$ = TRIM$(Mid$(REGISTRO$, 129, 8))
            LetraComprobante$ = Mid$(NumeroComprobante$, 4, 1)
            If InStr(1, UCase(ValorIva$), "NO GRAV") > 0 And (LetraComprobante$ <> "B" And LetraComprobante$ <> "C") Then
                ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
            End If
            If InStr(1, "011;012;013;006;007;008;019;020;021", TipoCbte$) > 0 Then ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
            
            PercepcionNoCategorizados$ = RIVFormatoNumero$(0, 15)

            ImporteOperacionExenta$ = RIVFormatoNumero$(0, 15)
            If PIVACLI%(NumeroProveedor%) = 5 Or PIVACLI%(NumeroProveedor%) = 7 Then
                ImporteOperacionExenta$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
            End If
            
            ImportePercepcionPagosACuenta$ = RIVFormatoNumero$(CVD(Mid$(REGISTRO$, 161, 8)) / TipoCambioCbte#, 15)
            
            ' Percepciones de iva + ganancias
                PercepcionesVarias# = CVD(Mid$(REGISTRO$, 169, 8))
            ImportePercepcionesNacionales$ = RIVFormatoNumero$(PercepcionesVarias# / TipoCambioCbte#, 15)
            
            ImportePercepcionIB$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 193, 8)) / TipoCambioCbte#, 15)
            
            ImportePercepcionMunicipales$ = RIVFormatoNumero$(0, 15)
            
            ImporteImpuestosInternos$ = RIVFormatoNumero$(0, 15)
                

                CANTALICUOTAS% = 1
            If InStr(1, "006;007;008;009;082;013;012;011;015;013;111", TipoCbte$) > 0 Or (InStr(1, "041", TipoCbte$) > 0 And LetraComprobante$ = "C") Then
                CANTALICUOTAS% = 0 ' Si es Comprobante "C" la cantidad debe ser cero
            End If
            CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
            
            Percepciones# = XVALO(ImportePercepcionesNacionales$) / 100 + XVALO(ImportePercepcionIB$) / 100 + XVALO(ImportePercepcionPagosACuenta$) / 100
            If PIVACLI%(NumeroProveedor%) = 7 Then
                CodigoOperacion$ = "Z"
            ElseIf PIVACLI%(NumeroProveedor%) = 5 Or (TRIM$(NumeroDespachoImportacion$) <> "" Or EsFacturaDespachante# > 0.005) Then
                CodigoOperacion$ = "X"
            ElseIf InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 Or (TRIM$(Mid$(REGISTRO$, 129, 8)) = "" And Percepciones# > 0.005) Then
                CodigoOperacion$ = "N"
            ElseIf PIVACLI%(NumeroProveedor%) = 4 Or InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "EXENTO") > 0 Then
                CodigoOperacion$ = "E"
            ElseIf CANTALICUOTAS% = 0 Then
                CodigoOperacion$ = "0"
            Else
                CodigoOperacion$ = " "
            End If
                
            ImporteIva$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)) / TipoCambioCbte#, 15)
            
            OtrosTributos$ = RIVFormatoNumero$(0, 15)
            
            CuitEmisor$ = "00000000000"
            
            DenominacionEmisor$ = AJUSTI$(" ", 30)
            
            IvaComision$ = RIVFormatoNumero("0", 15)
            
            ' Arma el registro para el regimen informativo
            RegistroRIV$ = FechaComprobante$ & TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & NumeroDespachoImportacion$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & razonSocial$ _
                               & ImporteTotal$ & ImporteNoGravado$ & ImporteOperacionExenta$ & ImportePercepcionPagosACuenta$ & ImportePercepcionesNacionales$ _
                               & ImportePercepcionIB$ & ImportePercepcionMunicipales$ & ImporteImpuestosInternos$ & CODMONEDA$ & TIPOCAMBIO$ & CantidadAlicuotas$ _
                               & CodigoOperacion$ & ImporteIva$ & OtrosTributos$ & CuitEmisor$ & DenominacionEmisor$ & IvaComision$
            Debug.Print RegistroRIV$
            UltimoNumeroComprobante$ = NumeroComprobante$
            UltimoCuitProveedor$ = CUITPROVEEDOR$

            'If (InStr(1, UCase(ValorIva$), "NO GRAV") = 0 And XVALO(ImporteIva$) > 0) Then
            If XVALO(CantidadAlicuotas$) > 0 Then
            
                ' Tratamiento para los comprobantes con parte no gravada o exenta
                If TRIM$(CodigoOperacion$) <> "" Then
                    ImporteNeto$ = RIVFormatoNumero(0, 15)
                    ALICUOTA$ = "0003"
                    ImporteIva0$ = RIVFormatoNumero(0, 15)

                    If (CodigoOperacion$ = "N") Then
                        ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%))) / TipoCambioCbte#
                        Percepciones# = XVALO(ImportePercepcionesNacionales$) / 100 + XVALO(ImportePercepcionIB$) / 100 + XVALO(ImportePercepcionPagosACuenta$) / 100
                        ImporteEx# = XVALO(VALOBADA("CAJABANC", "Sum(BaseImpo)", "VariMovim like 'Exento' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%))) / TipoCambioCbte#
                        If ConIva# < 0.01 Or (Percepciones# > 0.005 And ConIva# < 0.01) Or (ImporteEx# > 0.005 And ConIva# < 0.01) Then
                            ' cuando el comprobante es totalmente no gravado, dicho importe se escribe en el archivo de alicuotas
                            ImporteNeto$ = ImporteTotal$
                            If Percepciones# > 0.005 Or ImporteEx# > 0.005 Then ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                            ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
                            Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                        Else
                            ' Cuando el comprobante tiene parte gravada y parte que no lo esta, el gravado se escribe
                            ' en el archivo de comprobantes y los importes gravados en el de alicuotas
                            CANTALICUOTAS% = 0
                            CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                            Call REPLA(RegistroRIV$, " ", 239, 1)
                            Call REPLA(RegistroRIV$, CantidadAlicuotas$, 238, 1)
                            GoTo 35
                        End If
                    End If
                    If CodigoOperacion$ = "E" Then
                        ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%))) / TipoCambioCbte#
                        Percepciones# = XVALO(ImportePercepcionesNacionales$) / 100 + XVALO(ImportePercepcionIB$) / 100 + XVALO(ImportePercepcionPagosACuenta$) / 100
                        NoGravado# = XVALO(VALOBADA("CAJABANC", "Sum(BaseImpo)", "VariMovim like 'No Grav%' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                        If (ConIva# < 0.005 Or (Percepciones# > 0.005 And ConIva# < 0.01)) And NoGravado# < 0.005 Then
                            ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                            ImporteOperacionExenta$ = RIVFormatoNumero$(0, 15)
                            Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 135, 15)
                        ElseIf NoGravado# > 0.005 Then
                            ImporteOperacionExenta$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                            Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 135, 15)
                            CANTALICUOTAS% = 0
                            CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                            If Mid$(NumeroComprobante$, 1, 2) = "ND" Or Mid$(NumeroComprobante$, 1, 2) = "FC" Then
                                ImporteNeto$ = RIVFormatoNumero(0, 15)
                                ImporteIva0$ = RIVFormatoNumero(0, 15)
                                ALICUOTA$ = "0003"
                                CANTALICUOTAS% = 1
                                CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                            Else
                                GoTo 35
                            End If
                        Else
                            CANTALICUOTAS% = 0
                            CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                            Call REPLA(RegistroRIV$, " ", 239, 1)
                            Call REPLA(RegistroRIV$, CantidadAlicuotas$, 238, 1)
                            ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                            Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                            GoTo 35
                        End If
                    End If
                    If CodigoOperacion$ = "X" And (TRIM$(NumeroDespachoImportacion$) <> "" Or EsFacturaDespachante# > 0.005) Then
                        If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") > 0 Then
                            ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%))) / TipoCambioCbte#
                            ImpoNoGra# = XVALO(VALOBADA("CAJABANC", "Sum(BaseImpo)", "VariMovim='Aduana' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                            ImpoAduana# = XVALO(ImporteIva$) / 100
                            OtrosTributos$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)) / TipoCambioCbte#, 15)
                            ImporteIva$ = RIVFormatoNumero(0, 15)
                            Call REPLA(RegistroRIV$, ImporteIva$, 240, 15)
                            CantidadAlicuotas$ = "0"
                            If ImpoNoGra# > 0.005 Then
                                ImporteNGAnterior = XVALO(Mid$(RegistroRIV$, 120, 15)) / 100
                                ImpoNoGra# = ImpoNoGra# + ImporteNGAnterior
                                ImporteNoGravado$ = RIVFormatoNumero(ImpoNoGra#, 15, , 1)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                            End If
                            If (ConIva# < 0.005 And EsRegimenSimplificado% = 0) Then
                                Call REPLA(RegistroRIV$, RIVFormatoNumero(ImpoAduana#, 15), 150, 15)
                                ImporteNeto$ = RIVFormatoNumero(0, 15)
                                ImporteIva0$ = RIVFormatoNumero(0, 15)
                                ALICUOTA$ = "0003"
                                CantidadAlicuotas$ = "1"
                            Else
                                If MODOASEMBLI% > 0 And EsRegimenSimplificado% = 1 Then
                                    RegistroRIV_IvaAduana$ = String$(65, "0") & "0003" & RIVFormatoNumero(ImpoAduana#, 15)
                                End If
                                GoTo 35
                            End If
                        ElseIf InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "EXENTO") > 0 Then
                            'ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                            ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%))) / TipoCambioCbte#
                            If ConIva# > 0.005 Then
                                ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8) / TipoCambioCbte#), 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                CANTALICUOTAS% = 0
                                CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                                GoTo 35
                            Else
                                ImporteNoGravado$ = RIVFormatoNumero(0, 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                            End If
                        ElseIf InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 Then
                            ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%))) / TipoCambioCbte#
                            If ConIva# > 0.005 Then
                                ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                CANTALICUOTAS% = 0
                                CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                                GoTo 35
                            Else
                                ImporteNoGravado$ = RIVFormatoNumero(0, 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                            End If
                        ElseIf InStr(1, "0004;0005;0006;0008;0009", CodigoAlicuota(Mid$(REGISTRO$, 129, 8))) > 0 Then
                            ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                            ImporteIva0$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
                            ALICUOTA$ = CodigoAlicuota(ValorIva$)
                        End If
'                        If CodigoOperacion$ = "X" Then
'                            If CantRegistros("CAJABANC", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%) & " AND IDEBECOMP>0 AND TIPOMOVIM='COMPRA'") > 1 Then
'                                CANTALICUOTAS% = 0
'                                CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
'                                Call REPLA(RegistroRIV$, " ", 239, 1)
'                                Call REPLA(RegistroRIV$, CantidadAlicuotas$, 238, 1)
'                                GoTo 35
'                            Else
'                                ImporteOperacionExenta$ = RIVFormatoNumero(0, 15)
'                                Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 135, 15)
'                                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
'                            End If
'                        End If
                    ElseIf CodigoOperacion$ = "X" Then
'                        If CantRegistros("CAJABANC", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%) & " AND IDEBECOMP>0 AND TIPOMOVIM='COMPRA'") > 1 Then
'                                CANTALICUOTAS% = 0
'                                CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
'                                Call REPLA(RegistroRIV$, " ", 239, 1)
'                                Call REPLA(RegistroRIV$, CantidadAlicuotas$, 238, 1)
'                                GoTo 35
'                        Else
                                ImporteOperacionExenta$ = RIVFormatoNumero(0, 15)
                                Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 135, 15)
                                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
'                        End If
                    End If
                Else
                    ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                    ImporteIva0$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)) / TipoCambioCbte#, 15)
                    ALICUOTA$ = CodigoAlicuota(ValorIva$)
                End If
                
                If (TRIM$(NumeroDespachoImportacion$) = "" And EsFacturaDespachante# < 0.005) Or EsRegimenSimplificado% = 1 Then
                
                    RegistroRIV_Alicuotas$ = TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & ImporteNeto$ & ALICUOTA$ & ImporteIva0$
                    If MODOASEMBLI% = 1 And EsRegimenSimplificado% = 1 Then
                        RegistroRIV_IvaAduana$ = RegistroRIV_Alicuotas$
                    Else
                        Print #NX1%, RegistroRIV_Alicuotas$
                    End If
                Else
                    
                    RegistroRIV_Importaciones$ = NumeroDespachoImportacion$ & ImporteNeto$ & ALICUOTA$ & ImporteIva0$
                    'Print #NX2%, RegistroRIV_Importaciones$
                End If
            End If
        Else
            If FlagError% = 1 Then GoTo 35
            If TipoCbte$ = "006" Then GoTo 35
            
            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 Then
                ImporteNGAnterior = XVALO(Mid$(RegistroRIV$, 120, 15)) / 100
                ImporteNGActual = CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#
                'If Not (ImporteNGActual < 0 And Mid$(REGISTRO$, 11, 2) = "FC") Then
                If Mid$(REGISTRO$, 11, 2) <> "NC" Then
                    'If A <> 1 Then ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior + ImporteNGActual, 15, , 1)
                    ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior + ImporteNGActual, 15, , 1)
                Else
                    ConsideraSigno% = 0
                    If ImporteNGActual > 0 Then ' INCORPORADO PARA EL VALOR CIF
                        ImporteNGActual = (-1) * ImporteNGActual
                        ConsideraSigno% = 1
                    End If
                    ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior + ImporteNGActual, 15, , ConsideraSigno%)
                End If
                'End If
            End If
            
            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "EXENTO") > 0 Then
                ImporteExAnterior = XVALO(Mid$(RegistroRIV$, 135, 15)) / 100
                ImporteExActual = CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#
                ImporteOperacionExenta$ = RIVFormatoNumero(ImporteExAnterior + ImporteExActual, 15)
            End If
            
            If PIVACLI%(NumeroProveedor%) = 5 Or PIVACLI%(NumeroProveedor%) = 7 Then
                IMPO# = XVALO(Mid$(RegistroRIV$, 135, 15)) / 100
                ImporteOperacionExenta$ = RIVFormatoNumero(IMPO# + CVD(Mid$(REGISTRO$, 185, 8)) / TipoCambioCbte#, 15)
            End If
                                                                             
                IMPO# = XVALO(Mid$(RegistroRIV$, 150, 15)) / 100
            ImportePercepcionPagosACuenta$ = RIVFormatoNumero$(IMPO# + CVD(Mid$(REGISTRO$, 161, 8)) / TipoCambioCbte#, 15)
            
                IMPO# = XVALO(Mid$(RegistroRIV$, 165, 15)) / 100
            ImportePercepcionesNacionales$ = RIVFormatoNumero$(IMPO# + CVD(Mid$(REGISTRO$, 169, 8)) / TipoCambioCbte#, 15)

                IMPO# = XVALO(Mid$(RegistroRIV$, 180, 15)) / 100
            ImportePercepcionIB$ = RIVFormatoNumero(IMPO# + CVD(Mid$(REGISTRO$, 193, 8)) / TipoCambioCbte#, 15)
                                
            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") = 0 Then
                    IMPO# = XVALO(Mid$(RegistroRIV$, 240, 15)) / 100
                    ImporteIvaActual# = Abs(CVD(Mid$(REGISTRO$, 145, 8)) / TipoCambioCbte#)
                    If A = 0 Then
                        ImporteIva$ = RIVFormatoNumero(IMPO# + ImporteIvaActual#, 15)
                    Else
                        If ImporteIvaActual# > 0.005 Then
                            IMPO# = XVALO(Mid$(RegistroRIV$, 150, 15)) / 100
                            ImportePercepcionPagosACuenta$ = RIVFormatoNumero(IMPO# + ImporteIvaActual#, 15)
                        End If
                    End If
            End If

                CANTALICUOTAS% = XVALO(CantidadAlicuotas$)
            ImporteNeto0# = CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#
            ValorIva$ = TRIM$(REPLACAR$(Mid$(REGISTRO$, 129, 8), "%", ""))
            CantidadAlicuotaNuevoRegistro% = 0
            If (Not (ImporteNeto0# < 0 And Mid$(REGISTRO$, 11, 2) = "FC")) And XVALO(ImporteIvaActual#) > 0.005 And InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") = 0 Then CantidadAlicuotaNuevoRegistro% = 1
'            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 And (CodigoOperacion$ = "X" Or CodigoOperacion$ = "E") And _
'                CantRegistros("CAJABANC", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%) & " AND IDEBECOMP>0 AND (TIPOMOVIM='COMPRA' or TIPOMOVIM='REDCOM')") < 2 Then
'
'                CantidadAlicuotaNuevoRegistro% = 1
'                ImporteNGAnterior = XVALO(ImporteNoGravado$) / 100
'                ImporteNGActual = CVD(Mid$(REGISTRO$, 137, 8))
'                ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior - ImporteNGActual, 15)
'            End If
            CantidadAlicuotas$ = CStr(CANTALICUOTAS% + CantidadAlicuotaNuevoRegistro%)
            
            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") > 0 Then
'                IMPO# = XVALO(Mid$(RegistroRIV$, 255, 15)) / 100
'                OtAtributos# = CVD(Mid$(REGISTRO$, 145, 8))
'                OTROSTRIBUTOS$ = RIVFormatoNumero(IMPO# + OtAtributos#, 15)
'                IMPO# = XVALO(ImporteIva$) / 100
'                ImporteIva$ = RIVFormatoNumero(IMPO# - OtAtributos#, 15)
                IMPO# = XVALO(ImportePercepcionPagosACuenta$) / 100
                ImporteAduana# = CVD(Mid$(REGISTRO$, 145, 8)) / TipoCambioCbte#
                If MODOASEMBLI% = 0 Or (MODOASEMBLI% = 1 And EsRegimenSimplificado% = 0) Then
                    ImportePercepcionPagosACuenta$ = RIVFormatoNumero(IMPO# + ImporteAduana#, 15)
                End If
                ' SI NO TIENE ALICUOTA PREVIA, AGREGA UN REGISTRO EN EL ARCHIVO DE ALICUOTA. ES FICTICIO PORQUE
                ' TANTO EL NETO GRAVADO COMO EL IMPORTE DE IVA, ESTAN EN CERO.
                ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                If ConIva# < 0.005 And (MODOASEMBLI% = 0 Or (MODOASEMBLI% = 1 And EsRegimenSimplificado% = 0)) Then
                    CantidadAlicuotaNuevoRegistro% = 1
                    CantidadAlicuotas$ = CStr(CANTALICUOTAS% + CantidadAlicuotaNuevoRegistro%)
                End If
                
                If MODOASEMBLI% = 1 And EsRegimenSimplificado% = 1 And RegistroRIV_IvaAduana$ <> "" Then
                    CantidadAlicuotaNuevoRegistro% = 1
                End If
            End If
                        
            RegistroRIV$ = FechaComprobante$ & TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & NumeroDespachoImportacion$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & razonSocial$ _
                           & ImporteTotal$ & ImporteNoGravado$ & ImporteOperacionExenta$ & ImportePercepcionPagosACuenta$ & ImportePercepcionesNacionales$ _
                           & ImportePercepcionIB$ & ImportePercepcionMunicipales$ & ImporteImpuestosInternos$ & CODMONEDA$ & TIPOCAMBIO$ & CantidadAlicuotas$ _
                           & CodigoOperacion$ & ImporteIva$ & OtrosTributos$ & CuitEmisor$ & DenominacionEmisor$ & IvaComision$
            
            If (CantidadAlicuotaNuevoRegistro% > 0) Then
                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                ImporteIva0$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)) / TipoCambioCbte#, 15)
                ALICUOTA$ = CodigoAlicuota(ValorIva$)
                If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 And (CodigoOperacion$ = "X" Or CodigoOperacion$ = "E") Then
                    ALICUOTA$ = "0003"
                End If
                If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") > 0 And (MODOASEMBLI% = 0 Or (MODOASEMBLI% = 1 And EsRegimenSimplificado% = 0)) Then
                    ImporteNeto$ = RIVFormatoNumero(0, 15)
                    ImporteIva0$ = RIVFormatoNumero(0, 15)
                End If
                If A = 1 Then
                    ALICUOTA$ = "0003"
                    ImporteNeto$ = RIVFormatoNumero(0, 15)
                    ImporteIva0$ = RIVFormatoNumero(0, 15)
                End If
                
                If (TRIM$(NumeroDespachoImportacion$) = "" And EsFacturaDespachante# < 0.005) Or EsRegimenSimplificado% = 1 Then
                
                    RegistroRIV_Alicuotas$ = TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & ImporteNeto$ & ALICUOTA$ & ImporteIva0$
                    If MODOASEMBLI% = 1 And EsRegimenSimplificado% = 1 Then
                        If RegistroRIV_IvaAduana$ <> "" Then        ' EL IMPORTE DEL IVA ADUANA SE AGREGA AL CREDITO FISCAL DE ALGUNA DE LAS ALICUOTAS DEL COMPROBANTE
                            If Mid$(RegistroRIV_IvaAduana$, 66, 4) = "0003" Then
                                ImpoAduana# = XVALO(Mid$(RegistroRIV_IvaAduana$, 70, 15)) / 100
                                IMPOIVA# = XVALO(Mid$(RegistroRIV_Alicuotas$, 70, 15)) / 100
                                IMPOIVA# = IMPOIVA# + ImpoAduana#
                                RegistroRIV_Alicuotas$ = Mid$(RegistroRIV_Alicuotas$, 1, 69) & RIVFormatoNumero(IMPOIVA#, 15)
                            ElseIf ALICUOTA$ = "0003" Then
                                ImpoAduana# = XVALO(Mid$(RegistroRIV_IvaAduana$, 70, 15)) / 100
                                IMPOIVA# = XVALO(Mid$(RegistroRIV_Alicuotas$, 70, 15)) / 100
                                IMPOIVA# = IMPOIVA# + ImpoAduana#
                                RegistroRIV_Alicuotas$ = Mid$(RegistroRIV_IvaAduana$, 1, 69) & RIVFormatoNumero(IMPOIVA#, 15)
                            End If
                            RegistroRIV_IvaAduana$ = ""
                        Else
                            If RegistroRIV_IvaAduana$ = "" Then
                                RegistroRIV_IvaAduana$ = RegistroRIV_Alicuotas$
                                GoTo 35
                            End If
                        End If
                    End If
                    Print #NX1%, RegistroRIV_Alicuotas$
                Else
                    
                    RegistroRIV_Importaciones$ = NumeroDespachoImportacion$ & ImporteNeto$ & ALICUOTA$ & ImporteIva0$
                    'Print #NX2%, RegistroRIV_Importaciones$
                End If
            End If
        End If
35  Next i&
    
    If NumeroComprobante$ <> "" Then Print #NX%, RegistroRIV$
    
    Close #NX%
    Close #NX1%
    'Close #NX2%
    
    Screen.MousePointer = 1
    
'    Dim XX1 As String * 266
'
'    ARCHIIBB$ = "REGINFO_CV_COMPRAS_CBTE_" & PERIODO$ & ".TXT"
'    NARCHI% = FILEOPEN%(LUCOM$ + ARCHIIBB$, 1, 266)
'    Do While Not EOF(NARCHI%)
'        Line Input #NARCHI%, XX$
'        SIGNO% = 1
'        If Mid$(XX$, 9, 3) = "003" Or Mid$(XX$, 9, 3) = "008" Or Mid$(XX$, 9, 3) = "021" Or Mid$(XX$, 9, 3) = "013" Or Mid$(XX$, 9, 3) = "053" Then SIGNO% = -1
'
'        IMPO1# = IMPO1# + SIGNO% * XVALO(Mid$(XX$, 105, 15)) / 100
'        impo2# = impo2# + SIGNO% * XVALO(Mid$(XX$, 120, 15)) / 100
'        impo3# = impo3# + SIGNO% * XVALO(Mid$(XX$, 150, 15)) / 100
'        impo4# = impo4# + SIGNO% * XVALO(Mid$(XX$, 135, 15)) / 100
'        impo5# = impo5# + SIGNO% * XVALO(Mid$(XX$, 165, 15)) / 100
'        impo6# = impo6# + SIGNO% * XVALO(Mid$(XX$, 180, 15)) / 100
'        impo7# = impo7# + SIGNO% * XVALO(Mid$(XX$, 195, 15)) / 100
'        impo8# = impo8# + SIGNO% * XVALO(Mid$(XX$, 210, 15)) / 100
'        impo9# = impo9# + SIGNO% * XVALO(Mid$(XX$, 255, 15)) / 100
'        impo10# = impo10# + SIGNO% * XVALO(Mid$(XX$, 240, 15)) / 100
'20  Loop
'
'    MsgBox ("Total: " & TRIM$(FORMATNUM$(IMPO1#, "F15.2")) & vbCrLf _
'        & "No gravado: " & TRIM$(FORMATNUM$(impo2#, "F15.2")) & vbCrLf _
'        & "Iva: " & TRIM$(FORMATNUM$(impo10#, "F15.2")) & vbCrLf _
'        & "Percep. Iva: " & TRIM$(FORMATNUM$(impo3#, "F15.2")) & vbCrLf _
'        & "Op.Exentas: " & TRIM$(FORMATNUM$(impo4#, "F15.2")) & vbCrLf _
'        & "Imp.Nacionales: " & TRIM$(FORMATNUM$(impo5#, "F15.2")) & vbCrLf _
'        & "Imp.IIBB: " & TRIM$(FORMATNUM$(impo6#, "F15.2")) & vbCrLf _
'        & "Imp.Municipales: " & TRIM$(FORMATNUM$(impo7#, "F15.2")) & vbCrLf _
'        & "Imp.Internos: " & TRIM$(FORMATNUM$(impo8#, "F15.2")) & vbCrLf _
'        & "Otro Atributos: " & TRIM$(FORMATNUM$(impo9#, "F15.2")) & vbCrLf)
'
'    Close #NARCHI%
'
'    ARCHIIBB$ = "REGINFO_CV_COMPRAS_ALICUOTAS_" & PERIODO$ & ".TXT"
'    NARCHI1% = FILEOPEN%(LUCOM$ + ARCHIIBB$, 1, 62)
'    Do While Not EOF(NARCHI1%)
'        Line Input #NARCHI1%, XX$
'        SIGNO% = 1
'        If Mid$(XX$, 1, 3) = "003" Or Mid$(XX$, 1, 3) = "008" Or Mid$(XX$, 1, 3) = "021" Then
'              SIGNO% = -1
'        End If
'        IMPO20# = IMPO20# + SIGNO% * XVALO(Mid$(XX$, 51, 15)) / 100
'        impo21# = impo21# + SIGNO% * XVALO(Mid$(XX$, 70, 15)) / 100
'    Loop
'
'    MsgBox ("Neto: " & TRIM$(FORMATNUM$(IMPO20#, "F15.2")) & vbCrLf _
'        & "Iva: " & TRIM$(FORMATNUM$(impo21#, "F15.2")))
'
'    Close #NARCHI1%
    
    If TRIM$(ListaDeErrores$) <> "" Then
        Call COMUNI("Se Han Detectado Errores en la Generación\del IVA Digital de Compras.\Pulse Aceptar para Ver el Detalle")
        MSJESREGINFO.Text1 = ListaDeErrores$
        MSJESREGINFO.Caption = "Régimen Informativo de Compras"
        Call MSJESREGINFO.Command1_Click
        MSJESREGINFO.Show 1
        Call MENSERR(24, "Tener Presente que los Datos Generados Son Incompletos.\Debe Corregir los Errores y Volver a Generar")
    End If
    
99  Exit Sub
    
    
InicializarCamposVGD:
    
    VGD_FechaComprobante$(INDICE%) = "00000000"
    VGD_PuntoVta$(INDICE%) = ""
    VGD_NumeroCbtDesde$(INDICE%) = "0"
    VGD_NumeroCbtHasta$(INDICE%) = "0"
    VGD_ImporteTotal$(INDICE%) = "0"
    VGD_ImporteNoGravado$(INDICE%) = "0"
    VGD_ImporteOperacionExenta$(INDICE%) = "0"
    VGD_ImportePercepcionesNacionales$(INDICE%) = "0"
    VGD_ImportePercepcionIB$(INDICE%) = "0"
    VGD_CantidadAlicuotas$(INDICE%) = "0"
    VGD_RegistroRIV$(INDICE%) = ""

    Return
    
End Sub

Private Sub LibroIvaDigitalComprasImportaciones()
    
    On Error Resume Next
    Call TRALOCAL("LIVACOM", "")
    On Error GoTo 0
    
    If Err > 0 Then
        Call MENSERR(24, "Imposible Copiar Archivo con Datos del Período")
        Exit Sub
    End If
    
'    If InStr(1, EMPREAC$, "ASEMBLI") > 0 Then
'        MODOASEMBLI% = 1
'    End If
    
    PERIODO$ = TRIM$(MESAN)
    ' borrar archivo RIV
    NX% = FILEOPEN%(LUCOM$ & "LIBRO_IVA_DIGITAL_IMPORTACION_BIENES_CBTE_" & PERIODO$ & ".TXT", 2, 266)
    NX2% = FILEOPEN%(LUCOM$ & "LIBRO_IVA_DIGITAL_IMPORTACION_BIENES_ALICUOTA_" & PERIODO$ & ".TXT", 2, 62)

    UltimoNumeroComprobante$ = ""
    UltimoCuitProveedor$ = ""
    
    Dim VGD_FechaComprobante$(3)
    Dim VGD_TipoCbte$(3)
        VGD_TipoCbte$(1) = TipoComprobante("FC-0000-00000000-B")
        VGD_TipoCbte$(2) = TipoComprobante("ND-0000-00000000-B")
        VGD_TipoCbte$(3) = TipoComprobante("NC-0000-00000000-B")
    Dim VGD_PuntoVta$(3)
    Dim VGD_NumeroCbtDesde$(3)
    Dim VGD_NumeroCbtHasta$(3)
    VGD_CodigoDocComprador$ = "99"
    VGD_NumeroIdentificacionComprador$ = RIVFormatoNumero$(0, 20)
    VGD_RazonSocial$ = AJUSTI$("VENTA GLOBAL DIARIA", 30)
    Dim VGD_ImporteTotal$(3)
    Dim VGD_ImporteNoGravado$(3)
    VGD_PercepcionNoCategorizados$ = RIVFormatoNumero$(0, 15)
    Dim VGD_ImporteOperacionExenta$(3)
    Dim VGD_ImportePercepcionesNacionales$(3)
    Dim VGD_ImportePercepcionIB$(3)
    VGD_ImportePercepcionMunicipales$ = RIVFormatoNumero$(0, 15)
    VGD_ImporteImpuestosInternos$ = RIVFormatoNumero$(0, 15)
    VGD_CodMoneda$ = CodigoMoneda(1)
    VGD_TipoCambio$ = RIVFormatoNumero$(1, 10, 6)
    Dim VGD_CantidadAlicuotas$(3)
    VGD_CodigoOperacion$ = " "
    VGD_OtrosTributos$ = RIVFormatoNumero$(0, 15)
    VGD_FechaVencimiento$ = "00000000"
    Dim VGD_RegistroRIV$(3)
    Dim INDICE%
    
    Dim RegistroRIV_Alicuotas$
    Dim RegistroRIV_Importaciones$
    Dim FlagError%
    Dim ListaDeErrores$
    Dim EsRegimenSimplificado%
    
    INDICE% = 1: GoSub InicializarCamposVGD
    INDICE% = 2: GoSub InicializarCamposVGD
    INDICE% = 3: GoSub InicializarCamposVGD
    
    FIN& = ULTREG&("!LIVACOM")
    If FIN& = 0 Then GoTo 99
    
    Screen.MousePointer = 11
    
    For i& = 1 To ULTREG&("!LIVACOM")
        Call TRACE(24, i&, ULTREG&("!LIVACOM"))
        
        REGISTRO$ = REGLEIDO$("!LIVACOM", i&)
        NumeroComprobante$ = Mid$(REGISTRO$, 11, 18)
        
        CUITPROVEEDOR$ = TRIM$(Mid$(REGISTRO$, 63, 16))
        If TRIM$(NumeroComprobante$) = "" Then
            NumeroComprobante$ = UltimoNumeroComprobante$
            CUITPROVEEDOR$ = UltimoCuitProveedor$
            Call REPLA(REGISTRO$, MKI$(Abs(NumeroProveedor%)), 31, 2)
        End If
        
        ' VERIFICA SI SE TRATA DE UN DESPACHO DE IMPORTACION O UNA PARTICULAR PERO CON CUIT DE AFIP
        NumeroProveedor% = CVI(Mid$(REGISTRO$, 31, 2)) * (-1)
        Despa$ = SAFETEXT$(VALOBADA("CAJABANC", "DespaImpo", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%), "NOMESS"))
        NumeroDespachoImportacion$ = UCase(AJUSTI$(Despa$, 16))
        CuitAfip$ = "33693450239"
        If InStr(1, NumeroDespachoImportacion$, "PART") > 0 And REPLACAR$(CUITPROVEEDOR$, "-", "") <> CuitAfip$ Then
            NumeroDespachoImportacion$ = ""
        End If
        
        ' VERIFICA QUE SE TRATE DE UNA IMPORTACION (PARTICULAR O CON DESPACHO)
        If TRIM$(NumeroDespachoImportacion$) = "" Then
            GoTo 35
        End If
                
        If Mid$(NumeroComprobante$, 1, 2) = "TC" Then
            NUMMOVITARJETA$ = Mid$(NumeroComprobante$, 4, 6) & Mid$(NumeroComprobante$, 11, 2)
            NumeroComprobante$ = "CO-A-0001-" & NUMMOVITARJETA$
        End If

        If (NumeroComprobante$ <> UltimoNumeroComprobante$) Or ((NumeroComprobante$ = UltimoNumeroComprobante$) And (CUITPROVEEDOR$ <> UltimoCuitProveedor$)) Then
            '
            RegistroAnterior$ = REGISTRO$
            
            FechaComprobante$ = FECHATEXLAR$(CVI(Mid$(REGISTRO$, 3, 2)))
            FechaComprobante$ = Mid$(FechaComprobante$, 7, 4) & Mid$(FechaComprobante$, 4, 2) & Mid$(FechaComprobante$, 1, 2)

            If UltimoNumeroComprobante$ <> "" And FlagError% = 0 Then
                If TRIM$(RegistroRIV$) <> "" Then
                    Print #NX%, RegistroRIV$
                    RegistroRIV$ = ""
                    If RegistroRIV_IvaAduana$ <> "" Then
                        Print #NX1%, RegistroRIV_IvaAduana$
                        RegistroRIV_IvaAduana$ = ""
                    End If
                End If
            End If
            FlagError% = 0
            EsRegimenSimplificado% = 0

25          fechaCbte$ = FETEXCOR1$(CVI(Mid$(REGISTRO$, 3, 2)))
            NumeroProveedor% = CVI(Mid$(REGISTRO$, 31, 2)) * (-1)
            If NumeroProveedor% = 0 Then
                ListaDeErrores$ = ListaDeErrores$ & "No Se Pudo Encontrar Numero de Proveedor para el Comprobante '" & NumeroComprobante$ & "'"
                FlagError% = 1: GoTo 35
            End If

            TipoCbte$ = TipoComprobante$(NumeroComprobante$)
            
            NumeroCbt$ = NumeroCbte$(NumeroComprobante$)
                                         
            PUNTOVTA$ = "00000"
            
            TipoCbte$ = "066"
            
            CodigoDocComprador$ = CodigoDocumento$(PIVACLI%(NumeroProveedor%))
                        
            'NumeroIdentificacionComprador$ = SAFETEXT$(VALOBADA("CAJABANC", "Cuit", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%), "NOMESS"))
            'If Len(NumeroIdentificacionComprador$) < 11 Or VALICUIT%(NumeroIdentificacionComprador$) = 0 Then NumeroIdentificacionComprador$ = TRIM$(Mid$(REGISTRO$, 63, 16))
            'If NumeroIdentificacionComprador$ = "" Or VALICUIT%(NumeroIdentificacionComprador$) = 0 Then
            '    NumeroIdentificacionComprador$ = CUITCLI$(NumeroProveedor%)
            'End If
            NumeroIdentificacionComprador$ = TRIM$(Mid$(REGISTRO$, 63, 16))
            If NumeroIdentificacionComprador$ = "" Then
                ListaDeErrores$ = ListaDeErrores$ & "CUIT Inválido en el Proveedor '" & CStr(Abs(NumeroProveedor%)) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                FlagError% = 1: GoTo 35
            End If
            If PIVACLI%(NumeroProveedor%) <> 5 And PIVACLI%(NumeroProveedor%) <> 0 Then
                If VALICUIT%(NumeroIdentificacionComprador$) = 0 Then
                    ListaDeErrores$ = ListaDeErrores$ & "CUIT Inválido en el Proveedor '" & CStr(Abs(NumeroProveedor%)) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                    FlagError% = 1: GoTo 35
                End If
            End If
            NumeroIdentificacionComprador$ = SOLONUM$(NumeroIdentificacionComprador$)
            NumeroIdentificacionComprador$ = AJUSTD$(NumeroIdentificacionComprador$, 20)
            NumeroIdentificacionComprador$ = REPLACAR(NumeroIdentificacionComprador$, " ", "0")

            ' Toma la razon social desde la base de datos. Prevé la posibilidad de que
            ' se trate de proveedores varios
            razonSocial$ = AJUSTI$(SAFETEXT$(VALOBADA("CAJABANC", "RefeCom", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%), "NOMESS")), 30)
            If TRIM$(razonSocial$) = "" Then razonSocial$ = AJUSTI$(rasocli$(NumeroProveedor%), 30)
            LargoRS% = InStr(1, razonSocial$, "//")
            If LargoRS% > 0 Then razonSocial$ = AJUSTI$(Mid(razonSocial$, 1, LargoRS% - 1), 30)
            If PIVACLI%(NumeroProveedor%) = 0 Then
                razonSocial$ = AJUSTI$("CONSUMIDOR FINAL", 30)
            End If
            
                MONE% = XVALO(VALOBADA("CAJABANC", "Mone_Emis", "CODICOM_LAR='" & NumeroComprobante$ & "' AND FechConta='" & fechaCbte$ & "'"))
            CODMONEDA$ = CodigoMoneda(MONE%)
            If CODMONEDA$ = "000" Then CODMONEDA$ = "PES"
            
                TCAMBIO# = XVALO(VALOBADA("CAJABANC", "Tipo_Cam", "CODICOM_LAR='" & NumeroComprobante$ & "' AND FechConta='" & fechaCbte$ & "'"))
            TIPOCAMBIO$ = RIVFormatoNumero$(TCAMBIO#, 10, 6)
            If CODMONEDA$ = "PES" And TCAMBIO# > 1 Then
                    ListaDeErrores$ = ListaDeErrores$ & "Error en Tipo de Cambio en el Proveedor '" & CStr(Abs(NumeroProveedor%)) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                    FlagError% = 1: GoTo 35
            End If
            If XVALO(TIPOCAMBIO$) = 0 And CODMONEDA$ = "PES" Then TIPOCAMBIO$ = RIVFormatoNumero(1, 10, 6)

            TipoCambioCbte# = 1
            If MONE% > 1 Then TipoCambioCbte# = TCAMBIO#

            ImporteTotal$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 91, 8)) / TipoCambioCbte#, 15)
            
            ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
            ValorIva$ = TRIM$(Mid$(REGISTRO$, 129, 8))
            LetraComprobante$ = Mid$(NumeroComprobante$, 4, 1)
            If InStr(1, UCase(ValorIva$), "NO GRAV") > 0 And (LetraComprobante$ <> "B" And LetraComprobante$ <> "C") Then
                ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
            End If
            If InStr(1, "011;012;013;006;007;008;019;020;021", TipoCbte$) > 0 Then ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
            
            PercepcionNoCategorizados$ = RIVFormatoNumero$(0, 15)

            ImporteOperacionExenta$ = RIVFormatoNumero$(0, 15)
            If PIVACLI%(NumeroProveedor%) = 5 Or PIVACLI%(NumeroProveedor%) = 7 Then
                ImporteOperacionExenta$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
            End If
            
            ImportePercepcionPagosACuenta$ = RIVFormatoNumero$(CVD(Mid$(REGISTRO$, 161, 8)) / TipoCambioCbte#, 15)
            
            ' Percepciones de iva + ganancias
                PercepcionesVarias# = CVD(Mid$(REGISTRO$, 169, 8)) / TipoCambioCbte#
            ImportePercepcionesNacionales$ = RIVFormatoNumero$(PercepcionesVarias#, 15)
            
            ImportePercepcionIB$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 193, 8)) / TipoCambioCbte#, 15)
            
            ImportePercepcionMunicipales$ = RIVFormatoNumero$(0, 15)
            
            ImporteImpuestosInternos$ = RIVFormatoNumero$(0, 15)
                
                CANTALICUOTAS% = 1
            If InStr(1, "006;007;008;009;082;013;012;011;015;013;111", TipoCbte$) > 0 Then CANTALICUOTAS% = 0 ' Si es Comprobante "C" la cantidad debe ser cero
            CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
            
            Percepciones# = XVALO(ImportePercepcionesNacionales$) / 100 + XVALO(ImportePercepcionIB$) / 100 + XVALO(ImportePercepcionPagosACuenta$) / 100
            If PIVACLI%(NumeroProveedor%) = 7 Then
                CodigoOperacion$ = "Z"
            ElseIf PIVACLI%(NumeroProveedor%) = 5 Or (TRIM$(NumeroDespachoImportacion$) <> "" Or EsFacturaDespachante# > 0.005) Then
                CodigoOperacion$ = "X"
            ElseIf InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 Or (TRIM$(Mid$(REGISTRO$, 129, 8)) = "" And Percepciones# > 0.005) Then
                CodigoOperacion$ = "N"
            ElseIf PIVACLI%(NumeroProveedor%) = 4 Or InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "EXENTO") > 0 Then
                CodigoOperacion$ = "E"
            ElseIf CANTALICUOTAS% = 0 Then
                CodigoOperacion$ = "0"
            Else
                CodigoOperacion$ = " "
            End If
                
            ImporteIva$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
            
            OtrosTributos$ = RIVFormatoNumero$(0, 15)
            
            CuitEmisor$ = "00000000000"
            
            DenominacionEmisor$ = AJUSTI$(" ", 30)
            
            IvaComision$ = RIVFormatoNumero("0", 15)
            
            ' Arma el registro para el regimen informativo
            RegistroRIV$ = FechaComprobante$ & TipoCbte$ & PUNTOVTA$ & "00000000000000000000" & NumeroDespachoImportacion$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & razonSocial$ _
                               & ImporteTotal$ & ImporteNoGravado$ & ImporteOperacionExenta$ & ImportePercepcionPagosACuenta$ & ImportePercepcionesNacionales$ _
                               & ImportePercepcionIB$ & ImportePercepcionMunicipales$ & ImporteImpuestosInternos$ & CODMONEDA$ & TIPOCAMBIO$ & CantidadAlicuotas$ _
                               & CodigoOperacion$ & ImporteIva$ & OtrosTributos$ & CuitEmisor$ & DenominacionEmisor$ & IvaComision$
                               
            UltimoNumeroComprobante$ = NumeroComprobante$
            UltimoCuitProveedor$ = CUITPROVEEDOR$

            'If (InStr(1, UCase(ValorIva$), "NO GRAV") = 0 And XVALO(ImporteIva$) > 0) Then
            If XVALO(CantidadAlicuotas$) > 0 Then
            
                ' Tratamiento para los comprobantes con parte no gravada o exenta
                If TRIM$(CodigoOperacion$) <> "" Then
                    ImporteNeto$ = RIVFormatoNumero(0, 15)
                    ALICUOTA$ = "0003"
                    ImporteIva0$ = RIVFormatoNumero(0, 15)

                    If (CodigoOperacion$ = "N") Then
                        ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%))) / TipoCambioCbte#
                        Percepciones# = XVALO(ImportePercepcionesNacionales$) / 100 + XVALO(ImportePercepcionIB$) / 100 + XVALO(ImportePercepcionPagosACuenta$) / 100
                        ImporteEx# = XVALO(VALOBADA("CAJABANC", "Sum(BaseImpo)", "VariMovim like 'Exento' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%))) / TipoCambioCbte#
                        If ConIva# < 0.01 Or (Percepciones# > 0.005 And ConIva# < 0.01) Or (ImporteEx# > 0.005 And ConIva# < 0.01) Then
                            ' cuando el comprobante es totalmente no gravado, dicho importe se escribe en el archivo de alicuotas
                            ImporteNeto$ = ImporteTotal$
                            If Percepciones# > 0.005 Or ImporteEx# > 0.005 Then ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                            ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
                            Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                        Else
                            ' Cuando el comprobante tiene parte gravada y parte que no lo esta, el gravado se escribe
                            ' en el archivo de comprobantes y los importes gravados en el de alicuotas
                            CANTALICUOTAS% = 0
                            CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                            Call REPLA(RegistroRIV$, " ", 239, 1)
                            Call REPLA(RegistroRIV$, CantidadAlicuotas$, 238, 1)
                            GoTo 35
                        End If
                    End If
                    If CodigoOperacion$ = "E" Then
                        ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%))) / TipoCambioCbte#
                        Percepciones# = XVALO(ImportePercepcionesNacionales$) / 100 + XVALO(ImportePercepcionIB$) / 100 + XVALO(ImportePercepcionPagosACuenta$) / 100
                        NoGravado# = XVALO(VALOBADA("CAJABANC", "Sum(BaseImpo)", "VariMovim like 'No Grav%' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%))) / TipoCambioCbte#
                        If (ConIva# < 0.005 Or (Percepciones# > 0.005 And ConIva# < 0.01)) And NoGravado# < 0.005 Then
                            ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                            ImporteOperacionExenta$ = RIVFormatoNumero$(0, 15)
                            Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 135, 15)
                        ElseIf NoGravado# > 0.005 Then
                            ImporteOperacionExenta$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                            Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 135, 15)
                            CANTALICUOTAS% = 0
                            CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                            If Mid$(NumeroComprobante$, 1, 2) = "ND" Or Mid$(NumeroComprobante$, 1, 2) = "FC" Then
                                ImporteNeto$ = RIVFormatoNumero(0, 15)
                                ImporteIva0$ = RIVFormatoNumero(0, 15)
                                ALICUOTA$ = "0003"
                                CANTALICUOTAS% = 1
                                CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                            Else
                                GoTo 35
                            End If
                        Else
                            CANTALICUOTAS% = 0
                            CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                            Call REPLA(RegistroRIV$, " ", 239, 1)
                            Call REPLA(RegistroRIV$, CantidadAlicuotas$, 238, 1)
                            ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                            Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                            GoTo 35
                        End If
                    End If
                    If CodigoOperacion$ = "X" And (TRIM$(NumeroDespachoImportacion$) <> "" Or EsFacturaDespachante# > 0.005) Then
                        If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") > 0 Then
                            ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%))) / TipoCambioCbte#
                            ImpoNoGra# = XVALO(VALOBADA("CAJABANC", "Sum(BaseImpo)", "VariMovim='Aduana' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%))) / TipoCambioCbte#
                            ImpoAduana# = XVALO(ImporteIva$) / 100
                            OtrosTributos$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)) / TipoCambioCbte#, 15)
                            ImporteIva$ = RIVFormatoNumero(0, 15)
                            Call REPLA(RegistroRIV$, ImporteIva$, 240, 15)
                            CantidadAlicuotas$ = "0"
                            If ImpoNoGra# > 0.005 Then
                                ImporteNGAnterior = XVALO(Mid$(RegistroRIV$, 120, 15)) / 100
                                ImpoNoGra# = ImpoNoGra# + ImporteNGAnterior
                                ImporteNoGravado$ = RIVFormatoNumero(ImpoNoGra#, 15, , 1)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                            End If
                            If (ConIva# < 0.005 And EsRegimenSimplificado% = 0) Then
                                Call REPLA(RegistroRIV$, RIVFormatoNumero(ImpoAduana#, 15), 150, 15)
                                ImporteNeto$ = RIVFormatoNumero(0, 15)
                                ImporteIva0$ = RIVFormatoNumero(0, 15)
                                ALICUOTA$ = "0003"
                                CantidadAlicuotas$ = "1"
                            Else
                                If MODOASEMBLI% > 0 And EsRegimenSimplificado% = 1 Then
                                    RegistroRIV_IvaAduana$ = String$(65, "0") & "0003" & RIVFormatoNumero(ImpoAduana#, 15)
                                End If
                                GoTo 35
                            End If
                        ElseIf InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "EXENTO") > 0 Then
                            'ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                            ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%))) / TipoCambioCbte#
                            If ConIva# > 0.005 Then
                                ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                CANTALICUOTAS% = 0
                                CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                                GoTo 35
                            Else
                                ImporteNoGravado$ = RIVFormatoNumero(0, 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                            End If
                        ElseIf InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 Then
                            ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%))) / TipoCambioCbte#
                            If ConIva# > 0.005 Then
                                ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                CANTALICUOTAS% = 0
                                CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
                                GoTo 35
                            Else
                                ImporteNoGravado$ = RIVFormatoNumero(0, 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                            End If
                        ElseIf InStr(1, "0004;0005;0006;0008;0009", CodigoAlicuota(Mid$(REGISTRO$, 129, 8))) > 0 Then
                            ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                            ImporteIva0$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)) / TipoCambioCbte#, 15)
                            ALICUOTA$ = CodigoAlicuota(ValorIva$)
                        End If
'                        If CodigoOperacion$ = "X" Then
'                            If CantRegistros("CAJABANC", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%) & " AND IDEBECOMP>0 AND TIPOMOVIM='COMPRA'") > 1 Then
'                                CANTALICUOTAS% = 0
'                                CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
'                                Call REPLA(RegistroRIV$, " ", 239, 1)
'                                Call REPLA(RegistroRIV$, CantidadAlicuotas$, 238, 1)
'                                GoTo 35
'                            Else
'                                ImporteOperacionExenta$ = RIVFormatoNumero(0, 15)
'                                Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 135, 15)
'                                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
'                            End If
'                        End If
                    ElseIf CodigoOperacion$ = "X" Then
'                        If CantRegistros("CAJABANC", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%) & " AND IDEBECOMP>0 AND TIPOMOVIM='COMPRA'") > 1 Then
'                                CANTALICUOTAS% = 0
'                                CantidadAlicuotas$ = CStr(CANTALICUOTAS%)
'                                Call REPLA(RegistroRIV$, " ", 239, 1)
'                                Call REPLA(RegistroRIV$, CantidadAlicuotas$, 238, 1)
'                                GoTo 35
'                        Else
                                ImporteOperacionExenta$ = RIVFormatoNumero(0, 15)
                                Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 135, 15)
                                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
'                        End If
                    End If
                Else
                    ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                    ImporteIva0$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)) / TipoCambioCbte#, 15)
                    ALICUOTA$ = CodigoAlicuota(ValorIva$)
                End If
                
                If (TRIM$(NumeroDespachoImportacion$) = "" And EsFacturaDespachante# < 0.005) Or EsRegimenSimplificado% = 1 Then
                
                    RegistroRIV_Alicuotas$ = TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & ImporteNeto$ & ALICUOTA$ & ImporteIva0$
                    If MODOASEMBLI% = 1 And EsRegimenSimplificado% = 1 Then
                        RegistroRIV_IvaAduana$ = RegistroRIV_Alicuotas$
                    Else
                        Print #NX1%, RegistroRIV_Alicuotas$
                    End If
                Else
                    
                    RegistroRIV_Importaciones$ = NumeroDespachoImportacion$ & ImporteNeto$ & ALICUOTA$ & ImporteIva0$
                    Print #NX2%, RegistroRIV_Importaciones$
                End If
            End If
        Else
            If FlagError% = 1 Then GoTo 35
            If TipoCbte$ = "006" Then GoTo 35
            
            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 Then
                ImporteNGAnterior = XVALO(Mid$(RegistroRIV$, 120, 15)) / 100
                ImporteNGActual = CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#
                'If Not (ImporteNGActual < 0 And Mid$(REGISTRO$, 11, 2) = "FC") Then
                If Mid$(REGISTRO$, 11, 2) <> "NC" Then
                    'If A <> 1 Then ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior + ImporteNGActual, 15, , 1)
                    ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior + ImporteNGActual, 15, , 1)
                Else
                    ConsideraSigno% = 0
                    If ImporteNGActual > 0 Then ' INCORPORADO PARA EL VALOR CIF
                        ImporteNGActual = (-1) * ImporteNGActual
                        ConsideraSigno% = 1
                    End If
                    ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior + ImporteNGActual, 15, , ConsideraSigno%)
                End If
                'End If
            End If
            
            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "EXENTO") > 0 Then
                ImporteExAnterior = XVALO(Mid$(RegistroRIV$, 135, 15)) / 100
                ImporteExActual = CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#
                ImporteOperacionExenta$ = RIVFormatoNumero(ImporteExAnterior + ImporteExActual, 15)
            End If
            
            If PIVACLI%(NumeroProveedor%) = 5 Or PIVACLI%(NumeroProveedor%) = 7 Then
                IMPO# = XVALO(Mid$(RegistroRIV$, 135, 15)) / 100
                ImporteOperacionExenta$ = RIVFormatoNumero(IMPO# + CVD(Mid$(REGISTRO$, 185, 8)) / TipoCambioCbte#, 15)
            End If
                                                                             
                IMPO# = XVALO(Mid$(RegistroRIV$, 150, 15)) / 100
            ImportePercepcionPagosACuenta$ = RIVFormatoNumero$(IMPO# + CVD(Mid$(REGISTRO$, 161, 8)) / TipoCambioCbte#, 15)
            
                IMPO# = XVALO(Mid$(RegistroRIV$, 165, 15)) / 100
            ImportePercepcionesNacionales$ = RIVFormatoNumero$(IMPO# + CVD(Mid$(REGISTRO$, 169, 8)) / TipoCambioCbte#, 15)

                IMPO# = XVALO(Mid$(RegistroRIV$, 180, 15)) / 100
            ImportePercepcionIB$ = RIVFormatoNumero(IMPO# + CVD(Mid$(REGISTRO$, 193, 8)) / TipoCambioCbte#, 15)
                                
            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") = 0 Then
                    IMPO# = XVALO(Mid$(RegistroRIV$, 240, 15)) / 100
                    ImporteIvaActual# = Abs(CVD(Mid$(REGISTRO$, 145, 8)))
                    If A = 0 Then
                        ImporteIva$ = RIVFormatoNumero(IMPO# + ImporteIvaActual#, 15)
                    Else
                        If ImporteIvaActual# > 0.005 Then
                            IMPO# = XVALO(Mid$(RegistroRIV$, 150, 15)) / 100
                            ImportePercepcionPagosACuenta$ = RIVFormatoNumero(IMPO# + ImporteIvaActual#, 15)
                        End If
                    End If
            End If

                CANTALICUOTAS% = XVALO(CantidadAlicuotas$)
            ImporteNeto0# = CVD(Mid$(REGISTRO$, 137, 8))
            ValorIva$ = TRIM$(REPLACAR$(Mid$(REGISTRO$, 129, 8), "%", ""))
            CantidadAlicuotaNuevoRegistro% = 0
            If (Not (ImporteNeto0# < 0 And Mid$(REGISTRO$, 11, 2) = "FC")) And XVALO(ImporteIvaActual#) > 0.005 And InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") = 0 Then CantidadAlicuotaNuevoRegistro% = 1
'            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 And (CodigoOperacion$ = "X" Or CodigoOperacion$ = "E") And _
'                CantRegistros("CAJABANC", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%) & " AND IDEBECOMP>0 AND (TIPOMOVIM='COMPRA' or TIPOMOVIM='REDCOM')") < 2 Then
'
'                CantidadAlicuotaNuevoRegistro% = 1
'                ImporteNGAnterior = XVALO(ImporteNoGravado$) / 100
'                ImporteNGActual = CVD(Mid$(REGISTRO$, 137, 8))
'                ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior - ImporteNGActual, 15)
'            End If
            CantidadAlicuotas$ = CStr(CANTALICUOTAS% + CantidadAlicuotaNuevoRegistro%)
            
            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") > 0 Then
'                IMPO# = XVALO(Mid$(RegistroRIV$, 255, 15)) / 100
'                OtAtributos# = CVD(Mid$(REGISTRO$, 145, 8))
'                OTROSTRIBUTOS$ = RIVFormatoNumero(IMPO# + OtAtributos#, 15)
'                IMPO# = XVALO(ImporteIva$) / 100
'                ImporteIva$ = RIVFormatoNumero(IMPO# - OtAtributos#, 15)
                IMPO# = XVALO(ImportePercepcionPagosACuenta$) / 100
                ImporteAduana# = CVD(Mid$(REGISTRO$, 145, 8)) / TipoCambioCbte#
                If MODOASEMBLI% = 0 Or (MODOASEMBLI% = 1 And EsRegimenSimplificado% = 0) Then
                    ImportePercepcionPagosACuenta$ = RIVFormatoNumero(IMPO# + ImporteAduana#, 15)
                End If
                ' SI NO TIENE ALICUOTA PREVIA, AGREGA UN REGISTRO EN EL ARCHIVO DE ALICUOTA. ES FICTICIO PORQUE
                ' TANTO EL NETO GRAVADO COMO EL IMPORTE DE IVA, ESTAN EN CERO.
                ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                If ConIva# < 0.005 And (MODOASEMBLI% = 0 Or (MODOASEMBLI% = 1 And EsRegimenSimplificado% = 0)) Then
                    CantidadAlicuotaNuevoRegistro% = 1
                    CantidadAlicuotas$ = CStr(CANTALICUOTAS% + CantidadAlicuotaNuevoRegistro%)
                End If
                
                If MODOASEMBLI% = 1 And EsRegimenSimplificado% = 1 And RegistroRIV_IvaAduana$ <> "" Then
                    CantidadAlicuotaNuevoRegistro% = 1
                End If
            End If
                        
            RegistroRIV$ = FechaComprobante$ & TipoCbte$ & PUNTOVTA$ & "00000000000000000000" & NumeroDespachoImportacion$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & razonSocial$ _
                           & ImporteTotal$ & ImporteNoGravado$ & ImporteOperacionExenta$ & ImportePercepcionPagosACuenta$ & ImportePercepcionesNacionales$ _
                           & ImportePercepcionIB$ & ImportePercepcionMunicipales$ & ImporteImpuestosInternos$ & CODMONEDA$ & TIPOCAMBIO$ & CantidadAlicuotas$ _
                           & CodigoOperacion$ & ImporteIva$ & OtrosTributos$ & CuitEmisor$ & DenominacionEmisor$ & IvaComision$
            
            If (CantidadAlicuotaNuevoRegistro% > 0) Then
                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) / TipoCambioCbte#, 15)
                ImporteIva0$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)) / TipoCambioCbte#, 15)
                ALICUOTA$ = CodigoAlicuota(ValorIva$)
                If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 And (CodigoOperacion$ = "X" Or CodigoOperacion$ = "E") Then
                    ALICUOTA$ = "0003"
                End If
                If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") > 0 And (MODOASEMBLI% = 0 Or (MODOASEMBLI% = 1 And EsRegimenSimplificado% = 0)) Then
                    ImporteNeto$ = RIVFormatoNumero(0, 15)
                    ImporteIva0$ = RIVFormatoNumero(0, 15)
                End If
                If A = 1 Then
                    ALICUOTA$ = "0003"
                    ImporteNeto$ = RIVFormatoNumero(0, 15)
                    ImporteIva0$ = RIVFormatoNumero(0, 15)
                End If
                
                If (TRIM$(NumeroDespachoImportacion$) = "" And EsFacturaDespachante# < 0.005) Or EsRegimenSimplificado% = 1 Then
                
                    RegistroRIV_Alicuotas$ = TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & ImporteNeto$ & ALICUOTA$ & ImporteIva0$
                    If MODOASEMBLI% = 1 And EsRegimenSimplificado% = 1 Then
                        If RegistroRIV_IvaAduana$ <> "" Then        ' EL IMPORTE DEL IVA ADUANA SE AGREGA AL CREDITO FISCAL DE ALGUNA DE LAS ALICUOTAS DEL COMPROBANTE
                            If Mid$(RegistroRIV_IvaAduana$, 66, 4) = "0003" Then
                                ImpoAduana# = XVALO(Mid$(RegistroRIV_IvaAduana$, 70, 15)) / 100
                                IMPOIVA# = XVALO(Mid$(RegistroRIV_Alicuotas$, 70, 15)) / 100
                                IMPOIVA# = IMPOIVA# + ImpoAduana#
                                RegistroRIV_Alicuotas$ = Mid$(RegistroRIV_Alicuotas$, 1, 69) & RIVFormatoNumero(IMPOIVA#, 15)
                            ElseIf ALICUOTA$ = "0003" Then
                                ImpoAduana# = XVALO(Mid$(RegistroRIV_IvaAduana$, 70, 15)) / 100
                                IMPOIVA# = XVALO(Mid$(RegistroRIV_Alicuotas$, 70, 15)) / 100
                                IMPOIVA# = IMPOIVA# + ImpoAduana#
                                RegistroRIV_Alicuotas$ = Mid$(RegistroRIV_IvaAduana$, 1, 69) & RIVFormatoNumero(IMPOIVA#, 15)
                            End If
                            RegistroRIV_IvaAduana$ = ""
                        Else
                            If RegistroRIV_IvaAduana$ = "" Then
                                RegistroRIV_IvaAduana$ = RegistroRIV_Alicuotas$
                                GoTo 35
                            End If
                        End If
                    End If
                    Print #NX1%, RegistroRIV_Alicuotas$
                Else
                    
                    RegistroRIV_Importaciones$ = NumeroDespachoImportacion$ & ImporteNeto$ & ALICUOTA$ & ImporteIva0$
                    Print #NX2%, RegistroRIV_Importaciones$
                End If
            End If
        End If
35  Next i&
    
    If RegistroRIV$ <> "" Then Print #NX%, RegistroRIV$
    
    Close #NX%
    'Close #NX1%
    Close #NX2%
    
'    DespachosProcesados# = 0
'
'    NX_1% = FILEOPEN%(LUCOM$ & "xxLIBRO_IVA_DIGITAL_IMPORTACION_BIENES_CBTE_" & PERIODO$ & ".TXT", 1, 266)
'    For IX& = 1 To LOF(NX_1%)
'        Get NX_1%, IX&, XX$
'        Despacho1$ = Mid$(XX$, 37, 16)
'
'        If InStr(1, DespachosProcesados#, Despacho1$) = 0 Then
'            NX_2% = FILEOPEN%(LUCOM$ & "xxLIBRO_IVA_DIGITAL_IMPORTACION_BIENES_CBTE_" & PERIODO$ & ".TXT", 1, 266)
'            For ir& = IX& To LOF(NX_2%)
'                Get NX_2%, ir&, YY$
'                Despacho2$ = Mid$(YY$, 37, 16)
'
'                If Despacho1$ = Despacho2$ Then
'                    Total# = 0
'                    NoGravado# = 0
'                    Exento# = 0
'                    PercepcionesIva# = 0
'                    PercepcionesOtros# = 0
'                    PercepcionesIIBB# = 0
'                    ImpuestosMunicipales# = 0
'                    ImpuestosInternos# = 0
'                    CantidadAlicuotas% = 0
'                    CreditoFiscal# = 0
'                    OtrosTributos# = 0
'                End If
'
'            Next ir&
'        End If
'
'    Next IX&
'
'    NX_2% = FILEOPEN%(LUCOM$ & "xxLIBRO_IVA_DIGITAL_IMPORTACION_BIENES_CBTE_" & PERIODO$ & ".TXT", 1, 266)
    
    Screen.MousePointer = 1
    
'    Dim XX1 As String * 266
'
'    ARCHIIBB$ = "REGINFO_CV_COMPRAS_CBTE_" & PERIODO$ & ".TXT"
'    NARCHI% = FILEOPEN%(LUCOM$ + ARCHIIBB$, 1, 266)
'    Do While Not EOF(NARCHI%)
'        Line Input #NARCHI%, XX$
'        SIGNO% = 1
'        If Mid$(XX$, 9, 3) = "003" Or Mid$(XX$, 9, 3) = "008" Or Mid$(XX$, 9, 3) = "021" Or Mid$(XX$, 9, 3) = "013" Or Mid$(XX$, 9, 3) = "053" Then SIGNO% = -1
'
'        IMPO1# = IMPO1# + SIGNO% * XVALO(Mid$(XX$, 105, 15)) / 100
'        impo2# = impo2# + SIGNO% * XVALO(Mid$(XX$, 120, 15)) / 100
'        impo3# = impo3# + SIGNO% * XVALO(Mid$(XX$, 150, 15)) / 100
'        impo4# = impo4# + SIGNO% * XVALO(Mid$(XX$, 135, 15)) / 100
'        impo5# = impo5# + SIGNO% * XVALO(Mid$(XX$, 165, 15)) / 100
'        impo6# = impo6# + SIGNO% * XVALO(Mid$(XX$, 180, 15)) / 100
'        impo7# = impo7# + SIGNO% * XVALO(Mid$(XX$, 195, 15)) / 100
'        impo8# = impo8# + SIGNO% * XVALO(Mid$(XX$, 210, 15)) / 100
'        impo9# = impo9# + SIGNO% * XVALO(Mid$(XX$, 255, 15)) / 100
'        impo10# = impo10# + SIGNO% * XVALO(Mid$(XX$, 240, 15)) / 100
'20  Loop
'
'    MsgBox ("Total: " & TRIM$(FORMATNUM$(IMPO1#, "F15.2")) & vbCrLf _
'        & "No gravado: " & TRIM$(FORMATNUM$(impo2#, "F15.2")) & vbCrLf _
'        & "Iva: " & TRIM$(FORMATNUM$(impo10#, "F15.2")) & vbCrLf _
'        & "Percep. Iva: " & TRIM$(FORMATNUM$(impo3#, "F15.2")) & vbCrLf _
'        & "Op.Exentas: " & TRIM$(FORMATNUM$(impo4#, "F15.2")) & vbCrLf _
'        & "Imp.Nacionales: " & TRIM$(FORMATNUM$(impo5#, "F15.2")) & vbCrLf _
'        & "Imp.IIBB: " & TRIM$(FORMATNUM$(impo6#, "F15.2")) & vbCrLf _
'        & "Imp.Municipales: " & TRIM$(FORMATNUM$(impo7#, "F15.2")) & vbCrLf _
'        & "Imp.Internos: " & TRIM$(FORMATNUM$(impo8#, "F15.2")) & vbCrLf _
'        & "Otro Atributos: " & TRIM$(FORMATNUM$(impo9#, "F15.2")) & vbCrLf)
'
'    Close #NARCHI%
'
'    ARCHIIBB$ = "REGINFO_CV_COMPRAS_ALICUOTAS_" & PERIODO$ & ".TXT"
'    NARCHI1% = FILEOPEN%(LUCOM$ + ARCHIIBB$, 1, 62)
'    Do While Not EOF(NARCHI1%)
'        Line Input #NARCHI1%, XX$
'        SIGNO% = 1
'        If Mid$(XX$, 1, 3) = "003" Or Mid$(XX$, 1, 3) = "008" Or Mid$(XX$, 1, 3) = "021" Then
'              SIGNO% = -1
'        End If
'        IMPO20# = IMPO20# + SIGNO% * XVALO(Mid$(XX$, 51, 15)) / 100
'        impo21# = impo21# + SIGNO% * XVALO(Mid$(XX$, 70, 15)) / 100
'    Loop
'
'    MsgBox ("Neto: " & TRIM$(FORMATNUM$(IMPO20#, "F15.2")) & vbCrLf _
'        & "Iva: " & TRIM$(FORMATNUM$(impo21#, "F15.2")))
'
'    Close #NARCHI1%
    
    If TRIM$(ListaDeErrores$) <> "" Then
        Call COMUNI("Se Han Detectado Errores en la Generación\del IVA Digital de Bienes Importados.\Pulse Aceptar para Ver el Detalle")
        MSJESREGINFO.Text1 = ListaDeErrores$
        MSJESREGINFO.Caption = "Régimen Informativo de Compras"
        Call MSJESREGINFO.Command1_Click
        MSJESREGINFO.Show 1
        Call MENSERR(24, "Tener Presente que los Datos Generados Son Incompletos.\Debe Corregir los Errores y Volver a Generar")
    End If
    
99  Exit Sub
    
    
InicializarCamposVGD:
    
    VGD_FechaComprobante$(INDICE%) = "00000000"
    VGD_PuntoVta$(INDICE%) = ""
    VGD_NumeroCbtDesde$(INDICE%) = "0"
    VGD_NumeroCbtHasta$(INDICE%) = "0"
    VGD_ImporteTotal$(INDICE%) = "0"
    VGD_ImporteNoGravado$(INDICE%) = "0"
    VGD_ImporteOperacionExenta$(INDICE%) = "0"
    VGD_ImportePercepcionesNacionales$(INDICE%) = "0"
    VGD_ImportePercepcionIB$(INDICE%) = "0"
    VGD_CantidadAlicuotas$(INDICE%) = "0"
    VGD_RegistroRIV$(INDICE%) = ""

    Return
    
End Sub

Function CodigoAlicuota(PORCENTAJE$) As String

    Dim Codigo$, Porce$
    Porce$ = REPLACAR$(PORCENTAJE$, "%", "")
    If (XVALO(Porce$) = 10.5) Then
        Codigo$ = "0004"
    ElseIf (XVALO(Porce$) = 21) Then
        Codigo$ = "0005"
    ElseIf (XVALO(Porce$) = 27) Then
        Codigo$ = "0006"
    ElseIf (XVALO(Porce$) = 5) Then
        Codigo$ = "0008"
    ElseIf (XVALO(Porce$) = 2.5) Then
        Codigo$ = "0009"
    ElseIf (UCase(Porce$) = "EXENTO") Or (UCase(Porce$) = "ADUANA") Then
        Codigo$ = "0003"
    ElseIf (UCase(Porce$) = "NO GRAV") Then
        Codigo$ = "0001"
    End If
    
    CodigoAlicuota = Codigo$
    
End Function

Function AlicuotaSegunCodigo(Codigo$) As String

    Dim Porce$
    If (XVALO(Codigo$) = 4) Then
        Porce$ = "10.50%"
    ElseIf (XVALO(Codigo$) = 5) Then
        Porce$ = "21.00%"
    ElseIf (XVALO(Codigo$) = 6) Then
        Porce$ = "27.00%"
    ElseIf (XVALO(Codigo$) = 8) Then
        Porce$ = "5.00%"
    ElseIf (XVALO(Codigo$) = 9) Then
        Porce$ = "2.50%"
    ElseIf (XVALO(Codigo$) = 3) Then
        Porce$ = "Exento/Aduana"
    ElseIf (XVALO(Codigo$) = 1) Then
        Porce$ = "No Gravado"
    End If
    
    AlicuotaSegunCodigo = Porce$
    
End Function
Function RIVFormatoNumero(Numero#, Optional Longitud%, Optional Decimales%, Optional ConsiderarSigno%) As String
    
    If Longitud <= 0 Then Longitud = 15
    If Decimales% <= 0 Then Decimales% = 2
    'VALOR$ = TRIM$(FORMATNUM$(Abs(NUMERO#), "F" & CStr(Longitud%) & "." & Decimales%))
    valor$ = TRIM$(FORMATNUM$(Abs(Numero#), "F" & CStr(Longitud%) & "." & Decimales%))
    valor$ = REPLACAR$(valor$, ".", "")
    If valor$ = "" Then valor$ = "0"
    
    If ConsiderarSigno% = 0 Then
        Do While Len(valor$) < Longitud%
            valor$ = "0" & valor$
        Loop
    Else
        If Numero# >= 0 Then
            Do While Len(valor$) < Longitud%
                valor$ = "0" & valor$
            Loop
        Else
            'VALOR$ = Mid$(VALOR$, 2)
            Do While Len(valor$) < (Longitud% - 1)
                valor$ = "0" & valor$
            Loop
            valor$ = "-" & valor$
        End If
    End If
    
    RIVFormatoNumero = valor$
    
'    If Longitud <= 0 Then Longitud = 15
'    If Decimales% <= 0 Then Decimales% = 2
'    VALOR$ = TRIM$(FORMATNUM$(Abs(NUMERO#), "F" & CStr(Longitud%) & "." & Decimales%))
'    VALOR$ = REPLACAR$(VALOR$, ".", "")
'    If VALOR$ = "" Then VALOR$ = "0"
'
'    Do While Len(VALOR$) < Longitud%
'        VALOR$ = "0" & VALOR$
'    Loop
'
'    RIVFormatoNumero = VALOR$
    
End Function

Function TipoComprobante(NumeroComprobante$) As String

    IdCbte$ = Mid$(NumeroComprobante$, 1, 2) & Mid$(NumeroComprobante$, 4, 1)
    
    Tipo$ = ""
    Select Case IdCbte$
        Case "FCA": Tipo$ = "001"
        Case "NDA": Tipo$ = "002"
        Case "NCA": Tipo$ = "003"
        Case "NVA": Tipo$ = "005"
        Case "FCB": Tipo$ = "006"
        Case "NDB": Tipo$ = "007"
        Case "NCB": Tipo$ = "008"
        Case "NVB": Tipo$ = "010"
        Case "FCE": Tipo$ = "019"
        Case "NDE": Tipo$ = "020"
        Case "NCE": Tipo$ = "021"
        Case "FCC": Tipo$ = "011"
        Case "NDC": Tipo$ = "012"
        Case "NCC": Tipo$ = "013"
        Case "NVC": Tipo$ = "016"
        Case "RHA": Tipo$ = "004"
        Case "RCA": Tipo$ = "004"
        Case "RHB": Tipo$ = "009"
        Case "RCB": Tipo$ = "009"
        Case "RHC": Tipo$ = "015"
        Case "RCC": Tipo$ = "015"
        Case "RHM": Tipo$ = "054"
        Case "CPA": Tipo$ = "002"
        Case "CPB": Tipo$ = "007"
        Case "CPC": Tipo$ = "012"
        Case "CPM": Tipo$ = "052"
        Case "FCM": Tipo$ = "051"
        Case "NDM": Tipo$ = "052"
        Case "NCM": Tipo$ = "053"
        Case "TFA": Tipo$ = "081"
        Case "TFB": Tipo$ = "082"
        Case "TFC": Tipo$ = "111"
        Case "TFM": Tipo$ = "118"
        Case "COA": Tipo$ = "039"
        Case "COB": Tipo$ = "040"
        Case "COC": Tipo$ = "041"
        Case "COM": Tipo$ = "057"
        Case "DPA": Tipo$ = "003"
        Case "DPB": Tipo$ = "008"
        Case "DPC": Tipo$ = "013"
        Case "DPM": Tipo$ = "053"
        Case Else: Tipo$ = "001"
    End Select
        
    ' Posible asiento contable
    If TipoCbte$ = "" And Mid$(NumeroComprobante$, 7, 1) = "-" Then TipoCbte$ = "039"
    
    TipoComprobante$ = Tipo$
    
End Function

Function CodigoDocumento(PosicionIva%) As String

    Select Case PosicionIva%
        Case 0: CodiDoc$ = "96"
        Case Else: CodiDoc$ = 80
    End Select
    
    CodigoDocumento$ = CodiDoc$
    
End Function

Function CodigoMoneda(Moneda%) As String

    Select Case Moneda%
        Case 1: CodiMo$ = "PES"
        Case 2: CodiMo$ = "DOL"
        Case 3: CodiMo$ = "060"
        Case Else: CodiMo$ = "000"
    End Select
    
    CodigoMoneda = CodiMo$
    
End Function

Function NumeroCbte(NumeroComprobante$) As String

    NumeroComprobante$ = TRIM$(NumeroComprobante$)
    NroComprobante$ = REPLACAR$(NumeroComprobante$, "-", "")
    NroComprobante$ = SOLONUM$(Mid$(NroComprobante$, 8, 8))
    
    If Len(TRIM$(NumeroComprobante$)) < 18 Then NroComprobante$ = TRIM$(REPLACAR$(NumeroComprobante$, "-", ""))
    
    Do While Len(NroComprobante$) < 20
        NroComprobante$ = "0" & NroComprobante$
    Loop
    
    NumeroCbte = NroComprobante$
    
End Function

Function PuntoVenta(NumeroComprobante$) As String

    If Mid$(NumeroComprobante$, 1, 2) = "TC" Then
        PuntoVenta$ = "00000": Exit Function
    End If
    
    PVenta$ = Mid$(NumeroComprobante$, 6, 4)
    PVenta$ = SOLONUM$(PVenta$)
    
    If IsNumeric(PVenta$) = False Then PVenta$ = "0001"
    
    While Len(PVenta$) < 5: PVenta$ = "0" & PVenta$: Wend
    PuntoVenta$ = PVenta$
    
End Function
Private Sub Command12_Click()
    '
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
    End If
    Call FRESHALL
    '
End Sub

Sub Command13_Click()
    Command13.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Option2.Value = True
    Call GENIVACOM
    FIN& = ULTREG&("REIVACOM")
    If FIN& < 1 Then GoTo 99
    '
    OPX% = COMALTER%("Filtrado de Salida:\\Totales\Local\Com.Exterior")
    If OPX% = 1 Then
        Call FINAL("*LIREIVA")
      ElseIf OPX% = 2 Then
        JJ& = 0
        For U& = 1 To FIN&
          XXX$ = REGLEIDO$("REIVACOM", U&)
          If InStr(XXX$, "COM.EXT") < 1 Then
            JJ& = JJ& + 1
            Call GRAREG("REIVACOM", XXX$, JJ&)
          End If
        Next U&
        Call SETULTREG("REIVACOM", JJ&)
        Call FINAL("*LIREIVA")
      ElseIf OPX% = 3 Then
        JJ& = 0
        For U& = 1 To FIN&
          XXX$ = REGLEIDO$("REIVACOM", U&)
          If InStr(XXX$, "COM.EXT") > 0 Then
            JJ& = JJ& + 1
            Call GRAREG("REIVACOM", XXX$, JJ&)
          End If
        Next U&
        Call SETULTREG("REIVACOM", JJ&)
        Call FINAL("*LIREIVA")
    End If
    '
99  Command13.Enabled = True
End Sub

Sub Command16_Click()
    Command16.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Option2.Value = True
    Call GENSUDICOM
    Call FINAL("*SUDICOM")
    '
99  Command16.Enabled = True
End Sub

Private Sub Command17_Click()
    If Option1.Value = True Then
         Call COMUNI("Opción no Disponible por Pantalla.\Emita Listado Impreso\o Exporte a Archivo Texto.")
         Exit Sub
      ElseIf Option2.Value = True Then
         OPXX% = ALTERNA%("Ranking de Ordenes de Compra\Ranking de Facturación")
         If OPXX% = 1 Then
             Call RANKPRO
           ElseIf OPXX% = 2 Then
             Call RANKFACO
         End If
    End If
End Sub

Private Sub Command18_Click()
Ab1$ = VALIMESAN%
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   If DERACCE%("STPREMEC", "Configuración Resumen Mensual de Acreedores") = 2 Then
       FOLIVACO07.Show 1
   End If
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCHIG07/AMAPRO", "Archivos Maestros")
   Command15.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   '
10 OPX% = ALTERNA%("Base Subdiario\Analisis por Cuenta\Analisis por Proveedor\Analisis por Centro de Costos")
   If OPX% < 1 Or OPX% > 4 Then GoTo 99
   '
   'SI ES JYM Y ELEGIO ANALISIS POR CUENTA SALTEA VALIDACION
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "J Y M") > 0 And OPX% = 2 Then
     Call DESHASFECHA(Des%, Has%, PERIO$)
     If PERIO$ = "" Then GoTo 99
     DESDAT = FETEXCOR1$(Des%)
     HASDAT = FETEXCOR1$(Has%)
   Else
     If VALIMESAN% < 1 Then GoTo 99
   End If
   '
   Call GENSUDICOM(OK%)
   If OK% > 0 Then GoTo 99
   '
   Select Case OPX%
   Case 1
     If Option1.Value = True Then
          Call CONECRUN("*COSUCOM", "Consulta por Cuenta Contable")
       ElseIf Option2.Value = True Then
          Call CONECRUN("*SORANAI", "Listado Analitico de Imputacion de Compras")
     End If
   Case 2
     If InStr(EPAC$, "J Y M") > 0 Then
        Call GASPORCUECON
        GoTo 10
     Else
        If Option1.Value = True Then
             Call CONECRUN("*ANAICOM", "Analisis por Cuenta Contable")
          Else
             Call CONECRUN("*SORANAI", "Listado Analitico de Imputacion de Compras")
        End If
     End If
   Case 3
     If Option1.Value = True Then
20        Call CONECRUN("*ANAICOP", "Analisis Contable por Proveedor")
       Else
          Call MENSERR(24, "Solo Disponible por Pantalla")
          GoTo 20
     End If
   Case 4
     If Option1.Value = True Then
          Call MENSERR(24, "Solo Disponible como Listado")
          GoTo 30
       Else
30        Call CONECRUN("*ANCENCOS", "Analisis Contable por Centro de Costo")
     End If
  End Select
   '
99 Command19.Enabled = True
End Sub

Sub Command2_Click()
    Command2.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    MOX% = ALTERNA%("Cuentas Corrientes\Cuentas de Gastos")
    If MOX% < 1 Or MOX% > 2 Then
      GoTo 99
    End If
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM CAJABANC "
    'SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    'SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "Order by FechConta ASC, CodiCom_Lar ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    J1& = 0: J2& = 0: REBLA$ = REGBLAN$("DETPAGC")
    NIVCONXXX% = NIVCONFI%
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      Do While Not .EOF
        '\\\OT-08-1032-FL-12/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
        If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70
        '\\\OT-08-1032-FL-FIN
        ZZ$ = REBLA$
        NUMEFA$ = !CODICOM_LAR
        NCLIE% = !NUPROVE
        FECHAD = Int(CDbl(!fechconta))
        INETO# = !ValAbsComp
        ITOTA# = !ValAbsComp
        '
        Call REPLA(ZZ$, Left$(TRIM$(!TipoMovim), 2), 1, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!fechconta)), 3, 2)
        Call REPLA(ZZ$, !CODICOM_LAR, 5, 18)
        Call REPLA(ZZ$, MKI$(!NUPROVE), 23, 2)
        Call REPLA(ZZ$, MKD$(INETO#), 25, 8)
        Call REPLA(ZZ$, MKD$(IIVA#), 33, 8)
        Call REPLA(ZZ$, MKD$(IVADU#), 41, 8)
        Call REPLA(ZZ$, MKD$(RETEN#), 49, 8)
        Call REPLA(ZZ$, MKD$(IOTRO#), 57, 8)
        Call REPLA(ZZ$, MKD$(ITOTA#), 65, 8)
        '
        If TICUEPRO%((-1) * NCLIE%) = 0 Then
            J1& = J1& + 1
            Call GRAREG("DETPAGC", ZZ$, J1&)
          Else
            J2& = J2& + 1
            Call GRAREG("DETPAGG", ZZ$, J2&)
        End If
70    .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    '\\\OT-07-0098-RG-09-03-07///
    Call SETULTREG("DETPAGC", J1&)
    Call SETULTREG("DETPAGG", J2&)
    '\\\OT-07-0098-RG-FIN///
    '
    Screen.MousePointer = 1
    On MOX% GoTo 203, 204
    '
203 Call HEADERS("DETPAGC", "")
    Call HEADERS("DETPAGC", "Corresponde a: " + TRIM$(MESAN))
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODPAGC")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDPAGC")
    End If
    GoTo 99
    '
204 Call HEADERS("DETPAGG", "")
    Call HEADERS("DETPAGG", "Corresponde a: " + TRIM$(MESAN))
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODPAGG")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDPAGG")
    End If
    GoTo 99
    '
99  Command2.Enabled = True
    Screen.MousePointer = 1
    
End Sub


Private Sub Command20_Click()
   '
   Command20.Enabled = False
   If VALIMESAN% < 1 Then GoTo 99
   '
   Call GENSUDICOM
   '
   JJ& = 0
   TTXXY$ = Chr$(1) + Chr$(2) + MKI$(Hasta%)
   FIN& = ULTREG&("ASIDIAR")
   '
   For IL& = 1 To FIN&
     Call TRACE(20, IL& - IL0&, FIN& - IL0&)
     X$ = REGLEIDO$("ASIDIAR", IL&)
     If Left$(X$, 4) = TTXXY$ Then
       CUENUI% = CVI(Mid$(X$, 5, 2))
       IDEB# = CVD(Mid$(X$, 39, 8))
       If Abs(IDEB#) >= 0.005 Then
         IHAB# = 0
         If IDEB# < 0 Then
           IHAB# = Abs(IDEB#)
           IDEB# = 0
         End If
         CUES$ = CODCUE$(CUENUI%)
         DENOCUEX$ = ECOARCH$("ECUECON", CUES$)
         '
         Z$ = REGBLAN$("ASIDIAFI")
         Call REPLA(Z$, CUES$, 1, 12)
         Call REPLA(Z$, DENOCUEX$, 13, 30)
         Call REPLA(Z$, MKD$(IDEB#), 43, 8)
         Call REPLA(Z$, MKD$(IHAB#), 51, 8)
         '
         JJ& = JJ& + 1
         Call GRAREG("ASIDIAFI", Z$, JJ&)
       End If
     End If
   Next IL&
   Call SETULTREG("ASIDIAFI", JJ&)
   '
   Call FILESORT("ASIDIAFI", "", 1, 1, "ASC")
   Call TRALOCAL("ASIDIAFI", "")
   '
   Screen.MousePointer = 1
   VTB% = VENTABU%("ASIDIAFI/TITUPAN=Asiento Resumen Mensual de Compras - " + CORRESP$)
   '
99 Command20.Enabled = True
   '
End Sub

Sub Command21_Click()
    Command21.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Option2.Value = True
        '\\\OT-05-0264-WAR-09/05/05///
    If RESP1% = 0 Then
    '\\\OT-08-0855-OD-18/09/08///
    RESP1% = ALTERNA%("Selección de Comprobantes a Listar.\Listado Libro de IVA - Compras.\Consulta de Comprobantes Excluidos.")
     If RESP1% < 1 Or RESP1% > 3 Then
    '\\\OT-08-0855-OD-FIN///
       GoTo 99
     End If
    End If
    If RESP1% = 2 Then
        '\\\OT-05-0264-WAR-FIN///
        Call GENIVACOM
        If ULTREG&("LIVACOM") < 1 Then
           GoTo 99
        End If
        '
        FORIPRE$ = TRIM$(CONTROL$("FORIPRE", "LIVACOM"))
        If FORIPRE$ <> "" Then
          If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") > 0 Then
            If COMALTER("Elija Formato de Salida:\\Planilla Plana\Formulario") = 2 Then
              NOCONFIRM% = 1
              Call FORMIVACO
              NOCONFIRM% = 0
              GoTo 99
            End If
          End If
        End If
        'ACA ME FIJO EN EL CONTROL SI USA
        'FORMATO CAI ENTONCES OPTO POR A O B
        If CONTROL$("", "VALICAI") = "SI" Then
            Call FINAL("*LIVCOCAI") 'B
          Else
            Call FINAL("*LIVCOMEN") 'A
        End If
        Option2.Value = True
      ElseIf RESP1% = 1 Then
        Call EXCLUIVA
         '\\\OT-08-0855-OD-18/09/08///
      ElseIf RESP1% = 3 Then 'CONSULTA DE COMPROBANTES EXCLUIDOS
       
        Dim oReporte As cGenrep
        Dim Rst As ADODB.Recordset
        
        Set oReporte = New cGenrep
        ConsultaSql$ = "SELECT FECHCONTA, CODICOM_LAR, NUPROVE, IHABECOMP FROM CAJABANC WITH(NOLOCK)"
        ConsultaSql$ = ConsultaSql$ & " WHERE TIPOMOVIM='FCOM' AND IncSubIva LIKE '%NO%' AND FECHCONTA >='" & DESDAT & "' AND FECHCONTA <='" & HASDAT & "' ORDER BY FECHCONTA"
        Respuesta = oReporte.GeneraReporte("COMPEXCLUIVA", "COMPROBANTES EXCLUIDOS DEL LIBRO IVA", "Fecha/D8; Comprobante/A18; Cuenta(ECO=PROVED1-A48)/I10; Importe(TT=SI)/F14.2", ConsultaSql$)
        If Respuesta Then oReporte.MostrarReporte
        
        Set oReporte = Nothing
    
         'Call FINAL("?COMPEXCLUIVA")
         '\\\OT-08-0855-OD-FIN///
    End If
    '
99  Command21.Enabled = True
End Sub

Sub FORMIVACO()
    '
    FORIPRE$ = TRIM$(CONTROL$("FORIPRE", "LIVACOM"))
    FEX = HOY(HOS$)
    FECHDOC$ = FECHATEX$(FEX)
    NUMEDOC$ = MESAN
    TIPODOC$ = "Libro IVA Compras"
    '
    ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
    If ITPORHOJA% < 1 Then
      ITPORHOJA% = 32767
    End If
    '
    CODPARAM$(1) = "IMP-NETO" ': DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
    CODPARAM$(2) = "IMP-DESC" ': DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
    CODPARAM$(3) = "ITO-IVA"  ': DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
    CODPARAM$(4) = "IMP-RNI"  ': DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
    CODPARAM$(5) = "IMPNEDES" ': DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
    CODPARAM$(6) = "REF-MAT1" ': DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
    CODPARAM$(7) = "REF-MAT2" ': DOCPARAM$(7) = TRIM$(Mid$(DTP$, 159, 32))
    CODPARAM$(8) = "MEDIDA-3" ': DOCPARAM$(7) = TRIM$(Mid$(DTP$, 159, 32))
    CODPARAM$(9) = "MEDIDA-2" ': DOCPARAM$(7) = TRIM$(Mid$(DTP$, 159, 32))
    CAPARDOC% = 9
    '
    CODTABU1$(1) = "FECH-REC"
    CODTABU1$(2) = "TALLE"
    CODTABU1$(3) = "NRO-LOTE"
    CODTABU1$(4) = "RASO-CLI"
    CODTABU1$(5) = "CUIT-CLI"
    CODTABU1$(6) = "PIVA-CLI"
    CODTABU1$(7) = "NET-ITEM"
    CODTABU1$(8) = "BRU-ITEM"
    CODTABU1$(9) = "IMP-IVA"
    CODTABU1$(10) = "IMP-IVA1"
    CODTABU1$(11) = "IMP-TOTA"
    CODTABU1$(12) = "PERC-IVA"
    CODTABU1$(13) = "REF-MAT3"
    CODTABU1$(14) = "PERC-BRU"
    CODTABU1$(15) = "MEDIDA-1"
    '\\\OT-07-0096-RG-09-03-07///
    CODTABU1$(16) = "ESP-MAT"
    CODTABU1$(17) = "F-PLANO"
    '\\\OT-07-0096-RG-FIN///
    CACOLTAB1% = 17
    '
    DESTIDOCNUE = 0
    CAITAB1% = 0
    '
    SUMA1# = 0: SUMA2# = 0: SUMA3# = 0
    SUMA4# = 0: SUMA5# = 0: SUMA6# = 0: SUMA7# = 0: SUMA8# = 0: SUMA9# = 0
    URE& = ULTREG&("LIVACOM")
    For YX& = 1 To URE&
       '
       YY$ = REGLEIDO$("LIVACOM", YX&)
       FEMI = CVI(Mid$(YY$, 3, 2))
       NDOC$ = Mid$(YY$, 11, 18)
       Tipo$ = Left(NDOC$, 2) + Mid$(NDOC$, 4, 1)
       '\\\OT-07-0096-RG-09-03-07///
       FREAL = CVI(Mid$(YY$, 5, 2))
       '\\\OT-07-0096-RG-FIN///
       NUME$ = Mid$(NDOC$, 6, 13)
       RACLI$ = Mid$(YY$, 33, 30)
       CUITI$ = TRIM$(Mid$(YY$, 63, 16))
       PIVAI$ = TRIM$(Mid$(YY$, 79, 12))
       TOFA0# = CVD(Mid$(YY$, 91, 8))
       '
       NEGRA0# = 0: NEXEN0# = 0: NIVA21# = 0: NIVA105# = 0: RETIVA# = 0: RETGAN# = 0: RETIBR# = 0
       '
25     NETO1# = CVD(Mid$(YY$, 137, 8))
       IVA1# = CVD(Mid$(YY$, 145, 8))
       'TASA1 = CVD(Mid$(YY$, 153, 8)) SE CAMBIO ESTA LECTURA POR LA DE ABAJO
       TASA1 = XVALO(Mid$(YY$, 129, 8))
       If Abs(IVA1#) < 0.005 Then
            NEXEN0# = NEXEN0# + NETO1#
          Else
            NEGRA0# = NEGRA0# + NETO1#
            If TASA1 >= 10 And TASA1 <= 11 Then
                 NIVA105# = NIVA105# + IVA1#
               Else
                 NIVA21# = NIVA21# + IVA1#
            End If
       End If
       RETIVA# = RETIVA# + CVD(Mid$(YY$, 161, 8))
       RETGAN# = RETGAN# + CVD(Mid$(YY$, 169, 8))
       RETIBR# = RETIBR# + CVD(Mid$(YY$, 193, 8))
       '
       '\\\OT-07-0096-RG-09-03-07///
       If TRIM$(Mid$(YY$, 11, 18)) <> "" Then
         JURI$ = TRIM$(Mid$(YY$, 205, 24))
       End If
       '\\\OT-07-0096-RG-FIN///
       '
       If YX& < URE& Then
         YZ$ = REGLEIDO$("LIVACOM", YX& + 1)
         If TRIM$(Mid$(YZ$, 11, 18)) = "" Then
           YX& = YX& + 1
           YY$ = YZ$
           GoTo 25
         End If
       End If
       '
       CAITAB1% = CAITAB1% + 1
       TETABU1$(1, CAITAB1%) = FECHATEX$(FEMI)
       TETABU1$(2, CAITAB1%) = Tipo$
       TETABU1$(3, CAITAB1%) = NUME$
       TETABU1$(4, CAITAB1%) = RACLI$
       TETABU1$(5, CAITAB1%) = CUITI$
       TETABU1$(6, CAITAB1%) = PIVAI$
       TETABU1$(7, CAITAB1%) = FORMATNUM$(NEGRA0#, "F15.2")
          SUMA1# = SUMA1# + ROUND#(NEGRA0#)
       TETABU1$(8, CAITAB1%) = FORMATNUM$(NEXEN0#, "F15.2")
          SUMA2# = SUMA2# + ROUND#(NEXEN0#)
       TETABU1$(9, CAITAB1%) = FORMATNUM$(NIVA21#, "F15.2")
          SUMA3# = SUMA3# + ROUND#(NIVA21#)
       TETABU1$(10, CAITAB1%) = FORMATNUM$(NIVA105#, "F15.2")
          SUMA4# = SUMA4# + ROUND#(NIVA105#)
       TETABU1$(11, CAITAB1%) = FORMATNUM$(TOFA0#, "F15.2")
          SUMA5# = SUMA5# + ROUND#(TOFA0#)
       TETABU1$(12, CAITAB1%) = FORMATNUM$(RETIVA#, "F15.2")
          SUMA6# = SUMA6# + ROUND#(RETIVA#)
       TETABU1$(13, CAITAB1%) = FORMATNUM$(RETGAN#, "F15.2")
          SUMA7# = SUMA7# + ROUND#(RETGAN#)
       TETABU1$(14, CAITAB1%) = FORMATNUM$(RETIBR#, "F15.2")
          SUMA8# = SUMA8# + ROUND#(RETIBR#)
       TETABU1$(15, CAITAB1%) = FORMATNUM$(REINTE#, "F15.2")
          SUMA9# = SUMA9# + ROUND#(REINTE#)
       '\\\OT-07-0096-RG-09-03-07///
       TETABU1$(16, CAITAB1%) = JURI$
       TETABU1$(17, CAITAB1%) = FECHATEX$(FREAL)
       '\\\OT-07-0096-RG-FIN///
          '
       If CAITAB1% >= ITPORHOJA% Then
         If YX& < URE& Then
            DOCPARAM$(1) = TRIM$(FORMATNUM$(SUMA1#, "F15.2"))
            DOCPARAM$(2) = TRIM$(FORMATNUM$(SUMA2#, "F15.2"))
            DOCPARAM$(3) = TRIM$(FORMATNUM$(SUMA3#, "F15.2"))
            DOCPARAM$(4) = TRIM$(FORMATNUM$(SUMA4#, "F15.2"))
            DOCPARAM$(5) = TRIM$(FORMATNUM$(SUMA5#, "F15.2"))
            DOCPARAM$(6) = TRIM$(FORMATNUM$(SUMA6#, "F15.2"))
            DOCPARAM$(7) = TRIM$(FORMATNUM$(SUMA7#, "F15.2"))
            DOCPARAM$(8) = TRIM$(FORMATNUM$(SUMA8#, "F15.2"))
            DOCPARAM$(9) = TRIM$(FORMATNUM$(SUMA9#, "F15.2"))
            Call PRINTDOC("@" + FORIPRE$)
            CAITAB1% = 0
         End If
       End If
    Next YX&
    '
    DOCPARAM$(1) = TRIM$(FORMATNUM$(SUMA1#, "F15.2"))
    DOCPARAM$(2) = TRIM$(FORMATNUM$(SUMA2#, "F15.2"))
    DOCPARAM$(3) = TRIM$(FORMATNUM$(SUMA3#, "F15.2"))
    DOCPARAM$(4) = TRIM$(FORMATNUM$(SUMA4#, "F15.2"))
    DOCPARAM$(5) = TRIM$(FORMATNUM$(SUMA5#, "F15.2"))
    DOCPARAM$(6) = TRIM$(FORMATNUM$(SUMA6#, "F15.2"))
    DOCPARAM$(7) = TRIM$(FORMATNUM$(SUMA7#, "F15.2"))
    DOCPARAM$(8) = TRIM$(FORMATNUM$(SUMA8#, "F15.2"))
    DOCPARAM$(9) = TRIM$(FORMATNUM$(SUMA9#, "F15.2"))
    Call PRINTDOC("@" + FORIPRE$)
    '
End Sub
'

Private Sub Command22_Click()
    '
    If Option1.Value = True Then
         Call COMUNI("Opción no Disponible por Pantalla.\Emita Listado Impreso\o Exporte a Archivo Texto.")
         Exit Sub
      ElseIf Option2.Value = True Then
         Call RANKART
    End If
    '
End Sub

Private Sub CITICOMPVIE() 'C.I.T.I.COMPRAS
  '\\\OT-06-0271-WAR-06/07/06///
     '
     Screen.MousePointer = 11
     If VALIMESAN% < 1 Then GoTo 99
     '
5    Call GENIVACOM
     '
     If ULTREG&("LIVACOM") < 1 Then GoTo 99
     '
10   Call MESANO(MESAN, DESX1%, HASX1%)
     VDEF$ = MKI$(DESX1%) + MKI$(HASX1%)
     OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
     If OBX$ = "" Then GoTo 99
     DESX2% = CVI(Left$(OBX$, 2))
     HASX2% = CVI(Mid$(OBX$, 3, 2))
     If DESX2% > HASX2% Or DESX2% < DESX1% Or HASX2% > HASX1% Then
       Call MENSERR(24, "Rango de Fechas no Procesable")
       GoTo 10
     End If
     '
     VD1$ = ""
'15   CUIEM1$ = InputBox$("Ingrese Nº de C.U.I.T.", "C.U.I.T. del Informado", VD1$)
'     If TRIM$(CUIEM1$) = "" Then
'       Call MENSERR(24, "C.U.I.T. no Válido")
'       GoTo 5
'     End If
'     '
'     VD1$ = CUIEM1$
'     If VALICUIT%(CUIEM1$) <= 0 Then
'      Call MENSERR(24, "C.U.I.T. no Válido")
'      GoTo 15
'     End If
     '
     'CUIEM2$ = REPLACAR$(CUIEM1$, "-", "") 'CUIT DEL INFORMADO
     'CUIEM3$ = TRIM$(CUIEM2$)
     '
     L1% = InStr(1, EMPREAC$, "-")
     DENO1$ = Mid(EMPREAC$, 1, L1% - 1) 'APELLIDO Y NOMBRE O DENOMINACION
     '
     ASALID$ = "C:\FLEXOFT\TRICITVA.TXT"
     NX% = FILEOPEN%(ASALID$, 2, 128)
     CAREGI% = 0
     URE& = ULTREG&("LIVACOM")
     For U& = 1 To URE&
       XX$ = REGLEIDO$("LIVACOM", U&)
         FEX = CVI(Mid$(XX$, 3, 2)) 'FECHA CONTABLE
         If FEX < DESX2% Or FEX > HASX2% Then GoTo 29
       If CVD(Mid$(XX$, 145, 8)) > 0 Then ' SI IMPORTE I.V.A.
         TTXX$ = Space$(130)
         '
         'C.U.I.T. DEL PROVEEDOR
         CUIPRO1$ = Mid$(XX$, 63, 16)
         CUIPRO2$ = REPLACAR$(CUIPRO1$, "-", "") 'CUIT DEL VENDEDOR
         CUIPRO3$ = TRIM$(CUIPRO2$)
         'DENOMINACION DEL PROVEEDOR
         DENO2$ = Mid$(XX$, 33, 30)
         CUIEM3$ = CUIPRO3$
         DENO1$ = DENO2$
         '
         'CUIT DEL INFORMADO
         Call REPLA(TTXX$, CUIEM3$, 35, 11)
         'APELLIDO Y NOMBRE O DENOMINACION
         Call REPLA(TTXX$, DENO1$, 46, 25)
         'FECHA DE LA PERCEPCION
         FEXU$ = FECHATEX$(FEX)
         FEXU$ = Left$(FEXU$, 6) + "20" + Mid$(FEXU$, 7, 2)
         FEXU1$ = REPLACAR$(FEXU$, "/", "")
         FEXU2$ = TRIM$(FEXU1$)
         Call REPLA(TTXX$, FEXU2$, 27, 8)
         'PREFIJO DE FACTURA
         X11$ = Mid$(XX$, 11, 18)
         PREFI1$ = Mid$(X11$, 1, 2)
         VARIA1$ = Mid$(X11$, 4, 1)
         TIPO1$ = TICITICOM$(PREFI1$, VARIA1$)
         Call REPLA(TTXX$, TIPO1$, 1, 2)
         'NUMERO DE FACTURA
'         NUPREFI$ = Mid$(XX$, 16, 4)
'         X2 = CVS(Mid$(XX$, 7, 4))
'         NUMEFAC1$ = TRIM$(Str$(X2))
         NUPREFI$ = Mid$(X11$, 6, 4)
         NUMEFAC1$ = Mid$(X11$, 11)
         NUMEFAC1$ = REPLACAR$(NUMEFAC1$, ".", "") 'EXISTEN NUMEROS DE FC CON PUNTOS EN EL MEDIO
         While Len(NUMEFAC1$) < 8: NUMEFAC1$ = "0" + NUMEFAC1$: Wend
         NUMEFAC2$ = NUPREFI$ + "000000000000" + NUMEFAC1$
         If PREFI1$ = "CO" Then 'No esta Formateado
           XPRE$ = Mid$(X11$, 6, 4)
           GUI1% = InStr(7, X11$, "-")
           XNUM$ = Mid$(X11$, GUI1% + 1)
           XNUM$ = REPLACAR$(XNUM$, "-", "")
           XNUM$ = REPLACAR$(XNUM$, ".", "") 'EXISTEN NUMEROS DE FC CON PUNTOS EN EL MEDIO
           While Len(XNUM$) < 8: XNUM$ = "0" + XNUM$: Wend
           NUMEFAC2$ = XPRE$ + "000000000000" + XNUM$
         End If
24       Call REPLA(TTXX$, NUMEFAC2$, 3, 24)
         '
'         'C.U.I.T. DEL VENDEDOR
'         CUIPRO1$ = Mid$(XX$, 63, 16)
'         CUIPRO2$ = REPLACAR$(CUIPRO1$, "-", "") 'CUIT DEL VENDEDOR
'         CUIPRO3$ = TRIM$(CUIPRO2$)
'         'DENOMINACION DEL VENDEDOR
'         DENO2$ = Mid$(XX$, 33, 30)
         '
         If TIPO1$ = "63" Then
          Call REPLA(TTXX$, CUIPRO3$, 83, 11)
          Call REPLA(TTXX$, DENO2$, 94, 25)
          COMIVA1$ = String$(12, "0")
          Call REPLA(TTXX$, COMIVA1$, 119, 12)
         End If
         'SUMO TODAS LOS REGISTROS DEL COMPROBANTE CON IVA
         TOTIVA# = CVD(Mid$(XX$, 145, 8))
         KI% = 0
25       KI% = KI% + 1
         If U& + KI% <= URE& Then
           YY$ = REGLEIDO$("LIVACOM", U& + KI%)
           If TRIM$(Mid$(YY$, 1, 99)) = "" Then
             TOTIVA# = TOTIVA# + CVD(Mid$(YY$, 145, 8))
             GoTo 25
           End If
         End If
         IMPOL# = TOTIVA#
         IMPOX1$ = FORMATNUM$(IMPOL#, "F12.2")
         IMPOX2$ = REPLACAR$(IMPOX1$, ".", "")
         IMPOX3$ = TRIM$(IMPOX2$)
         While Len(IMPOX3$) < 12: IMPOX3$ = "0" + IMPOX3$: Wend
         Call REPLA(TTXX$, IMPOX3$, 71, 12)
         '
         Print #NX%, TTXX$
         CAREGI% = CAREGI% + 1
       End If
29   Next U&
     Close #NX%
     ASALID1$ = REPLACAR$(ASALID$, "\", "/")
     Call COMUNI("Archivo de Transferencia '" + ASALID1$ + "'\Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.")
99   Screen.MousePointer = 1
     '
End Sub

Private Sub CITICOMP() 'C.I.T.I.COMPRAS
  '\\\OT-06-0271-WAR-06/07/06///
     '
     Dim TIDOCU$(2048), NUDOCU$(2048), FEDOCU$(2048), CUIPROX$(2048), DENPROX$(2048), IMPOIVAX#(2048)
     
     Screen.MousePointer = 11
     If VALIMESAN% < 1 Then GoTo 99
     '
     N1% = FILEOPEN(LUDAT$ + "SVD147.EXE", 0, 41)
     Dim X As String * 41
     Dim Z As String * 41
     Get #N1%, 6, X
     CUITT$ = DESENCRI$(RTrim(X))
     CUIPRO1$ = TRIM$(InputBox("Ingrese CUIT", "CUIT de la Empresa", CUITT$))
     Valid% = VALICUIT%(CUIPRO1$)
     If Valid% < 1 Then
        Call COMUNI("C.U.I.T. de Empresa No Válido")
        GoTo 99
     Else
        X = ENCRIPTA$(CUIPRO1$)
        While Len(X) < 41: X = X + " ": Wend
        Put #N1%, 6, X
     End If
     '
     Get #N1%, 2, Z
     RASOCIALL$ = TRIM$(DESENCRI$(RTrim(Z)))
     RASOCIAL$ = TRIM$(InputBox("Ingrese Razon Social", "Razon Social de la Empresa", RASOCIALL$))
     If RASOCIAL$ = "" Then
        Call COMUNI("Falta Razon Social")
        GoTo 99
     Else
        Z = ENCRIPTA$(RASOCIAL$)
        While Len(Z) < 41: Z = Z + " ": Wend
        Put #N1%, 2, Z
     End If

     Close #N1%
     '
     CUIPRO2$ = REPLACAR$(CUIPRO1$, "-", "") 'CUIT DEL VENDEDOR
     Cuit$ = TRIM$(CUIPRO2$)
     
     '
5    Call GENIVACOM
     '
     If ULTREG&("LIVACOM") < 1 Then GoTo 99
     '
10   Call MESANO(MESAN, DESX1%, HASX1%)
     VDEF$ = MKI$(DESX1%) + MKI$(HASX1%)
     OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
     If OBX$ = "" Then GoTo 99
     'OBX$ = VDEF$
     DESX2% = CVI(Left$(OBX$, 2))
     HASX2% = CVI(Mid$(OBX$, 3, 2))
     If DESX2% > HASX2% Or DESX2% < DESX1% Or HASX2% > HASX1% Then
       Call MENSERR(24, "Rango de Fechas no Procesable")
       GoTo 10
     End If
     '
     VD1$ = ""
     L1% = InStr(1, EMPREAC$, "-")
     DENO1$ = Mid(EMPREAC$, 1, L1% - 1) 'APELLIDO Y NOMBRE O DENOMINACION
     '
     ASALID$ = "C:\FLEXOFT\TRICITVA.TXT"
     NX1% = FILEOPEN%(ASALID$, 2, 128)
     CAREGI% = 0
     ASALID_NC$ = "C:\FLEXOFT\TRICITVA_NC.TXT"
     NX2% = FILEOPEN%(ASALID_NC$, 2, 128)
     ULTIREGI% = 0
     '
     URE& = ULTREG&("LIVACOM")
     For U& = 1 To URE&
       XX$ = REGLEIDO$("LIVACOM", U&)
       FEX = CVI(Mid$(XX$, 3, 2)) 'FECHA CONTABLE
       AA = CVD(Mid$(XX$, 145, 8))
       'If FEX < DESX2% Or FEX > HASX2% Then GoTo 29
       If FEX = 0 Then
          Call REPLA(XX$, ZZX$, 1, 78)
          FEX = CVI(Mid$(XX$, 3, 2)) 'FECHA CONTABLE
       End If
       If FEX < DESX2% Or FEX > HASX2% Then GoTo 29
       If CVD(Mid$(XX$, 145, 8)) > 0 Then ' SI IMPORTE I.V.A.
         'C.U.I.T. DEL PROVEEDOR
         CUIPRO1$ = Mid$(XX$, 63, 16)
         CUIPRO2$ = REPLACAR$(CUIPRO1$, "-", "") 'CUIT DEL VENDEDOR
         CUIPRO3$ = TRIM$(CUIPRO2$)
         'DENOMINACION DEL PROVEEDOR
         DENO2$ = TRIM$(Mid$(XX$, 33, 30))
         For i% = 1 To Len(DENO2$)
           If Asc(Mid(DENO2$, i%, 1)) > 127 Then
            DENO2$ = REPLACE$(DENO2$, " ", i%, 1)
           End If
         Next i%
         DENO3$ = DENO2$
         'FECHA DE LA PERCEPCION
         FEXU$ = FECHATEX$(FEX)
         FEXU$ = Left$(FEXU$, 6) + "20" + Mid$(FEXU$, 7, 2)
         FEXU1$ = REPLACAR$(FEXU$, "/", "")
         FEXU3$ = TRIM$(FEXU1$)
         'PREFIJO DE FACTURA
         X11$ = Mid$(XX$, 11, 18)
         PREFI1$ = Mid$(X11$, 1, 2)
         VARIA1$ = Mid$(X11$, 4, 1)
         TIPO3$ = TICITICOM$(PREFI1$, VARIA1$)
         'NUMERO DE FACTURA
         NUPREFI$ = Mid$(X11$, 6, 4)
         NUMEFAC1$ = Mid$(X11$, 11)
         NUMEFAC1$ = REPLACAR$(NUMEFAC1$, ".", "") 'EXISTEN NUMEROS DE FC CON PUNTOS EN EL MEDIO
         NUMEFAC1$ = REPLACAR$(NUMEFAC1$, "-", "") 'EXISTEN NUMEROS DE FC CON GUION EN EL MEDIO
         While Len(NUMEFAC1$) < 8: NUMEFAC1$ = "0" + NUMEFAC1$: Wend
         NUMEFAC3$ = NUPREFI$ + "000000000000" + NUMEFAC1$
         'SUMO TODAS LOS REGISTROS DEL COMPROBANTE CON IVA
         TOTIVA# = CVD(Mid$(XX$, 145, 8))
         KI% = 0
'25       KI% = KI% + 1
'         If U& + KI% <= URE& Then
'           YY$ = REGLEIDO$("LIVACOM", U& + KI%)
'           If TRIM$(Mid$(YY$, 1, 99)) = "" Then
'             TOTIVA# = TOTIVA# + CVD(Mid$(YY$, 145, 8))
'             GoTo 25
'           End If
'         End If
         IMPOL3# = TOTIVA#
         '
         If IMPOL3# < 500 Then
           For CARE1% = 1 To CAREGI%
             If TIDOCU$(CARE1%) = "00" Then
               If NUDOCU$(CARE1%) = String$(24, "0") Then
                 If CUIPROX$(CARE1%) = CUIPRO3$ Then
                   'If DENPROX$(CARE1%) = DENO3$ Then
                     IMPOIVAX#(CARE1%) = IMPOIVAX#(CARE1%) + IMPOL3#
                     dd = dd + IMPOL3#
                     GoTo 29
                   'End If
                 End If
               End If
             End If
           Next CARE1%
           CAREGI% = CAREGI% + 1
           TIDOCU$(CAREGI%) = "00"
           NUDOCU$(CAREGI%) = String$(24, "0")
           FEDOCU$(CAREGI%) = "00000000"
           CUIPROX$(CAREGI%) = CUIPRO3$
           DENPROX$(CAREGI%) = DENO3$
           IMPOIVAX#(CAREGI%) = IMPOL3#
           dd = dd + IMPOL3#
           GoTo 29
         End If
         '
         'ESTO VERIFICA SI EL DOCUMENTO TIENE MAS DE UN REGISTRO. EN CUYO
         'CASO LO ACUMULA EN UNO SOLO. POR EJEMPLO, CUANDO HAY IVA COMUN Y ADUANA
         For CARE1% = 1 To CAREGI%
             If TIDOCU$(CARE1%) = TIPO3$ Then
               If NUDOCU$(CARE1%) = NUMEFAC3$ Then
                 If CUIPROX$(CARE1%) = CUIPRO3$ Then
                     IMPOIVAX#(CARE1%) = IMPOIVAX#(CARE1%) + IMPOL3#
                     dd = dd + IMPOL3#
                     GoTo 29
                 End If
               End If
             End If
         Next CARE1%
         '
         CAREGI% = CAREGI% + 1
         TIDOCU$(CAREGI%) = TIPO3$
         NUDOCU$(CAREGI%) = NUMEFAC3$
         FEDOCU$(CAREGI%) = FEXU3$
         CUIPROX$(CAREGI%) = CUIPRO3$
         DENPROX$(CAREGI%) = DENO3$
         IMPOIVAX#(CAREGI%) = IMPOL3#
       End If
29     ZZX$ = XX$
     Next U&
     '
     ULTIREGI% = CAREGI%
     '
     'NOTAS DE CREDITO DE VENTA
     DESDAT = FETEXCOR1$(DESX2%)
     HASDAT = FETEXCOR1$(HASX2%)
     SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE FECHCONTA >='" & DESDAT & "' AND FECHCONTA <='" & HASDAT & "' "
     SQLX$ = SQLX$ + " AND TIPOMOVIM='IVA' AND OPCIMOVIM='DEFI' AND IDEBECOMP>0.05"
     Set Rst = READSET(SQLX$)
     With Rst
          Do While Not .EOF
             FEX = CVINT(!fechconta)
             NCLI% = XVALO(!nuclien)
             'C.U.I.T. DE LA EMPRESA
             CUIPRO3$ = Cuit$
             'DENOMINACION DEL PROVEEDOR
             DENO3$ = RASOCIAL$
             'FECHA DE LA PERCEPCION
             FEXU$ = FECHATEX$(FEX)
             FEXU$ = Left$(FEXU$, 6) + "20" + Mid$(FEXU$, 7, 2)
             FEXU1$ = REPLACAR$(FEXU$, "/", "")
             FEXU3$ = TRIM$(FEXU1$)
             'PREFIJO DE FACTURA
             X11$ = SAFETEXT$(!CODICOM_LAR)
             PREFI1$ = Mid$(X11$, 1, 2)
             VARIA1$ = Mid$(X11$, 18, 1)
             TIPO3$ = TICITICOM$(PREFI1$, VARIA1$)
             'NUMERO DE FACTURA
             NUPREFI$ = Mid$(X11$, 4, 4)
             NUMEFAC1$ = Mid$(X11$, 9, 8)
             NUMEFAC1$ = REPLACAR$(NUMEFAC1$, ".", "") 'EXISTEN NUMEROS DE FC CON PUNTOS EN EL MEDIO
             NUMEFAC1$ = REPLACAR$(NUMEFAC1$, "-", "") 'EXISTEN NUMEROS DE FC CON GUION EN EL MEDIO
             While Len(NUMEFAC1$) < 8: NUMEFAC1$ = "0" + NUMEFAC1$: Wend
             NUMEFAC3$ = NUPREFI$ + "000000000000" + NUMEFAC1$
             TOTIVA# = XVALO(!Idebecomp)
             IMPOL3# = TOTIVA#
             '
             CAREGI% = CAREGI% + 1
             TIDOCU$(CAREGI%) = TIPO3$
             NUDOCU$(CAREGI%) = NUMEFAC3$
             FEDOCU$(CAREGI%) = FEXU3$
             CUIPROX$(CAREGI%) = CUIPRO3$
             DENPROX$(CAREGI%) = DENO3$
             IMPOIVAX#(CAREGI%) = IMPOL3#
             '
33           .MoveNext
          Loop
     End With
     '
     For CARE2% = 1 To ULTIREGI%
       TTXX$ = Space$(130)
       Call REPLA(TTXX$, TIDOCU$(CARE2%), 1, 2)
       Call REPLA(TTXX$, NUDOCU$(CARE2%), 3, 24)
       Call REPLA(TTXX$, FEDOCU$(CARE2%), 27, 8)
       Call REPLA(TTXX$, CUIPROX$(CARE2%), 35, 11)
       Call REPLA(TTXX$, DENPROX$(CARE2%), 46, 25)
         IMPOL# = IMPOIVAX#(CARE2%)
         IMPOX1$ = FORMATNUM$(IMPOL#, "F12.2")
         IMPOX2$ = REPLACAR$(IMPOX1$, ".", "")
         IMPOX3$ = TRIM$(IMPOX2$)
         While Len(IMPOX3$) < 12: IMPOX3$ = "0" + IMPOX3$: Wend
       Call REPLA(TTXX$, IMPOX3$, 71, 12)
       '
       'SE AGREGA PARA COMPLETAR EL NUEVO FORMATO DEL CITI COMPRAS
       Call REPLA(TTXX$, String(11, "0"), 83, 11)    'CUIT DEL VENDEDOR
       Call REPLA(TTXX$, Space(25), 94, 25)          'NOMBRE DEL VENDEDOR
       Call REPLA(TTXX$, String(12, "0"), 119, 12)   'IVA COMISION VENDEDOR
       '
       Print #NX1%, TTXX$
     Next CARE2%
     '
     Close #NX1%
     '
     For CARE2% = ULTIREGI% + 1 To CAREGI%
       TTXX$ = Space$(130)
       Call REPLA(TTXX$, TIDOCU$(CARE2%), 1, 2)
       Call REPLA(TTXX$, NUDOCU$(CARE2%), 3, 24)
       Call REPLA(TTXX$, FEDOCU$(CARE2%), 27, 8)
       Call REPLA(TTXX$, CUIPROX$(CARE2%), 35, 11)
       Call REPLA(TTXX$, DENPROX$(CARE2%), 46, 25)
         IMPOL# = IMPOIVAX#(CARE2%)
         IMPOX1$ = FORMATNUM$(IMPOL#, "F12.2")
         IMPOX2$ = REPLACAR$(IMPOX1$, ".", "")
         IMPOX3$ = TRIM$(IMPOX2$)
         While Len(IMPOX3$) < 12: IMPOX3$ = "0" + IMPOX3$: Wend
       Call REPLA(TTXX$, IMPOX3$, 71, 12)
       '
       'SE AGREGA PARA COMPLETAR EL NUEVO FORMATO DEL CITI COMPRAS
       Call REPLA(TTXX$, String(11, "0"), 83, 11)    'CUIT DEL VENDEDOR
       Call REPLA(TTXX$, Space(25), 94, 25)          'NOMBRE DEL VENDEDOR
       Call REPLA(TTXX$, String(12, "0"), 119, 12)   'IVA COMISION VENDEDOR
       '
       Print #NX2%, TTXX$
     Next CARE2%
     '
     Close #NX2%
     '
     ASALID1$ = REPLACAR$(ASALID$, "\", "/")
     Call COMUNI("Archivo de Transferencia '" + ASALID1$ + "'\Generado con " + TRIM$(Str$(ULTIREGI%)) + " Registros.")
     ASALID2$ = REPLACAR$(ASALID_NC$, "\", "/")
     Call COMUNI("Archivo de Transferencia '" + ASALID2$ + "'\Generado con " + TRIM$(Str$(CAREGI% - ULTIREGI%)) + " Registros.")
     '
99   Screen.MousePointer = 1
     '
End Sub

Function TICITICOM$(PRE1$, VAR1$)
  'Devuelve el Codigo del Tipo de Comprobante
  TICITICOM$ = "01" 'Información Global
  '
  Select Case PRE1$
    Case "FC"
      If VAR1$ = "A" Then TICITICOM$ = "01"
    Case "ND"
      If VAR1$ = "A" Then TICITICOM$ = "02"
    Case "NC"
      If VAR1$ = "A" Then TICITICOM$ = "03"
      If VAR1$ = "B" Then TICITICOM$ = "08"
      If VAR1$ = "M" Then TICITICOM$ = "53"
    Case "TF"
      If VAR1$ = "A" Then TICITICOM$ = "81"
  End Select
  '
End Function
'\\\OT-06-0271-WAR-FIN///

Private Sub Command23_Click()

    PERIODO$ = "JULIO 15"
    NX% = FILEOPEN%(LUCOM$ & "REGINFO_CV_COMPRAS_CBTE_" & PERIODO$ & ".TXT", 1, 266)
    NX1% = FILEOPEN%(LUCOM$ & "REGINFO_CV_COMPRAS_ALICUOTAS_" & PERIODO$ & ".TXT", 1, 62)
    NX2% = FILEOPEN%(LUCOM$ & "REGINFO_CV_COMPRAS_IMPORTACIONES_" & PERIODO$ & ".TXT", 1, 62)

    REGINFOR.Caption = "Régimen Informativo de Compras"
    
' ** COMPROBANTES *********************************************************************************************************

    With REGINFOR.GridComprobantes
            .Cols = 23
            .TextMatrix(0, 0) = "Fila": .ColAlignment(0) = flexAlignCenterTop
            .TextMatrix(0, 1) = "Fecha"
            .TextMatrix(0, 2) = "Tipo Cbte"
            .TextMatrix(0, 3) = "Punto Venta": .ColWidth(3) = .ColWidth(3) + 300
            .TextMatrix(0, 4) = "Numero": .ColWidth(4) = .ColWidth(4) + 500
            .TextMatrix(0, 5) = "Despacho": .ColWidth(5) = .ColWidth(5) + 1500
            .TextMatrix(0, 6) = "Pos. Iva": .ColWidth(6) = .ColWidth(6) + 500
            .TextMatrix(0, 7) = "Cuit Proveedor": .ColWidth(7) = .ColWidth(7) + 300: .ColAlignment(5) = flexAlignLeftTop
            .TextMatrix(0, 8) = "Razón Social": .ColWidth(8) = .ColWidth(8) + 1400
            .TextMatrix(0, 9) = "Importe Total": .ColWidth(9) = .ColWidth(9) + 1500: .ColAlignment(9) = flexAlignRightTop
            .TextMatrix(0, 10) = "No Gravado": .ColWidth(10) = .ColWidth(10) + 1500: .ColAlignment(10) = flexAlignRightTop
            .TextMatrix(0, 11) = "Exento": .ColWidth(11) = .ColWidth(11) + 1500: .ColAlignment(11) = flexAlignRightTop
            .TextMatrix(0, 12) = "Percep. Iva": .ColWidth(12) = .ColWidth(12) + 1500: .ColAlignment(12) = flexAlignRightTop
            .TextMatrix(0, 13) = "Otros Impuestos": .ColWidth(13) = .ColWidth(13) + 1500: .ColAlignment(13) = flexAlignRightTop
            .TextMatrix(0, 14) = "Percep. IIBB": .ColWidth(14) = .ColWidth(14) + 1500: .ColAlignment(14) = flexAlignRightTop
            .TextMatrix(0, 15) = "Impuesto Municipal": .ColWidth(15) = .ColWidth(15) + 1500: .ColAlignment(15) = flexAlignRightTop
            .TextMatrix(0, 16) = "Impuesto Interno": .ColWidth(16) = .ColWidth(16) + 1500: .ColAlignment(16) = flexAlignRightTop
            .TextMatrix(0, 17) = "Moneda"
            .TextMatrix(0, 18) = "Tipo De Cambio": .ColWidth(18) = .ColWidth(18) + 1000: .ColAlignment(18) = flexAlignRightTop
            .TextMatrix(0, 19) = "Alicuotas": .ColAlignment(19) = flexAlignRightTop
            .TextMatrix(0, 20) = "Codigo Operacion": .ColWidth(20) = .ColWidth(20) + 1500: .ColAlignment(20) = flexAlignRightTop
            .TextMatrix(0, 21) = "Credito Fiscal": .ColWidth(21) = .ColWidth(21) + 1500: .ColAlignment(21) = flexAlignRightTop
            .TextMatrix(0, 22) = "Otros": .ColWidth(22) = .ColWidth(22) + 1500: .ColAlignment(22) = flexAlignRightTop
    End With
    
    FILA& = 1
    
    Do While Not EOF(NX%)
        Line Input #NX%, XX$
            fechaCbte$ = Mid$(XX$, 1, 8)
        fechaCbte$ = Mid$(fechaCbte$, 7) & "/" & Mid$(fechaCbte$, 5, 2) & "/" & Mid$(fechaCbte$, 1, 4)
            TipoCbte$ = Mid$(XX$, 9, 3)
        TipoCbte$ = LetraTipoComprobante(TipoCbte$)
        PuntoDeVenta$ = Mid$(XX$, 12, 5)
            NumeroDeCbte$ = Mid$(XX$, 17, 20)
        NumeroDeCbte$ = CStr(XVALO(NumeroDeCbte$))
        Despacho$ = Mid$(XX$, 37, 16)
            IdPosicionIva$ = Mid$(XX$, 53, 2)
        IdPosicionIva$ = PosicionIvaSegunCodigoDocumento$(CInt(IdPosicionIva$))
        CUITPROVEEDOR$ = CStr(XVALO(Mid$(XX$, 55, 20)))
        razonSocial$ = Mid$(XX$, 75, 30)
            ImporteTotal$ = Mid$(XX$, 105, 15)
        ImporteTotal$ = FORMATNUM$(XVALO(ImporteTotal$) / 100, "F12.2")
            ImporteNoGravado$ = Mid$(XX$, 120, 15)
        ImporteNoGravado$ = FORMATNUM$(XVALO(ImporteNoGravado$) / 100, "F12.2")
            ImporteExento$ = Mid$(XX$, 135, 15)
        ImporteExento$ = FORMATNUM$(XVALO(ImporteExento$) / 100, "F12.2")
            PercepcionIva$ = Mid$(XX$, 150, 15)
        PercepcionIva$ = FORMATNUM$(XVALO(PercepcionIva$) / 100, "F12.2")
            OtrosImpuestos$ = Mid$(XX$, 165, 15)
        OtrosImpuestos$ = FORMATNUM$(XVALO(OtrosImpuestos$) / 100, "F12.2")
            IIBB0$ = Mid$(XX$, 180, 15)
        IIBB0$ = FORMATNUM$(XVALO(IIBB0$) / 100, "F12.2")
            ImpuestoMunicipal$ = Mid$(XX$, 195, 15)
        ImpuestoMunicipal$ = FORMATNUM$(XVALO(ImpuestoMunicipal$) / 100, "F12.2")
            ImpuestosInternos$ = Mid$(XX$, 210, 15)
        ImpuestosInternos$ = FORMATNUM$(XVALO(ImpuestosInternos$) / 100, "F12.2")
        CodigoDeMoneda$ = Mid$(XX$, 225, 3)
            TIPOCAMBIO$ = Mid$(XX$, 228, 10)
        TIPOCAMBIO$ = FORMATNUM$(XVALO(TIPOCAMBIO$) / 1000000, "F12.6")
        CantidadAlicuotas$ = Mid$(XX$, 238, 1)
        CodigoOperacion$ = Mid$(XX$, 239, 1)
            CreditoFiscal$ = Mid$(XX$, 240, 15)
        CreditoFiscal$ = FORMATNUM$(XVALO(CreditoFiscal$) / 100, "F12.2")
            OtrosAtributos$ = Mid$(XX$, 255, 15)
        OtrosAtributos$ = FORMATNUM$(XVALO(OtrosAtributos$) / 100, "F12.2")
        
        With REGINFOR.GridComprobantes
            
            .TextMatrix(FILA&, 0) = CStr(FILA&)
            .TextMatrix(FILA&, 1) = fechaCbte$
            .TextMatrix(FILA&, 2) = TipoCbte$
            .TextMatrix(FILA&, 3) = PuntoDeVenta$
            .TextMatrix(FILA&, 4) = NumeroDeCbte$
            .TextMatrix(FILA&, 5) = Despacho$
            .TextMatrix(FILA&, 6) = IdPosicionIva$
            .TextMatrix(FILA&, 7) = CUITPROVEEDOR$
            .TextMatrix(FILA&, 8) = razonSocial$
            .TextMatrix(FILA&, 9) = ImporteTotal$
            .TextMatrix(FILA&, 10) = ImporteNoGravado$
            .TextMatrix(FILA&, 11) = ImporteExento$
            .TextMatrix(FILA&, 12) = PercepcionIva$
            .TextMatrix(FILA&, 13) = OtrosImpuestos$
            .TextMatrix(FILA&, 14) = IIBB0$
            .TextMatrix(FILA&, 15) = ImpuestoMunicipal$
            .TextMatrix(FILA&, 16) = ImpuestosInternos$
            .TextMatrix(FILA&, 17) = CodigoDeMoneda$
            .TextMatrix(FILA&, 18) = TIPOCAMBIO$
            .TextMatrix(FILA&, 19) = CantidadAlicuotas$
            .TextMatrix(FILA&, 20) = CodigoOperacion$
            .TextMatrix(FILA&, 21) = CreditoFiscal$
            .TextMatrix(FILA&, 22) = OtrosAtributos$
    
            .Rows = .Rows + 1
        End With
        
        FILA& = FILA& + 1
    Loop
    
    
' ** ALICUOTAS *********************************************************************************************************

    With REGINFOR.GridAlicuota
            .Cols = 9
            .TextMatrix(0, 0) = "Fila": .ColAlignment(0) = flexAlignCenterTop
            .TextMatrix(0, 1) = "Tipo Cbte"
            .TextMatrix(0, 2) = "Punto Venta": .ColWidth(2) = .ColWidth(2) + 300: .ColAlignment(2) = flexAlignRightTop
            .TextMatrix(0, 3) = "Numero": .ColWidth(3) = .ColWidth(3) + 500
            .TextMatrix(0, 4) = "Pos. Iva": .ColWidth(4) = .ColWidth(4) + 500
            .TextMatrix(0, 5) = "Cuit Proveedor": .ColWidth(5) = .ColWidth(5) + 300: .ColAlignment(5) = flexAlignLeftTop
            .TextMatrix(0, 6) = "Importe Neto": .ColWidth(6) = .ColWidth(6) + 1500: .ColAlignment(6) = flexAlignRightTop
            .TextMatrix(0, 7) = "Alícuota": .ColWidth(7) = .ColWidth(7) + 500: .ColAlignment(7) = flexAlignRightTop
            .TextMatrix(0, 8) = "Impuesto": .ColWidth(8) = .ColWidth(8) + 1500: .ColAlignment(8) = flexAlignRightTop
    End With
    
    FILA& = 1
    
    Do While Not EOF(NX1%)
        Line Input #NX1%, XX$
            TipoCbte$ = Mid$(XX$, 1, 3)
        TipoCbte$ = LetraTipoComprobante(TipoCbte$)
        PuntoDeVenta$ = Mid$(XX$, 4, 5)
            NumeroDeCbte$ = Mid$(XX$, 9, 20)
        NumeroDeCbte$ = CStr(XVALO(NumeroDeCbte$))
            IdPosicionIva = Mid$(XX$, 29, 2)
        IdPosicionIva$ = PosicionIvaSegunCodigoDocumento$(CInt(IdPosicionIva$))
            CUITPROVEEDOR$ = Mid$(XX$, 31, 20)
        CUITPROVEEDOR$ = CStr(XVALO(CUITPROVEEDOR$))
            ImporteNeto$ = Mid$(XX$, 51, 15)
        ImporteNeto$ = FORMATNUM$(XVALO(ImporteNeto$) / 100, "F12.2")
            ALICUOTA$ = TRIM$(Mid$(XX$, 66, 4))
        ALICUOTA$ = AlicuotaSegunCodigo(ALICUOTA$)
            Impuesto$ = Mid$(XX$, 70, 15)
        Impuesto$ = FORMATNUM$(XVALO(Impuesto$) / 100, "F12.2")
        
        With REGINFOR.GridAlicuota
            
            .TextMatrix(FILA&, 0) = CStr(FILA&)
            .TextMatrix(FILA&, 1) = TipoCbte$
            .TextMatrix(FILA&, 2) = PuntoDeVenta$
            .TextMatrix(FILA&, 3) = NumeroDeCbte$
            .TextMatrix(FILA&, 4) = IdPosicionIva
            .TextMatrix(FILA&, 5) = CUITPROVEEDOR$
            .TextMatrix(FILA&, 6) = ImporteNeto$
            .TextMatrix(FILA&, 7) = ALICUOTA
            .TextMatrix(FILA&, 8) = Impuesto$
    
            .Rows = .Rows + 1
        End With
        
        FILA& = FILA& + 1
    Loop
    
    
' ** IMPORTACIONES *********************************************************************************************************

    With REGINFOR.GridImportaciones
            .Cols = 5
            .TextMatrix(0, 0) = "Fila"
            .TextMatrix(0, 1) = "Despacho": .ColWidth(1) = .ColWidth(1) + 1500
            .TextMatrix(0, 2) = "Importe Neto": .ColWidth(2) = .ColWidth(2) + 1500
            .TextMatrix(0, 3) = "Alícuota": .ColWidth(3) = .ColWidth(3) + 1500
            .TextMatrix(0, 4) = "Impuesto": .ColWidth(4) = .ColWidth(4) + 1500
    End With
    
    FILA& = 1
    
    Do While Not EOF(NX2%)
        Line Input #NX2%, XX$
        Despacho$ = Mid$(XX$, 1, 16)
        ImporteNeto$ = Mid$(XX$, 17, 15)
        ALICUOTA = Mid$(XX$, 32, 4)
        Impuesto$ = Mid$(XX$, 36, 15)
        
        With REGINFOR.GridImportaciones
            
            .TextMatrix(FILA&, 0) = CStr(FILA&)
            .TextMatrix(FILA&, 1) = Despacho$
            .TextMatrix(FILA&, 2) = ImporteNeto$
            .TextMatrix(FILA&, 3) = ALICUOTA
            .TextMatrix(FILA&, 4) = Impuesto$
    
            .Rows = .Rows + 1
        End With
        
        FILA& = FILA& + 1
    Loop
    
    Close #NX%
    Close #NX1%
    Close #NX2%
    
    REGINFOR.Show 1
End Sub

Function LetraTipoComprobante(Tipo$) As String
    
    TIPOCOM$ = ""
    Select Case Tipo$
        Case "001": TIPOCOM$ = "FCA"
        Case "002": TIPOCOM$ = "NDA"
        Case "003": TIPOCOM$ = "NCA"
        Case "006": TIPOCOM$ = "FCB"
        Case "007": TIPOCOM$ = "NDB"
        Case "008": TIPOCOM$ = "NCB"
        Case "019": TIPOCOM$ = "FCE"
        Case "020": TIPOCOM$ = "NDE"
        Case "021": TIPOCOM$ = "NCE"
        Case "011": TIPOCOM$ = "FCC"
        Case "012": TIPOCOM$ = "NDC"
        Case "013": TIPOCOM$ = "NCC"
        Case "004": TIPOCOM$ = "RHA"
        Case "009": TIPOCOM$ = "RHB"
        Case "015": TIPOCOM$ = "RHC"
        Case "054": TIPOCOM$ = "RHM"
        'Case "002": TipoCom$ = "CPA"
        'Case "007": TipoCom$ = "CPB"
        'Case "012": TipoCom$ = "CPC"
        Case "052": TIPOCOM$ = "CPM"
        Case "051": TIPOCOM$ = "FCM"
        Case "052": TIPOCOM$ = "NDM"
        Case "053": TIPOCOM$ = "NCM"
        Case "081": TIPOCOM$ = "TFA"
        Case "082": TIPOCOM$ = "TFB"
        Case "111": TIPOCOM$ = "TFC"
        Case "118": TIPOCOM$ = "TFM"
        'Case "001": TipoCom$ = "COA"
        'Case "006": TipoCom$ = "COB"
        'Case "019": TipoCom$ = "COE"
        'Case "011": TipoCom$ = "COC"
        'Case "051": TipoCom$ = "COM"
        'Case "003": TipoCom$ = "DPA"
        'Case "008": TipoCom$ = "DPB"
        'Case "013": TipoCom$ = "DPC"
        'Case "053": TipoCom$ = "DPM"
    End Select
            
    LetraTipoComprobante = TIPOCOM$
    
End Function

Function PosicionIvaSegunCodigoDocumento(Codigo%) As String

    Select Case Codigo%
        Case 96: PosIva$ = "Consumidor Final"
        Case Else: PosIva$ = "Responsable"
    End Select
    
    PosicionIvaSegunCodigoDocumento = PosIva$
    
End Function
Private Sub Command24_Click()
'    '\\\OT-05-0381-WAR-23/06/05///
'    Call COMUNI("No Disponible en la Presente Versión")
'    '\\\OT-05-0381-WAR-FIN///
'
'\\\OT-06-0180-WAR-23/06/05///
   Command24.Enabled = False
   Screen.MousePointer = 11
   Call LIESFACO
99 Screen.MousePointer = 1
   Command24.Enabled = True
'\\\OT-06-0180-WAR-FIN///
End Sub





Private Sub Command29_Click()
   Call HELPONLINE("COGECOM")
End Sub

Sub Command3_Click()
    'jandy sql
    
    Command3.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    
    SQLX$ = "SELECT * FROM CAJABANC "
'    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "AND OpciMovim = 'FC' "
    SQLX$ = SQLX$ + "Order by FechConta ASC "
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
26  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 26
    End If
    On Error GoTo 0
   
    
'    SQLY$ = "SELECT * FROM CAJABANC "
'    SQLY$ = SQLY$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
'    SQLY$ = SQLY$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
'    Set CABUSMEN = CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot)
'    '
    JJ& = 0: REBLA$ = REGBLAN$("DETFACC")
    NIVCONXXX% = NIVCONFI%
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      FIN& = .RecordCount
      KUL& = 0
      On Error GoTo 0
      .MoveFirst
      Do While Not .EOF
        '\\\OT-08-1032-FL-12/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
        If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70
        '\\\OT-08-1032-FL-FIN
        KUL& = KUL& + 1
        Call TRACE(20, KUL&, FIN&)
        ZZ$ = REBLA$
        NUMEFA$ = !CODICOM_LAR
        NCLIE% = !NUPROVE
        'FECHAD = Int(CDbl(!FechConta))
        FEFAXI% = CVINT(!fechconta)
        FECHAD = FECHATEXLAR$(FEFAXI%)
        '
        INETO# = 0
        IIVA# = 0
        IVADU# = 0
        RETEN# = 0
        IOTRO# = 0
        ITOTA# = !ValAbsComp
        '
        Dim CABUSMEN As ADODB.Recordset
        Set CABUSMEN = New ADODB.Recordset
        CABUSMEN.CursorLocation = adUseClient
        SQLX$ = "SELECT * FROM CAJABANC "
        SQLX$ = SQLX$ + "WHERE FechConta = '" & FECHAD & "'"
        SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX$ = SQLX$ + "AND NuProve = " & NCLIE% & " "
        '
25      On Error Resume Next
        CABUSMEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
        End If
        On Error GoTo 0
       
        With CABUSMEN
        If Not (.EOF And .BOF) Then
          .MoveFirst
'          .FindFirst SQLX$
'25        If .NoMatch = False Then
           Do While Not (.EOF)
            TIMOV$ = TRIM$(UCase$(!TipoMovim))
            IPARD# = !ValAbsComp
            If TIMOV$ = "COMPRA" Then
                INETO# = INETO# + IPARD#
              ElseIf TIMOV$ = "IVA" And (UCase$(!VARIMOVIM) = "ADUANA" Or UCase$("!VariMovim") = "IMPORT") Then
                IVADU# = IVADU# + IPARD#
              ElseIf TIMOV$ = "IVA" And UCase$(!VARIMOVIM) = "PER-GAN" Then
                RETEN# = RETEN# + IPARD#
              ElseIf TIMOV$ = "IVA" Then
                IIVA# = IIVA# + IPARD#
            End If
           ' .FindNext SQLX$
           ' GoTo 25
           .MoveNext
           Loop
          End If
        End With
        CABUSMEN.Close
        Set CABUSMEN = Nothing
        '
        IOTRO# = ITOTA# - INETO# - IIVA# - IVADU# - RETEN#
        '
        Call REPLA(ZZ$, Left$(TRIM$(!TipoMovim), 2), 1, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!fechconta)), 3, 2)
        Call REPLA(ZZ$, !CODICOM_LAR, 5, 18)
        Call REPLA(ZZ$, MKI$(!NUPROVE), 23, 2)
        Call REPLA(ZZ$, MKD$(INETO#), 25, 8)
        Call REPLA(ZZ$, MKD$(IIVA#), 33, 8)
        Call REPLA(ZZ$, MKD$(IVADU#), 41, 8)
        Call REPLA(ZZ$, MKD$(RETEN#), 49, 8)
        Call REPLA(ZZ$, MKD$(IOTRO#), 57, 8)
        Call REPLA(ZZ$, MKD$(ITOTA#), 65, 8)
        JJ& = JJ& + 1
        Call GRAREG("DETFACC", ZZ$, JJ&)
        '
70    .MoveNext
      Loop
    End With
    
    Call SETULTREG("DETFACC", JJ&)
    '
    Call HEADERS("DETFACC", "")
    Call HEADERS("DETFACC", "Corresponde a: " + TRIM$(CORRESP$))
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODFACC")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDFACC")
    End If

99
    Command3.Enabled = True
    Screen.MousePointer = 1
End Sub

Sub Command4_Click()
    'jandy sql
    Command4.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    
    SQLX$ = "SELECT * FROM CAJABANC "
'    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "AND (OpciMovim = 'ND' or substring(CodiCom_Cor,1, 2) = 'ND') "
    SQLX$ = SQLX$ + "Order by FechConta ASC "
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
 
 
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
'    SQLY$ = "SELECT * FROM CAJABANC "
'    SQLY$ = SQLY$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
'    SQLY$ = SQLY$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
'    Set CABUSMEN = CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot)
'    '
    JJ& = 0: REBLA$ = REGBLAN$("DETFACC")
    NIVCONXXX% = NIVCONFI%
    
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      Do While Not .EOF
        '\\\OT-08-1032-FL-12/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
        If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70
        '\\\OT-08-1032-FL-FIN
        ZZ$ = REBLA$
        NUMEFA$ = !CODICOM_LAR
        NCLIE% = !NUPROVE
        'FECHAD = Int(CDbl(!FechConta))
        FEFAXI% = CVINT(!fechconta)
        FECHAD = FECHATEXLAR$(FEFAXI%)
        '
        INETO# = 0
        IIVA# = 0
        IVADU# = 0
        RETEN# = 0
        IOTRO# = 0
        ITOTA# = !ValAbsComp
        '
        Dim CABUSMEN As ADODB.Recordset
        Set CABUSMEN = New ADODB.Recordset
        CABUSMEN.CursorLocation = adUseClient
     
        SQLX$ = "SELECT * FROM CAJABANC "
        SQLX$ = SQLX$ + " WHERE FechConta  = '" & FECHAD & "' "
        SQLX$ = SQLX$ + " AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX$ = SQLX$ + " AND NuProve = " & NCLIE% & " "
        '
26      On Error Resume Next
        CABUSMEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
  
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 26
        End If
        On Error GoTo 0
         
        With CABUSMEN
        If (Not .EOF And .BOF) Then
          .MoveFirst
'          .FindFirst SQLX$
'25        If .NoMatch = False Then
           Do While Not (.EOF)
            TIMOV$ = TRIM$(UCase$(!TipoMovim))
            IPARD# = !ValAbsComp
            If TIMOV$ = "COMPRA" Then
                INETO# = INETO# + IPARD#
              ElseIf TIMOV$ = "IVA" And (UCase$(!VARIMOVIM) = "ADUANA" Or UCase$("!VariMovim") = "IMPORT") Then
                IVADU# = IVADU# + IPARD#
              ElseIf TIMOV$ = "IVA" And UCase$(!VARIMOVIM) = "PER-GAN" Then
                RETEN# = RETEN# + IPARD#
              ElseIf TIMOV$ = "IVA" Then
                IIVA# = IIVA# + IPARD#
            End If
            '.FindNext SQLX$
            'GoTo 25
            .MoveNext
            Loop
          End If
        End With
        CABUSMEN.Close
        Set CABUSMEN = Nothing
        '
        IOTRO# = ITOTA# - INETO# - IIVA# - IVADU# - RETEN#
        '
        Call REPLA(ZZ$, Left$(TRIM$(!TipoMovim), 2), 1, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!fechconta)), 3, 2)
        Call REPLA(ZZ$, !CODICOM_LAR, 5, 18)
        Call REPLA(ZZ$, MKI$(!NUPROVE), 23, 2)
        Call REPLA(ZZ$, MKD$(INETO#), 25, 8)
        Call REPLA(ZZ$, MKD$(IIVA#), 33, 8)
        Call REPLA(ZZ$, MKD$(IVADU#), 41, 8)
        Call REPLA(ZZ$, MKD$(RETEN#), 49, 8)
        Call REPLA(ZZ$, MKD$(IOTRO#), 57, 8)
        Call REPLA(ZZ$, MKD$(ITOTA#), 65, 8)
        JJ& = JJ& + 1
        Call GRAREG("DETFACC", ZZ$, JJ&)
        '
70    .MoveNext
      Loop
    End With
        
    Call SETULTREG("DETFACC", JJ&)
    '
    Call HEADERS("DETFACC", "")
    Call HEADERS("DETFACC", "Corresponde a: " + TRIM$(CORRESP$))
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODFACC")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDFACC")
    End If

    CABATEMP.Close
99  Set CABATEMP = Nothing
   
    Command4.Enabled = True
    Screen.MousePointer = 1
End Sub


Sub Command5_Click()
   Command5.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    'JANDY SQL
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    
    SQLX$ = "SELECT * FROM CAJABANC "
    'SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    'SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "AND (OpciMovim = 'NC' or SUBSTRING(CodiCom_Cor ,1,2) = 'NC') "
    SQLX$ = SQLX$ + "Order by FechConta ASC "
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
26  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 26
    End If
    On Error GoTo 0
  
    '
    'SQLY$ = "SELECT * FROM CAJABANC "
    'SQLY$ = SQLY$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    'SQLY$ = SQLY$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    'Set CABUSMEN = CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot)
    '
    JJ& = 0: REBLA$ = REGBLAN$("DETFACC")
    NIVCONXXX% = NIVCONFI%
    
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      Do While Not .EOF
        '\\\OT-08-1032-FL-12/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
        If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70
        '\\\OT-08-1032-FL-FIN
    
        ZZ$ = REBLA$
        NUMEFA$ = !CODICOM_LAR
        NCLIE% = !NUPROVE
       'FECHAD = Int(CDbl(!FechConta))
        FEFAXI% = CVINT(!fechconta)
        FECHAD = FECHATEXLAR$(FEFAXI%)
    
        
        '
        INETO# = 0
        IIVA# = 0
        IVADU# = 0
        RETEN# = 0
        IOTRO# = 0
        ITOTA# = !ValAbsComp
        '
        Dim CABUSMEN As ADODB.Recordset
        Set CABUSMEN = New ADODB.Recordset
        CABUSMEN.CursorLocation = adUseClient
        SQLX$ = "SELECT * FROM CAJABANC "
        SQLX$ = SQLX$ + " WHERE FechConta = '" & FECHAD & "' "
        SQLX$ = SQLX$ + " AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX$ = SQLX$ + " AND NuProve = " & NCLIE% & " "
        '
25      On Error Resume Next
        CABUSMEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
        End If
        On Error GoTo 0
        With CABUSMEN
        If Not (.EOF And .BOF) Then
           .MoveFirst
'          .FindFirst SQLX$
'25        If .NoMatch = False Then
           Do While Not .EOF
            TIMOV$ = TRIM$(UCase$(!TipoMovim))
            IPARD# = !ValAbsComp
            If TIMOV$ = "COMPRA" Then
                INETO# = INETO# + IPARD#
              ElseIf TIMOV$ = "IVA" And (UCase$(!VARIMOVIM) = "ADUANA" Or UCase$("!VariMovim") = "IMPORT") Then
                IVADU# = IVADU# + IPARD#
              ElseIf TIMOV$ = "IVA" And UCase$(!VARIMOVIM) = "PER-GAN" Then
                RETEN# = RETEN# + IPARD#
              ElseIf TIMOV$ = "IVA" Then
                IIVA# = IIVA# + IPARD#
            End If
            '.FindNext SQLX$
            'GoTo 25
            .MoveNext
            Loop
          End If
        End With
        CABUSMEN.Close
        Set CABUSMEN = Nothing
                
        '
        IOTRO# = ITOTA# - INETO# - IIVA# - IVADU# - RETEN#
        '
        Call REPLA(ZZ$, Left$(TRIM$(!TipoMovim), 2), 1, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!fechconta)), 3, 2)
        Call REPLA(ZZ$, !CODICOM_LAR, 5, 18)
        Call REPLA(ZZ$, MKI$(!NUPROVE), 23, 2)
        Call REPLA(ZZ$, MKD$(INETO#), 25, 8)
        Call REPLA(ZZ$, MKD$(IIVA#), 33, 8)
        Call REPLA(ZZ$, MKD$(IVADU#), 41, 8)
        Call REPLA(ZZ$, MKD$(RETEN#), 49, 8)
        Call REPLA(ZZ$, MKD$(IOTRO#), 57, 8)
        Call REPLA(ZZ$, MKD$(ITOTA#), 65, 8)
        JJ& = JJ& + 1
        Call GRAREG("DETFACC", ZZ$, JJ&)
        '
70    .MoveNext
      Loop
    End With
    
    
    Call SETULTREG("DETFACC", JJ&)
    '
    Call HEADERS("DETFACC", "")
    Call HEADERS("DETFACC", "Corresponde a: " + TRIM$(CORRESP$))
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODFACC")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDFACC")
    End If
    CABATEMP.Close
99  Set CABATEMP = Nothing
    Command5.Enabled = True
    Screen.MousePointer = 1
End Sub

Sub Command6_Click()
    '
    Command6.Enabled = False
    If VALIMESAN% > 0 Then
       Call REMECUEN
    End If
    Command6.Enabled = True
    '
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Screen.MousePointer = 11
    'jandy sql
    Call APERBASDAT("CABAN")
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    
    SQLX$ = "SELECT * FROM CAJABANC "
    'SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    'SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    
    SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'AJUP' "
    'jandy esto estaba asi
    'SQLX$ = SQLX$ + "AND OpciMovim = 'FC' "
    'SQLX$ = SQLX$ + "Order by FechConta ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
  
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    
    JJ& = 0: REBLA$ = REGBLAN$("DETOTRC")
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      Do While Not .EOF
        ZZ$ = REBLA$
        '
        IIVA# = 0
        IVADU# = 0
        RETEN# = 0
        IOTRO# = 0
        INETO# = XVALO(!IHABECOMP) - XVALO(!Idebecomp)
        ITOTA# = INETO#
        '
        Call REPLA(ZZ$, Left$(!TipoMovim, 2), 1, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!fechconta)), 3, 2)
        Call REPLA(ZZ$, !CODICOM_LAR, 5, 18)
        Call REPLA(ZZ$, MKI$(!NUPROVE), 23, 2)
        Call REPLA(ZZ$, MKD$(INETO#), 25, 8)
        Call REPLA(ZZ$, MKD$(IIVA#), 33, 8)
        Call REPLA(ZZ$, MKD$(IVADU#), 41, 8)
        Call REPLA(ZZ$, MKD$(RETEN#), 49, 8)
        Call REPLA(ZZ$, MKD$(IOTRO#), 57, 8)
        Call REPLA(ZZ$, MKD$(ITOTA#), 65, 8)
        JJ& = JJ& + 1
        Call GRAREG("DETOTRC", ZZ$, JJ&)
        '
      .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    Call SETULTREG("DETOTRC", JJ&)
    '
    Call HEADERS("DETOTRC", "")
    Call HEADERS("DETOTRC", "Corresponde a: " + TRIM$(CORRESP$))
    Screen.MousePointer = 1
    '
    If Option1.Value = True Then
         Call FINAL("*CODOTRC")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDOTRC")
    End If
    '
99  Command7.Enabled = True
    Screen.MousePointer = 1
End Sub

Private Sub Command8_Click()
    Call ACCDIR("COGECOM")
End Sub

Private Sub Command9_Click()
    Screen.MousePointer = 11
    C1$ = TRIM$(RESUMENCO.MESAN.TEXT)
    C2$ = TRIM$(RESUMENCO.Text4(0).TEXT)
    C3$ = TRIM$(RESUMENCO.Text4(1).TEXT)
    Call HEADERS("DETFOFI", "")
    Call HEADERS("DETFOFI", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODFOFI")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDFOFI")
    End If
End Sub

Private Sub DespachosCourierTransf_Click()

            If VALIMESAN% < 1 Then Exit Sub
            
            ARCHIX$ = SELECCION_Archivo$("TXT", CommonDialog1)
            If TRIM$(ARCHIX$) = "" Then Exit Sub
            
            Screen.MousePointer = 11
            NARCHI2% = FILEOPEN%(ARCHIX$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
            GENERO_ARCHIVO% = 0: I1& = 0
            '
            SQLX$ = "SELECT SUM(VALABSCOMP) AS TOTALIVA, TIPOMOVIM, DESPAIMPO FROM CAJABANC WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE FECHCONTA >= '" & DESDAT & "'"
            SQLX$ = SQLX$ + " AND FECHCONTA <= '" & HASDAT$ & "'"
            SQLX$ = SQLX$ + " AND DESPAIMPO <> '' AND NUPROVE > 0 AND TIPOMOVIM='IVA' AND VARIMOVIM <> 'Per-IVA'"
            SQLX$ = SQLX$ + " GROUP BY TIPOMOVIM, DESPAIMPO"
            Set RS = READSET(SQLX$)
            With RS
              FILAS& = 0: j& = 0
              Do While Not .EOF
                    TX$ = Space(56)
                    
                    TipoDeclaracion$ = "0"
                    DespachoImportacion$ = SAFETEXT$(!DespaImpo)
                    CreditoFiscal# = XVALO(!TOTALIVA)
                    
                    Call REPLA(TX$, TipoDeclaracion$, 1, 1)
                    Call REPLA(TX$, DespachoImportacion$, 2, 26)
                        IvaCredito$ = TRIM$(FORMATNUM$(CreditoFiscal#, "F15.2"))
                        Do While Len(IvaCredito$) < 15: IvaCredito$ = "0" & IvaCredito$: Loop
                    Call REPLA(TX$, IvaCredito$, 27, 15)
                    Call REPLA(TX$, IvaCredito$, 42, 15)
                    
                    FILAS& = FILAS& + 1
                    Print #NARCHI2%, TX$
                 
                    .MoveNext
              Loop
            End With
            Close #NARCHI2%
            On Error Resume Next
            RS.Close
            Set RS = Nothing
            On Error GoTo 0
            Screen.MousePointer = 1
            '
            If FILAS& > 0 Then
                 Call COMUNI("Archivo Generado con Exito.\ \Se Han Encontrado '" & CStr(FILAS&) & "' Importaciones.")
                 GoTo 99
            Else
                 Call COMUNI("No Se Han Encontrado Importaciones en el Período Elegido")
                 GoTo 99
            End If

99          Screen.MousePointer = 1

End Sub


Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture9)
    '
    If InStr(1, EMPREAC$, "FERVI") > 0 Then
        Screen.MousePointer = 11
        CONDI$ = "nuprove=955 and codicom_lar='AS-A-0000-00000000'"
        If CantRegistros("cajabanc", CONDI$) > 0 Then
            FETX$ = "13/09/12"
            NRECIB$ = "AS-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
            ORMA% = (-1)
            Dim RECITEMP As ADODB.Recordset
            SQLX$ = "SELECT max(CodiCom_Cor) as Maximo FROM CAJABANC "
            SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUP' "
            SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "%' "
            Set RECITEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
            If Not IsNull(RECITEMP!maximo) Then
                ORMA% = XVALO(Mid$(RECITEMP!maximo, 11, 2))
            End If
            RECITEMP.Close
            Set RECITEMP = Nothing
            ORMX$ = TRIM$(Str$(ORMA% + 1))
            While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
            NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
            '
            Call ACTUREGISTRO("cajabanc", "codicom_cor", CONDI$, NRECIB$, REGAF&)
            NRECIB1$ = "AS-X-0000-00" & Mid$(NRECIB$, 4, 6)
            Call ACTUREGISTRO("cajabanc", "codicom_lar", CONDI$, NRECIB1$, REGAF&)
        End If
        
        CONDI$ = "nuprove=1090 and codicom_lar='AS-A-0000-00000000'"
        If CantRegistros("cajabanc", CONDI$) > 0 Then
            FETX$ = "25/10/12"
            NRECIB$ = "AS-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
            ORMA% = (-1)
            Dim RECITEMP1 As ADODB.Recordset
            SQLX$ = "SELECT max(CodiCom_Cor) as Maximo FROM CAJABANC "
            SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUP' "
            SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "%' "
            Set RECITEMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
            If Not IsNull(RECITEMP1!maximo) Then
                ORMA% = XVALO(Mid$(RECITEMP!maximo, 11, 2))
            End If
            RECITEMP1.Close
            Set RECITEMP1 = Nothing
            ORMX$ = TRIM$(Str$(ORMA% + 1))
            While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
            NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
            '
            Call ACTUREGISTRO("cajabanc", "codicom_cor", CONDI$, NRECIB$, REGAF&)
            NRECIB1$ = "AS-X-0000-00" & Mid$(NRECIB$, 4, 6)
            Call ACTUREGISTRO("cajabanc", "codicom_lar", CONDI$, NRECIB1$, REGAF&)
        End If
        Screen.MousePointer = 1
    End If
    '
End Sub

Function VALIMESAN%()
   VMA% = 0
   '
   If Len(MESAN) = 19 Then
     Desde% = FECHANUM(Left$(MESAN, 8))
     If Desde% > 0 Then
       Hasta% = FECHANUM(Mid$(MESAN, 12, 8))
       If Hasta% >= Desde% Then
         VMA% = 1
         'DESDAT = Int(CDbl(CVDAT(DESDE%)))
         'HASDAT = Int(CDbl(CVDAT(HASTA%)))
          DESDAT = FETEXCOR1$(Desde%)   ' FECHATEXLAR$(DESDE%)
          HASDAT = FETEXCOR1$(Hasta%)   ' FECHATEXLAR$(HASTA%)
         CORRESP$ = MESAN
         GoTo 98
       End If
     End If
   End If
   '
   CORRESP$ = MESANO$(MESAN, Desde%, Hasta%)
   If CORRESP$ <> "" Then
       VMA% = 1
       'DESDAT = Int(CDbl(CVDAT(DESDE%)))
       'HASDAT = Int(CDbl(CVDAT(HASTA%)))
       DESDAT = FETEXCOR1$(Desde%)
       HASDAT = FETEXCOR1$(Hasta%)
     Else
       Call MENSERR(24, "Falta Elegir Período Mensual")
   End If
   '
98 VALIMESAN% = VMA%
   '
End Function

Private Sub GananciasExpoRet_Click()

131     Call DESHASFECHA(Des%, Has%, PERIO$)
        If Des% < 1 Then GoTo 99
        DES1$ = FETEXCOR1$(Des%)
        HAS1$ = FETEXCOR1$(Has%)
        Call APERBASDAT("CABAN")

        SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
        SQLX$ = SQLX$ + "WHERE FECHCONTA >= '" & DES1$ & "'"
        SQLX$ = SQLX$ + "AND FECHCONTA <= '" & HAS1$ & "' "
        SQLX$ = SQLX$ + "AND TIPOMOVIM = 'RTGAN' "
        SQLX$ = SQLX$ + "AND NUPROVE > 0 "
        SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC "
135     On Error Resume Next
        Set PERLISTMP = READSET(SQLX$)

        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 135
        End If
        On Error GoTo 0
        '
        With PERLISTMP
           '
           On Error Resume Next
           .MoveFirst
           If Err Then
             Screen.MousePointer = 1
             Call COMUNI("No Existen Retenciones en el Rango de Fechas Elegido")
             On Error GoTo 0
             GoTo 131
           End If
           On Error GoTo 0

           CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
           CommonDialog1.DialogTitle = "Archivo Origen de Datos"
           CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
           CommonDialog1.FilterIndex = 1
           On Error GoTo 99
           CommonDialog1.ShowOpen
           On Error GoTo 0

           DoEvents
           NOMBASALID$ = UCase$(CommonDialog1.FileName)
           ASALID1$ = TRIM$(NOMBASALID$)
           If ASALID1$ = "" Then GoTo 99

           Screen.MousePointer = 11

           NX1% = FILEOPEN%(ASALID1$, 2, 128)
           CAREGI% = 0
           Do While .EOF = False
             NC% = XVALO(!NUPROVE)
             COMPROBANTELARGO$ = SAFETEXT$(!CODICOM_LAR)
             CODICOMPRO$ = "05"
             If UCase$(Left$(COMPROBANTELARGO$, 3)) = "NC-" Then CODICOMPRO$ = "03"
             If UCase$(Left$(COMPROBANTELARGO$, 3)) = "ND-" Then CODICOMPRO$ = "04"
             If UCase$(Left$(COMPROBANTELARGO$, 3)) = "FC-" Then CODICOMPRO$ = "01"
             If UCase$(Left$(COMPROBANTELARGO$, 3)) = "OP-" Then CODICOMPRO$ = "06"
             If UCase$(Left$(COMPROBANTELARGO$, 3)) = "RB-" Then CODICOMPRO$ = "02"
             '
             FECHA% = CVINT(!fechconta)
             FECHARET$ = FECHATEXLAR$(FECHA%)
             '
             NUMECOMPRO$ = Left(SAFETEXT$(!NumeComp), 16)
             '
             CONDI$ = "NUPROVE=" & NC% & " AND CODICOM_LAR='" & COMPROBANTELARGO$ & "' AND TIPOMOVIM='OPAG'"
             IMPORTECOMPROBANTE# = XVALO(VALOBADA("CAJABANC", "VALABSCOMP", CONDI$))
             '
             CODIGOIMP& = XVALO(!CodigoImpuesto)
             CODIGOREG& = XVALO(!CodigoRegimen)
             If CODIGOIMP& = 0 Or CODIGOREG& = 0 Then
               RECORD$ = FILTERGETRECORD$("DARGAPRO", 1, MKI$(NC%))
               CATIB1& = Asc(Mid$(RECORD$, 3, 1))
               REGISTROCATE$ = REGLEIDO$("TACRETG", CATIB1&)
               CODIGOIMP& = CVI(Mid$(REGISTROCATE$, 51, 2))
               CODIGOREG& = CVI(Mid$(REGISTROCATE$, 53, 2))
             End If
             '
             CodigoOperacion = 1
             '
             BASECALCULO# = XVALO(VALOBADA("CAJABANC", "BASEIMPO", CONDI$))
             '
             FechaRetencion$ = FECHARET$
             '
             POSIV% = PIVACLI((-1) * NC%)
             If POSIV% = 1 Then
               CODIGOCONDICION% = 1
             ElseIf POSIV% = 2 Then
               CODIGOCONDICION% = 2
             Else
               CODIGOCONDICION% = 3
             End If
             '
             SUJETOSUSPENDIDO$ = Space(1)
             '
             ImporteRetencion# = XVALO(!IHABECOMP) - XVALO(!Idebecomp)
             '
             PORCENTAJEEXCLUSION$ = Space(6)
             '
             FECHABOLETIN$ = Space(10)
             '
             TIPODOCUMENTO$ = "80"
             '
             Cuit$ = CuitCli$((-1) * NC%)
             Cuit$ = REPLACAR$(Cuit$, "-", "")
             '
             CERTIFICADOORIGINAL$ = Space(14)
             '
             DENOMINACIONORDENANTE$ = Space(30)
             '
             ACRECENTAMIENTO$ = Space(1)
             '
             CUITPAISRETENIDO$ = Space(11)
             '
             CUITDELORDENANTE$ = Space(11)
             '
             TX$ = Space(197)
             Call REPLA(TX$, CODICOMPRO$, 1, 2)
             Call REPLA(TX$, FECHARET$, 3, 10)
             Call REPLA(TX$, NUMECOMPRO$, 13, 16)
             Call REPLA(TX$, REPLACAR$(FORMATNUM$(IMPORTECOMPROBANTE#, "F16.2"), ".", ","), 29, 16)
             Call REPLA(TX$, FORMATNUM$(CODIGOIMP&, "I3"), 45, 3)
             Call REPLA(TX$, FORMATNUM$(CODIGOREG&, "I3"), 49, 3)
             Call REPLA(TX$, FORMATNUM$(CodigoOperacion, "I1"), 52, 1)
             Call REPLA(TX$, REPLACAR$(FORMATNUM$(BASECALCULO#, "F14.2"), ".", ","), 53, 14)
             Call REPLA(TX$, FechaRetencion$, 67, 10)
             Call REPLA(TX$, FORMATNUM$(CODIGOCONDICION%, "I2"), 77, 2)
             Call REPLA(TX$, SUJETOSUSPENDIDO$, 79, 1)
             Call REPLA(TX$, REPLACAR$(FORMATNUM$(ImporteRetencion#, "F14.2"), ".", ","), 80, 14)
             Call REPLA(TX$, PORCENTAJEEXCLUSION$, 94, 6)
             Call REPLA(TX$, FECHABOLETIN$, 100, 10)
             Call REPLA(TX$, TIPODOCUMENTO$, 110, 2)
             Call REPLA(TX$, Cuit$, 112, 20)
             Call REPLA(TX$, CERTIFICADOORIGINAL$, 132, 14)

             CAREGI% = CAREGI% + 1
             Print #NX1%, TX$
             .MoveNext
           Loop
        End With
        Close #NX1%
        ASALID2$ = REPLACAR(ASALID1$, "\", "/")
        Screen.MousePointer = 1
        Call COMUNI("Archivos de Transferencia\ '" + ASALID2$ + "' \Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.")
             
99      Exit Sub


'131     Call DESHASFECHA(Des%, Has%, PERIO$)
'        If Des% < 1 Then GoTo 99
'        DES1$ = FETEXCOR1$(Des%)
'        HAS1$ = FETEXCOR1$(Has%)
'        Call APERBASDAT("CABAN")
'
'        SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
'        SQLX$ = SQLX$ + "WHERE FECHCONTA >= '" & DES1$ & "'"
'        SQLX$ = SQLX$ + "AND FECHCONTA <= '" & HAS1$ & "' "
'        SQLX$ = SQLX$ + "AND TIPOMOVIM = 'RTGAN' "
'        SQLX$ = SQLX$ + "AND NUPROVE > 0 "
'        SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC "
'135     On Error Resume Next
'        Set PERLISTMP = READSET(SQLX$)
'
'        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'           On Error GoTo 0
'           Call CapturaErrorOpen
'           GoTo 135
'        End If
'        On Error GoTo 0
'        '
'        With PERLISTMP
'           '
'           On Error Resume Next
'           .MoveFirst
'           If Err Then
'             Screen.MousePointer = 1
'             Call COMUNI("No Existen Retenciones en el Rango de Fechas Elegido")
'             On Error GoTo 0
'             GoTo 131
'           End If
'           On Error GoTo 0
'
'           CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
'           CommonDialog1.DialogTitle = "Archivo Origen de Datos"
'           CommonDialog1.Filter = "Archivo de Texto(*.TXT)|*.TXT|"
'           CommonDialog1.FilterIndex = 1
'           On Error GoTo 99
'           CommonDialog1.ShowOpen
'           On Error GoTo 0
'
'           DoEvents
'           NOMBASALID$ = UCase$(CommonDialog1.FileName)
'           ASALID1$ = TRIM$(NOMBASALID$)
'           If ASALID1$ = "" Then GoTo 99
'
'           Screen.MousePointer = 11
'
'           NX1% = FILEOPEN%(ASALID1$, 2, 128)
'           CAREGI% = 0
'           Do While .EOF = False
'             NC% = XVALO(!NUPROVE)
'             COMPROBANTELARGO$ = SAFETEXT$(!CODICOM_LAR)
'             CODICOMPRO$ = "05"
'             If UCase$(Left$(COMPROBANTELARGO$, 3)) = "NC-" Then CODICOMPRO$ = "03"
'             If UCase$(Left$(COMPROBANTELARGO$, 3)) = "ND-" Then CODICOMPRO$ = "04"
'             If UCase$(Left$(COMPROBANTELARGO$, 3)) = "FC-" Then CODICOMPRO$ = "01"
'             If UCase$(Left$(COMPROBANTELARGO$, 3)) = "OP-" Then CODICOMPRO$ = "06"
'             If UCase$(Left$(COMPROBANTELARGO$, 3)) = "RB-" Then CODICOMPRO$ = "02"
'             '
'             FECHA% = CVINT(!fechconta)
'             FECHARET$ = FECHATEXLAR$(FECHA%)
'             '
'             NUMECOMPRO$ = Left(SAFETEXT$(!NumeComp), 16)
'             '
'             CONDI$ = "NUPROVE=" & NC% & " AND CODICOM_LAR='" & COMPROBANTELARGO$ & "' AND TIPOMOVIM='OPAG'"
'             IMPORTECOMPROBANTE# = XVALO(VALOBADA("CAJABANC", "VALABSCOMP", CONDI$))
'             '
'             CODIGOIMP& = XVALO(!CodigoImpuesto)
'             CODIGOREG& = XVALO(!CodigoRegimen)
'             If CODIGOIMP& = 0 Or CODIGOREG& = 0 Then
'               RECORD$ = FILTERGETRECORD$("DARGAPRO", 1, MKI$(NC%))
'               CATIB1& = Asc(Mid$(RECORD$, 3, 1))
'               REGISTROCATE$ = REGLEIDO$("TACRETG", CATIB1&)
'               CODIGOIMP& = CVI(Mid$(REGISTROCATE$, 51, 2))
'               CODIGOREG& = CVI(Mid$(REGISTROCATE$, 53, 2))
'             End If
'             '
'             CodigoOperacion = 1
'             '
'             BASECALCULO# = XVALO(VALOBADA("CAJABANC", "BASEIMPO", CONDI$))
'             '
'             FechaRetencion$ = FECHARET$
'             '
'             POSIV% = PIVACLI((-1) * NC%)
'             If POSIV% = 1 Then
'               CODIGOCONDICION% = 1
'             ElseIf POSIV% = 2 Then
'               CODIGOCONDICION% = 2
'             Else
'               CODIGOCONDICION% = 3
'             End If
'             '
'             SUJETOSUSPENDIDO$ = Space(1)
'             '
'             ImporteRetencion# = XVALO(!IHABECOMP) - XVALO(!Idebecomp)
'             '
'             PORCENTAJEEXCLUSION$ = Space(6)
'             '
'             FECHABOLETIN$ = Space(10)
'             '
'             TIPODOCUMENTO$ = "80"
'             '
'             Cuit$ = CuitCli$((-1) * NC%)
'             Cuit$ = REPLACAR$(Cuit$, "-", "")
'             '
'             CERTIFICADOORIGINAL$ = Space(14)
'             '
'             DENOMINACIONORDENANTE$ = Space(30)
'             '
'             ACRECENTAMIENTO$ = Space(1)
'             '
'             CUITPAISRETENIDO$ = Space(11)
'             '
'             CUITDELORDENANTE$ = Space(11)
'             '
'             TX$ = Space(197)
'             Call REPLA(TX$, CODICOMPRO$, 1, 2)
'             Call REPLA(TX$, FECHARET$, 3, 10)
'             Call REPLA(TX$, NUMECOMPRO$, 13, 16)
'             Call REPLA(TX$, REPLACAR$(FORMATNUM$(IMPORTECOMPROBANTE#, "F16.2"), ".", ","), 29, 16)
'             Call REPLA(TX$, FORMATNUM$(CODIGOIMP&, "I3"), 45, 3)
'             Call REPLA(TX$, FORMATNUM$(CODIGOREG&, "I3"), 48, 3)
'             Call REPLA(TX$, FORMATNUM$(CodigoOperacion, "I1"), 51, 1)
'             Call REPLA(TX$, REPLACAR$(FORMATNUM$(BASECALCULO#, "F14.2"), ".", ","), 52, 14)
'             Call REPLA(TX$, FechaRetencion$, 66, 10)
'             Call REPLA(TX$, FORMATNUM$(CODIGOCONDICION%, "I2"), 76, 2)
'             Call REPLA(TX$, SUJETOSUSPENDIDO$, 78, 1)
'             Call REPLA(TX$, REPLACAR$(FORMATNUM$(ImporteRetencion#, "F14.2"), ".", ","), 79, 14)
'             Call REPLA(TX$, PORCENTAJEEXCLUSION$, 93, 6)
'             Call REPLA(TX$, FECHABOLETIN$, 99, 10)
'             Call REPLA(TX$, TIPODOCUMENTO$, 109, 2)
'             Call REPLA(TX$, Cuit$, 111, 20)
'             Call REPLA(TX$, CERTIFICADOORIGINAL$, 131, 14)
'             Call REPLA(TX$, DENOMINACIONORDENANTE$, 145, 30)
'             Call REPLA(TX$, ACRECENTAMIENTO$, 175, 1)
'             Call REPLA(TX$, CUITPAISRETENIDO$, 176, 11)
'             Call REPLA(TX$, CUITDELORDENANTE$, 187, 11)
'
'             CAREGI% = CAREGI% + 1
'             Print #NX1%, TX$
'             .MoveNext
'           Loop
'        End With
'        Close #NX1%
'        ASALID2$ = REPLACAR(ASALID1$, "\", "/")
'        Screen.MousePointer = 1
'        Call COMUNI("Archivos de Transferencia\ '" + ASALID2$ + "' \Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.")
'
'99      Exit Sub

End Sub


Private Sub GananciasListaRet_Click()

131     Call DESHASFECHA(Des%, Has%, PERIO$)
        If Des% < 1 Then GoTo 99
        DES1$ = FETEXCOR1$(Des%)
        HAS1$ = FETEXCOR1$(Has%)
        Call APERBASDAT("CABAN")

        SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
        SQLX$ = SQLX$ + "WHERE FECHCONTA >= '" & DES1$ & "'"
        SQLX$ = SQLX$ + "AND FECHCONTA <= '" & HAS1$ & "' "
        SQLX$ = SQLX$ + "AND TIPOMOVIM = 'RTGAN' "
        SQLX$ = SQLX$ + "AND NUPROVE > 0 "
        SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC "
135     On Error Resume Next
        Set PERLISTMP = READSET(SQLX$)

        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 135
        End If
        On Error GoTo 0
        '
        With PERLISTMP
          '
          On Error Resume Next
          .MoveFirst
          If Err Then
            Call COMUNI("No existen Retenciones en el Rango de Fechas Elegido")
            On Error GoTo 0
            GoTo 131
          End If
          On Error GoTo 0
          JJ& = 0
          ZZ$ = REGBLAN("DERETGAN")
          NIVCONXXX% = NIVCONFI%
          Do While .EOF = False
            If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 170
            Call REPLA(ZZ$, MKI(CVINT(!fechconta)), 1, 2)
            Call REPLA(ZZ$, SAFETEXT$(!CODICOM_LAR), 3, 22)
            NC% = XVALO(!NUPROVE)
            CTP$ = rasocli$((-1) * NC%)
            Call REPLA(ZZ$, CTP$, 25, 32)
            IMPER# = XVALO(!IHABECOMP) - XVALO(!Idebecomp)
            Call REPLA(ZZ$, MKD(IMPER#), 70, 8)
            CPR$ = TRIM$(Mid$(SAFETEXT$(!OpciMovim), 4))
            CCTT$ = CuitCli$((-1) * NC%)
            Call REPLA(ZZ$, CCTT$, 57, 13)
            JJ& = JJ& + 1
            Call GRAREG("DERETGAN", ZZ$, JJ&)
170         .MoveNext
          Loop
          Call SETULTREG("DERETGAN", JJ&)
          Call FILESORT("DERETGAN", "DERETGAN", 1, 1, "ASC") ' POR FECHA
          Call FINAL("*DERETGAN")
          '
        End With
        PERLISTMP.Close
        Set PERLISTMP = Nothing
        
99      Exit Sub

End Sub


Private Sub IIBB_BsAsExpoPer_Click()

21         Call DESHASFECHA(Des%, Has%, PERIO$)
           If Des% < 1 Then GoTo 99
           'JANDY SQL
           DES1$ = FETEXCOR1$(Des%)
           HAS1$ = FETEXCOR1$(Has%)
           Call APERBASDAT("CABAN")
           Dim PERLISTMP As ADODB.Recordset
           Set PERLISTMP = New ADODB.Recordset
           PERLISTMP.CursorLocation = adUseClient
           SQLX$ = "SELECT * FROM CAJABANC "
           SQLX$ = SQLX$ + "WHERE FECHCONTA >= '" & DES1$ & "'"
           SQLX$ = SQLX$ + "AND FECHCONTA <= '" & HAS1$ & "' "
           SQLX$ = SQLX$ + "AND TIPOMOVIM = 'IBRU' "
           SQLX$ = SQLX$ + "AND VARIMOVIM = 'PERC' "
           SQLX$ = SQLX$ + "AND NUPROVE > 0 "
           SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC "
25         On Error Resume Next
           PERLISTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 25
           End If
           On Error GoTo 0
           '
           Screen.MousePointer = 11
           
           CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
           CommonDialog1.DialogTitle = "Archivo Origen de Datos"
           '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
           CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
           CommonDialog1.FilterIndex = 1
           On Error GoTo 99
           CommonDialog1.ShowOpen
           On Error GoTo 0
        
           DoEvents
           NOMBASALID$ = UCase$(CommonDialog1.FileName)
           ASALID1$ = TRIM$(NOMBASALID$)
           If ASALID1$ = "" Then
              Screen.MousePointer = 1: GoTo 99
           End If
           
           NX1% = FILEOPEN%(ASALID1$, 2, 128)
           CAREGI% = 0
            With PERLISTMP
            '
              On Error Resume Next
              .MoveFirst
              If Err Then
                Screen.MousePointer = 1
                Call COMUNI("No existen Percepciones en el Rango de Fechas Elegido")
                On Error GoTo 0
                GoTo 21
              End If
              On Error GoTo 0
              
              Call SELECHO("PROVINC")
              PROV$ = VALACT1$("PROVINC")
              
              Do While .EOF = False
                NC% = !NUPROVE
                
                PROVIPRO$ = PROVCLI$((-1) * NC%)
                If PROVIPRO$ <> PROV$ Then GoTo 80
                
                CCTT$ = CuitCli$((-1) * NC%)
                FECPERC$ = CVDAT((CVINT(!fechconta)))
                CODIKOMLAR$ = SAFETEXT$(!CODICOM_LAR)
                TICOMP$ = "F"
                If UCase$(Left$(CODIKOMLAR$, 3)) = "NC-" Then TICOMP$ = "C"
                If UCase$(Left$(CODIKOMLAR$, 3)) = "RB-" Then TICOMP$ = "R"
                If UCase$(Left$(CODIKOMLAR$, 3)) = "ND-" Then TICOMP$ = "D"
                If UCase$(Left$(CODIKOMLAR$, 3)) = "DP-" Then TICOMP$ = "C"
                If UCase$(Left$(CODIKOMLAR$, 3)) = "CP-" Then TICOMP$ = "D"
                '
                LECOMP$ = Mid$(CODIKOMLAR$, 4, 1)
                NUMSUCURSAL$ = Mid$(CODIKOMLAR$, 6, 4)
                NUEMISION$ = Right$(CODIKOMLAR$, 8)
                '
                'SE AGREGA EL MONTO IMPONIBLE PARA LA IMPORTACION AL SIAP
                'ES POR CT-102926-FL
                FEMMI = CVINT(!fechconta)
                BASEIMPON# = INEFAC#(CODIKOMLAR$, FEMMI, NC%)
                ' Verifica si es un credito
                SIGNO% = 0: If BASEIMPON# < 0 Then SIGNO% = 1
                BASEIMPON# = Abs(BASEIMPON#)
                '
                BASEIMPO$ = FORMATNUM(BASEIMPON#, "F12.2")
                BASEIMPO$ = TRIM$(REPLACAR$(BASEIMPO$, ".", ","))
                While Len(TRIM$(BASEIMPO$)) < 12: BASEIMPO$ = "0" + BASEIMPO$: Wend
                ' Agrega el signo cuando se trata de un credito
                If SIGNO% = 1 Then Call REPLA(BASEIMPO$, "-", 1, 1)
                '
                IMPOXRE# = XVALO(!Idebecomp) - XVALO(!IHABECOMP)
                IMPOXRE# = Abs(IMPOXRE#)
                '
                IMPOPERC$ = FORMATNUM(IMPOXRE#, "F11.2")
                '
                IMPOPERC$ = TRIM$(REPLACAR$(IMPOPERC$, ".", ","))
                While Len(TRIM$(IMPOPERC$)) < 11: IMPOPERC$ = "0" + IMPOPERC$: Wend
                ' Agrega el signo cuando se trata de un credito
                If SIGNO% = 1 Then Call REPLA(IMPOPERC$, "-", 1, 1)
                '
                Call REPLA(TX$, CCTT$, 1, 13) ' CUIT
                Call REPLA(TX$, FECPERC$, 14, 10) ' FECHA PERCEPCION
                Call REPLA(TX$, TICOMP$, 24, 1) ' TIPO DE COMPROBANTE
                Call REPLA(TX$, LECOMP$, 25, 1) ' LETRA COMPROBANTE
                Call REPLA(TX$, NUMSUCURSAL$, 26, 4) ' NUMERO SUCURSAL
                Call REPLA(TX$, NUEMISION$, 30, 8) ' NUMERO DE EMISION
                'Call REPLA(TX$, IMPOPERC$, 38, 11) ' IMPORTE DE PERCEPCION
                'SE AGREGA EL MONTO IMPONIBLE PARA EL ARCHIVO DE TRASNFERENCIA DE LAS PERCEPCIONES AL SIAP
                'POR CONSULTA DE DOSIVAC CT-102926-FL - OT-10-0154
                Call REPLA(TX$, BASEIMPO$, 38, 12) ' MONTO IMPONIBLE DE PERCEPCION
                Call REPLA(TX$, IMPOPERC$, 50, 11) ' IMPORTE DE PERCEPCION
                Call REPLA(TX$, "A", 61, 1) ' SE AGREGA SEGUN NUEVA NORMATIVA DE ARBA SETIEMBRE 2011
                
                CAREGI% = CAREGI% + 1
                Print #NX1%, TX$
80
                .MoveNext
              Loop
            End With
            Close #NX1%
            ASALID2$ = REPLACAR(ASALID1$, "\", "/")
            Screen.MousePointer = 1
            Call COMUNI("Archivos de Transferencia\ '" + ASALID2$ + "' \Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.")

99          Exit Sub

End Sub

Private Sub IIBB_BsAsExpoRet_Click()

         Call DESHASFECHA(Des%, Has%, PERIO$)
         If Des% < 1 Then GoTo 99
         '
         If InStr(1, EMPREAC$, "FERVI") > 0 Then
              FF1% = FECHANUM("01/01/80")
              FIN& = ULTREG&("SURETIB")
              For U& = 1 To FIN&
                  XX$ = REGLEIDO$("SURETIB", U&)
                  NOP& = CVS(Mid$(XX$, 7, 4))
                  If NOP& = 75157 Or NOP& = 75158 Then
                        Call REPLA(XX$, MKI(FF1%), 1, 2)
                        Call GRAREG("SURETIB", XX$, U&)
                  End If
              Next U&
         End If
         '
         CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
         CommonDialog1.DialogTitle = "Archivo Origen de Datos"
         '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
         CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
         CommonDialog1.FilterIndex = 1
         On Error GoTo 99
         CommonDialog1.ShowOpen
         On Error GoTo 0
         NOMBASALID$ = UCase$(CommonDialog1.FileName)
         ASALID$ = TRIM$(NOMBASALID$)
         If ASALID$ = "" Then GoTo 99
         
         NX% = FILEOPEN%(ASALID$, 2, 46)
         ASALID1$ = LUCOM$ + "DDJJRDET.TXT"
         NX1% = FILEOPEN%(ASALID1$, 2, 128)
         '
10                   cuitagen$ = CONTROL$("CUITAGEN", "CUITAGEN")
         cuitagen$ = InputBox$("Ingrese C.U.I.T. Agente de Retención", "C.U.I.T. Agente Retencion", cuitagen$)
         If VALICUIT%(cuitagen$) <> 1 Then
           Call MENSERR(24, "C.U.I.T. no Válido")
           GoTo 10
         End If
         '
         REBLA$ = Space$(219)
         '\\\OT-05-0285-WAR-12/05/05///
         Call REPLA(REBLA$, cuitagen$, 1, 13)
         'PL1$ = FECHATEX(CVINT(CVDAT(HAS%)))
         PL11 = CVDAT(Has%)
         PL1$ = Format(PL11, "YYYYMM")
         Call REPLA(REBLA$, PL1$, 14, 6)
         Call REPLA(REBLA$, "S", 73, 1)
         PL22 = CVDAT(Has%)
         PL2$ = Format(PL22, "YYYYMMDD")
         Call REPLA(REBLA$, PL2$, 144, 8)
         Call REPLA(REBLA$, "OTA0001", 152, 7)
         Call REPLA(REBLA$, "PSO", 167, 3)
         Call REPLA(REBLA$, "2.5000", 188, 6)
         Call REPLA(REBLA$, "S", 210, 1)
         '\\\OT-05-0285-WAR-FIN///
         CAREGI% = 0
         For U& = 1 To ULTREG&("SURETIB")
20                     XX$ = REGLEIDO$("SURETIB", U&)
           FEMIX% = CVI(Left$(XX$, 2))
           If FEMIX% >= Des% Then
             If FEMIX% <= Has% Then
               TTXX$ = Space$(45)
                 CUITTX$ = TRIM$(Mid$(XX$, 51, 14))
                 If TRIM$(CUITTX$) = "" Or VALICUIT%(CUITTX$) <> 1 Then
                   CUITTX$ = InputBox$("Ingrese Numero de CUIT", TRIM$(Mid$(XX$, 21, 30)))
                   Call REPLA(XX$, CUITTX$, 51, 14)
                   Call GRAREG("SURETIB", XX$, U&)
                   GoTo 20
                 End If
               Call REPLA(TTXX$, CUITTX$, 1, 13)
                 FEX = CVI(Left$(XX$, 2))
                 FEXU$ = FECHATEX$(FEX)
                 FEXU$ = Left$(FEXU$, 6) + "20" + Mid$(FEXU$, 7, 2)
               Call REPLA(TTXX$, FEXU$, 14, 10)
                 SUCUX$ = "0001"
               Call REPLA(TTXX$, SUCUX$, 24, 4)
                 NUCOX$ = TRIM$(Str$(CVS(Mid$(XX$, 3, 4))))
                 While Len(NUCOX$) < 8: NUCOX$ = "0" + NUCOX$: Wend
               Call REPLA(TTXX$, NUCOX$, 28, 8)
                 IMPOX$ = REPLACAR$(TRIM$(Format$(CVD(Mid$(XX$, 87, 8)), "#########0.00")), ".", ",")
                 While Len(IMPOX$) < 11: IMPOX$ = "0" + IMPOX$: Wend
                 'While Len(IMPOX$) < ANCHIMPOR%: IMPOX$ = "0" + IMPOX$: Wend
               'Call REPLA(TTXX$, IMPOX$, 36, ANCHIMPOR%)
               Call REPLA(TTXX$, IMPOX$, 36, 11)
               Call REPLA(TTXX$, "A", 47, 1) ' SE AGREGA SEGUN NUEVA NORMATIVA DE ARBA SETIEMBRE 2011
               If XVALO(IMPOX$) > 0.005 Then ' SE AGREGA PORQUE EL SIAP PROTESTA SI ES CERO
                    Print #NX%, TTXX$
               End If
               '\\\OT-05-0285-WAR-12/05/05///
               TTYY$ = REBLA$
               Call REPLA(TTYY$, CUITTX$, 20, 13)
               NC% = CVI(Mid$(XX$, 95, 2))
               RS1$ = rasocli(NC% * (-1))
               Call REPLA(TTYY$, AJUSTI(RS1$, 40), 33, 40)
               DO1$ = TRIM$(DOMICLI(NC% * (-1)))
               LO1$ = TRIM$(LOCACLI(NC% * (-1)))
               DP1$ = AJUSTI(DO1$ + " " + LO1$, 70)
               Call REPLA(TTYY$, DP1$, 74, 70)
               'NUMERO COMPROBANTE
                 NUCOX$ = TRIM$(Str$(CVS(Mid$(XX$, 3, 4))))
                 While Len(NUCOX$) < 8: NUCOX$ = "0" + NUCOX$: Wend
               Call REPLA(TTYY$, NUCOX$, 159, 8)
               'BASE IMPONIBLE
               IMPOX$ = FORMATNUM$(CVD(Mid$(XX$, 79, 8)), "F16.2")
               Call REPLA(TTYY$, AJUSTD(IMPOX$, 16), 170, 16)
             ' IMPORTE RETENIDO
               If CVD(Mid$(XX$, 87, 8)) < 0.005 Then GoTo 201 ' SALTEA REGISTROS SIN RETENCION
               IMPOX2$ = FORMATNUM$(CVD(Mid$(XX$, 87, 8)), "F16.2")
               Call REPLA(TTYY$, AJUSTD(IMPOX2$, 16), 194, 16)
               ' tasa retenida - MODIFICADO 3/8/2005 - ANTES ERA FIJO =2.5000
               TARET = 0
               On Error Resume Next
               TARET = CVD(Mid$(XX$, 87, 8)) / CVD(Mid$(XX$, 79, 8))
               On Error GoTo 0
               Call REPLA(TTYY$, FORMATNUM$(TARET, "F6.4"), 188, 6)
               'NUMERO DE ORDEN
               CAREGI% = CAREGI% + 1
               NUO$ = AJUSTD(CStr(CAREGI%), 9)
               Call REPLA(TTYY$, NUO$, 211, 9)
               Print #NX1%, TTYY$
               '\\\OT-05-0285-WAR-FIN///
             End If
           End If
201      Next U&
         ' PONE MARCA DE FIN DE ARCHIVO SEGúN MUESTRA RECIBIDA
         Print #NX1%, Chr$(26)
         ' FIN MARCA DE FIN DE ARCHIVO
         Close #NX%
         Close #NX1%
         ASALID3$ = REPLACAR(ASALID$, "\", "/")
         ASALID4$ = REPLACAR(ASALID1$, "\", "/")
         Call COMUNI("Archivos de Transferencia\ '" + ASALID3$ + "' y '" + ASALID4$ + "' \Generados con " + TRIM$(Str$(CAREGI%)) + " Registros.")
        
99       Exit Sub

End Sub


Private Sub IIBB_BsAsListaPer_Click()

21          Call DESHASFECHA(Des%, Has%, PERIO$)
            If Des% < 1 Then GoTo 99
            'JANDY SQL
            DES1$ = FETEXCOR1$(Des%)
            HAS1$ = FETEXCOR1$(Has%)
            Call APERBASDAT("CABAN")
            Dim PERLISTMP As ADODB.Recordset
            Set PERLISTMP = New ADODB.Recordset
            PERLISTMP.CursorLocation = adUseClient
            SQLX$ = "SELECT * FROM CAJABANC "
            SQLX$ = SQLX$ + "WHERE FECHCONTA >= '" & DES1$ & "'"
            SQLX$ = SQLX$ + "AND FECHCONTA <= '" & HAS1$ & "' "
            SQLX$ = SQLX$ + "AND TIPOMOVIM = 'IBRU' "
            SQLX$ = SQLX$ + "AND VARIMOVIM = 'PERC' "
            SQLX$ = SQLX$ + "AND NUPROVE > 0 "
            SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC "
25          On Error Resume Next
            PERLISTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
       
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 25
            End If
            On Error GoTo 0
            '
            With PERLISTMP
              '
              On Error Resume Next
              .MoveFirst
              If Err Then
                Call COMUNI("No existen Percepciones en el Rango de Fechas Elegido")
                On Error GoTo 0
                GoTo 21
              End If
              On Error GoTo 0
              JJ& = 0
              ZZ$ = REGBLAN("DEPERIB")
              NIVCONXXX% = NIVCONFI%
              Do While .EOF = False
                '\\\OT-08-1032-FL-12/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
                If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70
                '\\\OT-08-1032-FL-FIN
                Call REPLA(ZZ$, MKI(CVINT(!fechconta)), 1, 2)
                Call REPLA(ZZ$, !CODICOM_LAR, 3, 22)
                NC% = !NUPROVE
                CTP$ = CODICLI$((-1) * NC%)
                Call REPLA(ZZ$, CTP$, 25, 12)
                IMPER# = !Idebecomp - !IHABECOMP
                Call REPLA(ZZ$, MKD(IMPER#), 37, 8)
                CPR$ = TRIM$(Mid$(!OpciMovim, 4))
                CPR1$ = ECOARCH$("PROVINC", CPR$)
                Call REPLA(ZZ$, CPR$, 45, 1)
                CCTT$ = CuitCli$((-1) * NC%)
                Call REPLA(ZZ$, CCTT$, 70, 13)
                JJ& = JJ& + 1
                Call GRAREG("DEPERIB", ZZ$, JJ&)
70              .MoveNext
              Loop
              
              Call SETULTREG("DEPERIB", JJ&)
              Call FILESORT("DEPERIB", "DEPERIB", 1, 1, "ASC") ' POR FECHA
              Call FILESORT("DEPERIB", "DEPERIB", 6, 6, "ASC") ' POR JURISDICCIóN
              Call FINAL("*DEPERIB")
              '
            End With
            PERLISTMP.Close
            Set PERLISTMP = Nothing
            
99          Exit Sub

End Sub

Private Sub IIBB_BsAsListaRet_Click()

        Call DESHASFECHA(Des%, Has%, PERIO$)
        If Des% < 1 Then Exit Sub
        '
        If InStr(1, EMPREAC$, "FERVI") > 0 Then
             FF1% = FECHANUM("01/01/80")
             FIN& = ULTREG&("SURETIB")
             For U& = 1 To FIN&
                 XX$ = REGLEIDO$("SURETIB", U&)
                 NOP& = CVS(Mid$(XX$, 7, 4))
                 If NOP& = 75157 Or NOP& = 75158 Then
                       Call REPLA(XX$, MKI(FF1%), 1, 2)
                       Call GRAREG("SURETIB", XX$, U&)
                 End If
             Next U&
        End If
        '
        CORRESP$ = PERIO$
        Call HEADERS("SUBRETI", "")
        Call HEADERS("SUBRETI", "Corresponde a: " + CORRESP$)
        j& = 0: FIN& = ULTREG("SURETIB")
        For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          XX$ = REGLEIDO$("SURETIB", U&)
          FEMIX% = CVI(Left$(XX$, 2))
          If FEMIX% >= Des% Then
            If FEMIX% <= Has% Then
              j& = j& + 1
              Call GRAREG("SUBRETI", XX$, j&)
            End If
          End If
        Next U&
        Call SETULTREG("SUBRETI", j&)
        Call FINAL("*SUBRETI")

End Sub


Private Sub IIBB_CABAExpoRet_Click()

         Call DESHASFECHA(Des%, Has%, PERIO$)
         If Des% < 1 Then GoTo 99
         
         DESD$ = FETEXCOR1$(Des%)
         HAST$ = FETEXCOR1$(Has%)
         '
         CommonDialog1.CancelError = True
         CommonDialog1.DialogTitle = "Archivo Origen de Datos"
         CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
         CommonDialog1.FilterIndex = 1
         On Error GoTo 99
         CommonDialog1.ShowOpen
         On Error GoTo 0
         NOMBASALID$ = UCase$(CommonDialog1.FileName)
         ASALID$ = TRIM$(NOMBASALID$)
         If ASALID$ = "" Then GoTo 99
         '
         ' *** CREA ARCHIVO DE TRANSFERENCIA
         NX1% = FILEOPEN%(ASALID$, 2, 215)
         '
         Screen.MousePointer = 11
         '
         ' *** OBTIENE LAS RETENCIONES DEL PERIODO ELEGIDO Y COMPLETA EL ARCHIVO DE TRANSFERENCIA
         CONDICION$ = "FECHCONTA >= '" & DESD$ & "' AND FECHCONTA <= '" & HAST$ & "'" & " AND TIPOMOVIM='RTIB' AND OPCIMOVIM='CABA' AND IHABECOMP > 0.005"
         FIN& = CantRegistros&("CAJABANC", CONDICION$)
         ConsultaSql$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE " & CONDICION$ & " ORDER BY FECHCONTA"
         Dim RstIbCaba As ADODB.Recordset
         
         Set RstIbCaba = READSET(ConsultaSql$)
         With RstIbCaba
            Do While Not .EOF
                CAREGI& = CAREGI& + 1
                Call TRACE(24, CAREGI&, FIN&)
                
                REBLA$ = Space$(215)
                NroProveedor% = XVALO(!NUPROVE)
                DocRetenido$ = SAFETEXT$(!Notas)
                FechaRetencion% = CVINT(!fechconta)
                NroRetencion& = XVALO(!NroDeRetencion)
                ImporteRetencion# = XVALO(!IHABECOMP)

                ' TIPO DE OPERACIONES
                Call REPLA(REBLA$, "1", 1, 1)
                
                ' CODIGO DE NORMA
                Call REPLA(REBLA$, "029", 2, 3)
                
                ' FECHA RETENCION
                Call REPLA(REBLA$, FECHATEXLAR$(FechaRetencion%), 5, 10)

                ' TIPO DE COMPROBANTE
                    DocRet_Tipo$ = "09"
                    If Mid$(DocRetenido$, 1, 2) = "FC" Then DocRet_Tipo$ = "01"
                    If Mid$(DocRetenido$, 1, 2) = "ND" Then DocRet_Tipo$ = "01"
                Call REPLA(REBLA$, DocRet_Tipo$, 15, 2)
                
                ' LETRA DE COMPROBANTE
                    DocRet_Letra$ = Mid$(DocRetenido$, 4, 1)
                Call REPLA(REBLA$, DocRet_Letra$, 17, 1)
                
                ' NRO DE COMPROBANTE
                    DocRet_Numero$ = Mid$(DocRetenido$, 11, 8)
                    While Len(DocRet_Numero$) < 16: DocRet_Numero$ = "0" & DocRet_Numero$: Wend
                Call REPLA(REBLA$, DocRet_Numero$, 18, 16)
                
                ' FECHA DEL COMPROBANTE
                    DocRet_Fecha% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", "CODICOM_LAR='" & DocRetenido$ & "' AND NUPROVE=" & NroProveedor% & " AND TIPOMOVIM='FCOM'"))
                    If DocRet_Fecha% < FECHANUM("01/11/16") Then DocRet_Fecha% = FechaRetencion%
                Call REPLA(REBLA$, FECHATEXLAR$(DocRet_Fecha%), 34, 10)
                
                ' MONTO DEL COMPROBANTE
                    DocRet_Importe# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", "CODICOM_LAR='" & DocRetenido$ & "' AND NUPROVE=" & NroProveedor% & " AND TIPOMOVIM='FCOM'"))
                    DocRet_Impo$ = TRIM$(FORMATNUM$(DocRet_Importe#, "F12.2"))
                    While Len(DocRet_Impo$) < 16: DocRet_Impo$ = "0" & DocRet_Impo$: Wend
                    DocRet_Impo$ = REPLACAR$(DocRet_Impo$, ".", ",")
                Call REPLA(REBLA$, DocRet_Impo$, 44, 16)

                ' NRO DE CERTIFICADO
                    NroReten$ = TRIM$(FORMATNUM$(NroRetencion&, "I10"))
                    While Len(NroReten$) < 16: NroReten$ = "0" & NroReten$: Wend
                Call REPLA(REBLA$, NroReten$, 60, 16)
                
                ' TIPO DE DOCUMENTO (CUIT, CUIL)
                    PosicionIva% = PIVACLI%((-1) * NroProveedor%)
                    PosIva$ = "3": If PosicionIva% = 0 Then PosIva$ = "2"
                Call REPLA(REBLA$, PosIva$, 76, 1)
                
                ' NUMERO DE DOCUMENTO DEL RETENIDO
                    Cuit$ = TRIM$(CuitCli$((-1) * NroProveedor%))
                    Cuit$ = REPLACAR$(Cuit$, "-", "")
                Call REPLA(REBLA$, Cuit$, 77, 11)
                
                ' SITUACION IIBB DEL RETENIDO (CATEGORIA)
'                    NroIB$ = TRIM$(NIBRCLI$((-1) * NroProveedor%))
'                    NroIB$ = REPLACAR$(NroIB$, "-", "")
'                    CategoriaIB% = 2                        ' CONVENIO MULTILATERAL
'                    If CUIT$ = NroIB$ Then CategoriaIB% = 5 ' REGIMEN SIMPLIFICADO
NroIB$ = Cuit$
CategoriaIB% = 5
                Call REPLA(REBLA$, CStr(CategoriaIB%), 88, 1)

                ' NUMERO DE INSCRIPCION
                Call REPLA(REBLA$, NroIB$, 89, 11)
                
                ' SITUACION FRENTE AL IVA
                    SituacionIva% = 1 ' RI
                    If PosicionIva% = 6 Then SituacionIva% = 4
                    If PosicionIva% = 4 Then SituacionIva% = 3
                Call REPLA(REBLA$, CStr(PosicionIva%), 100, 1)
                
                ' RAZON SOCIAL DEL RETENIDO
                    RAZONSOCIAL0$ = rasocli$((-1) * NroProveedor%)
                    razonSocial$ = TextoEnFormatoUTF8(RAZONSOCIAL0$)
                Call REPLA(REBLA$, razonSocial$, 101, 30)
               
                ' *** DATOS DEL COMPROBANTE RETENIDO
                    CondicionTotal$ = "CODICOM_LAR='" & DocRetenido$ & "' AND NUPROVE=" & NroProveedor% & " AND TIPOMOVIM='FCOM'"
                    condicionIva$ = "CODICOM_LAR='" & DocRetenido$ & "' AND NUPROVE=" & NroProveedor% & " AND TIPOMOVIM='IVA' AND VARIMOVIM<>'EXENTO' AND VARIMOVIM<>'NO GRAV' AND VARIMOVIM<>'ADUANA' AND VARIMOVIM<>'PER-IVA'"
                    CondicionNeto$ = "CODICOM_LAR='" & DocRetenido$ & "' AND NUPROVE=" & NroProveedor% & " AND TIPOMOVIM='COMPRA'"
                    
                    DocRet_Total# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", CondicionTotal$))
                    DocRet_Iva# = XVALO(VALOBADA("CAJABANC", "IDEBECOMP", condicionIva$))
                    DocRet_Neto# = XVALO(VALOBADA("CAJABANC", "IDEBECOMP", CondicionNeto$))
                    DocRet_OtConceptos# = DocRet_Total# - DocRet_Neto# - DocRet_Iva#
                    
                ' IMPORTE OTROS CONCEPTOS
                    DocRet_OtCon$ = TRIM$(FORMATNUM$(DocRet_OtConceptos#, "F12.2"))
                    If DocRet_OtCon$ = "" Then DocRet_OtCon$ = "0.00"
                    While Len(DocRet_OtCon$) < 16: DocRet_OtCon$ = "0" & DocRet_OtCon$: Wend
                    DocRet_OtCon$ = REPLACAR$(DocRet_OtCon$, ".", ",")
                Call REPLA(REBLA$, DocRet_OtCon$, 131, 16)
                
                ' IMPORTE IVA
                    DocRet_ImporteIva$ = TRIM$(FORMATNUM$(DocRet_Iva#, "F12.2"))
                    While Len(DocRet_ImporteIva$) < 16: DocRet_ImporteIva$ = "0" & DocRet_ImporteIva$: Wend
                    DocRet_ImporteIva$ = REPLACAR$(DocRet_ImporteIva$, ".", ",")
                Call REPLA(REBLA$, DocRet_ImporteIva$, 147, 16)
                
                ' IMPORTE SUJETO DE RETENCION (NETO)
                    DocRet_ImporteNeto$ = TRIM$(FORMATNUM$(DocRet_Neto#, "F12.2"))
                    While Len(DocRet_ImporteNeto$) < 16: DocRet_ImporteNeto$ = "0" & DocRet_ImporteNeto$: Wend
                    DocRet_ImporteNeto$ = REPLACAR$(DocRet_ImporteNeto$, ".", ",")
                Call REPLA(REBLA$, DocRet_ImporteNeto$, 163, 16)

                ' ALICUOTA
                    ALICUOTA# = (ImporteRetencion# / DocRet_Neto#) * 100
                    Alicuo$ = TRIM$(FORMATNUM$(ALICUOTA#, "F5.2"))
                    While Len(Alicuo$) < 5: Alicuo$ = "0" & Alicuo$: Wend
                    Alicuo$ = REPLACAR$(Alicuo$, ".", ",")
                Call REPLA(REBLA$, Alicuo$, 179, 5)

                ' RETENCION PRACTICADA
                    ImporteRet$ = TRIM$(FORMATNUM$(ImporteRetencion#, "F12.2"))
                    While Len(ImporteRet$) < 16: ImporteRet$ = "0" & ImporteRet$: Wend
                    ImporteRet$ = REPLACAR$(ImporteRet$, ".", ",")
                Call REPLA(REBLA$, ImporteRet$, 184, 16)

                ' MONTO TOTAL RETENIDO
                Call REPLA(REBLA$, ImporteRet$, 200, 16)
                
                Print #NX1%, REBLA$
                .MoveNext
            Loop
         End With
         Close #NX1%
         On Error Resume Next
         RstIbCaba.Close
         Set RstIbCaba = Nothing
         On Error GoTo 0
         
         Screen.MousePointer = 1
         
         ASALID3$ = REPLACAR(ASALID$, "\", "/")
         Call COMUNI("Archivos de Transferencia\ '" + ASALID3$ + "\Generado con " + TRIM$(Str$(CAREGI&)) + " Registros.")
        
99       Exit Sub

End Sub

Private Sub IIBB_CMExpoPer_Click(Index As Integer)

86     Call DESHASFECHA(Des%, Has%, PERIO$)
       If Des% < 1 Then GoTo 99
       DES1$ = FETEXCOR1$(Des%)
       HAS1$ = FETEXCOR1$(Has%)
       Call APERBASDAT("CABAN")
       Dim PERLISTMP0 As ADODB.Recordset
       Set PERLISTMP0 = New ADODB.Recordset
       PERLISTMP0.CursorLocation = adUseClient
       SQLX$ = "SELECT * FROM CAJABANC "
       SQLX$ = SQLX$ + "WHERE FECHCONTA >= '" & DES1$ & "'"
       SQLX$ = SQLX$ + "AND FECHCONTA <= '" & HAS1$ & "' "
       SQLX$ = SQLX$ + "AND TIPOMOVIM = 'IBRU' "
       SQLX$ = SQLX$ + "AND VARIMOVIM = 'PERC' "
       SQLX$ = SQLX$ + "AND NUPROVE > 0 "
       SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC "
87     On Error Resume Next
       PERLISTMP0.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 87
       End If
       On Error GoTo 0
       '
       Screen.MousePointer = 11
       CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
       CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
       CommonDialog1.FilterIndex = 1
       On Error GoTo 99
       CommonDialog1.ShowOpen
       On Error GoTo 0
    
       DoEvents
       NOMBASALID$ = UCase$(CommonDialog1.FileName)
       ASALID1$ = TRIM$(NOMBASALID$)
       If ASALID1$ = "" Then
         Screen.MousePointer = 1: GoTo 99
       End If
       ASALID2$ = Mid$(NOMBASALID$, 1, Len(NOMBASALID$) - 4) & " - ADUANERAS.TXT"
       
       NX1% = FILEOPEN%(ASALID1$, 2, 51)
       NX2% = FILEOPEN%(ASALID2$, 2, 56)
       CAREGI% = 0
       With PERLISTMP0
        '
          On Error Resume Next
          .MoveFirst
          If Err Then
            Screen.MousePointer = 1
            Call COMUNI("No Existen Percepciones en el Rango de Fechas Elegido")
            On Error GoTo 0
            GoTo 99
          End If
          On Error GoTo 0
          
          Do While .EOF = False
            NC% = !NUPROVE
            CCTT$ = CuitCli$((-1) * NC%)
            FECPERC$ = CVDAT((CVINT(!fechconta)))
            CODIKOMLAR$ = SAFETEXT$(!CODICOM_LAR)
            TICOMP$ = "F"
            If UCase$(Left$(CODIKOMLAR$, 3)) = "NC-" Then TICOMP$ = "C"
            If UCase$(Left$(CODIKOMLAR$, 3)) = "RB-" Then TICOMP$ = "R"
            If UCase$(Left$(CODIKOMLAR$, 3)) = "ND-" Then TICOMP$ = "D"
            If UCase$(Left$(CODIKOMLAR$, 3)) = "DP-" Then TICOMP$ = "C"
            If UCase$(Left$(CODIKOMLAR$, 3)) = "CP-" Then TICOMP$ = "D"
            '
            LECOMP$ = Mid$(CODIKOMLAR$, 4, 1)
                 NUMSUCURSAL$ = REPLACAR$(Mid$(CODIKOMLAR$, 6, 4), "/", "")
                 NUMSUCURSAL$ = REPLACAR$(NUMSUCURSAL$, "-", "")
                 NUSUC& = XVALO(NUMSUCURSAL$)
                 If NUSUC& = 0 Then NUSUC& = 1
                 NUMSUCURSAL$ = CStr(NUSUC&)
                 
            Do While Len(NUMSUCURSAL$) < 4: NUMSUCURSAL$ = "0" & NUMSUCURSAL$: Loop
                 NUEMISION$ = REPLACAR$(Right$(CODIKOMLAR$, 8), "/", "")
                 NUEMISION$ = REPLACAR$(NUEMISION$, "-", "")
                 NUEMISION$ = REPLACAR$(NUEMISION$, ".", "")
                 NUCBTE& = XVALO(NUEMISION$)
                 NUEMISION$ = CStr(NUCBTE&)
            Do While Len(NUEMISION$) < 8: NUEMISION$ = "0" & NUEMISION$: Loop
            
            'NUEMISION$ = Right$(CODIKOMLAR$, 8)
            '
            FEMMI = CVINT(!fechconta)
            ' Verifica si es un credito
            SIGNO% = 0: If TICOMP$ = "C" Or TICOMP$ = "R" Then SIGNO% = 1
            ' Agrega el signo cuando se trata de un credito
            If SIGNO% = 1 Then Call REPLA(BASEIMPO$, "-", 1, 1)
            '
            IMPOXRE# = XVALO(!Idebecomp) - XVALO(!IHABECOMP)
            IMPOXRE# = Abs(IMPOXRE#)
            '
            IMPOPERC$ = FORMATNUM(IMPOXRE#, "F10.2")
            '
            IMPOPERC$ = TRIM$(REPLACAR$(IMPOPERC$, ".", ","))
            While Len(TRIM$(IMPOPERC$)) < 11: IMPOPERC$ = "0" + IMPOPERC$: Wend
            ' Agrega el signo cuando se trata de un credito
            If SIGNO% = 1 Then Call REPLA(IMPOPERC$, "-", 1, 1)
            '
            CUECONTABLE% = XVALO(!CUECONTAI)
            encontro% = 0: YY$ = ""
            For IIK& = 1 To ULTREG&("PROVINB")
                 XX$ = REGLEIDO$("PROVINB", IIK&)
                 CC& = CVI(Mid$(XX$, 27, 2))
                 If CC& = CUECONTABLE% Then
                     encontro% = 1: YY$ = XX$
                     Exit For
                 End If
            Next IIK&
            If encontro% = 1 Then
                 CODPROV_CM$ = CStr(CVI(Mid$(YY$, 37, 2)))
                 If TRIM$(CODPROV_CM$) = "" Then
                     Call COMUNI("No Se Ha Encontrado la Provincia\para la Percepción del Comprobante '" & CODIKOMLAR$ & "'.\" _
                                 & "Dicha Percepción No Se Incluye en el Archivo de Transferencia.")
                     GoTo 89
                 End If
            Else
                 Call COMUNI("No Se Ha Encontrado la Jurisdicción\para la Percepción del Comprobante '" & CODIKOMLAR$ & "'.\" _
                             & "del Proveedor '" & TRIM$(AJUSTI$(rasocli$((-1) * NC%), 32)) & "'.\ \" _
                             & "Dicha Percepción No Se Incluye en el Archivo de Transferencia.")
                 GoTo 89
            End If
            '
            DespachoImportacion$ = TRIM$(SAFETEXT$(!DespaImpo))
            '
            TX$ = ""
            If DespachoImportacion = "" Then
                 Call REPLA(TX$, CODPROV_CM$, 1, 3) ' CODIGO DE PROVINCIA S/ CONVENIO MULTILATERAL
                 Call REPLA(TX$, CCTT$, 4, 13) ' CUIT
                 Call REPLA(TX$, FECPERC$, 17, 10) ' FECHA PERCEPCION
                 Call REPLA(TX$, NUMSUCURSAL$, 27, 4) ' NUMERO SUCURSAL
                 Call REPLA(TX$, NUEMISION$, 31, 8) ' NUMERO DE EMISION
                 Call REPLA(TX$, TICOMP$, 39, 1) ' TIPO DE COMPROBANTE
                 Call REPLA(TX$, LECOMP$, 40, 1) ' LETRA COMPROBANTE
                 Call REPLA(TX$, IMPOPERC$, 41, 11) ' IMPORTE DE PERCEPCION
                 CAREGI% = CAREGI% + 1
                 Print #NX1%, TX$
            Else
                 Call REPLA(TX$, CODPROV_CM$, 1, 3) ' CODIGO DE PROVINCIA S/ CONVENIO MULTILATERAL
                 Call REPLA(TX$, CCTT$, 4, 13) ' CUIT
                 Call REPLA(TX$, FECPERC$, 17, 10) ' FECHA PERCEPCION
                 Call REPLA(TX$, DespachoImportacion$, 27, 20) ' IMPORTE DE PERCEPCION
                 Call REPLA(TX$, AJUSTD$(TRIM$(IMPOPERC$), 10), 47, 10) ' IMPORTE DE PERCEPCION
                 CAREGIADUA% = CAREGIADUA% + 1
                 Print #NX2%, TX$
            End If

89          .MoveNext
          Loop
        End With
        Close #NX1%
        Close #NX2%
        ASALID11$ = REPLACAR(ASALID1$, "\", "/")
        ASALID22$ = REPLACAR(ASALID2$, "\", "/")
        Screen.MousePointer = 1
        MensajeFinProceso$ = "Archivos de Transferencia\ '" + ASALID11$ + "' \Generado con " + TRIM$(Str$(CAREGI%)) + " Percepciones."
        If CAREGIADUA% > 0 Then MensajeFinProceso = MensajeFinProceso + "\ \Además, Se Generó el Archivo '" + ASALID22$ + "'\con " + TRIM$(Str$(CAREGIADUA%)) + " Percepciones Aduaneros"
        Call COMUNI(MensajeFinProceso$)
        
99      Exit Sub

End Sub

Private Sub IIBB_CMExpoRet_Click()

    Call COMUNI("Opción No Disponible")
    
End Sub

Private Sub IIBB_CMListaPer_Click(Index As Integer)

86       Call DESHASFECHA(Des%, Has%, PERIO$)
         If Des% < 1 Then GoTo 99
         DES1$ = FETEXCOR1$(Des%)
         HAS1$ = FETEXCOR1$(Has%)
         Call APERBASDAT("CABAN")
         Dim PERLISTMP0 As ADODB.Recordset
         Set PERLISTMP0 = New ADODB.Recordset
         PERLISTMP0.CursorLocation = adUseClient
         SQLX$ = "SELECT * FROM CAJABANC "
         SQLX$ = SQLX$ + "WHERE FECHCONTA >= '" & DES1$ & "'"
         SQLX$ = SQLX$ + "AND FECHCONTA <= '" & HAS1$ & "' "
         SQLX$ = SQLX$ + "AND TIPOMOVIM = 'IBRU' "
         SQLX$ = SQLX$ + "AND VARIMOVIM = 'PERC' "
         SQLX$ = SQLX$ + "AND NUPROVE > 0 "
         SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC "
87       On Error Resume Next
         PERLISTMP0.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    
         If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
            On Error GoTo 0
            Call CapturaErrorOpen
            GoTo 87
         End If
         On Error GoTo 0
         '
         With PERLISTMP0
             '
             On Error Resume Next
             .MoveFirst
             If Err Then
               Call COMUNI("No Existen Percepciones en el Rango de Fechas Elegido")
               On Error GoTo 0
               GoTo 86
             End If
             On Error GoTo 0
             JJ& = 0
             ZZ$ = REGBLAN("DEPERIB")
             NIVCONXXX% = NIVCONFI%
             Do While .EOF = False
               '\\\OT-08-1032-FL-12/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
               If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 88
               '\\\OT-08-1032-FL-FIN
               Call REPLA(ZZ$, MKI(CVINT(!fechconta)), 1, 2)
               Call REPLA(ZZ$, !CODICOM_LAR, 3, 22)
               NC% = !NUPROVE
               CTP$ = CODICLI$((-1) * NC%)
               Call REPLA(ZZ$, CTP$, 25, 12)
               IMPER# = !Idebecomp - !IHABECOMP
               Call REPLA(ZZ$, MKD(IMPER#), 37, 8)
               CPR$ = TRIM$(Mid$(!OpciMovim, 4))
               CPR1$ = ECOARCH$("PROVINC", CPR$)
               Call REPLA(ZZ$, CPR$, 45, 1)
               CCTT$ = CuitCli$((-1) * NC%)
               Call REPLA(ZZ$, CCTT$, 70, 13)
               JJ& = JJ& + 1
               Call GRAREG("DEPERIB", ZZ$, JJ&)
88             .MoveNext
             Loop
             Call SETULTREG("DEPERIB", JJ&)
             Call FILESORT("DEPERIB", "DEPERIB", 1, 1, "ASC") ' POR FECHA
             Call FILESORT("DEPERIB", "DEPERIB", 6, 6, "ASC") ' POR JURISDICCIóN
             Call FINAL("*DEPERIB")
             '
         End With
         PERLISTMP0.Close
         Set PERLISTMP0 = Nothing
         
99       Exit Sub
         
End Sub

Private Sub IIBB_CMListaRet_Click()

    Call COMUNI("Opción No Disponible")
    
End Sub

Private Sub IIBB_CMSircrebCoeficientes_Click()

    Call CargaCoeficientesSIRCREB
                    
End Sub

Private Sub IIBB_CMSircrebExpoRet_Click()

    Call ExportaRecaudacionesBancarias
                    
End Sub


Private Sub IIBBPercepRetenCABAExpo_Click()

         ' *** FACTURACION Y ORDENES DE PAGO

         Call DESHASFECHA(Des%, Has%, PERIO$)
         If Des% < 1 Then GoTo 99

         DESD$ = FETEXCOR1$(Des%)
         HAST$ = FETEXCOR1$(Has%)
         '
         CommonDialog1.CancelError = True
         CommonDialog1.DialogTitle = "Archivo Origen de Datos"
         CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
         CommonDialog1.FilterIndex = 1
         On Error GoTo 99
         CommonDialog1.ShowOpen
         On Error GoTo 0
         NOMBASALID$ = UCase$(CommonDialog1.FileName)
         ASALID$ = TRIM$(NOMBASALID$)
         If ASALID$ = "" Then GoTo 99
         '
         ' *** CREA ARCHIVO DE TRANSFERENCIA
         NX1% = FILEOPEN%(ASALID$, 2, 215)
         '
         Screen.MousePointer = 11
         '
         ' *** OBTIENE LAS PERCEPCIONES DEL PERIODO ELEGIDO Y COMPLETA EL ARCHIVO DE TRANSFERENCIA
         '
         CONDICION$ = "FECHCONTA >= '" & DESD$ & "' AND FECHCONTA <= '" & HAST$ & "'" & " AND TIPOMOVIM='IBRU' AND OPCIMOVIM='IB-C' AND VARIMOVIM='PERC' AND IHABECOMP > 0.005 AND NUCLIEN > 0"
         FIN& = CantRegistros&("CAJABANC", CONDICION$)
         ConsultaSql$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE " & CONDICION$ & " ORDER BY FECHCONTA"
         Dim RstIbCaba As ADODB.Recordset

         Set RstIbCaba = READSET(ConsultaSql$)
         With RstIbCaba
            On Error GoTo ControlErrorVentas
            Do While Not .EOF
                CAREGI& = CAREGI& + 1
                Call TRACE(24, CAREGI&, FIN&)

                REBLA$ = Space$(215)
                NroCliente% = XVALO(!nuclien)
                DocRetenido$ = SAFETEXT$(!CODICOM_LAR)
                FechaRetencion% = CVINT(!fechconta)
                NroRetencion& = XVALO(!NumeComp)
                ImporteRetencion# = XVALO(!IHABECOMP)

                ' *** DATOS DEL COMPROBANTE RETENIDO
                condicionIva$ = "CODICOM_LAR='" & DocRetenido$ & "' AND NUCLIEN=" & NroCliente% & " AND TIPOMOVIM='IVA'"
                CondicionNeto$ = "CODICOM_LAR='" & DocRetenido$ & "' AND NUCLIEN=" & NroCliente% & " AND TIPOMOVIM='VENTA'"

                DocRet_Importe# = XVALO(VALOBADA("CAJABANC", "IDEBECOMP", "CODICOM_LAR='" & DocRetenido$ & "' AND NUCLIEN=" & NroCliente% & " AND TIPOMOVIM='FVEN'"))
                DocRet_Total# = DocRet_Importe#
                DocRet_Iva# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", condicionIva$))
                DocRet_Neto# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CondicionNeto$))
                DocRet_OtConceptos# = DocRet_Total# - DocRet_Neto# - DocRet_Iva# - ImporteRetencion#


                ' TIPO DE OPERACIONES
                Call REPLA(REBLA$, "2", 1, 1)

                ' CODIGO DE NORMA
                Call REPLA(REBLA$, "029", 2, 3)

                ' FECHA RETENCION
                Call REPLA(REBLA$, FECHATEXLAR$(FechaRetencion%), 5, 10)

                ' TIPO DE COMPROBANTE
                    DocRet_Tipo$ = "09"
                    If Mid$(DocRetenido$, 1, 2) = "FC" Then DocRet_Tipo$ = "01"
                Call REPLA(REBLA$, DocRet_Tipo$, 15, 2)

                ' LETRA DE COMPROBANTE
                    DocRet_Letra$ = Mid$(DocRetenido$, 18, 1)
                    If DocRet_Tipo$ = "09" Then DocRet_Letra$ = " "
                Call REPLA(REBLA$, DocRet_Letra$, 17, 1)

                ' NRO DE COMPROBANTE
                    DocRet_Numero$ = Mid$(DocRetenido$, 4, 4) & Mid$(DocRetenido$, 9, 8)
                    While Len(DocRet_Numero$) < 16: DocRet_Numero$ = "0" & DocRet_Numero$: Wend
                Call REPLA(REBLA$, DocRet_Numero$, 18, 16)

                ' FECHA DEL COMPROBANTE
                    DocRet_Fecha% = FechaRetencion%
                    Call REPLA(REBLA$, FECHATEXLAR$(DocRet_Fecha%), 34, 10)

                ' MONTO DEL COMPROBANTE
                    If DocRet_Letra$ <> "B" Then
                        DocRet_Importe# = DocRet_Importe# - ImporteRetencion#
                    Else
                        DocRet_Importe# = DocRet_Neto# + DocRet_OtConceptos#
                    End If
                    DocRet_Impo$ = TRIM$(FORMATNUM$(DocRet_Importe#, "F12.2"))
                    While Len(DocRet_Impo$) < 16: DocRet_Impo$ = "0" & DocRet_Impo$: Wend
                    DocRet_Impo$ = REPLACAR$(DocRet_Impo$, ".", ",")
                Call REPLA(REBLA$, DocRet_Impo$, 44, 16)

                ' NRO DE CERTIFICADO
                    NroReten$ = "0"
                    While Len(NroReten$) < 16: NroReten$ = "0" & NroReten$: Wend
                Call REPLA(REBLA$, NroReten$, 60, 16)

                ' TIPO DE DOCUMENTO (CUIT, CUIL)
                    PosicionIva% = PIVACLI%(NroCliente%)
                    PosIva$ = "3": If PosicionIva% = 0 Then PosIva$ = "2"
                Call REPLA(REBLA$, PosIva$, 76, 1)

                ' NUMERO DE DOCUMENTO DEL RETENIDO
                    Cuit$ = TRIM$(CuitCli$(NroCliente%))
                    Cuit$ = REPLACAR$(Cuit$, "-", "")
                Call REPLA(REBLA$, Cuit$, 77, 11)

                ' SITUACION IIBB DEL RETENIDO (CATEGORIA)
'                    NroIB$ = TRIM$(NIBRCLI$(NroCliente%))
'                    NroIB$ = REPLACAR$(NroIB$, "-", "")
'                    CategoriaIB% = 2
'                    If NroIB$ = CUIT$ Or NroIB$ = "" Then CategoriaIB% = 5
CategoriaIB% = 5
                Call REPLA(REBLA$, CStr(CategoriaIB%), 88, 1)

                ' NUMERO DE INSCRIPCION
NroIB$ = Cuit$
                Call REPLA(REBLA$, NroIB$, 89, 11)

                ' SITUACION FRENTE AL IVA
                    SituacionIva% = 1 ' RI
                    If PosicionIva% = 6 Then SituacionIva% = 4
                    If PosicionIva% = 4 Then SituacionIva% = 3
                Call REPLA(REBLA$, CStr(SituacionIva%), 100, 1)

                ' RAZON SOCIAL DEL RETENIDO
                    razonSocial$ = TextoEnFormatoUTF8(rasocli$(NroCliente%))
                Call REPLA(REBLA$, razonSocial$, 101, 30)

                ' IMPORTE OTROS CONCEPTOS
                    DocRet_OtCon$ = TRIM$(FORMATNUM$(DocRet_OtConceptos#, "F12.2"))
                    If DocRet_OtCon$ = "" Or XVALO(DocRet_OtCon$) = 0 Then DocRet_OtCon$ = "0.00"
                    While Len(DocRet_OtCon$) < 16: DocRet_OtCon$ = "0" & DocRet_OtCon$: Wend
                    DocRet_OtCon$ = REPLACAR$(DocRet_OtCon$, ".", ",")
                Call REPLA(REBLA$, DocRet_OtCon$, 131, 16)

                ' IMPORTE IVA
                    If DocRet_Letra$ = "B" Then DocRet_Iva# = 0
                    DocRet_ImporteIva$ = TRIM$(FORMATNUM$(DocRet_Iva#, "F12.2"))
                    If DocRet_ImporteIva$ = "" Or XVALO(DocRet_ImporteIva$) = 0 Then DocRet_ImporteIva$ = "0.00"
                    While Len(DocRet_ImporteIva$) < 16: DocRet_ImporteIva$ = "0" & DocRet_ImporteIva$: Wend
                    DocRet_ImporteIva$ = REPLACAR$(DocRet_ImporteIva$, ".", ",")
                Call REPLA(REBLA$, DocRet_ImporteIva$, 147, 16)

                ' IMPORTE SUJETO DE RETENCION (NETO)
                    DocRet_ImporteNeto$ = TRIM$(FORMATNUM$(DocRet_Neto#, "F12.2"))
                    While Len(DocRet_ImporteNeto$) < 16: DocRet_ImporteNeto$ = "0" & DocRet_ImporteNeto$: Wend
                    DocRet_ImporteNeto$ = REPLACAR$(DocRet_ImporteNeto$, ".", ",")
                Call REPLA(REBLA$, DocRet_ImporteNeto$, 163, 16)

                ' ALICUOTA
On Error GoTo ControlErrorPercepciones
                    ALICUOTA# = (ImporteRetencion# / DocRet_Neto#) * 100
                    Alicuo$ = TRIM$(FORMATNUM$(ALICUOTA#, "F5.2"))
                    While Len(Alicuo$) < 5: Alicuo$ = "0" & Alicuo$: Wend
                    Alicuo$ = REPLACAR$(Alicuo$, ".", ",")
                Call REPLA(REBLA$, Alicuo$, 179, 5)
On Error GoTo 0

                ' RETENCION PRACTICADA
                    ImporteRet$ = TRIM$(FORMATNUM$(ImporteRetencion#, "F12.2"))
                    While Len(ImporteRet$) < 16: ImporteRet$ = "0" & ImporteRet$: Wend
                    ImporteRet$ = REPLACAR$(ImporteRet$, ".", ",")
                Call REPLA(REBLA$, ImporteRet$, 184, 16)

                ' MONTO TOTAL RETENIDO
                Call REPLA(REBLA$, ImporteRet$, 200, 16)

                TotalPercepcion# = TotalPercepcion# + ImporteRetencion#

                Print #NX1%, REBLA$
25              .MoveNext
            Loop
            On Error GoTo 0
         End With
         'Close #NX1%
         On Error Resume Next
         RstIbCaba.Close
         Set RstIbCaba = Nothing
         On Error GoTo 0

         '
         ' *** OBTIENE LAS RETENCIONES DEL PERIODO ELEGIDO Y COMPLETA EL ARCHIVO DE TRANSFERENCIA
         '
         CONDICION$ = "FECHCONTA >= '" & DESD$ & "' AND FECHCONTA <= '" & HAST$ & "'" & " AND TIPOMOVIM='RTIB' AND OPCIMOVIM='CABA' AND IHABECOMP > 0.005 AND NUPROVE > 0"
         FIN& = CantRegistros&("CAJABANC", CONDICION$)
         ConsultaSql$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE " & CONDICION$ & " ORDER BY FECHCONTA"
         Dim RstIBRETCABA As ADODB.Recordset

         Set RstIBRETCABA = READSET(ConsultaSql$)
         With RstIBRETCABA
            Do While Not .EOF
                CAREGI& = CAREGI& + 1
                Call TRACE(24, CAREGI&, FIN&)

                REBLA$ = Space$(215)
                NroProveedor% = XVALO(!NUPROVE)
                DocRetenido$ = SAFETEXT$(!Notas)
                FechaRetencion% = CVINT(!fechconta)
                NroRetencion& = XVALO(!NroDeRetencion)
                ImporteRetencion# = XVALO(!IHABECOMP)

                ' TIPO DE OPERACIONES
                Call REPLA(REBLA$, "1", 1, 1)

                ' CODIGO DE NORMA
                Call REPLA(REBLA$, "029", 2, 3)

                ' FECHA RETENCION
                Call REPLA(REBLA$, FECHATEXLAR$(FechaRetencion%), 5, 10)

                ' TIPO DE COMPROBANTE
                    DocRet_Tipo$ = "09"
                    If Mid$(DocRetenido$, 1, 2) = "FC" Then DocRet_Tipo$ = "01"
                    If Mid$(DocRetenido$, 1, 2) = "ND" Then DocRet_Tipo$ = "02"
                    If Mid$(DocRetenido$, 1, 2) = "RC" Then DocRet_Tipo$ = "07"
                Call REPLA(REBLA$, DocRet_Tipo$, 15, 2)

                ' LETRA DE COMPROBANTE
                    DocRet_Letra$ = Mid$(DocRetenido$, 4, 1)
                Call REPLA(REBLA$, DocRet_Letra$, 17, 1)

                ' NRO DE COMPROBANTE
                    DocRet_Numero$ = Mid$(DocRetenido$, 6, 4) & Mid$(DocRetenido$, 11, 8)
                    While Len(DocRet_Numero$) < 16: DocRet_Numero$ = "0" & DocRet_Numero$: Wend
                Call REPLA(REBLA$, DocRet_Numero$, 18, 16)

                ' FECHA DEL COMPROBANTE
                    DocRet_Fecha% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", "CODICOM_LAR='" & DocRetenido$ & "' AND NUPROVE=" & NroProveedor% & " AND TIPOMOVIM='FCOM'"))
                    If DocRet_Fecha% < FECHANUM("01/11/16") Then DocRet_Fecha% = FechaRetencion%
                Call REPLA(REBLA$, FECHATEXLAR$(DocRet_Fecha%), 34, 10)

                ' MONTO DEL COMPROBANTE
                    DocRet_Importe# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", "CODICOM_LAR='" & DocRetenido$ & "' AND NUPROVE=" & NroProveedor% & " AND TIPOMOVIM='FCOM'"))
                    DocRet_Impo$ = TRIM$(FORMATNUM$(DocRet_Importe#, "F12.2"))
                    While Len(DocRet_Impo$) < 16: DocRet_Impo$ = "0" & DocRet_Impo$: Wend
                    DocRet_Impo$ = REPLACAR$(DocRet_Impo$, ".", ",")
                Call REPLA(REBLA$, DocRet_Impo$, 44, 16)

                ' NRO DE CERTIFICADO
                    NroReten$ = TRIM$(FORMATNUM$(NroRetencion&, "I10"))
                    While Len(NroReten$) < 16: NroReten$ = "0" & NroReten$: Wend
                Call REPLA(REBLA$, NroReten$, 60, 16)

                ' TIPO DE DOCUMENTO (CUIT, CUIL)
                    PosicionIva% = PIVACLI%((-1) * NroProveedor%)
                    PosIva$ = "3": If PosicionIva% = 0 Then PosIva$ = "2"
                Call REPLA(REBLA$, PosIva$, 76, 1)

                ' NUMERO DE DOCUMENTO DEL RETENIDO
                    Cuit$ = TRIM$(CuitCli$((-1) * NroProveedor%))
                    Cuit$ = REPLACAR$(Cuit$, "-", "")
                Call REPLA(REBLA$, Cuit$, 77, 11)

                ' SITUACION IIBB DEL RETENIDO (CATEGORIA)
'                    NroIB$ = TRIM$(NIBRCLI$((-1) * NroProveedor%))
'                    NroIB$ = REPLACAR$(NroIB$, "-", "")
'                    CategoriaIB% = 2                        ' CONVENIO MULTILATERAL
'                    If CUIT$ = NroIB$ Then CategoriaIB% = 5 ' REGIMEN SIMPLIFICADO
NroIB$ = Cuit$
CategoriaIB% = 5
                Call REPLA(REBLA$, CStr(CategoriaIB%), 88, 1)

                ' NUMERO DE INSCRIPCION
                Call REPLA(REBLA$, NroIB$, 89, 11)

                ' SITUACION FRENTE AL IVA
                    SituacionIva% = 1 ' RI
                    If PosicionIva% = 6 Then SituacionIva% = 4
                    If PosicionIva% = 4 Then SituacionIva% = 3
                Call REPLA(REBLA$, CStr(SituacionIva%), 100, 1)

                ' RAZON SOCIAL DEL RETENIDO
                    RAZONSOCIAL0$ = rasocli$((-1) * NroProveedor%)
                    razonSocial$ = TextoEnFormatoUTF8(RAZONSOCIAL0$)
                Call REPLA(REBLA$, razonSocial$, 101, 30)

                ' *** DATOS DEL COMPROBANTE RETENIDO
                    CondicionTotal$ = "CODICOM_LAR='" & DocRetenido$ & "' AND NUPROVE=" & NroProveedor% & " AND TIPOMOVIM='FCOM'"
                    condicionIva$ = "CODICOM_LAR='" & DocRetenido$ & "' AND NUPROVE=" & NroProveedor% & " AND TIPOMOVIM='IVA' AND VARIMOVIM<>'EXENTO' AND VARIMOVIM<>'NO GRAV' AND VARIMOVIM<>'ADUANA' AND VARIMOVIM<>'PER-IVA'"
                    CondicionNeto$ = "CODICOM_LAR='" & DocRetenido$ & "' AND NUPROVE=" & NroProveedor% & " AND TIPOMOVIM='COMPRA'"

                    DocRet_Total# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", CondicionTotal$))
                    DocRet_Iva# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", condicionIva$))
                    DocRet_Neto# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CondicionNeto$))
                    DocRet_OtConceptos# = DocRet_Total# - DocRet_Neto# - DocRet_Iva#

                ' IMPORTE OTROS CONCEPTOS
                    DocRet_OtCon$ = TRIM$(FORMATNUM$(DocRet_OtConceptos#, "F12.2"))
                    If DocRet_OtCon$ = "" Then DocRet_OtCon$ = "0.00"
                    While Len(DocRet_OtCon$) < 16: DocRet_OtCon$ = "0" & DocRet_OtCon$: Wend
                    DocRet_OtCon$ = REPLACAR$(DocRet_OtCon$, ".", ",")
                Call REPLA(REBLA$, DocRet_OtCon$, 131, 16)

                ' IMPORTE IVA
                    DocRet_ImporteIva$ = TRIM$(FORMATNUM$(DocRet_Iva#, "F12.2"))
                    If DocRet_ImporteIva$ = "" Then DocRet_ImporteIva$ = "0.00"
                    While Len(DocRet_ImporteIva$) < 16: DocRet_ImporteIva$ = "0" & DocRet_ImporteIva$: Wend
                    DocRet_ImporteIva$ = REPLACAR$(DocRet_ImporteIva$, ".", ",")
                Call REPLA(REBLA$, DocRet_ImporteIva$, 147, 16)

                ' IMPORTE SUJETO DE RETENCION (NETO)
                    DocRet_ImporteNeto$ = TRIM$(FORMATNUM$(DocRet_Neto#, "F12.2"))
                    If DocRet_ImporteNeto$ = "" Then DocRet_ImporteNeto$ = "0.00"
                    While Len(DocRet_ImporteNeto$) < 16: DocRet_ImporteNeto$ = "0" & DocRet_ImporteNeto$: Wend
                    DocRet_ImporteNeto$ = REPLACAR$(DocRet_ImporteNeto$, ".", ",")
                Call REPLA(REBLA$, DocRet_ImporteNeto$, 163, 16)

On Error GoTo ControlErrorRetenciones
                ' ALICUOTA
                    ALICUOTA# = (ImporteRetencion# / DocRet_Neto#) * 100
                    Alicuo$ = TRIM$(FORMATNUM$(ALICUOTA#, "F5.2"))
                    While Len(Alicuo$) < 5: Alicuo$ = "0" & Alicuo$: Wend
                    Alicuo$ = REPLACAR$(Alicuo$, ".", ",")
                Call REPLA(REBLA$, Alicuo$, 179, 5)
On Error GoTo 0

                ' RETENCION PRACTICADA
                    ImporteRet$ = TRIM$(FORMATNUM$(ImporteRetencion#, "F12.2"))
                    While Len(ImporteRet$) < 16: ImporteRet$ = "0" & ImporteRet$: Wend
                    ImporteRet$ = REPLACAR$(ImporteRet$, ".", ",")
                Call REPLA(REBLA$, ImporteRet$, 184, 16)

                ' MONTO TOTAL RETENIDO
                Call REPLA(REBLA$, ImporteRet$, 200, 16)

                TotalRetencion# = TotalRetencion# + ImporteRetencion#

                Print #NX1%, REBLA$
50              .MoveNext
            Loop
         End With
         Close #NX1%
         On Error Resume Next
         RstIBRETCABA.Close
         Set RstIBRETCABA = Nothing
         On Error GoTo 0

         'MsgBox "Total Retencion: " & CStr(TotalRetencion#) & vbCrLf & "Total Percepcion: " & CStr(TotalPercepcion#)
         'MsgBox "Archivo Transferencia Generado"
         If EXISTE%(ASALID$) > 0 Then

            ' *** CARGA UNA MATRIZ CON LAS PERCEPCIONES Y RETENCIONES PARA LUEGO ORDENAR
            NARCHI2% = FILEOPEN%(ASALID$, 1, 215) 'ABRE EL ARCHIVO DE LECTURA
            CantidadRegistros& = LOF(NARCHI2%)
            Dim MatrizOrdenaRegistros() As String
            ReDim MatrizOrdenaRegistros(CantidadRegistros&, 2)

            FilaMatriz& = 0
            Do While Not EOF(NARCHI2%) ' RECORRE
                Line Input #NARCHI2%, XX$ ' LEE CADA REGISTRO
                FilaMatriz& = FilaMatriz& + 1
                    FechaRetPer$ = Mid$(XX$, 5, 10)
                MatrizOrdenaRegistros(FilaMatriz&, 1) = Mid$(FechaRetPer$, 7, 4) & Mid$(FechaRetPer$, 4, 2) & Mid$(FechaRetPer$, 1, 2)
                MatrizOrdenaRegistros(FilaMatriz&, 2) = XX$
            Loop
            Close #NARCHI2%
            'MsgBox "Matriz generada"

            ' *** ORDENAMIENTO BURBUJA POR FECHA - ES REQUERIDO POR EL IMPORTADOR DEL GOB DE LA CIUDAD
            Dim iMin    As Long
            Dim iMax    As Long
            Dim Vectemp1 As String                   ' -- variable temporal
            Dim Vectemp2 As String                   ' -- variable temporal
            Dim Pos     As Long
            Dim i       As Long

            iMin = 1
            iMax = FilaMatriz&

            While iMax > iMin
                Pos = iMin
                For i = iMin To iMax - 1
                    If MatrizOrdenaRegistros(i, 1) > MatrizOrdenaRegistros(i + 1, 1) Then
                        Vectemp1 = MatrizOrdenaRegistros(i + 1, 1)
                        Vectemp2 = MatrizOrdenaRegistros(i + 1, 2)

                        MatrizOrdenaRegistros(i + 1, 1) = MatrizOrdenaRegistros(i, 1)
                        MatrizOrdenaRegistros(i + 1, 2) = MatrizOrdenaRegistros(i, 2)

                        MatrizOrdenaRegistros(i, 1) = Vectemp1
                        MatrizOrdenaRegistros(i, 2) = Vectemp2

                        Pos = i
                    End If
                Next i
                iMax = Pos
            Wend
            'MsgBox "Matriz ordenada"

            ' *** REGENERA EL ARCHIVO CON PERCEPCIONES Y RETENCIONES, PERO ESTA VEZ ORDENADO POR FECHA
            On Error Resume Next
            Kill ASALID$
            If Err = 0 Then
                NX1% = FILEOPEN%(ASALID$, 2, 215)
                For i = 1 To FilaMatriz
                    REBLA$ = MatrizOrdenaRegistros(i, 2)
                    Print #NX1%, REBLA$
                Next i
                Close #NX1%
                'MsgBox "Archivo ordenado"
            Else
                Call MENSERR(24, "Error al Ordenar Archivo de Retenciones y Percepciones IIBB.\ \No Considerar el Mismo para Importar en ARCIBA.")
            End If
            On Error GoTo 0

         End If

         Screen.MousePointer = 1

         ASALID3$ = REPLACAR(ASALID$, "\", "/")
         Call COMUNI("Archivos de Transferencia\ '" + ASALID3$ + "\Generado con " + TRIM$(Str$(CAREGI&)) + " Registros.")

99       Exit Sub


ControlErrorVentas:

    Call MENSERR(24, "Error en Percepciones IIBB CABA\ \Cliente: " & rasocli$(NroCliente%) & "\Comprobante: " & DocRetenido$)
    Screen.MousePointer = 1
    GoTo 99

ControlErrorPercepciones:

    Call MENSERR(24, "Error en Percepciones IIBB CABA\ \Cliente: " & rasocli$(NroCliente%) & "\Comprobante: " & DocRetenido$ & ".\ \Movimiento No Incluido")
    Screen.MousePointer = 1
    GoTo 25

ControlErrorRetenciones:

    Call MENSERR(24, "Error en Percepciones IIBB CABA\ \Cliente: " & rasocli$(NroProveedor% * -1) & "\Comprobante: " & DocRetenido$ & ".\ \Movimiento No Incluido")
    Screen.MousePointer = 1
    GoTo 50
    
End Sub

Private Sub Image2_Click()
   TXT$ = MESAN.TEXT
   If TXT$ = "" Then Exit Sub
   If MESANO$(TXT$, Des%, Has%) = TXT$ Then
      If DERACCE%("CIEMECOM", "Cierre y Re-Apertura Mensual Compras") >= 2 Then
         If COMALTER%("Confirme Cierre del Período Mensual " + TXT$ + "\\Cancelar\Cerrar") = 2 Then
           Call GRACONTROL("", "CIEMECOM", TRIM$(Str$(Has%)))
           Image2.Visible = False
           Image1.Visible = True
         End If
      End If
   End If
End Sub

Private Sub Image1_Click()
   TXT$ = MESAN.TEXT
   If TXT$ = "" Then Exit Sub
   If MESANO$(TXT$, Des%, Has%) = TXT$ Then
      If DERACCE%("CIEMECOM", "Cierre y Re-Apertura Mensual Compras") >= 2 Then
         If COMALTER%("Confirme Re-Apertura del Período Mensual\\Cancelar\Re-Abrir") = 2 Then
            Call GRACONTROL("", "CIEMECOM", TRIM$(Str$(Des% - 1)))
            Image1.Visible = False
            Image2.Visible = True
         End If
      End If
   End If
End Sub


Private Sub IvaExpoPer_Click()
        
    If VALIMESAN% < 1 Then Exit Sub
    Call GeneraArchTransfPercepIVA(Desde%, Hasta%)

End Sub


Private Sub MESAN_Change()
   TXT$ = MESAN.TEXT
   Image2.Visible = True
   Image1.Visible = False
   If TXT$ = "" Then Exit Sub
   If MESANO$(TXT$, Des%, Has%) = TXT$ Then
     If XVALO(CONTROL("", "CIEMECOM")) >= Has% Then
       Image2.Visible = False
       Image1.Visible = True
     End If
   End If
End Sub

Private Sub MESAN_DblClick()
   'Call Command12_Click
   Call DESHASFECHA(Des%, Has%, PERIO$)
   MESAN = PERIO$
   '
End Sub

Sub GENSUDICOM(Optional OK%)
    '
    'jandy sql
    Screen.MousePointer = 11
    '
    Dim IMASI#(32767)
    CUEMX% = 0
    '
    OK% = 0
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM CAJABANC "
'    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "where FechConta >=  '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + "Order by FechConta ASC "
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
'    SQLX$ = "SELECT * FROM CAJABANC "
'    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
'
'    Set CABAMENS = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'    Screen.MousePointer = 1
    '
    JJ& = 0: REBLA$ = REGBLAN$("SUDICOM")
    J2& = 0: REBL2$ = REGBLAN$("ANAICOM")
    J3& = 0: REBL3$ = REGBLAN$("ANCENCOS") '\\\OT-06-0238-WAR-09/06/06///
    
    Screen.MousePointer = 11
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        OK% = 1
        GoTo 99
      End If
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      KU& = 0
      NIVCONXXX% = NIVCONFI%
      Do While Not .EOF
        KU& = KU& + 1
        Call TRACE(20, KU&, FIN&)
        '\\\OT-08-1032-OD-11/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
        If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70 'VER QUE FACTURAS MOSTRAR EN LISTADO
        '\\\OT-08-1032-OD-FIN

        ZZ$ = REBLA$
        NUMEFA$ = !CODICOM_LAR
        NCLIE% = !NUPROVE
        'FECHAD = Int(CDbl(!FechConta))
        FEFAXI% = CVINT(!fechconta)
        FECHAD = FECHATEXLAR$(FEFAXI%)
        
        ITOTA# = !ValAbsComp
        SIG% = Sgn(!IHABECOMP - !Idebecomp)
        MOBASI$ = TRIM$(!TipoMovim)
        SUMACO# = 0
        '
        ' LO QUE SIGUE ES PARA CONSIDERAR CLIENTES VARIOS
        RSCL$ = TRIM$(rasocli$((-1) * NCLIE%))
        CTCL$ = CuitCli$((-1) * NCLIE%)
        If UCase$(Left$(RSCL$, 6)) = "VARIOS" Then
          REFEMO$ = !RefeCom
          PPXX% = InStr(REFEMO$, "//")
          If PPXX% > 0 Then
            RSCL1$ = TRIM$(Left$(REFEMO$, PPXX% - 1))
            CTCL1$ = TRIM$(Mid$(REFEMO$, PPXX% + 2))
            If RSCL1$ <> "" Then
              If CTCL1$ <> "" Then
                RSCL$ = RSCL1$
                CTCL$ = CTCL1$
              End If
            End If
          End If
        End If
        '
        If MOBASI$ = "FCOM" Then
          Call REPLA(ZZ$, MKI$(CVINT(!fechconta)), 1, 2)
          Call REPLA(ZZ$, !CODICOM_LAR, 3, 18)
          Call REPLA(ZZ$, RSCL$, 21, 30)
          Call REPLA(ZZ$, CTCL$, 51, 16)
          Call REPLA(ZZ$, MKD$(ITOTA# * SIG%), 67, 8)
          Call REPLA(ZZ$, CODCUE$(!CUECONTAI), 75, 12)
          Call REPLA(ZZ$, MKD$(!Idebecomp), 87, 8)
          Call REPLA(ZZ$, MKD$(!IHABECOMP), 95, 8)
          JJ& = JJ& + 1
          Call GRAREG("SUDICOM", ZZ$, JJ&)
        End If
        '
        SUMACO# = SUMACO# + !Idebecomp - !IHABECOMP
        '
        CUEII% = !CUECONTAI
        If CUEII% > 0 Then
          IMASI#(CUEII%) = IMASI#(CUEII%) + (!Idebecomp - !IHABECOMP)
          If CUEII% > CUEMX% Then CUEMX% = CUEII%
        End If
        '
        'Z2$ = REBL2$
        'Call REPLA(Z2$, MKI$(!CueContaI), 1, 2)
        'Call REPLA(Z2$, MKI$(NCLIE%), 3, 2)
        'Call REPLA(Z2$, !CodiCom_Lar, 5, 18)
        'Call REPLA(Z2$, MKI$(CVINT(!FechConta)), 23, 2)
        'Call REPLA(Z2$, MKD$(ITOTA# * SIG%), 25, 8)
        'J2& = J2& + 1
        'Call GRAREG("ANAICOM", Z2$, J2&)
        '
        Dim CABAMENS As ADODB.Recordset
        Set CABAMENS = New ADODB.Recordset
        CABAMENS.CursorLocation = adUseClient
        SQLX$ = "Select * from CAJABANC "
        SQLX$ = SQLX$ + "WHERE FechConta = '" & FECHAD & "' "
        SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX$ = SQLX$ + "AND NuProve = " & NCLIE% & " "
        SQLX$ = SQLX$ + "AND TipoMovim <> '" & MOBASI$ & "' "
        SQLX$ = SQLX$ + "AND TipoMovim <> 'OPAG' "    ' PARA EVITAR PAGOS CONTADO
        SQLX$ = SQLX$ + "AND TipoMovim <> 'CHEMI' "   ' PARA EFITAR PAGOS CONTADO
        '
27      On Error Resume Next
        CABAMENS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 27
        End If
        On Error GoTo 0
         With CABAMENS
         If Not (.EOF And .BOF) Then
           .MoveFirst
          Do While Not .EOF
'          .FindFirst SQLX$
'25        If .NoMatch = False Then
            ZZ$ = REBLA$
            '
            If MOBASI$ = "FCOM" Then
              Call REPLA(ZZ$, MKI$(CVINT(!fechconta)), 1, 2)
              Call REPLA(ZZ$, !CODICOM_LAR, 3, 18)
              Call REPLA(ZZ$, RSCL$, 21, 30)
              Call REPLA(ZZ$, CTCL$, 51, 16)
              Call REPLA(ZZ$, CODCUE$(!CUECONTAI), 75, 12)
              Call REPLA(ZZ$, MKD$(!Idebecomp), 87, 8)
              Call REPLA(ZZ$, MKD$(!IHABECOMP), 95, 8)
              JJ& = JJ& + 1
              Call GRAREG("SUDICOM", ZZ$, JJ&)
            End If
            '
            SUMACO# = SUMACO# + !Idebecomp - !IHABECOMP
            '
            Z2$ = REBL2$
            Call REPLA(Z2$, MKI$(!CUECONTAI), 1, 2)
            Call REPLA(Z2$, MKI$(NCLIE%), 3, 2)
            Call REPLA(Z2$, !CODICOM_LAR, 5, 18)
            Call REPLA(Z2$, MKI$(CVINT(!fechconta)), 23, 2)
            Call REPLA(Z2$, MKD$(!Idebecomp - !IHABECOMP), 25, 8)
            J2& = J2& + 1
            Call GRAREG("ANAICOM", Z2$, J2&)
            '
            Z3$ = Space$(8) + Z2$
            CENCO$ = "": If IsNull(!CenCosto) = False Then CENCO$ = !CenCosto
            Call REPLA(Z3$, CENCO$, 1, 8)
            '\\\OT-06-0238-WAR-09/06/06///
            If TRIM$(CENCO$) <> "" Then
              J3& = J3& + 1
              Call GRAREG("ANCENCOS", Z3$, J3&)
            End If
            '\\\OT-06-0238-WAR-FIN///
            '
            CUEII% = !CUECONTAI
            If CUEII% > 0 Then
              IMASI#(CUEII%) = IMASI#(CUEII%) + (!Idebecomp - !IHABECOMP)
              If CUEII% > CUEMX% Then CUEMX% = CUEII%
            End If
            '
           '.FindNext SQLX$
           ' GoTo 25
           .MoveNext
           Loop
          End If
        End With
        CABAMENS.Close
        Set CABAMENS = Nothing
        '
        If Abs(SUMACO#) >= 0.005 Then
          Call MENSERR(24, "No Balancea Comprobante " + NUMEFA$ + "\de " + TRIM$(rasocli$((-1) * NCLIE%)) + "\  \Debe Anularse y Pasarse de Nuevo.")
        End If
        '
70    .MoveNext
      Loop
    End With
    
    Call SETULTREG("SUDICOM", JJ&)
    Call SETULTREG("ANAICOM", J2&)
    '\\\OT-06-0238-WAR-09/06/06///'AGREGO CENTRO DE COSTOS POR OP
    
    Dim CAJABAMENS As ADODB.Recordset
    Set CAJABAMENS = New ADODB.Recordset
    CAJABAMENS.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND  TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "AND VariMovim = 'CGAST' "
26  On Error Resume Next
    CAJABAMENS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 26
    End If
    On Error GoTo 0
    
    '
    With CAJABAMENS        ' CajaBanc
'      .FindFirst SQLX1$
'33    If .NoMatch = False Then
      If Not (.EOF And .BOF) Then
       .MoveFirst
       Do While Not .EOF
        '\\\OT-08-1032-OD-11/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
        If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 80 'VER QUE FACTURAS MOSTRAR EN LISTADO
        '\\\OT-08-1032-OD-FIN

        Z3$ = REBL3$
        CENCO$ = ""
        If IsNull(!CenCosto) = False Then
          If TRIM$(!CenCosto) <> "" Then
            CENCO$ = !CenCosto
          End If
        End If
        Call REPLA(Z3$, CENCO$, 1, 8)
        If TRIM$(CENCO$) <> "" Then
          '
          Call REPLA(Z3$, MKI$(!CUECONTAI), 9, 2)
          Call REPLA(Z3$, MKI$(!NUPROVE), 11, 2)
          Call REPLA(Z3$, SAFETEXT$(!CODICOM_LAR), 13, 18)
          Call REPLA(Z3$, MKI$(CVINT(!fechconta)), 31, 2)
          Call REPLA(Z3$, MKD$(!Idebecomp - !IHABECOMP), 33, 8)
          '
          J3& = J3& + 1
          Call GRAREG("ANCENCOS", Z3$, J3&)
        End If
        '
        If InStr(1, EMPREAC$, "J Y M") > 0 Then
          '
          Z2$ = REBL2$
          Call REPLA(Z2$, MKI$(!CUECONTAI), 1, 2)
          Call REPLA(Z2$, MKI$(!NUPROVE), 3, 2)
          Call REPLA(Z2$, SAFETEXT$(!CODICOM_LAR), 5, 18)
          Call REPLA(Z2$, MKI$(CVINT(!fechconta)), 23, 2)
          Call REPLA(Z2$, MKD$(!Idebecomp - !IHABECOMP), 25, 8)
          J2& = J2& + 1
          Call GRAREG("ANAICOM", Z2$, J2&)
          '
        End If
        '.FindNext SQLX1$
        'GoTo 33
80      .MoveNext
        Loop
      End If
    End With
    CAJABAMENS.Close
    Set CAJABAMENS = Nothing
    
    ' AGREGA LOS ASIENTOS DE DIARIOS CON CENTRO DE COSTO
    'Call APERBASDAT("PLANCUEN")
    Dim ASIENDIAR As ADODB.Recordset
    Set ASIENDIAR = New ADODB.Recordset
    SQLX$ = "SELECT * FROM DETAASIEN "
    SQLX$ = SQLX$ + "where FechAsien >= '" & DESDAT & "'"
    SQLX$ = SQLX$ + "AND FechAsien <= '" & HASDAT & "'"
    SQLX$ = SQLX$ + "AND STATPASE > 0" ' SE INCLUYE PARA NO CONSIDERAR LOS ASIENTOS ANULADOS26  On Error Resume Next
    ASIENDIAR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    '
    With ASIENDIAR
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           If IsNull(!CenCos) = False Then
              CENCO$ = !CenCos
              If TRIM(CENCO$) <> "" Then ' EXCLUYE LOS QUE NO TIENEN CENTRO
                 Call REPLA(Z3$, CENCO$, 1, 8)
                 Call REPLA(Z3$, MKI$(!CUECOINT), 9, 2)
                 Call REPLA(Z3$, MKI$(0), 11, 2)
                 Call REPLA(Z3$, !nuidasien, 13, 18)
                 Call REPLA(Z3$, MKI$(CVINT(!FECHASIEN)), 31, 2)
                 Call REPLA(Z3$, MKD$(!idebepase - !ihabepase), 33, 8)
                 '
                 J3& = J3& + 1
                 Call GRAREG("ANCENCOS", Z3$, J3&)
              End If
           End If
         .MoveNext
         Loop
      End If
      On Error GoTo 0
    End With
    Call SETULTREG("ANCENCOS", J3&)
    '
    Call HEADERS("SUDICOM", "")
    Call HEADERS("SUDICOM", "Corresponde a: " + TRIM$(CORRESP$))
    Call HEADERS("ANAICOM", "")
    Call HEADERS("ANAICOM", "Corresponde a: " + TRIM$(CORRESP$))
    Call HEADERS("ANAICOL", "")
    Call HEADERS("ANAICOL", "Corresponde a: " + TRIM$(CORRESP$))
    Call HEADERS("ANCENCOS", "")
    Call HEADERS("ANCENCOS", "Corresponde a: " + TRIM$(CORRESP$))
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "J Y M") > 0 Then
       CORRESP$ = "Desde: " & DESDAT & " - Hasta: " & HASDAT
       Call HEADERS("ANAICOL", "")
       Call HEADERS("ANAICOL", TRIM$(CORRESP$))
    End If

    Call FILESORT("ANCENCOS", "", 1, 5, "ASC")
    '
    TTXX$ = Chr$(1) + Chr$(2) + MKI$(Hasta%)
    v& = 0: FIN& = ULTREG&("ASIDIAR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("ASIDIAR", U&)
      If Left$(X$, 4) <> TTXX$ Then
        v& = v& + 1
        Call GRAREG("ASIDIAR", X$, v&)
      End If
    Next U&
    '
    AGRECU% = 0
    For k% = 1 To CUEMX%
      If Abs(IMASI#(k%)) >= 0.005 Then
        X$ = REGBLAN$("ASIDIAR")
        Call REPLA(X$, TTXX$, 1, 4)
        Call REPLA(X$, MKI$(k%), 5, 2)
        Call REPLA(X$, MKD$(IMASI#(k%)), 39, 8)
        v& = v& + 1
        Call GRAREG("ASIDIAR", X$, v&)
        '
        If ECOARCH$("CUECON", MKI$(k%)) = "" Then
          Z3$ = REGBLAN$("CUECON")
          Call REPLA(Z3$, MKI$(k%), 1, 2)
          Call REPLA(Z3$, DENOCUE$(k%), 3, 30)
          Call REPLA(Z3$, CODCUE$(k%), 33, 12)
          Call REPLA(Z3$, CODCUE$(k%), 45, 12)
          Call REGAPP("CUECON", Z3$)
          AGRECU% = 1
        End If
      End If
    Next k%
    '
    If AGRECU% > 0 Then
      Call FILESORT("CUECON", "", 1, 1, "ASC")
    End If
    Call CIERRARCH("CUECON")
    '
99  CABATEMP.Close
    Set CABATEMP = Nothing
    Screen.MousePointer = 1
    '
End Sub

'\\\OT-05-0264-WAR-09/05/05///
Sub EXCLUIVA()
 '
    If DERACCE%("SELCOIVA", "Selección de Comprobantes a Listar") < 2 Then
      Exit Sub
    End If
    '
    ' 1.- Ingresar Numero de Proveedor
    '     ****************************
    '
'    FECHALIM% = XVALO(CONTROL$("", "CIEMECOM"))
'    FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    '
10  Call SELECHO("PROVED1")
    NC% = XVALO(VALACT1$("PROVED1"))
    If NC% < 1 Then Exit Sub
    '
    'jandy SQL
  
    Call APERBASDAT("CABAN")
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "AND NuProve = " & NC% & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) >= " & DESDAT & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "Order by CODICOM_LAR ASC "
    '
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NCLIE% = !NUPROVE
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!CODICOM_LAR, 13, 6))
            If Len(!CODICOM_LAR) < 18 Then
              NPX& = XVALO(Right$(!CODICOM_LAR, 6))
            End If
            y$ = MKS$(NPX&) + AJUSTI$(!CODICOM_LAR, 18) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUECO", y$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("!SELCUECO", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUECO") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no Registra\Facturas de Compra que Puedan Seleccionarse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUECO")
    Call FRESHECHO("!SELCUECO")
    '
    Call SELECHO("!SELCUECO")
'    If VALACT1$("!SELCUECO") = "" Then GoTo 10
'    REGCUEC& = XVALO(VALACT1$("!SELCUECO"))
'    If REGCUEC& < 1 Then GoTo 10
    If VALACT2$("!SELCUECO") = "" Then GoTo 10
    REGCUE$ = TRIM$(VALACT2$("!SELCUECO"))
    If REGCUE$ = "" Then GoTo 10
    '
    DOCULA$ = Left$(VALACT2$("!SELCUECO"), 18)
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    '
    Call APERBASDAT("CABAN")
    TTXX$ = "Libro de I.V.A. Compras\  \Comprobante '" + TRIM$(DOCULA$) + "' \de Cuenta '" + TRIM$(rasocli$((-1) * NC%)) + "' ?"
    If COMALTER%(TTXX$ + "\\Incluye\Excluye de Listado") = 1 Then
      COINC = 1
    End If
    
    
'    With CajaBanc
'      .FindFirst "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
'20     If .NoMatch = False Then
'        .Edit
'         If COINC = 1 Then
'           !IncSubIva = ""
'         Else
'           !IncSubIva = "NO"
'         End If
'        .Update
'        .FindNext "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & ""
'        GoTo 20
'      End If
'    End With
    '
    strsql = "Update CajaBanc "
    If COINC = 1 Then
       strsql = strsql & "set IncSubIva='' "
    Else
       strsql = strsql & "set IncSubIva='No' "
    End If
    strsql = strsql & " where CodiCom_Lar='" & DOCULA$ & "' AND NuProve = " & NC%
    Call EJECUTACOMANDO(strsql)

    GoTo 10
    '
End Sub

Sub RANKPRO()
Dim TOTCOMPRA#(32767)
    '
10  FEX = HOY(HO$)
    Has% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
20  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Des% > Has% Then
        GoTo 10
    End If
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    VDEF$ = OBX$
    '
    Screen.MousePointer = 11
    j& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTCOMPRA#(U&) = 0
    Next U&
    '
    'jandy sql
    
    'DESDAT = Int(CDbl(CVDAT(DES%)))
    'HASDAT = Int(CDbl(CVDAT(HAS%)))
    DESDAT = FETEXCOR1$(Des%)
    HASDAT = FETEXCOR1$(Has%)
    
    
    '
    Call APERBASDAT("COMPRAS")
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    
    SQLX$ = "SELECT * FROM OCOMENCA "
    'SQLX$ = SQLX$ + "WHERE INT(CDBL(Femi_Ocom)) >= " & DESDAT & " "
    'SQLX$ = SQLX$ + "AND INT(CDBL(Femi_Ocom)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "WHERE Femi_Ocom >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND Femi_Ocom <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
    '
    'Set CABATEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    PROMAX% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      Do While Not .EOF
         NPRO% = !Npro_Ocom
         If NPRO% > 0 Then
           If NPRO% > PROMAX% Then PROMAX% = NPRO%
           TOTCOMPRA#(NPRO%) = TOTCOMPRA#(NPRO%) + (!TOBRU_OCOM * TICAMONE(!Mone_Emis) / 1)
         End If
         .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    For i% = 1 To PROMAX
      TOTAGEN# = TOTAGEN# + TOTCOMPRA#(i%)
    Next i%
    '
    List1.Clear
    For NCI% = 1 To 32766
       If Abs(TOTCOMPRA#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTCOMPRA#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTCOMPRA#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTCOMPRA#(NCI%)), 4, 12, 2, 2)
          List1.AddItem RELIB$
       End If
    Next NCI%
    '
    j& = 0: PORACU = 0
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        List1.ListIndex = U& - 1
        NC% = XVALO(Mid$(List1.TEXT, 14, 6))
        TONEVEN# = XVALO(Mid$(List1.TEXT, 28, 12))
        PORPAR = XVALO(Mid$(List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKCLI")
        j& = j& + 1: JI% = j&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call GRAREG("RANKCLI", Z$, j&)
        '
    Next U&
    '
    Call SETULTREG("RANKCLI", j&)
    Call CIERRARCH("RANKCLI")
    Call HEADERS("RANKCLI", "")
    Call HEADERS("RANKCLI", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    List1.Clear
    '
    If Option1.Value = True Then
         Call FINAL("*RANKCLI")
       Else
         Call FINAL("*LRANPRO")
    End If
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
End Sub
'\\\OT-05-0381-WAR-FIN///

'\\\OT-05-0264-WAR-FIN///
'\\\OT-05-0381-WAR-07/06/05///
Sub RANKFACO()
Dim TOTVENTA#(32767)
    '
10  FEX = HOY(HO$)
    Has% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
20  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Des% > Has% Then
        GoTo 10
    End If
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    VDEF$ = OBX$
    '
    Screen.MousePointer = 11
    j& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTVENTA#(U&) = 0
    Next U&
    '
    'JANDY SQL
    
    'DESDAT = Int(CDbl(CVDAT(DES%)))
    'HASDAT = Int(CDbl(CVDAT(HAS%)))
    DESDAT = FETEXCOR1$(Des%)
    HASDAT = FETEXCOR1$(Has%)
    
    '
    Call APERBASDAT("CABAN")
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM CAJABANC "
'    SQLX$ = SQLX$ + "WHERE INT(CDBL(FechConta)) >= " & DESDAT & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "WHERE FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'COMPRA' "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
   

    Screen.MousePointer = 11
    CLIMAX% = 0
    NIVCONXXX% = NIVCONFI%
    '
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      Do While Not .EOF
         '\\\OT-08-1032-FL-12/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
         If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70
         '\\\OT-08-1032-FL-FIN
         NCLIE% = !NUPROVE
         If NCLIE% > 0 Then
           If NCLIE% > CLIMAX% Then CLIMAX% = NCLIE%
           TOTVENTA#(NCLIE%) = TOTVENTA#(NCLIE%) + (!Idebecomp - !IHABECOMP)
         End If
70       .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    For i% = 1 To CLIMAX
      TOTAGEN# = TOTAGEN# + TOTVENTA#(i%)
    Next i%
    '
    List1.Clear
    For NCI% = 1 To 32766
       If Abs(TOTVENTA#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTVENTA#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTVENTA#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTVENTA#(NCI%)), 4, 12, 2, 2)
          List1.AddItem RELIB$
       End If
    Next NCI%
    '
    j& = 0: PORACU = 0
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        List1.ListIndex = U& - 1
        NC% = XVALO(Mid$(List1.TEXT, 14, 6))
        TONEVEN# = XVALO(Mid$(List1.TEXT, 28, 12))
        PORPAR = XVALO(Mid$(List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKCLI")
        j& = j& + 1: JI% = j&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call GRAREG("RANKCLI", Z$, j&)
        '
    Next U&
    '
    Call SETULTREG("RANKCLI", j&)
    Call CIERRARCH("RANKCLI")
    Call HEADERS("RANKCLI", "")
    Call HEADERS("RANKCLI", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    List1.Clear
    '
    If Option1.Value = True Then
         Call FINAL("*RANKCLI")
       Else
         Call FINAL("*LRANPRO")
    End If
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
End Sub
'\\\OT-05-0381-WAR-FIN///
Sub LIESFACO()
'\\\OT-06-0180-WAR-12/05/06///
    Dim NPROV1%(), PERI1$(), TOTIMP#()
    '
'1.- RANGO DE FECHAS
    Call DESHASFECHA(Des%, Has%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    '
    Screen.MousePointer = 11
    '
    'JANDY SQL
    'DESDAT = Int(CDbl(CVDAT(DES%)))
    'HASDAT = Int(CDbl(CVDAT(HAS%)))
    DESDAT = FETEXCOR1$(Des%)
    HASDAT = FETEXCOR1$(Has%)
    
    ' JANDY SQL (NO HABRIA Q CALCULAR EL SALDO EN LA SELECT DIRECTAMENTE
    ' CON UN GROUP BY POR PROVEEDOR?)

'2.- CONSULTA SQL
    Call APERBASDAT("CABAN")
     Dim LIFACOTMP As ADODB.Recordset
     Set LIFACOTMP = New ADODB.Recordset
     LIFACOTMP.CursorLocation = adUseClient
    
    SQLX$ = "SELECT * FROM CAJABANC"
    SQLX$ = SQLX$ + " WHERE TipoMovim = 'COMPRA' "
    'SQLX$ = SQLX$ + " AND   CDBL(FechConta) >= " & DESDAT & ""
    'SQLX$ = SQLX$ + " AND   CDBL(FechConta) <= " & HASDAT & ""
    SQLX$ = SQLX$ + " AND   FechConta >= '" & DESDAT & "'"
    SQLX$ = SQLX$ + " AND   FechConta <= '" & HASDAT & "'"
    'Set LIFACOTMP = CueCorri.OpenRecordset(SQLX$, 4)
25  On Error Resume Next
    LIFACOTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
      NIVCONXXX% = NIVCONFI%

    With LIFACOTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Movimientos para el \Rango de Fechas Elegido.")
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      '
      CAREG1% = 0
      Do While Not .EOF
        '\\\OT-08-1032-FL-12/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
        If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 30
        '\\\OT-08-1032-FL-FIN
        NPX1% = !NUPROVE
        FEX = CVINT(!fechconta)
        FE1$ = FECHATEX$(FEX)
        FE2$ = "20" + Mid$(FE1$, 7, 2) + "/" + Mid$(FE1$, 4, 2)
        IMPOX1# = !Idebecomp - !IHABECOMP 'ValAbsComp
        '
        For KKI% = 1 To CAREG1%
          If PERI1$(KKI%) = FE2$ Then
            If NPROV1%(KKI%) = NPX1% Then
              TOTIMP#(KKI%) = TOTIMP#(KKI%) + IMPOX1#
              GoTo 30
            End If
          End If
        Next KKI%
        '
        CAREG1% = CAREG1% + 1
        ReDim Preserve PERI1$(CAREG1%), NPROV1%(CAREG1%), TOTIMP#(CAREG1%)
        PERI1$(CAREG1%) = FE2$
        NPROV1%(CAREG1%) = NPX1%
        TOTIMP#(CAREG1%) = IMPOX1#
        '
30    .MoveNext
      Loop
    End With
    LIFACOTMP.Close
    Set LIFACOTMP = Nothing
    '
    Screen.MousePointer = 1
    '4.-AQUI VA A GENERAR ARCHIVO
    y$ = REGBLAN$("LIESFACO")
    JJ& = 0
    For U% = 1 To CAREG1%
      Call REPLA(y$, PERI1$(U%), 1, 8)
      Call REPLA(y$, MKI$(NPROV1%(U%)), 9, 2)
      Call REPLA(y$, MKD$(TOTIMP#(U%)), 11, 8)
      JJ& = JJ& + 1
      Call GRAREG("LIESFACO", y$, JJ&)
    Next U%
    '
    Call SETULTREG("LIESFACO", JJ&)
    '5.- SORT
    Call FILESORT("LIESFACO", "LIESFACO", 4, 4)
    Call FILESORT("LIESFACO", "LIESFACO", 2, 2)
    Call FILESORT("LIESFACO", "LIESFACO", 1, 1)
    '
    Call HEADERS("LIESFACO", "")
    Call HEADERS("LIESFACO", "Período: " + PERIO$)
    '
    '6.- LIST
    Call FINAL("*LIESFACO")
'\\\OT-06-0180-WAR-FIN///
End Sub
Sub RANKART()
  '
  Dim TOTART#(32767)
    '
10  FEX = HOY(HO$)
    Has% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
20  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Des% > Has% Then
        GoTo 10
    End If
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    VDEF$ = OBX$
    '
    Screen.MousePointer = 11
    j& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTART#(U&) = 0
    Next U&
    '
    'JANDY SQL
    'DESDAT = Int(CDbl(CVDAT(DES%)))
    'HASDAT = Int(CDbl(CVDAT(HAS%)))
    DESDAT = FETEXCOR1$(Des%)
    HASDAT = FETEXCOR1$(Has%)
    '
    Call APERBASDAT("COMPRAS")
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM OCOMDETA "
'    SQLX$ = SQLX$ + "WHERE INT(CDBL(Femi_Ocom)) >= " & DESDAT & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(Femi_Ocom)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "WHERE Femi_Ocom >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND Femi_Ocom <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
    '
    'Set CABATEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    PROMAX% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      Do While Not .EOF
         NART% = XVALO(!COD_INTE)
         If NART% > 0 Then
           If NART% > PROMAX% Then PROMAX% = NART%
           TOTART#(NART%) = TOTART#(NART%) + (XVALO(!inet_item) * TICAMONE(XVALO(!Mone_Emis)) / 1)
         End If
         .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    For i% = 1 To PROMAX
      TOTAGEN# = TOTAGEN# + TOTART#(i%)
    Next i%
    '
    List1.Clear
    For NCI% = 1 To 32766
       If Abs(TOTART#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTART#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTART#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTART#(NCI%)), 4, 12, 2, 2)
          List1.AddItem RELIB$
       End If
    Next NCI%
    '
    j& = 0: PORACU = 0
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        List1.ListIndex = U& - 1
        NC% = XVALO(Mid$(List1.TEXT, 14, 6))
        TONEVEN# = XVALO(Mid$(List1.TEXT, 28, 12))
        PORPAR = XVALO(Mid$(List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKPIE")
        j& = j& + 1: JI% = j&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call GRAREG("RANKPIE", Z$, j&)
        '
    Next U&
    '
    Call SETULTREG("RANKPIE", j&)
    Call CIERRARCH("RANKPIE")
    Call HEADERS("RANKPIE", "")
    Call HEADERS("RANKPIE", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    List1.Clear
    '
    If Option1.Value = True Then
         Call FINAL("*RANKCLI")
       Else
         Call FINAL("*LRANART")
    End If
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
End Sub
'\\\OT-05-0559-WAR-02/08/05///
Sub GENIVACOM()
    Screen.MousePointer = 11
    '
    Dim RELIVA$(32)
    Dim IMASI#(32767)
    CUEMX% = 0
    Call SETULTREG("LIVACOM", 0)
    Call SETULTREG("LIVACCAI", 0)
    '
    ' Leer las Cuentas Contables de IVA
    Dim CUEIVA%(), TASIVANOM$(), ICUEN#(), CACUEIVA%
    Dim NUMEASI$(), FECHASI%(), CANASIX%
    CACUEIVA% = 0
    CANASIX% = 0
    '
    For IXX% = 0 To 6
      TXXX$ = TRIM$(Str$(IXX%))
      CIVX$ = TRIM$(CONTROL$("", "CUEIVA-" + TXXX$))
      If CIVX$ <> "" Then
        CIVY% = CUEINT%(CIVX$)
        If CIVY% > 0 Then
          TASAX$ = CONTROL$("", "TASIVA-" + TXXX$)
          If XVALO(TASAX$) > 0 Or IXX% > 3 Then
            If IXX% = 4 Then TASAX$ = "Aduana"
            If IXX% = 5 Then TASAX$ = "Per-IVA"
            If IXX% = 6 Then TASAX$ = "Per-GAN"
            CACUEIVA% = CACUEIVA% + 1
            ReDim Preserve CUEIVA%(CACUEIVA%), TASIVANOM$(CACUEIVA%), ICUEN#(CACUEIVA%)
            CUEIVA%(CACUEIVA%) = CIVY%
            TASIVANOM$(CACUEIVA%) = TASAX$
          End If
        End If
      End If
    Next IXX%
    '
    ' SUPRIME ASIENTOS DE DIARIO
    If CONTROL$("LIVACOM", "SUPRIVA") = "SI" Then GoTo 22
    '
    SQLX$ = "SELECT * FROM DetaAsien WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE Fechasien >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + " AND Fechasien <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + " AND IDEBEPASE >= 0.005 "
    SQLX$ = SQLX$ + " AND statpase > 0  "
    SQLX$ = SQLX$ + " ORDER BY Fechasien "
    Set DETACONTA = READSET(SQLX$)
    '
    With DETACONTA
      Do While Not .EOF
        NUMEFA$ = AJUSTI$(!nuidasien, 16)
        For KUI% = 1 To CANASIX%
          If NUMEASI$(KUI%) = NUMEFA$ Then GoTo 12
        Next KUI%
        icuexi% = !CUECOINT
        For IXX% = 1 To CACUEIVA%
          If icuexi% = CUEIVA%(IXX%) Then
            CANASIX% = CANASIX% + 1
            ReDim Preserve NUMEASI$(CANASIX%), FECHASI%(CANASIX%)
            NUMEASI$(CANASIX%) = NUMEFA$
            FECHASI%(CANASIX%) = CVINT%(!FECHASIEN)
            GoTo 12
          End If
        Next IXX%
12    .MoveNext
      Loop
    End With
    '
19  CANASIP% = 0
    For IXX% = 1 To CANASIX%
      NUMEFA$ = TRIM$(NUMEASI$(IXX%))
      If CantRegistros("ENCABEASIEN", "NuidAsien = '" & NUMEFA$ & "' AND StatAsien > 0 AND TipoAsien < 1") > 0 Then
          CANASIP% = CANASIP% + 1
          NUMEASI$(CANASIP%) = NUMEASI$(IXX%)
          FECHASI%(CANASIP%) = FECHASI%(IXX%)
      End If
    Next IXX%
    CANASIX% = CANASIP%

    On Error Resume Next
    DETACONTA.Close
    Set DETACONTA = Nothing
    On Error GoTo 0
    
22  Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta<= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'FCOM' "
    '\\\OT-05-0264-WAR-09/05/05///
    SQLX$ = SQLX$ + "AND (IncSubIva <> 'NO' "
    SQLX$ = SQLX$ + "OR IncSubIva is null OR IncSubIva = '') "
    '\\\OT-05-0264-WAR-FIN///
    '\\\OT-07-0063-RG-15/02/07///
    SQLX$ = SQLX$ + "Order by FechConta ASC, FechAcred ASC, FeHoReal ASC "
    '\\\OT-07-0063-RG-FIN///
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    JJ& = 0
    REBLA$ = REGBLAN$("LIVACCAI")
    REBLB$ = REGBLAN$("REIVACOM")
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      FIN& = .RecordCount
      .MoveFirst
      On Error GoTo 0
      U& = 0
      NIVCONXXX% = NIVCONFI%
      Do While Not .EOF
        U& = U& + 1
        Call TRACE(20, U&, FIN&)
        '\\\OT-08-1032-OD-11/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
        If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70 'VER QUE FACTURAS MOSTRAR EN LISTADO
        '\\\OT-08-1032-OD-FIN

        'ver ernesto
        'FECHAD = Int(CDbl(!FechConta))
          'GoSub INCLUASDIAR
        'jandy sql
        FECHAD = FETEXCOR1$(CVINT%(!fechconta))
        NUMEFA$ = SAFETEXT$(!CODICOM_LAR)
        NCLIE% = XVALO(!NUPROVE)
        ITOTA# = XVALO(!ValAbsComp)
        SIG% = Sgn(XVALO(!IHABECOMP) - XVALO(!Idebecomp))
        MOBASI$ = TRIM$(SAFETEXT$(!TipoMovim))
        OBSERVAFAC$ = TRIM$(Mid$(SAFETEXT$(!Notas), 1, 18))
        '
        ZZ$ = REBLA$
        Call REPLA(ZZ$, MKI$(CVINT(!fechconta)), 3, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!fechacred)), 5, 2)
        Call REPLA(ZZ$, MKS$(XVALO(!NumeComp)), 7, 4)
        Call REPLA(ZZ$, SAFETEXT$(!CODICOM_LAR), 11, 18)
        Call REPLA(ZZ$, MKI$(NCLIE%), 31, 2)
        Call REPLA(ZZ$, rasocli$((-1) * NCLIE%), 33, 30)
            Cuit$ = CuitCli$((-1) * NCLIE%)
            CuitEnSql$ = SAFETEXT$(!Cuit)
            If Len(TRIM$(SAFETEXT$(!Cuit))) > 10 And SAFETEXT$(!Cuit) <> "00-00000000-0" And VALICUIT%(CuitEnSql$) = 1 Then
                Cuit$ = SOLONUM$(CuitEnSql$)
                Cuit$ = Mid$(Cuit$, 1, 2) & "-" & Mid$(Cuit$, 3, 8) & "-" & Mid$(Cuit$, 11, 1)
            End If
        Call REPLA(ZZ$, Cuit$, 63, 16)
        Call REPLA(ZZ$, POSIVAC$((-1) * NCLIE%), 79, 12)
        '\\\OT-05-0264-WAR-09/05/05///
        If IsNull(!NROCAI) Then
           NCAI$ = ""
          Else
           NCAI$ = SAFETEXT$(!NROCAI)
        End If
        If IsNull(!FEVTOCAI) Then
            FVCAI = 0
          Else
            FVCAI = !FEVTOCAI
        End If
        Call REPLA(ZZ$, NCAI$, 257, 24)
        Call REPLA(ZZ$, MKI$(CVINT(FVCAI)), 281, 2)
        '\\\OT-05-0264-WAR-FIN///
          ' LO QUE SIGUE ES PARA CONSIDERAR CLIENTES VARIOS
          'If UCase$(Mid$(ZZ$, 33, 6)) = "VARIOS" Then
          If InStr(1, UCase$(Mid$(ZZ$, 33, 30)), "VARIOS") > 0 Then
            REFEMO$ = SAFETEXT$(!RefeCom)
            PPXX% = InStr(REFEMO$, "//")
            If PPXX% > 0 Then
              RSCL$ = TRIM$(Left$(REFEMO$, PPXX% - 1))
              CTCL$ = TRIM$(Mid$(REFEMO$, PPXX% + 2))
              If RSCL$ <> "" Then
                If CTCL$ <> "" Then
                  Call REPLA(ZZ$, RSCL$, 33, 30)
                  Call REPLA(ZZ$, CTCL$, 63, 16)
                End If
              End If
            End If
          End If
          '
        Call REPLA(ZZ$, MKD$(ITOTA# * SIG%), 91, 8)
        '\\\OT-05-0709-WAR-02/11/05///
        Pi% = PIVACLI%(NCLIE% * (-1))
        If Pi% = 5 Then
          Call REPLA(ZZ$, MKD$(ITOTA# * SIG%), 185, 8)
        End If
        '\\\OT-05-0709-WAR-FIN///
        RESTOBRU# = ITOTA# * SIG%
        '
        '
        RELIVA$(0) = ZZ$
        For KU% = 1 To 32
          RELIVA$(KU%) = REBLA$
        Next KU%
        '
        'VER ACA
        Dim CABAFETMP As ADODB.Recordset
        Set CABAFETMP = New ADODB.Recordset
        CABAFETMP.CursorLocation = adUseClient
        SQLX$ = "SELECT * FROM CAJABANC "
        'SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
        'SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
        SQLX$ = SQLX$ + "WHERE FechConta = '" & FECHAD & "' "
        SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX$ = SQLX$ + "AND NuProve = " & NCLIE% & " "
        SQLX$ = SQLX$ + "AND TipoMovim <> '" & MOBASI$ & "'"
24      On Error Resume Next
        CABAFETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 24
        End If
        On Error GoTo 0
        '
        ' VERIFICA SI ALGUNO DE LOS ITEMS DEL DETALLE DEL COMPROBANTE TIENE ALGUNA DE LAS CUENTAS CONTABLES DE IVA
        ' DE SER ASI, LA CONSIDERA COMO FACTURA DE IMPORTACION CON SOLO IVA Y DICHO IMPORTE LO PONE EN LA COLUMNA
        ' DE IVA EN LUGAR DE LA DE EXENTO
        CONDI$ = "CodiCom_Lar = '" & NUMEFA$ & "' AND NuProve = " & NCLIE% & " AND TipoMovim = 'COMPRA' AND ("
        CONDI1$ = ""
        For II% = 1 To CACUEIVA%
            CONDI1$ = CONDI1$ & "CUECONTAI=" & CUEIVA%(II%) & " OR "
        Next II%
        FACIMPO% = 0
        If TRIM$(CONDI1$) <> "" And InStr(1, EMPREAC$, "FUINYTER") < 1 Then
            CONDI1$ = Mid$(TRIM$(CONDI1$), 1, Len(CONDI1$) - 3)
            CONDI$ = CONDI$ & CONDI1$ & ")"
            FACIMPO% = CantRegistros("CAJABANC", CONDI$)
        End If
        '
        PROIVA% = 0: ULIVA% = 0: PROIBRU% = 0: ULIBRU% = 0
        With CABAFETMP
        If Not (.EOF And .BOF) Then
        .MoveFirst
        Do While Not .EOF
            TIMO$ = UCase$(TRIM$(SAFETEXT$(!TipoMovim)))
            OPMO$ = SAFETEXT$(!OpciMovim)
            VAMO$ = SAFETEXT$(!VARIMOVIM)
            IPAX# = Abs(XVALO(!ValAbsComp))
            SIGMO% = Sgn(XVALO(!Idebecomp) - XVALO(!IHABECOMP))
            Base# = 0
            On Error Resume Next
            Base# = Abs(XVALO(!BASEIMPO))
            On Error GoTo 0
            '
            If Abs(IPAX#) >= 0.005 Or Base# >= 0.005 Then
              If TIMO$ = "IVA" And UCase$(VAMO$) = "PER-IVA" Then
                  ZZ$ = RELIVA$(0)
                  Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 161, 8)
                  RELIVA$(0) = ZZ$
                  RESTOBRU# = RESTOBRU# - IPAX# * SIG%
                ElseIf (TIMO$ = "GAN" Or TIMO$ = "IVA") And UCase$(VAMO$) = "PER-GAN" Then
                  ZZ$ = RELIVA$(0)
                  Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 169, 8)
                  RELIVA$(0) = ZZ$
                  RESTOBRU# = RESTOBRU# - IPAX# * SIG%
                ElseIf TIMO$ = "IBRU" And (UCase$(VAMO$) = "PERC" Or UCase$(VAMO$) = "RETEN") Then
                  ZZ$ = RELIVA$(0)
                  ACUANTEX# = CVD(Mid$(ZZ$, 193, 8))
                  Call REPLA(ZZ$, MKD$(ACUANTEX# + IPAX# * SIG%), 193, 8)
                  RESTOBRU# = RESTOBRU# - IPAX# * SIG%
                  Call REPLA(ZZ$, MKI$(!CUECONTAI), 121, 2)
                    JURIB$ = TRIM$(Mid$(OPMO$, 4))
                  Call REPLA(ZZ$, JURIB$, 203, 1)
                  Call REPLA(ZZ$, ECOARCH$("PROVINC", JURIB$), 205, 24)
                  If Abs(ACUANTEX#) >= 0.005 Then
                    Call REPLA(ZZ$, " ", 203, 1)
                    Call REPLA(ZZ$, "JURISD. VARIAS", 205, 24)
                  End If
                  RELIVA$(0) = ZZ$
                
'                ElseIf TIMO$ = "IBRU" And UCase$(VAMO$) = "RETEN" Then
'                  ZZ$ = RELIVA$(0)
'                  ACUANTEX# = CVD(Mid$(ZZ$, 193, 8))
'                  Call REPLA(ZZ$, MKD$(ACUANTEX# + IPAX# * SIG%), 193, 8)
'                  RESTOBRU# = RESTOBRU# - IPAX# * SIG%
'                  PROIVA% = PROIVA% + 1
'                  RELIVA$(1) = ZZ$
'                  HUBORETIIBB% = 1
                  
                ElseIf TIMO$ = "IVA" And Abs(IPAX#) >= 0.005 Then
'                  If HUBORETIIBB% = 1 Then
'                    HUBORETIIBB% = -1
'                    PROIVA% = 0
'                  End If
                  ZZ$ = RELIVA$(PROIVA%)
                  Call REPLA(ZZ$, VAMO$, 129, 8)
                  Call REPLA(ZZ$, MKD$(Base# * SIGMO%), 137, 8)
                  Call REPLA(ZZ$, MKD$(IPAX# * SIGMO%), 145, 8)
                  RESTOBRU# = RESTOBRU# - IPAX# * SIGMO% - Base# * SIGMO%
                  'If BASE# > 0.005 Then
                  '  Call REPLA(ZZ$, MKD$(100 * IPAX# / BASE#), 153, 8)
                  'End If
                  '
                  ' CONSIDERA QUE EL IMPORTE EXENTO CORRESPONDE A UNA FACTURA DE IMPORTACION
                  If FACIMPO% >= 1 Then
                    Call REPLA(ZZ$, MKD$(0), 137, 8)
                    Call REPLA(ZZ$, MKD$(Base# * SIGMO%), 145, 8)
                    Call REPLA(ZZ$, OBSERVAFAC$, 11, 18)
                    Call REPLA(ZZ$, Space(8), 129, 8)
                  End If
                  '
                  RELIVA$(PROIVA%) = ZZ$
                  ULIVA% = PROIVA%
                  PROIVA% = PROIVA% + 1
              End If
            End If
            '
            .MoveNext
            Loop
            '
          End If
        End With
        CABAFETMP.Close
        Set CABAFETMP = Nothing
        '
        'If HUBORETIIBB% = -1 And PROIVA% = 0 Then ULIVA% = 1
        '
        If Abs(RESTOBRU#) > 0.005 Then
          For KU% = 0 To ULIVA%
            ZZ$ = RELIVA$(KU%)
            If TRIM$(Mid$(ZZ$, 129, 8)) = "" Then
              If Abs(CVD(Mid$(ZZ$, 137, 8))) < 0.005 Then
                If Abs(CVD(Mid$(ZZ$, 145, 8))) < 0.005 Then
                  Call REPLA(ZZ$, "No Grav.", 129, 8)
                  Call REPLA(ZZ$, MKD$(RESTOBRU#), 137, 8)
                  RELIVA$(KU%) = ZZ$
                  GoTo 50
                End If
              End If
            End If
          Next KU%
          ULIVA% = ULIVA% + 1
          'ZZ$ = RELIVA$(ULIVA%)
          ZZ$ = REBLA$: Call REPLA(ZZ$, RELIVA$(0), 1, 91)
          VAMO$ = "No Grav."
          Call REPLA(ZZ$, VAMO$, 129, 8)
          Call REPLA(ZZ$, MKD$(RESTOBRU#), 137, 8)
          RELIVA$(ULIVA%) = ZZ$
        End If
        '
50      If ULIBRU% > ULIVA% Then ULIVA% = ULIBRU%
        For KU% = 0 To ULIVA%
          JJ& = JJ& + 1
          Call GRAREG("LIVACOM", RELIVA$(KU%), JJ&)
          Call GRAREG("LIVACCAI", RELIVA$(KU%), JJ&)
        Next KU%
        '
70    .MoveNext
      Loop
    End With

    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    FECHAD = CDbl(CDate("31/12/2050"))
    GoSub INCLUASDIAR
    GoSub INCLUYETARJETAS
    '
    Call SETULTREG("LIVACOM", JJ&)
    Call SETULTREG("LIVACCAI", JJ&)
    '
    Call HEADERS("LIVACOM", "")
    Call HEADERS("LIVACOM", "Corresponde a: " + TRIM$(CORRESP$))
    '
    '
    ' CARGA RESUMEN MENSUAL DE IVA - COMPRAS
    Dim DECATIV$(128)
    CACATIV% = 0
    Dim VECREIVA$(256)
    UVECRE% = 0
    '
    For KUL& = 1 To JJ&
      RECUE$ = REGLEIDO$("LIVACOM", KUL&)
      If TRIM$(Mid$(RECUE$, 11, 18)) <> "" Then
        NUCO$ = TRIM$(Mid$(RECUE$, 11, 18))
        SIG% = Sgn(CVD(Mid$(RECUE$, 91, 8)))
      End If
      '
      CORUBRO% = 1: DERUBRO$ = "COMPRA"
      COCONCE% = 8: DECONCE$ = "COMPROBANTES VARIOS"
      '
      If SIG% < 0 Then
          COCONCE% = 9: DECONCE$ = "NOTAS DE CREDITO"
        Else
          If Left$(NUCO$, 2) = "FC" Then COCONCE% = 1: DECONCE$ = "FACTURAS"
          If Left$(NUCO$, 2) = "ND" Then COCONCE% = 2: DECONCE$ = "NOTAS DE DEBITO"
          If Left$(NUCO$, 2) = "TF" Then COCONCE% = 3: DECONCE$ = "TICKET-FACTURA"
          If Left$(NUCO$, 2) = "RH" Then COCONCE% = 4: DECONCE$ = "RECIBO HONORARIOS"
      End If
      '
      CAIVA$ = Mid$(RECUE$, 79, 12)
      If TRIM$(CAIVA$) = "" Then CAIVA$ = CAIVANT$
      DECATE$ = TRIM$(Mid$(RECUE$, 129, 8)) + " - " + CAIVA$
      CAIVANT$ = CAIVA$
      For KLI% = 1 To CACATIV%
        If DECATE$ = DECATIV$(KLI%) Then
          COCATE% = KLI%
          GoTo 20
        End If
      Next KLI%
      CACATIV% = CACATIV% + 1
      DECATIV$(CACATIV%) = DECATE$
      COCATE% = CACATIV%
      '
20    BASEIMPO# = CVD(Mid$(RECUE$, 137, 8))
      VALOIMPU# = CVD(Mid$(RECUE$, 145, 8))
      
82    If (Abs(BASEIMPO#) + Abs(VALOIMPU#)) >= 0.005 Then
        ABUS$ = Chr$(CORUBRO%) + AJUSTI$(DERUBRO$, 15)
        ABUS$ = ABUS$ + Chr$(COCONCE%) + AJUSTI$(DECONCE$, 23)
        ABUS$ = ABUS$ + Chr$(COCATE%) + AJUSTI$(DECATE$, 23)
        '
        
        For UVE% = 1 To UVECRE%
          If Left$(VECREIVA$(UVE%), 64) = ABUS$ Then
             VCC$ = VECREIVA$(UVE%)
             GoTo 84
          End If
        Next UVE%
        UVECRE% = UVECRE% + 1
        UVE% = UVECRE%
        VCC$ = REBLB$
        Call REPLA(VCC$, ABUS$, 1, 64)
        '
84      BASEIMPO# = BASEIMPO# + CVD(Mid$(VCC$, 65, 8))
        VALOIMPU# = VALOIMPU# + CVD(Mid$(VCC$, 81, 8))
        Call REPLA(VCC$, MKD$(BASEIMPO#), 65, 8)
        Call REPLA(VCC$, MKD$(VALOIMPU#), 81, 8)
        Call REPLA(VCC$, MKD$(BASEIMPO# + VALOIMPU#), 97, 8)
        VECREIVA$(UVE%) = VCC$
      End If
      '
      CORUBRO% = 2: DERUBRO$ = "PERCEPCION"        ' PERCEPCIONES
      For COCO% = 1 To 3
        COCONCE% = 10 + COCO%
        BASEIMPO# = 0
        COCATE% = 0: DECATE$ = ""
        VALOIMPU# = CVD(Mid$(RECUE$, 153 + 8 * COCO%, 8))
        If COCO% = 1 Then
           DECONCE$ = "PERCEPCION I.V.A."
          ElseIf COCO% = 2 Then
           DECONCE$ = "PERCEPCION GANANCIAS"
          ElseIf COCO% = 3 Then
           DECONCE$ = "PERCEPC.ING.BRUTOS"
           DECATE$ = Mid$(RECUE$, 205, 24)
           VALOIMPU# = CVD(Mid$(RECUE$, 193, 8))
        End If
        '
92      If (Abs(BASEIMPO#) + Abs(VALOIMPU#)) >= 0.005 Then
          ABUS$ = Chr$(CORUBRO%) + AJUSTI$(DERUBRO$, 15)
          ABUS$ = ABUS$ + Chr$(COCONCE%) + AJUSTI$(DECONCE$, 23)
          ABUS$ = ABUS$ + Chr$(COCATE%) + AJUSTI$(DECATE$, 23)
          '
          For UVE% = 1 To UVECRE%
            If Left$(VECREIVA$(UVE%), 64) = ABUS$ Then
               VCC$ = VECREIVA$(UVE%)
               GoTo 94
            End If
          Next UVE%
          UVECRE% = UVECRE% + 1
          UVE% = UVECRE%
          VCC$ = REBLB$
          Call REPLA(VCC$, ABUS$, 1, 64)
          '
94        BASEIMPO# = BASEIMPO# + CVD(Mid$(VCC$, 65, 8))
          VALOIMPU# = VALOIMPU# + CVD(Mid$(VCC$, 81, 8))
          Call REPLA(VCC$, MKD$(BASEIMPO#), 65, 8)
          Call REPLA(VCC$, MKD$(VALOIMPU#), 81, 8)
          Call REPLA(VCC$, MKD$(BASEIMPO# + VALOIMPU#), 97, 8)
          VECREIVA$(UVE%) = VCC$
        End If
         '
      Next COCO%
      '
    Next KUL&
    '
    JJK& = 0
    For UI1% = 1 To UVECRE%
      JJK& = JJK& + 1
      Call GRAREG("REIVACOM", VECREIVA$(UI1%), JJK&)
    Next UI1%
    Call SETULTREG("REIVACOM", JJK&)
    Call CIERRARCH("REIVACOM")
    '
    Call FILESORT("REIVACOM", "", 6, 6, "ASC")
    Call FILESORT("REIVACOM", "", 3, 3, "ASC")
    Call FILESORT("REIVACOM", "", 1, 1, "ASC")
    '
    Call HEADERS("REIVACOM", "")
    Call HEADERS("REIVACOM", "Corresponde a: " + TRIM$(CORRESP$))
    '
99  Screen.MousePointer = 1
    Exit Sub
    '
    '
'ver ernesto contabilidad

INCLUASDIAR:
    FEMOSUDIA% = CVINT(FECHAD)
    For KUI% = 1 To CANASIX%
      If FECHASI%(KUI%) > 0 Then
        If FECHASI%(KUI%) < FEMOSUDIA% Then
          NUMEFA$ = TRIM$(NUMEASI$(KUI%))
          NCLIE% = 0
          For IXX% = 1 To CACUEIVA%
            ICUEN#(IXX%) = 0
          Next IXX%
          '
          SQLX$ = "SELECT * FROM DetaAsien "
          SQLX$ = SQLX$ + " WHERE Fechasien >= '" & DESDAT & "' "
          SQLX$ = SQLX$ + " AND Fechasien <= '" & HASDAT & "' "
          SQLX$ = SQLX$ + " AND statpase > 0  "
          SQLX$ = SQLX$ + " AND NuIdAsien = '" & NUMEFA$ & "' "
          SQLX$ = SQLX$ + "ORDER BY OrNuPase "
          'Set DETACONTA = Contable.OpenRecordset(SQLX$, 4)
          Dim DETACONTA0 As ADODB.Recordset
          Set DETACONTA0 = New ADODB.Recordset
28        On Error Resume Next
          DETACONTA0.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 28
          End If
          On Error GoTo 0
           '
          With DETACONTA0
            On Error Resume Next
            .MoveFirst
            If Err < 1 Then
              On Error GoTo 0
              Do While Not .EOF
                icuexi% = XVALO(!CUECOINT)
                For IXX% = 1 To CACUEIVA%
                  If icuexi% = CUEIVA%(IXX%) Then
                    ICUEN#(IXX%) = ICUEN#(IXX%) + XVALO(!idebepase) - XVALO(!ihabepase)
                    REPAS$ = SAFETEXT$(!RefePase)
                    GoTo 192
                  End If
                Next IXX%
                If NCLIE% < 1 Then
                  Call APERBASDAT("PROVE")
'                  Proved12.FindFirst "CuMa_Cli = " & icuexi%
'                  If Proved12.NoMatch = False Then
                   SQLX$ = "SELECT * FROM Proved12 "
                   SQLX$ = SQLX$ + " WHERE CuMa_Cli = " & icuexi% & ""
                   Dim PROVVVED12 As ADODB.Recordset
                   Set PROVVVED12 = FLEXCONN.Execute(FILTSQL$(SQLX$))
                   With PROVVVED12
                    On Error Resume Next
                    If Not (.EOF And .BOF) Then
                      On Error GoTo 0
                      NCLIE% = XVALO(!nume_cli)
                    End If
                    On Error GoTo 0
                   End With
                End If
192           .MoveNext
              Loop
            End If
          End With
          On Error Resume Next
          PROVVVED12.Close
          Set PROVVVED12 = Nothing
          DETACONTA0.Close
          Set DETACONTA0 = Nothing
          On Error GoTo 0
          '
          ZZ$ = REBLA$
          Call REPLA(ZZ$, MKI$(FECHASI%(KUI%)), 3, 2)
          Call REPLA(ZZ$, MKI$(FECHASI(KUI%)), 5, 2)
          Call REPLA(ZZ$, MKS$(0), 7, 4)
          Call REPLA(ZZ$, NUMEFA$, 11, 18)
          Call REPLA(ZZ$, MKI$(NCLIE%), 31, 2)
          Call REPLA(ZZ$, rasocli$((-1) * NCLIE%), 33, 30)
          Call REPLA(ZZ$, CuitCli$((-1) * NCLIE%), 63, 16)
          Call REPLA(ZZ$, POSIVAC$((-1) * NCLIE%), 79, 12)
          '
          If Abs(NCLIE%) < 1 Then
            Call REPLA(ZZ$, REPAS$, 33, 30)
            Call REPLA(ZZ$, "N/Registrado", 63, 16)
            Call REPLA(ZZ$, "N/Def", 79, 12)
          End If
          '
          PROIVA% = 0: ULIVA% = 0
          RELIVA$(0) = ZZ$
          For KU% = 1 To 32
            RELIVA$(KU%) = REBLA$
          Next KU%
          '
          ITOTAX# = 0
          For IXX% = 1 To CACUEIVA%
            If Abs(ICUEN#(IXX%)) >= 0.005 Then
              ZZ$ = RELIVA$(PROIVA%)
              VAMO$ = TASIVANOM$(IXX%)
              If TRIM$(UCase$(VAMO$)) = "PER-IVA" Then
                   ZZ$ = RELIVA$(0)
                   ITOTAA# = ICUEN#(IXX%) + CVD(Mid$(ZZ$, 161, 8))
                   Call REPLA(ZZ$, MKD$(ITOTAA#), 161, 8)
                   ITOTAX# = ITOTAX# + ICUEN#(IXX%)
                   RELIVA$(0) = ZZ$
                   GoTo 199
                 ElseIf TRIM$(UCase$(VAMO$)) = "PER-GAN" Then
                   ZZ$ = RELIVA$(0)
                   ITOTAA# = ICUEN#(IXX%) + CVD(Mid$(ZZ$, 169, 8))
                   Call REPLA(ZZ$, MKD$(ITOTAA#), 169, 8)
                   ITOTAX# = ITOTAX# + ICUEN#(IXX%)
                   RELIVA$(0) = ZZ$
                   GoTo 199
              End If
              Base# = 0
              If XVALO(TASIVANOM$(IXX%)) > 0.5 Then
                Base# = 100 * ICUEN#(IXX%) / XVALO(TASIVANOM$(IXX%))
              End If
              Call REPLA(ZZ$, VAMO$, 129, 8)
              Call REPLA(ZZ$, MKD$(Base#), 137, 8)
              Call REPLA(ZZ$, MKD$(ICUEN#(IXX%)), 145, 8)
              ITOTAX# = ITOTAX# + Base# + ICUEN#(IXX%)
              If XVALO(TASIVANOM$(IXX%)) > 0.5 Then
                Call REPLA(ZZ$, MKD$(XVALO(TASIVANOM$(IXX%))), 153, 8)
              End If
              RELIVA$(PROIVA%) = ZZ$
              ULIVA% = PROIVA%
              PROIVA% = PROIVA% + 1
              '
            End If
199       Next IXX%
          '
          FECHASI%(KUI%) = 0
          ZZ$ = RELIVA$(0)
          Call REPLA(ZZ$, MKD$(ITOTAX#), 91, 8)
          RELIVA$(0) = ZZ$
          '
          For KU% = 0 To ULIVA%
            JJ& = JJ& + 1
            Call GRAREG("LIVACOM", RELIVA$(KU%), JJ&)
            Call GRAREG("LIVACCAI", RELIVA$(KU%), JJ&)
          Next KU%
          '
        End If
      End If
    Next KUI%
    
Return

INCLUYETARJETAS:
    
    ConsultaSql$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
    ConsultaSql$ = ConsultaSql$ & " WHERE FECHCONTA >= '" & DESDAT & "' AND FECHCONTA <= '" & HASDAT & "'"
    ConsultaSql$ = ConsultaSql$ & " AND TIPOMOVIM='CONCITC' AND IDEBECOMP > 0"
    Dim RstIvaTarjetas As ADODB.Recordset
    Set RstIvaTarjetas = READSET(ConsultaSql$)
    
    With RstIvaTarjetas
        Do While Not .EOF
            NroDocumento$ = SAFETEXT$(!CODICOM_LAR)
            CuentaContableBanco% = XVALO(!CUECONTAI)
            FechaMovimiento% = CVINT(!fechconta)
            For IXX% = 1 To CACUEIVA%
                ICUEN#(IXX%) = 0
            Next IXX%
            
            ConsultaSqlA$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
            ConsultaSqlA$ = ConsultaSqlA$ & " WHERE CODICOM_LAR='" & NroDocumento$ & "' AND TIPOMOVIM='GASTC'"
            Dim RstIva As ADODB.Recordset
            Set RstIva = READSET(ConsultaSqlA$)
        
            With RstIva
                Do While Not .EOF
                    CuentaContable% = XVALO(!CUECONTAI)
                    For IXX% = 1 To CACUEIVA%
                        If CuentaContable% = CUEIVA%(IXX%) Then
                            ICUEN#(IXX%) = ICUEN#(IXX%) + XVALO(!Idebecomp) - XVALO(!IHABECOMP)
                            Exit For
                        End If
                    Next IXX%
                    
                    .MoveNext
                Loop
            End With
            On Error Resume Next
            RstIva.Close
            Set RstIva = Nothing
            On Error GoTo 0
            
            NroProveedor% = XVALO(VALOBADA("PROVED12", "NUME_CLI", "CuMa_Cli = " & CuentaContableBanco%))
            
            For IXX% = 1 To CACUEIVA%
                If ICUEN#(IXX%) > 0.005 Then GoTo 230
            Next IXX%
            GoTo 320
            
230         ZZ$ = REBLA$
            Call REPLA(ZZ$, MKI$(FechaMovimiento%), 3, 2)
            Call REPLA(ZZ$, MKI$(FechaMovimiento%), 5, 2)
            Call REPLA(ZZ$, MKS$(0), 7, 4)
            Call REPLA(ZZ$, NroDocumento$, 11, 18)
            Call REPLA(ZZ$, MKI$(NroProveedor%), 31, 2)
            Call REPLA(ZZ$, rasocli$((-1) * NroProveedor%), 33, 30)
            Call REPLA(ZZ$, CuitCli$((-1) * NroProveedor%), 63, 16)
            Call REPLA(ZZ$, POSIVAC$((-1) * NroProveedor%), 79, 12)
            '
            If Abs(NroProveedor%) < 1 Then
              Call REPLA(ZZ$, REPAS$, 33, 30)
              Call REPLA(ZZ$, "N/Registrado", 63, 16)
              Call REPLA(ZZ$, "N/Def", 79, 12)
            End If
            '
            PROIVA% = 0: ULIVA% = 0
            RELIVA$(0) = ZZ$
            For KU% = 1 To 32
              RELIVA$(KU%) = REBLA$
            Next KU%
            '
            ITOTAX# = 0
            For IXX% = 1 To CACUEIVA%
              If Abs(ICUEN#(IXX%)) >= 0.005 Then
                ZZ$ = RELIVA$(PROIVA%)
                VAMO$ = TASIVANOM$(IXX%)
                If TRIM$(UCase$(VAMO$)) = "PER-IVA" Then
                     ZZ$ = RELIVA$(0)
                     ITOTAA# = ICUEN#(IXX%) + CVD(Mid$(ZZ$, 161, 8))
                     Call REPLA(ZZ$, MKD$(ITOTAA#), 161, 8)
                     ITOTAX# = ITOTAX# + ICUEN#(IXX%)
                     RELIVA$(0) = ZZ$
                     GoTo 299
                   ElseIf TRIM$(UCase$(VAMO$)) = "PER-GAN" Then
                     ZZ$ = RELIVA$(0)
                     ITOTAA# = ICUEN#(IXX%) + CVD(Mid$(ZZ$, 169, 8))
                     Call REPLA(ZZ$, MKD$(ITOTAA#), 169, 8)
                     ITOTAX# = ITOTAX# + ICUEN#(IXX%)
                     RELIVA$(0) = ZZ$
                     GoTo 299
                End If
                Base# = 0
                If XVALO(TASIVANOM$(IXX%)) > 0.5 Then
                  Base# = 100 * ICUEN#(IXX%) / XVALO(TASIVANOM$(IXX%))
                End If
                Call REPLA(ZZ$, VAMO$, 129, 8)
                Call REPLA(ZZ$, MKD$(Base#), 137, 8)
                Call REPLA(ZZ$, MKD$(ICUEN#(IXX%)), 145, 8)
                ITOTAX# = ITOTAX# + Base# + ICUEN#(IXX%)
                If XVALO(TASIVANOM$(IXX%)) > 0.5 Then
                  Call REPLA(ZZ$, MKD$(XVALO(TASIVANOM$(IXX%))), 153, 8)
                End If
                RELIVA$(PROIVA%) = ZZ$
                ULIVA% = PROIVA%
                PROIVA% = PROIVA% + 1
                '
              End If
299         Next IXX%
            '
            ZZ$ = RELIVA$(0)
            Call REPLA(ZZ$, MKD$(ITOTAX#), 91, 8)
            RELIVA$(0) = ZZ$
            '
            UltimoRegistro& = ULTREG&("LIVACOM")
            '
            For KU% = 0 To ULIVA%
              JJ& = JJ& + 1
              Call GRAREG("LIVACOM", RELIVA$(KU%), JJ&)
              Call GRAREG("LIVACCAI", RELIVA$(KU%), JJ&)
            Next KU%
            
            ' INTERCALA PARA MANTENER EL ORDEN
            PuntoDeInsercion& = 0
            For KKL& = 1 To ULTREG&("LIVACOM")
                XY$ = REGLEIDO$("LIVACOM", KKL&)
                FechaRegistro% = CVI(Mid$(XY$, 3, 2))
                If FechaRegistro% = FechaMovimiento% Or FechaRegistro% > FechaMovimiento% Then
                    PuntoDeInsercion& = KKL&
                    Exit For
                End If
            Next KKL&
            
            If PuntoDeInsercion& > 0 Then
                Desplazamiento& = ULIVA% + 1
                For KKL& = (ULTREG&("LIVACOM") - Desplazamiento&) To PuntoDeInsercion& Step -1
                    XY$ = REGLEIDO$("LIVACOM", KKL&)
                    Call GRAREG("LIVACOM", XY$, KKL& + Desplazamiento&)
                Next KKL&
                
                For KU% = 0 To ULIVA%
                  Call GRAREG("LIVACOM", RELIVA$(KU%), PuntoDeInsercion&)
                  PuntoDeInsercion& = PuntoDeInsercion& + 1
                Next KU%
            
            End If
            
320         .MoveNext
        Loop
    End With
    
    On Error Resume Next
    RstIvaTarjetas.Close
    Set RstIvaTarjetas = Nothing
    On Error GoTo 0

Return
End Sub
'\\\OT-05-0559-WAR-FIN///
'
Sub REMECUEN_NUE()
    '
    Screen.MousePointer = 11
    Dim SALINI#(32767), SUFAC#(32767), SUPAG#(32767)
    '
    Call APERBASDAT("CABAN")
    TABORI$ = "CAJABANC"
    CLIMAX% = 0
    NIVCONXXX% = NIVCONFI%
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    
10  SQLX$ = "SELECT NuProve, SUM(IHABECOMP-IDEBECOMP) AS SALDOIN FROM " + TABORI$ + " "
    SQLX$ = SQLX$ + "WHERE FechConta < '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "or TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + "AND (Niv_Clasi IS NULL OR Niv_Clasi <= " & NIVCONXXX% & ") "
    SQLX$ = SQLX$ + "GROUP BY NuProve "
    '
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       If TABORI$ = "CABANVIE" Then GoTo 40
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    With CABATEMP
       If Not .BOF And Not .EOF Then
         .MoveLast
         FIN& = .RecordCount
         .MoveFirst
         JU& = 0
         Do While Not .EOF
           JU& = JU& + 1
           Call TRACE(20, JU&, FIN&)
           NCLIE% = !NUPROVE
           If NCLIE% > 0 Then
             If TICUEPRO((-1) * NCLIE%) = 0 Then ' solo cuentas corrientes
               If NCLIE% > CLIMAX% Then CLIMAX% = NCLIE%
               SALINI#(NCLIE%) = SALINI#(NCLIE%) + !SALDOIN
             End If
           End If
50       .MoveNext
         Loop
       End If
    End With
    CABATEMP.Close
    '
20  SQLX$ = "SELECT NuProve, SUM(IHABECOMP-IDEBECOMP) AS IFAC FROM " + TABORI$ + " "
    SQLX$ = SQLX$ + "WHERE FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + "AND (Niv_Clasi IS NULL OR Niv_Clasi <= " & NIVCONXXX% & ") "
    SQLX$ = SQLX$ + "GROUP BY NuProve "
    '
'    On Error Resume Next
'
'    If Err Then
'       On Error GoTo 0
'       GoTo 30
'    End If
'    On Error GoTo 0
26  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       If TABORI$ = "CABANVIE" Then GoTo 40
       Call CapturaErrorOpen
       GoTo 26
    End If
    On Error GoTo 0
    '
    With CABATEMP
       If Not .BOF And Not .EOF Then
         .MoveLast
         FIN& = .RecordCount
         .MoveFirst
         JU& = 0
         Do While Not .EOF
           JU& = JU& + 1
           Call TRACE(20, JU&, FIN&)
           NCLIE% = !NUPROVE
           If NCLIE% > 0 Then
             If TICUEPRO((-1) * NCLIE%) = 0 Then ' solo cuentas corrientes
               If NCLIE% > CLIMAX% Then CLIMAX% = NCLIE%
               SUFAC#(NCLIE%) = SUFAC#(NCLIE%) + !IFAC
             End If
           End If
60       .MoveNext
         Loop
       End If
    End With
    CABATEMP.Close
    '
30  SQLX$ = "SELECT NuProve, SUM(IHABECOMP-IDEBECOMP) AS IPAG FROM " + TABORI$ + " "
    SQLX$ = SQLX$ + "WHERE FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "AND (Niv_Clasi IS NULL OR Niv_Clasi <= " & NIVCONXXX% & ") "
    SQLX$ = SQLX$ + "GROUP BY NuProve "
    '
'    On Error Resume Next
'    CABATEMP.Open (SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'    If Err Then
'       On Error GoTo 0
'       GoTo 40
'    End If
'    On Error GoTo 0
27  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       If TABORI$ = "CABANVIE" Then GoTo 40
       Call CapturaErrorOpen
       GoTo 27
    End If
    On Error GoTo 0
    '
    With CABATEMP
       If Not .BOF And Not .EOF Then
         .MoveLast
         FIN& = .RecordCount
         .MoveFirst
         JU& = 0
         Do While Not .EOF
           JU& = JU& + 1
           Call TRACE(20, JU&, FIN&)
           NCLIE% = !NUPROVE
           If NCLIE% > 0 Then
             If TICUEPRO((-1) * NCLIE%) = 0 Then ' solo cuentas corrientes
               If NCLIE% > CLIMAX% Then CLIMAX% = NCLIE%
               SUPAG#(NCLIE%) = SUPAG#(NCLIE%) - !IPAG
             End If
           End If
70       .MoveNext
         Loop
       End If
    End With
    CABATEMP.Close
    '
40  On Error GoTo 0
    If TABORI$ = "CAJABANC" Then
      TABORI$ = "CABANVIE"
      GoTo 10
    End If
    Set CABATEMP = Nothing
    
    JJ& = 0
    For k% = 1 To CLIMAX%
      If Abs(SALINI#(k%)) >= 0.005 Or Abs(SUFAC#(k%)) >= 0.005 Or Abs(SUPAG#(k%)) >= 0.005 Then
        X$ = REGBLAN$("RESUCCP")
        Call REPLA(X$, MKI$(k%), 1, 2)
        Call REPLA(X$, MKD$(SALINI#(k%)), 3, 8)
        Call REPLA(X$, MKD$(SUFAC#(k%)), 11, 8)
        Call REPLA(X$, MKD$(SUPAG#(k%)), 19, 8)
        Call REPLA(X$, MKD$(SALINI#(k%) + SUFAC#(k%) - SUPAG#(k%)), 27, 8)
        JJ& = JJ& + 1
        Call GRAREG("RESUCCP", X$, JJ&)
       End If
    Next k%
    Call SETULTREG("RESUCCP", JJ&)
    '
    Call HEADERS("RESUCCP", "")
    Call HEADERS("RESUCCP", "Corresponde a: " + TRIM$(CORRESP$))
    Screen.MousePointer = 1
    '
    If Option1.Value = True Then
         Call FINAL("*RESUCCP")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIRECCP")
    End If
    '
End Sub

Sub REMECUEN()
    '
    Screen.MousePointer = 11
    Dim SALINI#(32767), SUFAC#(32767), SUPAG#(32767)
    '
    Call APERBASDAT("CABAN")
    TABORI$ = "CAJABANC"
    CLIMAX% = 0
    NIVCONXXX% = NIVCONFI%
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    
10  SQLX$ = "SELECT NuProve, SUM(IHABECOMP-IDEBECOMP) AS SALDOIN FROM " + TABORI$ + " "
    SQLX$ = SQLX$ + "WHERE FechConta < '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "or TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + "AND (Niv_Clasi IS NULL OR Niv_Clasi <= " & NIVCONXXX% & ") "
    SQLX$ = SQLX$ + "GROUP BY NuProve "
    '
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       If TABORI$ = "CABANVIE" Then GoTo 40
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    With CABATEMP
       If Not .BOF And Not .EOF Then
         .MoveLast
         FIN& = .RecordCount
         .MoveFirst
         JU& = 0
         Do While Not .EOF
           JU& = JU& + 1
           Call TRACE(20, JU&, FIN&)
           NCLIE% = !NUPROVE
           If NCLIE% > 0 Then
             If TICUEPRO((-1) * NCLIE%) = 0 Then ' solo cuentas corrientes
               If NCLIE% > CLIMAX% Then CLIMAX% = NCLIE%
               SALINI#(NCLIE%) = SALINI#(NCLIE%) + !SALDOIN
             End If
           End If
50       .MoveNext
         Loop
       End If
    End With
    CABATEMP.Close
    '
20  SQLX$ = "SELECT NuProve, SUM(IHABECOMP-IDEBECOMP) AS IFAC FROM " + TABORI$ + " "
    SQLX$ = SQLX$ + "WHERE FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + "AND (Niv_Clasi IS NULL OR Niv_Clasi <= " & NIVCONXXX% & ") "
    SQLX$ = SQLX$ + "GROUP BY NuProve "
    '
'    On Error Resume Next
'
'    If Err Then
'       On Error GoTo 0
'       GoTo 30
'    End If
'    On Error GoTo 0
26  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       If TABORI$ = "CABANVIE" Then GoTo 40
       Call CapturaErrorOpen
       GoTo 26
    End If
    On Error GoTo 0
    '
    With CABATEMP
       If Not .BOF And Not .EOF Then
         .MoveLast
         FIN& = .RecordCount
         .MoveFirst
         JU& = 0
         Do While Not .EOF
           JU& = JU& + 1
           Call TRACE(20, JU&, FIN&)
           NCLIE% = !NUPROVE
           If NCLIE% > 0 Then
             If TICUEPRO((-1) * NCLIE%) = 0 Then ' solo cuentas corrientes
               If NCLIE% > CLIMAX% Then CLIMAX% = NCLIE%
               SUFAC#(NCLIE%) = SUFAC#(NCLIE%) + !IFAC
             End If
           End If
60       .MoveNext
         Loop
       End If
    End With
    CABATEMP.Close
    '
30  SQLX$ = "SELECT NuProve, SUM(IHABECOMP-IDEBECOMP) AS IPAG FROM " + TABORI$ + " "
    SQLX$ = SQLX$ + "WHERE FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "AND (Niv_Clasi IS NULL OR Niv_Clasi <= " & NIVCONXXX% & ") "
    SQLX$ = SQLX$ + "GROUP BY NuProve "
    '
'    On Error Resume Next
'    CABATEMP.Open (SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'    If Err Then
'       On Error GoTo 0
'       GoTo 40
'    End If
'    On Error GoTo 0
27  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       If TABORI$ = "CABANVIE" Then GoTo 40
       Call CapturaErrorOpen
       GoTo 27
    End If
    On Error GoTo 0
    '
    With CABATEMP
       If Not .BOF And Not .EOF Then
         .MoveLast
         FIN& = .RecordCount
         .MoveFirst
         JU& = 0
         Do While Not .EOF
           JU& = JU& + 1
           Call TRACE(20, JU&, FIN&)
           NCLIE% = !NUPROVE
           If NCLIE% > 0 Then
             If TICUEPRO((-1) * NCLIE%) = 0 Then ' solo cuentas corrientes
               If NCLIE% > CLIMAX% Then CLIMAX% = NCLIE%
               SUPAG#(NCLIE%) = SUPAG#(NCLIE%) - !IPAG
             End If
           End If
70       .MoveNext
         Loop
       End If
    End With
    CABATEMP.Close
    '
40  On Error GoTo 0
    If TABORI$ = "CAJABANC" Then
      TABORI$ = "CABANVIE"
      GoTo 10
    End If
    Set CABATEMP = Nothing
    
    JJ& = 0
    For k% = 1 To CLIMAX%
      If Abs(SALINI#(k%)) >= 0.005 Or Abs(SUFAC#(k%)) >= 0.005 Or Abs(SUPAG#(k%)) >= 0.005 Then
        X$ = REGBLAN$("RESUCCP")
        Call REPLA(X$, MKI$(k%), 1, 2)
        Call REPLA(X$, MKD$(SALINI#(k%)), 3, 8)
        Call REPLA(X$, MKD$(SUFAC#(k%)), 11, 8)
        Call REPLA(X$, MKD$(SUPAG#(k%)), 19, 8)
        Call REPLA(X$, MKD$(SALINI#(k%) + SUFAC#(k%) - SUPAG#(k%)), 27, 8)
        JJ& = JJ& + 1
        Call GRAREG("RESUCCP", X$, JJ&)
       End If
    Next k%
    Call SETULTREG("RESUCCP", JJ&)
    '
    Call HEADERS("RESUCCP", "")
    Call HEADERS("RESUCCP", "Corresponde a: " + TRIM$(CORRESP$))
    Screen.MousePointer = 1
    '
    If Option1.Value = True Then
         Call FINAL("*RESUCCP")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIRECCP")
    End If
    '
End Sub
'

Private Sub mnuAjustesOtrosMovimientos_Click()
    Call Command7_Click
End Sub

Private Sub mnuAsociarReporte_Click()
    Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub MnuCerrarPeriodoActivo_Click()
   Call Image2_Click
End Sub

Private Sub MnuConfigFormIva_Click()
    Call Command14_Click
End Sub

Private Sub mnuconsultaAsientocontable_Click()
    Call Command20_Click
End Sub

Private Sub mnuCtasxPagar_Click()
   Call COGECOM07.Command3_Click
End Sub

Private Sub mnuDetalleOrdenPago_Click()
   Call Command2_Click
End Sub

Private Sub mnuEstadisticoCompras_Click()
    Call Command24_Click
End Sub

Private Sub mnuFacturaCompra_Click()
   Call Command3_Click
End Sub

Private Sub mnuGenerarArchivoTransferencia_Click()

End Sub

Private Sub mnuGastosPorProvincia_Click()

    Dim l As New ListaSeleccion
    Dim rp As New ReporteGrilla
    Dim Grilla As MSFlexGrid
    
    Set Grilla = FRMZELECHO.msf2
    
    Call DESHASFECHA(Des%, Has%, PERIODO$)
    If PERIODO$ = "" Then Exit Sub
    
    DesdeFecha$ = FETEXCOR1$(Des%)
    HastaFecha$ = FETEXCOR1$(Has%)
    
    l.TITULO = "Proveedores sin Provincia"
    l.AnchoLista = 8000
    l.EncabezadoDeColumnas = "Cta/I10;Razon Social/A48"
    l.ConsultaSql = "select distinct p.nume_cli, p.raso_cli from" _
    & " (select c.nuprove, c.idebecomp, c.ihabecomp from cajabanc c with(nolock) inner join plancuen p with(nolock) on (c.cuecontai=p.codincue) where c.tipomovim='compra' and p.tipocue='gasto' and c.fechconta>='" & DesdeFecha$ & "' and c.fechconta<='" & HastaFecha$ & "') g" _
    & " left join proved12 p with(nolock) on (g.nuprove=p.nume_cli)" _
    & " where p.prov_cli = '' or p.prov_cli is null"
    
    Set Rst = READSET(l.ConsultaSql)
    If Rst.RecordCount > 0 Then
        Call COMUNI("Se Han Detectado Proveedores sin Provincia.\ \Acepte para Obtener una Lista.")
        Call l.Mostrar
    End If
       
    rp.TITULO = "GASTOS POR PROVINCIA - (" & PERIODO$ & ")"
    rp.CodigoReporte = "GASXPROV"
    rp.MENU1.Visible = True: rp.MENU1.Caption = "Imprimir"
    rp.MenuEdicion = False
    For i& = 1 To ULTREG&("PROVINC")
        CodigoProvincia$ = Mid$(REGLEIDO$("PROVINC", i&), 1, 1)
        Provincia$ = Mid$(REGLEIDO$("PROVINC", i&), 2, 20)
        EncabezadoProvincias$ = EncabezadoProvincias$ & CodigoProvincia$ & "-" & Provincia$ & "/F12.2;"
    Next i&
    rp.EncabezadoDeColumnas = "C.Contable/A10;Denominación/A24;Sin Prov./F12.2;" & EncabezadoProvincias$ & ";Totales/F12.2"
       
    Call CARGA_ENCABEZADO(Grilla, rp.EncabezadoDeColumnas, FRMZELECHO)
    ConsultaSql$ = "select cuecontai, prov_cli, (debe-habe) saldo from " _
    & " (select c.cuecontai, d.prov_cli, sum(idebecomp) debe, sum(ihabecomp) habe from cajabanc c with(nolock) inner join proved12 d with(nolock) on (c.nuprove=d.nume_cli)" _
    & " where c.fechconta>='" & DesdeFecha$ & "' and c.fechconta<='" & HastaFecha$ & "' and c.tipomovim='compra' group by c.cuecontai, d.prov_cli) g" _
    & " inner join plancuen p with(nolock) on (g.cuecontai=p.codincue) where p.tipocue='gasto'" _
    & " order by cuecontai, prov_cli"
    Set Rst = READSET(ConsultaSql$)
    With Rst
        CuentaContable% = 0
        Do While Not .EOF
            If CuentaContable% <> !CUECONTAI Then
                Grilla.Rows = Grilla.Rows + 1
                CuentaContable = XVALO(!CUECONTAI)
                Grilla.TextMatrix(Grilla.Rows - 1, 0) = "*"
                Grilla.TextMatrix(Grilla.Rows - 1, 1) = CODCUE$(CuentaContable%)
                Grilla.TextMatrix(Grilla.Rows - 1, 2) = DENOCUE$(CuentaContable%)
            End If
            
            CodigoProvincia$ = SAFETEXT$(!Prov_Cli)
            COL% = 0
            For c% = 4 To Grilla.Cols - 2
                If CodigoProvincia$ = Mid$(Grilla.TextMatrix(0, c%), 1, 1) Then
                    COL% = c%
                    Exit For
                End If
            Next c%
            
            If COL% = 0 Then
                ImporteActual# = XVALO(Grilla.TextMatrix(Grilla.Rows - 1, 2))
                ImporteActual# = ImporteActual# + !SALDO
                Grilla.TextMatrix(Grilla.Rows - 1, 3) = TRIM$(FORMATNUM$(ImporteActual#, "F12.2"))
            Else
                ImporteActual# = XVALO(Grilla.TextMatrix(Grilla.Rows - 1, COL%))
                ImporteActual# = ImporteActual# + !SALDO
                Grilla.TextMatrix(Grilla.Rows - 1, COL%) = TRIM$(FORMATNUM$(ImporteActual#, "F12.2"))
            End If
                
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    For F% = 1 To Grilla.Rows - 1
        TotalPorCuenta# = 0
        For c% = 3 To Grilla.Cols - 2
            TotalPorCuenta# = TotalPorCuenta# + XVALO(Grilla.TextMatrix(F%, c%))
        Next c%
        Grilla.TextMatrix(F%, Grilla.Cols - 1) = TRIM$(FORMATNUM$(TotalPorCuenta#, "F12.2"))
    Next F%
    
    Set rp.Grilla = Grilla
    Call rp.Mostrar
    
End Sub

Private Sub mnuLibroIvaDigitaGenerarArchivo_Click()

    If VALIMESAN% < 1 Then Exit Sub
    Call GENIVACOM
    Call LibroIvaDigitalComprasLocales
    Call LibroIvaDigitalComprasImportaciones
    
    Call COMUNI("Libro Iva Digital de Compras: Generación Ok")

End Sub

Private Sub mnuLibroIvaDigitalCompararMisComprobantes_Click()

    If VALIMESAN% < 1 Then Exit Sub

    Campos$ = "F.Emisión/D8;F.Contable/D8;F.Acred/D0;F.Real/D0;Comprobante/A18;Cuit/A13;Proveedor/A20;Imp. Total/F12.2;Observaciones/A24"
    Call CARGA_ENCABEZADO(MISCBTESAFIP.mfgComprobantesSistema, Campos$, FRMZELECHO)

    ConsultaSql$ = "select c.fechemisi, c.fechconta, c.fechacred, c.fehoreal, c.codicom_lar, replace(p.cuit_cli,'-','') cuit, p.raso_cli, c.valabscomp, '' Observaciones from" _
    & " (select fechemisi, fechconta, fechacred, fehoreal, codicom_lar, nuprove, valabscomp from cajabanc with (nolock)" _
    & " where fechConta >= '" & DESDAT & "'" _
    & " and fechConta<= '" & HASDAT & "'" _
    & " and tipoMovim = 'fcom'" _
    & " and (incsubiva <> 'no'" _
    & " or incsubiva is null or incsubiva = '')) c" _
    & " inner join proved12 p with(nolock) on (c.nuprove=p.nume_cli)" _
    & " order by c.fechconta asc, c.fechacred asc, c.fehoreal asc"

    Set Rst = READSET(ConsultaSql$)
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    FilaLibre% = 0
    If InStr(EPAC$, "AHP") > 0 Then
    FilaLibre% = 1
    End If
    Call Carga_MSF_Recordset(Rst, Campos$, MISCBTESAFIP.mfgComprobantesSistema, FilaLibre%)
    If Rst.RecordCount = 0 Then MISCBTESAFIP.mfgComprobantesSistema.Rows = 1
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    MISCBTESAFIP.PeriodoDesde = Desde%
    MISCBTESAFIP.PeriodoHasta = Hasta%
    MISCBTESAFIP.Show 1

End Sub

Private Sub mnulistadoCtaCble_Click()
   Call Command19_Click
End Sub

Private Sub mnuMaestroProved_Click()
   Call Command15_Click
End Sub

Private Sub mnuNotaCredito_Click()
   Call Command5_Click
End Sub

Private Sub mnuNotaDebito_Click()
    Call Command4_Click
End Sub

Private Sub mnuRankingArticuloxFecha_Click()
   Call Command22_Click
End Sub

Private Sub mnuReimputacionComprobanteCompras_Click()
   Call Command10_Click
End Sub

Private Sub mnuReportesVarios_Click()
   Call AGenRep_Click
End Sub

Private Sub mnuResumenMensualCta_Click()
   Call Command6_Click
End Sub

Private Sub mnuResumenMensualivacomp_Click()
   Call Command13_Click
End Sub

Private Sub mnuSaldosydeudasproveedores_Click()
   Call COGECOM07.Command5_Click
End Sub

Private Sub mnuSubdiarioCompras_Click()
   Call Command16_Click
End Sub

Private Sub mnuSubdirioivacompras_Click()
   Call Command21_Click
End Sub

Private Sub mnuVer_Click()

End Sub

Private Sub nnuAbrirPeriodoActivo_Click()
  Call Image1_Click
End Sub

Private Sub RankingComprasProvxFecha_Click()
   Call Command17_Click
End Sub

Private Sub RegimenCompras_Click()

    If VALIMESAN% < 1 Then Exit Sub
    Call GENIVACOM
    Call RegimenInformativoCompras
    
End Sub


