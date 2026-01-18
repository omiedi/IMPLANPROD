VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ORPAGO07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Pago (Operaciones Comerciales de Compras)"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ORPAGO07.frx":0000
   ScaleHeight     =   8295
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C7D9CC&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8500
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command20 
         Caption         =   "Command20"
         Height          =   375
         Left            =   7440
         TabIndex        =   111
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00C7D9CC&
         Caption         =   "M/Extr."
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
         Left            =   6000
         TabIndex        =   98
         Top             =   120
         Width           =   2565
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "T-C:"
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
            Left            =   120
            TabIndex        =   101
            ToolTipText     =   "Doble Click para Configurar Cliente"
            Top             =   285
            Width           =   345
         End
         Begin VB.Label MONEMIS 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   11.25
               Charset         =   0
               Weight          =   900
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   1540
            TabIndex        =   100
            ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
            Top             =   150
            Width           =   960
         End
         Begin VB.Label TICAMBIS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   525
            TabIndex        =   99
            ToolTipText     =   "Doble Click para Ajustar Tipo de Cambio"
            Top             =   210
            Width           =   960
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   210
         ScaleHeight     =   360
         ScaleWidth      =   10425
         TabIndex        =   76
         Top             =   6830
         Width           =   10460
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   80
            Top             =   120
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Roboto Mono Medium"
               Size            =   9.75
               Charset         =   0
               Weight          =   500
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   6720
            TabIndex        =   79
            Top             =   35
            Width           =   1530
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8415
            TabIndex        =   78
            Top             =   45
            Width           =   1605
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Total General"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   77
            Top             =   50
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00C7D9CC&
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   5880
         ScaleHeight     =   540
         ScaleWidth      =   960
         TabIndex        =   67
         Top             =   680
         Visible         =   0   'False
         Width           =   960
         Begin VB.Image Image2 
            Height          =   480
            Left            =   0
            Picture         =   "ORPAGO07.frx":0342
            Top             =   0
            Width           =   480
         End
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   420
            TabIndex        =   68
            Top             =   120
            Width           =   450
         End
      End
      Begin VB.TextBox Text7 
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
         Height          =   330
         Left            =   1470
         TabIndex        =   60
         Top             =   1260
         Width           =   3690
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H0050765A&
         ForeColor       =   &H80000008&
         Height          =   9000
         Left            =   10920
         ScaleHeight     =   8970
         ScaleWidth      =   945
         TabIndex        =   30
         Top             =   0
         Width           =   975
         Begin VB.PictureBox MARCOBARRA 
            Appearance      =   0  'Flat
            BackColor       =   &H000C93DC&
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   140
            ScaleHeight     =   165
            ScaleWidth      =   615
            TabIndex        =   104
            Top             =   6840
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H0076C9F5&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   -120
               TabIndex        =   105
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.CommandButton Command2 
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
            Left            =   140
            Picture         =   "ORPAGO07.frx":064C
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Terminar - Salir de la Aplicación"
            Top             =   200
            Width           =   650
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Cancel"
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
            Left            =   140
            Picture         =   "ORPAGO07.frx":0796
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Anulación de Orden de Pago Emitida"
            Top             =   5220
            Visible         =   0   'False
            Width           =   650
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00404000&
            Height          =   8310
            Left            =   10950
            ScaleHeight     =   8250
            ScaleWidth      =   915
            TabIndex        =   39
            Top             =   0
            Width           =   975
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
               Left            =   140
               Picture         =   "ORPAGO07.frx":0AA0
               Style           =   1  'Graphical
               TabIndex        =   50
               ToolTipText     =   "Cambia Ordenamiento de Valores (Fecha / Importe)"
               Top             =   3360
               Width           =   650
            End
            Begin VB.PictureBox Picture5 
               BackColor       =   &H000000FF&
               Height          =   135
               Left            =   140
               ScaleHeight     =   75
               ScaleWidth      =   585
               TabIndex        =   48
               Top             =   6720
               Width           =   650
               Begin VB.Frame Frame7 
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   0  'None
                  Height          =   500
                  Left            =   0
                  TabIndex        =   49
                  Top             =   0
                  Width           =   12000
               End
            End
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
               Height          =   640
               Left            =   140
               Picture         =   "ORPAGO07.frx":0DAA
               Style           =   1  'Graphical
               TabIndex        =   47
               ToolTipText     =   "Configuración de Funcionamiento"
               Top             =   5775
               Width           =   650
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
               Height          =   640
               Left            =   140
               Picture         =   "ORPAGO07.frx":11EC
               Style           =   1  'Graphical
               TabIndex        =   46
               ToolTipText     =   "Anular, Re-Imprimir, Imputación de Créditos, Tablas"
               Top             =   5145
               Width           =   650
            End
            Begin VB.CommandButton Command16 
               Caption         =   "Imprimir"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   960
               Left            =   140
               Picture         =   "ORPAGO07.frx":1336
               Style           =   1  'Graphical
               TabIndex        =   45
               ToolTipText     =   "Aprueba - Graba e Imprime Orden de Pago"
               Top             =   7140
               Width           =   650
            End
            Begin VB.CommandButton Command15 
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
               Left            =   140
               Picture         =   "ORPAGO07.frx":19A0
               Style           =   1  'Graphical
               TabIndex        =   44
               ToolTipText     =   "Archivo Maestro de Proveedores"
               Top             =   1680
               Width           =   650
            End
            Begin VB.CommandButton Command14 
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
               Left            =   140
               Picture         =   "ORPAGO07.frx":1CAA
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Resumen de Cuenta Corriente"
               Top             =   2310
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
               Height          =   640
               Left            =   140
               Picture         =   "ORPAGO07.frx":1FB4
               Style           =   1  'Graphical
               TabIndex        =   42
               ToolTipText     =   "Notas de Proveedores"
               Top             =   3990
               Width           =   650
            End
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
               Height          =   540
               Left            =   105
               Picture         =   "ORPAGO07.frx":22BE
               TabIndex        =   41
               ToolTipText     =   "Abandona la Aplicación"
               Top             =   105
               Width           =   650
            End
            Begin VB.CommandButton Command10 
               Caption         =   "Next"
               Enabled         =   0   'False
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
               Left            =   140
               Picture         =   "ORPAGO07.frx":25C8
               Style           =   1  'Graphical
               TabIndex        =   40
               ToolTipText     =   "Carga de Valores de Pago Propios"
               Top             =   735
               Width           =   650
            End
         End
         Begin VB.CommandButton BOTNEXT 
            Caption         =   "Next"
            Enabled         =   0   'False
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
            Left            =   140
            Picture         =   "ORPAGO07.frx":28D2
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Carga de Valores de Pago Propios"
            Top             =   890
            Width           =   650
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Note"
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
            Left            =   140
            Picture         =   "ORPAGO07.frx":2BDC
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Notas de Proveedores"
            Top             =   3080
            Visible         =   0   'False
            Width           =   650
         End
         Begin VB.CommandButton Command4 
            Caption         =   "R-Cue"
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
            Left            =   140
            Picture         =   "ORPAGO07.frx":2EE6
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   1700
            Width           =   650
         End
         Begin VB.CommandButton Command11 
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
            Left            =   140
            Picture         =   "ORPAGO07.frx":31F0
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Archivo Maestro de Proveedores"
            Top             =   3770
            Visible         =   0   'False
            Width           =   650
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Print"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   140
            Picture         =   "ORPAGO07.frx":34FA
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Aprueba - Graba e Imprime Orden de Pago"
            Top             =   7140
            Width           =   650
         End
         Begin VB.CommandButton Command6 
            Caption         =   "More"
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
            Left            =   140
            Picture         =   "ORPAGO07.frx":3B64
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Anular, Re-Imprimir, Imputación de Créditos, Tablas"
            Top             =   2390
            Width           =   650
         End
         Begin VB.CommandButton Command7 
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
            Left            =   140
            Picture         =   "ORPAGO07.frx":3CAE
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   6000
            Visible         =   0   'False
            Width           =   650
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Sort"
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
            Left            =   140
            Picture         =   "ORPAGO07.frx":40F0
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Cambia Ordenamiento de Valores (Fecha / Importe)"
            Top             =   4530
            Width           =   650
         End
      End
      Begin VB.ListBox LINDICE 
         Height          =   1035
         Left            =   5355
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   2520
         Visible         =   0   'False
         Width           =   1170
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
         Height          =   510
         Left            =   6930
         TabIndex        =   28
         Top             =   2835
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C7D9CC&
         BorderStyle     =   0  'None
         Height          =   640
         Left            =   8625
         TabIndex        =   27
         Top             =   110
         Width           =   2100
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Orden de Pago"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   600
            Left            =   900
            TabIndex        =   97
            Top             =   50
            Value           =   -1  'True
            Width           =   1100
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Fondo Fijo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   600
            Left            =   10
            TabIndex        =   96
            Top             =   50
            Visible         =   0   'False
            Width           =   900
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   2540
         Left            =   210
         ScaleHeight     =   2505
         ScaleWidth      =   10425
         TabIndex        =   16
         Top             =   1750
         Width           =   10450
         Begin VB.Frame Frame4 
            Caption         =   "Frame11"
            Height          =   2900
            Left            =   10100
            TabIndex        =   87
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   600
            TabIndex        =   86
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame2 
            Caption         =   "Frame7"
            Height          =   2900
            Left            =   2760
            TabIndex        =   85
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame8 
            Caption         =   "Frame8"
            Height          =   2900
            Left            =   4560
            TabIndex        =   84
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame9 
            Caption         =   "Frame9"
            Height          =   2900
            Left            =   5640
            TabIndex        =   83
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame10 
            Caption         =   "Frame10"
            Height          =   2900
            Left            =   6960
            TabIndex        =   82
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame11 
            Caption         =   "Frame11"
            Height          =   2900
            Left            =   8400
            TabIndex        =   81
            Top             =   0
            Width           =   15
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Roboto Mono Medium"
               Size            =   9.75
               Charset         =   0
               Weight          =   500
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2190
            Left            =   0
            TabIndex        =   74
            Top             =   330
            Width           =   10425
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " M-P  Referencia        Número         Emisión Vencimiento   Importe      Selección"
            BeginProperty Font 
               Name            =   "Roboto Mono"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   0
            TabIndex        =   102
            Top             =   60
            Width           =   10410
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2275
         Left            =   210
         ScaleHeight     =   2250
         ScaleWidth      =   10425
         TabIndex        =   15
         Top             =   4550
         Width           =   10460
         Begin VB.Frame Frame17 
            Caption         =   "Frame11"
            Height          =   2900
            Left            =   8350
            TabIndex        =   95
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame16 
            Caption         =   "Frame10"
            Height          =   2900
            Left            =   6700
            TabIndex        =   94
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame15 
            Caption         =   "Frame9"
            Height          =   2900
            Left            =   5400
            TabIndex        =   93
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame14 
            Caption         =   "Frame8"
            Height          =   2900
            Left            =   3650
            TabIndex        =   92
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame13 
            Caption         =   "Frame7"
            Height          =   2900
            Left            =   2500
            TabIndex        =   91
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame5 
            Caption         =   "Frame11"
            Height          =   2900
            Left            =   10100
            TabIndex        =   90
            Top             =   0
            Width           =   15
         End
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Roboto Mono Medium"
               Size            =   9.75
               Charset         =   0
               Weight          =   500
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1920
            Left            =   0
            TabIndex        =   75
            Top             =   330
            Width           =   10425
         End
         Begin VB.Label Label36 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " Documento Número     Emisión        Importe Vencimiento       Saldo       Acredita       "
            BeginProperty Font 
               Name            =   "Roboto Mono"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   0
            TabIndex        =   103
            Top             =   60
            Width           =   10410
         End
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   5280
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   210
         Width           =   540
      End
      Begin VB.TextBox Text11 
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
         Left            =   9030
         TabIndex        =   4
         Text            =   " "
         Top             =   1260
         Width           =   1650
      End
      Begin VB.TextBox Text9 
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
         Left            =   9345
         TabIndex        =   2
         Text            =   " "
         Top             =   870
         Width           =   1170
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10500
         TabIndex        =   3
         Top             =   870
         Width           =   180
      End
      Begin VB.TextBox Text5 
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
         Height          =   300
         Left            =   2310
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   920
         Width           =   2850
      End
      Begin VB.TextBox Text4 
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
         Height          =   300
         Left            =   1470
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   920
         Width           =   750
      End
      Begin VB.TextBox Text3 
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
         Height          =   300
         Left            =   1470
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   560
         Width           =   3690
      End
      Begin VB.TextBox Text2 
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
         Height          =   300
         Left            =   1470
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   210
         Width           =   3690
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1100
         TabIndex        =   1
         Top             =   560
         Width           =   180
      End
      Begin VB.TextBox Text1 
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
         Left            =   240
         TabIndex        =   0
         Text            =   " "
         Top             =   560
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7800
         TabIndex        =   118
         ToolTipText     =   "Doble Click para Editar Valores"
         Top             =   7920
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Ret. IIBB Tucumán"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6360
         TabIndex        =   117
         Top             =   8040
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6300
         TabIndex        =   116
         ToolTipText     =   "Doble Click para Editar Valores"
         Top             =   7560
         Width           =   1380
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "Retención de IVA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6300
         TabIndex        =   115
         Top             =   7320
         Width           =   1380
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Ret. IIBB Misiones "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3360
         TabIndex        =   114
         Top             =   8040
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4800
         TabIndex        =   113
         ToolTipText     =   "Doble Click para Editar Valores"
         Top             =   7920
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6480
         TabIndex        =   110
         Top             =   7680
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         TabIndex        =   109
         Top             =   6960
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9600
         TabIndex        =   108
         ToolTipText     =   "Doble Click para Editar Valores"
         Top             =   6840
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4800
         TabIndex        =   107
         ToolTipText     =   "Doble Click para Editar Valores"
         Top             =   7560
         Width           =   1380
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Ret. I.Brutos CABA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4800
         TabIndex        =   106
         Top             =   7320
         Width           =   1380
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C. Imp:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   0
         TabIndex        =   71
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label44 
         Height          =   225
         Left            =   780
         TabIndex        =   70
         Top             =   30
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6960
         TabIndex        =   66
         Top             =   7200
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   7980
         TabIndex        =   65
         Top             =   7035
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7800
         TabIndex        =   64
         ToolTipText     =   "Doble-Click para Ajustar Importe Retenido"
         Top             =   7560
         Width           =   1380
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Valores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1750
         TabIndex        =   26
         Top             =   7320
         Width           =   780
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Ret. I.Brutos BsAs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3280
         TabIndex        =   21
         Top             =   7320
         Width           =   1380
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Ret. Ganancias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7800
         TabIndex        =   63
         Top             =   7320
         Width           =   1620
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9345
         TabIndex        =   62
         ToolTipText     =   "Doble Click para Editar Valores"
         Top             =   7035
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O/Cheque:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -420
         TabIndex        =   61
         Top             =   1320
         Width           =   1800
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5985
         TabIndex        =   59
         Top             =   7245
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label LabeL28 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4725
         TabIndex        =   58
         Top             =   7245
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         TabIndex        =   57
         Top             =   7245
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   56
         Top             =   7245
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   55
         Top             =   7245
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   54
         ToolTipText     =   "Doble - Click para visualzar aplicación de Créditos y Anticipos"
         Top             =   300
         Width           =   960
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   240
         TabIndex        =   53
         Top             =   915
         Width           =   1035
      End
      Begin VB.Line Line21 
         X1              =   210
         X2              =   10660
         Y1              =   4010
         Y2              =   4010
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1800
         TabIndex        =   25
         Top             =   7560
         Width           =   1380
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9300
         TabIndex        =   24
         Top             =   7560
         Width           =   1380
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9300
         TabIndex        =   23
         Top             =   7320
         Width           =   780
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   240
         TabIndex        =   22
         Top             =   7560
         Width           =   1380
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3780
         TabIndex        =   20
         Top             =   5775
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Valores:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         TabIndex        =   19
         Top             =   5880
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3280
         TabIndex        =   18
         ToolTipText     =   "Doble Click para Editar Valores"
         Top             =   7560
         Width           =   1380
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Créditos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   7320
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Documento no Válido como Factura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5130
         TabIndex        =   14
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   1200
         Width           =   2115
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Numero"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7560
         TabIndex        =   12
         Top             =   1350
         Width           =   1380
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8760
         TabIndex        =   11
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label46 
         Height          =   225
         Left            =   780
         TabIndex        =   69
         Top             =   270
         Visible         =   0   'False
         Width           =   585
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GRID2 
      Height          =   765
      Left            =   360
      TabIndex        =   88
      Top             =   2400
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   1349
      _Version        =   393216
      Rows            =   23
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      Redraw          =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid GRID1 
      Height          =   765
      Left            =   1320
      TabIndex        =   89
      Top             =   3480
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   1349
      _Version        =   393216
      Rows            =   23
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      Redraw          =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   72
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ORPAGO07.frx":43FA
      TabIndex        =   73
      Top             =   960
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "ORPAGO07.frx":4474
      Top             =   1200
   End
   Begin VB.Label Label43 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Roboto Mono Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      TabIndex        =   112
      ToolTipText     =   "Doble Click para Editar Valores"
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pos.IVA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu NueOPs 
         Caption         =   "Nueva Orden de Pago"
      End
      Begin VB.Menu mnuGuardar 
         Caption         =   "Grabar Orden de Pago"
      End
      Begin VB.Menu AnuOps 
         Caption         =   "Anulación de Orden de Pago"
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFondoFijo 
         Caption         =   "Fondo Fijo"
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSali 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu ConsuLis 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuReCuCor 
         Caption         =   "Resumen de Cuenta Corriente"
      End
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuApliCredAnt 
         Caption         =   "Aplicación de Créditos y Anticipos"
      End
      Begin VB.Menu ListOP 
         Caption         =   "Listado de Ordenes de Pago"
      End
      Begin VB.Menu mnuAnalisisEstaPagos 
         Caption         =   "Análisis Estadístico de Pagos"
      End
      Begin VB.Menu s4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportesVarios 
         Caption         =   "Reportes Varios"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu veryreimp 
         Caption         =   "Ver y Re-Imprimir Orden de Pago"
      End
      Begin VB.Menu mnuVeryReimprmirRetencionesGanancia 
         Caption         =   "Ver y Re-Imprimir Retenciones de Ganancias e IIBB"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuCambarOrdenamiento 
         Caption         =   "Cambia Ordenamiento de Valores (Fecha / Importe)"
      End
      Begin VB.Menu linea2 
         Caption         =   "-"
      End
      Begin VB.Menu impucredi 
         Caption         =   "Imputación de Créditos"
      End
      Begin VB.Menu mnuDesaplicarOP 
         Caption         =   "Desaplicar Orden de Pago"
      End
      Begin VB.Menu AjuSaldAc 
         Caption         =   "Ajuste de Saldos Acreedores"
      End
      Begin VB.Menu mnuCompensacionClteProv 
         Caption         =   "Compensación Cliente-Proveedor"
      End
      Begin VB.Menu mnuTransfPago 
         Caption         =   "Transferencia de Orden de Pago a otro Proveedor"
      End
      Begin VB.Menu NotProve 
         Caption         =   "Notas de Proveedores"
      End
      Begin VB.Menu mnuAplicaReparacionesAPagos 
         Caption         =   "Aplicación de Pagos a Reparaciones"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuconfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuAsociarReporte 
         Caption         =   "Vincular Reportes al Formulario"
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTaMePa 
         Caption         =   "Tabla de Bancos y Valores de Pago"
      End
      Begin VB.Menu ActualizaPadronIBCABA 
         Caption         =   "Actualizacion Padron de IIBB CABA"
      End
      Begin VB.Menu ActualizaPadronIIBBTucuman 
         Caption         =   "Actualización Padrón de IIBB Tucumán"
      End
      Begin VB.Menu ControlDeEstadoValoresEntregados 
         Caption         =   "Control de Estado de Valores Entregados"
      End
      Begin VB.Menu linea3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfiGral 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu mnuAccDir 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuMaeProv 
         Caption         =   "Maestro de Proveedores"
      End
   End
End
Attribute VB_Name = "ORPAGO07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$
Dim MODOSORVAL%
Dim YACARFO%

Dim RETIBMAN%, RETGANMAN%
'
Dim DETANT$(), CADETANT%
Dim MODOMONEI%, TICAMBX#
Dim OrdenDePagoActiva$
Dim RetieneIBMisiones%, MedioPagoIBMisiones%, PorcentajeRetencionIBMisiones#, MinimoRetencionIBMisiones#, MontoSujetoIBMisiones#, NumeroRetencionIBMisiones&
Dim MinimoImponibleRetencionIva#, MedioPagoRetencionIva%, NumeroRetencionIva&

Dim NumeroRetencionIBTucuman&, MedioPagoIBTucuman%, MontoSujetoIBTucuman#

Sub CALRETIBRU_CABA()

    RetencionTotal# = 0
    For IndiceComprobantes% = 1 To List2.ListCount - 1
        Renglon$ = List2.LIST(IndiceComprobantes%)
        ImporteRetenido# = XVALO(Mid$(Renglon$, 165, 12))
        RetencionTotal# = RetencionTotal# + ImporteRetenido#
    Next IndiceComprobantes%
    
    Label38.Caption = TRIM$(FORMATNUM$(RetencionTotal#, "F12.2"))
    
End Sub

Sub CALRETIBRU_TUCUMAN()

    ' *** VERIFICA SI TIENE QUE RETENER IIBB POR CABA, A PARTIR DE QUE EXISTA EL PADRON DE ALICUOTA
    If EXISTE%(LUDAT$ + "IIBBTUCU.TXT") > 0 Then
        
        NumeroProveedor% = XVALO(Text1)
        ' *** OBTIENE ALICUOTA, MINIMO IMPONIBLE Y MEDIO DE PAGO
10      TasaRetencion# = ALIRETBRU_TUCUMAN#(NumeroProveedor%, MinimoImponible#, MedioPagoIBTucuman%)

        If TasaRetencion# > 0.005 Then
            NumeroProveedor% = (-1) * NumeroProveedor%
            PosicionIva% = PIVACLI(NumeroProveedor%)
            If PosicionIva% = 1 Then
                BASECALCULO# = NETOCANCE#           ' PARA RESP. INSCRIPTOS: NETO DE LOS COMPROBANTES ELEGIDOS
            Else
                BASECALCULO# = IMPOCANCE_TUCUMAN#   ' PARA EL RESTO: TOTAL DE LOS COMPROBANTES ELEGIDOS MENOS OTROS IMPUESTOS NACIONALES O PROVINCIALES
            End If
            If BASECALCULO# > MinimoImponible# Then
                ImporteRetenido# = BASECALCULO# * (TasaRetencion# / 100)
                Label12 = FORMATNUM$(ImporteRetenido#, "F10.2")
                
                MontoSujetoIBTucuman# = BASECALCULO#
            End If
            
        Else
            ' *** SI EL CONTRIBUYENTE NO FIGURA EN EL PADRON Y ES DE LA PROVINCIA DE TUCUMAN,
            ' *** SE LE ADVIERTE AL USUARIO PARA QUE CONFIGURE ALGUNA DE LAS CATEGORIAS
            CodigoProvinciaConfigurada$ = Control("IIBBTUCU", "CODPROVI")
            NumeroProveedor% = (-1) * NumeroProveedor%
            codigoProvincia$ = UCase(PROVCLI$(NumeroProveedor%))
            
            If codigoProvincia$ = CodigoProvinciaConfigurada$ And CodigoProvinciaConfigurada$ <> "" Then
                Call COMUNI("El Proveedor Elegido Corresponde a la Provincia de Túcuman\y No se Encuentra Registrado en el Padrón.\ \Deberá Aplicar Alguna de las Categorías de Retención\que Aparecerón luego de Aceptar el presente mensaje.")
                
                Call SELECHO("TACAIBTU")
                CATIBTUCU% = VALACT1$("TACAIBTU")
                If CATIBTUCU% > 0 Then
                    Call ACTUREGISTRO("PROVED12", "", "NUME_CLI=" & Text1, CStr(CATIBTUCU%), REGAF&)
                    GoTo 10
                Else
                    Call COMUNI("No Se Ha Elegido una Categoría.\Se Obviará el Cálculo de la Retención de IIBB Tucumán.")
                End If
                
            End If
            
        End If
    End If
    
End Sub

Sub CALRETIVA()

    RetencionTotal# = 0
    For IndiceComprobantes% = 1 To List2.ListCount - 1
        Renglon$ = List2.LIST(IndiceComprobantes%)
        ImporteRetenido# = XVALO(Mid$(Renglon$, 177, 12))
        RetencionTotal# = RetencionTotal# + ImporteRetenido#
    Next IndiceComprobantes%
    
    Label49.Caption = ""
    If RetencionTotal# > 0.005 Then
        Label49.Caption = TRIM$(FORMATNUM$(RetencionTotal#, "F12.2"))
    End If
    
End Sub

Sub CALRETGANANCIAS_FACTURAS_M()

    RetencionTotal# = 0
    For IndiceComprobantes% = 1 To List2.ListCount - 1
        Renglon$ = List2.LIST(IndiceComprobantes%)
        
        Comprobante$ = Mid$(Renglon$, 1, 18)
        ImporteACancelar# = XVALO(TRIM$(Mid$(Renglon$, 70, 14)))
        ' *** VERIFICA QUE HAYA ELEGIDO EL COMPROBANTE MEDIANTE LA EXISTENCIA DEL IMPORTE QUE ESTA POR CANCELAR
        If Mid$(Comprobante$, 4, 1) = "M" And Abs(ImporteACancelar#) > 0.005 Then
            ImporteNeto# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP) - SUM(IHABECOMP)", "CODICOM_LAR='" & Comprobante$ & "' AND TIPOMOVIM='COMPRA' AND NUPROVE=" & Text1))
            ImporteNetoTotal# = ImporteNetoTotal# + ImporteNeto#
        End If
        
    Next IndiceComprobantes%
    
    Label30.Caption = TRIM$(FORMATNUM$(ImporteNetoTotal# * 0.03, "F12.2"))
    
End Sub

Sub CALRETIBRU_MISIONES()

    ' *** VERIFICA SI TIENE QUE RETENER POR ORDEN Y CUENTA DE MISIONES
    ImporteRetenido# = 0
    If RetieneIBMisiones% > 0 Then
        ImporteNetoACancelar# = NETOCANCE#
        If ImporteNetoACancelar# > 0.005 Then
            MontoSujetoIBMisiones# = ImporteNetoACancelar#
            ImporteRetenido# = ImporteNetoACancelar# * (PorcentajeRetencionIBMisiones# / 100)
            If Abs(ImporteRetenido# - MinimoRetencionIBMisiones#) < 0.005 Then
                ImporteRetenido# = 0
                MontoSujetoIBMisiones# = 0
            End If
        End If
    End If
    
    Label45.Caption = TRIM$(FORMATNUM$(ImporteRetenido#, "F12.2"))
    
End Sub

Private Sub ActualizaPadronIBCABA_Click()

    Screen.MousePointer = 11
    
    Dim AlicuotaPadronIBCABA() As String
    
    ' VERIFICA LA EXISTENCIA DEL PADRON DE IIBB
    If EXISTE%(LUDAT$ + "IIBBCABA.TXT") = 0 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Falta padrón de alícuotas de IIBB CABA")
        GoTo 99
    End If
    
    ' ELIMINA PADRON SIMPLIFICADO ACTUAL
    If EXISTE%(LUDAT$ + "PAIBCABA.TXT") > 0 Then
        On Error Resume Next
        Call Kill(LUDAT$ & "PAIBCABA.TXT")
        If Err > 0 Then
            Screen.MousePointer = 1
            Call MENSERR(24, "Error al Actualizar Padrón de Alicuotas de IIBB CABA")
            On Error GoTo 0
            GoTo 99
        End If
        On Error GoTo 0
    End If

    ' CARGA EL NUEVO PADRON A UN VECTOR
    IndiceV& = -1
    NARCHI1% = FILEOPEN%(LUDAT$ + "IIBBCABA.TXT", 1, 122)
    While Not EOF(NARCHI1%)
        Line Input #NARCHI1%, RegistroPadron$
        
        IndiceV& = IndiceV& + 1
        ReDim Preserve AlicuotaPadronIBCABA(IndiceV&)
        AlicuotaPadronIBCABA(IndiceV&) = RegistroPadron$
    Wend
    Close #NARCHI1%
    
    ' ORDENAMIENTO DEL VECTOR POR EL NUMERO DE CUIT. SE DETECTO QUE EL PADRON CABA NO VIENE ORDENADO.
    ' ESTO ES INDISPENSABLE PARA LA BUSQUEDA BINARIA QUE REALIZAN LAS FUNCIONES QUE DEVUELVEN LAS ALICUOTAS
    ' DE PERCEPCION Y RETENCION.
    RegistroAuxiliar$ = ""
    For i& = 0 To IndiceV& - 1
        Call TRACE(24, i&, IndiceV& - 1)
        Cuit1# = XVALO(Mid$(AlicuotaPadronIBCABA(i&), 28, 13))
        Cuit2# = XVALO(Mid$(AlicuotaPadronIBCABA(i& + 1), 28, 13))
        
        If Cuit1# > Cuit2# Then
            RegistroAuxiliar$ = AlicuotaPadronIBCABA(i&)
            AlicuotaPadronIBCABA(i&) = AlicuotaPadronIBCABA(i& + 1)
            AlicuotaPadronIBCABA(i& + 1) = RegistroAuxiliar$
        End If
    Next i&
    
    ' ACTUALIZA EL PADRON, VOLCANDO LOS DATOS DE LOS CONTRIBUYENTES
    ' SOBRE UN ARCHIVO CON ANCHO FIJO, NECESARIO PARA LA BUSQUEDA BINARIA UTILIZADA
    ' PARA RECUPERAR LAS ALICUOTAS.
    NARCHI2% = FILEOPEN%(LUDAT$ + "PAIBCABA.TXT", 0, 61)
    Dim X As String * 61
    For i& = 0 To IndiceV&
        Call TRACE(24, i&, IndiceV&)
        X = Mid$(AlicuotaPadronIBCABA(i&), 12, 61)
        Put #NARCHI2%, i& + 1, X
    Next i&
    Close #NARCHI2%

    Call COMUNI("Ha finalizado correctamente\la actualizacion del padrón de Ingresos Brutos CABA.")

99  Screen.MousePointer = 1
    
End Sub


Private Sub ActualizaPadronIIBBTucuman_Click()

    Screen.MousePointer = 11
    
    ' VERIFICA LA EXISTENCIA DEL PADRON DE IIBB
    If EXISTE%(LUDAT$ + "IIBBTUCU.TXT") = 0 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Falta padrón de alícuotas de IIBB TUCUMAN")
        GoTo 99
    End If
    
    ' ELIMINA PADRON SIMPLIFICADO ACTUAL
    If EXISTE%(LUDAT$ + "PAIBTUCU.TXT") > 0 Then
        On Error Resume Next
        Call Kill(LUDAT$ & "PAIBTUCU.TXT")
        If Err > 0 Then
            Screen.MousePointer = 1
            Call MENSERR(24, "Error al actualizar padrón de alicuotas de IIBB TUCUMAN")
            On Error GoTo 0
            GoTo 99
        End If
        On Error GoTo 0
    End If

    ' CARGA EL NUEVO PADRON A UN VECTOR
    ConsultaSql$ = "create table #padronibtucuman(" _
    & " cuit varchar(11), datos varchar(256));"
    FLEXCONN.Execute ConsultaSql$
    
    ConsultaSql$ = "delete #padronibtucuman"
    FLEXCONN.Execute ConsultaSql$
    
    Dim rst_padron0 As New ADODB.Recordset
    ConsultaSql$ = "select * from #padronibtucuman"
    rst_padron0.Open ConsultaSql$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    
    IndiceV& = -1
    NARCHI1% = FILEOPEN%(LUDAT$ + "IIBBTUCU.TXT", 1, 198)
    FilaComienzo& = 0
    While Not EOF(NARCHI1%)
        Line Input #NARCHI1%, RegistroPadron$
        
        FilaComienzo& = FilaComienzo& + 1
        If FilaComienzo& > 7 Then ' EXCLUYE LAS PRIMERAS 7 FILAS QUE SON ENCABEZADO DEL PADRON
            IndiceV& = IndiceV& + 1
            With rst_padron0
                .AddNew
                !Cuit = Mid$(RegistroPadron$, 1, 11)
                !datos = Mid$(RegistroPadron$, 1, 19) & " " & Mid$(RegistroPadron$, 193, 4)
                .Update
            End With
        End If
    Wend
    Close #NARCHI1%
    rst_padron0.Close
    Set rst_padron0 = Nothing
    
    ' ORDENAMIENTO DEL VECTOR POR EL NUMERO DE CUIT. SE DETECTO QUE EL PADRON CABA NO VIENE ORDENADO.
    ' ESTO ES INDISPENSABLE PARA LA BUSQUEDA BINARIA QUE REALIZAN LAS FUNCIONES QUE DEVUELVEN LAS ALICUOTAS
    ' DE PERCEPCION Y RETENCION.
    Dim rst_padron1 As ADODB.Recordset
    ConsultaSql$ = "select * from #padronibtucuman with(nolock) order by cuit"
    Set rst_padron1 = READSET(ConsultaSql$)
    
    ' ACTUALIZA EL PADRON, VOLCANDO LOS DATOS DE LOS CONTRIBUYENTES
    ' SOBRE UN ARCHIVO CON ANCHO FIJO, NECESARIO PARA LA BUSQUEDA BINARIA UTILIZADA
    ' PARA RECUPERAR LAS ALICUOTAS.
    NARCHI2% = FILEOPEN%(LUDAT$ + "PAIBTUCU.TXT", 0, 32)
    i& = 0
    With rst_padron1
        Do While Not .EOF
            i& = i& + 1
            Call TRACE(24, i&, IndiceV&)
            X$ = Mid$(!datos + Space(32), 1, 30)
            Put #NARCHI2%, i& + 1, X$
        
            .MoveNext
        Loop
    End With
    Close #NARCHI2%
    rst_padron1.Close
    Set rst_padron1 = Nothing

    ConsultaSql$ = "delete #padronibtucuman"
    FLEXCONN.Execute ConsultaSql$

    Call COMUNI("Ha finalizado correctamente\la actualizacion del padrón de Ingresos Brutos TUCUMAN.")

99  Screen.MousePointer = 1

End Sub


Private Sub AjuSaldAc_Click()
    Call OPORPA07.Command4_Click
End Sub

Private Sub AnuOps_Click()
   Call OPORPA07.Command25_Click
End Sub

Private Sub BOTNEXT_Click()
      '
    If XVALO(Label18.Caption) < 0.005 Then
       Exit Sub
    End If
    '
    FF% = FECHANUM(Text9)
    FECHALIM% = XVALO(Control$("", "CIEMECOM"))
    If FF% <= FECHALIM% Then
      Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
      Text9.SetFocus
      Exit Sub
    End If
     '
    If XVALO(Label20.Caption) < 0 Then
         Call MENSERR(24, "Valores Cargados exceden Créditos")
         Exit Sub
       ElseIf XVALO(Label20.Caption) >= 0.005 Then
         Call CIERRARCH("!CARMEPA")
         Call BLANARCH("!CARMEPA")
         X$ = REGBLAN$("CARMEPA")
         HOI% = HOY(HO$)
         Call REPLA(X$, MKD$(XVALO(Label20.Caption)), 31, 8)
         Call REPLA(X$, MKI$(HOI%), 39, 2)
         Call REPLA(X$, MKI$(HOI%), 41, 2)
         Call REGAPP("!CARMEPA", X$)
         '
10       VTB% = VENTABU%("CARMEPA")
         If VTB% < 0 Then
            Call CIERRARCH("!CARMEPA")
            Call BLANARCH("!CARMEPA")
            Exit Sub
         End If
         SUVAPROP# = 0
         For U& = 1 To ULTREG&("!CARMEPA")
            X$ = REGLEIDO$("!CARMEPA", U&)
            IVAL# = CVD(Mid$(X$, 31, 8))
            If IVAL# < 0.005 Then
               Call MENSERR(24, "No se Admiten Valores Negativos")
               GoTo 10
            End If
            '
            '\\\OT-06-0570-RG-27/12/06///
            FECHEMIS% = CVI(Mid$(X$, 39, 2))
            FechVenc% = CVI(Mid$(X$, 41, 2))
            If FechVenc% < FECHEMIS% Then
               Call MENSERR(24, "La Fecha de Vencimiento Debe Ser Mayor a La Fecha De Emisión")
               GoTo 10
            End If
            '\\\OT-06-0570-RG-FIN///
            '
            '\\\OT-06-0570-RG-27/12/06///
            BANC% = CVI(Left$(X$, 2))
            NUCHE$ = TRIM$(Mid$(X$, 19, 12))
            '\\\OT-06-0570-RG-FIN///
            '
            '\\\OT-08-0662-OD-11/08/08///
            'Valida que el medio de pago no este como no vigente ('N/V')
            MEDIOPAGO$ = TRIM$(ECOARCH$("MEPAGO", MKI$(BANC%)))
            If InStr(1, MEDIOPAGO$, "N/V") > 0 Then
              Call MENSERR(24, "Medio de Pago " + "'" + MEDIOPAGO$ + "'" + " No Vigente ")
              GoTo 10
            End If
            '
            '\\\OT-08-0662-OD-FIN///
            '
            If CVI(Left$(X$, 2)) < 1 Or ECOARCH$("MEPAGO", Left$(X$, 2)) = "" Then
               Call MENSERR(24, "Tipo de Valor no Admitido")
               GoTo 10
            End If
            '
            ' SE MODIFICO POR PEDIDO DE TANIT QUE EN EL CASO QUE HAYA CHEQUE INGRESADO ANTERIORMENTE
            ' SOLO AVISE Y PERMITA CONTINUAR.
            'VALIDACION DE CHEQUE ENTREGADO
            If YAESTACHEQUE_OP%(MEDIOPAGO$, NUCHE$, opago$) > 0 Then
               If COMALTER%("Cheque Nro.: " + NUCHE$ + " Banco: '" + MEDIOPAGO$ + "' Ya Utilizado En '" & opago$ & "'\\Corregir\Continuar") < 2 Then GoTo 10
            End If
            '
            'Verifica si el cheque esta repetido
            If NUCHE$ <> "" Then 'validacion para que no valide medios de pago sin numero como efectivo
                For i& = 1 To (U& - 1)
                    X1$ = REGLEIDO$("!CARMEPA", i&)
                    If Mid$(X1$, 1, 30) = Mid$(X$, 1, 30) Then
                      Call MENSERR(24, "Cheque N.: " + NUCHE$ + "  Banco : " + MEDIOPAGO$ + " Repetido")
                      GoTo 10
                    End If
                Next i&
            End If

            SUVAPROP# = SUVAPROP# + IVAL#
         Next U&
         If Abs(SUVAPROP# - XVALO(Label20.Caption)) >= 0.005 Then
            Call MENSERR(24, "Imposible Cerrar - Diferencia de Valores")
            GoTo 10
         End If
         '
         Call CIERRARCH("!CARMEPA")
         Call SIGUEORPA4
         '
       Else
         '
         Call CIERRARCH("!CARMEPA")
         Call BLANARCH("!CARMEPA")
         Call SIGUEORPA4
         '
    End If
    '
End Sub
'

Private Sub Command1_Click()
    '
    'Call SELECHO("PROVED1")
    'NCXI% = XVALO(TRIM$(VALACT1$("PROVED1")))
    NCXI% = ListaSeleccionProveedores()
    If NCXI% > 0 Then
      '\\\LIMPIAR TOTALES -OT-08-0889-OD-25/09/08///
        Call LimpTotOrPgo
      '\\\OT-08-0889-OD-FIN///
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = rasocli$((-1) * NC%)
        Text3.TEXT = DOMICLI$((-1) * NC%)
        Text4.TEXT = CPOSCLI$((-1) * NC%)
        Text5.TEXT = LOCACLI$((-1) * NC%)
        Text7.TEXT = ORDECHEQ$((-1) * NC%)
        Label26 = CODICLI((-1) * NC%)
          Text7.ForeColor = vbBlack
          If Text7 <> rasocli(NC% * (-1)) Then Text7.ForeColor = vbRed
        Call VERIRECHA(NCXI%)
        
        ' *** CONFIGURA RETENCION IB MISIONES
        RetieneIBMisiones% = XVALO(VALOBADA("PROVED12", "CategoriaIBMISIONES_Cli", "NUME_CLI=" & CStr(NC%)))
        If RetieneIBMisiones% > 0 Then ' RetieneIBMisiones% tiene la categoria.
            PorcentajeRetencionIBMisiones# = ALIRETBRU_MISIONES#(XVALO(Text1), MinimoRetencionIBMisiones#, MedioPagoIBMisiones%)
        End If
        
        Call SIGUEORPA1
        
    End If
    '
End Sub


Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("ARCHIG07/AMAPRO", "Archivos Maestros")
    Command11.Enabled = True
End Sub

Private Sub Command2_Click()
    Screen.MousePointer = 11
    If List1.ListCount > 0 Or List2.ListCount > 0 Then
          Call CIERRARCH("*.*")
          Call BLANARCH("!CARGAST")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label18.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label45.Caption = ""
          Label12.Caption = ""
          Label38.Caption = ""
          Label49.Caption = ""
          Label30.Caption = ""
          Label20.Caption = ""
          Label21.Caption = ""
          BOTNEXT.Enabled = False
          Command5.Enabled = False
          Screen.MousePointer = 1
          EPAC$ = UCase$(EMPREAC$)
          If InStr(EPAC$, "NEUMAN") > 0 Then
            Text1.TEXT = ""
            Text2.TEXT = ""
            Text3.TEXT = ""
            Text4.TEXT = ""
            Text5.TEXT = ""
          End If
          Text1.SetFocus
        Else
          If TRIM$(Control$("ORPASTA", "AVAHOJA")) = "NEWPAGE" Then
              Call SETSPOOL("ENDDOC", "")
          End If
          Hide
          Screen.MousePointer = 1
    End If
    'Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
   If List1.ListCount > 0 Or List2.ListCount > 0 Then
        Command2.ToolTipText = "Limpia y Refresca Pantalla"
      Else
        Command2.ToolTipText = "Cierra y Abandona la Aplicación"
   End If
End Sub




Private Sub Command20_Click()
    
    FORIPRE$ = TRIM$(Control$("", "FORPASTA"))
        
    ConsultaSql$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE CODICOM_LAR LIKE 'OP-X-9999%' AND TIPOMOVIM='OPAG' ORDER BY CODICOM_LAR"
    Dim RST As ADODB.Recordset
    Set RST = READSET(ConsultaSql$)
    
    With RST
        Do While Not .EOF
            '
            NUMEORPA$ = SAFETEXT$(!CODICOM_LAR)
            Set RST0 = READSET("SELECT * FROM CAJABANC WITH(NOLOCK) WHERE CODICOM_LAR='" & NUMEORPA$ & "' ORDER BY TIPOMOVIM")
            'MsgBox NUMEORPA$
            With RST0
                Do While Not .EOF
                    TIMO$ = SAFETEXT$(!TipoMovim)
                    If TIMO$ = "RTIB" Then
                        TIPOS$ = "R.IB"
                        BANCO$ = "Ret.IB CABA"
                        NUVA$ = SAFETEXT$(!NroDeRetencion)
                        IMVA# = XVALO(!IHABECOMP)
                        '
                        Z$ = REGBLAN$("!PRINRECI")
                        Call REPLA(Z$, TIPOS$, 1, 4)
                        Call REPLA(Z$, NUVA$, 5, 12)
                        Call REPLA(Z$, BANCO$, 17, 16)
                        Call REPLA(Z$, MKD$(IMVA#), 33, 8)
                        FEX = CVINT(!fechconta)
                        If FEX > 0 Then
                            Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
                        End If
                        Call REGAPP("!PRINRECI", Z$)
                        'MsgBox "paso retencion"
                    End If
                    
                    If TIMO$ = "OPAG" Then
                        Z$ = REGBLAN$("!PRINRECI")
                        DOCU$ = SAFETEXT$(!Notas)
                        CANCE# = XVALO(!Idebecomp)
                        SALPE# = 0
                        Call REPLA(Z$, DOCU$, 53, 18)
                        Call REPLA(Z$, MKD$(CANCE#), 71, 8)
                        Call REPLA(Z$, MKD$(SALPE#), 79, 8)
                        Call REGAPP("!PRINRECI", Z$)
                        'MsgBox "paso factura"
                    End If
                    .MoveNext
                Loop
            End With
            RST0.Close
            Set RST0 = Nothing
            '
            Call CIERRARCH("!PRINRECI")

            'MsgBox "impresion 1"
            CANCELPRINT% = 1
            '
            CODFOR$ = TRIM$(UCase$(FORIPRE$))
            MARIZW = XVALO(ATRIFORM$(CODFOR$, "MARIZQMM"))
            MARSUW = XVALO(ATRIFORM$(CODFOR$, "MARSUPMM"))
            '
            INTERLI = XVALO(ATRIFORM$(CODFOR$, "INTERLIN"))
            If INTERLI < 1 Then
                Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
                INTERLI = 1
            End If
            '
            DESMM = XVALO(ATRIFORM$(CODFOR$, "DETDESMM"))
            HASMM = XVALO(ATRIFORM$(CODFOR$, "DETHASMM"))
            '
            'MsgBox "impresion 2"
            CACOPI% = XVALO(ATRIFORM$(CODFOR$, "CAN-COP"))
            If CACOPI% < 1 Or CACOPI% > 6 Then
                Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
                CACOPI% = 1
            End If
            '
            ORIENTA% = 1
            PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
            If PORI$ = "PAISAJE" Or PORI$ = "2" Then
                ORIENTA% = 2
            End If
            '
            PORIAN = Printer.Orientation  ' retiene orientacion
            '
            PRINTERCONTROL% = 1
            '
            'MsgBox "impresion 3"
            SUMAVA# = 0: SUMACAN# = SUMAVA#
            Call CARFORWIN
            '
            NC% = XVALO(!NUPROVE)
            NC1% = NC% * (-1)
            Call PRINTFORWIN(TRIM$(Str$(NC%)), "CUEN-CLI", 0, 0)
            Call PRINTFORWIN(TRIM$(rasocli$(NC1%)), "RASO-CLI", 0, 0)
            Call PRINTFORWIN(DOMICLI$(NC1%), "DOMI-CLI", 0, 0)
            Call PRINTFORWIN(CPOSCLI$(NC1%), "CPOS-CLI", 0, 0)
            Call PRINTFORWIN(LOCACLI$(NC1%), "LOCA-CLI", 0, 0)
            Call PRINTFORWIN(CuitCli$(NC1%), "CUIT-CLI", 0, 0)
            Call PRINTFORWIN(POSIVAL$(NC1%), "PIVA-CLI", 0, 0)
            Call PRINTFORWIN(TEL1CLI$(NC1%), "TELE-CLI", 0, 0)
                FF% = CVINT(!fechconta)
            Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
                NRORPA$ = Str(XVALO(Mid$(!CODICOM_LAR, 11)))
            Call PRINTFORWIN(NRORPA$, "NUME-COM", 0, 0)
                MONEMI% = 1
            Call PRINTFORWIN(DEMONECO$(MONEMI%), "MONEDEMI", 0, 0)
            '
            'MsgBox "impresion 4 - " & ULTREG&("!PRINRECI")
            YAC = DESMM + INTERLI
            For YX& = 1 To ULTREG&("!PRINRECI")
               '
               'YAC = YAC + INTERLI
               '
               If YAC > HASMM - 2 * INTERLI Then
                   YAC = YAC + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   '
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-VAL", 0, 0)
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-CAN", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET1", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET2", 0, 0)
                   '
                   Call CARFORWIN
                   Call PRINTFORWIN(TRIM$(Str$(NC1%)), "CUEN-CLI", 0, 0)
                   Call PRINTFORWIN(TRIM$(rasocli$(NC1%)), "RASO-CLI", 0, 0)
                   Call PRINTFORWIN(DOMICLI$(NC1%), "DOMI-CLI", 0, 0)
                   Call PRINTFORWIN(CPOSCLI$(NC1%), "CPOS-CLI", 0, 0)
                   Call PRINTFORWIN(LOCACLI$(NC1%), "LOCA-CLI", 0, 0)
                   Call PRINTFORWIN(CuitCli$(NC1%), "CUIT-CLI", 0, 0)
                   Call PRINTFORWIN(POSIVAL$(NC1%), "PIVA-CLI", 0, 0)
                   Call PRINTFORWIN(TEL1CLI$(NC1%), "TELE-CLI", 0, 0)
                   FEX = FF%
                   Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
                   Call PRINTFORWIN(NRORPA$, "NUME-COM", 0, 0)
                   '
                   YAC = DESMM + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   '
                   YAC = YAC + 2 * INTERLI
                   '
                End If
                '
                YY$ = REGLEIDO$("!PRINRECI", YX&)
                'MsgBox YY$
                Call PRINTFORWIN(Left$(YY$, 4), "TIPO-VAL", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 5, 12), "NUME-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 17, 16), "BANC-CHE", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 33, 8), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                SUMAVA# = SUMAVA# + CVD(Mid$(YY$, 33, 8))
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 41, 8), 6, 8, 0, 1), "FECH-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 53, 18), "DOCU-APL", 0, YAC)
                '
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 71, 8), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                SUMACAN# = SUMACAN# + CVD(Mid$(YY$, 71, 8))
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 79, 8), 4, 12, 2, 2), "IMPO-PEN", 0, YAC)
                '
            Next YX&
            '
            'MsgBox "impresion 5"
33          ITEXTO$ = NUMTEX$(SUMAVA#) + " ctvs"
            LOTEX% = XVALO(ATRIFORM$(CODFOR$, "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
            Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
            IMPLET1$ = "": IMPLET2$ = "": IMPLET3$ = ""
            If TRIM$(NRX1$) <> "" Then
               IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
            End If
            If TRIM$(NRX2$) <> "" Then
               IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
            End If
            If TRIM$(NRX3$) <> "" Then
               IMPLET3$ = TRIM$(NRX3$) + " " + "***************************************************************************************************************"
            End If
            
            Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "TOTA-VAL", 0, 0)
            Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "TOTA-CAN", 0, 0)
            Call PRINTFORWIN(IMPLET1$, "IMP-LET1", 0, 0)
            Call PRINTFORWIN(IMPLET2$, "IMP-LET2", 0, 0)
            Call PRINTFORWIN(USERNAME, "USUARIO", 0, 0)
            '
            Call PRINTFORWIN("CORRESPONDE A LA RETENCION DE INGRESOS BRUTOS C.A.B.A.", "OBSERVA", 0, YAC)
            '
            Call SETULTREG("!PRINRECI", 0)
            Call BLANARCH("!PRINRECI")
            'If AVAHOJA$ <> "NEWPAGE" Then
                 'Printer.EndDoc
                 'MsgBox "impresion 6"
                 Call SETSPOOL("ENDDOC", "")
            'End If
            '
            '
            'MsgBox "impresion 7"
            Call CIERRARCH("*.*")
            DESCRIDOC$ = "Orden de Pago Nro. " + TRIM$(NRORPA$) + " - " + TRIM$(rasocli$(NC1%))
            TIPODOC$ = "Orden de Pago"
            NUMEDOC$ = NRORPA$
            FEX = FF%
            DESCRIDOC$ = DESCRIDOC$ + " - " + FECHATEX$(FEX)
            '
            'MsgBox "impresion 8"
            Call SAVESPOOL(DESCRIDOC$, OKX%)
            Screen.MousePointer = 1

            'If COMALTER%("Orden de Pago Numero\'" + NRORPA$ + "'\Emitida y Registrada\\Salir\Seguir") = 1 Then GoTo 30
            '
            'If AVAHOJA$ <> "NEWPAGE" Then
                Printer.Orientation = PORIAN
            'End If
            
            .MoveNext
        Loop
    End With
30  Call CIERRARCH("*.*")
'
Call COMUNI("PROCESO FINALIZADO")
    '
End Sub

Private Sub Command25_Click()
    Call OPORPA07.Command25_Click
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Text1.TEXT)
    Call CONECRUN("*NOTAPRO/" + NCLI$, "Notas de Proveedores")
    Command3.Enabled = True
End Sub

Sub Command4_Click()
    Command4.Enabled = False
    'If EXISTE%("COGECOM.EXE") < 1 Then
    '    Call OPNOIN("")
    '    Command4.Enabled = True
    '    Exit Sub
    'End If
    RECUEP07.Text1 = Text1
    'PROLLAMA$ = Name
    'Hide
    RECUEP07.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    '
'MsgBox 1
    'validacion que coincida el improte de la grilla con el total aplicado
    NC% = XVALO(Text1)
    If TICUEPRO%((-1) * NC%) = 1 Then 'caso gastos
        TOTAACUM# = 0
        V& = 0
        FIN& = ULTREG&("!CARGAST")
        For U& = 1 To FIN&
           X$ = REGLEIDO$("!CARGAST", U&)
           If ECOARCH$("ECUECON", Left$(X$, 12)) <> "" Then
            TOTAACUM# = TOTAACUM# + CVD(Mid$(X$, 13, 8))
           End If
        Next U&
        TOTALAPLI# = XVALO(Label21) + XVALO(Label45.Caption) + XVALO(Label12.Caption) + XVALO(Label14) + XVALO(Label30) + XVALO(Label38) + XVALO(Label49)
        If ROUND(TOTALAPLI#, 4) <> ROUND(TOTAACUM#, 4) Then
           Call MENSERR(24, "Importe Valor Difiere de Saldo a Cancelar")
           Command5.Enabled = True
           Exit Sub
        End If
    End If
    
    FF% = FECHANUM(Text9)
    FECHALIM% = XVALO(Control$("", "CIEMECOM"))
    If FF% <= FECHALIM% Then
      Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
      Text9.SetFocus
      Command5.Enabled = True
      Exit Sub
    End If
     '
    If XVALO(Label21) < 0.005 Then
       Call MENSERR(24, "No Valido Emitir Orden de Pago\sin Ingresar Valores.")
       Command5.Enabled = True
       Exit Sub
    End If
    '
    TOTAVAL# = CDbl(XVALO(Label30.Caption)) + CDbl(XVALO(Label14.Caption)) + CDbl(XVALO(Label45.Caption)) + CDbl(XVALO(Label12.Caption)) + CDbl(XVALO(Label21.Caption)) + CDbl(XVALO(Label38.Caption)) + CDbl(XVALO(Label49.Caption))
    TOTACAN# = CDbl(XVALO(Label18.Caption))
    If Abs(TOTACAN# - TOTAVAL#) >= 0.005 Then
        Call COMUNI("Operación Incorrecta.\  \Las Cancelaciones Difieren\de la Suma de Valores a Entregar.")
        Command5.Enabled = True
        Exit Sub
    End If
    '
    ' *** VALIDA QUE EXISTE EL MEDIO DE PAGO PARA RETENCION DE IVA
    RetencionTotal# = XVALO(Label49)
    If RetencionTotal# > 0.005 And MedioPagoRetencionIva% = 0 Then
        Call MENSERR(24, "Imposible Registrar la Orden de Pago.\ \Falta Cuenta Contable\Correspondiente a Retención de IVA")
        GoTo 10
    End If
    
    VACON% = 1       ' PARA QUE NO BLANQUEE LISTADOS
    NiveCla% = 0
    If Picture6.Visible = True Then NiveCla% = XVALO(Mid$(Label40, 2))
    '
    Call GRAORPASTA(OKX%)
    Call LIBARCH("CUECORR")
    If OKX% < 0 Then GoTo 10   ' abortó en observaciones
    If OKX% = 0 Then           ' RE-NUMERO ORDEN DE PAGO O NO VERIFICO CHEQUES
       Screen.MousePointer = 1
       Command5.Enabled = True
       VACON% = 0
       Exit Sub
    End If
    '
    If OKX% <> 7 Then
        Call COMUNI("La Registración de esta Orden de Pago\no Pudo Completarse. (" + TRIM$(Str$(OKX%)) + ")\  \Consulte a su Soporte de Sistemas FLEXOFT.")
        Call FINAL("")
    End If
    
    ' *** IMPRESION DE ORDEN DE PAGO
    Call PRIORPASTA
    ' *** IMPRESION RETENCIONES DE IIBB
    Call PRICORETIB
    ' *** IMPRESION RETENCION DE GANANCIAS
    Call PRICORETGA
    ' *** IMPRESION RETENCION DE IVA
    Call PRICORETIVA
    
    VACON% = 0
    '
10  FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    NROPRO& = 0
    Call NUORPA(NROPRO&)
    Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
    '
    Call GRAORDECHPRO   'PARA QUE GUARDE LA ORDEN DEL CHEQUE
    '
    Call CIERRARCH("!CARMEPA")
    Call BLANARCH("!CARMEPA")
    Call CIERRARCH("!CARGAPLI")
    Call BLANARCH("!CARGAPLI")
    Call CIERRARCH("!CARGAST")
    Call BLANARCH("!CARGAST")
    List1.Clear
    List2.Clear
    'Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
    Label4.Caption = ""
    Label8.Caption = ""
    Label18.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label45.Caption = ""
    Label12.Caption = ""
    Label38.Caption = ""
    Label49.Caption = ""
    Label30.Caption = ""
    Label20.Caption = ""
    Label21.Caption = ""
    'BOTNEXT.Enabled = True
    Command5.Enabled = False
    EPAC$ = UCase$(EMPREAC$)
    If InStr(EPAC$, "NEUMAN") > 0 Then
      Text1.TEXT = ""
      Text2.TEXT = ""
      Text3.TEXT = ""
      Text4.TEXT = ""
      Text5.TEXT = ""
    End If
    MONEMI% = 0
    Call MONEMIS_DblClick
    Text1.SetFocus
    '
End Sub

Sub Command6_Click()
    Command6.Enabled = False
    OPORPA07.Show 1
    '
    NROPRO& = 0
    Call NUORPA(NROPRO&)
    Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
    '
    Command6.Enabled = True
End Sub

Sub Command7_Click()
    Command7.Enabled = False
    If DERACCE%("STPORPAG", "Configuración de Funcionamiento de Ordenes de Pago") = 2 Then
        STORPA07.Show 1
        Call COMUNI("Re-Inicie Ahora la Aplicación.")
        Call FINAL("")
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   If List1.ListCount > 0 Then
     Screen.MousePointer = 11
     MODOSORVAL% = (MODOSORVAL% + 1) Mod 2
     List4.Clear
     '
     For U& = 1 To List1.ListCount
       Z1$ = List1.LIST(U& - 1)
       VALI1 = FECHANUM(Mid$(Z1$, 49, 8))
       If MODOSORVAL% = 1 Then VALI1 = XVALO(Mid$(Z1$, 57, 12))
       For U4& = 1 To List4.ListCount
         Z4$ = List4.LIST(U4& - 1)
         VALI4 = FECHANUM(Mid$(Z4$, 49, 8))
         If MODOSORVAL% = 1 Then VALI4 = XVALO(Mid$(Z4$, 57, 12))
         If VALI4 > VALI1 Then
           List4.AddItem Z1$, U4& - 1
           GoTo 19
         End If
       Next U4&
       List4.AddItem Z1$
19   Next U&
     '
     List1.Clear
     For U& = 1 To List4.ListCount
       List1.AddItem List4.LIST(U& - 1)
     Next U&
     List1.Refresh
     'Call REMPLAGRID(Me, GRID2, Picture1, Label12, List1)
     Screen.MousePointer = 1
     '
   End If
   Command8.Enabled = True
   '
End Sub

Private Sub Command9_Click()
    If DERACCE%("CAMFEOP", "Cambio de Fecha de Orden de Pago") >= 1 Then
        Call SELECHO("SELFECHA")
        VDEVV$ = VALACT1$("SELFECHA")
        If VDEVV$ <> "" Then
            Text9.TEXT = VDEVV$
            FF% = FECHANUM(Text9.TEXT)
        End If
    End If
End Sub

Private Sub ControlDeEstadoValoresEntregados_Click()
    '
    '  ESTA RUTINA EVALUA SI EL ESTADO DE UN CHEQUE ES CORRECTO.
    '  1. SI EL CHEQUE FUE ENTREGADO COMO PARTE DE PAGO, VERIFICA SI LA O/P EXISTE. CASO CONTRARIO,
    '     HABILITA LA POSIBILIDAD DE REVERTIR EL ESTADO Y DEJARLO NUEVAMENTE DISPONIBLE EN LA CARTERA DE VALORES
    '
    Call DESHASFECHA(Des%, Has%, PERIO$)
    
    If PERIO$ = "" Then Exit Sub
    Desde$ = FETEXCOR1$(Des%)
    Hasta$ = FETEXCOR1$(Has%)
     
10  Call CARGA_ENCABEZADO(CTRLCHEQENTRE.GRID, "Recibo/A18;Cheque/A36;Importe/F12.2;O-Pago/A18", CTRLCHEQENTRE)
       
    Dim RstCheques As ADODB.Recordset
    ConsultaSql$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE TIPOMOVIM='CHRECI' AND STATUS=2 AND OPCIMOVIM='CH'"
    ConsultaSql$ = ConsultaSql$ & " AND FECHCONTA>=' " & Desde$ & "' AND FECHCONTA <='" & Hasta$ & "'"
    
    Set RstCheques = READSET(ConsultaSql$)

    FilaGrid& = 0
    With RstCheques
        Do While Not .EOF
            DocumentoSalida$ = SAFETEXT$(!Document)
            If Mid$(TRIM$(DocumentoSalida$), 1, 2) = "OP" Then
                ImporteCheque$ = TRIM$(FORMATNUM$(XVALO(!Idebecomp), "F12.2"))
                If CantRegistros("CAJABANC", "CODICOM_COR='" & DocumentoSalida$ & "' AND ABS(IHABECOMP - " & ImporteCheque$ & ") < 0.005") = 0 Then
                    CTRLCHEQENTRE.GRID.Rows = MUESTRADATOS1.GRID.Rows + 1
                    FilaGrid& = FilaGrid& + 1
                    CTRLCHEQENTRE.GRID.TextMatrix(FilaGrid&, 1) = SAFETEXT$(!CODICOM_LAR)
                    CTRLCHEQENTRE.GRID.TextMatrix(FilaGrid&, 2) = SAFETEXT$(!DESCRIMOV)
                    CTRLCHEQENTRE.GRID.TextMatrix(FilaGrid&, 3) = FORMATNUM$(XVALO(!ValAbsComp), "F12.2")
                    CTRLCHEQENTRE.GRID.TextMatrix(FilaGrid&, 4) = SAFETEXT$(!Document)
                End If
            End If
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    RstCheques.Close
    Set RstCheques = Nothing
    On Error GoTo 0
    
    If FilaGrid& > 0 Then
        CTRLCHEQENTRE.GRID.Rows = FilaGrid& + 1
        CTRLCHEQENTRE.Show 1
    Else
        Call COMUNI("No Se Han Encontrado Cheques\Con Estado Erróneo.")
    End If
    
End Sub

Private Sub Form_Activate()
   '
    'AGREGADO POR MULTIMONEDA PARA QUE NO VUELVA A PESOS CADA VEZ QUE VUELVE FORM_ACTIVATE
    If YACARFO% > 0 And MODOMONEI% = 2 Then Exit Sub

    MONEMI% = 1
    'MONEMIS = DEMONECO$(1)
    TICAMBIS = "1.0000"
    TICAMBX# = 1
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESACRED", "MODOMONE"))), 8)
    '
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
       Frame12.Visible = True
    End If

   If YACARFO% < 1 Then
      '
      DoEvents
      YACARFO% = 1
      '
10    FORIPRE$ = TRIM$(Control$("", "FORPASTA"))
      If FORIPRE$ = "" Then
          Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Orden de Pago.")
          STORPA07.Show 1
          List1.Clear
          List2.Clear
          'Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
          Label4.Caption = ""
          Label8.Caption = ""
          Label18.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label45.Caption = ""
          Label12.Caption = ""
          Label38.Caption = ""
          Label49.Caption = ""
          Label30.Caption = ""
          Label20.Caption = ""
          Command5.Enabled = False
          GoTo 10
      End If
      '
      If FORIPRE$ <> "" Then
          LU$ = LUDAT$
          If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
              RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
              Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Pago Standard:\  \'" + RMCC$ + "'.")
              STORPA07.Show 1
              GoTo 10
          End If
      End If
      '
      FF% = HOY(HO$)
      Text9.TEXT = HO$
      '
      NROPRO& = 0
      Call NUORPA(NROPRO&)
      Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
      '
   End If
   
End Sub



Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If NIVCONFI% > 0 Then
       Option1.ToolTipText = "Doble Click para Clasificar"
    End If
    '
    Call CENTRAFORM(Me)
    Refresh
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture3)
    
    If InStr(1, EPAC$, "MEDITEA") > 0 Then
        mnuAplicaReparacionesAPagos.Visible = True
    End If
    
'    Call AplicaFuenteAnchoFijo(List1, , 10):  Call AplicaFuenteAnchoFijo(Picture1, , 10)
'    Call AplicaFuenteAnchoFijo(List2, , 10):  Call AplicaFuenteAnchoFijo(Picture2, , 10)
'    Call AplicaFuenteAnchoFijo(Label8, , 10)
'    Call AplicaFuenteAnchoFijo(Label18, , 10)
'    Call AplicaFuenteAnchoFijo(Label21, , 10)
'    Call AplicaFuenteAnchoFijo(Label14, , 10)
'    Call AplicaFuenteAnchoFijo(Label38, , 10)
'    Call AplicaFuenteAnchoFijo(Label30, , 10)
'    Call AplicaFuenteAnchoFijo(Label20, , 10)
    '
    'SE COMENTA ESTA LLAMADA POR HABERSE AGREGADO EN RECIBOS LA VALIDACION
    'DE INGRESO DE CUIT DEL LIBRADOR
    'Call CUITLIBCHE ' PROCESO DE RECUPERACION DE CUIT DE CHEQUES
    DoEvents
    '
    If RECLEN%("PRINRECI") <> 256 Then           'VERIFICA LONGITUD DEL ARCHIVO PRINRECI
       Call MENSERR(24, "Longitud del Archivo PRINRECI.RFS Menor a la Longitud Necesaria (256).\Debe Actualizar el Archivo.\Contactarse con Mesa de Ayuda Para su Actualización.")
    End If
    ' Carpetas de Importacion
    CODCARPE$ = BUSCAVALOR(PARAPLI$, "CODCARPE =") ' codigo de carpeta de import / export
    If CODCARPE$ <> "" Then
        Condi$ = "CodcarpImp ='" & CODCARPE$ & "' AND STATUS < 9"
        If CantRegistros("CARPEIMP", Condi$) > 0 Then
           CCOS$ = VALOBADA("CARPEIMP", "CodCenCos", Condi$, "NOMESS")
           Label46 = CCOS$
           Label44 = CODCARPE$
        End If
    End If
    'Call REMPLAGRID(Me, GRID2, Picture1, Label12, List1)
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESACRED", "MODOMONE"))), 8)

    If InStr(1, EMPREAC$, "FERVI") > 0 Or InStr(1, EMPREAC$, "AMBEST") > 0 Or InStr(1, EMPREAC$, "AHP") > 0 Then
        Option2.Visible = True
    End If
    
    ' *** RETENCION DE IVA - PARAMETROS DE CONFIGURACION
    MinimoImponibleRetencionIva# = XVALO(Control$("RETENIVA", "MINIMPO"))
    MedioPagoRetencionIva% = CUEINT%(TRIM$(Control$("RETENIVA", "CUECONTA")))
    
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
       Frame8.Visible = True
       GoTo 90
    End If

    'Me.Height = 8745
    
    ' *** HABILITA RETENCION DE IIBB MISIONES SI ESTA CONFIGURADO EN EL SETUO DE O/P
    If Control$("", "RETIBMI") = "SI" Then
        Label47.Visible = True
        Label45.Visible = True
    End If
    
    ' *** HABILITA RETENCION DE IIBB TUCUMAN SI ESTA CONFIGURADO EN EL SETUO DE O/P
    If Control$("", "RETIBTU") = "SI" Then
        Label9.Visible = True
        Label12.Visible = True
    End If
    
90
    '
End Sub
'
Sub CUITLIBCHE()
     '
     ' SE AGREGA VARIABLE STATIC PARA QUE NO INGRESE CADA VEZ QUE SE ABRE EL FRM
     Static YAPASO%
     If YAPASO% > 0 Then Exit Sub
     YAPASO% = 1
     '
     Screen.MousePointer = 11
     
     CUECON$ = Control$("", "CUECHECA")
     CUECONTA% = CUEINT%(CUECON$)
     '
     Call ABRECABAN
     SQLX$ = "SELECT * FROM CAJABANC "
     SQLX$ = SQLX$ & " WHERE "
     Condi$ = " TIPOMOVIM = 'CHRECI'"
     Condi$ = Condi$ & " AND (OBSERGEN NOT  LIKE 'CUIT%' OR OBSERGEN IS  NULL)"
     Condi$ = Condi$ & " AND CUECONTAI = " & CUECONTA%
     Condi$ = Condi$ & " AND NUCLIEN > 0 "
     SQLX$ = SQLX$ + Condi$
     '
     FIN& = CantRegistros&("CAJABANC", Condi$)
     '
     Dim CAJTMP As ADODB.Recordset
     Set CAJTMP = New ADODB.Recordset
25   On Error Resume Next
     CAJTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 25
     End If
     On Error GoTo 0
    With CAJTMP
        Do While Not .EOF
           Call TRACE(20, i&, FIN&)
           i& = i& + 1
           NC% = XVALO(!NUCLIEN)
           CUITLIB$ = CuitCli$(NC%)
           If TRIM$(CUITLIB$) = "" Then
              GoTo 50
           End If
           !OBSERGEN = "CUIT Libr: " & CUITLIB$
           .Update
50         .MoveNext
        Loop
    End With
    CAJTMP.Close
    Set CAJTMP = Nothing
    Screen.MousePointer = 1
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If TRIM$(Control$("ORPASTA", "AVAHOJA")) = "NEWPAGE" Then
       Call SETSPOOL("ENDDOC", "")
    End If
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    YACARFO% = 0
End Sub



Private Sub GRID1_Click()
    If GRID1.Row > 0 Then
       On Error Resume Next
       List2.ListIndex = GRID1.Row - 1
       Call List2_Click
       On Error GoTo 0
    End If
End Sub

Private Sub GRID1_DblClick()
    If GRID1.Row > 0 Then
       On Error Resume Next
       List2.ListIndex = GRID1.Row - 1
       Call List2_DblClick
       On Error GoTo 0
    End If
End Sub


Private Sub GRID2_Click()
    If GRID2.Row > 0 Then
       On Error Resume Next
       List1.ListIndex = GRID2.Row - 1
       Call LIST1_Click
       On Error GoTo 0
    End If
End Sub

Private Sub GRID2_DblClick()
    If GRID2.Row > 0 Then
       On Error Resume Next
       List1.ListIndex = GRID2.Row - 1
       Call List1_DblClick
       ' Call REMPLAGRID(Me, GRID2, Picture3, Label27, List1)
    End If
End Sub

Private Sub GRID2_RowColChange()
    If SUPRIMECLICK% <> 1 Then
       Call GRID2_Click
    End If
End Sub

Private Sub impucredi_Click()
    Call OPORPA07.Command3_Click
End Sub

Private Sub Label10_Change()
   Call CALRETIBRU
   Call CALRETIBRU_MISIONES
   Call CALRETIBRU_TUCUMAN
   If XVALO(Label10.Caption) >= 0.005 Then
        BOTNEXT.Enabled = True
      Else
        BOTNEXT.Enabled = False
   End If
End Sub

Private Sub Label12_Change()
   Label20.Caption = TRIM$(CSTRING$(MKD$(XVALO(Label18.Caption) - XVALO(Label30.Caption) - XVALO(Label14.Caption) - XVALO(Label45.Caption) - XVALO(Label12.Caption) - XVALO(Label21.Caption) - XVALO(Label38) - XVALO(Label49)), 4, 14, 2, 2))
   Label20.Refresh
End Sub

Private Sub Label12_DblClick()
   If XVALO(Label18) > 0 Then
     RetencionRecalculado# = XVALO(Label12)
     Porcentaje# = 0
     If MontoSujetoIBTucuman# > 0.005 Then Porcentaje# = (RetencionRecalculado# / MontoSujetoIBTucuman#) * 100
     VDEF$ = MKD$(MontoSujetoIBTucuman#) + MKS$(Porcentaje#) + MKS$(100)
10   BASIMPON# = CVD(Left$(VDEF$, 8))
     IRET# = XVALO(Label12)
     VDEF$ = VDEF$ + MKD$(RetencionRecalculado#)
     OBX$ = OBJPLA$("CARIRETIB", VDEF$)
     If OBX$ = "" Then Exit Sub
     
     IRET1# = CVD(Left$(OBX$, 8)) * CVS(Mid$(OBX$, 9, 4)) / 100 * CVS(Mid$(OBX$, 13, 4)) / 100
     If IRET1# <= MinimoRetencionIBTucuman# Then
        Call REPLA(OBX$, MKD$(0), 17, 8)
     Else
        Call REPLA(OBX$, MKD$(IRET1#), 17, 8)
     End If
     '
     Label12 = TRIM$(CSTRING$(Mid$(OBX$, 17, 8), 3, 12, 2, 2))
     RETIBMAN% = 1
     '
   End If
End Sub

Private Sub Label14_Change()
   Label20.Caption = TRIM$(CSTRING$(MKD$(XVALO(Label18.Caption) - XVALO(Label30.Caption) - XVALO(Label14.Caption) - XVALO(Label45.Caption) - XVALO(Label12.Caption) - XVALO(Label21.Caption) - XVALO(Label38) - XVALO(Label49)), 4, 14, 2, 2))
   Label20.Refresh
End Sub


Private Sub Label30_Change()
   Label20.Caption = TRIM$(CSTRING$(MKD$(XVALO(Label18.Caption) - XVALO(Label30.Caption) - XVALO(Label14.Caption) - XVALO(Label45.Caption) - XVALO(Label12.Caption) - XVALO(Label21.Caption) - XVALO(Label38) - XVALO(Label49)), 4, 14, 2, 2))
   Label20.Refresh
End Sub

Private Sub Label14_DblClick()
   If XVALO(Label18) > 0 Then
     VDEF$ = MKD$(XVALO(Label23)) + MKS$(XVALO(Label29)) + MKS$(XVALO(Label27))
10   BASIMPON# = CVD(Left$(VDEF$, 8))
     IRET# = 0
     If BASIMPON# >= XVALO(Label28) Then
       IRET# = CVD(Left$(VDEF$, 8)) * CVS(Mid$(VDEF$, 9, 4)) / 100 * CVS(Mid$(VDEF$, 13, 4)) / 100
     End If
     VDEF$ = VDEF$ + MKD$(IRET#)
     OBX$ = OBJPLA$("CARIRETIB", VDEF$)
     If OBX$ = "" Then Exit Sub
     IRET1# = 0
     If CVD(Left$(OBX$, 8)) >= XVALO(Label28) Then
       IRET1# = CVD(Left$(OBX$, 8)) * CVS(Mid$(OBX$, 9, 4)) / 100 * CVS(Mid$(OBX$, 13, 4)) / 100
     End If
     If Abs(IRET# - IRET1#) > 0.0005 Then
       VDEF$ = Left$(OBX$, 16)
       GoTo 10
     End If
     '
     Label23 = TRIM$(CSTRING$(Left$(OBX$, 8), 3, 12, 2, 2))
     Label29 = TRIM$(CSTRING$(Mid$(OBX$, 9, 4), 3, 12, 2, 2))
     Label27 = TRIM$(CSTRING$(Mid$(OBX$, 13, 4), 3, 12, 2, 2))
     Label14 = TRIM$(CSTRING$(Mid$(OBX$, 17, 8), 3, 12, 2, 2))
     RETIBMAN% = 1
     '
   End If
End Sub

Private Sub Label18_Change()
   Call CALRETGAN
   Call CALRETIBRU
   Call CALRETIBRU_MISIONES
   Call CALRETIBRU_TUCUMAN
   Call CALRETIVA
   If XVALO(Label18.Caption) > 0 Then
        BOTNEXT.Enabled = True
      Else
        BOTNEXT.Enabled = False
   End If
End Sub

Private Sub Label2_DblClick()
    STORPA07.Show 1
End Sub

Private Sub Label20_Change()
   
    If XVALO(Label20) < (-0.005) Then
      If List2.ListCount < 1 Then Exit Sub
10    VACON% = 1
      NETOCRE# = XVALO(Label21) + XVALO(Label14) + XVALO(Label45) + XVALO(Label30) + XVALO(Label38) + XVALO(Label49)
      '
      For KKI% = 2 To List2.ListCount
        NETOCRE# = NETOCRE# - XVALO(Mid$(List2.LIST(KKI% - 1), 70, 14))
      Next KKI%
      If NETOCRE# < 0 Then NETOCRE# = 0
      
      Z$ = List2.LIST(0)
      IACRE$ = FORMATNUM$(NETOCRE#, "F14.2")
      Call REPLA(Z$, IACRE$, 70, 14)
      List2.AddItem Z$, 0
      List2.RemoveItem 1
      VACON% = 0
      '
      SUMACRE# = 0
      For KKI% = 1 To List2.ListCount
        SUMACRE# = SUMACRE# + XVALO(Mid$(List2.LIST(KKI% - 1), 70, 14))
      Next KKI%
      '
      Label20.Caption = ""
      Label20.Refresh
      Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
      Label18.Refresh
      Label10.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
      Label10.Refresh
      If XVALO(Label20) >= 0.005 Then GoTo 10
    End If
    '
    If XVALO(Label20.Caption) >= 0.005 Then
        Command5.Enabled = False
        BOTNEXT.Enabled = True
      Else
        Command5.Enabled = True
        BOTNEXT.Enabled = False
    End If
    '
    'Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
    '
End Sub

Private Sub Label26_Click()
   Call SELECHO("INDIACRE")
   CPROX$ = TRIM$(VALACT1$("INDIACRE"))
   If CPROX$ <> "" Then
     NCXI% = NUMEPRO%(CPROX$)
     If NCXI% > 0 Then
      '\\\LIMPIAR TOTALES -OT-08-0889-OD-25/09/08///
        Call LimpTotOrPgo
      '\\\OT-08-0889-OD-FIN///
       If REGICLI&((-1) * NCXI%) > 0 Then
         Label26 = CPROX$
         NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NC%))
         Text2.TEXT = rasocli$((-1) * NC%)
         Text3.TEXT = DOMICLI$((-1) * NC%)
         Text4.TEXT = CPOSCLI$((-1) * NC%)
         Text5.TEXT = LOCACLI$((-1) * NC%)
         Text7.TEXT = ORDECHEQ$((-1) * NC%)
           Text7.ForeColor = vbBlack
           If Text7 <> rasocli(NC% * (-1)) Then Text7.ForeColor = vbRed
         Call SIGUEORPA1
       End If
     End If
   End If
End Sub

Private Sub Label30_DblClick()
    SHOWRETGA.Show 1
    If UCase$(SHOWRETGA.Label24) = "OK" Then
      If DERACCE%("AJURETGA", "Ajustar Importe Retención Ganancias") = 2 Then
        IRETGA# = XVALO(SHOWRETGA.Text1)
        Label30 = TRIM$(FORMATNUM$(IRETGA#, "F10.2"))
        RETGANMAN% = 1
        Call COMUNI("Para la Presente O/Pago\se Encuentra SUPRIMIDO\el Cálculo Automatico de Retención de Ganancias.")
        Exit Sub
      End If
    End If
    SHOWRETGA.Text1 = Label30
End Sub


Private Sub Label33_DblClick()
   '\\\OT-07-0190-OD-24/04/08///
   If TRIM$(Text1) = "" Or TRIM$(Text2) = "" Then
     Exit Sub
   End If
   NC% = XVALO(Text1)
   Call OPORPA07.VAPLIOP(NC%)
   '
End Sub




Private Sub Label38_Change()
   Label20.Caption = TRIM$(CSTRING$(MKD$(XVALO(Label18.Caption) - XVALO(Label30.Caption) - XVALO(Label14.Caption) - XVALO(Label45.Caption) - XVALO(Label12.Caption) - XVALO(Label21.Caption) - XVALO(Label38) - XVALO(Label49)), 4, 14, 2, 2))
   Label20.Refresh
End Sub

Private Sub Label38_DblClick()

   If XVALO(Label18) > 0 Then
     Proveedor% = XVALO(Text1)
     For idList2% = 0 To List2.ListCount - 1
        RenglonList2$ = List2.LIST(idList2%)
        ImporteRetenido# = XVALO(Mid$(RenglonList2$, 165, 12))
        If ImporteRetenido# > 0.005 Then
            Comprobante$ = TRIM$(Mid$(RenglonList2$, 1, 18))
            ImporteNetoComprobante# = XVALO(VALOBADA("CAJABANC", "VALABSCOMP", "CODICOM_LAR='" & Comprobante$ & "' AND NUPROVE=" & CStr(Proveedor%) & " AND TIPOMOVIM='COMPRA'"))
            
            ImporteNetoTotal# = ImporteNetoTotal# + ImporteNetoComprobante#
        End If
     Next idList2%
     
     ImporteRetenido# = XVALO(Label38)
     If ImporteNetoTotal# > 0.005 Then
        TasaRetencion# = XVALO(FORMATNUM$(ImporteRetenido# / ImporteNetoTotal#, "F12.2")) * 100
     End If
     
     VDEF$ = MKD$(ImporteNetoTotal#) + MKS$(TasaRetencion#) + MKS$(100)
     BASIMPON# = CVD(Left$(VDEF$, 8))
     VDEF$ = VDEF$ + MKD$(ImporteRetenido#)
10   OBX$ = OBJPLA$("CARIRETIB", VDEF$)
     If OBX$ = "" Then Exit Sub
     IRET1# = 0
     If CVD(Left$(OBX$, 8)) >= MinimoImponible# Then
       IRET1# = CVD(Left$(OBX$, 8)) * CVS(Mid$(OBX$, 9, 4)) / 100
     End If
     '
     Label38 = TRIM$(FORMATNUM$(IRET1#, "F12.2"))
     '
   End If
   
End Sub

Private Sub Label45_Change()
   Label20.Caption = TRIM$(CSTRING$(MKD$(XVALO(Label18.Caption) - XVALO(Label30.Caption) - XVALO(Label14.Caption) - XVALO(Label45.Caption) - XVALO(Label12.Caption) - XVALO(Label21.Caption) - XVALO(Label38) - XVALO(Label49)), 4, 14, 2, 2))
   Label20.Refresh
End Sub

Private Sub Label45_DblClick()
   If XVALO(Label18) > 0 Then
     RetencionRecalculado# = XVALO(Label45)
     Porcentaje# = (RetencionRecalculado# / XVALO(Label23)) * 100
     VDEF$ = MKD$(XVALO(Label23)) + MKS$(Porcentaje#) + MKS$(100)
10   BASIMPON# = CVD(Left$(VDEF$, 8))
     IRET# = XVALO(Label45)
     VDEF$ = VDEF$ + MKD$(IRET#)
     OBX$ = OBJPLA$("CARIRETIB", VDEF$)
     If OBX$ = "" Then Exit Sub
     
     IRET1# = CVD(Left$(OBX$, 8)) * CVS(Mid$(OBX$, 9, 4)) / 100 * CVS(Mid$(OBX$, 13, 4)) / 100
     If IRET1# <= MinimoRetencionIBMisiones# Then
        Call REPLA(OBX$, MKD$(0), 17, 8)
     Else
        Call REPLA(OBX$, MKD$(IRET1#), 17, 8)
     End If
     '
     Label45 = TRIM$(CSTRING$(Mid$(OBX$, 17, 8), 3, 12, 2, 2))
     RETIBMAN% = 1
     '
   End If
End Sub

Private Sub Label49_Click()
   Label20.Caption = TRIM$(CSTRING$(MKD$(XVALO(Label18.Caption) - XVALO(Label30.Caption) - XVALO(Label14.Caption) - XVALO(Label45.Caption) - XVALO(Label12.Caption) - XVALO(Label21.Caption) - XVALO(Label38) - XVALO(Label49)), 4, 14, 2, 2))
   Label20.Refresh
End Sub

Private Sub LIST1_Click()
    '
    If VACON% <> 0 Then
        Exit Sub
    End If
    '
    Z$ = List1.TEXT
    LIST1ANT$ = Z$
    ILISTA& = List1.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = "  " + Mid$(Z$, 57, 12)
    SUMAVAL# = XVALO(Label21.Caption)
    MAXAPLI# = XVALO(Label18.Caption) - XVALO(Label21.Caption) - XVALO(Label14.Caption) - XVALO(Label45.Caption) - XVALO(Label12.Caption) - XVALO(Label30.Caption) - XVALO(Label38.Caption) - XVALO(Label49.Caption)
    If IACRE$ = "" Then
'         If XVALO(IPEN$) <= MAXAPLI# + 0.005 Then
             Call REPLA(Z$, IPEN$, 70, 14)
    List1.AddItem Z$, ILISTA&
    List1.RemoveItem ILISTA& + 1
             SUMAVAL# = SUMAVAL# + CDbl(XVALO(IPEN$))
'           Else
Call APLIAUTO
'             Call MENSERR(24, "Importe Valor Excede Saldo a Cancelar")
'         End If
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMAVAL# = SUMAVAL# - CDbl(XVALO(IACRE$))
    End If
    VACON% = 1
    List1.AddItem Z$, ILISTA&
    List1.RemoveItem ILISTA& + 1
    List1.Refresh
    List1.ListIndex = ILISTA&
    VACON% = 0
    '
    Label21.Caption = TRIM$(CSTRING$(MKD$(SUMAVAL#), 4, 14, 2, 2))
    Label21.Refresh
    Label20.Caption = TRIM$(CSTRING$(MKD$(XVALO(Label18.Caption) - XVALO(Label14.Caption) - XVALO(Label45.Caption) - XVALO(Label12.Caption) - XVALO(Label30.Caption) - XVALO(Label38.Caption) - XVALO(Label49.Caption) - SUMAVAL#), 4, 14, 2, 2))
    Label20.Refresh
    'Call REMPLAGRID(Me, GRID2, Picture1, Label12, List1)
    '
End Sub

Sub APLIAUTO()
    '
10  RETIBANT# = (XVALO(Label14.Caption) + XVALO(Label38.Caption) + XVALO(Label45.Caption) + XVALO(Label12.Caption))
    retganant# = XVALO(Label30)
    RETIVA# = XVALO(Label49.Caption)
    '
    SUMAVAL# = 0
    For KU% = 1 To List1.ListCount
      SUMAVAL# = SUMAVAL# + XVALO(Mid$(List1.LIST(KU% - 1), 70, 14))
    Next KU%
    '
    VACON% = 1
    REVALTO# = SUMAVAL# + RETIBANT# + retganant#
    SUMACRE# = 0
    For KU% = 1 To List2.ListCount
      Z$ = List2.LIST(KU% - 1)
      LIST2ANT$ = Z$
      IACRE$ = TRIM$(Mid$(Z$, 70, 14))
      REVALTO# = REVALTO# - XVALO(IACRE$)
      SUMACRE# = SUMACRE# + XVALO(IACRE$)
    Next KU%
    If REVALTO# < 0.005 Then GoTo 60
    '
    REVALTO# = SUMAVAL# + RETIBANT# + retganant#
    SUMACRE# = 0
    '
    If List2.ListCount < 2 Then
      Z$ = List2.LIST(0)
      IACRE$ = FORMATNUM$(REVALTO#, "F14.2")
      Call REPLA(Z$, IACRE$, 70, 14)
      List2.AddItem Z$, 0
      List2.RemoveItem 1
      REVALTO# = REVALTO# - XVALO(IACRE$)
      SUMACRE# = SUMACRE# + XVALO(IACRE$)
      GoTo 60
    End If
    '
    For KU% = 1 To List2.ListCount
      Z$ = List2.LIST(KU% - 1)
      LIST2ANT$ = Z$
      IACRE$ = TRIM$(Mid$(Z$, 70, 14))
      IPEN$ = " " + Mid$(Z$, 56, 13)
      If KU% > 1 Then
        If XVALO(IACRE$) < XVALO(IPEN$) Then
          IAPLI# = REVALTO#
          If IAPLI# > XVALO(IPEN$) Then IAPLI# = XVALO(IPEN$)
          If IAPLI# < 0 Then IAPLI# = 0
          IACRE$ = FORMATNUM$(IAPLI#, "F14.2")
          Call REPLA(Z$, IACRE$, 70, 14)
          List2.AddItem Z$, KU% - 1
          List2.RemoveItem KU%
        End If
      End If
      REVALTO# = REVALTO# - XVALO(IACRE$)
      SUMACRE# = SUMACRE# + XVALO(IACRE$)
    Next KU%
    '
    If REVALTO# >= 0.005 Then
      Z$ = List2.LIST(0)
      LIST2ANT$ = Z$
      IACRE$ = FORMATNUM$(REVALTO# + XVALO(Mid$(Z$, 70, 14)), "F14.2")
      Call REPLA(Z$, IACRE$, 70, 14)
      List2.AddItem Z$, 0
      List2.RemoveItem 1
      REVALTO# = 0  ' REVALTO# - XVALO(IACRE$)
      'SUMACRE# = SUMACRE# + XVALO(IACRE$)
    End If
    '
60  VACON% = 0
    Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label18.Refresh
    Label10.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label10.Refresh
    Label20.Caption = TRIM$(CSTRING$(MKD$(XVALO(Label18.Caption) - XVALO(Label14.Caption) - XVALO(Label38.Caption) - XVALO(Label45.Caption) - XVALO(Label12.Caption) - XVALO(Label21.Caption) - XVALO(Label49.Caption) - XVALO(Label30.Caption)), 4, 14, 2, 2))
    Label20.Refresh
    '
    'If XVALO(Label14) <> RETIBANT# Then GoTo 10
    If (XVALO(Label14.Caption) + XVALO(Label38.Caption) + XVALO(Label45.Caption) + XVALO(Label12.Caption)) <> RETIBANT# Then GoTo 10
    If XVALO(Label30) <> retganant# Then GoTo 10
    If XVALO(Label49.Caption) <> RETIVA# Then GoTo 10
    '
    'Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
    '
End Sub
'
Private Sub List1_DblClick()
    '
    ' Call SIGUEORPA2
    '
End Sub

Private Sub List2_Click()
    '
    If VACON% <> 0 Then
        Exit Sub
    End If
    '
    NPROVE% = XVALO(Text1)
    If TICUEPRO%((-1) * NPROVE%) = 1 Then   ' cuenta de gastos
       Call SIGUEORPA5
       Exit Sub
    End If
    '
    Z$ = List2.TEXT
    LIST2ANT$ = Z$
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = " " + Mid$(Z$, 56, 13)
    SUMACRE# = XVALO(Label18.Caption)
    If IACRE$ = "" Then
         Call REPLA(Z$, IPEN$, 70, 14)
         SUMACRE# = SUMACRE# + CDbl(XVALO(IPEN$))
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
         If Left$(Z$, 20) = "Anticipo no Aplicado" Then
            Call REPLA(Z$, Space$(13), 32, 13)
            Call REPLA(Z$, Space$(13), 56, 13)
         End If
    End If
    
    ' *** VERIFICA SI TIENE QUE RETENER IIBB POR CABA, A PARTIR DE QUE EXISTA EL PADRON DE ALICUOTA
    If EXISTE%(LUDAT$ + "IIBBCABA.TXT") > 0 Then
        
        NumeroProveedor& = XVALO(Text1)
        codigoProvincia$ = UCase(PROVCLI$((-1) * NumeroProveedor&))
        If codigoProvincia$ = "C" Then
            Comprobante$ = TRIM$(Mid$(Z$, 1, 18))
            ImporteACancelar# = XVALO(TRIM$(Mid$(Z$, 70, 14)))
            ' *** VERIFICA QUE HAYA ELEGIDO EL COMPROBANTE MEDIANTE LA EXISTENCIA DEL IMPORTE QUE ESTA POR CANCELAR
            If ImporteACancelar# > 0.005 Then
                ' *** VERIFICA QUE EL COMPROBANTE NO HAYA SIDO PREVIAMENTE RETENIDO
                If CantRegistros&("CAJABANC", "TIPOMOVIM='RTIB' AND OPCIMOVIM='CABA' AND NOTAS LIKE '%" & Comprobante$ & "%'") = 0 Then
                   ImporteNetoComprobante# = XVALO(VALOBADA("CAJABANC", "SUM(VALABSCOMP)", "CODICOM_LAR='" & Comprobante$ & "' AND NUPROVE=" & Text1 & " AND TIPOMOVIM='COMPRA'"))
                   TasaRetencion# = ALIRETBRU_CABA#(XVALO(Text1), MinimoImponible#, MEDIOPAGO%)
                   ' *** VERIFICA QUE EL IMPORTE NETO DEL COMPROBANTE SEA MAYOR AL MINIMO IMPONIBLE
                   If ImporteNetoComprobante# > 0.005 And ImporteNetoComprobante# > MinimoImponible# Then
                        ImporteRetenido# = ImporteNetoComprobante# * TasaRetencion# / 100
                        Z1$ = Left$(Z$ + Space$(256), 256)
                        Call REPLA(Z1$, FORMATNUM$(ImporteRetenido#, "F12.2"), 165, 12)
                        Z$ = Z1$
                        Label39 = CStr(MEDIOPAGO%)
                    End If
                End If
            Else
                ' *** DESELECCIONO EL COMPROBANTE, CON LO CUAL SE ELIMINA LA RETENCION
                Z1$ = Left$(Z$ + Space$(256), 256)
                Call REPLA(Z1$, FORMATNUM$(0, "F12.2"), 165, 12)
                Z$ = Z1$
            End If
        End If
    End If
    
    VACON% = 1
    List2.AddItem Z$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    List2.Refresh
    List2.ListIndex = ILISTA&
    
    ' *** RETENCION DE IVA SOBRE FACTURAS 'M'
    '
    ImporteRetenido# = 0
    NumeroProveedor& = CLng(Text1)
    IndiceComprobante% = List2.ListIndex
    Renglon$ = List2.LIST(IndiceComprobante%)
    ImporteACancelar# = XVALO(TRIM$(Mid$(Z$, 70, 14)))
    If Abs(ImporteACancelar#) > 0.005 Then
        Comprobante$ = TRIM$(Mid$(Renglon$, 1, 18))
        
        ' VERIFICA QUE SE TRATE DE UN COMPROBANTE 'M'
        'If (Mid$(Comprobante$, 1, 2) = "FC" Or Mid$(Comprobante$, 1, 2) = "ND") And Mid$(Comprobante$, 4, 1) = "M" Then
        If Mid$(Comprobante$, 4, 1) = "M" Then
            CondicionSql$ = "NUPROVE=" & CStr(NumeroProveedor&) & " AND CODICOM_LAR='" & Comprobante$ & "' AND TIPOMOVIM='COMPRA'"
            ImporteNetoComprobante# = XVALO(VALOBADA("CAJABANC", "SUM(VALABSCOMP)", CondicionSql$))
            
            ' VALIDA QUE NO SE HAYA RETENIDO ANTERIORMENTE Y QUE EL VALOR NETO DEL COMPROBANTE SEA MAYOR O IGUAL AL MINIMO IMPONIBLE
            ' SI SE TRATA DE UNA NOTA DE CREDITO, NO SE APLICA LA VALIDACION DEL NETO MAYOR AL MINIMO IMPONIBLE
            If CantRegistros&("CAJABANC", "TIPOMOVIM='RTIVA' AND NOTAS LIKE '%" & Comprobante$ & "%'") = 0 And (ImporteNetoComprobante# >= MinimoImponibleRetencionIva# Or Mid$(Comprobante$, 1, 2) = "NC") Then
                CondicionSql$ = "NUPROVE=" & CStr(NumeroProveedor&) & " AND CODICOM_LAR='" & Comprobante$ & "'"
                CondicionSql$ = CondicionSql$ & " AND TIPOMOVIM='IVA' and VARIMOVIM<>'Aduana' and VARIMOVIM<>'Per-Iva' and VARIMOVIM<>'Per-Gan'"
                ImporteRetenido# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)-SUM(IHABECOMP)", CondicionSql$))
            End If
        End If
    End If
    ' ACTUALIZA LA LISTA AGREGANDO EL IMPORTE RETENIDO DEL COMPROBANTE
    Z1$ = Left$(Z$ + Space$(256), 256)
    Call REPLA(Z1$, FORMATNUM$(ImporteRetenido#, "F12.2"), 177, 12)
    Z$ = Z1$
    List2.AddItem Z$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    List2.Refresh
    List2.ListIndex = ILISTA&
    
    If ImporteRetenido# > 0.005 And MedioPagoRetencionIva% = 0 Then
        Call MENSERR(24, "La Orden de Pago en Proceso\Recibirá una Retención de IVA.\ \Falta Configurar la Cuenta Contable Correspondiente.\Debe Realizar la Configurar y luego volver Procesar el Pago!")
    End If
    ' *** FIN CALCULO DE RETENCION IVA
    '
    
    'Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
    ' *** ACTUALIZA TOTAL DE RETENCIONES CABA
    Call CALRETIBRU_CABA
    Call CALRETIBRU_MISIONES
    Call CALRETIBRU_TUCUMAN
    Call CALRETIVA
    VACON% = 0
    '
    Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label18.Refresh
    Label10.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label10.Refresh
    Label20.Caption = TRIM$(CSTRING$(MKD$(XVALO(Label18.Caption) - XVALO(Label14.Caption) - XVALO(Label45.Caption) - XVALO(Label12.Caption) - XVALO(Label38.Caption) - XVALO(Label49.Caption) - XVALO(Label21.Caption) - XVALO(Label30.Caption)), 4, 14, 2, 2))
    Label20.Refresh
    '
End Sub
'

Private Sub List2_DblClick()
    '
    NPROVE% = XVALO(Text1)
    If TICUEPRO%((-1) * NPROVE%) = 1 Then   ' cuenta de gastos
       Call SIGUEORPA5
       Exit Sub
    End If
    '
    VACON% = 1
    ILIST2& = List2.ListIndex
    TOLIST2& = List2.TopIndex
    '
    Z$ = List2.TEXT
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = Mid$(Z$, 56, 14)
    SUMACRE# = XVALO(Label18.Caption)
    SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
    SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(LIST2ANT$, 70, 14)))
    If ILISTA& >= 0 Then
      List2.Visible = False
      List2.AddItem LIST2ANT$, ILISTA&
      List2.RemoveItem ILISTA& + 1
    End If
    '
    Call SIGUEORPA3
    '
    VACON% = 1
    List2.TopIndex = TOLIST2&
    List2.Visible = True
    If List2.ListCount > ILIST2& Then
          List2.ListIndex = ILIST2&
        Else
          Call CIERRARCH("!CARMEPA")
          Call BLANARCH("!CARMEPA")
          Call CIERRARCH("!CARGAPLI")
          Call BLANARCH("!CARGAPLI")
          Call CIERRARCH("!CARGAST")
          Call BLANARCH("!CARGAST")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label45.Caption = ""
          Label12.Caption = ""
          Label38.Caption = ""
          Label49.Caption = ""
          Label30.Caption = ""
          Label18.Caption = ""
          Label20.Caption = ""
          Label21.Caption = ""
          'BOTNEXT.Enabled = True
          Command5.Enabled = False
          EPAC$ = UCase$(EMPREAC$)
          If InStr(EPAC$, "NEUMAN") > 0 Then
            Text1.TEXT = ""
            Text2.TEXT = ""
            Text3.TEXT = ""
            Text4.TEXT = ""
            Text5.TEXT = ""
          End If
          Text1.SetFocus
    End If
    VACON% = 0
    '
    'Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
    '
End Sub

Private Sub ListOP_Click()
    Call OPORPA07.Command5_Click
End Sub

Private Sub MaeProves_Click()
    Call Command11_Click
End Sub

Private Sub mnuAnalisisEstaPagos_Click()
  Call OPORPA07.ANAPACHEQ
End Sub

Private Sub mnuAplicaReparacionesAPagos_Click()
    
    APLIREPA.Show 1
    
End Sub

Private Sub MnuApliCredAnt_Click()
   '\\\OT-07-0190-OD-24/04/08///
   Call OPORPA07.Command11_Click
   '\\\OT-07-0190-OD-FIN///

End Sub

Private Sub mnuAsociarReporte_Click()
  Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub mnuCambarOrdenamiento_Click()
   Call Command8_Click
End Sub

Private Sub mnuCompensacionClteProv_Click()
   Call OPORPA07.Command10_Click
End Sub

Private Sub mnuConfiGral_Click()
    Call Command7_Click
End Sub

Private Sub mnuDesaplicarOP_Click()
    OPORPA07.Command12_Click
End Sub

Private Sub MnuFondoFijo_Click()
 Call Option2_Click
End Sub

Private Sub mnuGuardar_Click()
   Call Command5_Click
End Sub

Private Sub mnuLiCuCor_Click()
    Call RECUEP07.Command3_Click
End Sub

Private Sub mnuMaeProv_Click()
    Call Command11_Click
End Sub

Private Sub mnuOtrAccDir_Click()
    
    Call COMUNI("No Existen Accesos Directos \ para esta Aplicación")

End Sub

Sub mnuReCuCor_Click()
    Call Command4_Click
End Sub

Private Sub mnuReportesVarios_Click()
   Call OPORPA07.AGenRep_Click
End Sub

Private Sub mnuRetenIIBB_Click()
   OPORPA07.Command17_Click
End Sub

Private Sub mnuSali_Click()
    Call Command2_Click
End Sub

Private Sub mnuTaMePa_Click()
    Call OPORPA07.Command1_Click
End Sub

Private Sub mnuTransfPago_Click()
  Call OPORPA07.Command13_Click
End Sub

Private Sub mnuVeryReimprmirRetencionesGanancia_Click()
   OPORPA07.Command16_Click
End Sub

Private Sub MONEMIS_Change()
    '
    '\\\OT-07-0086-RG-06/03/07///
    ENTRATICA.Label6.Caption = MONEMIS.Caption
    ENTRATICA.Text4 = TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4"))
    Call SIGUEORPA1
    '\\\OT-07-0086-RG-FIN///
    '
End Sub


Private Sub MONEMIS_DblClick()
  Static CAMON%, DMONE$(), TCMONE()
   ' Exit Sub    ' PORQUE TODAVÍA NO ESTA TERMINADO BIMONETARIO
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
10   If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         U2& = 0
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           If TRIM$(DMX$) <> "" Then
             VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
             U2& = U2& + 1
             DMONE$(U2&) = DMX$
             TCMONE(U2&) = VHX
           End If
         Next U&
         CAMON% = U2&
         If CAMON% < 1 Then GoTo 10
     End If
   End If
   '
   MONEMIA% = MONEMI%
   MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   MONEMIS = DMONE$(MONEMI%)
   MONEMIS.ForeColor = RGB(0, 0, 255)
   MONEMIS.FontSize = 11
   If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
      MONEMIS.FontSize = 10
   End If
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   TICAMBX# = TICAMONE(MONEMI%)
   TICAMBIS = FORMATNUM(TICAMBX#, "F9.4")
   '
   List1.Clear
   List2.Clear
   Label4.Caption = ""
   Label8.Caption = ""
   Label10.Caption = ""
   Label14.Caption = ""
   Label45.Caption = ""
   Label12.Caption = ""
   Label38.Caption = ""
   Label49.Caption = ""
   Label18.Caption = ""
   Label20.Caption = ""
   Text7 = ""
   BOTNEXT.Enabled = True
   '
   Call SIGUEORPA1
   '
End Sub

Private Sub NotProve_Click()
    Call Command3_Click
End Sub

Private Sub NueOPs_Click()
    Call Command1_Click
End Sub

Private Sub Option1_DblClick()
   TOTACAN# = CDbl(XVALO(Label18.Caption))
   If TOTACAN# > 0.005 Then
      NICON% = NIVCONFI%
      If NICON% > 0 Then
         Label40 = "<" + TRIM$(Str$(NICON%))
         Picture6.Visible = True
      End If
   End If
End Sub

Private Sub Option2_Click()
    'Hide
    Call FORLIFOF.LIFOFIJ1
    FORLIFOF.Show 1
    Call FRESHALL

End Sub



Private Sub Text1_Change()
    Call CIERRARCH("!CARGAST")
    Call BLANARCH("!CARGAST")
    List1.Clear
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label18.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label45.Caption = ""
    Label12.Caption = ""
    Label38.Caption = ""
    Label49.Caption = ""
    Label30.Caption = ""
    Label20.Caption = ""
    Label21.Caption = ""
    Picture6.Visible = False
    'BOTNEXT.Enabled = True
    '
    SHOWRETGA.Label2 = Text1
    SHOWRETGA.Label7 = ""
    SHOWRETGA.Label9 = ""
    SHOWRETGA.Label10 = ""
    SHOWRETGA.Label11 = ""
    SHOWRETGA.Label12 = ""
    SHOWRETGA.Label14 = ""
    SHOWRETGA.Label15 = ""
    SHOWRETGA.Label16 = ""
    SHOWRETGA.Label17 = ""
    SHOWRETGA.Text1 = ""
    SHOWRETGA.List1.Clear
    '
    RETIBMAN% = 0
    RETGANMAN% = 0
    '
    If TRIM$(Text1) = "" Then A% = PIVACLI%(0)  ' ES PARA REFRESCAR LAS RUTINAS DEL TIPO RASOCLI, PIVACLI, ETC.
    '
    Command5.Enabled = False
    '
    'Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
    '
End Sub

Private Sub Text1_DblClick()
    '
    Call Command1_Click
    '
End Sub

Private Sub Text1_GotFocus()
    Static CTXX%
    If CTXX% = 0 Then
        Command5.Enabled = False
        Refresh
        CTXX% = 1
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If XVALO(Text1.TEXT) > 0 Then
          If XVALO(Text1.TEXT) <= 32767 Then
             NC% = XVALO(Text1.TEXT)
             'If ECOARCH$("PROVED1", MKI$(NC%)) <> "" Then
             NPRO& = NC%  ' SE CAMBIO LO DE ARRIBA COMENTADO POR LO SIGUIENTE PARA QUE BUSQUE EN LA BASE
             Call ABREPROVE
             PROV& = XVALO(VALOBADA("Proved12", "Nume_Cli", " Nume_Cli = " & NPRO&))
             If PROV& > 0 Then
                Text2.TEXT = rasocli$((-1) * NC%)
                Text3.TEXT = DOMICLI$((-1) * NC%)
                Text4.TEXT = CPOSCLI$((-1) * NC%)
                Text5.TEXT = LOCACLI$((-1) * NC%)
                Text7.TEXT = ORDECHEQ$((-1) * NC%)
                Label26 = CODICLI((-1) * NC%)
                  Text7.ForeColor = vbBlack
                  If Text7 <> rasocli(NC% * (-1)) Then Text7.ForeColor = vbRed
                  
                ' *** CONFIGURA RETENCION IB MISIONES
                RetieneIBMisiones% = XVALO(VALOBADA("PROVED12", "CategoriaIBMISIONES_Cli", "NUME_CLI=" & CStr(NC%)))
                If RetieneIBMisiones% > 0 Then ' RetieneIBMisiones% tiene la categoria.
                    PorcentajeRetencionIBMisiones# = ALIRETBRU_MISIONES#(XVALO(Text1), MinimoRetencionIBMisiones#, MedioPagoIBMisiones%)
                End If
                  
                Call VERIRECHA(NC%) ' MUESTRA RECHAZOS EN LINEA (DOSIVAC)
                Call SIGUEORPA1
                Exit Sub
             End If
          End If
       End If
       '
       Beep
       Text1.TEXT = ""
       '
    End If
    '
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        List2.SetFocus
    End If
End Sub

Private Sub Text11_LostFocus()
    '
    VDEFF$ = TRIM$(Text11.TEXT)
    Text11.TEXT = VDEFF$
    If Left$(VDEFF$, 3) <> "OP." Then GoTo 90
    NRONUE& = XVALO(Mid$(VDEFF$, 4))
    NROPRO& = NRONUE&
    Call NUORPA(NROPRO&)
    If NROPRO& <> NRONUE& Then
        Beep
        Call MENSERR(24, "Número no Válido o Pre-existente")
        GoTo 95
    End If
    Exit Sub
    '
90  Beep
    Call MENSERR(24, "Número no Válido")
    NROPRO& = 0
    Call NUORPA(NROPRO&)
95  Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
    Text11.SetFocus
    '
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text3_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text4_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text5_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text6_Click()
    Text1.SetFocus
End Sub

Private Sub Text6_DblClick()
    '\\\OT-06-0570-RG-28/12/06///
    STORPA07.Show 1
    Call COMUNI("Re-Inicie Ahora la Aplicación.")
    Call FINAL("")
    '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text7_GotFocus()
    '
    Text7.SelStart = 0
    Text7.SelLength = Len(Text7)
    '
End Sub

Private Sub Text9_Change()
    If FECHANUM(Text9) > 0 Then
       Call CALRETGAN
    End If
End Sub

Private Sub Text9_DblClick()
    If DERACCE%("CAMFEOP", "Cambio de Fecha de Orden de Pago") >= 1 Then
        Call SELECHO("SELFECHA")
        VDEVV$ = VALACT1$("SELFECHA")
        If VDEVV$ <> "" Then
            Text9.TEXT = VDEVV$
            FF% = FECHANUM(Text9.TEXT)
        End If
    End If
End Sub

Private Sub Text9_GotFocus()

    If DERACCE%("CAMFEOP/NOMESS", "Cambio de Fecha de Orden de Pago") < 1 Then
        Text1.SetFocus
    End If
    
    Screen.MousePointer = 1
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text11.SetFocus
    End If
End Sub

Private Sub Text9_LostFocus()

    If DERACCE%("CAMFEOP/NOMESS", "Cambio de Fecha de Orden de Pago") >= 1 Then
        VDEVV$ = TRIM$(Text9.TEXT)
        If FECHANUM(VDEVV$) > 0 Then
              FF% = FECHANUM(VDEVV$)
              FEX = FF%
              Text9 = FECHATEX$(FEX)
            Else
              Call MENSERR(24, "Fecha no Válida")
              FEX = HOY(HO$)
              Text9.TEXT = HO$
              Text9.SetFocus
        End If
    End If
End Sub

Sub NUORPA(NROPRO&)
    '
    Screen.MousePointer = 11
    '
    FENUE$ = Control$("", "FENUORPA")
    FEXII% = FECHANUM(FENUE$)
    If FEXII% < 33 Then FEXII% = FECHANUM("01/01/80")
    FLINUME = FETEXCOR1$(FEXII%)
    '
    NROMA& = 0
22  Call APERBASDAT("CABAN")
    Dim MIRS As ADODB.Recordset
    SQLX$ = "SELECT NumeComp FROM CajaBanc WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE FechEmisi >= '" & FLINUME & "'"
    SQLX$ = SQLX$ + " AND TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + " AND CodiCom_Cor LIKE 'OP%' "
    SQLX$ = SQLX$ + " ORDER BY NumeComp DESC "
    '
    Set MIRS = New ADODB.Recordset
    MIRS.CursorLocation = adUseClient
'    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    'MODIFICADO PARA QUE DENTRO DE LA FUNCION READSET CAPTURE EL ERROR DE T.ESPERA
    Set MIRS = READSET(SQLX$)
    If Not (MIRS.BOF And MIRS.EOF) Then
        NROMA& = MIRS!NumeComp
'25      If VERINUMOP%(NROMA& + 1) < 0 Then
'           NROMA& = NROMA& + 1
'           GoTo 25
'        End If
        GoTo 80
    End If
    '
80  On Error Resume Next
    MIRS.Close
    Set MIRS = Nothing
    On Error GoTo 0
    '
    If NROPRO& > 0 Then
       Dim OrpaTemp As ADODB.Recordset
       SQLX$ = "SELECT * FROM CajaBanc WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE FechEmisi >= '" & FLINUME & "' "
       SQLX$ = SQLX$ + " AND TipoMovim = 'OPAG' "
       SQLX$ = SQLX$ + " AND CodiCom_Cor LIKE 'OP%' "
       SQLX$ = SQLX$ + " AND NumeComp = " & NROPRO&
'       Set OrpaTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
       Set OrpaTemp = READSET(SQLX$)
       If Not (OrpaTemp.EOF And OrpaTemp.BOF) Then 'si lo encuentra
          NROPRO& = 0
       End If
       OrpaTemp.Close
       Set OrpaTemp = Nothing
    End If
    '
90  If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    Screen.MousePointer = 1
    '
End Sub
'
Sub SIGUEORPA1()
    '
    Screen.MousePointer = 11
    If MODORPA% = 0 Then
      MODORPA% = 1       ' por emision
      If Control$("ORPASTA", "MODORDE") = "VENCIM" Then
        MODORPA% = 2
      End If
    End If
    '
    NCXI = XVALO(Text1.TEXT)
    If NCXI < 1 Or NCXI > 32767 Then
      Call MENSERR(24, "Numero de Cuenta no Válido")
      Text1.SetFocus
      Exit Sub
    End If
    '
'    NC% = NCXI
'
'    If ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
'      Call MENSERR(24, "Cuenta no Figura\en Base de Datos de Proveedores")
'      Text1.SetFocus
'      Exit Sub
'    End If
    '
    NPRO& = NCXI  ' SE CAMBIO LO DE ARRIBA COMENTADO POR LO SIGUIENTE PARA QUE BUSQUE EN LA BASE
    Call ABREPROVE
    PROV& = XVALO(VALOBADA("Proved12", "Nume_Cli", " Nume_Cli = " & NPRO&))
    If PROV& < 1 Then
      Call MENSERR(24, "Cuenta no Figura\en Base de Datos de Proveedores")
      Text1.SetFocus
      Exit Sub
    End If
    
    FE$ = TRIM$(Text9.TEXT)
    FF% = FECHANUM(FE$)
    HOI% = HOY(HO$)
    If FF% < 1 Or FF% > HOI% + 1 Then
      Call MENSERR(24, "Fecha no Válida")
      Text9.SetFocus
      Exit Sub
    End If
    '
    If DIENTRE%(FF%, HOI%) > 55 Then
      Call MENSERR(24, "Fecha no Válida")
      Text9.SetFocus
      Exit Sub
    End If
    '
    FECHALIM% = XVALO(Control$("", "CIEMECOM"))
    If FF% <= FECHALIM% Then
      Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
      Text9.SetFocus
      Exit Sub
    End If
     '
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 11
    NC% = XVALO(Text1.TEXT)
    If NC% > 0 Then
      Text2.TEXT = rasocli$((-1) * NC%)
      Text3.TEXT = DOMICLI$((-1) * NC%)
      Text4.TEXT = CPOSCLI$((-1) * NC%)
      Text5.TEXT = LOCACLI$((-1) * NC%)
      'Text7.TEXT = TEL1CLI$(NC%)
      'Text10.TEXT = CSTRING$(MKD$(SALDCLI#(NC%)), 4, 13, 2, 2)
      Refresh
    End If
    '
100 List1.Clear
    List1.Refresh
    List2.Clear
    List2.Refresh
    '
    NICONFLIM% = NIVCONFI%
    '
    Z$ = Space$(128)
    Call REPLA(Z$, "Anticipo no Aplicado", 1, 20)
    Call REPLA(Z$, HO$, 23, 8)
    Call REPLA(Z$, "       0", 119, 8)
    Call REPLA(Z$, " 0", 127, 2)
    List2.AddItem Z$
    '
    SUMADEU# = 0: SUMAFAC# = 0
    If TICUEPRO%((-1) * NC%) = 1 Then        ' GASTOS
      Call BLOQARCH("!CARGAST")
      Call BLANARCH("!CARGAST")
      GoTo 200
    End If
    '
    Call APERBASDAT("CABAN")
    ' JANDY SQL
    Dim FacPenTemp As ADODB.Recordset
    SQLX$ = "SELECT * FROM SACREPEN "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    'Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Set FacPenTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With FacPenTemp
    '  On Error Resume Next
    '  .MoveLast
    '  If Err < 1 Then
      If Not (.EOF And .BOF) Then
        .MoveFirst
        Do While Not .EOF
          '
          NICON% = XVALO(!Niv_Clasi)          ' NIVEL DE CLASIFICACION
          If NICON% > NICONFLIM% Then
             GoTo 993
          End If
          '
          TIX$ = UCase$(Left$(!CODICOM_LAR, 2))
            If Len(TRIM$(!CODICOM_LAR)) < 18 Then TIX$ = "CO"
          X$ = !CODICOM_LAR
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "TF" Then TIPOS$ = "Ticket-Factura"
                              If TIX$ = "RH" Then TIPOS$ = "Rec.Honorarios"
                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                              If TIX$ = "CO" Then TIPOS$ = "Comprobante"
                              If TIX$ = "CP" Then TIPOS$ = "Crédito Propio"
                              If TIX$ = "OP" Then TIPOS$ = "Orden de Pago": SG% = 1
                              If TIX$ = "FF" Then TIPOS$ = "O/P Fondo Fijo": SG% = 1
                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
                              If TIX$ = "DP" Then TIPOS$ = "Debito Propio": SG% = 1
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Cuenta": SG% = 1
                              '
          'If TIX$ = "FC" Or TIX$ = "ND" Or TIX$ = "TF" Or TIX$ = "RH" Or TIX$ = "CP" Or TIX$ = "CO" Or TIX$ = "AS" Or (TIX$ = "OP" And Mid$(X$, 6, 4) = "anul") Then
            SALDO# = XVALO(!SaldAcre) - XVALO(!SaldDebe)
            SALDOLA# = 0
            '
            If MODOMONEI% = 2 Then
               If IsNull(!Mone_Emis) = True Then
                  'call RECUSALDOL
               End If
               On Error Resume Next
               SALDOLA# = !SaldAcre_DOLA - !SaldDebe_DOLA
               MONEDOCU% = XVALO(!Mone_Emis)
               If MONEMI% > 1 Then
                 If MONEMI% <> MONEDOCU% Then GoTo 993
                 SALDO# = SALDOLA#
               End If
               On Error GoTo 0
            End If

            If Abs(SALDO#) >= 0.005 Then       ' AQUI SE INTERVINO PARA QUE TRAIGA LOS CREDITOS
              ' VALIDA TIPO DE COMPROBANTE Y CONFIGURACION (SUPRIME NOTA DE CREDTIO U O.P.)
              If TIX$ = "NC" And Control$("ORPASTA", "SUPRICRE") = "SI" Then GoTo 993
              If TIX$ = "OP" And Control$("ORPASTA", "SUPRIPAG") = "SI" Then GoTo 993
              '
              IMPO# = (-1) * !IBruComp * SG%
              SUMAFAC# = SUMAFAC# + IMPO#
              IMPO# = Abs(IMPO#)
              '
              SUMADEU# = SUMADEU# + SALDO#
              '
              NROX$ = !CODICOM_LAR
              FEMI = CVINT(!FECHEMISI)
              FEVE = CVINT(!FEVENCIM)
              '
              MEMI% = 0
              On Error Resume Next
              MEMI% = !Mone_Emis Mod 6
              On Error GoTo 0

              IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 13, 2, 2)
              IVATEX$ = CSTRING$(MKD$(IVAO#), 3, 13, 2, 2)
              SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 13, 2, 2)
              '
                    REGTEX$ = CSTRING$(MKS$(REG&), 3, 8, 0, 2)
                    IRETEX$ = CSTRING$(MKI$(KKI%), 1, 2, 0, 2)
              '
              Z$ = Space$(128)
              Call REPLA(Z$, NROX$, 1, 18)
              Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
              Call REPLA(Z$, IMPTEX$, 32, 13)
              Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
              Call REPLA(Z$, SALTEX$, 56, 13)
              Call REPLA(Z$, REGTEX$, 119, 8)
              Call REPLA(Z$, IRETEX$, 127, 2)
              '
              If MODOMONEI% = 2 Then
                 If MONEMI% <= 1 Then
                   Z$ = Left$(Z$ + Space$(160), 160)
                   Call REPLA(Z$, FORMATNUM$(MEMI%, "I2"), 147, 2)
                   Call REPLA(Z$, FORMATNUM$(SALDOLA#, "F12.2"), 149, 12)
                 End If
              End If
              
              List2.AddItem Z$
              '
            End If
          'End If
993     .MoveNext
        Loop
      End If
    End With
    FacPenTemp.Close
    Set FacPenTemp = Nothing
    '
200 List2.ListIndex = -1    ' CARGA DE GASTOS
    List2.Refresh
    Label8.Caption = TRIM$(CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2))
    Label8.Refresh
    Label10.Caption = ""
    Label10.Refresh
    List2.Visible = True
    List2.SetFocus
    If List2.ListCount > 0 Then
        VACON% = 1
        List2.ListIndex = 0
        VACON% = 0
    End If
    '
    Screen.MousePointer = 11
    List1.Clear
    '
    LINDICE.Clear
    If TRIM$(Control$("ORPASTA", "UTICHECA")) = "SI" Then
      'JANDY SQL
      Dim CHETEMP As ADODB.Recordset
      SQLX$ = "SELECT * FROM CAJABANC "
      SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
      SQLX$ = SQLX$ + "AND Status = 1 "
      SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
      If MODOSORVAL% = 1 Then
          SQLX$ = SQLX$ + "ORDER BY FechAcred ASC"
        Else
          SQLX$ = SQLX$ + "ORDER BY IdebeComp ASC"
      End If
      'Set CHETEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      Set CHETEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With CHETEMP
        'On Error Resume Next
        '.MoveFirst
        'If Err < 1 Then
        '  On Error GoTo 0
          
      If Not (.EOF And .BOF) Then
          .MoveFirst
          Do While Not .EOF
            '
            If CVINT%(!FECHENTRE) < 1 Then
              FF% = CVINT(!fechacred)
              ICHE# = !Idebecomp
              TIVA$ = TRIM$(Left$(!DESCRIMOV, 18))
              NUCHE$ = TRIM$(Mid$(!DESCRIMOV, 19, 14))
              FEVECH$ = TRIM$(Mid$(!DESCRIMOV, 33, 8))
              '
              Z$ = Space$(128)
              Call REPLA(Z$, "Ch.3", 1, 4)
                If Left$(UCase$(TIVA$), 3) = "DOC" Then
                  Call REPLA(Z$, "Doc3", 1, 4)
                End If
              Call REPLA(Z$, TIVA$, 7, 18)
              Call REPLA(Z$, NUCHE$, 25, 14)
                FEMI = CVINT(!FECHEMISI)
              Call REPLA(Z$, FECHATEX$(FEMI), 39, 8)
                'FEVE = CVINT(!FechAcred)
              Call REPLA(Z$, FEVECH$, 49, 8)
              Call REPLA(Z$, CSTRING$(MKD$(ICHE#), 3, 12, 2, 2), 57, 12)
                REGBA1& = !NumeComp: If REGBA1& < 1 Then REGBA1& = 123456
              Call REPLA(Z$, CSTRING$(MKS$(REGBA1&), 3, 8, 0, 2), 121, 8)
              List1.AddItem Z$
              '
            End If
          .MoveNext
          Loop
        End If
        'On Error GoTo 0
      End With
      CHETEMP.Close
      Set CHETEMP = Nothing
      '
    End If
    '
    List1.Refresh
    'Call REMPLAGRID(Me, GRID2, Picture1, Label12, List1)
    'Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
    Screen.MousePointer = 1
    '
    If TICUEPRO%((-1) * NC%) = 1 Then        ' GASTOS
       Call SIGUEORPA5
    End If
    '
End Sub
'
Sub SIGUEORPA3()
    '
    If List2.ListCount < 1 Then
        Exit Sub
    End If
    '
    Call BLANARCH("!CARGAPLI")
    For U& = 1 To List2.ListCount
        'List2.ListIndex = U& - 1
        Z$ = List2.LIST(U& - 1)
        NROX$ = Mid$(Z$, 1, 20)
        FEMII% = FECHANUM(Mid$(Z$, 23, 8))
        IPED# = CDbl(XVALO(Mid$(Z$, 32, 13)))
        FEVEI% = FECHANUM(Mid$(Z$, 47, 8))
        SALD# = CDbl(XVALO(Mid$(Z$, 56, 13)))
        ACRE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 1, 18)
        REG2& = XVALO(Mid$(Z$, 119, 8))
        IRE% = XVALO(Mid$(Z$, 127, 2))
        '
        y$ = REGBLAN$("!CARGAPLI")
        Call REPLA(y$, NROX$, 1, 20)
        Call REPLA(y$, MKI$(FEMII%), 21, 2)
        Call REPLA(y$, MKD$(IPED#), 23, 8)
        Call REPLA(y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(y$, MKD$(SALD#), 33, 8)
        Call REPLA(y$, MKD$(ACRE#), 41, 8)
        Call REPLA(y$, DOCU$, 49, 18)
        Call REPLA(y$, MKS$(REG2&), 67, 4)
        Call REPLA(y$, MKI$(IRE%), 71, 2)
        '
        Call REGAPP("!CARGAPLI", y$)
    Next U&
    Call CIERRARCH("!CARGAPLI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Aplicación de Pagos"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture2.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture2.Left))
    '
    VTB% = VENTABU%("CARGAPLI" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        List2.Clear
        For U& = 1 To ULTREG&("!CARGAPLI")
            y$ = REGLEIDO$("!CARGAPLI", U&)
            NROX$ = Mid$(y$, 1, 20)
            If TRIM$(NROX$) = "Anticipo no Aplicado" Then
               HOI% = HOY(HO$)
               ACRE# = CVD(Mid$(y$, 41, 8))
               Call REPLA(y$, MKI$(HOI%), 31, 2)
               Call REPLA(y$, MKD$(ACRE#), 23, 8)
               Call REPLA(y$, MKD$(ACRE#), 33, 8)
               Call GRAREG("!CARGAPLI", y$, U&)
            End If
            FEMI = CVI(Mid$(y$, 21, 2))
            IPED# = CVD(Mid$(y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 13, 2, 2)
            FEVE = CVI(Mid$(y$, 31, 2))
            SALD# = CVD(Mid$(y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 13, 2, 2)
            ACRE# = CVD(Mid$(y$, 41, 8))
            '
            If Abs(ACRE#) >= 0.005 Then
              If Sgn(ACRE#) <> Sgn(SALD#) Then
                Call MENSERR(24, "No se Permite Aplicación Negativa")
                GoTo 200
              End If
            End If
            If U& > 1 Then
              If Abs(ACRE#) > Abs(SALD#) Then
                Call MENSERR(24, "El Importe Acreditado Debe Ser\Inferior al Saldo del Documento")
                 GoTo 200
              End If
            End If
            ACRTEX$ = CSTRING$(MKD$(ACRE#), 4, 14, 2, 2)
            DOCU$ = Mid$(y$, 49, 18)
            REG2& = CVS(Mid$(y$, 67, 4))
            IRE% = CVI(Mid$(y$, 71, 2))
            REGTEX$ = CSTRING$(MKS$(REG2&), 3, 8, 0, 2)
            IRETEX$ = CSTRING$(MKI$(IRE%), 1, 2, 0, 2)
            '
            Z$ = Space$(128)
            Call REPLA(Z$, NROX$, 1, 20)
            Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(Z$, IMPTEX$, 32, 13)
            Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
            Call REPLA(Z$, SALTEX$, 56, 13)
            Call REPLA(Z$, ACRTEX$, 70, 14)
            Call REPLA(Z$, REGTEX$, 119, 8)
            Call REPLA(Z$, IRETEX$, 127, 2)
            List2.AddItem Z$
            '
        Next U&
    End If
    '
    SUMACRE# = 0
    NumeroProveedor& = XVALO(Text1)
    For U& = 1 To List2.ListCount
        Z$ = List2.LIST(U& - 1)
        
        ' *** CALCULA RETENCION DE IVA PARA FACTURAS 'M'
        
        ImporteRetenido# = 0
        IndiceComprobante% = U& - 1
        Renglon$ = List2.LIST(IndiceComprobante%)
        ImporteACancelar# = XVALO(TRIM$(Mid$(Renglon$, 70, 14)))
        If Abs(ImporteACancelar#) > 0.005 Then
            Comprobante$ = TRIM$(Mid$(Renglon$, 1, 18))
        
            ' VERIFICA QUE SE TRATE DE UN COMPROBANTE 'M'
            'If (Mid$(Comprobante$, 1, 2) = "FC" Or Mid$(Comprobante$, 1, 2) = "ND") And Mid$(Comprobante$, 4, 1) = "M" Then
            If Mid$(Comprobante$, 4, 1) = "M" Then
                CondicionSql$ = "NUPROVE=" & CStr(NumeroProveedor&) & " AND CODICOM_LAR='" & Comprobante$ & "' AND TIPOMOVIM='COMPRA'"
                ImporteNetoComprobante# = XVALO(VALOBADA("CAJABANC", "SUM(VALABSCOMP)", CondicionSql$))
                
                ' VALIDA QUE NO SE HAYA RETENIDO ANTERIORMENTE Y QUE EL VALOR NETO DEL COMPROBANTE SEA MAYOR O IGUAL AL MINIMO IMPONIBLE
                If CantRegistros&("CAJABANC", "TIPOMOVIM='RTIVA' AND NOTAS LIKE '%" & Comprobante$ & "%'") = 0 And (ImporteNetoComprobante# >= MinimoImponibleRetencionIva# Or Mid$(Comprobante$, 1, 2) = "NC") Then
                
                    CondicionSql$ = "NUPROVE=" & CStr(NumeroProveedor&) & " AND CODICOM_LAR='" & Comprobante$ & "'"
                    CondicionSql$ = CondicionSql$ & " AND TIPOMOVIM='IVA' and VARIMOVIM<>'Aduana' and VARIMOVIM<>'Per-Iva' and VARIMOVIM<>'Per-Gan'"
                    ImporteRetenido# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)-SUM(IHABECOMP)", CondicionSql$))
                End If
            
            End If
        End If
        ' ACTUALIZA LA LISTA AGREGANDO EL IMPORTE RETENIDO DEL COMPROBANTE
        Z1$ = Left$(Z$ + Space$(256), 256)
        Call REPLA(Z1$, FORMATNUM$(ImporteRetenido#, "F12.2"), 177, 12)
        Z$ = Z1$
        ILISTA& = U& - 1
        List2.AddItem Z$, ILISTA&
        List2.RemoveItem ILISTA& + 1
        List2.Refresh
        
        ' *** FIN RETENCION DE IVA
        
        SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(Z$, 70, 14)))
    Next U&
    '
    Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label18.Refresh
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label10.Refresh
    Label20.Caption = TRIM$(CSTRING$(MKD$(XVALO(Label18.Caption) - XVALO(Label14.Caption) - XVALO(Label45.Caption) - XVALO(Label12.Caption) - XVALO(Label21.Caption) - XVALO(Label30.Caption) - XVALO(Label38.Caption) - XVALO(Label49.Caption)), 4, 14, 2, 2))
    Label20.Refresh
    '
    List2.Refresh
    VACON% = 1
    '
    'Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
    '
End Sub
'

Sub SIGUEORPA4()
   '
   FEMIOP% = FECHANUM(Text9)
   VACON% = 1
   List4.Clear
10 For U& = 1 To List1.ListCount
     Z$ = List1.LIST(U& - 1)
     If XVALO(Mid$(Z$, 70, 14)) >= 0.005 Then
       List4.AddItem Z$
     End If
   Next U&
   List1.Clear
   For U& = 1 To List4.ListCount
     List1.AddItem List4.LIST(U& - 1)
   Next U&
   'Call REMPLAGRID(Me, GRID2, Picture1, Label12, List1)
   List1.Refresh
   '
   SUMAVAL# = XVALO(Label21.Caption)
   For U& = 1 To ULTREG&("!CARMEPA")
      X$ = REGLEIDO$("!CARMEPA", U&)
      ICHE# = CVD(Mid$(X$, 31, 8))
      COVA% = CVI(Left$(X$, 2))
      TIVA$ = ECOARCH$("MEPAGO", MKI$(COVA%))
      NUCHE$ = TRIM$(Mid$(X$, 19, 12))
      '
      Z$ = Space$(128)
      Call REPLA(Z$, "Prop", 1, 4)
      Call REPLA(Z$, TIVA$, 7, 18)
      Call REPLA(Z$, NUCHE$, 25, 14)
      FEMI = CVI(Mid$(X$, 39, 2))
        If FEMI < 1 Then
          FEMI = FEMIOP%
          Call REPLA(X$, MKI$(FEMIOP%), 39, 2)
          Call GRAREG("!CARMEPA", X$, U&)
        End If
      Call REPLA(Z$, FECHATEX$(FEMI), 39, 8)
      FEVE = CVI(Mid$(X$, 41, 2))
        If FEVE < 1 Then
          FEVE = FEMIOP%
          Call REPLA(X$, MKI$(FEMIOP%), 41, 2)
          Call GRAREG("!CARMEPA", X$, U&)
        End If
      Call REPLA(Z$, FECHATEX$(FEVE), 49, 8)
      Call REPLA(Z$, CSTRING$(MKD$(ICHE#), 3, 12, 2, 2), 57, 12)
      Call REPLA(Z$, CSTRING$(MKD$(ICHE#), 3, 14, 2, 2), 70, 14)
      Call REPLA(Z$, CSTRING$(MKI$(COVA%), 1, 6, 0, 2), 115, 6)
      List1.AddItem Z$
      SUMAVAL# = SUMAVAL# + ICHE#
      Label21.Caption = TRIM$(CSTRING$(MKD$(SUMAVAL#), 4, 14, 2, 2))
    '
   Next U&
   '
   List1.Refresh
   'Call REMPLAGRID(Me, GRID2, Picture1, Label12, List1)
   'Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
   Label21.Refresh
   Label20.Caption = TRIM$(CSTRING$(MKD$(XVALO(Label18.Caption) - XVALO(Label14.Caption) - XVALO(Label45.Caption) - XVALO(Label12.Caption) - XVALO(Label30.Caption) - XVALO(Label38.Caption) - XVALO(Label49.Caption) - SUMAVAL#), 4, 14, 2, 2))
   Label20.Refresh
   '
   U& = 0
20 If U& < List2.ListCount Then
      'List2.ListIndex = U&
      Z$ = List2.LIST(U&)  ' TEXT
      If Abs(XVALO(Mid$(Z$, 70, 14))) < 0.005 Then
           List2.RemoveItem U&
         Else
           U& = U& + 1
      End If
      GoTo 20
   End If
   '
   VACON% = 0
   '
   Command5.Enabled = True
   '
End Sub
'

'
Sub SIGUEORPA5()
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Distribución de Pagos"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture2.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture2.Left))
    If Control("FACOMPRA", "EDITARCC") = "NO" Then 'DESHABILITA LA EDICION DE CENTRO DE COSTOS
     ATRI$ = ATRI$ + "/POSEDIT(4)=2"
     ATRI$ = ATRI$ + "/POSEDIT(5)=2"
    End If

    '
305 VTB% = VENTABU%("CARGAST" + ATRI$)
    If VTB% < 0 Then
       If COMALTER%("Realmente Desea Cancelar\la Distribución Ingresada ?\\No, Continuar\Si, Cancelar") <> 2 Then
           GoTo 305
       End If
       Call CIERRARCH("!CARGAST")
       Call BLANARCH("!CARGAST")
    End If
    '
    SUMACRE# = 0
    V& = 0: UF& = ULTREG&("!CARGAST")
    For U& = 1 To UF&
       X$ = REGLEIDO$("!CARGAST", U&)
       If ECOARCH$("ECUECON", Left$(X$, 12)) <> "" Then
       CUE% = CVI(Mid$(REGACT$("ECUECON"), 41, 2))
       Call REPLA(X$, MKI$(CUE%), 21, 2)
       SUMACRE# = SUMACRE# + CVD(Mid$(X$, 13, 8))
       V& = V& + 1
       Call GRAREG("!CARGAST", X$, V&)
       End If
    Next U&
    '
    Call SETULTREG("!CARGAST", V&)
    If V& < UF& Then
       Call COMUNI("Se han Detectado Importes Mal Aplicados.\  \Revise y Confirme la Distribución de Pagos.")
       GoTo 305
    End If
    '
    Label18.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label18.Refresh
    Label20.Caption = TRIM$(CSTRING$(MKD$(XVALO(Label18.Caption) - XVALO(Label14.Caption) - XVALO(Label45.Caption) - XVALO(Label12.Caption) - XVALO(Label21.Caption) - XVALO(Label30.Caption) - XVALO(Label38.Caption) - XVALO(Label49.Caption)), 4, 14, 2, 2))
    Label20.Refresh
    '
    If List2.ListCount > 0 Then
        List2.ListIndex = 0
    End If
    Z$ = List2.TEXT
    IACRE$ = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Call REPLA(Z$, IACRE$, 70, 14)
    VACON% = 1
    List2.AddItem Z$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    List2.Refresh
    List2.ListIndex = 0
    VACON% = 0
    Screen.MousePointer = 1
    '
    'Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
    '
End Sub
'
Sub GRAORPASTA(OKX%)
    '
    HO% = HOY(HOS$)
    NRO& = XVALO(Mid$(Text11.TEXT, 4))
    NC% = XVALO(Text1.TEXT)
    NC1% = (-1) * NC%
    FC% = FECHANUM(Text9.TEXT)
    '
    VA% = 1: MO$ = "$"
    TOFAC#(0) = XVALO(Label14.Caption) + XVALO(Label45.Caption) + XVALO(Label12.Caption) + XVALO(Label21.Caption) + XVALO(Label30.Caption) + XVALO(Label38.Caption) + XVALO(Label49.Caption)
    TOFAC#(1) = 0
    TOFAC#(2) = TOFAC#(0)
    '
    'AGREGADO ASEMBLI-MULTIMONEDA COMIENZO
    SuBalan# = 0
    If MODOMONEI% = 2 Then
      IAJU% = CUEINT%(Control$("", "CUEIVA-7"))
      If IAJU% < 1 Then
         Call MENSERR(24, "Inposible Grabar esta Orden de Pago.\   \Falta Definir Cuenta Contable\Correspondiente a Diferencia de Cambio.")
         If DERACCE%("SETUPACR", "Configuración de Funcionamiento ACREEDORES") = 2 Then
            OPFACOM07.Show 1
         End If
         Call COMUNI("Re-Inicie Ahora la Aplicación")
         Call FINAL("")
      End If
      '
      If MONEMI% <= 1 Then
        For U& = 1 To List2.ListCount
          Z$ = List2.LIST(U& - 1)
          DOCU$ = Left$(Z$, 20)
          If DOCU$ <> "Anticipo no Aplicado" Then
             CANCE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
             MEMI% = XVALO(Mid$(Z$, 147, 2))
             SALDOPESI# = XVALO(Mid$(Z$, 56, 13))
             SALDODOLI# = 0
             CANCEDOL# = 0
             If MEMI% > 1 Then
                CANCEDOL# = ROUND(CANCE# / TICAMONE(2))
                SALDODOLI# = XVALO(Mid$(Z$, 149, 12))
             End If
             If CANCEDOL# > SALDODOLI# + 0.1 Then
                Call MENSERR(24, "Imposible Grabar.\Las Aplicaciones en Moneda Extranjera\Exceden el Saldo Pendiente\del Comprobante '" + TRIM$(DOCU$) + "'.")
                List1.Clear
                Label21 = ""
                Exit Sub
             End If
             If SALDOPESI# - CANCE# <= 0.005 Then ' ANULARÁ SALDO EN PESOS
               If SALDODOLI# - CANCEDOL# >= 0.005 Then ' MANTENDRÁ SALDO EN DOLARES
                 TTXX$ = "Esta Transacción Anulará el Saldo en Pesos\del Comprobante '" + TRIM$(DOCU$) + "' Manteniendo\un Saldo no Cancelado en Moneda Extranjera.\  \Se Recomienda Generar O/P en Moneda Extranjera."
                 Call MENSERR(24, TTXX$)
                 List1.Clear
                 Label21 = ""
                 Exit Sub
               End If
             End If
          End If
        Next U&
      End If
      '
    End If
   'AGREGADO ASEMBLI-MULTIMONEDA FIN
    '
    GoSub VALIDACUENRET
    GoSub ENTRAOBSERVA
    'GoSub VERINUMERO
    '
    ' verificar que los cheques todavía están en cartera
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    
    Dim CHETEMP1 As ADODB.Recordset
    'SQLX$ = "SELECT * FROM CAJABANC "
    'SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
    'SQLX$ = SQLX$ + "AND Status = 1 "
    'SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
    'Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    GoSub VERICHECAR
    '
    ' verificar que las facturas estan aun pendientes de cancelación
    Dim FacPenTemp As ADODB.Recordset
    'SQLX$ = "SELECT * FROM SACREPEN "
    'SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    'Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    GoSub VERIFACPEN
    '
    GoSub ARMANTICIPO
    GoSub VERINUMERO
    GoSub ARMANUMEROS
    '
    Screen.MousePointer = 11
    NOCLOSE% = 1
    COTIMONEM_ANTERIOR# = COTIMONEM#
    COTIMONEM# = 0
    GoSub GRABATORPAG
    GoSub GRABACHENTRE
    COTIMONEM# = COTIMONEM_ANTERIOR#
    '
    OrdenDePagoActiva$ = RECILAR$
    '
    ' Ret I.Brutos Buenos Aires
    If XVALO(Label14) >= 0.005 Then
       NURETBA& = 1
       If ULTREG&("SURETIB") > 0 Then
         Z1$ = REGLEIDO$("SURETIB", ULTREG&("SURETIB"))
         NURETBA& = 1 + CVS(Mid$(Z1$, 3, 4))
       End If
       If Control$("", "NURETGAN") = "MANUAL" Then
795      NURETBA& = XVALO(InputBox$("Numero de Comprobante I.B.", "Retencion Ingresos Brutos", TRIM$(Str$(NURETBA&))))
         If NURETBA& < 1 Then GoTo 795
       End If
       Label25 = TRIM$(Str$(NURETBA&))
       '
       Z1$ = REGBLAN$("SURETIB")
       Call REPLA(Z1$, MKI$(FC%), 1, 2)
       Call REPLA(Z1$, MKS$(NURETBA&), 3, 4)
       Call REPLA(Z1$, MKS$(NRO&), 7, 4)
       Call REPLA(Z1$, rasocli$(NC1%), 21, 30)
       Call REPLA(Z1$, CuitCli$(NC1%), 51, 14)
       Call REPLA(Z1$, NIBRCLI$(NC1%), 65, 14)
       Call REPLA(Z1$, MKD$(XVALO(Label23)), 79, 8)
       Call REPLA(Z1$, MKD$(XVALO(Label14)), 87, 8)
       If MONEMI% > 1 Then
          Call REPLA(Z1$, MKD$(TICAMONE(MONEMI%) * XVALO(Label14)), 87, 8)
       End If
      '\\\OT-05-0285-WAR-12/05/05///
       Call REPLA(Z1$, MKI$(NC%), 95, 2)
       '\\\OT-05-0285-WAR-FIN///
       Call REGAPP("SURETIB", Z1$)
       Call CIERRARCH("SURETIB")
       '
       CoComCor$ = NRECIB$
       CoComLar$ = RECILAR$
'       CodiEmp% = 0
'       CodiSuc% = SUC%
       NumeClie = 0
       NumeProv% = NC%
       FechEmi = CVDAT(FC%)
         FEX = FC%
       Referen$ = rasocli$(NC1%) + " - " + FECHATEX$(FEX)
       FeContab = CVDAT(FC%)
       '
       FeAcredi = CVDAT(FC%)
         Tipo% = XVALO(Label24)
       CuentaCo% = CUEMEPAG%(Tipo%)
       TiMovi$ = "RTIB"
       OpMovi$ = "Bs-As"
       VaMovi$ = ""
       NumeroRetencion& = NURETBA&
       StatInit = 0
       'Descrip$ = "Ret.IB.Bs.As." + TRIM$(Str$(NURETBA&))
       Descrip$ = "Ret.I.Brutos." + TRIM$(Str$(NURETBA&))
       NumeTal = COBRACT%
       NumeCom = NRO&
       ImpoDeb# = (-1) * Abs(XVALO(Label14))
       If MONEMI% > 1 Then
          ImpoDeb# = (-1) * Abs(TICAMONE(MONEMI%) * XVALO(Label14))
       End If
       Observa$ = "O.Pago " + TRIM$(Str$(NRO&)) + " - " + rasocli$(NC1%)
       Call GRACABAN
    End If
    '
    
    ' Ret I.Brutos CABA
    If XVALO(Label38) >= 0.005 Then
       NumeroDeRetencion& = XVALO(VALOBADA("CAJABANC", "MAX(NRODERETENCION)", "NUPROVE>0 AND TIPOMOVIM='RTIB' AND OPCIMOVIM='CABA'")) + 1
       
       If Control$("", "NURETGAN") = "MANUAL" Then
790      NumeroDeRetencion& = XVALO(InputBox$("Numero de Comprobante I.B.", "Retencion Ingresos Brutos", CStr(NumeroDeRetencion&)))
         If NumeroDeRetencion& < 1 Then GoTo 790
       End If
       '
       Label41.Caption = CStr(NumeroDeRetencion&)
       For IndiceComprobantes% = 0 To List2.ListCount - 1
        
            ImporteRetenido# = XVALO(Mid$(List2.LIST(IndiceComprobantes%), 165, 12))
            If ImporteRetenido# > 0.005 Then
                CoComCor$ = NRECIB$
                CoComLar$ = RECILAR$
                NumeClie = 0
                NumeProv% = NC%
                FechEmi = CVDAT(FC%)
                  FEX = FC%
                Referen$ = rasocli$(NC1%) + " - " + FECHATEX$(FEX)
                FeContab = CVDAT(FC%)
                '
                FeAcredi = CVDAT(FC%)
                  Tipo% = XVALO(Label39)
                CuentaCo% = CUEMEPAG%(Tipo%)
                TiMovi$ = "RTIB"
                OpMovi$ = "CABA"
                VaMovi$ = ""
                StatInit = 0
                Descrip$ = "Ret.I.Brutos." + TRIM$(Str$(NumeroDeRetencion&))
                NumeroRetencion& = NumeroDeRetencion&
                NumeTal = COBRACT%
                NumeCom = NRO&
                ImpoDeb# = (-1) * ImporteRetenido#
                If MONEMI% > 1 Then
                   ImpoDeb# = (-1) * TICAMONE(MONEMI%) * ImporteRetenido#
                End If
                Observa$ = "O.Pago " + TRIM$(Str$(NRO&)) + " - " + rasocli$(NC1%)
                ObservaFactu$ = TRIM$(Mid$(List2.LIST(IndiceComprobantes%), 1, 18))
                Call GRACABAN
            End If
            
        Next IndiceComprobantes%
    End If
    
    ' Ret I.Brutos Misiones
    If XVALO(Label45) >= 0.005 Then
       NumeroDeRetencion& = XVALO(VALOBADA("CAJABANC", "MAX(NRODERETENCION)", "NUPROVE>0 AND TIPOMOVIM='RTIB' AND OPCIMOVIM='Misiones'")) + 1
       
       If Control$("", "NURETGAN") = "MANUAL" Then
785      NumeroDeRetencion& = XVALO(InputBox$("Numero de Comprobante I.B.", "Retencion Ingresos Brutos", TRIM$(Str$(NumeroDeRetencion&))))
         If NumeroDeRetencion& < 1 Then GoTo 785
       End If
       '
       NumeroRetencionIBMisiones& = NumeroDeRetencion&
       '
       CoComCor$ = NRECIB$
       CoComLar$ = RECILAR$
       NumeClie = 0
       NumeProv% = NC%
       FechEmi = CVDAT(FC%)
         FEX = FC%
       Referen$ = rasocli$(NC1%) + " - " + FECHATEX$(FEX)
       FeContab = CVDAT(FC%)
       '
       FeAcredi = CVDAT(FC%)
       CuentaCo% = CUEMEPAG%(MedioPagoIBMisiones%)
       TiMovi$ = "RTIB"
       OpMovi$ = "Misiones"
       VaMovi$ = ""
       NumeroRetencion& = NumeroDeRetencion&
       StatInit = 0
       Descrip$ = "Ret.I.Brutos." + TRIM$(Str$(NumeroDeRetencion&))
       NumeTal = COBRACT%
       NumeCom = NRO&
       ImpoDeb# = (-1) * Abs(XVALO(Label45))
       If MONEMI% > 1 Then
          ImpoDeb# = (-1) * Abs(TICAMONE(MONEMI%) * XVALO(Label45))
       End If
       BasImpo# = MontoSujetoIBMisiones#
       Observa$ = "O.Pago " + TRIM$(Str$(NRO&)) + " - " + rasocli$(NC1%)
       Call GRACABAN
    End If

    ' Ret I.Brutos TUCUMAN
    If XVALO(Label12) >= 0.005 Then
       NumeroDeRetencion& = XVALO(VALOBADA("CAJABANC", "MAX(NRODERETENCION)", "NUPROVE>0 AND TIPOMOVIM='RTIB' AND OPCIMOVIM='TUCUMAN'")) + 1
       
       If Control$("", "NURETGAN") = "MANUAL" Then
787      NumeroDeRetencion& = XVALO(InputBox$("Numero de Comprobante I.B.", "Retencion Ingresos Brutos", TRIM$(Str$(NumeroDeRetencion&))))
         If NumeroDeRetencion& < 1 Then GoTo 787
       End If
       '
       NumeroRetencionIBTucuman& = NumeroDeRetencion&
       '
       CoComCor$ = NRECIB$
       CoComLar$ = RECILAR$
       NumeClie = 0
       NumeProv% = NC%
       FechEmi = CVDAT(FC%)
         FEX = FC%
       Referen$ = rasocli$(NC1%) + " - " + FECHATEX$(FEX)
       FeContab = CVDAT(FC%)
       '
       FeAcredi = CVDAT(FC%)
       CuentaCo% = CUEMEPAG%(MedioPagoIBTucuman%)
       TiMovi$ = "RTIB"
       OpMovi$ = "TUCUMAN"
       VaMovi$ = ""
       NumeroRetencion& = NumeroDeRetencion&
       StatInit = 0
       Descrip$ = "Ret.I.Brutos." + TRIM$(Str$(NumeroDeRetencion&))
       NumeTal = COBRACT%
       NumeCom = NRO&
       ImpoDeb# = (-1) * Abs(XVALO(Label12))
       If MONEMI% > 1 Then
          ImpoDeb# = (-1) * Abs(TICAMONE(MONEMI%) * XVALO(Label2))
       End If
       BasImpo# = MontoSujetoIBTucuman#
       Observa$ = "O.Pago " + TRIM$(Str$(NRO&)) + " - " + rasocli$(NC1%)
       Call GRACABAN
    End If


    ' Retencion de IVA
    If XVALO(Label49) >= 0.005 Then
       NumeroDeRetencion& = XVALO(VALOBADA("CAJABANC", "MAX(NRODERETENCION)", "NUPROVE>0 AND TIPOMOVIM='RTIVA'")) + 1
       
       If Control$("", "NURETGAN") = "MANUAL" Then
786      NumeroDeRetencion& = XVALO(InputBox$("Numero de Comprobante Ret, Iva.", "Retencion IVA", TRIM$(Str$(NumeroDeRetencion&))))
         If NumeroDeRetencion& < 1 Then GoTo 786
       End If
       '
       NumeroRetencionIva& = NumeroDeRetencion&
       For IndiceComprobantes% = 0 To List2.ListCount - 1
        
            ImporteRetenido# = XVALO(Mid$(List2.LIST(IndiceComprobantes%), 177, 12))
            If ImporteRetenido# > 0.005 Then
                CoComCor$ = NRECIB$
                CoComLar$ = RECILAR$
                NumeClie = 0
                NumeProv% = NC%
                FechEmi = CVDAT(FC%)
                  FEX = FC%
                Referen$ = rasocli$(NC1%) + " - " + FECHATEX$(FEX)
                FeContab = CVDAT(FC%)
                '
                FeAcredi = CVDAT(FC%)
                CuentaCo% = MedioPagoRetencionIva%
                TiMovi$ = "RTIVA"
                OpMovi$ = ""
                VaMovi$ = ""
                StatInit = 0
                Descrip$ = "Ret.IVA Nro. " + TRIM$(Str$(NumeroDeRetencion&))
                NumeroRetencion& = NumeroDeRetencion&
                NumeTal = COBRACT%
                NumeCom = NRO&
                ImpoDeb# = (-1) * Abs(ImporteRetenido#)
                If MONEMI% > 1 Then
                   ImpoDeb# = (-1) * TICAMONE(MONEMI%) * ImporteRetenido#
                End If
                Observa$ = "O.Pago " + TRIM$(Str$(NRO&)) + " - " + rasocli$(NC1%)
                ObservaFactu$ = TRIM$(Mid$(List2.LIST(IndiceComprobantes%), 1, 18))
                Call GRACABAN
            End If
            
        Next IndiceComprobantes%
    End If
    
    ' Retencion de Ganancias
    If XVALO(Label30) >= 0.005 Then
       NURETGA& = 1
       FEXU = FC%
       ANX0$ = Mid$(FECHATEX$(FEXU), 7, 2)
       
       NURETGA& = XVALO(VALOBADA("CAJABANC", "MAX(NRODERETENCION)", "NUPROVE>0 AND TIPOMOVIM='RTGAN'")) + 1
       
       If ULTREG&("SUREGAN") > 0 And NURETGA& = 1 Then
         Call FILESORT("SUREGAN", "", 1, 2, "ASC")
         Call CIERRARCH("SUREGAN")
         FINI& = ULTREG&("SUREGAN")
         INII& = FINI& - 128: If INII& < 1 Then INII& = 1
         For KKII& = INII& To FINI&
           Z1$ = REGLEIDO$("SUREGAN", KKII&)
           FEXU = CVI(Left$(Z1$, 2))
           ANXX$ = Mid$(FECHATEX$(FEXU), 7, 2)
           If ANXX$ = ANX0$ Then
             NURETGA1& = 1 + CVS(Mid$(Z1$, 3, 4))
             If NURETGA1& > NURETGA& Then NURETGA& = NURETGA1&
           End If
         Next KKII&
       End If
       
       If Control$("", "NURETGAN") = "MANUAL" Then
796      NURETGA& = XVALO(InputBox$("Numero de Comprobante Ganancias", "Retencion Ganancias", TRIM$(Str$(NURETGA&))))
         If NURETGA& < 1 Then GoTo 796
       End If
       Label34 = TRIM$(Str$(NURETGA&))
       '
       Z1$ = REGBLAN$("SUREGAN")
       Call REPLA(Z1$, MKI$(FC%), 1, 2)
       Call REPLA(Z1$, MKS$(NURETGA&), 3, 4)
       Call REPLA(Z1$, MKS$(NRO&), 7, 4)
       Call REPLA(Z1$, rasocli$(NC1%), 21, 30)
       Call REPLA(Z1$, CuitCli$(NC1%), 51, 14)
       Call REPLA(Z1$, NIBRCLI$(NC1%), 65, 14)
       Call REPLA(Z1$, MKD$(XVALO(Label23)), 79, 8)
       'Call REPLA(Z1$, MKD$(XVALO(Label14)), 87, 8)
       Call REPLA(Z1$, MKD$(XVALO(Label30)), 87, 8) ' VER ANTES ERA LABEL14
       If MONEMI% > 1 Then
          Call REPLA(Z1$, MKD$(TICAMONE(MONEMI%) * XVALO(Label30)), 87, 8)
       End If

        RECORD$ = FILTERGETRECORD$("DARGAPRO", 1, MKI$(Abs(NC1%)))
        CATIB11% = Asc(Mid$(RECORD$, 3, 1))
        RECORD$ = FILTERGETRECORD$("TACRETG", 1, Chr$(CATIB11%))
        CODIGOIMP& = CVI(Mid$(RECORD$, 51, 2))
        CODIGOREG& = CVI(Mid$(RECORD$, 53, 2))
       Call REPLA(Z1$, MKI$(CODIGOIMP&), 95, 2)
       Call REPLA(Z1$, MKI$(CODIGOREG&), 97, 2)
       Call REGAPP("SUREGAN", Z1$)
            FINN& = ULTREG&("SUREGAN")
       Call SETULTREG("SUREGAN", FINN&)
       Call CIERRARCH("SUREGAN")
       '
       CoComCor$ = NRECIB$
       CoComLar$ = RECILAR$
'       CodiEmp% = 0
'       CodiSuc% = SUC%
       NumeClie = 0
       NumeProv% = NC%
       FechEmi = CVDAT(FC%)
         FEX = FC%
       Referen$ = rasocli$(NC1%) + " - " + FECHATEX$(FEX)
       FeContab = CVDAT(FC%)
       '
       FeAcredi = CVDAT(FC%)
         Tipo% = XVALO(Label32)
       CuentaCo% = CUEMEPAG%(Tipo%)
       TiMovi$ = "RTGAN"
       OpMovi$ = ""
       VaMovi$ = ""
       NumeroRetencion& = NURETGA&
       StatInit = 0
       Descrip$ = "Ret.Ganancias " + TRIM$(Str$(NURETGA&))
       NumeTal = COBRACT%
       'NumeCom = NURETGA&
       BasImpo# = XVALO(Label18)
       ImpoDeb# = (-1) * Abs(XVALO(Label30))
       If MONEMI% > 1 Then
          ImpoDeb# = (-1) * Abs(TICAMONE(MONEMI%) * XVALO(Label30))
       End If
       Observa$ = "O.Pago " + TRIM$(Str$(NRO&)) + " - " + rasocli$(NC1%)
       CodigoImpuesto = CODIGOIMP&
       CodigoRegimen = CODIGOREG&
       Call GRACABAN
       '
       CodigoImpuesto = 0
       CodigoRegimen = 0
    End If
    '
    OKX% = 3          ' grabo cheques de terceros
    OKX% = 4          ' grabo cheques emitidos
    '
800 ' ******************** aqui blanquear facturas canceladas
    ' ******************** tambien grabar archivo APLIPAG.DAT
    '
    If TICUEPRO%(NC1%) = 0 Then
       '
       CoComCor$ = NRECIB$
       CoComLar$ = RECILAR$
 '      CodiEmp% = 0
 '      CodiSuc% = SUC%
       NumeClie = 0
       NumeProv% = NC%
       FechEmi = CVDAT(FC%)
       RefApli$ = "Orden de Pago " + NRECIB$
       '
       XX% = 0
       For U& = 1 To List2.ListCount
         Z$ = List2.LIST(U& - 1)
         DOCU$ = Left$(Z$, 20)
         If DOCU$ <> "Anticipo no Aplicado" Then
              CANCE# = CDbl(XVALO(Mid$(Z$, 70, 14))) ' 800
              MEMI% = XVALO(Mid$(Z$, 147, 2))
              CANCEDOL# = CANCE# / TICAMONE(2)        '100
              'AGREGADO ASEMBLI MULTIMONEDA
              If MONEMI% > 1 Then
                 CANCEDOL# = CANCE#  '100
                 CANCE# = ROUND(CANCEDOL * TICAMONE(MONEMI%)) '800
               End If
              'AGREGADO ASEMBLI MULTIMONEDA FIN
              If Abs(CANCE#) >= 0.005 Then  '800
                DOCUCA$ = TRIM$(Left$(Z$, 18))
                ICAN# = CANCE#  ' 800
                SICAN% = Sgn(CANCE#)
                ICANDOL# = CANCEDOL#  ' 100
                TOTAPLI# = 0
                '
                DocOrCor$ = ""
                TotOrig# = 0
                IvaOrig# = 0
                DiferenciaPorTipoCambio# = 0
                '
                Call APERBASDAT("CABAN")
                Dim Sacrepen As ADODB.Recordset
                Set Sacrepen = New ADODB.Recordset
                strsql = "Select * from SacrePen "
                strsql = strsql & " where NuProve = " & NumeProv%
                strsql = strsql & " and CODICOM_LAR = '" & DOCUCA$ & "'"
                strsql = strsql & " and ( ABS(SALDACRE) > 0 OR ABS(SALDDEBE) > 0 ) "
                '
25              On Error Resume Next
                Sacrepen.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           
                If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                   On Error GoTo 0
                   Call CapturaErrorOpen
                   GoTo 25
                End If
                On Error GoTo 0
               With Sacrepen
                If Not (.EOF And .BOF) Then
                  .MoveFirst
                  Do While Not (.EOF)
850                 PARCAN# = ICAN#  '  : If !SALDACRE < ICAN# Then PARCAN# = !SALDACRE ' 800
                    PARCANDOL# = ICANDOL# ' 100
                    ' *** CALCULA LA DIFERENCIA EN PESOS CUANDO EL TIPO DE CAMBIO DE LA O/PAGO ES MAYOR AL DE LA FACTURA
                    If PARCAN# > !SaldAcre Then PARCAN# = !SaldAcre ' 800
                    If PARCANDOL# > !SaldAcre_DOLA Then PARCANDOL# = !SaldAcre_DOLA  '100
                    If SICAN% * (!SaldAcre - !SaldDebe) < SICAN% * ICAN# Then PARCAN# = !SaldAcre - !SaldDebe
                    If SICAN% * (!SaldAcre_DOLA - !SaldDebe_DOLA) < SICAN% * ICANDOL# Then PARCANDOL# = !SaldAcre_DOLA - !SaldDebe_DOLA
                    '
                    MonedaFactura% = MONEMI%
                    If MonedaFactura% > 1 Then
                       ' SI ICAN# < PARCAN#, QUIERE DECIR QUE FALTAN PESOS EN LA FACTURA PARA COMPENZAR LA O-PAGO. ESTO ESTA REFLEJADO
                       ' EN EL SALDO EN PESOS PENDIENTE QUE TIENE LA FACTURA EN EL SACREPEN. DICHA FALTA SE CONSIDERA COMO UNA
                       ' DIFERENCIA DE CAMBIO Y POR ENDE UN AJUSTE DEL LADO DEL DEBE.
                       If ICAN# > PARCAN# Then
                           DiferenciaPorTipoCambio# = DiferenciaPorTipoCambio# + (ICAN# - PARCAN#)
                       Else
                           ' *** CALCULA LA DIFERENCIA EN PESOS CUANDO EL TIPO DE CAMBIO DE LA O/PAGO (PARCAN#) ES MENOR AL DE LA FACTURA (AplicacionEnPesosSegunFactura#).
                           ' *** PARA ESO TOMA LOS DOLARES A CANCELAR (PARCANDOL#), LOS CONVIERTE A PESOS SEGUN EL TIPO DE CAMBIO
                           ' *** DE LA FACTURA Y LO COMPARA CON LOS PESOS DE LA ORDEN DE PAGO (PARCAN#)
                           TipoDeCambioFactura# = XVALO(VALOBADA("CAJABANC", "Tipo_Cam", "CODICOM_LAR='" & DOCUCA$ & "' AND NuProve= " & NumeProv%))
                           If TipoDeCambioFactura# > 0.00005 Then AplicacionEnPesosSegunFactura# = PARCANDOL# * TipoDeCambioFactura#  ' CONVIERTE LOS DOLARES DE LA O-P EN PESOS S/TIPO DE CAMBIO DE LA FACTURA
                           If (AplicacionEnPesosSegunFactura# - PARCAN#) >= 0.01 Then DiferenciaPorTipoCambio# = DiferenciaPorTipoCambio# + (PARCAN# - AplicacionEnPesosSegunFactura#)
                       End If
                    End If

                    If PARCAN# > 0 Then
                        !SaldAcre = !SaldAcre - PARCAN#  ' ES 817 - 800 QUEDA EN 17
                      Else
                        !SaldDebe = !SaldDebe + PARCAN#
                    End If
                    If PARCANDOL# > 0 Then
                        !SaldAcre_DOLA = !SaldAcre_DOLA - PARCANDOL# 'QUEDA EN 0
                      Else
                        !SaldDebe_DOLA = !SaldDebe_DOLA + PARCANDOL#
                    End If
                    .Update
                    TOTAPLI# = TOTAPLI# + PARCAN#   ' SUMA APLICACIONES -  agregado por tema bimonetario 22/03
                    '
                    DocOrCor$ = !CODICOM_COR
                    TotOrig# = !IBruComp ' AQUI TIENE LOS 817
                    IIVAORIG# = 0
                    On Error Resume Next
                    IvaOrig# = !IIvaComp
                    On Error GoTo 0
                    '
                    ICAN# = ICAN# - PARCAN#
                    ICANDOL# = ICANDOL# - PARCANDOL#
                    If Abs(ICAN#) >= 0.005 Or ICANDOL# >= 0.005 Then
                       .MoveNext
                     Else
                       Exit Do
                    End If
                  Loop
                 End If
                End With
                Sacrepen.Close
                Set Sacrepen = Nothing
                '
                SUCANCE# = SUCANCE# + TOTAPLI#
                DocOrLar$ = TRIM$(Left$(Z$, 18))
                'ImpApli# = CANCE#
                ImpApli# = TOTAPLI#
                MonOri% = MONEMISORP%(TRIM$(Left$(Z$, 18)), NC%)
                ImpApliOr# = CANCE#
                If MonOri% > 1 Then
                   ImpApliOr# = CANCE# / TICAMONE(MonOri%)
                End If
                ' *** SI SE CUMPLE TOTAPLI# < CANCE#, SIGNIFICA QUE LOS PESOS DE LA O-PAGO SON MAYOR
                ' *** A LOS PESOS SEGUN LA FACTURA GUARDADOS EN SACREPEN.
                ' *** SI EN CAMBIO Abs(DiferenciaPorTipoCambio#) > 0.005, SIGNIFICA QUE LOS PESOS DE LA O-PAGO
                ' *** SON MENOR A LOS PESOS SEGUN LA FACTURA.
                If (TOTAPLI# < CANCE# - 0.005 Or Abs(DiferenciaPorTipoCambio#) > 0.005) And MONEMI% > 1 Then
                    ' acumula la diferencia para ajuste de saldo
                    Z1$ = Left$(Z$ + Space$(256), 256)
                    DIFAPLI$ = FORMATNUM$(CANCE# - TOTAPLI#, "F12.4")
                    If Abs(DiferenciaPorTipoCambio#) >= 0.01 Then DIFAPLI$ = FORMATNUM$(DiferenciaPorTipoCambio#, "F12.4")
                    Call REPLA(Z1$, DIFAPLI$, 241, 12)
                    List2.RemoveItem U& - 1
                    List2.AddItem Z1$, U& - 1
                End If
                
                Call GRAPLIPAG
              End If
            Else
              CANCE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
              If Abs(CANCE#) > 0 Then
                IBruTot# = XVALO(Label18)
                NuorVen% = 0
                FeVenc = CVDAT(FC%)
                ImVenc# = (-1) * CANCE#
                If MONEMI% > 1 Then
                   IBruTot# = ROUND(IBruTot# * XVALO(TICAMBIS))
                   ImVenc# = ROUND(ImVenc# * XVALO(TICAMBIS))
                End If
                Call GRACREPEN
              End If
           End If
       Next U&
       '
       ' AGREGADO PARA MULTIMON ASEMBLI COMIENZO
       MONEMANT% = MONEMI%
       MONEMI% = 1
       COTIMONEM_ANTERIOR# = COTIMONEM#
       COTIMONEM# = 0
       ' graba ajuste diferencia de cambio
       CoComCor$ = "AS" + Mid$(NRECIB$, 3)
       CoComLar$ = "AS" + Mid$(RECILAR$, 3)
       FeContab = CVDAT(FC%)
       SuBalan# = 0
       SuDicam# = 0
       SuBalan0# = 0
       SuDicam0# = 0
       
       For U5& = 1 To List2.ListCount
         DIFE1# = XVALO(Mid$(List2.LIST(U5& - 1), 241, 12))
         If Abs(DIFE1#) >= 0.005 Then
         
            NumeroOrdenAjuste% = NumeroOrdenAjuste% + 1
            PREFIJO$ = CStr(NumeroOrdenAjuste%)
            Do While Len(PREFIJO$) < 4: PREFIJO$ = "0" & PREFIJO$: Loop
            CoComCor$ = "AS" + Mid$(NRECIB$, 3, 8) + Right(PREFIJO$, 2)
            CoComLar$ = "AS" + Mid$(RECILAR$, 3, 3) & PREFIJO$ & Mid$(RECILAR$, 10)
            
               CUECO% = CUECOCLI%(NC1%)
            CuentaCo% = CUECO%
            NumeClie = 0
            NumeProv% = NC%
            TiMovi = "AJUP"            ' AJUSTE DE SALDO DEUDORES
            OpMovi = "AS"              ' AJUSTE DE SALDO
            VaMovi = "CCORR"           ' CUENTA CORRIENTE
            Descrip = "Aj.Saldo - " + rasocli$(NC1%)
              FechaMovimiento% = FC%
            FechEmi = CVDAT(FechaMovimiento%)
            FeContab = CVDAT(FechaMovimiento%)
            FeAcredi = CVDAT(FechaMovimiento%)
            NumeTal = COBRACT%
            NumeCom = NRO&
            StatInit = 0
            ImpoDeb# = (-1) * DIFE1#
            Call GRACABAN
            '
            ' graba contrapartida del ajuste
            CuentaCo% = IAJU%
            TiMovi = "AJUSALP"         ' AJUSTE DE SALDO DEUDORES
            OpMovi = "AS"              ' AJUSTE DE SALDO
            VaMovi = "NO-IMP"          ' NO IMPUTADO
            Descrip = "Aj.Saldo - " + rasocli$(NC1%)
            ImpoDeb# = DIFE1#
            Call GRACABAN
            
            ' *** GRABA MOVIMIENTO EN APLIPAGO
            If DIFE1# > 0.05 Then
              ImpApli# = DIFE1#
              RefApli$ = "Aj.Saldo - " + rasocli$(NC1%)
              DocOrCor$ = "AS" + Mid$(NRECIB$, 3, 8) + Right(PREFIJO$, 2)
              DocOrLar$ = "AS" + Mid$(RECILAR$, 3, 3) & PREFIJO$ & Mid$(RECILAR$, 10) ' EL QUE APLICA
              CoComCor$ = NRECIB$
              CoComLar$ = RECILAR$                                                     ' AL QUE SE APLICA
              TotOrig# = DIFE1#
              IvaOrig# = 0
              ImpApliOr# = DIFE1#
            Else
              ImpApli# = Abs(DIFE1#)
              RefApli$ = "Aj.Saldo - " + rasocli$(NC1%)
              DocOrCor$ = "AS" + Mid$(NRECIB$, 3, 8) + Right(PREFIJO$, 2)
              DocOrLar$ = "AS" + Mid$(RECILAR$, 3, 3) & PREFIJO$ & Mid$(RECILAR$, 10) ' EL QUE APLICA
                  DOCU$ = TRIM$(Left$(List2.LIST(U5& - 1), 18))
              CoComCor$ = Mid$(DOCU$, 1, 5) + Mid$(DOCU$, 9, 1) + Mid$(DOCU$, 13, 6)
              CoComLar$ = DOCU$
              TotOrig# = Abs(DIFE1#)
              IvaOrig# = 0
              ImpApliOr# = Abs(DIFE1#)
            End If
            
            Call GRAPLIPAG
          End If
       Next U5&
       '
       If SuDicam# >= 0.005 Then
          ' graba contrapartida del ajuste
          CoComCor$ = "AS" + Mid$(NRECIB$, 3)
          CoComLar$ = "AS" + Mid$(RECILAR$, 3)
          CUECO% = CUECOCLI%(NC1%)
          CuentaCo% = CUECO%
          TiMovi = "AJUP"           ' AJUSTE DE SALDO DEUDORES
          OpMovi = "AS"             ' AJUSTE DE SALDO
          VaMovi = "CCORR"          ' CUENTA CORRIENTE
          Descrip = "Aj.Saldo - " + rasocli$(NC1%)
          ImpoDeb# = (-1) * SuDicam
          SuBalan# = SuBalan# + ImpoDeb#
          Call GRACABAN
       End If

       If SuDicam0# < -0.005 Then
          ' graba contrapartida del ajuste
          CoComCor$ = "AS" + Mid$(NRECIB$, 3, 7) + "-01"
          CoComLar$ = "AS" + Mid$(RECILAR$, 3, 3) + "0001" + Mid$(RECILAR$, 10)
          CUECO% = CUECOCLI%(NC1%)
          CuentaCo% = CUECO%
          TiMovi = "AJUP"           ' AJUSTE DE SALDO DEUDORES
          OpMovi = "AS"             ' AJUSTE DE SALDO
          VaMovi = "CCORR"          ' CUENTA CORRIENTE
          Descrip = "Aj.Saldo - " + rasocli$(NC1%)
          ImpoDeb# = Abs(SuDicam0#)
          SuBalan0# = SuBalan0# + ImpoDeb#
          Call GRACABAN
       End If
              
       If Abs(SuDicam#) >= 0.005 Then
          ' graba contrapartida del ajuste
          CUECO% = CUECOCLI%(NC1%)
          CuentaCo% = CUECO%
          TiMovi = "AJUP"           ' AJUSTE DE SALDO DEUDORES
          OpMovi = "AS"             ' AJUSTE DE SALDO
          VaMovi = "CCORR"          ' CUENTA CORRIENTE
          Descrip = "Aj.Saldo - " + rasocli$(NC1%)
          ImpoDeb# = (-1) * SuDicam
          SuBalan# = SuBalan# + ImpoDeb#
          Call GRACABAN
       End If
       '
       If MODOMONEI% = 2 Then
          If Abs(SuBalan# + SuBalan0#) >= 0.005 Then
             CuentaCo% = IAJU%
             NumeClie = 0
             NumeProv% = NC
             TiMovi = "AJUSALP"          ' AJUSTE DE SALDO DEUDORES
             OpMovi = "AS"               ' AJUSTE DE SALDO
             VaMovi = "NO-IMP"           ' CUENTA CORRIENTE
             Descrip = "Redondeo Dif.Cambio"
             FechEmi = CVDAT(FC%)
             FeAcredi = CVDAT(FC%)
             NumeTal = COBRACT%
             'CodiEmp% = 0
             'CodiSuc% = SUCUI%
             NumeCom = NRO&
             StatInit = 0
             ImpoDeb# = (-1) * (SuBalan# + SuBalan0#)
             Call GRACABAN
          End If
       End If
       '
       MONEMI% = MONEMANT%
       COTIMONEM# = COTIMONEM_ANTERIOR#
       '
    'FIN DE AGREGADO MULTIMON ASEMBLI FINAL
    End If
    '
    Call RAPLIPRO(NC%)
    '
    Call LIBARCH("CUECORR")
    '
    OKX% = 5         ' depuró facturas pendientes
    OKX% = 6         ' grabó aplicaciones
    OKX% = 7          ' LIBERO ARCHIVOS
    NOCLOSE% = 0
    'Call ACTULISTAOP ' LA MISMA SE REALIZA EN FRM OPORPA07 EN EL COMMAND23 DE VER Y REIMPRIMIR
    
    Exit Sub
    '
' *****************************************************************************
VALIDACUENRET:
    ' valida I.B.
    If XVALO(Label14) >= 0.005 Then
      Tipo% = XVALO(Label24)
      If Tipo% < 1 Or CUEMEPAG%(Tipo%) < 1 Then
        Call MENSERR(24, "Imposible Registrar esta O.P.\  \Cuenta Contable Incorrecta\para Categoría I.B. de este Proveedor.")
        OKX% = (-1)
        Exit Sub
      End If
    End If
    '
    ' valida I.B. Misiones
    If XVALO(Label45) >= 0.005 Then
      Tipo% = MedioPagoIBMisiones%
      If Tipo% < 1 Or CUEMEPAG%(Tipo%) < 1 Then
        Call MENSERR(24, "Imposible Registrar esta O.P.\  \Cuenta Contable Incorrecta\para Categoría I.B. de Misiones de este Proveedor.")
        OKX% = (-1)
        Exit Sub
      End If
    End If
    
    ' valida Ret.Ganancias
    If XVALO(Label30) >= 0.005 Then
      Tipo% = XVALO(Label32)
      If Tipo% < 1 Or CUEMEPAG%(Tipo%) < 1 Then
        Call MENSERR(24, "Imposible Registrar esta O.P.\  \Cuenta Contable Incorrecta\para Categoría Ganancias de este Proveedor.")
        OKX% = (-1)
        Exit Sub
      End If
    End If
    '
Return
'
' *****************************************************************************
ENTRAOBSERVA:
    Call CIERRARCH("!OBSERVOF")
    Call BLANARCH("!OBSERVOF")
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "GUIDI") > 0 Then
       ' TRASLADA LAS REEFERENCIAS A LAS OBSERVACIONES
       Call APERBASDAT("CABAN")
       'JANDY SQL
       Dim FACTSELE As ADODB.Recordset
       'SQLX$ = "SELECT * FROM CAJABANC "
       'SQLX$ = SQLX$ + "where NuProve = " & NC% & " "
       'SQLX$ = SQLX$ + "AND TipoMovim = 'FCOM' "
       'Set FACTSELE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
       
       '
       XX1& = 0
       For U& = 1 To List2.ListCount
         Z$ = List2.LIST(U& - 1)
         DOCU$ = Left$(Z$, 20)
         If DOCU$ <> "Anticipo no Aplicado" Then
           CANCE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
           If Abs(CANCE#) >= 0.005 Then
              DOCUCA$ = TRIM$(Left$(Z$, 18))
              ABUSQUE$ = "CODICOM_LAR = '" & DOCUCA$ & "' "
              SQLX$ = "SELECT * FROM CAJABANC  WITH(NOLOCK) "
              SQLX$ = SQLX$ + " where NuProve = " & NC% & " "
              SQLX$ = SQLX$ + " And TipoMovim = 'FCOM' "
              SQLX$ = SQLX$ + " And CODICOM_LAR = '" & DOCUCA$ & "'"
              'Set FACTSELE = FLEXCONN.Execute(FILTSQL$(SQLX$))
              'MODIFICADO PARA QUE SE CAPTURE EL T.ESPERA DENTRO DE LA FUNCION READSET
              Set FACTSELE = READSET(SQLX$)
              With FACTSELE
              '  .FindFirst ABUSQUE$
              '  If .NoMatch = False Then
              If Not (.EOF And .BOF) Then
                  obsex$ = TRIM$(!OBSERGEN)
                  If obsex$ <> "" Then
                     If obsex$ <> TRIM$(rasocli$((-1) * NC%)) Then
                       XX1& = XX1& + 1
                       Call GRAREG("!OBSERVOF", obsex$, XX1&)
                     End If
                  End If
              End If
              End With
              FACTSELE.Close
              Set FACTSELE = Nothing
           End If
         End If
       Next U&
       Call CIERRARCH("!OBSERVOF")
    End If
    '
    VALIDA% = VENTABU%("OBSERVOF/TITUPAN=Orden de Pagos Numero " + TRIM$(Text11))
    If VALIDA% < 0 Then
       OKX% = (-1): Exit Sub
    End If
Return
'
' *****************************************************************************
VERINUMERO:
    ' ESTO ES PARA VERIFICAR QUE EL ARCHIVO ESTÉ BLOQUEADO
    'JANDY SQL se comento
    'Call BLOQARCH("CUECORR")
    'CajaBanc.Close
    'If ISLOCKED%("CUECORR") < 1 Then
    '   Call LIBARCH("CUECORR")
    '   Call MENSERR(24, "No fue Posible Abrir la Base en Modo Exclusivo.\  \Espere un Instante y Vuelva a Grabar.")
    '   OKX% = 0
    '   Exit Sub
    'End If
    '
    
    If VERINUMOP%(NRO&) < 0 Then
       Call LIBARCH("CUECORR")
       NROPRO& = 0
       Call NUORPA(NROPRO&)
       Call MENSERR(24, "Al Grabar esta Orden de Pago\el Número Propuesto Estaba Ocupado.\  \Verifique el Nuevo Número Propuesto\y Vuelva a Grabar e Imprimir.")
       Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
       OKX% = 0
       Exit Sub
    End If
    '
Return
'
' *****************************************************************************
VERICHECAR:
    '
    For U& = 1 To List1.ListCount
       Z$ = List1.LIST(U& - 1)
       BANCO$ = Mid$(Z$, 7, 18)
       NUCHE$ = Mid$(Z$, 25, 14)
       IMCHE# = XVALO(Mid$(Z$, 70, 14))
       FECHE% = FECHANUM(Mid$(Z$, 39, 8)): If FECHE% < 1 Then FECHE% = HO%
       FEVEN% = FECHANUM(Mid$(Z$, 49, 8))
         FEX = FEVEN%: FEVECHE$ = FECHATEX$(FEX)
       IDEVA$ = AJUSTI$(BANCO$, 18) + AJUSTI$(NUCHE$, 14) + FECHATEX$(FEX)
       '
       HOII% = HOY(HOS$)
       If IMCHE# > 0 Then
          REGBA& = XVALO(TRIM$(Mid$(Z$, 121, 8)))
          If REGBA& > 0 Then ' es un cheque de cartera
            '
            VUELTA% = 0: ENCUEN% = 0
            IMCHEX$ = TRIM$(FORMATNUM$(IMCHE#, "F14.2"))
             'jandy sql
113         SQLX$ = "SELECT * FROM CAJABANC  WITH(NOLOCK) "
            SQLX$ = SQLX$ + " WHERE TipoMovim = 'CHRECI'"
            SQLX$ = SQLX$ + " AND Status = 1 "
            SQLX$ = SQLX$ + " AND OPCIMOVIM = 'CH' "
            SQLX$ = SQLX$ + " AND DescriMov = '" & IDEVA$ & "' "
            
'113        'SQLX$ = "DescriMov = '" & IDEVA$ & "' "
            SQLX$ = SQLX$ + " AND ABS(ValAbsComp - " & IMCHEX$ & ") < 0.005 "
            SQLX$ = SQLX$ + " AND (NumeComp = " & REGBA& & " OR NumeComp = 0) "
            SQLX$ = SQLX$ + " AND Status = 1"
            '
            'Set CHETEMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
            'MODIFICADO PARA QUE SE CAPTURE EL T.ESPERA DENTRO DE LA FUNCION READSET
            Set CHETEMP1 = READSET(SQLX$)
            With CHETEMP1
            '  .FindFirst SQLX$
            '  If .NoMatch = False Then
            '    On Error GoTo 0
              If Not (.EOF And .BOF) Then
                ENCUEN% = 1
                GoTo 119
              End If
              '
              If ENCUEN% < 1 Then
                If VUELTA% < 1 Then
                  IDEVA$ = AJUSTI$(BANCO$, 20) + AJUSTI$(NUCHE$, 12) + FECHATEX$(FEX)
                  VUELTA% = 1
                  GoTo 113
                End If
              End If
              '
            End With
            CHETEMP1.Close
            Set CHETEMP1 = Nothing
            '
            If ENCUEN% < 1 Then
               Call LIBARCH("CUECORR")
               Call MENSERR(24, "El Cheque '" + TRIM$(IDEVA$) + "'.\Aparece Como Salido de Cartera.\   \Corrija y Vuelva a Grabar O/Pago.")
               OKX% = 0
               Exit Sub
            End If
          End If
          '
       End If
119 Next U&
    '
Return
'
' *****************************************************************************
VERIFACPEN:
    '
    If TICUEPRO%(NC1%) = 0 Then      ' SOLO CUENTAS DE GASTOS
       '
       XX% = 0
       For U& = 1 To List2.ListCount
         Z$ = List2.LIST(U& - 1)
         DOCU$ = Left$(Z$, 20)
         If DOCU$ <> "Anticipo no Aplicado" Then
             CANCE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
             If CANCE# > 0 Then
                DOCUCA$ = TRIM$(Left$(Z$, 18))
                FEVEN% = FECHANUM(Mid$(Z$, 47, 8))
                SALPENX# = XVALO(Mid$(Z$, 56, 13))
                '
                ABUSQUE$ = "CODICOM_LAR = '" & DOCUCA$ & "' "
                  'JANDY SQL
                SQLX$ = "SELECT * FROM SACREPEN "
                SQLX$ = SQLX$ + "WHERE NuProve = " & NC%
                SQLX$ = SQLX$ + " AND  CODICOM_LAR = '" & DOCUCA$ & "'"
                Set FacPenTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
                
                With FacPenTemp
                '  .FindFirst ABUSQUE$
'643               If .NoMatch = False Then
                   If Not (.EOF And .BOF) Then
                     Do While Not .EOF
                     SADEUBAS# = !SaldAcre - !SaldDebe
                     If MONEMI% > 1 Then
                        SADEUBAS# = !SaldAcre_DOLA - !SaldDebe_DOLA
                     End If
                      If Abs(SADEUBAS# - SALPENX#) < 0.005 Then
                        If Abs(FEVEN% - CVINT(!FEVENCIM)) < 0.5 Then
                           ' ENCONTRO EL REGISTRO
                           GoTo 649
                        End If
                      End If
                     '.FindNext ABUSQUE$
                     'GoTo 643
                     .MoveNext
                     Loop
                  End If
                End With
                FacPenTemp.Close
                Set FacPenTemp = Nothing
                Call MENSERR(24, "Imposible Registrar Esta O/Pago.\   \Diferencia en el Saldo Pendiente\del Documento '" + DOCUCA$ + "'.\Verifique y Vuelva a Grabar.")
                OKX% = 0
                Exit Sub
                '
             End If
         End If
649    Next U&
       '
    End If
    '
Return
'
' *****************************************************************************
ARMANTICIPO:
    FVII%(1) = FC%
    IMV#(1) = 0
    '\\\OT-06-0570-RG-27/12/06///
    List2.ListIndex = -1
    '\\\OT-06-0570-RG-FIN///
    X$ = List2.TEXT
    DOCU$ = Mid$(X$, 1, 20)
    If DOCU$ = "Anticipo no Aplicado" Then
      CANCE# = CDbl(XVALO(Mid$(X$, 70, 14)))
      If Abs(CANCE#) >= 0.005 Then
        IMV#(1) = (-1) * CANCE#
      End If
    End If
Return
'
' *****************************************************************************
ARMANUMEROS:
    COBRACT% = 0
    NRECIB$ = "0" ' TALONARIO
    While Len(NRECIB$) < 2: NRECIB$ = "0" + NRECIB$: Wend
    NRECIB$ = TRIM$(Str$(NRO&)) + "-" + NRECIB$
    While Len(NRECIB$) < 9: NRECIB$ = "0" + NRECIB$: Wend
    NRECIB$ = "OP-" + NRECIB$
    '
    RECILAR$ = TRIM$(Str$(NRO&))
    While Len(RECILAR$) < 8: RECILAR$ = "0" + RECILAR$: Wend
    RECILAR$ = TRIM$(Str$(COBRACT%)) + "-" + RECILAR$
    While Len(RECILAR$) < 13: RECILAR = "0" + RECILAR$: Wend
    RECILAR$ = "OP-X-" + RECILAR$
Return
'
' *****************************************************************************
GRABATORPAG:                              ' grabar orden de pago en CAJABANC
    '
    CoComCor$ = NRECIB$
    CoComLar$ = RECILAR$
'    CodiEmp% = 0
'    CodiSuc% = SUC%
    NumeClie = 0
    NumeProv% = NC%
    FechEmi = CVDAT(FC%)
    Referen$ = rasocli$(NC1%)
    FeContab = CVDAT(FC%)
    FeAcredi = CVDAT(FC%)
    CuentaCo% = CUECOCLI%(NC1%)
    TiMovi$ = "OPAG"
    OpMovi$ = "OP"
    VaMovi$ = "CCORR"
    Descrip$ = "Orden de Pago " + NRECIB$
    NumeTal = COBRACT%
    NumeCom = NRO&
    StatInit = 0
    ImpoDeb# = Abs(TOFAC#(0))
    BasImpo# = XVALO(Label23)    ' aqui va entonces el neto sin IVA
    Observa$ = ""
    If ULTREG&("!OBSERVOF") > 0 Then
      Observa$ = TRIM$(REGLEIDO$("!OBSERVOF", 1))
    End If
    If Observa$ = "" Then Observa$ = rasocli$(NC1%)
    '
    'agregado multimoneda
    If MONEMI% > 1 Then
       ValoDola# = ImpoDeb#
       ImpoDeb# = ImpoDeb# * TICAMONE(MONEMI%)    ' PARA PASARLO A PESOS
    End If
    
    If TICUEPRO%(NC1%) = 0 Then
         Call GRACABAN
       Else
         For K1& = 1 To ULTREG&("!CARGAST")
           Z$ = REGLEIDO$("!CARGAST", K1&)
           CuentaCo% = CUEINT%(Left$(Z$, 12))
           ImpoDeb# = CVD(Mid$(Z$, 13, 8))
           If MONEMI% > 1 Then
              ImpoDeb# = ImpoDeb# * TICAMONE(MONEMI%)    ' PARA PASARLO A PESOS
           End If
           VaMovi$ = "CGAST"
           Referen$ = DENOCUE$(CuentaCo%)
           CentroCos$ = Mid$(Z$, 25, 8)
           'SI SE CARGA CENTRO DE COSTO EN EL DETALLE DE LA FACTURA MANDA ESTE
           'SI NO SE CARGA TRAE EL CONFIGURADO EN LA TABLA DE ASOCIACION DE C.CBLES A CENTRO DE COSTOS ("CTAYCOST")
           'TEXTO CTAYCOST
           If TRIM$(CentroCos$) = "" Then
               CentroCos$ = CENTROCOSTOASOCIADO$(CuentaCo%)
           End If
           ItConcep$ = Mid$(Z$, 33, 8) 'P/ conceptos
           Call GRACABAN
         Next K1&
         CentroCos$ = ""
         ItConcep$ = ""
         ' para que no repita sobre los cheques
    End If
    OKX% = 2         ' grabó Orden de Pago
Return
'
' *****************************************************************************
GRABACHENTRE:        ' grabar los valores entregados
    '
    For U& = 1 To List1.ListCount
       Z$ = List1.LIST(U& - 1)
       TIDOS$ = Left$(Z$, 4)  ' Prop
       Tipo% = XVALO(Mid$(Z$, 115, 6))
       CUECONTABINT% = CUEMEPAG%(Tipo%)
       BANCO$ = Mid$(Z$, 7, 18)
       NUCHE$ = Mid$(Z$, 25, 14)
       IMCHE# = XVALO(Mid$(Z$, 70, 14))
       If MONEMI% > 1 Then
          IMCHE# = IMCHE# * TICAMBX#
       End If
       FECHE% = FECHANUM(Mid$(Z$, 39, 8)): If FECHE% < 1 Then FECHE% = HO%
       FEVEN% = FECHANUM(Mid$(Z$, 49, 8))
         FEX = FEVEN%: FEVECHE$ = FECHATEX$(FEX)
       IDEVA$ = AJUSTI$(BANCO$, 18) + AJUSTI$(NUCHE$, 14) + FECHATEX$(FEX)
       '
       HOII% = HOY(HOS$)
       If IMCHE# > 0 Then
          REGBA& = XVALO(TRIM$(Mid$(Z$, 121, 8)))
          If REGBA& > 0 Then ' es un cheque de cartera
            '
            VUELTA% = 0: ENCUEN% = 0
            IMCHEX$ = TRIM$(FORMATNUM$(IMCHE#, "F14.2"))

13          SQLX$ = "SELECT * FROM CAJABANC "
            SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
            SQLX$ = SQLX$ + "AND Status = 1 "
            SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
            SQLX$ = SQLX$ + "AND DescriMov = '" & IDEVA$ & "' "
'13          SQLX$ = "DescriMov = '" & IDEVA$ & "' "
            SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IMCHEX$ & ") < 0.005 "
            SQLX$ = SQLX$ + "AND (NumeComp = " & REGBA& & " OR NumeComp = 0) "
            SQLX$ = SQLX$ + "AND Status = 1"
            Set CHETEMP1 = New ADODB.Recordset
14         On Error Resume Next
            CHETEMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 14
            End If
            On Error GoTo 0
            '
            With CHETEMP1
            If Not (.EOF And .BOF) Then
'15            .FindFirst SQLX$
'              If .NoMatch = False Then
'                On Error GoTo 0
'                .Edit
                !FEDEPOSENT = CVDAT(FC%)
                !Document = NRECIB$
                !DESTICHEQ = rasocli$(NC1%)
                 CUECONTABINT% = !CUECONTAI  ' PARA QUE LA SALIDA SEA DE LA MISMA CUENTA QUE ENTRO
                !REFESALCHE = "Orden de Pago Nro." + Str$(NRO&)
                !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
                !USUASALCHE = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
                !Status = 2                     ' Marca Salida de Cartera
                !IdenBaja = IDACTU$             ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
                .Update
                ENCUEN% = 1
            End If
              '
              If ENCUEN% < 1 Then
                If VUELTA% < 1 Then
                  IDEVA$ = AJUSTI$(BANCO$, 20) + AJUSTI$(NUCHE$, 12) + FECHATEX$(FEX)
                  VUELTA% = 1
                  GoTo 13
                End If
              End If
              '
            End With
            CHETEMP1.Close
            Set CHETEMP1 = Nothing
            '
            If ENCUEN% < 1 Then
              Call MENSERR(24, "Imposible Registrar la Salida\del Cheque '" + TRIM$(IDEVA$) + "'.\Consulte a su Soporte de Sistemas")
            End If
          End If
          '
          CoComCor$ = NRECIB$
          CoComLar$ = RECILAR$
'          CodiEmp% = 0
 '         CodiSuc% = SUC%
          NumeClie = 0
          NumeProv% = NC%
          FechEmi = CVDAT(FECHE%)
            FEX = FC%
          Referen$ = rasocli$(NC1%) + " - " + FECHATEX$(FEX)
          FeContab = CVDAT(FC%)
          '
          FeAcredi = CVDAT(FEVEN%)
          CuentaCo% = CUECONTABINT%
          TiMovi$ = "CHENTR"
          OpMovi$ = "OP"
          VaMovi$ = ""
          StatInit = 0
          Descrip$ = IDEVA$
          NumeTal = COBRACT%
          NumeCom = NRO&
          ImpoDeb# = (-1) * Abs(IMCHE#)
          BasImpo# = 0               ' CERO PARA LA SALIDA DEL CHEQUE
          Observa$ = "O.Pago " + TRIM$(Str$(NRO&)) + " - " + rasocli$(NC1%)
          Z$ = TTS$
          Call REPLA(Z$, Mid$(REBLA$, 209), 209, 304)
          '
          Call GRACABAN
       End If
    Next U&
    '
Return

End Sub

'
Sub PRIORPASTA()
    '
    HOI% = HOY(HO$)
    NC% = XVALO(Text1)
    '
    FORIPRE$ = TRIM$(Control$("", "FORPASTA"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Orden de Pago.")
        Exit Sub
    End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Recibo de Cobranza:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    CODFOR$ = FORIPRE$
    '
    CONTA% = List1.ListCount
    If CONTA% < 1 Then
        Call COMUNI("No se ha Ingresado Detalle de Valores")
        Exit Sub
    End If
    '
    If XVALO(Label14.Caption) + XVALO(Label45.Caption) + XVALO(Label12.Caption) + XVALO(Label21.Caption) + XVALO(Label30.Caption) + XVALO(Label38.Caption) + XVALO(Label49.Caption) < 0.005 Then
        Call COMUNI("No se ha Ingresado Detalle de Valores")
        Exit Sub
    End If
    '
    AVAHOJA$ = TRIM$(Control$("ORPASTA", "AVAHOJA"))
    Screen.MousePointer = 11
    Call BLANARCH("!PRINRECI")
    '
    ILISTA& = List1.ListIndex
    TOLISTA& = List1.TopIndex
    '
    For U& = 1 To CONTA%
      'List1.ListIndex = U& - 1
      X$ = List1.LIST(U& - 1)
      If XVALO(Mid$(X$, 70, 14)) >= 0.005 Then
        If TRIM$(X$) = "" Then GoTo 19
        TIPOS$ = Left$(X$, 4)
        BANCO$ = Mid$(X$, 7, 18)
        NUVA$ = Mid$(X$, 25, 14)
        IMVA# = XVALO(Mid$(X$, 70, 14))
        CUECONTT% = XVALO(Mid$(X$, 115, 6))
        If IMVA# < 0.005 Then GoTo 19
        '
        FEMII% = FECHANUM(Mid$(X$, 39, 8)): If FEMII% < 1 Then FEMII% = HOI%
        FEVEI% = FECHANUM(Mid$(X$, 49, 8)): If FEVEI% < 1 Then FEMII% = HOI%
        PROTER$ = ""
        CLEA% = 48
        '
10      Z$ = REGBLAN$("!PRINRECI")
        Call REPLA(Z$, TIPOS$, 1, 4)
        Call REPLA(Z$, NUVA$, 5, 12)
        Call REPLA(Z$, BANCO$, 17, 16)
        Call REPLA(Z$, MKD$(IMVA#), 33, 8)
        FEX = FEVEI%
        If FEX > 0 Then
            Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
        End If
        Call REPLA(Z$, MKI$(CUECONTT%), 129, 2)
        Call REGAPP("!PRINRECI", Z$)
      End If
      '
19  Next U&
    '
    If XVALO(Label14) >= 0.005 Then
       Z$ = REGBLAN$("!PRINRECI")
       Call REPLA(Z$, "R.IB", 1, 4)
       Call REPLA(Z$, TRIM$(Label25), 5, 12)
       'Call REPLA(Z$, "Ret.IB.Bs.As.", 17, 16)
       Call REPLA(Z$, "Ret.I.Brutos.", 17, 16)
       Call REPLA(Z$, MKD$(XVALO(Label14)), 33, 8)
       FEX = HOY(HOS$)
       Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
       Call REGAPP("!PRINRECI", Z$)
    End If
    '
    If XVALO(Label38) >= 0.005 Then
       Z$ = REGBLAN$("!PRINRECI")
       Call REPLA(Z$, "R.IB", 1, 4)
       Call REPLA(Z$, TRIM$(Label41), 5, 12)
       Call REPLA(Z$, "Ret.IB CABA", 17, 16)
       Call REPLA(Z$, MKD$(XVALO(Label38)), 33, 8)
       FEX = HOY(HOS$)
       Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
       Call REGAPP("!PRINRECI", Z$)
    End If
    '
    If XVALO(Label45) >= 0.005 Then
       Z$ = REGBLAN$("!PRINRECI")
       Call REPLA(Z$, "R.IB", 1, 4)
       Call REPLA(Z$, CStr(NumeroRetencionIBMisiones&), 5, 12)
       Call REPLA(Z$, "Ret.IB Misiones", 17, 16)
       Call REPLA(Z$, MKD$(XVALO(Label45)), 33, 8)
       FEX = HOY(HOS$)
       Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
       Call REGAPP("!PRINRECI", Z$)
    End If
    '
    If XVALO(Label12) >= 0.005 Then
       Z$ = REGBLAN$("!PRINRECI")
       Call REPLA(Z$, "R.IB", 1, 4)
       Call REPLA(Z$, CStr(NumeroRetencionIBTucuman&), 5, 12)
       Call REPLA(Z$, "Ret.IB Tucuman", 17, 16)
       Call REPLA(Z$, MKD$(XVALO(Label12)), 33, 8)
       FEX = HOY(HOS$)
       Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
       Call REGAPP("!PRINRECI", Z$)
    End If
    '
    If XVALO(Label49) >= 0.005 Then
       Z$ = REGBLAN$("!PRINRECI")
       Call REPLA(Z$, "RIVA", 1, 4)
       Call REPLA(Z$, CStr(NumeroRetencionIva&), 5, 12)
       Call REPLA(Z$, "Ret.IVA", 17, 16)
       Call REPLA(Z$, MKD$(XVALO(Label49)), 33, 8)
       FEX = HOY(HOS$)
       Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
       Call REGAPP("!PRINRECI", Z$)
    End If
    '
    If XVALO(Label30) >= 0.005 Then
       Z$ = REGBLAN$("!PRINRECI")
       Call REPLA(Z$, "RGAN", 1, 4)
       Call REPLA(Z$, TRIM$(Label34), 5, 12)
       Call REPLA(Z$, "Ret.Ganancias", 17, 16)
       Call REPLA(Z$, MKD$(XVALO(Label30)), 33, 8)
       FEX = HOY(HOS$)
       Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
       Call REGAPP("!PRINRECI", Z$)
    End If
    '
    V0& = ULTREG&("!PRINRECI")
    '
    ILISTA& = List2.ListIndex
    TOLISTA& = List2.TopIndex
    '
    V& = 0
    If TICUEPRO%((-1) * NC%) = 1 Then
       For U& = 1 To ULTREG&("!CARGAST")
          X$ = REGLEIDO$("!CARGAST", U&)
          DOCU$ = Left$(X$, 12)
          SALPE# = 0
          CANCE# = CVD(Mid$(X$, 13, 8))
          If Abs(CANCE#) >= 0.005 Then
             V& = V& + 1
             If V& <= V0& Then
                  Z$ = REGLEIDO$("!PRINRECI", V&)
                Else
                  Z$ = REGBLAN$("!PRINRECI")
             End If
             Call REPLA(Z$, DOCU$, 53, 18)
             Call REPLA(Z$, MKD$(CANCE#), 71, 8)
             Call REPLA(Z$, MKD$(SALPE#), 79, 8)
             Call GRAREG("!PRINRECI", Z$, V&)
          End If
       Next U&
       Call CIERRARCH("!CARGAST")
       GoTo 25
    End If
    '
    For U& = 1 To List2.ListCount
        'List2.ListIndex = U& - 1
        X$ = List2.LIST(U& - 1)
        DOCU$ = Mid$(X$, 1, 20)
        SALPE# = CDbl(XVALO(Mid$(X$, 56, 13)))
        CANCE# = CDbl(XVALO(Mid$(X$, 70, 14)))
        If Abs(CANCE#) >= 0.005 Then
            V& = V& + 1
            If V& <= V0& Then
                  Z$ = REGLEIDO$("!PRINRECI", V&)
               Else
                  Z$ = REGBLAN$("!PRINRECI")
            End If
            Call REPLA(Z$, DOCU$, 53, 18)
            Call REPLA(Z$, MKD$(CANCE#), 71, 8)
            Call REPLA(Z$, MKD$(SALPE#), 79, 8)
            Call GRAREG("!PRINRECI", Z$, V&)
        End If
    Next U&
    '
25  V0& = ULTREG&("!PRINRECI")
    '
    TOTAVAL# = CDbl(XVALO(Label14.Caption) + XVALO(Label21.Caption) + XVALO(Label30.Caption)) + XVALO(Label38.Caption) + XVALO(Label45.Caption) + XVALO(Label12.Caption) + XVALO(Label49.Caption)
    TOTACAN# = CDbl(XVALO(Label18.Caption)) + CDbl(XVALO(Label20.Caption))
33  ITEXTO$ = NUMTEX$(TOTAVAL#) + " ctvs"
    LOTEX% = XVALO(ATRIFORM$(CODFOR$, "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = "": IMPLET2$ = "": IMPLET3$ = ""
    If TRIM$(NRX1$) <> "" Then
       IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    End If
    If TRIM$(NRX2$) <> "" Then
       IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
    End If
    If TRIM$(NRX3$) <> "" Then
       IMPLET3$ = TRIM$(NRX3$) + " " + "***************************************************************************************************************"
    End If
    '
    Call CIERRARCH("!PRINRECI")
    '
30  If TRIM$(FORIPRE$) <> "" Then
        '
        CANCELPRINT% = 0
        If Control$("ORPASTA", "SUPRINT") = "SI" Then CANCELPRINT% = 1
        '
        CODFOR$ = TRIM$(UCase$(FORIPRE$))
        MARIZW = XVALO(ATRIFORM$(CODFOR$, "MARIZQMM"))
        MARSUW = XVALO(ATRIFORM$(CODFOR$, "MARSUPMM"))
        '
        INTERLI = XVALO(ATRIFORM$(CODFOR$, "INTERLIN"))
        If INTERLI < 1 Then
            Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
            INTERLI = 1
        End If
        '
        DESMM = XVALO(ATRIFORM$(CODFOR$, "DETDESMM"))
        HASMM = XVALO(ATRIFORM$(CODFOR$, "DETHASMM"))
        '
        CACOPI% = XVALO(ATRIFORM$(CODFOR$, "CAN-COP"))
        If CACOPI% < 1 Or CACOPI% > 6 Then
            Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
            CACOPI% = 1
        End If
        '
        ORIENTA% = 1
        PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
        If PORI$ = "PAISAJE" Or PORI$ = "2" Then
            ORIENTA% = 2
        End If
        '
        PORIAN = Printer.Orientation  ' retiene orientacion
        '
        NRORPA$ = TRIM$(Mid$(Text11.TEXT, 4))
        '
        Screen.MousePointer = 11
        If AVAHOJA$ <> "NEWPAGE" Then
            'Printer.Orientation = 1
            Call SETSPOOL("ORIENTATION", "1")
            If ORIENTA% = 2 Then
                'Printer.Orientation = 2
                Call SETSPOOL("ORIENTATION", "2")
            End If
        End If
        '
        NC1% = (-1) * NC%
        For COP% = 1 To CACOPI%
            '
            If AVAHOJA$ = "NEWPAGE" Then
               If PRINTERCONTROL% = 1 Then
                  'Printer.NewPage
                  Call SETSPOOL("NEWPAGE", "")
               End If
            End If
            PRINTERCONTROL% = 1
            '
            SUMAVA# = 0: SUMACAN# = 0
            Call CARFORWIN
            '
            Call PRINTFORWIN(TRIM$(Str$(NC%)), "CUEN-CLI", 0, 0)
            Call PRINTFORWIN(TRIM$(rasocli$(NC1%)), "RASO-CLI", 0, 0)
            Call PRINTFORWIN(DOMICLI$(NC1%), "DOMI-CLI", 0, 0)
            Call PRINTFORWIN(CPOSCLI$(NC1%), "CPOS-CLI", 0, 0)
            Call PRINTFORWIN(LOCACLI$(NC1%), "LOCA-CLI", 0, 0)
            Call PRINTFORWIN(CuitCli$(NC1%), "CUIT-CLI", 0, 0)
            Call PRINTFORWIN(POSIVAL$(NC1%), "PIVA-CLI", 0, 0)
            Call PRINTFORWIN(TEL1CLI$(NC1%), "TELE-CLI", 0, 0)
            FF% = FECHANUM(TRIM$(Text9.TEXT))
            Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
            Call PRINTFORWIN(NRORPA$, "NUME-COM", 0, 0)
            Call PRINTFORWIN(DEMONECO$(MONEMI%), "MONEDEMI", 0, 0)
            Call PRINTFORWIN(TICAMBIS, "TICAMBIO", 0, 0)
            '
            YAC = 0
            For U& = 1 To ULTREG&("!OBSERVOF")
               TXOB$ = TRIM$(REGLEIDO$("!OBSERVOF", U&))
               Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
               YAC = YAC + INTERLI
            Next U&
            Call CIERRARCH("!OBSERVOF")
            '
            YAC = DESMM
            For YX& = 1 To ULTREG&("!PRINRECI")
               '
               YAC = YAC + INTERLI
               '
               If YAC > HASMM - 2 * INTERLI Then
                   YAC = YAC + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   '
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-VAL", 0, 0)
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-CAN", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET1", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET2", 0, 0)
                   '
                   If AVAHOJA$ = "NEWPAGE" Then
                        'Printer.NewPage
                        Call SETSPOOL("NEWPAGE", "")
                      Else
                       'Printer.EndDoc
                       Call SETSPOOL("ENDDOC", "")
                   End If
                   '
                   Call CARFORWIN
                   Call PRINTFORWIN(TRIM$(Str$(NC1%)), "CUEN-CLI", 0, 0)
                   Call PRINTFORWIN(TRIM$(rasocli$(NC1%)), "RASO-CLI", 0, 0)
                   Call PRINTFORWIN(DOMICLI$(NC1%), "DOMI-CLI", 0, 0)
                   Call PRINTFORWIN(CPOSCLI$(NC1%), "CPOS-CLI", 0, 0)
                   Call PRINTFORWIN(LOCACLI$(NC1%), "LOCA-CLI", 0, 0)
                   Call PRINTFORWIN(CuitCli$(NC1%), "CUIT-CLI", 0, 0)
                   Call PRINTFORWIN(POSIVAL$(NC1%), "PIVA-CLI", 0, 0)
                   Call PRINTFORWIN(TEL1CLI$(NC1%), "TELE-CLI", 0, 0)
                   FEX = FF%
                   Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
                   Call PRINTFORWIN(NRORPA$, "NUME-COM", 0, 0)
                   '
                   YAC = DESMM + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   '
                   YAC = YAC + 2 * INTERLI
                   '
                End If
                '
                YY$ = REGLEIDO$("!PRINRECI", YX&)
                Call PRINTFORWIN(Left$(YY$, 4), "TIPO-VAL", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 5, 12), "NUME-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 17, 16), "BANC-CHE", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 33, 8), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                SUMAVA# = SUMAVA# + CVD(Mid$(YY$, 33, 8))
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 41, 8), 6, 8, 0, 1), "FECH-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 53, 18), "DOCU-APL", 0, YAC)
                '
                'Obtiene CUENTA MADRE
                If CVD(Mid$(YY$, 71, 8)) >= 0.005 Then
                    CTAMADRE$ = TRIM(CUEMADRE((-1) * NC%))
                    DET$ = CTAMADRE$ & " - " & TRIM(DENOCUE$(CUEINT%(CTAMADRE$)))
                    Call PRINTFORWIN(DET$, "DET-PASE", 0, YAC)
                End If
                '
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 71, 8), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                SUMACAN# = SUMACAN# + CVD(Mid$(YY$, 71, 8))
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 79, 8), 4, 12, 2, 2), "IMPO-PEN", 0, YAC)
                '
                'Obtiene Descripcion de Cuenta Contable a partir del !CARMEPA
'                XCP$ = REGLEIDO("!CARMEPA", YX&)
'                NCARMEPA% = CVI(Mid(XCP$, 1, 2))
'                CCC% = CUEMEPAG%(NCARMEPA%) 'CODIGO DE CUENTA CONTABLE
'                CUEDENOM$ = TRIM(CODCUE$(CCC%)) & " - " & TRIM$(DENOCUE$(CCC%)) 'CUENTA Y DENOMINACION DE CUENTA CONTABLE
'                Call PRINTFORWIN(CUEDENOM$, "DEN-CUEN", 0, YAC)
                '
                'Obtiene Descripcion de Cuenta Contable
                CCC% = CUEMEPAG%(CVI(Mid$(YY$, 129, 2))) 'CODIGO DE CUENTA CONTABLE
                CUEDENOM$ = TRIM(CODCUE$(CCC%)) & " - " & TRIM$(DENOCUE$(CCC%)) 'CUENTA Y DENOMINACION DE CUENTA CONTABLE
                Call PRINTFORWIN(CUEDENOM$, "DEN-CUEN", 0, YAC)
                '
                'Obtiene Fecha de Emision de la Lista
                FLAN$ = Mid$(List1.LIST(YX& - 1), 39, 8)
                Call PRINTFORWIN(FLAN$, "FECH-LAN", 0, YAC)
                '
                'Obtiene Cuit del Librador
                If UCase(Mid$(List1.LIST(YX& - 1), 1, 4)) = "CH.3" Then
                   BANCONUMX$ = Mid$(List1.LIST(YX& - 1), 7, 32)
                   FechVenc$ = Mid$(List1.LIST(YX& - 1), 49, 8)
                   IMPOX# = XVALO(Mid$(List1.LIST(YX& - 1), 59, 12))
                   IMPOS$ = TRIM$(FORMATNUM(IMPOX#, "F16.6"))
                   
                   CONDIX$ = "TipoMovim ='CHRECI' AND OpciMovim='CH'"
                   CONDIX$ = CONDIX$ + " AND Descrimov = '" & BANCONUMX$ & FechVenc$ & "'"
                   CONDIX$ = CONDIX$ + " AND ABS(IDebeComp - " & IMPOS$ & ")< 0.005"
                   
                   CUITLIBX$ = VALOBADA("CAJABANC", "ObserGen", CONDIX$)
                   CUITLIB$ = Mid$(CUITLIBX$, 12, 11)
                   Call PRINTFORWIN(CUITLIB$, "CUIT-TRA", 0, YAC)
                End If
'                'Obtiene CUENTA MADRE
'                CTAMADRE$ = TRIM(CUEMADRE((-1) * NC%))
'                DET$ = CTAMADRE$ & " - " & TRIM(DENOCUE$(CUEINT%(CTAMADRE$)))
'                Call PRINTFORWIN(DET$, "DET-PASE", 0, YAC)
            Next YX&
            '
            Call PRINTFORWIN(CSTRING$(MKD$(TOTAVAL#), 4, 12, 2, 2), "TOTA-VAL", 0, 0)
            Call PRINTFORWIN(CSTRING$(MKD$(TOTACAN#), 4, 12, 2, 2), "TOTA-CAN", 0, 0)
            Call PRINTFORWIN(IMPLET1$, "IMP-LET1", 0, 0)
            Call PRINTFORWIN(IMPLET2$, "IMP-LET2", 0, 0)
            Call PRINTFORWIN(USERNAME, "USUARIO", 0, 0)         '
            '
            If AVAHOJA$ <> "NEWPAGE" Then
                 'Printer.EndDoc
                 Call SETSPOOL("ENDDOC", "")
            End If
            '
        Next COP%
        '
        Call CIERRARCH("*.*")
        DESCRIDOC$ = "Orden de Pago Nro. " + TRIM$(NRORPA$) + " - " + TRIM$(rasocli$(NC1%))
        TIPODOC$ = "Orden de Pago"
        NUMEDOC$ = NRORPA$
        FEX = FF%
        DESCRIDOC$ = DESCRIDOC$ + " - " + FECHATEX$(FEX)
        '
        Call SAVESPOOL(DESCRIDOC$, OKX%)
        Screen.MousePointer = 1
        If COMALTER%("Orden de Pago Numero\'" + NRORPA$ + "'\Emitida y Registrada\\Repetir Impresion\Otra Orden de Pago") = 1 Then GoTo 30
        '
        If AVAHOJA$ <> "NEWPAGE" Then
            Printer.Orientation = PORIAN
        End If
        '
        CANCELPRINT% = 0
        '
    End If
    Call CIERRARCH("*.*")
    '
End Sub

Private Sub TICAMBIS_Change()
   '
   If TICAMBIS.Enabled = True Then
      TICACH(MONEMI%) = XVALO(TICAMBIS)
   End If
   TICAMBX# = 1
   If MONEMI% > 1 Then
     TICAMBX# = XVALO(TICAMBIS)
   End If
   '
   COTIMONEM# = XVALO(TICAMBIS)
   '
End Sub

Private Sub TICAMBIS_DblClick()
   'Call CARTAMONED
   '
   '\\\OT-07-0086-RG-06/03/07///
   If MONEMI% > 1 Then
     If DERACCE%("COTILOCA", "Ajustar Tipo de Cambio Local") >= 2 Then
        ENTRATICA.Frame4.Caption = "Cotización Local para " + MONEMIS.Caption
        ENTRATICA.Text4 = FORMATNUM(XVALO(TICAMBIS.Caption), "F10.4")
        ENTRATICA.Show 1
        If TRIM$(ENTRATICA.Text4) <> "" Then
          TICAMBIS = TRIM$(ENTRATICA.Text4)
        End If
     End If
     List1.Clear
     List2.Clear
     Label4.Caption = ""
     Label8.Caption = ""
     Label10.Caption = ""
     Label14.Caption = ""
     Label45.Caption = ""
     Label12.Caption = ""
     Label38.Caption = ""
     Label49.Caption = ""
     Label18.Caption = ""
     Label20.Caption = ""
     Text7 = ""
     BOTNEXT.Enabled = True
     '
     Call SIGUEORPA1
     '
   End If
   '\\\OT-07-0086-RG-FIN///
   '
End Sub


Private Sub veryreimp_Click()
    Call OPORPA07.Command23_Click
End Sub

Sub CALRETIBRU()
   '
'   Static CTCALCU%, ALIBRU(8), MINIMPO#(8), MEPAGX%(8)
   Static NPROVEA%, NPROVEB%, NOPAGAN$, MESANANT$, CATIB1%, PORCONVE
   Static ACUBASANT#, ACURETANT#
   '
   If XVALO(Label18) < 0.005 Then
     Label14 = "": Label23 = ""
     GoTo 99
   End If
   '
   If RETIBMAN = 1 Then GoTo 99
   '
'   Screen.MousePointer = 11
'   If CTCALCU% = 0 Then
'     CTCALCU% = (-1)
'     For U& = 1 To ULTREG&("TACATIBP")
'       TCX$ = REGLEIDO$("TACATIBP", U&)
'       TCY% = Asc(Mid$(TCX$, 1, 1))
'       If TCY% > 0 Then
'         If TCY% <= 8 Then
'           If CVS(Mid$(TCX$, 33, 4)) > 0 Then CTCALCU% = 1
'         End If
'       End If
'     Next U&
'     Call CIERRARCH("TACATIBP")
'   End If
'   '
'   If CTCALCU% < 0 Then
'     Label23 = FORMATNUM$(NETOCANCE#, "F12.2")
'     GoTo 99
'   End If
'   '
   ' determinar categoria de I.Brutos
   NPROVEI% = XVALO(Text1)
   If TICUEPRO%((-1) * NPROVEI%) = 1 Then GoTo 99 ' RETENCION CERO EN CUENTAS DE GASTOS
   '
   If NPROVEI% <> NPROVEB% Then
     REBLA$ = REGBLAN$("DARIBPRO")
     RERET$ = FILTERGETRECORD$("DARIBPRO", 1, MKI$(NPROVEI%))
     CATIB1% = 1: PORCONVE = 100   ' VALORES POR DEFAULT
     If RERET$ <> REBLA$ Then
       CATIB1% = Asc(Mid$(RERET$, 3, 1))
       PORCONVE = 100 ' CVS(Mid$(RERET$, 5, 4))
     End If
     Label24 = Str$(MEDIPAGO%)
     Label29 = FORMATNUM$(PORCEIBRU#, "F8.2")
     NPROVEB% = NPROVEI%
   End If
   If CATIB1% < 1 Then
     Label23 = FORMATNUM$(NETOCANCE#, "F12.2")
     GoTo 99 ' exento
   End If
   '
   ' determinar monto imponible
   '\\\OT-05-0319-WAR-24/05/05///
   '
   If NPROVEI% <> NPROVEA% Or Text11 <> NOPAGAN$ Or Mid$(Text9, 4) <> MESANANT$ Then
      ACUBASANT# = 0: ACURETANT# = 0
      NPROVEA% = NPROVEI%: NOPAGAN$ = Text11: MESANANT$ = Mid$(Text9, 4)
      '
      If Control("ORPASTA", "ACUMENS") = "SI" Then
         DESMEI% = FECHANUM("01" + Mid$(Text9, 3))
         DESDAT1 = FETEXCOR1$(DESMEI%)
         HASMEI = DIPOST(DESMEI%)
         While Mid$(FECHATEX$(DIPOST(HASMEI)), 4, 2) = Mid$(Text9, 4, 2)
           HASMEI = DIPOST(HASMEI)
         Wend
         HASMEII% = HASMEI
         HASDAT1 = FETEXCOR1$(HASMEII%)
         ' determinar monto imponible - ACUMULADO DEL MES
         DIVIAN = 1
         If PIVACLI%((-1) * NPROVEI%) = 1 Then
            DIVIAN = 1 + ALIVA / 100
         End If
         '
         Call APERBASDAT("CABAN")
         SQLX$ = "SELECT * FROM CAJABANC "
         SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT1 & "' "
         SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT1 & "' "
         SQLX$ = SQLX$ + "AND TipoMovim = 'OPAG' "
         SQLX$ = SQLX$ + "AND NuProve = " & NPROVEI%
'        Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
         '
         Dim CABATEMP As ADODB.Recordset
         Set CABATEMP = New ADODB.Recordset
25       On Error Resume Next
         CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

         If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
            On Error GoTo 0
            Call CapturaErrorOpen
           GoTo 25
         End If
         On Error GoTo 0
  
         With CABATEMP
           On Error Resume Next
           If Not (.EOF And .BOF) Then
             On Error GoTo 0
             Do While Not .EOF
               If !Status < 9 Then   ' para no condiderar anulados
                 VALBASIMP# = 0
                 On Error Resume Next
                 VALBASIMP# = !BASEIMPO
                 On Error GoTo 0
                 If VALBASIMP# < 0.005 Then
                   VALBASIMP# = !ValAbsComp / DIVIAN
                 End If
                 ACUBASANT# = ACUBASANT# + VALBASIMP#
               End If
             .MoveNext
             Loop
           End If
           On Error GoTo 0
         End With
         CABATEMP.Close
         Set CABATEMP = Nothing
         '
         ACURETIB# = 0
         Call APERBASDAT("CABAN")
         SQLX$ = "SELECT * FROM CAJABANC "
         SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= '" & DESDAT1 & "' "
         SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= '" & HASDAT1 & "' "
         SQLX$ = SQLX$ + "AND TipoMovim = 'RTIB' "
         SQLX$ = SQLX$ + "AND NuProve = " & NPROVEI%
         'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
         '
         Set CABATEMP = New ADODB.Recordset
14     On Error Resume Next
         CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
         If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
            On Error GoTo 0
            Call CapturaErrorOpen
            GoTo 14
         End If
         On Error GoTo 0
         '
         With CABATEMP
           On Error Resume Next
           If Not (.EOF And .BOF) Then
             On Error GoTo 0
             Do While Not .EOF
               ACURETANT# = ACURETANT# + !ValAbsComp
             .MoveNext
             Loop
           End If
           On Error GoTo 0
         End With
         CABATEMP.Close
         Set CABATEMP = Nothing
       '\\\OT-05-0319-WAR-24/05/05///
      End If
   End If
   '
   Label29 = "": Label27 = ""
   NTCAN# = NETOCANCE#
   BASEIMTO# = ACUBASANT# + NTCAN#  ' EN LA OPCION NORMAL, NO ACUMULA MENSUALMENTE
   IRETIB# = 0
   '
   If CATIB1% <= 8 Then
     PORCEIBRU# = ALIRETBRU#(NPROVEI%, MINOIMPO#, MEDIPAGO%)
     Label28 = FORMATNUM$(MINOIMPO#, "F12.2")
     If BASEIMTO# >= MINOIMPO# Then
       IRETIB# = BASEIMTO# * PORCONVE / 100 * PORCEIBRU# / 100 - ACURETANT#
       Label24 = Str$(MEDIPAGO%)
       Label29 = FORMATNUM$(PORCEIBRU#, "F8.2")
       Label27 = FORMATNUM$(PORCONVE, "F8.2")
     End If
   End If
   '
   If MONEMI% > 1 Then
     IRETIB# = ROUND(IRETIB# / TICAMONE(MONEMI%))
   End If
   '
   Label14 = TRIM$(FORMATNUM$(IRETIB#, "F12.2"))
   Label23 = TRIM$(FORMATNUM$(NTCAN#, "F12.2"))
99 Screen.MousePointer = 1
   '
End Sub

'
Sub CALRETGAN()
   '
   Static CTCALCU%, ALIGAN(8), MINIMPO#(8), MEPAGX%(8), RETEMIN(8) '27/04/07 (OT-07-0148)
   Static NPROVEA%, NPROVEB%, NOPAGAN$, MESANANT$, CATIB1%
   Static ACUBASANT#, ACURETANT#
   '
'Call CALRETGANVIE
'Exit Sub
   '
   If XVALO(Label18) < 0.005 Then
     Label30 = ""
     Label23 = FORMATNUM$(NETOCANCE#(2), "F12.2")
     GoTo 99
   End If
   '
   If RETGANMAN% = 1 Then GoTo 99
   '
   NPROVEI% = XVALO(Text1) ' SI ES PROVEEDOR COMERCIO EXTERIOR SE VA
   If PIVACLI%(NPROVEI% * (-1)) = 5 Then
      GoTo 99
   End If
   '
   ' CARGA LA TABLA DE CATEGORIAS DE RETENCION
   Screen.MousePointer = 11
   If CTCALCU% = 0 Then
     CTCALCU% = -1
     For U& = 1 To ULTREG&("TACRETG")
        TCX$ = REGLEIDO$("TACRETG", U&)
        TCY% = Asc(Mid$(TCX$, 1, 1))
        If TCY% > 0 Then
          If TCY% <= 8 Then
            ALIGAN(TCY%) = CVS(Mid$(TCX$, 41, 4))
              If ALIGAN(TCY%) > 0 Then CTCALCU% = 1
            MINIMPO#(TCY%) = CVD(Mid$(TCX$, 33, 8))
            MEPAGX%(TCY%) = CVI(Mid$(TCX$, 45, 2)) 'ANTERIORMENTE EXISTIA UNA LONGITUD DE 12 ERRONEA
            RETEMIN(TCY%) = CVS(Mid$(TCX$, 47, 4)) 'CARGA EL VECTOR CON EL VALOR DE LA RETENCION MINIMA
          End If
        End If
     Next U&
     Call CIERRARCH("TACRETG")
   End If
   '
   If CTCALCU% < 0 Then
     Label23 = FORMATNUM$(NETOCANCE#(2), "F12.2")
     GoTo 99
   End If
   '
   'NPROVEI% = XVALO(Text1) LO ESTA LEYENDO ARRIBA
   SHOWRETGA.Label2 = TRIM$(Str$(NPROVEI%))
   If TICUEPRO%((-1) * NPROVEI%) = 1 Then
     Label23 = FORMATNUM$(NETOCANCE#(2), "F12.2")
     GoTo 99 ' RETENCION CERO EN CUENTAS DE GASTOS
   End If
   '
   ' determinar categoria de Ganancias SEGúN PROVEEDOR
   If NPROVEI% <> NPROVEB% Then
     CATIB1% = 0
     REBLA$ = REGBLAN$("DARGAPRO")
     RERET$ = FILTERGETRECORD$("DARGAPRO", 1, MKI$(NPROVEI%))
     CATIB1% = 0   ' VALOR POR DEFAULT
     If RERET$ <> REBLA$ Then CATIB1% = Asc(Mid$(RERET$, 3, 1))
     If CTCALCU% < 0 Then CATIB1% = 0   ' EMPRESA NO AGENTE DE RETENCION
     If TICUEPRO%((-1) * NPROVEI%) = 1 Then CATIB1% = 0 ' CUENTA DE GASTOS
     Label32 = (MEPAGX%(CATIB1%))
     NPROVEB% = NPROVEI%
   End If
   SHOWRETGA.Label7 = TRIM$(CStr(CATIB1%))
   SHOWRETGA.Label15 = FORMATNUM(MINIMPO#(CATIB1%), "F12.2")
   '
   ' CARGA LIST CON ORDENES DE PAGO Y RETENCIONES ANTERIORES
   If NPROVEI% <> NPROVEA% Or Text11 <> NOPAGAN$ Or Mid$(Text9, 4) <> MESANANT$ Then
      SHOWRETGA.Label9 = "": SHOWRETGA.Label10 = ""
      ACUBASANT# = 0: ACURETANT# = 0
      NPROVEA% = NPROVEI%: NOPAGAN$ = Text11: MESANANT$ = Mid$(Text9, 4)
      '
      CADETANT% = 0
      ReDim Preserve DETANT$(CADETANT%)
      '
      SHOWRETGA.Label9 = ""
      SHOWRETGA.Label10 = ""
      SHOWRETGA.Label11 = ""
      SHOWRETGA.Label12 = ""
      SHOWRETGA.Label14 = ""
      SHOWRETGA.Label15 = ""
      SHOWRETGA.Label16 = ""
      SHOWRETGA.Label17 = ""
      SHOWRETGA.Text1 = ""
      SHOWRETGA.List1.Clear
      '
      DESMEI% = FECHANUM("01" + Mid$(Text9, 3))
      'DESDAT = CDbl(CVDAT(DESMEI%))
      'JANDY SQL
      DESDAT = FETEXCOR1$(DESMEI%)
      
      HASMEI = DIPOST(DESMEI%)
      While Mid$(FECHATEX$(DIPOST(HASMEI)), 4, 2) = Mid$(Text9, 4, 2)
        HASMEI = DIPOST(HASMEI)
      Wend
      HASMEII% = HASMEI
      'HASDAT = CDbl(CVDAT(HASMEII%))
      'JANDY SQL
      HASDAT = FETEXCOR1$(HASMEII%)
      
      ' determinar monto imponible - ACUMULADO DEL MES
      DIVIAN = 1
      If PIVACLI%((-1) * NPROVEI%) = 1 Then
        DIVIAN = 1 + ALIVA / 100
      End If
      'BASEIMTO# = 0
      Call APERBASDAT("CABAN")
      'JANDY SQL
      
      Dim CABATEMP As ADODB.Recordset
      Set CABATEMP = New ADODB.Recordset
      CABATEMP.CursorLocation = adUseClient
      
      SQLX$ = "SELECT * FROM CAJABANC "
      'SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
      'SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
      SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
      SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
      SQLX$ = SQLX$ + "AND (TipoMovim = 'OPAG' OR TipoMovim = 'RTGAN') "
      SQLX$ = SQLX$ + "AND NuProve = " & NPROVEI% & " "
      SQLX$ = SQLX$ + "ORDER BY CodiCom_Lar ASC"
25    On Error Resume Next
      CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
      End If
      On Error GoTo 0
      'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      With CABATEMP
        'On Error Resume Next
        '.MoveFirst
        'If Err < 1 Then
        '  On Error GoTo 0
        If Not (.EOF And .BOF) Then
          .MoveFirst
          Do While Not .EOF
            If !Status < 9 Then   ' para no condiderar anulados
               VALBRUIMP# = 0
               VALBASIMP# = 0
               VALRETGAN# = 0
               If !TipoMovim = "OPAG" Then
                   VALBRUIMP# = !ValAbsComp
                   VALBASIMP# = 0
                   On Error Resume Next
                   VALBASIMP# = !BASEIMPO
                   On Error GoTo 0
                   If VALBASIMP# > !ValAbsComp Then VALBASIMP# = 0
                   If VALBASIMP# < 0.005 Then
                     VALBASIMP# = !ValAbsComp / DIVIAN
                   End If
                   '
                   ' VERIFICA SI LA ORDEN DE PAGO CANCELO UNA NOTA DE DEBITO POR CHEQUE RECHAZADO.
                   ' EN CASO AFIRMATIVO LA OBVIA PARA NO CALCULAR GANANCIAS POR SEGUNDA VEZ
                   DOCU$ = SAFETEXT$(!CODICOM_LAR)
                   DOCUAPLI$ = SAFETEXT$(VALOBADA("APLIPAGO", "CODOCOR_LAR", "CODOCAP_LAR='" & DOCU$ & "' AND CODOCOR_LAR LIKE 'ND-%'"))
                   If TRIM$(DOCUAPLI$) <> "" Then
                        CUERECHACHEQUE% = CUEINT(Control$("SVD451S", "CUEVE-8"))
                        Condi$ = "CODICOM_LAR='" & DOCUAPLI$ & "' AND NUPROVE=" & Text1 & " AND TIPOMOVIM='COMPRA' AND CUECONTAI=" & CUERECHACHEQUE%
                        If CantRegistros("CAJABANC", Condi$) > 0 Then VALBASIMP# = 0 ' LO EXCLUYE
                   End If
                   '
                   ACUBASANT# = ACUBASANT# + VALBASIMP#
                 ElseIf !TipoMovim = "RTGAN" Then
                   VALRETGAN# = !ValAbsComp
                   ACURETANT# = ACURETANT# + VALRETGAN#
               End If
               '
               For KKI% = 1 To CADETANT%
                 If TRIM$(Left$(DETANT$(KKI%), 18)) = TRIM$(!CODICOM_LAR) Then
                   IBRUCOM# = XVALO(Mid$(DETANT$(KKI%), 21, 14)) + VALBRUIMP#
                   IBASIMP# = XVALO(Mid$(DETANT$(KKI%), 37, 14)) + VALBASIMP#
                   ACURETE# = XVALO(Mid$(DETANT$(KKI%), 53, 12)) + VALRETGAN#
                   TTXYZ$ = DETANT$(KKI%)
                   Call REPLA(TTXYZ$, FORMATNUM$(IBRUCOM#, "F14.2"), 21, 14)
                   Call REPLA(TTXYZ$, FORMATNUM$(IBASIMP#, "F14.2"), 37, 14)
                   Call REPLA(TTXYZ$, FORMATNUM$(ACURETE#, "F12.2"), 53, 12)
                   DETANT$(KKI%) = TTXYZ$
                   GoTo 19
                 End If
               Next KKI%
               CADETANT% = CADETANT% + 1
               ReDim Preserve DETANT$(CADETANT%)
               TTXYZ$ = Space$(64)
               IBRUCOM# = VALBRUIMP#
               IBASIMP# = VALBASIMP#
               ACURETE# = VALRETGAN#
               Call REPLA(TTXYZ$, !CODICOM_LAR, 1, 18)
               Call REPLA(TTXYZ$, FORMATNUM$(IBRUCOM#, "F14.2"), 21, 14)
               Call REPLA(TTXYZ$, FORMATNUM$(IBASIMP#, "F14.2"), 37, 14)
               Call REPLA(TTXYZ$, FORMATNUM$(ACURETE#, "F12.2"), 53, 12)
               
               DETANT$(KKI%) = TTXYZ$
               '
19           End If
          .MoveNext
          Loop
        End If
      End With
      CABATEMP.Close
      Set CABATEMP = Nothing
      
      '
      ACUBASANT# = 0: ACURETANT# = 0
      For KKI% = 1 To CADETANT%
        SHOWRETGA.List1.AddItem DETANT$(KKI%)
        ACUBASANT# = ACUBASANT# + XVALO(Mid$(DETANT$(KKI%), 37, 14))
        ACURETANT# = ACURETANT# + XVALO(Mid$(DETANT$(KKI%), 53, 12))
     Next KKI%
   End If
   '
   SHOWRETGA.Label12 = Label18
   SHOWRETGA.Label11 = FORMATNUM(NETOCANCE#(2), "F12.2")
   SHOWRETGA.Label9 = FORMATNUM(ACUBASANT#, "F12.2")
   SHOWRETGA.Label10 = FORMATNUM(ACURETANT#, "F10.2")
   '
   BASEIMTO# = ACUBASANT# + NETOCANCE#(2)
   SHOWRETGA.Label14 = FORMATNUM(BASEIMTO#, "F12.2")
   MINOIM# = MINIMPO#(CATIB1%)
   SHOWRETGA.Label15 = FORMATNUM(MINOIM#, "F12.2")
   BASEIMNE# = BASEIMTO# - MINOIM#
     If BASEIMNE# < 0 Then BASEIMNE# = 0
   SHOWRETGA.Label16 = FORMATNUM(BASEIMNE#, "F12.2")
   '
   RETGATO# = 0
   If BASEIMNE# > 0 Then
     RETGATO# = BASEIMNE# * ALIGAN(CATIB1%) / 100 ' - ACURETANT#
     RFF$ = RECFILT$("ESCALGAN", 1, MKI$(CATIB1%))
     If Len(RFF$) > 0 Then
       Call FILTERFILE("ESCALGAN", "!ESCALGAN", 1, MKI$(CATIB1%))
       For KU& = 1 To ULTREG&("!ESCALGAN")
         ESCX$ = REGLEIDO$("!ESCALGAN", KU&)
         If BASEIMNE# > CVD(Mid$(ESCX$, 3, 8)) Then
           If BASEIMNE# <= CVD(Mid$(ESCX$, 11, 8)) Then
             RETGATO# = CVD(Mid$(ESCX$, 19, 8)) + (BASEIMNE# - CVD(Mid$(ESCX$, 3, 8))) * CVS(Mid$(ESCX$, 27, 4)) / 100
             GoTo 98
           End If
         End If
       Next KU&
     End If
   End If
   '
98 SHOWRETGA.Label17 = FORMATNUM(RETGATO#, "F10.2")
   RETGACAL# = RETGATO# - ACURETANT#
   If RETGACAL# < 0 Then RETGACAL# = 0
   '27/04/07 (OT-07-0148)
   If RETGACAL# < RETEMIN(CATIB1%) Then 'SI LA RETENCION ES MENOR AL VALOR MINIMO DE RETENCION
       RETGACAL# = 0    'NO SE RETIENE
   End If
   '27/04/07 (FIN)
   SHOWRETGA.Text1 = FORMATNUM(RETGACAL#, "F10.2")
   '
   Label30 = TRIM$(FORMATNUM$(RETGACAL#, "F12.2"))
   Label32 = (MEPAGX%(CATIB1%))
   '
   ' VERIFICA SI EXISTEN COMPROBANTES 'M' DONDE TENGA QUE RETENER
   ' SOLO SI EL CALCULO GENERAL DE RETENCION DE GANANCIAS NO ARROJO UN IMPORTE A RETENER
   RetencionTotal# = 0
   For IndiceComprobantes% = 1 To List2.ListCount - 1
        Renglon$ = List2.LIST(IndiceComprobantes%)
        
        Comprobante$ = Mid$(Renglon$, 1, 18)
        ImporteACancelar# = XVALO(TRIM$(Mid$(Renglon$, 70, 14)))
        ' *** VERIFICA QUE HAYA ELEGIDO EL COMPROBANTE MEDIANTE LA EXISTENCIA DEL IMPORTE QUE ESTA POR CANCELAR
        If Mid$(Comprobante$, 4, 1) = "M" And Abs(ImporteACancelar#) > 0.005 Then
            ImporteNeto# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP) - SUM(IHABECOMP)", "CODICOM_LAR='" & Comprobante$ & "' AND TIPOMOVIM='COMPRA' AND NUPROVE=" & Text1))
            ImporteNetoTotal# = ImporteNetoTotal# + ImporteNeto#
        End If
        
   Next IndiceComprobantes%
   ' SI TIENE COMPROBANTES 'M', TOMA EL CALCULO DE GANANCIAS DE ESTE Y PASA POR ENCIMA EL GENERAL
   If ImporteNetoTotal# > 0.005 Then Label30.Caption = TRIM$(FORMATNUM$(ImporteNetoTotal# * 0.06, "F12.2"))
   '
99 Screen.MousePointer = 1
   '
End Sub
'

Sub PRICORETGA()
'COPIE LA RUTINA DE IMPRSION DE FORMULARIOI DE RET ING. BRU.
'HASTA ACA LLEGUE EL 04/02/05
   If XVALO(Label30) >= 0.005 Then
      NURETGA& = XVALO(Label34) 'FIJARSE QUE LABEL TIENE EL NUME RET
      Screen.MousePointer = 11
      FECHDOC$ = FETEXTO$(Int(FECHANUM(Text9.TEXT)))
      NUMEDOC$ = TRIM$(Str$(NURETGA&))
      While Len(NUMEDOC$) < 8
         NUMEDOC$ = "0" + NUMEDOC$
      Wend
      'NUMEDOC$ = "0000-20" + Mid$(Text9, 7, 2) + "-" + NUMEDOC$
      NUMEDOC$ = NUMEDOC$
      TIPODOC$ = "Retencion Ganancias"
      '
      CODPARAM$(1) = "IMP-TOTA"
      CODPARAM$(2) = "IMPO-CAN"
      CODPARAM$(3) = "DOCUORIG"
      '\\\OT-5-596-FG-24/08/05///
      CODPARAM$(4) = "REF-MAT1" ' REGIMEN DE LA RETENCION
      '\\\OT-5-596-FG-FIN///
      CAPARDOC% = 4
      '
      DOCPARAM$(1) = TRIM$(Label18) 'MONTO QUE ORIGINA RET
      DOCPARAM$(2) = TRIM$(Label30) 'importe retenido
      DOCPARAM$(3) = "Orden de Pago Nro. " & Mid$(TRIM$(Text11.TEXT), 4)
      '
'      For U& = 1 To List2.ListCount
'        'List2.ListIndex = U& - 1
'        X$ = List2.List(U& - 1)
'        DOCU$ = TRIM$(Mid$(X$, 1, 20))
'        SALPE# = CDbl(XVALO(Mid$(X$, 56, 13)))
'        CANCE# = CDbl(XVALO(Mid$(X$, 70, 14)))
'        If Abs(CANCE#) >= 0.005 Then
'          If DOCPARAM$(3) <> "" Then
'            DOCPARAM$(3) = DOCPARAM$(3) + ", "
'          End If
'          If Len(DOCPARAM$(3)) < 24 Then
'              DOCPARAM$(3) = DOCPARAM$(3) + DOCU$ ' Comprobantes que originan la retencion
'            Else
'              DOCPARAM$(3) = DOCPARAM$(3) + ".........................................."
'          End If
'        End If
'
'      Next U&
      '
      '\\\OT-5-596-FG-24/08/05///
      DOCPARAM$(4) = ""
      NPROVE% = CInt(Text1) ' NUMERO DE PROVEEDOR
'      JJ$ = RECFILT$("DARGAPRO", 1, MKI$(NPROVE%))
'      XX$ = REGLEIDO$("DARGAPRO", CVS(JJ$)) ' LEE CONFIGURACION DE LA RETENCION
      XX$ = FILTERGETRECORD$("DARGAPRO", 1, MKI$(NPROVE%)) ' LEE CONFIGURACION DE LA RETENCION
      CATEG1& = Asc(Mid$(XX$, 3, 1)) ' OBTIENE NUMERO DE CATEGORIA
      XX$ = REGLEIDO$("TACRETG", CATEG1& + 1) ' LEE TABLA DE CATEGORIAS
      DESREG$ = TRIM$(Mid$(XX$, 2, 31)) ' OBTIENE DESCRIPCION DE LA CATEGORIA
      DOCPARAM$(4) = DESREG$
      '\\\OT-5-596-FG-FIN///
      '
      CACOLTAB1% = 0
      '
      DESTIDOCNUE = (-1) * XVALO(Text1)
      CAITAB1% = 0
      '
      Call PRINTDOC("FORREGAN")   ' Retencion Ganancias
      '
   End If
   '
End Sub

Sub PRICORETIB()
   '
   ' *** RETENCION DE IIBB BUENOS AIRES
   If XVALO(Label14) >= 0.005 Then
      NURETBA& = XVALO(Label25)
      Screen.MousePointer = 11
      FECHDOC$ = FETEXTO$(Int(FECHANUM(Text9.TEXT)))
      NUMEDOC$ = TRIM$(Str$(NURETBA&))
      While Len(NUMEDOC$) < 8
         NUMEDOC$ = "0" + NUMEDOC$
      Wend
      NUMEDOC$ = "0001-" + NUMEDOC$
      TIPODOC$ = "Retencion I.Brutos"
      '
      CODPARAM$(1) = "IMPO-CHE"
      CODPARAM$(2) = "IMPO-PEN"
      CODPARAM$(3) = "IMPO-CAN"
      CODPARAM$(4) = "POR-DESC"
      CODPARAM$(5) = "DOCU-APL"
      CODPARAM$(6) = "TIPO-VAL"
      CAPARDOC% = 6
      '\\\OT-05-0514-WAR-14/7/05///
        ' determinar categoria de I.Brutos
      CATEIB2$ = ""
      NPROVEI% = XVALO(Text1)
      If TICUEPRO%((-1) * NPROVEI%) = 0 Then ' RETENCION CERO EN CUENTAS DE GASTOS
        RERET$ = FILTERGETRECORD$("DARIBPRO", 1, MKI$(NPROVEI%))
        CATIB1% = Asc(Mid$(RERET$, 3, 1))
        CATEIB2$ = ECOARCH("TACATIBP", Chr$(CATIB1%))
      End If
      '\\\OT-05-0514-WAR-FIN///
      DOCPARAM$(1) = TRIM$(Label23)
      DOCPARAM$(2) = TRIM$(Label29)
      DOCPARAM$(3) = TRIM$(Label14)
      DOCPARAM$(4) = TRIM$(Label27)
      DOCPARAM$(5) = TRIM(Text11)
      DOCPARAM$(6) = TRIM(CATEIB2$) 'CATEGORIA
      '
      CACOLTAB1% = 0
      '
      DESTIDOCNUE = (-1) * XVALO(Text1)
      CAITAB1% = 0
      '
      Call PRINTDOC("FORRETIB")   ' Retencion I.Brutos
      '
   End If
   
   '
   ' *** RETENCION DE IIBB C.A.B.A.
   If XVALO(Label38) >= 0.005 Then
   
        FormularioIIBB$ = TRIM$(Control$("", "FOREIBCA"))
        NroRetencion$ = TRIM$(Label41)
        While Len(NroRetencion$) < 8: NroRetencion$ = "0" & NroRetencion$: Wend
        
        FILTX$ = NroRetencion$ & ";" & Text1 & ";" & Text9 & ";" & CStr(XVALO(NroRetencion$))
        Call STDFORM(FormularioIIBB$, FILTX$, "PRINT")
        
        RESPUESTA% = COMALTER%("Retención de IIBB CABA Impresa.\\Repetir Impresión\Continuar")
        If RESPUESTA% = 1 Then
            FILTX$ = NroRetencion$ & ";" & Text1 & ";" & Text9 & ";" & CStr(XVALO(NroRetencion$))
            Call STDFORM(FormularioIIBB$, FILTX$ & ";", "PRINT")
        End If
                
   End If
   
   '
   ' *** RETENCION DE IIBB MISIONES
   If XVALO(Label45) >= 0.005 Then
   
        FormularioIIBB$ = TRIM$(Control$("", "FOREIBMI"))
        NroRetencion$ = CStr(NumeroRetencionIBMisiones&)
        While Len(NroRetencion$) < 8: NroRetencion$ = "0" & NroRetencion$: Wend
        
        FILTX$ = NroRetencion$ & ";" & Text1 & ";" & Text9 & ";" & CStr(XVALO(NroRetencion$))
        Call STDFORM(FormularioIIBB$, FILTX$, "PRINT")
        
        FILTX$ = NroRetencion$ & ";" & Text1 & ";" & Text9 & ";" & CStr(XVALO(NroRetencion$))
        Call STDFORM(FormularioIIBB$, FILTX$ & ";", "PRINT")
                
   End If
   
   '
   ' *** RETENCION DE IIBB TUCUMAN
   If XVALO(Label12) >= 0.005 Then
   
        FormularioIIBB$ = TRIM$(Control$("", "FOREIBTU"))
        NroRetencion$ = CStr(NumeroRetencionIBTucuman&)
        While Len(NroRetencion$) < 8: NroRetencion$ = "0" & NroRetencion$: Wend
        
        FILTX$ = NroRetencion$ & ";" & Text1 & ";" & Text9 & ";" & CStr(XVALO(NroRetencion$)) & ";" & "ORIGINAL"
        Call STDFORM(FormularioIIBB$, FILTX$, "PRINT")
        
        FILTX$ = NroRetencion$ & ";" & Text1 & ";" & Text9 & ";" & CStr(XVALO(NroRetencion$)) & ";" & "DUPLICADO"
        Call STDFORM(FormularioIIBB$, FILTX$ & ";", "PRINT")
                
   End If
   
End Sub

Sub PRICORETIVA()
   '
   ' *** RETENCION DE IVA
   If XVALO(Label49) >= 0.005 Then
   
        FormularioRetIva$ = TRIM$(Control$("", "FORETIVA"))
        
        NroRetencion$ = CStr(NumeroRetencionIva&)
        CondicionSql$ = "TIPOMOVIM='RTIVA' AND CODICOM_LAR='" & OrdenDePagoActiva$ & "' AND NRODERETENCION=" & NroRetencion$
        FechaRetencion% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", CondicionSql$))
        ANIO$ = Format(CVDAT(FechaRetencion%), "yyyy")
        While Len(NroRetencion$) < 8: NroRetencion$ = "0" & NroRetencion$: Wend
        NroRetencion$ = "0000-" & ANIO$ & "-" & NroRetencion$
        
        FILTX$ = NroRetencion$ & ";" & Text1 & ";" & Text9 & ";" & CStr(XVALO(Mid$(NroRetencion$, 11)))
        Call STDFORM(FormularioRetIva$, FILTX$, "PRINT")
        
        FILTX$ = NroRetencion$ & ";" & Text1 & ";" & Text9 & ";" & CStr(XVALO(Mid$(NroRetencion$, 11)))
        Call STDFORM(FormularioRetIva$, FILTX$ & ";", "PRINT")
                
   End If
   '
End Sub

Sub GRAORDECHPRO()
    '
    Call APERBASDAT("PROVE")
    NRX12& = XVALO(Text1)
    'With Proved12
      XXX$ = ""
      '/// OT-8-717-28/08/08
      '/// OT-8-942-14/10/08 - ESTE CAMBIO SE HIZO PORQUE SI LA RAZON
      'SOCIAL CONTIENE UNA COMILLA SIMPLE, EN EL EJECUTACOMANDO SE PRODUCE UN ERROR
      'AL EJECUTARSE LA CONSULTA SQL
      TXTX$ = AJUSTI$(Text7, 48)
      TXTX$ = REPLACAR$(TXTX$, "'", "")
      strsql = "Update Proved12 set Ord_Cheq = '" & TXTX$ & "'"
      '/// OT-8-942-14/10/08
      '/// OT-8-717-FIN
      strsql = strsql & " where Nume_Cli = " & NRX12&
      Call EJECUTACOMANDO(strsql)
    
      '.FindFirst ("Nume_Cli = " & NRX12&)
      'If .NoMatch = False Then
      '  .Edit
      '    !Ord_Cheq = AJUSTI$(Text7, 48)
      '  .Update
      'End If
    'End With
    '
End Sub

Function NETOCANCE#(Optional MODO%)
      NPROVEI% = Abs(XVALO(Text1))
      SUNECAN# = 0
      NDEXCLUIDA% = 0
      '
      ' CUENTA CONTABLE DE RECHAZO DE VALORES. SE CONSIDERA PARA OBVIAR ESTA OP
      ' SI SE ESTA PAGANDO UNA NOTA DE DEBITO POR CHEQUE RECHAZADO
      CUERECHACHEQUE% = CUEINT(Control$("SVD451S", "CUEVE-8"))
      For U& = 1 To List2.ListCount
        Z$ = List2.LIST(U& - 1)
        CANCE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
        'AGREGADO PARA MULTIMONEDA
        If MONEMI% > 1 Then
           CANCE# = ROUND(CANCE# * TICAMONE(MONEMI%))
        End If
        'FIN AGREGADO MULTIMONEDA
        
        'If CANCE# > 0 Then
        If Abs(CANCE#) > 0 Then ' SE INCLUYE ABS PARA CONSIDERAR LAS NOTAS DE CREDITO
          DOCU$ = TRIM$(Left$(Z$, 20))
          If DOCU$ = "Anticipo no Aplicado" Then
              INECAN# = 0
            ElseIf Left$(DOCU$, 2) = "OP" Or Left$(DOCU$, 2) = "NC" Then
              INECAN# = 0
            ElseIf Left$(DOCU$, 2) = "ND" Then
              ' EXCLUYE EL CALCULO SOBRE NOTAS DE DEBITO POR CHEQUE RECHAZADO. PARA ESO VERIFICA QUE LA MISMA
              ' FORME PARTE DEL DETALLE
              If MODO% = 2 Then
                Condi$ = "CODICOM_LAR='" & DOCU$ & "' AND NUPROVE=" & Text1 & " AND TIPOMOVIM='COMPRA' AND CUECONTAI=" & CUERECHACHEQUE%
                NDEXCLUIDA% = 0
                If CantRegistros("CAJABANC", Condi$) > 0 Then NDEXCLUIDA% = 1
              Else
                FEMIA = FECHANUM(Mid$(Z$, 23, 8))
                INEFA# = INEFAC#(DOCU$, FEMIA, (-1) * NPROVEI%)
                INECAN# = 0
                On Error Resume Next
                If XVALO(Mid$(Z$, 32, 13)) > 0.00005 Then INECAN# = CANCE# * INEFA# / XVALO(Mid$(Z$, 32, 13))
                On Error GoTo 0
              End If
            Else
              FEMIA = FECHANUM(Mid$(Z$, 23, 8))
              INEFA# = INEFAC#(DOCU$, FEMIA, (-1) * NPROVEI%)
              INECAN# = 0
              On Error Resume Next
              If XVALO(Mid$(Z$, 32, 13)) > 0.00005 Then INECAN# = CANCE# * INEFA# / XVALO(Mid$(Z$, 32, 13))
              On Error GoTo 0
'              If INECAN# < 0.005 Then
'                 TTXX$ = "Error Recuperable de Ejecución:\Imposible Determinar Base Imponible de " + DOCU$
'                 TTXX$ = TTXX$ + "\del " + FECHATEX$(FEMIA) + " - Proveedor " + TRIM$(Str$(NPROVEI%))
'                 TTXX$ = TTXX$ + "\INEFA# = " + TRIM$(Str$(INEFA#))
'                 TTXX$ = TTXX$ + "\XVALO = " + TRIM$(Str$(XVALO(Mid$(z$, 32, 13))))
'                 TTXX$ = TTXX$ + "\   \Envíe Esta Pantalla s su Soporte FLEXOFT"
'                 Call MENSERR(24, TTXX$)
'              End If
          End If
          If INECAN# < 0.005 Then
            DIVIAN = 1
            If PIVACLI%((-1) * NPROVEI%) = 1 Then
              DIVIAN = 1 + ALIVA / 100
            End If
            INECAN# = CANCE# / DIVIAN
          End If
          If NDEXCLUIDA% = 1 Then INECAN# = 0 ' EXCLUYE A LA NOTA DEBITO DE LA BASE NETA DE CALCULO DE RETENCION
          SUNECAN# = SUNECAN# + INECAN#
        End If
      Next U&
      NETOCANCE# = SUNECAN#
End Function

Function IMPOCANCE_TUCUMAN#(Optional MODO%)

      NPROVEI% = Abs(XVALO(Text1))
      SUNECAN# = 0
      NDEXCLUIDA% = 0
      '
      ' CUENTA CONTABLE DE RECHAZO DE VALORES. SE CONSIDERA PARA OBVIAR ESTA OP
      ' SI SE ESTA PAGANDO UNA NOTA DE DEBITO POR CHEQUE RECHAZADO
      CUERECHACHEQUE% = CUEINT(Control$("SVD451S", "CUEVE-8"))
      For U& = 1 To List2.ListCount
        Z$ = List2.LIST(U& - 1)
        CANCE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
        'AGREGADO PARA MULTIMONEDA
        If MONEMI% > 1 Then
           CANCE# = ROUND(CANCE# * TICAMONE(MONEMI%))
        End If
        'FIN AGREGADO MULTIMONEDA
        
        'If CANCE# > 0 Then
        If Abs(CANCE#) > 0 Then ' SE INCLUYE ABS PARA CONSIDERAR LAS NOTAS DE CREDITO
          DOCU$ = TRIM$(Left$(Z$, 20))
          If DOCU$ = "Anticipo no Aplicado" Then
              INECAN# = 0
            ElseIf Left$(DOCU$, 2) = "OP" Or Left$(DOCU$, 2) = "NC" Then
              INECAN# = 0
            ElseIf Left$(DOCU$, 2) = "ND" Then
              ' EXCLUYE EL CALCULO SOBRE NOTAS DE DEBITO POR CHEQUE RECHAZADO. PARA ESO VERIFICA QUE LA MISMA
              ' FORME PARTE DEL DETALLE
              If MODO% = 2 Then
                Condi$ = "CODICOM_LAR='" & DOCU$ & "' AND NUPROVE=" & Text1 & " AND TIPOMOVIM='COMPRA' AND CUECONTAI=" & CUERECHACHEQUE%
                NDEXCLUIDA% = 0
                If CantRegistros("CAJABANC", Condi$) > 0 Then NDEXCLUIDA% = 1
              Else
                FEMIA = FECHANUM(Mid$(Z$, 23, 8))
                INEFA# = INEIVAFAC#(DOCU$, FEMIA, (-1) * NPROVEI%)
                INECAN# = 0
                On Error Resume Next
                If XVALO(Mid$(Z$, 32, 13)) > 0.00005 Then INECAN# = CANCE# * INEFA# / XVALO(Mid$(Z$, 32, 13))
                On Error GoTo 0
              End If
            Else
              FEMIA = FECHANUM(Mid$(Z$, 23, 8))
              INEFA# = INEIVAFAC#(DOCU$, FEMIA, (-1) * NPROVEI%)
              INECAN# = 0
              On Error Resume Next
              If XVALO(Mid$(Z$, 32, 13)) > 0.00005 Then INECAN# = CANCE# * INEFA# / XVALO(Mid$(Z$, 32, 13))
              On Error GoTo 0
'              If INECAN# < 0.005 Then
'                 TTXX$ = "Error Recuperable de Ejecución:\Imposible Determinar Base Imponible de " + DOCU$
'                 TTXX$ = TTXX$ + "\del " + FECHATEX$(FEMIA) + " - Proveedor " + TRIM$(Str$(NPROVEI%))
'                 TTXX$ = TTXX$ + "\INEFA# = " + TRIM$(Str$(INEFA#))
'                 TTXX$ = TTXX$ + "\XVALO = " + TRIM$(Str$(XVALO(Mid$(z$, 32, 13))))
'                 TTXX$ = TTXX$ + "\   \Envíe Esta Pantalla s su Soporte FLEXOFT"
'                 Call MENSERR(24, TTXX$)
'              End If
          End If
          If INECAN# < 0.005 Then
            DIVIAN = 1
            If PIVACLI%((-1) * NPROVEI%) = 1 Then
              DIVIAN = 1 + ALIVA / 100
            End If
            INECAN# = CANCE# / DIVIAN
          End If
          If NDEXCLUIDA% = 1 Then INECAN# = 0 ' EXCLUYE A LA NOTA DEBITO DE LA BASE NETA DE CALCULO DE RETENCION
          SUNECAN# = SUNECAN# + INECAN#
        End If
      Next U&
      IMPOCANCE_TUCUMAN# = SUNECAN#
      
End Function

Function VERINUMOP%(NROPRO&)
    '
    Static FLINUME
    '
    Screen.MousePointer = 11
    '
    'jandy sql Flinume es string
    'If FLINUME < 1 Then
    If FLINUME = "" Then
       FENUE$ = Control$("", "FENUORPA")
       FLIMI = CVDAT(FECHANUM(FENUE$))
       If FLIMI > Now Then FLIMI = CVDAT(HOY(HO$))
       'FLINUME = CDbl(FLIMI)
        FEX = FECHANUM(FENUE$)
        FEXII% = FEX
        FLINUME = FETEXCOR1$(FEXII%)
    End If
    '
    NROMA& = 0
22  Call APERBASDAT("CABAN")
    Dim OrpaTemp As ADODB.Recordset
    Set OrpaTemp = New ADODB.Recordset
    OrpaTemp.CursorLocation = adUseClient
    '
    SQLX$ = "SELECT * FROM CajaBanc WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE FechEmisi >= '" & FLINUME & "' "
    SQLX$ = SQLX$ + " AND TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + " AND CodiCom_Cor LIKE 'OP%' "
    SQLX$ = SQLX$ + " AND NumeComp = " & NROPRO&
    Set OrpaTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With OrpaTemp
      If (.EOF And .BOF) Then 'SI NO LO ENCUENTRA
           VERINUMOP% = 1
      Else
           VERINUMOP% = (-1)
      End If
    End With
    OrpaTemp.Close
    Set OrpaTemp = Nothing
    '
End Function

Sub VERIRECHA(NPRO%)
  '
  Call BLANARCH("!INDIVARE")
  For J& = 1 To ULTREG&("RECHLIN")
    X$ = REGLEIDO$("RECHLIN", J&)
    NPRO2% = CVI(Mid$(X$, 9, 2))
    If NPRO2% = NPRO% Then
      FE2% = CVI(Mid$(X$, 113, 2))
      USUPRO% = CVI(Mid$(X$, 115, 2))
      If FE2% <= 0 Then 'SI NO TIENE FECHA DE PROCESADO
        If USUPRO% <= 0 Then 'SI NO TIENE USUARIO DE PROCESADO
          'AGREGA LA LISTA DE SELECCION
          NROREC& = CVS(Mid$(X$, 1, 4))
          FEMI1$ = FECHATEX$(CVI(Mid$(X$, 7, 2)))
          REFERE2$ = FEMI1$ + " - " + Left$(rasocli$(NPRO% * (-1)), 30)
          'CODIGO VIEJO
'          ZZ$ = Space(64)
'          Call REPLA(ZZ$, MKS$(NROREC&), 1, 4)
'          Call REPLA(ZZ$, REFERE2$, 5, 60)
'          Call REGAPP("!INDIVARE", ZZ$)
          '\\\OT-05-0665-WAR-07/10/05///
          ZZ$ = Space(64)
          Call REPLA(ZZ$, MKS$(NROREC&), 1, 4)
             IMPORE# = CVD(Mid$(X$, 105, 8))
             IMPOR1$ = FORMATNUM$(IMPORE#, "F10.2")
          Call REPLA(ZZ$, REFERE2$ + IMPOR1$, 5, 60)
          Call REGAPP("!INDIVARE", ZZ$)
          '\\\OT-05-0665-WAR-FIN///
          HAYPEND% = 1
        End If
      End If
    End If
  Next J&
  Call CIERRARCH("!INDIVARE")
  If HAYPEND% = 0 Then Exit Sub 'NO HAY RECHAZOS PENDIENTES DE PROCESO
  Call FRESHECHO("!INDIVARE")
  '
  Call SELECHO("!INDIVARE")
  NPX& = Val(VALACT1$("!INDIVARE"))
  '
  If NPX& > 0 Then
    Call MUESTRARECH(NPX&)
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = " " + NUDOCU$ + " "
    If COMALTER%("¿Desea Procesar el Rechazo en Línea Número: " & NUDOCU$ & " ?\\Si, Procesar.\No, Continuar. ") = 1 Then
      If DERACCE%("PROCEREC", "Procesar Rechazos en Línea.") = 2 Then
        Call PROCERECH(NPX&) 'PROCESA RECHAZO EN LÍNEA
      End If
    End If
  End If
  '
99 End Sub

Sub MUESTRARECH(NPX&)
 '
    TIDOCUM$ = "Rechazo en Línea"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = " " + NUDOCU$ + " "
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 99
    '
5   Call MUESTRADOC(DOCUNU&)
    '
99 End Sub

Sub PROCERECH(NPX&)
  '
  NPRX% = XVALO(Text1)
  RFF$ = RECFILT$("RECHLIN", 1, MKS$(NPX&))
  For KU& = 1 To Len(RFF$) Step 4
     REGO& = CVS(Mid$(RFF$, KU&, 4))
     XXX$ = REGLEIDO$("RECHLIN", REGO&)
     If CVI(Mid$(XXX$, 9, 2)) = NPRX% Then
       FEX1% = HOY(HO$)
       Call REPLA(XXX$, MKI$(USNBR%), 115, 2)
       Call REPLA(XXX$, MKI$(FEX1%), 113, 2)
       Call GRAREG("RECHLIN", XXX$, REGO&)
     End If
  Next KU&
  '
End Sub
Sub LimpTotOrPgo()
  '\\\OT-08-0889-OD-25/09/08///
  Label14 = ""
  Label45 = ""
  Label38 = ""
  Label49 = ""
  Label18 = ""
  Label20 = ""
  Label30 = ""
  Label21 = ""
  '\\\OT-08-0889-OD-FIN///


End Sub

Sub ACTULISTAOP()
    'OD-02/12/11 SE AGREGA DISTINCT
    'NOTA PARA EJECUTAR UN DISTINCT CON ORDER BY TENER EN CUENTA QUE LOS CAMPOS DEL ORDENAMIENTO
    'DEBEN ESTAR INCLUIDOS EN LA SELECT SI NO GENERA ERROR
    'EL DISTINCT SE REALIZA SOLO SOBRE EL PRIMER CAMPO INMEDIATO.
    'SI NO SE QUIEREN PRESENTAR LOS CAMPOS AGREGADOS EN LA SELECT INDICAR LONGITUD 0
    Condi$ = "TipoMovim = 'OPAG' AND CODICOM_LAR LIKE 'OP%' Order by FECHEMISI ASC, CODICOM_LAR ASC "
    Call CARGALISEL("!INDIORPA", "CAJABANC", "NUMECOMP/I8;FECHCONTA/D8;REFECOM/A32;FECHEMISI/D0;CODICOM_LAR/A0 ", Condi$, "DISTINCT/NOMESS")
    Exit Sub
    
    'ESTA RUTINA ESTABA EN EL COMMAND23 DE VER Y REIMPRIMIR ORDENES DE PAGO FRM OPORPA07
    Call SETULTREG("INDIORPA", 0)
    Call COPYSTRU("INDIPRES", "INDIORPA")
    SQLX$ = " SELECT NuProve , SUM(ValAbsComp) AS SUVALOP, Codicom_Lar , FechEmisi , RefeCom, Niv_Clasi FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + " AND CODICOM_LAR LIKE 'OP%' "
    SQLX$ = SQLX$ + "GROUP BY NUPROVE,CODICOM_LAR, FECHEMISI, REFECOM, NIV_CLASI "
    SQLX$ = SQLX$ + "Order by FECHEMISI ASC, CODICOM_LAR ASC "
    '
    Set CABATEMP = READSET(SQLX$)
    JJ& = 0
    NIVCONXXX% = NIVCONFI%
    With CABATEMP
       Do While Not .EOF
          Call TRACE(20, JJ&, FIN&)
          If XVALO(!Niv_Clasi) <= NIVCONFI% Then    '\\\OT-08-1032-FL-12/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
             NCLIE% = !NUPROVE
             IMPO# = !SUVALOP                       ' ValAbsComp
             If IMPO# >= 0.005 Then
                NPX& = XVALO(Mid$(!CODICOM_LAR, 11, 8))
                FEX = CVINT(!FECHEMISI)
                YY$ = MKS$(NPX&) + TRIM$(!RefeCom) + " - " + FECHATEX$(FEX)
                JJ& = JJ& + 1
                Call GRAREG("!INDIORPA", YY$, JJ&)
             End If
          End If
70     .MoveNext
       Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    Call SETULTREG("!INDIORPA", JJ&)
    Call TRACENTRAL("INDIORPA", "")
    Screen.MousePointer = 1
    

End Sub




