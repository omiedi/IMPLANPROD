VERSION 5.00
Begin VB.Form REMEVE04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Gestión - Resúmenes Mensuales de Ventas"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture10 
      Height          =   5295
      Left            =   200
      ScaleHeight     =   5235
      ScaleWidth      =   9195
      TabIndex        =   60
      Top             =   1080
      Visible         =   0   'False
      Width           =   9260
      Begin VB.Image Image5 
         Height          =   5475
         Left            =   0
         Picture         =   "REMEVE04.frx":0000
         Stretch         =   -1  'True
         Top             =   -240
         Width           =   9735
      End
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00004000&
      Height          =   6735
      Left            =   9600
      ScaleHeight     =   6675
      ScaleWidth      =   900
      TabIndex        =   42
      Top             =   0
      Width           =   960
      Begin VB.PictureBox Picture16 
         Height          =   6400
         Left            =   -120
         ScaleHeight     =   6345
         ScaleWidth      =   1020
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image4 
            Height          =   6345
            Left            =   0
            Picture         =   "REMEVE04.frx":2E8A2
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
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
         Height          =   690
         Left            =   105
         Picture         =   "REMEVE04.frx":30C3A
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Archivos de Transferencia Percep. I.V.A. / Ingresos Brutos / Ret. Ganancias"
         Top             =   2790
         Width           =   650
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   55
         Top             =   4230
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Cust."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "REMEVE04.frx":30F44
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Acceso a Maestro de Clientes"
         Top             =   2100
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "REMEVE04.frx":31386
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   315
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
         Left            =   105
         Picture         =   "REMEVE04.frx":314D0
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   4725
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
         Height          =   630
         Left            =   105
         Picture         =   "REMEVE04.frx":31912
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   955
         Width           =   650
      End
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
         Height          =   630
         Left            =   105
         Picture         =   "REMEVE04.frx":31C1C
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3480
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   43
         Top             =   5460
         Width           =   645
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   44
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -315
         Picture         =   "REMEVE04.frx":31F26
         Top             =   5775
         Width           =   1515
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Dispositivo de Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   210
      TabIndex        =   30
      Top             =   5355
      Width           =   4635
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0C0&
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
         Left            =   840
         TabIndex        =   32
         Top             =   460
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0C0&
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
         TabIndex        =   31
         Top             =   460
         Width           =   1170
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Período Mensual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   210
      TabIndex        =   26
      Top             =   210
      Width           =   4680
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
         Left            =   3360
         TabIndex        =   28
         Top             =   410
         Width           =   175
      End
      Begin VB.TextBox MESAN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1575
         TabIndex        =   27
         Top             =   410
         Width           =   1800
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mes y año:"
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
         Index           =   1
         Left            =   -735
         TabIndex        =   29
         Top             =   450
         Width           =   2220
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   3885
         Picture         =   "REMEVE04.frx":33E48
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3885
         Picture         =   "REMEVE04.frx":34152
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Gestión de Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   210
      TabIndex        =   25
      Top             =   1260
      Width           =   4635
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1365
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   33
         Top             =   1470
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
            Height          =   600
            Left            =   0
            Picture         =   "REMEVE04.frx":3445C
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Facturación en Cuenta Corriente y de Contado"
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
            Left            =   735
            TabIndex        =   35
            Top             =   0
            Width           =   1515
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   36
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
            Height          =   600
            Left            =   -20
            Picture         =   "REMEVE04.frx":34766
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Notas de Débito"
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
            TabIndex        =   38
            Top             =   105
            Width           =   1410
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   39
         Top             =   420
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
            Height          =   600
            Left            =   0
            Picture         =   "REMEVE04.frx":34A70
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Notas de Crédito"
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
            Left            =   735
            TabIndex        =   41
            Top             =   105
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Libros de Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   210
      TabIndex        =   18
      Top             =   3780
      Width           =   4635
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2940
         TabIndex        =   19
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
            Height          =   600
            Left            =   0
            Picture         =   "REMEVE04.frx":34D7A
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Emite Subdiario de Ventas por Impresora"
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
            Height          =   600
            Left            =   1840
            Picture         =   "REMEVE04.frx":35084
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Emite Subdiario de I.V.A.- Ventas por Impresora"
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
            Picture         =   "REMEVE04.frx":3538E
            Style           =   1  'Graphical
            TabIndex        =   21
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
            Height          =   600
            Left            =   2400
            Picture         =   "REMEVE04.frx":35698
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Resumen Mensual I.V.A. Ventas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Emisión de Subdiarios"
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
            Top             =   105
            Width           =   1485
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Otros Reportes Mensuales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6030
      Left            =   5040
      TabIndex        =   0
      Top             =   210
      Width           =   4425
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0C0&
         Caption         =   "Estadísticas de Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1225
         Left            =   210
         TabIndex        =   50
         Top             =   4620
         Width           =   4005
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   230
            ScaleHeight     =   585
            ScaleWidth      =   3480
            TabIndex        =   51
            Top             =   420
            Width           =   3540
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
               Height          =   600
               Left            =   2940
               Picture         =   "REMEVE04.frx":359A2
               Style           =   1  'Graphical
               TabIndex        =   58
               ToolTipText     =   "Detalle de Facturación por Rango de Fechas"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command30 
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
               Left            =   1815
               Picture         =   "REMEVE04.frx":35CAC
               Style           =   1  'Graphical
               TabIndex        =   54
               ToolTipText     =   "Ranking de Ventas por Cliente por Rango de Fechas"
               Top             =   0
               Width           =   555
            End
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
               Left            =   2370
               Picture         =   "REMEVE04.frx":360EE
               Style           =   1  'Graphical
               TabIndex        =   56
               ToolTipText     =   "Ranking de Ventas por Articulo por Rango de Fechas"
               Top             =   0
               Width           =   560
            End
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
               Left            =   0
               Picture         =   "REMEVE04.frx":363F8
               Style           =   1  'Graphical
               TabIndex        =   52
               ToolTipText     =   "Resumen por Provincia y Grupo"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Resúmenes Estadísticos"
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
               Left            =   680
               TabIndex        =   53
               Top             =   90
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0E0C0&
         Caption         =   "Cobranza y Ajustes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   210
         TabIndex        =   3
         Top             =   315
         Width           =   4005
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2475
            TabIndex        =   4
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
               Left            =   -20
               Picture         =   "REMEVE04.frx":36702
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Detalle de Cobranza en Cuenta Corriente"
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
               TabIndex        =   6
               Top             =   105
               Width           =   1905
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2475
            TabIndex        =   7
            Top             =   420
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
               Height          =   600
               Left            =   0
               Picture         =   "REMEVE04.frx":36A0C
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Ajustes y Otros Movimientos"
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
               Index           =   0
               Left            =   735
               TabIndex        =   9
               Top             =   0
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0E0C0&
         Caption         =   "Procesos  Contables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1225
         Left            =   210
         TabIndex        =   2
         Top             =   3320
         Width           =   4005
         Begin VB.PictureBox Picture15 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3135
            TabIndex        =   13
            Top             =   420
            Width           =   3200
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
               Left            =   2020
               Picture         =   "REMEVE04.frx":36D16
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Consulta de Asiento Contable por Pantalla"
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
               Height          =   600
               Left            =   0
               Picture         =   "REMEVE04.frx":37020
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Consulta / Listado por Cuenta Contable"
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
               Height          =   600
               Left            =   2580
               Picture         =   "REMEVE04.frx":3732A
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Re-Imputación de Comprobantes de Ventas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Contabilización de Ventas"
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
               TabIndex        =   17
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0E0C0&
         Caption         =   "Resumen Operativo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1225
         Left            =   210
         TabIndex        =   1
         Top             =   2020
         Width           =   4005
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2055
            TabIndex        =   10
            Top             =   420
            Width           =   2115
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
               Picture         =   "REMEVE04.frx":37474
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Resumen Men- sual por Cuenta"
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
               TabIndex        =   12
               Top             =   105
               Width           =   1485
            End
         End
      End
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu edicion 
      Caption         =   "Edición"
      Begin VB.Menu reimputaciondecomprobantesdeventas 
         Caption         =   "Re-Imputación de Comprobantes de Ventas"
      End
      Begin VB.Menu reabrirmes 
         Caption         =   "Re-Abrir (Se Permitirá Realizar Cambios a Este Mes)"
      End
      Begin VB.Menu cerrarmes 
         Caption         =   "Cerrar Mes (No Podran Realizarse Cambios  a Este Mes)"
      End
      Begin VB.Menu ssss 
         Caption         =   "-"
      End
      Begin VB.Menu archivostransferenciaingresobrutoganancia 
         Caption         =   "Archivos de Transferencia Ingresos Brutos / C.I.T.I. Compras/Perc.Ganancias"
      End
   End
   Begin VB.Menu Consultasylistados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu notascredito 
         Caption         =   "Notas de Crédito (Del Mes Activo)"
         Begin VB.Menu pantalla0 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora0 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu notadebito 
         Caption         =   "Notas de Dédito (Del Mes Activo)"
         Begin VB.Menu pantalla1 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora1 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu facturasdeventas 
         Caption         =   "Facturas de Ventas  (Del Mes Activo)"
         Begin VB.Menu pantalla4 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora4 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu librosdeventas 
         Caption         =   "Libros de Ventas"
         Begin VB.Menu subdiariodeventas 
            Caption         =   "Subdiario de Ventas por Impresora"
         End
         Begin VB.Menu subiarioivaventasporimpresora 
            Caption         =   "Subdiario de I.V.A.- Ventass por Impresora"
         End
         Begin VB.Menu resumenmensualivaventas 
            Caption         =   "Resumen Mensual I.V.A. Ventas"
         End
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu resumenmensualporcuenta 
         Caption         =   "Resumen Mensual por Cuenta (Saldo de Clientes al Comenzar y al Terminar el mes)"
         Begin VB.Menu pantalla5 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora5 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu sss 
         Caption         =   "-"
      End
      Begin VB.Menu listadoporcuentacontable 
         Caption         =   "Consulta / Listado por Cuenta Contable"
         Begin VB.Menu pantalla6 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora6 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu consultadeasientocontableporpantalla 
         Caption         =   "Consulta de Asiento Contable por Pantalla"
      End
      Begin VB.Menu rankingporclienterporrangodefechas 
         Caption         =   "Ranking de Ventas por Cliente por Rango de Fechas"
         Begin VB.Menu pantalla7 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora7 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu rankingdeclientesporarticuloyfecha 
         Caption         =   "Ranking de Ventas por Articulo por Rango de Fechas"
         Begin VB.Menu pantalla8 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora8 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu resumenprovinciaygrupos 
         Caption         =   "Resumen de Ventas por Provincia y por Grupo"
         Visible         =   0   'False
      End
      Begin VB.Menu detalledeventas 
         Caption         =   "Detalle de Facturación por Rango de Fechas"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu recibosdecobranzayajustes 
      Caption         =   "Recibos de Cobranza y Ajustes"
      Begin VB.Menu ajustesyotrosmovimientos 
         Caption         =   "Ajustes y Otros Movimientos"
         Begin VB.Menu pantalla 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu detallederecibos 
         Caption         =   "Detalle de Recibos de Cobranza"
         Begin VB.Menu pantalla2 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora2 
            Caption         =   "Impresora"
         End
      End
   End
   Begin VB.Menu accesosdirectos 
      Caption         =   "Accesos Directos"
      Begin VB.Menu maestrocientes 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu varios 
         Caption         =   "Liquidación de Comisiones\Listado de Debitos y Creditos\Subdiario de Retenciones de Ingresos Brutos"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu configuracion 
      Caption         =   "Configuración"
      Begin VB.Menu config 
         Caption         =   "Configuración"
      End
   End
End
Attribute VB_Name = "REMEVE04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DESDE%, HASTA%, CORRESP$
Dim DESDAT As Double
Dim HASDAT As Double




Private Sub archivostransferenciaingresobrutoganancia_Click()
    Call Command24_Click
End Sub

Private Sub cerrarmes_Click()
   Call Image2_Click
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub


Private Sub Command10_Click()
    Command10.Enabled = False
    '
    TTXX$ = "Ventas por Provincia y Artículo\Ventas por Provincia y Cliente"
    OPMAX% = 2
    If EXISTE%("IGESVE04.EXE") > 0 Then
       TTXX$ = TTXX$ + "\Informe Estadístico de Gestión"
       OPMAX% = 3
    End If
    OPX% = ALTERNA%(TTXX$)
    If OPX% < 1 Or OPX% > OPMAX% Then GoTo 99
    If OPX% = 2 Then
         Call VENPROCLI
         GoTo 99
       ElseIf OPX% = 3 Then
         Call CONECRUN("IGESVE04", "Informe Estadístico de Gestión")
         GoTo 99
    End If
    '
    Dim CPROVI$(), CIXX%(), TOVEN#(), GRUCO%()
    '
    difacu# = 0
    MODOTANIT% = 0
    If EXISTE%("LIVENTAN.PRF") > 0 Then MODOTANIT% = 1
    '
'1- Ingresar rango de fechas desde - hasta
10   Call DESHASFECHA(DES%, HAS%, PERIO$)
     If DES% < 1 Then GoTo 99
     DESX = Int(CDbl(CVDAT(DES%)))
     HASX = Int(CDbl(CVDAT(HAS%)))
     '\\\OT-08-0591-od-31/07/08///
     OPX% = COMALTER%("Opciones de Formato de Salida\\Todos\Nacionales\Extranjeros")
     If OPX% < 1 Or OPX% > 3 Then GoTo 10
     '\\\OT-08-0591-od-31/07/08///
     CAREGIS% = 0
     '
'1'- Copiar la estructura del archivo GRUCOVEN.rfs
    Call COPYSTRU("GRUCOVEN", "GRUCOVE1")
'1''-Copio los registros del GRUCOVEN.DAT al GRUCOVE1.DAT
    YY$ = REGBLAN$("GRUCOVE1")
    U& = 0
    For J& = 1 To ULTREG&("GRUCOVEN")
      x$ = REGLEIDO("GRUCOVEN", J&)
      U& = U& + 1
      Call GRAREG("GRUCOVE1", x$, U&)
    Next J&
'1'''-Agregar un ultimo registro a GRUCOVE1.DAT
    YY$ = REGBLAN$("GRUCOVE1")
    Call REPLA(YY$, MKI(99), 1, 2)
    Call REPLA(YY$, "No Clasificado", 3, 30)
    Call GRAREG("GRUCOVE1", YY$, U& + 1)
    Call SETULTREG("GRUCOVE1", U& + 1)
    '
    If MODOTANIT% = 1 Then
       DESDAT = DESX
       HASDAT = HASX
       Call GENIVAVEN
    End If
    '
'2- Seleccionar de CajaBanc las facturas del período - similar a libro IVA
'    con el cliente obtener la provincia
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where INT(CDBL(FechEmisi)) >= " & DESX & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechEmisi)) <= " & HASX & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'VENTA' or TipoMovim = 'NDCLI') "
    SQLX$ = SQLX$ + "Order by FechEmisi ASC, NumeComp ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    SQLY$ = "SELECT * FROM CAJABANC "
    SQLY$ = SQLY$ + "where INT(CDBL(FechEmisi)) >= " & DESX & " "
    SQLY$ = SQLY$ + "AND INT(CDBL(FechEmisi)) <= " & HASX & " "
    SQLY$ = SQLY$ + "AND TipoMovim = 'EMBAL' "
    Set SELEMBA = CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot)
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      FIN& = .RecordCount
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      U& = 0
      .MoveFirst
      Do While Not .EOF
        U& = U& + 1
        Call TRACE(20, U&, FIN&)
        Screen.MousePointer = 11
        If TRIM$(UCase$(!OPCIMOVIM)) <> "ANUL" Then
          NUMEFA$ = !CODICOM_LAR
          NCLIE% = !NuClien
          '\\\OT-08-0591-od-31/07/08///
          If OPX% = 2 And PIVACLI%(NCLIE%) = 5 Then GoTo 40
          If OPX% = 3 And PIVACLI%(NCLIE%) <> 5 Then GoTo 40
          '\\\OT-08-0591-od-31/07/08///
          ITOTA# = !IHABECOMP - !IDEBECOMP
          SIG% = 1: If ITOTA# < 0 Then SIG% = (-1)
          '
          ' AQUI OBTENGO EL IMPORTE EMBALAJE
          IEMBAL# = 0
          SELEMBA.FindFirst "CodiCom_Lar = '" & NUMEFA$ & "'"
          If SELEMBA.NoMatch = False Then
            IEMBAL# = SELEMBA!IHABECOMP - SELEMBA!IDEBECOMP
          End If
          '
          ' LO QUE SIQUE ES PARA OBTENER EL NETO GRAVADO DE DÉBITOS LIBRES - TANIT
          If Left$(UCase$(NUMEFA$), 2) = "ND" Then
            If MODOTANIT% = 1 Then
              RFVT$ = FILTERGETRECORD$("LIVENTAN", 5, NUMEFA$)
              If TRIM$(Mid$(RFVT$, 11, 18)) = TRIM$(NUMEFA$) Then
                ITOTA# = SIG% * Abs(CVD(Mid$(RFVT$, 137, 8)))
              End If
            End If
          End If
          '
          MOBASI$ = !TipoMovim
          COPROVIN$ = PROVCLI$(NCLIE%)
          If TRIM$(COPROVIN$) = "" Or ECOARCH$("PROVINC", COPROVIN$) = "" Then
            Call MENSERR(24, "Falta Codigo de Provincia\para Cliente Numero " + TRIM$(Str$(NCLIE%)) + ".")
          End If
          SUMAFA# = 0
          '
' COEFICIENTE DE PRECIOS
          Call APERBASDAT("STOCKS")
          SQLY$ = "SELECT * FROM MOVISTO "
          SQLY$ = SQLY$ + "where DoSecLar = '" & NUMEFA$ & "' "
          SQLY$ = SQLY$ + "AND (Nume_Clie = " & NCLIE% & ""
             ORGD% = ORIGRUDI%(NCLIE%)
             If ORGD% > 0 Then
               If ORGD% <> NCLIE% Then
                  SQLY$ = SQLY$ + " OR Nume_Clie = " & ORGD% & ""
               End If
             End If
          SQLY$ = SQLY$ + ") "
          SQLY$ = SQLY$ + "Order by NuorItem ASC "
          '
          Set MOVISTEMP = Stocks.OpenRecordset(SQLY$, dbOpenSnapshot)
          On Error Resume Next
          MOVISTEMP.MoveFirst
          If Err > 0 Then
            On Error GoTo 0
            GoTo 35 'Stop ' PROXIMA FACTURA
          End If
          On Error GoTo 0
          SUMATVEN# = 0
          Do While Not MOVISTEMP.EOF
            CIXI% = MOVISTEMP!Cod_Inte
            'CANTU = (Abs(MOVISTEMP!CantSalid - MOVISTEMP!CantIngre)) * SIG%
            CANTU = ((MOVISTEMP!CANTSALID - MOVISTEMP!CANTINGRE))
            PREUN# = MOVISTEMP!VALOUNIT
            SUMATVEN# = SUMATVEN# + PREUN# * CANTU
            MOVISTEMP.MoveNext
          Loop
          COEFIMA# = 1
          If Abs(SUMATVEN#) > 0.005 Then
            COEFIMA# = Abs(ITOTA# / SUMATVEN#)
          End If
          '
'3.- para cada factura, con codicom_cor seleccionar los movimientos de stock correspondientes y obtener CI%, Preuni, cantidad y con eso el importe.
25        Call APERBASDAT("STOCKS")
          SQLY$ = "SELECT * FROM MOVISTO "
          SQLY$ = SQLY$ + "where DoSecLar = '" & NUMEFA$ & "' "
          SQLY$ = SQLY$ + "AND (Nume_Clie = " & NCLIE% & ""
             ORGD% = ORIGRUDI%(NCLIE%)
             If ORGD% > 0 Then
               If ORGD% <> NCLIE% Then
                  SQLY$ = SQLY$ + " OR Nume_Clie = " & ORGD% & ""
               End If
             End If
          SQLY$ = SQLY$ + ") "
          SQLY$ = SQLY$ + "Order by NuorItem ASC "
          '
          Set MOVISTEMP = Stocks.OpenRecordset(SQLY$, dbOpenSnapshot)
          On Error Resume Next
          MOVISTEMP.MoveFirst
          If Err > 0 Then
            On Error GoTo 0
            GoTo 35 'Stop ' PROXIMA FACTURA
          End If
          On Error GoTo 0
          TDC$ = UCase$(Left$(MOVISTEMP!DoSecLar, 2))
          Do While Not MOVISTEMP.EOF
            CIXI% = MOVISTEMP!Cod_Inte
            'CANTU = (Abs(MOVISTEMP!CantSalid - MOVISTEMP!CantIngre)) * SIG%
            CANTU = ((MOVISTEMP!CANTSALID - MOVISTEMP!CANTINGRE))
            PREUN# = MOVISTEMP!VALOUNIT * COEFIMA#
            GRUMA% = GRUCOVEI%(CIXI%)
            If GRUMA% < 1 Then GRUMA% = 99
            SUMAFA# = SUMAFA# + PREUN# * CANTU
            '
            ' aqui eliminar notas de debito y credito de texto libre
'            If TDC$ = "NC" Or TDC$ = "ND" Then
'              If CIXI% < 1 Then
'                GoTo 32
'              End If
'            End If
'            '
'4.- Buscar en el vector de detalle coincidencia de provincia y articulo. si machea sumas cantidad e importe
'    Si no machea agregar un elemento al vector
            For KU% = CAREGIS% To 1 Step -1
              If CPROVI$(KU%) = COPROVIN$ Then
                If CIXX%(KU%) = CIXI% Then
                  TOVEN#(KU%) = TOVEN#(KU%) + PREUN# * CANTU
                  GoTo 32
                End If
              End If
            Next KU%
            CAREGIS% = CAREGIS% + 1
            ReDim Preserve CPROVI$(CAREGIS%), CIXX%(CAREGIS%), TOVEN#(CAREGIS%), GRUCO%(CAREGIS%)
            CPROVI$(CAREGIS%) = COPROVIN$
            CIXX%(CAREGIS%) = CIXI%
            TOVEN#(CAREGIS%) = PREUN# * CANTU
            GRUCO%(CAREGIS%) = GRUMA%
            '
'5.- repetir para los restantes items de la factura
32        MOVISTEMP.MoveNext
          Loop
35        DIFIMP# = IEMBAL# + ITOTA# - SUMAFA#
' MsgBox NUMEFA$ & "   " & ITOTA# & "   " & DIFIMP#
          If Abs(DIFIMP#) >= 0.005 Then
            CIXI% = 0
            PREUN# = DIFIMP#
            CANTU = 1
            GRUMA% = 99
            For KU% = CAREGIS% To 1 Step -1
              If CPROVI$(KU%) = COPROVIN$ Then
                If CIXX%(KU%) = CIXI% Then
                  TOVEN#(KU%) = TOVEN#(KU%) + (Int(100 * PREUN# * CANTU + 0.5)) / 100
                  GoTo 40
                End If
              End If
            Next KU%
            CAREGIS% = CAREGIS% + 1
            ReDim Preserve CPROVI$(CAREGIS%), CIXX%(CAREGIS%), TOVEN#(CAREGIS%), GRUCO%(CAREGIS%)
            CPROVI$(CAREGIS%) = COPROVIN$
            CIXX%(CAREGIS%) = CIXI%
            TOVEN#(CAREGIS%) = (Int(100 * PREUN# * CANTU + 0.5)) / 100
            GRUCO%(CAREGIS%) = GRUMA%
          End If
          '
        End If
'6.- siguiente factura
40    .MoveNext
      Loop
    End With
    '
'7.- ordenar creciente por provincia y por tipo de material
'8.- generar listador
    '
    U& = 0
    YYY$ = REGBLAN("VEPROGRU")
    For J& = 1 To CAREGIS%
      Call REPLA(YYY$, CPROVI$(J&), 1, 1)
      Call REPLA(YYY$, MKI$(GRUCO%(J&)), 3, 2)
      Call REPLA(YYY$, MKI(CIXX%(J&)), 5, 2)
      Call REPLA(YYY$, MKD(TOVEN#(J&)), 7, 8)
      Call REPLA(YYY$, MKD(TOVEN#(J&)), 15, 8)
      U& = U& + 1
      Call GRAREG("VEPROGRU", YYY$, U&)
    Next J&
    '
    Call SETULTREG("VEPROGRU", U&)
    Call FILESORT("VEPROGRU", "VEPROGRU", 1, 5, "ASC")
    '
    REBLA$ = REGBLAN$("VEPROGRU")
    SUMAGRU# = 0: U2& = 0: XANT$ = REBLA$
    FIN& = ULTREG&("VEPROGRU")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      x$ = REGLEIDO$("VEPROGRU", U&)
      Y$ = REBLA$
      If U& < FIN& Then
        Y$ = REGLEIDO$("VEPROGRU", U& + 1)
      End If
      '
      SUMAGRU# = SUMAGRU# + CVD(Mid$(x$, 7, 8))
      CORTECON% = 0
      '
      Call REPLA(x$, String$(16, 0), 23, 16)
      If Left$(Y$, 1) <> Left$(x$, 1) Then     'Or CVI(Mid$(Y$, 3, 2)) <> CVI(Mid$(X$, 3, 2)) Then ' CAMBIO PROVINCIA O GRUPO
          Call REPLA(x$, MKD$(SUMAGRU#), 23, 8)
          SUMAGRU# = 0
          UGRU% = CVI(Mid$(x$, 3, 2))
          CORTECON% = 1
      End If
      U2& = U2& + 1
      Call GRAREG("!VEPROGRU", x$, U2&)
    Next U&
    Call SETULTREG("!VEPROGRU", U2&)
    Call CIERRARCH("!VEPROGRU")
    Call TRACENTRAL("VEPROGRU", "")
    '\\\OT-08-0591-OD-31/07/08///
    '
    Select Case OPX%
    Case Is = 1
      OpcionElegida$ = "Todos "
    Case Is = 2
      OpcionElegida$ = "Nacionales"
    Case Is = 3
      OpcionElegida$ = "Extranjeros"
    End Select
    '
    '\\\OT-08-0591-OD-31/07/08///
    Call FILESORT("VEPROGRU", "", 1, 3)
    Call HEADERS("VEPROGRU", "")
    Call HEADERS("VEPROGRU", "Corresponde a: " + PERIO$)
    '\\\OT-08-0591-OD-31/07/08///
    Call HEADERS("VEPROGRU", "Filtro: " + OpcionElegida$)
    '\\\OT-08-0591-OD-31/07/08///

    
    '
    Screen.MousePointer = 1
    Call FINAL("*VEPROGRU")
    '
99  Command10.Enabled = True

End Sub

Sub VENPROCLI()
    Dim TOTVENTA#(32767)
    '
10  Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    '\\\OT-08-0591-OD-31/07/08///
    OPX% = COMALTER%("Opciones de Formato de Salida\\Todos\Nacionales\Extranjeros")
    If OPX% < 1 Or OPX% > 3 Then GoTo 10

    '\\\OT-08-0591-OD-FIN///

    '
    Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTVENTA#(U&) = 0
    Next U&
    '
    DESDAT = Int(CDbl(CVDAT(DES%)))
    HASDAT = Int(CDbl(CVDAT(HAS%)))
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where INT(CDBL(FechEmisi)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechEmisi)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'VENTA' or TipoMovim = 'NDCLI') "
    SQLX$ = SQLX$ + "Order by FechEmisi ASC, NumeComp ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    SQLY$ = "SELECT * FROM CAJABANC "
    SQLY$ = SQLY$ + "where INT(CDBL(FechEmisi)) >= " & DESDAT & " "
    SQLY$ = SQLY$ + "AND INT(CDBL(FechEmisi)) <= " & HASDAT & " "
    SQLY$ = SQLY$ + "AND TipoMovim = 'EMBAL' "
    Set SELEMBA = CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot)
    '
    Screen.MousePointer = 11
    CLIMAX% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      FIN& = .RecordCount
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      JJ& = 0
      Do While Not .EOF
         JJ& = JJ& + 1
         Call TRACE(20, JJ&, FIN&)
         NCLIE% = !NuClien
         '\\\OT-08-0591-od-31/07/08///
         If OPX% = 2 And PIVACLI%(NCLIE%) = 5 Then GoTo 35
         If OPX% = 3 And PIVACLI%(NCLIE%) <> 5 Then GoTo 35
         '\\\OT-08-0591-od-31/07/08///
         If NCLIE% > 0 Then
           NUMEFA$ = !CODICOM_LAR
           If NCLIE% > CLIMAX% Then CLIMAX% = NCLIE%
           '
           TOTVENTA#(NCLIE%) = TOTVENTA#(NCLIE%) + (!IHABECOMP - !IDEBECOMP)
           ' agrega EL IMPORTE EMBALAJE
           IEMBAL# = 0
           SELEMBA.FindFirst "CodiCom_Lar = '" & NUMEFA$ & "'"
           If SELEMBA.NoMatch = False Then
             IEMBAL# = SELEMBA!IHABECOMP - SELEMBA!IDEBECOMP
           End If
           TOTVENTA#(NCLIE%) = TOTVENTA#(NCLIE%) + IEMBAL#
          '
         End If
35    .MoveNext
      Loop
    End With
    '
    JJ& = 0
    REBLA$ = REGBLAN$("VEPROVCL")
    For NCI% = 1 To 32766
       If Abs(TOTVENTA#(NCI%)) > 0.005 Then
          NCLIE% = NCI%
          COPROVIN$ = PROVCLI$(NCLIE%)
          XX$ = REBLA$
          Call REPLA(XX$, COPROVIN$, 1, 2)
          Call REPLA(XX$, MKI$(NCLIE%), 3, 2)
          Call REPLA(XX$, MKD$(TOTVENTA#(NCI%)), 5, 8)
          JJ& = JJ& + 1
          Call GRAREG("VEPROVCL", XX, JJ&)
       End If
    Next NCI%
    Call SETULTREG("VEPROVCL", JJ&)
    '\\\OT-08-0591-OD-31/07/08///
    '
    Select Case OPX%
    Case Is = 1
      OpcionElegida$ = "Todos "
    Case Is = 2
      OpcionElegida$ = "Nacionales"
    Case Is = 3
      OpcionElegida$ = "Extranjeros"
    End Select
    '
    '\\\OT-08-0591-odOD
    Call FILESORT("VEPROVCL", "", 1, 3)
    Call HEADERS("VEPROVCL", "")
    Call HEADERS("VEPROVCL", "Corresponde a: " + PERIO$)
    '\\\OT-08-0591-OD-31/07/08///
    Call HEADERS("VEPROVCL", "Filtro: " + OpcionElegida$)
    '\\\OT-08-0591-OD-31/07/08///

    Call FINAL("*VEPROVCL")
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")

End Sub
'
Sub Command11_Click()
    Command11.Enabled = False
    Dim TOTVENTA#(32767)
    Dim TOTCANT#(32767)
    Dim TOTCOSTO#(32767)
10   Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    '\\\OT-06-0494-RG-20/11/06///
    OPT% = COMALTER%("Opciones de Formato de Salida\\Ranking de Ventas\Ranking de Margen Bruto")
    If OPT% = 0 Then GoTo 99
    Call SELECHO("GRUCOVEN")
    LINEA% = XVALO(VALACT1$("GRUCOVEN"))
    DescripLinea$ = TRIM$(VALACT2$("GRUCOVEN"))
    '
    '\\\OT-08-0581-OD-23/07/08///
    'Seleccion de Grupo de Clientes
    Call SELECHO("GRUCOCLI")
    NroGruCli% = XVALO(VALACT1$("GRUCOCLI"))
    DenomGruCli$ = TRIM$(VALACT2$("GRUCOCLI"))    '
    '\\\OT-08-0581-OD-fin///
    '\\\OT-06-0494-RG-FIN///
    
    Screen.MousePointer = 11
    Call ACOSREPVIE                  ' ACTUALIZA COSTOS DE REPOSICIÓN VIEJOS
    '
    OPXX% = COMALTER%("Opcion de Ordenamiento\\Por Importe\Por Unidades")
    If OPXX% < 1 Or OPXX% > 2 Then GoTo 99
    ' incluye o excluye material no codificados
    OPXX1% = COMALTER%("Opciones de Listado\\Incluye Todos\Excluye Sin Código")
    If OPXX1% < 1 Or OPXX1% > 2 Then GoTo 99

    J& = 0: TOTAGEN# = 0: TOTACAN# = 0
    For U& = 0 To 32767
      TOTVENTA#(U&) = 0
      TOTCANT#(U&) = 0
      TOTCOSTO#(U&) = 0
    Next U&
    '
    DESDAT = Int(CDbl(CVDAT(DES%)))
    HASDAT = Int(CDbl(CVDAT(HAS%)))
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'VENTA' "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    SQLY$ = "SELECT * FROM CAJABANC "
    SQLY$ = SQLY$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLY$ = SQLY$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLY$ = SQLY$ + "AND TipoMovim = 'EMBAL' "
    Set SELEMBA = CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot)
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      '
      Call APERBASDAT("STOCKS")
      Screen.MousePointer = 11
      ARTMAX% = 0
      KU& = 0
      '
      Do While Not .EOF
        KU& = KU& + 1
        Call TRACE(20, KU&, FIN&)
        TOTFAC# = ROUND(!IHABECOMP - !IDEBECOMP, 2)
        'NCLIE% = !NuClien
        NCLIE% = !Cli_Real
        '\\\OT-08-0581-OD-23/07/08///
        If NroGruCli% = GRUCOCLI%(NCLIE%) Or NroGruCli% = 0 Then
        '\\\OT-08-0581-OD-fin///
          If Abs(TOTFAC#) >= 0.005 Then
            NROFAC$ = !CODICOM_LAR
            '
            ' AQUI OBTENGO EL IMPORTE EMBALAJE
            IEMBAL# = 0
            SELEMBA.FindFirst "CodiCom_Lar = '" & NROFAC$ & "'"
            If SELEMBA.NoMatch = False Then
              IEMBAL# = SELEMBA!IHABECOMP - SELEMBA!IDEBECOMP
            End If
            '
            SQLX$ = "SELECT * FROM MOVISTO "
            SQLX$ = SQLX$ + "WHERE DoSecLar = " & "'" & NROFAC$ & "'"
            SQLX$ = SQLX$ + "AND (Nume_Clie = " & NCLIE% & " "
               ORGD% = ORIGRUDI%(NCLIE%)
               If ORGD% > 0 Then
                 If ORGD% <> NCLIE% Then
                    SQLX$ = SQLX$ + " OR Nume_Clie = " & ORGD% & ""
                 End If
               End If
            SQLX$ = SQLX$ + ") "
            SQLX$ = SQLX$ + "Order by NuorItem ASC "
            '
            Set STOTEMP1 = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
            '
            With STOTEMP1
              On Error Resume Next
              .MoveFirst
              If Err Then
                 If OPXX1% = 1 Then TOTVENTA#(0) = TOTVENTA#(0) + TOTFAC# + IEMBAL# ' es para incluir solo si selecciono todos
                 On Error GoTo 0
                 GoTo 30
              End If
              On Error GoTo 0 '
              '
              'COEFICIENTE DE PRECIOS
              SUMATVEN# = 0
              Do While Not .EOF
                CIXI% = !Cod_Inte
                CANTU = (!CANTSALID - !CANTINGRE)
                PREUN# = !VALOUNIT
                If !MoneVal > 1 Then
                  PREUN# = PREUN# * !TIPOCAM
                End If
                SUMATVEN# = SUMATVEN# + PREUN# * CANTU
              .MoveNext
              Loop
              COEFIMA# = 1
              If Abs(SUMATVEN#) > 0.005 Then
                COEFIMA# = Abs(TOTFAC# / SUMATVEN#)
              End If
              '
              .MoveFirst
              Do While Not .EOF
                NART% = !Cod_Inte: If NART% < 0 Then NART% = 0
                If NART% > ARTMAX% Then ARTMAX% = NART%
                TIPOCOM$ = Mid$(NROFAC$, 1, 2)
                If NART% = 0 And OPXX1% = 2 Then GoTo 30
                '  CONTINUA SOLO SI EL ARTICULO PERTENECE AL GRUPO O GRUPO ES CERO
                If GRUCOVEI%(NART%) = LINEA% Or LINEA% = 0 Then
                    '
                    CANTI# = (!CANTSALID - !CANTINGRE)
                    TOTACAN# = TOTACAN# + CANTI#
                    NETO# = (!CANTSALID - !CANTINGRE) * !VALOUNIT * COEFIMA#
                       If !MoneVal > 1 Then NETO# = NETO# * !TIPOCAM
                       NETO# = ROUND(NETO#, 2)
                    COSTONETO# = ROUND((!CANTSALID - !CANTINGRE) * !CostoRep, 2)
                    '
                    TOTCANT#(NART%) = TOTCANT#(NART%) + CANTI#
                    TOTVENTA#(NART%) = TOTVENTA#(NART%) + NETO#
                    TOTCOSTO#(NART%) = TOTCOSTO#(NART%) + COSTONETO#
                    TOTFAC# = TOTFAC# - NETO#
                End If
                .MoveNext
              Loop
              '
              ' NO INCLUYE MATERIAL NO CODIFICADO CUANDO SE ELIGE UN GRUPO
              If LINEA% = 0 Then TOTVENTA#(0) = TOTVENTA#(0) + TOTFAC# + IEMBAL#
              '
            End With                    ' SOBRE MOVISTO
          End If
        End If
        '
30    .MoveNext
      Loop
      '
    End With                          ' SOBLE CAJABANC
    '
    For i% = 0 To ARTMAX%
      TOTAGEN# = TOTAGEN# + TOTVENTA#(i%)
    Next i%
    '
    List1.Clear
    For NCI% = 0 To 32766
     If NCI% > 0 Or LINEA% = 0 Then
       If Abs(TOTVENTA#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTVENTA#(NCI%)), 4, 12, 2, 2)
          If OPXX% = 2 Then RELIB$ = "A" + CSTRING$(MKD$(TOTACAN# - TOTCANT#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTVENTA#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTVENTA#(NCI%)), 4, 12, 2, 2)
          List1.AddItem RELIB$
       End If
     End If
    Next NCI%
    '
    J& = 0: PORACU = 0
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
        W$ = REGBLAN$("RANMABRU")
        
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call REPLA(Z$, MKD$(TOTCANT#(NC%)), 21, 8) '10/05/07 (OT-07-0188)
        Call GRAREG("RANKPIE", Z$, J&)
        '\\\OT-06-0494-RG-20/11/06///
        Call REPLA(W$, MKI$(JI%), 1, 2)
        Call REPLA(W$, MKI$(NC%), 3, 2)
        Call REPLA(W$, MKD$(TOTCANT#(NC%)), 5, 8)
        Call REPLA(W$, MKD$(TONEVEN#), 13, 8)
        '//OT-07-0458//
        'CU# = COSPESOS(NC%)
        CU# = 0
        If TOTCANT#(NC%) > 0.005 Then CU# = TOTCOSTO#(NC%) / TOTCANT#(NC%)
        Call REPLA(W$, MKD$(CU#), 21, 8)
        CT# = TOTCOSTO#(NC%)
        '//OT-07-0458//
        Call REPLA(W$, MKD$(CT#), 29, 8)
        UB# = TONEVEN# - CT#
        Call REPLA(W$, MKD$(UB#), 37, 8)
        Call REPLA(W$, MKS$(PORPAR), 45, 4)
        Call REPLA(W$, MKS$(PORACU), 49, 4)
        '\\\OT-07-0014-RG-24/1/07///
        IMPIVA# = (TONEVEN# * PORCEIVA(NC%)) / 100
        Call REPLA(W$, MKD$(IMPIVA#), 53, 8)
        ITOVE# = TONEVEN# + IMPIVA#
        Call REPLA(W$, MKD$(ITOVE#), 61, 8)
        '\\\OT-07-0014-RG-FIN///
        Call GRAREG("RANMABRU", W$, J&)
        '\\\OT-06-0494-RG-FIN///
    Next U&
    '
    
    Call SETULTREG("RANKPIE", J&)
    Call CIERRARCH("RANKPIE")
    Call HEADERS("RANKPIE", "")
    Call HEADERS("RANKPIE", "Periodo: " + PERIO$)
    If NroGruCli% > 0 Then
      Call HEADERS("RANKPIE", "Grupo Clientes : " + DenomGruCli$)
    End If
    If LINEA% > 0 Then
      Call HEADERS("RANKPIE", "Linea: " + DescripLinea$)
    End If
    '\\\OT-06-0494-RG-20/11/06///
    Call SETULTREG("RANMABRU", J&)
    Call CIERRARCH("RANMABRU")
    Call HEADERS("RANMABRU", "")
    Call HEADERS("RANMABRU", "Periodo: " + PERIO$)
    MOVIM$ = "Todos"
    If OPXX1% = 2 Then MOVIM$ = "Excluye Material No Codificado"
    Call HEADERS("RANMABRU", "Movimientos: " + MOVIM$)
    '\\\OT-06-0494-RG-FIN///
    Screen.MousePointer = 1
    List1.Clear
    '
    
    If OPT% = 1 Then
      If Option1.Value = True Then
         Call FINAL("*RANKPIE")
      Else
         Call FINAL("*LRANPIE")
      End If
      '\\\OT-06-0494-RG-20/11/06///
      ElseIf OPT% = 2 Then
          If Option1.Value = True Then
             Call FINAL("*RANMABRU")
          Else
             Call COMUNI("Solamente válido para impresión")
             Call FINAL("*RANMABRU")
          End If
      '\\\OT-06-0494-RG-FIN///
    End If
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    Command11.Enabled = True
    '
    '\\\OT-5-0384-WAR-07/06/05///
End Sub

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

Private Sub Command13_Click()
    Command13.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Option2.Value = True
    Call GENIVAVEN
    FIN& = ULTREG&("REIVAVEN")
    If FIN& < 1 Then GoTo 99
    '
    OPX% = COMALTER%("Filtrado de Salida:\\Totales\Local\Com.Exterior")
    If OPX% = 1 Then
        Call FINAL("*SOREIVEN")
      ElseIf OPX% = 2 Then
        JJ& = 0
        For U& = 1 To FIN&
          XXX$ = REGLEIDO$("REIVAVEN", U&)
          If InStr(XXX$, "COM.EXT") < 1 Then
            JJ& = JJ& + 1
            Call GRAREG("REIVAVEN", XXX$, JJ&)
          End If
        Next U&
        Call SETULTREG("REIVAVEN", JJ&)
        Call FINAL("*SOREIVEN")
      ElseIf OPX% = 3 Then
        JJ& = 0
        For U& = 1 To FIN&
          XXX$ = REGLEIDO$("REIVAVEN", U&)
          If InStr(XXX$, "COM.EXT") > 0 Then
            JJ& = JJ& + 1
            Call GRAREG("REIVAVEN", XXX$, JJ&)
          End If
        Next U&
        Call SETULTREG("REIVAVEN", JJ&)
        Call FINAL("*SOREIVEN")
    End If
    '
99  Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   '
   '\\\OT-06-0558-RG-21/12/06///
   OPT% = ALTERNA%("Imputación Contable\Libro de IVA")
   If OPT% = 1 Then
       Call COGEVE04.Command18_Click
     ElseIf OPT% = 2 Then
       FOLIVAVE04.Show 1
   End If
   '\\\OT-06-0558-RG-FIN///
   '
99 Command14.Enabled = True
End Sub

Sub Command15_Click()
    Command15.Enabled = False
    Dim TOTVENTA#(32767), TOTCANT#(32767)
    '
10  Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    '
15  OPX% = COMALTER%("Opciones de Trabajo:\\Por Cliente\Por Articulo\Por Fecha")
    If OPX% < 1 Or OPX% > 3 Then GoTo 99
    '
    If OPX% = 2 Then 'Si la opcion elegida es Por Articulo, permite filtrar por codigo de articulo
    VDEF$ = MKI$(PRICOD%) + MKI$(ULTCOD%) + Chr$(0)
20  RACOD$ = OBJPLA$("DESHASCOD", VDEF$)
      If RACOD$ = "" Then
        GoTo 15
      End If
      CI1% = CVI(Left$(RACOD$, 2)): COD1$ = CODEXT$(CI1%)
      CI2% = CVI(Mid$(RACOD$, 3, 2)): COD2$ = CODEXT$(CI2%)
      If CI1% = 0 Or CI2% = 0 Then GoTo 20
    End If
    
25  Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTVENTA#(U&) = 0
      TOTCANT#(U&) = 0
    Next U&
    '
    DESDAT = Int(CDbl(CVDAT(DES%)))
    HASDAT = Int(CDbl(CVDAT(HAS%)))
    REBLA$ = REGBLAN$("DETFACPE")
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'VENTA' "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      '
      Call APERBASDAT("STOCKS")
      Screen.MousePointer = 11
      ARTMAX% = 0
      KU& = 0
      '
      Do While Not .EOF
        KU& = KU& + 1
        Call TRACE(20, KU&, FIN&)
        TOTFAC# = ROUND(!IHABECOMP - !IDEBECOMP, 2)
        NCLIE% = !Cli_Real
        If NCLIE% = 0 Then NCLIE% = !NuClien
        If Abs(TOTFAC#) >= 0.005 Then
          FF% = CVINT(!FechConta)
          NROFAC$ = !CODICOM_LAR
          '
          If Left$(NROFAC$, 2) <> "FC" Then GoTo 30 ' solo facturas
          '
          SQLX$ = "SELECT * FROM MOVISTO "
          SQLX$ = SQLX$ + "WHERE DoSecLar = " & "'" & NROFAC$ & "'"
          SQLX$ = SQLX$ + "AND (Nume_Clie = " & NCLIE% & " "
' ESTE CODIGO SE COMENTA PORQUE A PARTIR DE ESTE MOMENTO SE UTILIZA
' EL CLIENTE REAL (CLI_REAL)
'             ORGD% = ORIGRUDI%(NCLIE%)
'             If ORGD% > 0 Then
'               If ORGD% <> NCLIE% Then
'                  SQLX$ = SQLX$ + " OR Nume_Clie = " & ORGD% & ""
'               End If
'             End If
          SQLX$ = SQLX$ + ") "
          SQLX$ = SQLX$ + "Order by NuorItem ASC "
          '
          Set STOTEMP1 = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
          '
          With STOTEMP1
            On Error Resume Next
            .MoveFirst
            If Err > 0 Then
               On Error GoTo 0
               NPED& = 0
               NART% = 0
               CANTU = 0
               PRUNI# = 0
               NETO# = TOTFAC#
               GoSub AGREDET
               GoTo 30
            End If
            On Error GoTo 0 '
            '
            'COEFICIENTE DE PRECIOS
            SUMATVEN# = 0
            Do While Not .EOF
              CIXI% = !Cod_Inte
              CANTU = (!CANTSALID - !CANTINGRE)
              PREUN# = !VALOUNIT
              If !MoneVal > 1 Then
                PREUN# = PREUN# * !TIPOCAM
              End If
              SUMATVEN# = SUMATVEN# + PREUN# * CANTU
            .MoveNext
            Loop
            COEFIMA# = 1
            If Abs(SUMATVEN#) > 0.005 Then
              COEFIMA# = Abs(TOTFAC# / SUMATVEN#)
            End If
            '
            .MoveFirst
            Do While Not .EOF
              NART% = !Cod_Inte: If NART% < 0 Then NART% = 0
              If NART% > ARTMAX% Then ARTMAX% = NART%
              NPED& = !NUPEDCLI
              CANTU = (!CANTSALID - !CANTINGRE)
              PRUNI# = !VALOUNIT * COEFIMA#
                If !MoneVal > 1 Then PRUNI# = PRUNI# * !TIPOCAM
              NETO# = ROUND(CANTU * PRUNI#)
              TOTFAC# = TOTFAC# - NETO#
              GoSub AGREDET
            .MoveNext
            Loop
            If Abs(TOTFAC#) >= 0.005 Then
               NPED& = 0
               NART% = 0
               CANTU = 0
               PRUNI# = 0
               NETO# = TOTFAC#
               GoSub AGREDET
            End If
          End With                    ' SOBRE MOVISTO
        End If
        '
30    .MoveNext
      Loop
      '
    End With                          ' SOBLE CAJABANC
    '
    Call SETULTREG("DETFACPE", J&)
    Call CIERRARCH("DETFACPE")
    Call HEADERS("DETFACPE", "")
    Call HEADERS("DETFACPE", "Periodo: " + PERIO$)
    If OPX% = 1 Then
        Call FILESORT("DETFACPE", "", 1, 4, "ASC")  'ORDENADO POR CLIENTE Y Nº FACTURA
    ElseIf OPX% = 3 Then
        Call FILESORT("DETFACPE", "", 3, 4, "ASC")  'ORDENADO POR FECHA Y Nº FACTURA
    End If
    
    '
    Call SETULTREG("DETFACOD", i&)
    Call CIERRARCH("DETFACOD")
    Call HEADERS("DETFACOD", "")
    Call HEADERS("DETFACOD", "Periodo: " + PERIO$)
    Call HEADERS("DETFACOD", "Desde Codigo: " + CODEXT$(CI1%) + " - " + DESCRIT$(CI1%))
    Call HEADERS("DETFACOD", "Hasta Codigo: " + CODEXT$(CI2%) + " - " + DESCRIT$(CI2%))
    Call FILESORT("DETFACOD", "", 3, 3, "ASC")
    Call FILESORT("DETFACOD", "", 1, 1, "ASC")
    '
    Screen.MousePointer = 1
    '
    If OPX% = 1 Then
        Call FINAL("*DETFACPE")     'ORDENADO POR CLIENTE Y Nº FACTURA
    ElseIf OPX% = 2 Then
        Call FINAL("*DETFACOD")
    ElseIf OPX% = 3 Then
        Call FINAL("*DETFACPE")     'ORDENADO POR FECHA Y Nº FACTURA
    End If
    GoTo 99
    '
AGREDET:
    'DETALLE DE FACTURACION POR CLIENTE
    Z$ = REBLA$
    Call REPLA(Z$, MKI$(NCLIE%), 1, 2)
    Call REPLA(Z$, MKI$(FF%), 3, 2)
    Call REPLA(Z$, NROFAC$, 5, 18)
    Call REPLA(Z$, NROREM$, 23, 18)
    Call REPLA(Z$, MKS$(NPED&), 41, 4)
    Call REPLA(Z$, MKI$(NART%), 45, 2) '10/05/07 (OT-07-0188)
    Call REPLA(Z$, MKS$(CANTU), 47, 4)
    Call REPLA(Z$, MKD$(PRUNI#), 51, 8)
    Call REPLA(Z$, MKD$(NETO#), 59, 8)
    J& = J& + 1
    Call GRAREG("DETFACPE", Z$, J&)
    '
    'DETALLE DE FACTURACION POR ARTICULO
    If CODEXT$(NART%) >= COD1$ Then
        If CODEXT$(NART%) <= COD2$ Then
            XZ$ = REBLA$
            Call REPLA(XZ$, MKI$(NART%), 1, 2)
            Call REPLA(XZ$, MKI$(NCLIE%), 3, 2)
            Call REPLA(XZ$, MKS$(CANTU), 5, 4)
            Call REPLA(XZ$, MKD$(PRUNI#), 9, 8)
            Call REPLA(XZ$, MKD$(NETO#), 17, 8)
            i& = i& + 1
            Call GRAREG("DETFACOD", XZ$, i&)
        End If
    End If
    
Return

99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    Command15.Enabled = True
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Option2.Value = True
    Call GENSUDIVEN
    Call FINAL("*SUDIVEN")
    '
99  Command16.Enabled = True
End Sub

Sub Command19_Click()
    Command19.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Call GENSUDIVEN
    If COPYFILE%(LUDAT$ + "SUDIVEN.DAT", LUDAT$ + "COSUDIV.DAT") > 0 Then
      Call FILESORT("COSUDIV", "", 1, 3, "ASC")
      If Option1.Value = True Then
          Call CONECRUN("*COSUDIV", "Consulta por Cuenta Contable")
          GoTo 99
        ElseIf Option2.Value = True Then
          Call FILESORT("COSUDIV", "", 6, 6, "ASC")
          Call CONECRUN("*LISUDIV", "Listado de Ventas por Cuenta Contable")
          GoTo 99
      End If
    End If
99  Command19.Enabled = True
End Sub

Sub Command2_Click()
    Command2.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    MOX% = ALTERNA%("Secuencia por Recibo\Detalle por Cliente")
    If MOX% < 1 Or MOX% > 2 Then
      GoTo 99
    End If
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "Order by FechConta ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    J1& = 0: J2& = 0
    REBLA1$ = REGBLAN$("DETACCX")
    REBLA1$ = REGBLAN$("RECOBCLI")
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
        ZZ$ = REBLA1$
        NCLIE% = !NuClien
        FECHAD = Int(CDbl(!FechConta))
        INETO# = !ValAbsComp
        ITOTA# = !ValAbsComp
        '
        Call REPLA(ZZ$, Chr$(4), 1, 1)
        Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 3, 2)
        Call REPLA(ZZ$, MKS$(!NumeComp), 5, 4)
        Call REPLA(ZZ$, MKI$(!NuClien), 9, 2)
        Call REPLA(ZZ$, MKD$(INETO#), 11, 8)
        Call REPLA(ZZ$, MKD$(ITOTA#), 51, 8)
        J1& = J1& + 1
        Call GRAREG("DETACCX", ZZ$, J1&)
        '
        ZZ$ = REBLA2$
        Call REPLA(ZZ$, MKI$(!NuClien), 1, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 3, 2)
        Call REPLA(ZZ$, MKS$(!NumeComp), 5, 4)
        Call REPLA(ZZ$, MKD$(INETO#), 25, 8)
        J2& = J2& + 1
        Call GRAREG("RECOBCLI", ZZ$, J2&)
        '
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("DETACCX", J1&)
    Call SETULTREG("RECOBCLI", J2&)
    '
    Screen.MousePointer = 1
    On MOX% GoTo 203, 204
    GoTo 99
    '
203 Screen.MousePointer = 11
    Call HEADERS("DETACCX", "")
    Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(MESAN))
    Call HEADERS("DETACCX", "Rubro: Recibos de Cobranza")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODECCX")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDECCX")
    End If
    GoTo 99
    '
204 Call HEADERS("RECOBCLI", "")
    Call HEADERS("RECOBCLI", "Corresponde a: " + TRIM$(MESAN))
    Call HEADERS("RECOBCLI", "Rubro: Recibos de Cobranza")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*COCOBCL")
      ElseIf Option2.Value = True Then
         Call FINAL("*SOCOBCL")
    End If
    GoTo 99
    '
99 Command2.Enabled = True
   Screen.MousePointer = 1
End Sub
'

Sub Command20_Click()
   '
   Command20.Enabled = False
   If DERACCE%("ASIENCON", "Asientos Contables por Pantalla") < 2 Then
     GoTo 99
   End If
  
   If VALIMESAN% < 1 Then GoTo 99
   '
   Call GENSUDIVEN
   '
   JJ& = 0
   TTXXY$ = Chr$(1) + Chr$(1) + MKI$(HASTA%)
   FIN& = ULTREG&("ASIDIAR")
   '
   For IL& = 1 To FIN&
     Call TRACE(20, IL& - IL0&, FIN& - IL0&)
     x$ = REGLEIDO$("ASIDIAR", IL&)
     If Left$(x$, 4) = TTXXY$ Then
       CUENUI% = CVI(Mid$(x$, 5, 2))
       IDEB# = CVD(Mid$(x$, 39, 8))
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
   '\\\OT-05-0710-WAR-03/11/05///
   If ULTREG&("ASIDIAFI") > 0 Then
   '\\\OT-05-0710-WAR-FIN///
    VTB% = VENTABU%("ASIDIAFI/TITUPAN=Asiento Resumen Mensual de Ventas - " + CORRESP$)
   End If
   '
99 Command20.Enabled = True
   '
End Sub

Sub Command21_Click()
    Command21.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    RESP1% = ALTERNA%("Selección de Comprobantes a Listar.\Listado Libro de IVA - Ventas.")
    If RESP1% = 1 Then
        Call EXCLUIVA
      ElseIf RESP1% = 2 Then
        Option2.Value = True
        Call GENIVAVEN
        If EXISTE%("LIVENTAN.PRF") > 0 Then
             Call FINAL("*LIVENTAN")
           Else
             FORIPRE$ = TRIM$(CONTROL$("FORIPRE", "LIVAVEN"))
             If FORIPRE$ <> "" Then
               If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") > 0 Then
                 If COMALTER("Elija Formato de Salida:\\Planilla Plana\Formulario") = 2 Then
                   Call FORMIVAVE
                   GoTo 99
                 End If
               End If
             End If
             Call FINAL("*LIVENMEN")
        End If
    End If
    '
99  Command21.Enabled = True
End Sub

Sub FORMIVAVE()
    '
    FORIPRE$ = TRIM$(CONTROL$("FORIPRE", "LIVAVEN"))
    FEX = HOY(HOS$)
    FECHDOC$ = FECHATEX$(FEX)
    NUMEDOC$ = MESAN
    TIPODOC$ = "Libro IVA Ventas"
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
    CAPARDOC% = 7
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
    CODTABU1$(13) = "PERC-BRU"
    CACOLTAB1% = 13
    '
    DESTIDOC% = 0
    CAITAB1% = 0
    '
    SUMA1# = 0: SUMA2# = 0: SUMA3# = 0
    SUMA4# = 0: SUMA5# = 0: SUMA6# = 0: SUMA7# = 0
    URE& = ULTREG&("LIVAVEN")
    For YX& = 1 To URE&
       '
       YY$ = REGLEIDO$("LIVAVEN", YX&)
       FEMI = CVI(Mid$(YY$, 3, 2))
       NDOC$ = Mid$(YY$, 11, 18)
       TIPO$ = Left(NDOC$, 2) + Mid$(NDOC$, 18, 1)
       NUME$ = Mid$(NDOC$, 4, 13)
       RACLI$ = Mid$(YY$, 33, 30)
       CUITI$ = TRIM$(Mid$(YY$, 63, 16))
       PIVAI$ = TRIM$(Mid$(YY$, 79, 12))
       TOFA0# = CVD(Mid$(YY$, 91, 8))
       '
       NEGRA0# = 0: NEXEN0# = 0: NIVA21# = 0: NIVA105# = 0: RETIVA# = 0: RETGAN# = 0
       '
25     NETO1# = CVD(Mid$(YY$, 137, 8)) + CVD(Mid$(YY$, 185, 8))
       IVA1# = CVD(Mid$(YY$, 145, 8))
       TASA1 = CVD(Mid$(YY$, 153, 8))
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
       '
       If YX& < URE& Then
         YZ$ = REGLEIDO$("LIVAVEN", YX& + 1)
         If TRIM$(Mid$(YZ$, 11, 18)) = "" Then
           YX& = YX& + 1
           YY$ = YZ$
           GoTo 25
         End If
       End If
       '
       CAITAB1% = CAITAB1% + 1
       TETABU1$(1, CAITAB1%) = FECHATEX$(FEMI)
       TETABU1$(2, CAITAB1%) = TIPO$
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
    Call PRINTDOC("@" + FORIPRE$)
    '
End Sub
'
Sub Command22_Click()
    '
    Command22.Enabled = False
    '
    If DERACCE%("REIMPUVE", "Re-Imputación de Comprobantes de Ventas") < 2 Then
      GoTo 99
    End If
    '
    If VALIMESAN% < 1 Then GoTo 99
    DES% = DESDE%: DESDEX = CDbl(CVDAT(DES%))
    HAS% = HASTA%: HASTAX = CDbl(CVDAT(HAS%))
    '
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then GoTo 99
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' OR TIPOMOVIM = 'AJUD') "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
    SQLX$ = SQLX$ + "AND FechConta >= " & DESDEX & " "
    SQLX$ = SQLX$ + "AND FechConta <= " & HASTAX & " "
    SQLX$ = SQLX$ + "Order by CODICOM_LAR ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!CODICOM_LAR, 11, 8))
            Y$ = MKS$(JJ& + 1) + AJUSTI$(!CODICOM_LAR, 18) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUERE", Y$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("!SELCUERE", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no Tiene\Comprobantes de Venta\que Puedan Re-Imputarse.")
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
20  ATRI$ = DOCULA$ + " - $ " + TRIM$(Mid$(VALACT2$("!SELCUERE"), 19, 14)) + " - " + TRIM$(RASOCLI$(NC%))
    '
    Call BLANARCH("!REIMPUTA")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(FechConta)) >= " & DESDEX & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASTAX & " "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & DOCULA$ & "' "
    SQLX$ = SQLX$ + "AND NuClien = " & NC% & " "
    '
    Set CABASEL = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    SUIMPU# = 0
    '
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
        CUE% = !CueContaI
        'If CUE% > 0 Then
           IPA# = !IDEBECOMP - !IHABECOMP
           Y$ = REGBLAN$("!REIMPUTA")
           Call REPLA(Y$, CODCUE$(CUE%), 1, 12)
           Call REPLA(Y$, MKD$(IPA#), 13, 8)
           Call REGAPP("!REIMPUTA", Y$)
           SUIMPU# = SUIMPU# + IPA#
        'End If
      .MoveNext
      Loop
    End With
    Call CIERRARCH("!REIMPUTA")
    '
    If Abs(SUIMPU#) >= 0.005 Then
      Call MENSERR(24, "Imposible Editar Imputación Contable.\Comprobante No Balanceado.\  \Debe Anularse y Registrarse Nuevamente.")
      GoTo 10
    End If
    '
30  ATRI$ = "TITUPAN=Re-Imputación Contable: " + ATRI$
    VTB% = VENTABU%("REIMPUTA/" + ATRI$)
    If VTB% < 0 Then GoTo 10
    '
    SUMANE# = 0
    For U& = 1 To ULTREG&("!REIMPUTA")
      Y$ = REGLEIDO$("!REIMPUTA", U&)
      CUENUE% = CUEINT%(Left$(Y$, 12))
      IMPOX# = CVD(Mid$(Y$, 13, 8))
      SUMANE# = SUMANE# + IMPOX#
      '
      If CUENUE% < 1 Then
        Call MENSERR(24, "Numero de Cuenta no Valido\en Registro Numero " + TRIM$(Str$(U&)))
        GoTo 30
      End If
    Next U&
    If Abs(SUMANE#) >= 0.005 Then
      Call MENSERR(24, "Comprobante no Balanceado")
      GoTo 20
    End If
    '
    REIMPU% = 0
    For U& = 1 To ULTREG&("!REIMPUTA")
      Y$ = REGLEIDO$("!REIMPUTA", U&)
      CUENUE% = CUEINT%(Left$(Y$, 12))
      IMPOX# = CVD(Mid$(Y$, 13, 8))
      '
      With CABASEL
        .MoveFirst
        Do While Not .EOF
32        CUE% = !CueContaI
          If Abs(!IDEBECOMP - !IHABECOMP - IMPOX#) < 0.005 Then
            .Edit
            !CueContaI = CUENUE%
            .Update
            REIMPU% = 1
            GoTo 35
          End If
        .MoveNext
        Loop
      End With
      '
35  Next U&
    '
    If REIMPU% > 0 Then
        Call COMUNI("Re-Imputación Completada")
      Else
        Call COMUNI("No se Pudo Completar la Re-Imputación")
    End If
    '
    GoTo 16
    '
99  Command22.Enabled = True
    '
End Sub

Private Sub INGBRUT()
                '\\\OT-06-0158-WAR-31/05/06///
   'Command23.Enabled = False
   '
   'If VALIMESAN% < 1 Then GoTo 99
   '
   'Screen.MousePointer = 11
   'Call GENIVAVEN
   'Screen.MousePointer = 1
   If ULTREG&("SUPERIB") < 1 Then GoTo 99
   '
   OPX% = COMALTER%("Opciones de Trabajo:\\Subdiario I.Brutos\Archivo Transferencia")
   If OPX% = 1 Then
       Call HEADERS("SUPERIB", "")
       Call HEADERS("SUPERIB", "Corresponde a: " + MESAN)
       Call FINAL("*SUPERIB")
     ElseIf OPX% = 2 Then
       '\\\OT-06-0076-WAR-02/03/06///'Para Tanit
       'para filtrar dentro del mes y año elegido.
10     Call MESANO(MESAN, DESX1%, HASX1%)
       VDEF$ = MKI$(DESX1%) + MKI$(HASX1%)
       OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
       If OBX$ = "" Then GoTo 99
       DESX2% = CVI(Left$(OBX$, 2))
       HASX2% = CVI(Mid$(OBX$, 3, 2))
       If DESX2% > HASX2% Or DESX2% < DESX1% Or HASX2% > HASX1% Then
         Call MENSERR(24, "Rango de Fechas no Procesable")
         GoTo 10
       End If
       ''\\\OT-06-0076-WAR-FIN///
       ASALID$ = "C:\FLEXOFT\TRANIBRU.TXT"
       NX% = FILEOPEN%(ASALID$, 2, 128)
       CAREGI% = 0
       For U& = 1 To ULTREG&("SUPERIB")
         XX$ = REGLEIDO$("SUPERIB", U&)
           FEX = CVI(Left$(XX$, 2))
           If FEX < DESX2% Or FEX > HASX2% Then GoTo 29
         TTXX$ = Space$(71)
           CUITTX$ = TRIM$(Mid$(XX$, 51, 14))
         Call REPLA(TTXX$, CUITTX$, 1, 13)
           FEXU$ = FECHATEX$(FEX)
           FEXU$ = Left$(FEXU$, 6) + "20" + Mid$(FEXU$, 7, 2)
         Call REPLA(TTXX$, FEXU$, 14, 10)
           LECOX$ = "F"
           If Mid$(XX$, 3, 2) = "ND" Then LECOX$ = "D"
           If Mid$(XX$, 3, 2) = "NC" Then LECOX$ = "C"
         Call REPLA(TTXX$, LECOX$, 24, 1)
           LECO$ = Mid$(XX$, 20, 1)
         Call REPLA(TTXX$, LECO$, 25, 1)
           SUCUX$ = Mid$(XX$, 6, 4)
         Call REPLA(TTXX$, SUCUX$, 26, 4)
           NUCOX$ = Mid$(XX$, 11, 8)
         Call REPLA(TTXX$, NUCOX$, 30, 8)
           IMPOL# = CVD(Mid$(XX$, 79, 8))
           IMPOX$ = REPLACAR$(TRIM$(Format$(Abs(IMPOL#), "#########0.00")), ".", ",")
           While Len(IMPOX$) < 12: IMPOX$ = "0" + IMPOX$: Wend
           If LECOX$ = "C" Then IMPOX$ = "-" + Mid$(IMPOX$, 2)
         Call REPLA(TTXX$, IMPOX$, 38, 12)
           IMPOL# = CVD(Mid$(XX$, 87, 8))
           IMPOX$ = REPLACAR$(TRIM$(Format$(Abs(IMPOL#), "#########0.00")), ".", ",")
           While Len(IMPOX$) < 11: IMPOX$ = "0" + IMPOX$: Wend
           If LECOX$ = "C" Then IMPOX$ = "-" + Mid$(IMPOX$, 2)
         Call REPLA(TTXX$, IMPOX$, 50, 11)
         Call REPLA(TTXX$, FEXU$, 61, 10)
         Call REPLA(TTXX$, "A", 71, 1) ' SE AGREGA SEGUN NUEVA NORMATIVA DE ARBA SETIEMBRE 2011
         Print #NX%, TTXX$
         CAREGI% = CAREGI% + 1
29     Next U&
       Close #NX%
       ASALID1$ = REPLACAR$(ASALID$, "\", "/")
       Call COMUNI("Archivo de Transferencia '" + ASALID1$ + "'\Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.")
   End If
   '
99 'Command23.Enabled = True
   On Error GoTo 0
End Sub


Private Sub Command24_Click()
   '\\\OT-06-0158-WAR-31/05/06///
   Command24.Enabled = False
   If VALIMESAN% < 1 Then GoTo 99
   '
5  OPXR% = COMALTER%("Opciones de Trabajo: \\ Percepciones I.V.A./I.Brutos\Retenciones Ganancias\Regimen Informativo de Ventas")
   If OPXR% < 1 Or OPXR% > 3 Then GoTo 99
   '
   If OPXR% = 1 Then 'Percepciones IVA e Ingresos Brutos
       Screen.MousePointer = 11
       Call GENIVAVEN
       Screen.MousePointer = 1
       '
       OPX1% = COMALTER%("Opciones de Percepción: \\ I.V.A.\Ingresos Brutos")
       If OPX1% = 1 Then                       'PERCEPCIONES DE IVA
            If ULTREG&("LIVAVEN") < 1 Then GoTo 99
            '
10          Call MESANO(MESAN, DESX1%, HASX1%)
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
15          CUIEM1$ = InputBox$("Ingrese Nº de C.U.I.T.", "Agente de Percepción", VD1$)
            If TRIM$(CUIEM1$) = "" Then
               Call MENSERR(24, "C.U.I.T. no Válido")
               GoTo 5
            End If
            '
            VD1$ = CUIEM1$
            If VALICUIT%(CUIEM1$) <= 0 Then
             Call MENSERR(24, "C.U.I.T. no Válido")
             GoTo 15
            End If
            '
            ASALID$ = "C:\FLEXOFT\TRAPEIVA.TXT"
            NX% = FILEOPEN%(ASALID$, 2, 128)
            CAREGI% = 0
            For U& = 1 To ULTREG&("LIVAVEN")
               XX$ = REGLEIDO$("LIVAVEN", U&)
               FEX = CVI(Mid$(XX$, 3, 2))
               If FEX < DESX2% Or FEX > HASX2% Then GoTo 29
               If CVD(Mid$(XX$, 161, 8)) > 0 Then ' SI PERCEPCION I.V.A.
                  '
                  TTXX$ = Space$(60)
                  CODPER1$ = CONTROL$("COREPER", "COREPER")
                  If TRIM$(CODPER1$) = "" Then
                     Call MENSERR(24, "Código de Regimen de Percepción no Válido.\Se ha Cancelado la Exportación.")
                     GoTo 99
                  End If
                  Call REPLA(TTXX$, CUIEM1$, 4, 13)     'CUIT DEL AGENTE
                  Call REPLA(TTXX$, CODPER1$, 1, 3)     'CODIGO DE REGIMEN DE PERCEPCION
                  FEXU$ = FECHATEX$(FEX)                'FECHA DE LA PERCEPCION
                  FEXU$ = Left$(FEXU$, 6) + "20" + Mid$(FEXU$, 7, 2)
                  Call REPLA(TTXX$, FEXU$, 17, 10)
                  X11$ = Mid$(XX$, 11, 18)              'PREFIJO DE FACTURA
                  PREFI1$ = Mid$(X11$, 4, 4)
                  While Len(PREFI1$) < 8: PREFI1$ = "0" + PREFI1$: Wend
                  Call REPLA(TTXX$, PREFI1$, 27, 8)
                  X2 = CVS(Mid$(XX$, 7, 4))             'NUMERO DE FACTURA
                  NUMEFAC1$ = TRIM$(Str$(X2))
                  While Len(NUMEFAC1$) < 8: NUMEFAC1$ = "0" + NUMEFAC1$: Wend
                  Call REPLA(TTXX$, NUMEFAC1$, 35, 8)
                  IMPOL# = CVD(Mid$(XX$, 161, 8))       'MONTO DE LA PERCEPCION
                  IMPOX$ = REPLACAR$(TRIM$(Format$(Abs(IMPOL#), "###########0.00")), ".", ",")
                  While Len(IMPOX$) < 16: IMPOX$ = "0" + IMPOX$: Wend
                  Call REPLA(TTXX$, IMPOX$, 43, 16)
                  Print #NX%, TTXX$
                  CAREGI% = CAREGI% + 1
                  '
               End If
29          Next U&
            Close #NX%
            ASALID1$ = REPLACAR$(ASALID$, "\", "/")
            Call COMUNI("Archivo de Transferencia '" + ASALID1$ + "'\Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.")
          ElseIf OPX1% = 2 Then 'INGRESOS BRUTOS
            Call INGBRUT
       End If
       '
    ElseIf OPXR% = 2 Then ' Retenciones
       '
25     Call DESHASFECHA(DES%, HAS%, PERIO$)       ' SELECCION DE FECHAS
       If PERIO$ = "" Then GoTo 99
       DESDAT = Int(CDbl(CVDAT(DES%)))
       HASDAT = Int(CDbl(CVDAT(HAS%)))
       '
       CMC% = 0
       RFF$ = RECFILT$("MECOBRA", 6, "RGAN")
       If Len(RFF$) >= 4 Then
          REME& = CVS(Mid$(RFF$, 1, 4))
          If REME& > 0 Then
            CUECU$ = Mid$(REGLEIDO$("MECOBRA", REME&), 45, 12)
            CMC% = CUECOINT%(CUECU$)
          End If
       End If
       '
       If CMC% = 0 Then
           Call MENSERR(24, "Falta Establecer Medio de Cobranza de Ganancias")
           GoTo 99
       End If
       '
       'CONSULTA AL CAJABANC
       Call APERBASDAT("CABAN")
       SQLX$ = "SELECT * FROM CAJABANC"
       SQLX$ = SQLX$ + " WHERE TipoMovim = 'CHRECI' "
       SQLX$ = SQLX$ + " AND   CueContai= " & CMC% & ""
       SQLX$ = SQLX$ + " AND   CDBL(FechConta) >= " & DESDAT & ""
       SQLX$ = SQLX$ + " AND   CDBL(FechConta) <= " & HASDAT & ""
       Set GananTemp = CueCorri.OpenRecordset(SQLX$, 4)
       With GananTemp
          On Error Resume Next
          .MoveFirst
          If Err Then
             Call COMUNI("No existen Retenciones en el Rango de Fechas Elegido")
             On Error GoTo 0
             GoTo 25
          End If
          On Error GoTo 0
          '
          JX& = 0
          CAREG1% = 0
          Call SETULTREG("!PERETGAN", 0)
          Do While Not .EOF
             '
             FECHA = !FechEmisi
             COMPROBANTE$ = !CODICOM_LAR
             CERTIFICADO$ = Mid$(!Descrimov, 19, 12)
             CLIENTE% = !NuClien
             CUIT$ = CUITCLI$(CLIENTE%)
             IMPORTE = !IDEBECOMP
             '\\\OT-08-0916-OD-01/10/08///
             CODIRETEGAN$ = SAFETEXT$(!CodRetGan)
             If CODIRETEGAN$ = "" Then CODIRETEGAN$ = "19"
             '\\\OT-08-0916-OD-FIN//
             Call REPLA(TXXX$, MKI$(CVINT(FECHA)), 1, 2)
             Call REPLA(TXXX$, COMPROBANTE$, 3, 18)
             Call REPLA(TXXX$, CERTIFICADO$, 21, 12)
             Call REPLA(TXXX$, MKI$(CLIENTE%), 33, 2)
             Call REPLA(TXXX$, CUIT$, 35, 13)
             Call REPLA(TXXX$, MKS(IMPORTE), 48, 4)
             Call REPLA(TXXX$, CODIRETEGAN$, 52, 3)
             JX& = JX& + 1
             Call GRAREG("!PERETGAN", TXXX$, JX&)
          .MoveNext
          Loop
       End With

40     Call SETULTREG("!PERETGAN", JX&)
       Screen.MousePointer = 1
       '\\\OT-08-0916-OD-01/10/08///
       VTB% = VENTABU%("PERETGA1")
       If VTB% >= 0 Then
           For U& = 1 To ULTREG&("!PERETGAN")
              XX$ = REGLEIDO$("!PERETGAN", U&)
              Call GRAREG("PERETGAN", XX$, U&)
              '
              ABUS$ = Mid$(XX$, 3, 18)
              FEC% = CVI(Mid$(XX$, 1, 2))
              FECCHH = Int(CDbl(CVDAT(FEC%)))
              NuClien% = CVI(Mid$(XX$, 33, 2))
              XODRETGAN$ = TRIM$(Mid$(XX$, 52, 3))
              '
              'que coincida documento ,fecha y Cliente
              SQLX$ = "UPDATE CAJABANC SET CodRetGan = '" & XODRETGAN$ & "'"
              SQLX$ = SQLX$ + " WHERE CodiCom_Lar = '" & ABUS$ & "'"
              SQLX$ = SQLX$ + " AND INT(CDBL(FechEmisi)) = " & FECCHH & ""
              SQLX$ = SQLX$ + " AND NuClien = " & NuClien% & ""
   
              CueCorri.Execute (SQLX$)
           Next U&
           Call SETULTREG("PERETGAN", JX&)
         Else
           GoTo 25
       End If
       '\\\OT-08-0916-OD-FIN///
       OPPCC% = COMALTER%("Opciones de Trabajo:\\Subdiarios de Ganancias\Archivo de Transferencia")
       If OPPCC% < 1 Or OPPCC% > 2 Then GoTo 25
       '
       If OPPCC% = 1 Then
            Call FILESORT("PERETGAN", "", 2, 2, "ASC")
            Call FILESORT("PERETGAN", "", 1, 1, "ASC")
            '
            Call HEADERS("PERETGAN", "")
            Call HEADERS("PERETGAN", "Periodo: " & PERIO$)
            Call FINAL("*RETEGAN")
         Else
             Dim XXX As String * 52
            ASALID$ = LUCOM$ + "RETGAN.TXT"
            NX% = FILEOPEN%(ASALID$, 2, 128)
            '
            FIN& = ULTREG("PERETGAN")
            Screen.MousePointer = 11
            For i& = 1 To FIN&
                 XXX = ""
                 TTXX$ = REGLEIDO$("PERETGAN", i&)
                 CUIT$ = Mid$(TTXX$, 35, 13)
                 NuCuit$ = REPLACAR(CUIT$, "-", "")
               Call REPLA(XXX, "1", 1, 1)
               Call REPLA(XXX, NuCuit$, 2, 11)
                 FECH% = CVI(Mid$(TTXX$, 1, 2))
               Call REPLA(XXX, FECHATEXLAR(FECH%), 13, 10)
                 REGIMEN$ = Mid$(TTXX$, 52, 3)
               Call REPLA(XXX, REGIMEN$, 23, 3)
                 IMPO# = CVS(Mid$(TTXX$, 48, 4))
               Call REPLA(XXX, FORMATNUM(IMPO#, "F15.2"), 26, 15)
                 CERTIFICADO$ = AJUSTI$(TRIM$(Mid$(TTXX$, 21, 12)), 12)
               Call REPLA(XXX, CERTIFICADO$, 41, 12)
               Print #NX%, XXX
            Next i&
            Close #NX%
            ASALID11$ = REPLACAR(ASALID$, "\", "/") 'INVERTIMOS LAS BARRAS PARA QUE _
             NO LO PRESENTE EN DIFERENTES RENGLONES EL CALL COMUNI
            Screen.MousePointer = 1
            Call COMUNI("Archivo de Transferencia '" + ASALID11$ + "'\Generado con " + TRIM$(Str$(FIN&)) + " Registros.")
       End If
       '\\\OT-08-0589-OD-FIN///
   ElseIf OPXR% = 3 Then
     Call GENIVAVEN
     Call RegimenInformativoVentas
   End If
   
   
99 Command24.Enabled = True
   '\\\OT-06-0158-WAR-FIN///
End Sub

Sub Command3_Click()
   Command3.Enabled = False
   If VALIMESAN% < 1 Then GoTo 99
   '
   Screen.MousePointer = 11
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   'SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
   'SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
   SQLX$ = SQLX$ + "where FechConta >= " & DESDAT & " "
   SQLX$ = SQLX$ + "AND FechConta <= " & HASDAT & " "
   SQLX$ = SQLX$ + "AND TipoMovim = 'FVEN' "
   SQLX$ = SQLX$ + "AND OpciMovim = 'FC' "
   SQLX$ = SQLX$ + "Order by FechConta ASC, CodiCom_Lar ASC "
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where FechConta >= " & DESDAT & " "
   SQLX$ = SQLX$ + "AND FechConta <= " & HASDAT & " "
   Set CABASELE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   JJ& = 0: REBLA$ = REGBLAN$("DETACCX")
   With CABATEMP
     On Error Resume Next
     .MoveLast
     If Err > 0 Then
       On Error GoTo 0
       Call MENSERR(24, "Sin Movimientos en el Período Elegido")
       GoTo 99
     End If
     On Error GoTo 0
     FIN& = .RecordCount
     .MoveFirst
     KU& = 0
     Do While Not .EOF
       KU& = KU& + 1
       Call TRACE(20, KU&, FIN&)
       ZZ$ = REBLA$
       NUMEFA$ = !CODICOM_LAR
       NCLIE% = !NuClien
       FEFAXI% = CVINT(!FechConta)
       FECHAD = Int(CDbl(!FechConta))
       NUMECO = !NumeComp
       '
       INETO# = 0
       IMIVA# = 0
       IVRNI# = 0
       RETEN# = 0
       IOTRO# = 0
       ITOTA# = !ValAbsComp
       '
       SQLX$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
       SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
       SQLX$ = SQLX$ + "AND NuClien = " & NCLIE% & " "
       '
       With CABASELE
         .FindFirst SQLX$
25        If .NoMatch = False Then
           TIMOV$ = TRIM$(UCase$(!TipoMovim))
           IPARD# = !ValAbsComp
           If TIMOV$ = "VENTA" Or TIMOV$ = "EMBAL" Then
               INETO# = INETO# + IPARD#
             ElseIf TIMOV$ = "IVA" And UCase$(!OPCIMOVIM) = "DEFI" And UCase$(!VARIMOVIM) <> "RNI" Then
               IMIVA# = IMIVA# + IPARD#
             ElseIf TIMOV$ = "IVA" And UCase$(!OPCIMOVIM) = "DEFI" And UCase$(!VARIMOVIM) = "RNI" Then
               RVRNI# = IVRNI# + IPARD#
             ElseIf UCase$(!VARIMOVIM) Like "PER*" Then
               RETEN# = RETEN# + IPARD#
           End If
           .FindNext SQLX$
           GoTo 25
         End If
       End With
       '
       IOTRO# = ITOTA# - INETO# - IMIVA# - IVRNI# - RETEN#
       '
       Call REPLA(ZZ$, Chr$(2), 1, 1)
       Call REPLA(ZZ$, MKI$(FEFAXI%), 3, 2)
       Call REPLA(ZZ$, MKS$(NUMECO), 5, 4)
       Call REPLA(ZZ$, MKI$(NCLIE%), 9, 2)
       Call REPLA(ZZ$, MKD$(INETO#), 11, 8)
       Call REPLA(ZZ$, MKD$(IMIVA#), 19, 8)
       Call REPLA(ZZ$, MKD$(IVRNI#), 27, 8)
       Call REPLA(ZZ$, MKD$(RETEN#), 35, 8)
       Call REPLA(ZZ$, MKD$(IOTRO#), 43, 8)
       Call REPLA(ZZ$, MKD$(ITOTA#), 51, 8)
       JJ& = JJ& + 1
       Call GRAREG("DETACCX", ZZ$, JJ&)
       '
     .MoveNext
     Loop
   End With
   Call SETULTREG("DETACCX", JJ&)
   '
   Call HEADERS("DETACCX", "")
   Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(MESAN))
   Call HEADERS("DETACCX", "Rubro: Facturacion en Cuenta Corriente / Contado")
   Screen.MousePointer = 1
   If Option1.Value = True Then
        Call FINAL("*CODECCX")
     ElseIf Option2.Value = True Then
        Call FINAL("*LIDECCX")
   End If
99 Screen.MousePointer = 1
   Command3.Enabled = True
End Sub

Sub Command30_Click()
   Command30.Enabled = False
'\\\OT-5-0354-FG-03/06/05///
    Dim TOTVENTA#(32767)
    Dim TOTDOLA#(32767)
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    '
    OPX% = COMALTER%("Opciones de Formato de Salida\\Todos\Nacionales\Extranjeros")
'    Call SELECHO("VENDEDOR")
'    VENDEDOR% = XVALO(VALACT1$("VENDEDOR"))
'    DescVendedor$ = TRIM$(VALACT2$("VENDEDOR"))
    OPX1% = COMALTER%("Desea Incluir Notas de Débito\\Incluir\Excluir")
    EXCLUYE_ND% = 0
    If OPX1% > 1 Then EXCLUYE_ND% = 1
    NCMIN% = 32767: NCMAX% = 0
    IXXX$ = LOADFILE$(LUDAT$ + "DEUDOR1.X01")
   '
    For KKL& = 1 To Len(IXXX$) Step 2
     NCMX% = CVI(Mid$(IXXX$, KKL&, 2))
     If NCMX% > 0 Then
       If NCMX% < NCMIN% Then NCMIN% = NCMX%
       If NCMX% > NCMAX% Then NCMAX% = NCMX%
     End If
    Next KKL&
    '
10  VDEF$ = MKI$(NCMIN%) + MKI$(NCMAX%) + Chr$(0)
    OBX$ = OBJPLA$("DESHASCLIEVEN", VDEF$)
    If OBX$ = "" Then
     NCL1% = 0: NCL2% = 0
     GoTo 99
    End If
    '
    NCL1% = CVI(Left$(OBX$, 2))
    NCL2% = CVI(Mid$(OBX$, 3, 2))
    If NCL1% < 1 Or NCL2% < 1 Or NCL2% < NCL1% Then
     Call MENSERR(24, "Rango no Válido")
     GoTo 10
    End If
    VENDEDOR% = Asc(Mid$(OBX$, 5, 1))
    DescVendedor$ = DEVENDE$(VENDEDOR%)
    '
    Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTVENTA#(U&) = 0
    Next U&
    '
    DESDAT = Int(CDbl(CVDAT(DES%)))
    HASDAT = Int(CDbl(CVDAT(HAS%)))
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'VENTA' "
    '\\\OT-08-00677-OD-26/08/08///
    If EXCLUYE_ND% > 0 Then SQLX$ = SQLX$ + "AND Opcimovim <> 'ND'" ' para que no considere notas de Débito
    '\\\OT-08-00677-OD-FIN///
    SQLX$ = SQLX$ + "AND NuClien >= " & NCL1% & ""
    SQLX$ = SQLX$ + "AND NuClien <= " & NCL2% & ""
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Screen.MousePointer = 11
    CLIMAX% = 0
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
         NCLIE% = !NuClien
         If OPX% = 2 And PIVACLI%(NCLIE%) = 5 Then GoTo 35
         If OPX% = 3 And PIVACLI%(NCLIE%) <> 5 Then GoTo 35
         If NCLIE% > 0 Then
           If NCLIE% > CLIMAX% Then CLIMAX% = NCLIE%
           If VENDCLI%(NCLIE%) = VENDEDOR% Or VENDEDOR% = 0 Then
              TOTVENTA#(NCLIE%) = TOTVENTA#(NCLIE%) + (!IHABECOMP - !IDEBECOMP)
              If NCLIE% = 66 Then
              AA = BB
              COMPRO$ = !CODICOM_LAR
              End If
              If XVALO(!Mone_Emis) > 1 Then
                If Abs(!IHABECOMP - !IDEBECOMP) > 0.005 Then
                  SIGNO% = 1: If !IDEBECOMP > 0.005 Then SIGNO% = (-1)
                  TOTDOLA#(NCLIE%) = TOTDOLA#(NCLIE%) + XVALO(!Impo_Dola) * SIGNO%
                End If
              End If
           End If
         End If
35       .MoveNext
      Loop
    End With
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
    J& = 0: PORACU = 0
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
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call REPLA(Z$, MKD$(TOTDOLA#(NC%)), 21, 8)
        Call GRAREG("RANKCLI", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKCLI", J&)
    Call CIERRARCH("RANKCLI")
    Call HEADERS("RANKCLI", "")
    Call HEADERS("RANKCLI", "Periodo: " + PERIO$)
    If VENDEDOR% > 0 Then
      Call HEADERS("RANKCLI", "Vendedor: " + DescVendedor$)
    End If
    '
    Select Case OPX%
    Case Is = 1
      OpcionElegida$ = "Todos "
    Case Is = 2
      OpcionElegida$ = "Nacionales"
    Case Is = 3
      OpcionElegida$ = "Extranjeros"
    End Select
    '
    Call HEADERS("RANKCLI", "Filtro: " + OpcionElegida$)
    Screen.MousePointer = 1
    List1.Clear
    '
    If Option1.Value = True Then
         Call FINAL("*RANKCLI")
       Else
         Call FINAL("*LRANCLI")
    End If
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    Command30.Enabled = True
    '
'\\\OT-5-0354-FG-FIN///
End Sub

Sub Command4_Click()
   Command4.Enabled = False
   If VALIMESAN% < 1 Then GoTo 99
   '
   Screen.MousePointer = 11
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   'SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
   'SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
   SQLX$ = SQLX$ + "where FechConta >= " & DESDAT & " "
   SQLX$ = SQLX$ + "AND FechConta <= " & HASDAT & " "
   SQLX$ = SQLX$ + "AND TipoMovim = 'FVEN' "
   SQLX$ = SQLX$ + "AND OpciMovim = 'ND' "
   SQLX$ = SQLX$ + "Order by FechConta ASC, CodiCom_Lar ASC "
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where FechConta >= " & DESDAT & " "
   SQLX$ = SQLX$ + "AND FechConta <= " & HASDAT & " "
   Set CABASELE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   JJ& = 0: REBLA$ = REGBLAN$("DETACCX")
   With CABATEMP
     On Error Resume Next
     .MoveLast
     If Err > 0 Then
       On Error GoTo 0
       Call MENSERR(24, "Sin Movimientos en el Período Elegido")
       GoTo 99
     End If
     On Error GoTo 0
     FIN& = .RecordCount
     .MoveFirst
     KU& = 0
     Do While Not .EOF
       KU& = KU& + 1
       Call TRACE(20, KU&, FIN&)
       ZZ$ = REBLA$
       NUMEFA$ = !CODICOM_LAR
       NCLIE% = !NuClien
       FEFAXI% = CVINT(!FechConta)
       FECHAD = Int(CDbl(!FechConta))
       NUMECO = !NumeComp
       '
       INETO# = 0
       IMIVA# = 0
       IVRNI# = 0
       RETEN# = 0
       IOTRO# = 0
       ITOTA# = !ValAbsComp
       '
       SQLX$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
       SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
       SQLX$ = SQLX$ + "AND NuClien = " & NCLIE% & " "
       '
       With CABASELE
         .FindFirst SQLX$
25        If .NoMatch = False Then
           TIMOV$ = TRIM$(UCase$(!TipoMovim))
           IPARD# = !ValAbsComp
           If TIMOV$ = "VENTA" Or TIMOV$ = "EMBAL" Then
               INETO# = INETO# + IPARD#
             ElseIf TIMOV$ = "IVA" And UCase$(!OPCIMOVIM) = "DEFI" And UCase$(!VARIMOVIM) <> "RNI" Then
               IMIVA# = IMIVA# + IPARD#
             ElseIf TIMOV$ = "IVA" And UCase$(!OPCIMOVIM) = "DEFI" And UCase$(!VARIMOVIM) = "RNI" Then
               RVRNI# = IVRNI# + IPARD#
             ElseIf UCase$(!VARIMOVIM) Like "PER*" Then
               RETEN# = RETEN# + IPARD#
           End If
           .FindNext SQLX$
           GoTo 25
         End If
       End With
       '
       IOTRO# = ITOTA# - INETO# - IMIVA# - IVRNI# - RETEN#
       '
       Call REPLA(ZZ$, Chr$(2), 1, 1)
       Call REPLA(ZZ$, MKI$(FEFAXI%), 3, 2)
       Call REPLA(ZZ$, MKS$(NUMECO), 5, 4)
       Call REPLA(ZZ$, MKI$(NCLIE%), 9, 2)
       Call REPLA(ZZ$, MKD$(INETO#), 11, 8)
       Call REPLA(ZZ$, MKD$(IMIVA#), 19, 8)
       Call REPLA(ZZ$, MKD$(IVRNI#), 27, 8)
       Call REPLA(ZZ$, MKD$(RETEN#), 35, 8)
       Call REPLA(ZZ$, MKD$(IOTRO#), 43, 8)
       Call REPLA(ZZ$, MKD$(ITOTA#), 51, 8)
       JJ& = JJ& + 1
       Call GRAREG("DETACCX", ZZ$, JJ&)
       '
     .MoveNext
     Loop
   End With
   Call SETULTREG("DETACCX", JJ&)
   '
   Call HEADERS("DETACCX", "")
   Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(MESAN))
   Call HEADERS("DETACCX", "Rubro: Facturacion en Cuenta Corriente / Contado")
   Screen.MousePointer = 1
   If Option1.Value = True Then
        Call FINAL("*CODECCX")
     ElseIf Option2.Value = True Then
        Call FINAL("*LIDECCX")
   End If
99 Screen.MousePointer = 1
   Command4.Enabled = True
End Sub

Sub Command5_Click()
   Command5.Enabled = False
   If VALIMESAN% < 1 Then GoTo 99
   '
   Screen.MousePointer = 11
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   'SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
   'SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
   SQLX$ = SQLX$ + "where FechConta >= " & DESDAT & " "
   SQLX$ = SQLX$ + "AND FechConta <= " & HASDAT & " "
   SQLX$ = SQLX$ + "AND TipoMovim = 'FVEN' "
   SQLX$ = SQLX$ + "AND OpciMovim = 'NC' "
   SQLX$ = SQLX$ + "Order by FechConta ASC, CodiCom_Lar ASC "
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where FechConta >= " & DESDAT & " "
   SQLX$ = SQLX$ + "AND FechConta <= " & HASDAT & " "
   Set CABASELE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   JJ& = 0: REBLA$ = REGBLAN$("DETACCX")
   With CABATEMP
     On Error Resume Next
     .MoveLast
     If Err > 0 Then
       On Error GoTo 0
       Call MENSERR(24, "Sin Movimientos en el Período Elegido")
       GoTo 99
     End If
     On Error GoTo 0
     FIN& = .RecordCount
     .MoveFirst
     KU& = 0
     Do While Not .EOF
       KU& = KU& + 1
       Call TRACE(20, KU&, FIN&)
       ZZ$ = REBLA$
       NUMEFA$ = !CODICOM_LAR
       NCLIE% = !NuClien
       FEFAXI% = CVINT(!FechConta)
       FECHAD = Int(CDbl(!FechConta))
       NUMECO = !NumeComp
       '
       INETO# = 0
       IMIVA# = 0
       IVRNI# = 0
       RETEN# = 0
       IOTRO# = 0
       ITOTA# = !ValAbsComp
       '
       SQLX$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
       SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
       SQLX$ = SQLX$ + "AND NuClien = " & NCLIE% & " "
       '
       With CABASELE
         .FindFirst SQLX$
25        If .NoMatch = False Then
           TIMOV$ = TRIM$(UCase$(!TipoMovim))
           IPARD# = !ValAbsComp
           If TIMOV$ = "VENTA" Or TIMOV$ = "EMBAL" Then
               INETO# = INETO# + IPARD#
             ElseIf TIMOV$ = "IVA" And UCase$(!OPCIMOVIM) = "DEFI" And UCase$(!VARIMOVIM) <> "RNI" Then
               IMIVA# = IMIVA# + IPARD#
             ElseIf TIMOV$ = "IVA" And UCase$(!OPCIMOVIM) = "DEFI" And UCase$(!VARIMOVIM) = "RNI" Then
               RVRNI# = IVRNI# + IPARD#
             ElseIf UCase$(!VARIMOVIM) Like "PER*" Then
               RETEN# = RETEN# + IPARD#
           End If
           .FindNext SQLX$
           GoTo 25
         End If
       End With
       '
       IOTRO# = ITOTA# - INETO# - IMIVA# - IVRNI# - RETEN#
       '
       Call REPLA(ZZ$, Chr$(2), 1, 1)
       Call REPLA(ZZ$, MKI$(FEFAXI%), 3, 2)
       Call REPLA(ZZ$, MKS$(NUMECO), 5, 4)
       Call REPLA(ZZ$, MKI$(NCLIE%), 9, 2)
       Call REPLA(ZZ$, MKD$(INETO#), 11, 8)
       Call REPLA(ZZ$, MKD$(IMIVA#), 19, 8)
       Call REPLA(ZZ$, MKD$(IVRNI#), 27, 8)
       Call REPLA(ZZ$, MKD$(RETEN#), 35, 8)
       Call REPLA(ZZ$, MKD$(IOTRO#), 43, 8)
       Call REPLA(ZZ$, MKD$(ITOTA#), 51, 8)
       JJ& = JJ& + 1
       Call GRAREG("DETACCX", ZZ$, JJ&)
       '
     .MoveNext
     Loop
   End With
   Call SETULTREG("DETACCX", JJ&)
   '
   Call HEADERS("DETACCX", "")
   Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(MESAN))
   Call HEADERS("DETACCX", "Rubro: Facturacion en Cuenta Corriente / Contado")
   Screen.MousePointer = 1
   If Option1.Value = True Then
        Call FINAL("*CODECCX")
     ElseIf Option2.Value = True Then
        Call FINAL("*LIDECCX")
   End If
99 Screen.MousePointer = 1
   Command5.Enabled = True
End Sub

Sub Command6_Click()
'////*******///////
Command6.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
10  Screen.MousePointer = 11
    Dim SALINI#(32768), SUFAC#(32768), SUCOB#(32768)
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "or TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUD' "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    CLIMAX% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      UKL& = 0
      Do While Not .EOF
        UKL& = UKL& + 1
        Call TRACE(20, UKL&, FIN&)
        FEMII% = CVINT(!FechConta)
        If FEMII% <= HASTA% Then
          NCLIE% = !NuClien
          'NCLIE% = !NuProve
          If NCLIE% > 0 Then
            'If TICUEPRO((-1) * NCLIE%) = 0 Then ' solo cuentas corrientes
              If NCLIE% > CLIMAX% Then CLIMAX% = NCLIE%
              If !TipoMovim = "FVEN" Or !TipoMovim = "AJUD" Then
                  IFAC# = !IDEBECOMP - !IHABECOMP
                  IPAG# = 0
                Else
                  IFAC# = 0
                  IPAG# = !IHABECOMP - !IDEBECOMP
              End If
              '
              If FEMII% < DESDE% Then
                  SALINI#(NCLIE%) = SALINI#(NCLIE%) + IFAC# - IPAG#
                Else
                  SUFAC#(NCLIE%) = SUFAC#(NCLIE%) + IFAC#
                  SUCOB#(NCLIE%) = SUCOB#(NCLIE%) + IPAG#
              End If
            'End If
          End If
        End If
      .MoveNext
      Loop
    End With
    '
    JJ& = 0
    For k& = 1 To CLIMAX%
      If Abs(SALINI#(k&)) >= 0.005 Or Abs(SUFAC#(k&)) >= 0.005 Or Abs(SUCOB#(k&)) >= 0.005 Then
        x$ = REGBLAN$("RESUCCD")
        K1% = k&
        Call REPLA(x$, MKI$(K1%), 1, 2)
        Call REPLA(x$, MKD$(SALINI#(k&)), 3, 8)
        Call REPLA(x$, MKD$(SUFAC#(k&)), 11, 8)
        Call REPLA(x$, MKD$(SUCOB#(k&)), 19, 8)
        Call REPLA(x$, MKD$(SALINI#(k&) + SUFAC#(k&) - SUCOB#(k&)), 27, 8)
        JJ& = JJ& + 1
        Call GRAREG("RESUCCD", x$, JJ&)
       End If
    Next k&
    '
    Call SETULTREG("RESUCCD", JJ&)
        '
    Call HEADERS("RESUCCD", "")
    Call HEADERS("RESUCCD", "Corresponde a: " + TRIM$(CORRESP$))
    Screen.MousePointer = 1
    '
    If Option1.Value = True Then
        Call FINAL("*RESUCCD")
     ElseIf Option2.Value = True Then
    '\\\OT-08-0903-OD-26/09/08/// LISTADO DE RESUMEN MENSUAL CON CUIT
        OPXX% = ALTERNA%("Normal \Incluye C.U.I.T")
        If OPXX% < 1 Or OPXX% > 2 Then GoTo 99
        LISTAD$ = "*LIRECCD" 'LISTADO NORMAL
        If OPXX% = 1 Then GoTo 98
        If OPXX% = 2 Then LISTAD$ = "*LIRCCD1" ' LISTADO CON CUIT
        '
        '\\\OT-08-0903-OD-26/09/08/// LISTADO DE RESUMEN MENSUAL CON CUIT
        JK& = 0
        FIN& = ULTREG&("RESUCCD")
        For II& = 1 To FIN&
            JK& = JK& + 1
            XXO$ = REGLEIDO$("RESUCCD", II&) 'LEO EL ARCHIVO YA CARGADO Y LE AGREGO EL CUIT
            '
            Call REPLA(TTXX$, Mid$(XXO$, 1, 2), 1, 2)
            K1% = CVI(Mid$(XXO$, 1, 2))
            Call REPLA(TTXX$, CUITCLI$(K1%), 3, 16)
            Call REPLA(TTXX$, Mid$(XXO$, 3, 8), 19, 8)
            Call REPLA(TTXX$, Mid$(XXO$, 11, 8), 27, 8)
            Call REPLA(TTXX$, Mid$(XXO$, 19, 8), 35, 8)
            Call REPLA(TTXX$, Mid$(XXO$, 27, 8), 43, 8)
            Call GRAREG("RESCCD1", TTXX$, II&) 'GRABO EN ARCHIVO NUEVO CON CUIT
            '
        Next II&
        Call SETULTREG("RESCCD1", JK&)
        '
        Call HEADERS("RESCCD1", "")
        Call HEADERS("RESCCD1", "Corresponde a: " + TRIM$(CORRESP$))
            '
98      Call FINAL(LISTAD$)
    End If
    '\\\OT-08-0903-OD-FIN///

    
99 Command6.Enabled = True
'////***antigua***///
'    If Option1.Value = True Then
'         Call FINAL("*RESUCCD")
'      ElseIf Option2.Value = True Then
'         Call FINAL("*LIRECCD")
'    End If
End Sub

Sub Command7_Click()
    '
    Command7.Enabled = False
    '
    OPXX% = ALTERNA%("Listado de Control Mensual\Control de Errores en Ajustes")
    If OPXX% = 1 Then
        Screen.MousePointer = 11
        LU$ = LUDAT$
        C1$ = TRIM$(REMEVE04.MESAN)
        '
        If VALIMESAN% < 1 Then GoTo 99
        '
        Screen.MousePointer = 11
        Call APERBASDAT("CABAN")
        SQLX$ = "SELECT * FROM CAJABANC "
        SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
        SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
        SQLX$ = SQLX$ + "AND TipoMovim = 'AJUD' "
        SQLX$ = SQLX$ + "Order by FechConta ASC, NumeComp ASC, NuClien ASC "
        '
        Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
        '
        J1& = 0
        REBLA1$ = REGBLAN$("DETACCX")
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
            ZZ$ = REBLA1$
            NCLIE% = !NuClien
            FECHAD = Int(CDbl(!FechConta))
            INETO# = (!IDEBECOMP - !IHABECOMP)
            ITOTA# = (!IDEBECOMP - !IHABECOMP)
            '
            Call REPLA(ZZ$, Chr$(4), 1, 1)
            Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 3, 2)
            Call REPLA(ZZ$, MKS$(!NumeComp), 5, 4)
            Call REPLA(ZZ$, MKI$(!NuClien), 9, 2)
            Call REPLA(ZZ$, MKD$(INETO#), 11, 8)
            Call REPLA(ZZ$, MKD$(ITOTA#), 51, 8)
            J1& = J1& + 1
            Call GRAREG("DETACCX", ZZ$, J1&)
            '
            .MoveNext
          Loop
        End With
        Call SETULTREG("DETACCX", J1&)
        '////*****FINNUEVO***/////
    
        Call HEADERS("DETACCX", "")
        Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
        Call HEADERS("DETACCX", "Rubro: Otros Movimientos de Cuenta Corriente")
        Screen.MousePointer = 1
        If Option1.Value = True Then
             Call FINAL("*CODECCX")
          ElseIf Option2.Value = True Then
             Call FINAL("*LIDECCX")
        End If
        '
      ElseIf OPXX% = 2 Then
        '
        Call CONTROLAJUSTE
      Else
        GoTo 99
    End If
     
99  Command7.Enabled = True
    Screen.MousePointer = 1
 End Sub

Sub CONTROLAJUSTE()

    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where TipoMovim = 'AJUD'"
    SQLX$ = SQLX$ + "Order by FechConta ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    J1& = 0
    REBLA1$ = REGBLAN$("DETACCX")
    '
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "No se Registraron Movimientos de Ajustes")
        GoTo 99
      End If
      On Error GoTo 0
      Do While Not .EOF
        If TRIM$(Mid$(!Descrimov, 12, 30)) <> TRIM$(Mid$(!Refecom, 1, 30)) Then
          If UCase$(Left$(!Descrimov, 12)) <> "COMPENSACION" And XVALO(!ValAbsComp) >= 0.05 Then
          ZZ$ = REBLA1$
          NCLIE% = !NuClien
          FECHAD = Int(CDbl(!FechConta))
          INETO# = !ValAbsComp
          ITOTA# = !ValAbsComp
          '
          Call REPLA(ZZ$, Chr$(4), 1, 1)
          Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 3, 2)
          Call REPLA(ZZ$, MKS$(!NumeComp), 5, 4)
          Call REPLA(ZZ$, MKI$(!NuClien), 9, 2)
          Call REPLA(ZZ$, MKD$(INETO#), 11, 8)
          Call REPLA(ZZ$, MKD$(ITOTA#), 51, 8)
          J1& = J1& + 1
          Call GRAREG("DETACCX", ZZ$, J1&)
          '
          End If
        End If
        .MoveNext
      Loop
    End With
    '
    If J1& < 1 Then
        Call MENSERR(24, "No se Registraron Errores por Movimientos de Ajuste")
        GoTo 99
    End If
    '
    Call SETULTREG("DETACCX", J1&)
    '////*****FINNUEVO***/////

    Call HEADERS("DETACCX", "")
    Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Call HEADERS("DETACCX", "Rubro: Errores por Movimientos de Ajuste")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODECCX")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDECCX")
    End If
99
End Sub


Sub Command8_Click()
    Command8.Enabled = False
    '
    If CONTROL$("FACTURA", "SALIDA") = "CFISCAL" Then
       ' INICIALIZAR CONTROLADOR FISCAL
       Call FORHASAR.REPORTES
       GoTo 99
    End If
    '
    If EXISTE%("LCOMIS04.EXE") > 0 Then
      OPXX% = ALTERNA%("Liquidación de Comisiones\Listado de Debitos y Creditos\Subdiario de Retenciones de Ingresos Brutos\Otros Accesos Directos")
      If OPXX% < 1 Or OPXX% > 2 Then GoTo 99
      If OPXX% = 1 Then
          Call CONECRUN("LCOMIS04", "Liquidación de Comisiones")
        ElseIf OPXX% = 2 Then
          Call LISTADEBCRE
        ElseIf OPXX% = 3 Then
          Call RECICO04.LISRETIIBB
        ElseIf OPXX% = 4 Then
          GoTo 20
      End If
    End If
    GoTo 99
    '
    OPXX% = ALTERNA%("Listado de Debitos y Creditos\Subdiario de Retenciones de Ingresos Brutos\Otros Accesos Directos")
    If OPXX% < 1 Or OPXX% > 2 Then GoTo 99
    If OPXX% = 1 Then
      Call LISTADEBCRE
    ElseIf OPXX% = 2 Then
      Call RECICO04.LISRETIIBB
    End If
    GoTo 99
    '
20  Call ACCDIR("COGEVEN")
    '
99  Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCHIG04/AMACLI", "Archivos Maestros")
   Command9.Enabled = True
End Sub

Private Sub config_Click()

   If CLACONTRA% > 0 Then
      Call Command14_Click
   End If
End Sub

Private Sub consultadeasientocontableporpantalla_Click()
   Call Command20_Click
End Sub

Private Sub estadisticodecompras_Click()

End Sub

Private Sub detalledeventas_Click()
    Call Command15_Click
End Sub

Private Sub facturasdecompra_Click()
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      Me.Frame9.BackColor = &HFCD7B6
      Me.Frame10.BackColor = &HFCD7B6
      Me.Frame11.BackColor = &HFCD7B6
      Me.Frame8.Width = 9240
      Me.Option1.BackColor = &HFCD7B6
      Me.Option2.BackColor = &HFCD7B6
      
      MARCOBARRA.Top = 5950
      
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
      Picture10.Visible = True
    End If
    If CONTROL("FACTURA", "SISTMANT") = "SI" Then
        archivostransferenciaingresobrutoganancia.Visible = False
        consultadeasientocontableporpantalla.Visible = False
        resumenprovinciaygrupos.Visible = False
        detalledeventas.Visible = False
        varios.Visible = False
        notascredito.Visible = False
        notadebito.Visible = False
        facturasdeventas.Visible = False
    End If
    
    If EXISTE%("LCOMIS04.EXE") > 0 Then
       varios.Enabled = True
    End If
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub



Private Sub HASAR1_ImpresoraNoResponde(ByVal CantidadReintentos As Long)

End Sub

Private Sub RegimenInformativoVentas()
    
    On Error Resume Next
    Call TRALOCAL("LIVAVEN", "")
    On Error GoTo 0
    
    If Err > 0 Then
        Call MENSERR(24, "Imposible Copiar Archivo con Datos del Período")
        Exit Sub
    End If
    
    PERIODO$ = TRIM$(MESAN)
    ' borrar archivo RIV
    NX% = FILEOPEN%(LUCOM$ & "REGINFO_CV_VENTAS_CBTE_" & PERIODO$ & ".TXT", 2, 266)
    NX1% = FILEOPEN%(LUCOM$ & "REGINFO_CV_VENTAS_ALICUOTAS_" & PERIODO$ & ".TXT", 2, 62)

    UltimoNumeroComprobante$ = ""
    
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
    
    Dim FlagError%
    Dim ListaDeErrores$
    
    INDICE% = 1: GoSub InicializarCamposVGD
    INDICE% = 2: GoSub InicializarCamposVGD
    INDICE% = 3: GoSub InicializarCamposVGD
    
    FIN& = ULTREG&("!LIVAVEN")
    If FIN& = 0 Then GoTo 99
    
    Call ABRECABAN
    
    Screen.MousePointer = 11
    For i& = 1 To FIN&
        Call TRACE(24, i&, ULTREG&("!LIVAVEN"))
        
        REGISTRO$ = REGLEIDO$("!LIVAVEN", i&)
        NumeroComprobante$ = Mid$(REGISTRO$, 11, 18)
        
        If TRIM$(NumeroComprobante$) = "" Then
            NumeroComprobante$ = UltimoNumeroComprobante$
        End If
        
        If NumeroComprobante$ <> UltimoNumeroComprobante$ Then
            '
            FechaComprobante$ = FECHATEXLAR$(CVI(Mid$(REGISTRO$, 3, 2)))
            FechaComprobante$ = Mid$(FechaComprobante$, 7, 4) & Mid$(FechaComprobante$, 4, 2) & Mid$(FechaComprobante$, 1, 2)
            
            If UltimoNumeroComprobante$ <> "" And FlagError% = 0 Then
                If TRIM$(RegistroRIV$) <> "" Then
                    Print #NX%, RegistroRIV$
                    RegistroRIV$ = ""
                    RegistroRIV_Alicuotas$ = ""
                End If
            End If
            FlagError% = 0
            
25          'NumeroCliente& = XVALO(VALOBADA("CAJABANC", "NuClien", "CODICOM_LAR='" & NumeroComprobante$ & "' AND FechConta='" & FechaCbte$ & "'"))
            NumeroCliente& = CVI(Mid$(REGISTRO$, 31, 2))
            If NumeroCliente& = 0 Then
                'Call MENSERR(24, "Imposible Procesar el Movimiento.\No Se Pudo Encontrar Numero de Cliente para el Comprobante '" & NumeroComprobante$ & "'")
                ListaDeErrores$ = ListaDeErrores$ & "Imposible Procesar el Movimiento. No Se Pudo Encontrar Numero de Cliente para el Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                FlagError% = 1: GoTo 50
            End If
    
            TipoCbte$ = TipoComprobante$(NumeroComprobante$)
            
            PuntoVta$ = PuntoVenta$(NumeroComprobante$)
            
            NumeroCbt$ = NumeroCbte$(NumeroComprobante$)
            
            CodigoDocComprador$ = CodigoDocumento$(PIVACLI%(NumeroCliente&))
            
            NumeroIdentificacionComprador$ = TRIM$(Mid$(REGISTRO$, 63, 16))
            If NumeroIdentificacionComprador$ = "" Or VALICUIT%(NumeroIdentificacionComprador$) = 0 Then
                NumeroIdentificacionComprador$ = CUITCLI$(NumeroCliente&)
            End If
            If NumeroIdentificacionComprador$ = "" Then
                'Call MENSERR(24, "Imposible Procesar el Movimiento.\CUIT Inválido para el Comprobante '" & NumeroComprobante$ & "'")
                ListaDeErrores$ = ListaDeErrores$ & "CUIT Inválido en el Cliente '" & CStr(NumeroCliente&) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                FlagError% = 1: GoTo 50
            End If
            If PIVACLI%(NumeroCliente&) <> 5 And PIVACLI%(NumeroCliente&) <> 0 Then
                If VALICUIT%(NumeroIdentificacionComprador$) = 0 Then
                    'Call MENSERR(24, "Imposible Procesar el Movimiento.\CUIT Inválido para el Comprobante '" & NumeroComprobante$ & "'")
                    ListaDeErrores$ = ListaDeErrores$ & "CUIT Inválido en el Cliente '" & CStr(NumeroCliente&) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                    FlagError% = 1: GoTo 50
                End If
            Else
                NumeroIdentificacionComprador$ = CUITCLI$(NumeroCliente&)
            End If
            NumeroIdentificacionComprador$ = REPLACAR(NumeroIdentificacionComprador$, "-", "")
            NumeroIdentificacionComprador$ = AJUSTD$(NumeroIdentificacionComprador$, 20)
            NumeroIdentificacionComprador$ = REPLACAR(NumeroIdentificacionComprador$, " ", "0")
            
            RazonSocial$ = AJUSTI$(RASOCLI$(NumeroCliente&), 30)
            If PIVACLI%(NumeroCliente&) = 0 Then
                RazonSocial$ = AJUSTI$("CONSUMIDOR FINAL", 30)
            End If
            
            ImporteTotal$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 91, 8)), 15)
            
            ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
            If Mid$(REGISTRO$, 129, 8) = "NO GRAV." Then
                ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) + CVD(Mid$(REGISTRO$, 153, 8)) + CVD(Mid$(REGISTRO$, 161, 8)), 15)
            End If
            
            PercepcionNoCategorizados$ = RIVFormatoNumero$(0, 15)
            
            ImporteOperacionExenta$ = RIVFormatoNumero$(0, 15)
            
            If PIVACLI%(NumeroCliente&) = 5 Or PIVACLI%(NumeroCliente&) = 7 Then
                ImporteOperacionExenta$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 185, 8)), 15)
                ImporteNoGravado$ = RIVFormatoNumero$(0, 15) ' O es No Gravado o es Exento
            End If
            
            ImportePercepcionesNacionales$ = RIVFormatoNumero$(0, 15)
            
            ImportePercepcionIB$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 193, 8)), 15)
            
            ImportePercepcionMunicipales$ = RIVFormatoNumero$(0, 15)
            
            ImporteImpuestosInternos$ = RIVFormatoNumero$(0, 15)
                
                MONE% = XVALO(VALOBADAMDB("CAJABANC", "Mone_Emis", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuClien=" & NumeroCliente&))
            CodMoneda$ = CodigoMoneda(MONE%)
            If TRIM$(CodMoneda$) = "" Then CodMoneda$ = "PES"
                
                TCambio# = XVALO(VALOBADAMDB("CAJABANC", "Tipo_Cam", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuClien=" & NumeroCliente&))
            TipoCambio$ = RIVFormatoNumero$(TCambio#, 10, 6)
            If CodMoneda$ = "PES" And (XVALO(TipoCambio$) / 1000000) > 1 Then
                    ListaDeErrores$ = ListaDeErrores$ & "Error en Tipo de Cambio en el Cliente '" & CStr(NumeroCliente&) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                    FlagError% = 1: GoTo 50
            End If
            If XVALO(TipoCambio$) = 0 And (CodMoneda$ = "" Or CodMoneda$ = "PES") Then TipoCambio$ = RIVFormatoNumero(1, 10, 6)
                
                            
            ' Debe existir por lo menos una alicuota aunque sea no gravado o exento. No permite 0 (cero)
                CantAlicuotas% = 1
            CantidadAlicuotas$ = CStr(CantAlicuotas%)
            
            If PIVACLI%(NumeroCliente&) = 7 Then
                CodigoOperacion$ = "Z"
            ElseIf PIVACLI%(NumeroCliente&) = 5 Then
                CodigoOperacion$ = "X"
            ElseIf Mid$(REGISTRO$, 129, 8) = "NO GRAV." Then
                CodigoOperacion$ = "N"
            ElseIf PIVACLI%(NumeroCliente&) = 4 Then
                CodigoOperacion$ = "E"
            ElseIf UCase$(InStr(1, REGISTRO$, "COMPROBANTE ANULADO")) > 0 Then
                CodigoOperacion$ = "E"
            Else
                CodigoOperacion$ = " "
            End If
            
            OtrosTributos$ = RIVFormatoNumero$(0, 15)
            
            ' Si es una Exportacion o Consumidor Final, es la misma que la fecha del comprobante.
            ' Caso contrario la calcula
            FechaVencimientoCbte$ = FechaComprobante$
            If PIVACLI%(NumeroCliente&) <> 5 And PIVACLI%(NumeroCliente&) <> 7 And PIVACLI%(NumeroCliente&) <> 0 Then
                FechaVencimientoCbte$ = FechaComprobante$ ' calcular!!!
            ElseIf PIVACLI%(NumeroCliente&) = 5 Then
                FechaVencimientoCbte$ = "00000000"
            End If
            
               ' Arma el registro para el regimen informativo
               RegistroRIV$ = FechaComprobante$ & TipoCbte$ & PuntoVta$ & NumeroCbt$ & NumeroCbt$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & RazonSocial$ _
                               & ImporteTotal$ & ImporteNoGravado$ & PercepcionNoCategorizados$ & ImporteOperacionExenta$ & ImportePercepcionesNacionales$ _
                               & ImportePercepcionIB$ & ImportePercepcionMunicipales$ & ImporteImpuestosInternos$ & CodMoneda$ & TipoCambio$ & CantidadAlicuotas$ _
                               & CodigoOperacion$ & OtrosTributos$ & FechaVencimientoCbte$
                
               UltimoNumeroComprobante$ = NumeroComprobante$
            
            ValorIva$ = TRIM$(REPLACAR$(Mid$(REGISTRO$, 129, 8), "%", ""))
            If InStr(1, UCase(ValorIva$), "NO GRAV.") = 0 And (XVALO(ImporteOperacionExenta$) / 100) < 0.01 Then
                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                ImporteIva$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
                Alicuota$ = CodigoAlicuota(ValorIva$)
                If UCase$(InStr(1, REGISTRO$, "COMPROBANTE ANULADO")) > 0 Then Alicuota$ = "0003"
                
                RegistroRIV_Alicuotas$ = TipoCbte$ & PuntoVta$ & NumeroCbt$ & ImporteNeto$ & Alicuota$ & ImporteIva$
                Print #NX1%, RegistroRIV_Alicuotas$
            Else
                ConIva# = XVALO(VALOBADAMDB("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuClien=" & NumeroCliente&))
                If (ConIva# < 0.01 Or (XVALO(ImportePercepcionIB$) / 100) > 0.005) And (XVALO(ImporteOperacionExenta$) / 100) < 0.01 Then
                    ImporteNeto$ = ImporteTotal$
                    If (XVALO(ImportePercepcionIB$) / 100) > 0.005 Then ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                    
                    ImporteIva$ = RIVFormatoNumero(0, 15)
                    Alicuota$ = "0003"
                    ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
                    Call REPLA(RegistroRIV$, ImporteNoGravado$, 124, 15)
                    
                    RegistroRIV_Alicuotas$ = TipoCbte$ & PuntoVta$ & NumeroCbt$ & ImporteNeto$ & Alicuota$ & ImporteIva$
                    Print #NX1%, RegistroRIV_Alicuotas$
                    
                ElseIf (XVALO(ImporteOperacionExenta$) / 100) > 0.005 Then
                    
                    ImporteNeto$ = ImporteTotal$
                    ImporteIva$ = RIVFormatoNumero(0, 15)
                    Alicuota$ = "0003"
                    ImporteOperacionExenta$ = RIVFormatoNumero$(0, 15)
                    Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 154, 15)
                    
                    RegistroRIV_Alicuotas$ = TipoCbte$ & PuntoVta$ & NumeroCbt$ & ImporteNeto$ & Alicuota$ & ImporteIva$
                    Print #NX1%, RegistroRIV_Alicuotas$
                End If
            End If
'            ValorIva$ = TRIM$(REPLACAR$(Mid$(REGISTRO$, 129, 8), "%", ""))
'            If InStr(1, UCase(ValorIva$), "NO GRAV.") = 0 Then
'                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
'                ImporteIva$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
'                Alicuota$ = CodigoAlicuota(ValorIva$)
'
'                RegistroRIV_Alicuotas$ = TipoCbte$ & PuntoVta$ & NumeroCbt$ & ImporteNeto$ & Alicuota$ & ImporteIva$
'                Print #NX1%, RegistroRIV_Alicuotas$
'            ElseIf PIVACLI%(NumeroCliente&) = 5 Then
'                ImporteNeto$ = ImporteOperacionExenta$
'                ImporteIva$ = RIVFormatoNumero(0, 15)
'                Alicuota$ = "0003"
'                ImporteOperacionExenta$ = RIVFormatoNumero(0, 15)
'
'                RegistroRIV_Alicuotas$ = TipoCbte$ & PuntoVta$ & NumeroCbt$ & ImporteNeto$ & Alicuota$ & ImporteIva$
'                Print #NX1%, RegistroRIV_Alicuotas$
'            End If
        Else
            If Mid$(REGISTRO$, 129, 8) = "NO GRAV." Then
                ImporteNGAnterior = XVALO(Mid$(RegistroRIV$, 124, 15)) / 100
                ImporteNGActual = CVD(Mid$(REGISTRO$, 137, 8)) + CVD(Mid$(REGISTRO$, 153, 8)) + CVD(Mid$(REGISTRO$, 161, 8))
                If Not (ImporteNGActual < 0 And Mid$(REGISTRO$, 11, 2) = "FC") Then
                    ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior + CVD(Mid$(REGISTRO$, 137, 8)) + CVD(Mid$(REGISTRO$, 153, 8)) + CVD(Mid$(REGISTRO$, 161, 8)), 15)
                End If
            End If
           
            If PIVACLI%(NumeroCliente&) = 5 Or PIVACLI%(NumeroCliente&) = 7 Then
                IMPO# = XVALO(Mid$(RegistroRIV$, 154, 15)) / 100
                ImporteOperacionExenta$ = RIVFormatoNumero(IMPO# + CVD(Mid$(REGISTRO$, 185, 8)), 15)
                ImporteNoGravado$ = RIVFormatoNumero$(0, 15) ' O es No Gravado o es Exento
            End If
                          
            IMPO# = XVALO(Mid$(RegistroRIV$, 184, 15))
            ImportePercepcionIB$ = RIVFormatoNumero(IMPO# + CVD(Mid$(REGISTRO$, 193, 8)))
                               
                CantAlicuotas% = XVALO(CantidadAlicuotas$)
                If Mid$(REGISTRO$, 129, 8) <> "NO GRAV." Then CantAlicuotas% = CantAlicuota% + 1
            CantidadAlicuotas$ = CStr(CantAlicuotas%)
               
            RegistroRIV$ = FechaComprobante$ & TipoCbte$ & PuntoVta$ & NumeroCbt$ & NumeroCbt$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & RazonSocial$ _
                           & ImporteTotal$ & ImporteNoGravado$ & PercepcionNoCategorizados$ & ImporteOperacionExenta$ & ImportePercepcionesNacionales$ _
                           & ImportePercepcionIB$ & ImportePercepcionMunicipales$ & ImporteImpuestosInternos$ & CodMoneda$ & TipoCambio$ & CantidadAlicuotas$ _
                           & CodigoOperacion$ & OtrosTributos$ & FechaVencimientoCbte$
               
            ImporteNeto0# = CVD(Mid$(REGISTRO$, 137, 8))
            ValorIva$ = TRIM$(REPLACAR$(Mid$(REGISTRO$, 129, 8), "%", ""))
            If Not (ImporteNeto0# < 0 And Mid$(REGISTRO$, 11, 2) = "FC") And InStr(1, UCase(ValorIva$), "NO GRAV.") = 0 Then
                ImporteNeto$ = RIVFormatoNumero(ImporteNeto0#, 15)
                Alicuota$ = CodigoAlicuota(ValorIva$)
                ImporteIva$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
                   
                RegistroRIV_Alicuotas$ = TipoCbte$ & PuntoVta$ & NumeroCbt$ & ImporteNeto$ & Alicuota$ & ImporteIva$
                Print #NX1%, RegistroRIV_Alicuotas$
            End If
        End If
50  Next i&
    
    If NumeroComprobante$ <> "" Then Print #NX%, RegistroRIV$
    
    Close #NX%
    Close #NX1%
        
    Screen.MousePointer = 1
    
'    Dim XX1 As String * 266
'
'    ARCHIIBB$ = "REGINFO_CV_VENTAS_CBTE_" & PERIODO$ & ".TXT"
'    NARCHI% = FILEOPEN%(LUCOM$ + ARCHIIBB$, 1, 266)
'    Do While Not EOF(NARCHI%)
'        Line Input #NARCHI%, XX$
'        SIGNO% = 1
'        If Mid$(XX$, 9, 3) = "003" Or Mid$(XX$, 9, 3) = "008" Or Mid$(XX$, 9, 3) = "021" Then SIGNO% = -1
'
'        IMPO1# = IMPO1# + SIGNO% * XVALO(Mid$(XX$, 109, 15)) / 100
'        impo2# = impo2# + SIGNO% * XVALO(Mid$(XX$, 124, 15)) / 100
'        impo3# = impo3# + SIGNO% * XVALO(Mid$(XX$, 139, 15)) / 100
'        impo4# = impo4# + SIGNO% * XVALO(Mid$(XX$, 154, 15)) / 100
'        impo5# = impo5# + SIGNO% * XVALO(Mid$(XX$, 169, 15)) / 100
'        impo6# = impo6# + SIGNO% * XVALO(Mid$(XX$, 184, 15)) / 100
'        impo7# = impo7# + SIGNO% * XVALO(Mid$(XX$, 199, 15)) / 100
'        impo8# = impo8# + SIGNO% * XVALO(Mid$(XX$, 214, 15)) / 100
'        impo9# = impo9# + SIGNO% * XVALO(Mid$(XX$, 244, 15)) / 100
'20  Loop
'
'    MsgBox ("Total: " & TRIM$(FORMATNUM$(IMPO1#, "F15.2")) & vbCrLf _
'        & "No gravado: " & TRIM$(FORMATNUM$(impo2#, "F15.2")) & vbCrLf _
'        & "Percep. No Cat: " & TRIM$(FORMATNUM$(impo3#, "F15.2")) & vbCrLf _
'        & "Op.Exentas: " & TRIM$(FORMATNUM$(impo4#, "F15.2")) & vbCrLf _
'        & "Imp.Nacionales: " & TRIM$(FORMATNUM$(impo5#, "F15.2")) & vbCrLf _
'        & "Imp.IIBB: " & TRIM$(FORMATNUM$(impo6#, "F15.2")) & vbCrLf _
'        & "Imp.Municipales: " & TRIM$(FORMATNUM$(impo7#, "F15.2")) & vbCrLf _
'        & "Imp.Internos: " & TRIM$(FORMATNUM$(impo8#, "F15.2")) & vbCrLf _
'        & "Otro Atributos: " & TRIM$(FORMATNUM$(impo9#, "F15.2")) & vbCrLf)
'
'    Close #NARCHI%
'
'    ARCHIIBB$ = "REGINFO_CV_VENTAS_ALICUOTAS_" & PERIODO$ & ".TXT"
'    NARCHI% = FILEOPEN%(LUCOM$ + ARCHIIBB$, 1, 62)
'    Do While Not EOF(NARCHI%)
'        Line Input #NARCHI%, XX$
'        SIGNO% = 1
'        If Mid$(XX$, 1, 3) = "003" Or Mid$(XX$, 1, 3) = "008" Or Mid$(XX$, 1, 3) = "021" Then SIGNO% = -1
'        IMPO20# = IMPO20# + SIGNO% * XVALO(Mid$(XX$, 29, 15)) / 100
'        impo21# = impo21# + SIGNO% * XVALO(Mid$(XX$, 48, 15)) / 100
'    Loop
'
'    MsgBox ("Neto: " & TRIM$(FORMATNUM$(IMPO20#, "F15.2")) & vbCrLf _
'        & "Iva: " & TRIM$(FORMATNUM$(impo21#, "F15.2")))
'
'    Close #NARCHI%
    
    If TRIM$(ListaDeErrores$) <> "" Then
        Call COMUNI("Se Han Detectado Errores en la Generación.\Pulse Aceptar para Ver el Detalle")
        MSJESREGINFO.Text1 = ListaDeErrores$
        MSJESREGINFO.Caption = "Régimen Informativo de Ventas"
        Call MSJESREGINFO.Command1_Click
        MSJESREGINFO.Show 1
        Call MENSERR(24, "Tener Presente que los Datos Generados Son Incompletos.\Debe Corregir los Errores y Volver a Generar")
    Else
        Call COMUNI("Regimen Informativo de Ventas: Generación Ok")
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

Function NumeroCbte(NumeroComprobante$) As String

    NroComprobante$ = REPLACAR$(NumeroComprobante$, "-", "")
    NroComprobante$ = Mid$(NroComprobante$, 7, 8)
    
    Do While Len(NroComprobante$) < 20
        NroComprobante$ = "0" & NroComprobante$
    Loop
    
    NumeroCbte = NroComprobante$
    
End Function

Function PuntoVenta(NumeroComprobante$) As String

    PVenta$ = Mid$(NumeroComprobante$, 4, 4)
    
    PuntoVenta$ = "0" & PVenta$
    
End Function
Function CodigoAlicuota(Porcentaje$) As String

    Dim CODIGO$
    If (XVALO(Porcentaje$) = 10.5) Then
        CODIGO$ = "0004"
    ElseIf (XVALO(Porcentaje$) = 21) Then
        CODIGO$ = "0005"
    ElseIf (Porcentaje$ = "NO GRAV.") Then
        CODIGO$ = "0000"
    End If
    
    CodigoAlicuota = CODIGO$
    
End Function


Function RIVFormatoNumero(NUMERO#, Optional Longitud%, Optional Decimales%) As String
    
    If Longitud <= 0 Then Longitud = 15
    If Decimales% <= 0 Then Decimales% = 2
    VALOR$ = TRIM$(FORMATNUM$(Abs(NUMERO#), "F" & CStr(Longitud%) & "." & Decimales%))
    VALOR$ = REPLACAR$(VALOR$, ".", "")
    If VALOR$ = "" Then VALOR$ = "0"
    
    Do While Len(VALOR$) < Longitud%
        VALOR$ = "0" & VALOR$
    Loop
    
    RIVFormatoNumero = VALOR$
    
End Function

Function TipoComprobante(NumeroComprobante$) As String

    IdCbte$ = Mid$(NumeroComprobante$, 1, 2) & Mid$(NumeroComprobante$, Len(NumeroComprobante$), 1)
    
    TIPO$ = ""
    Select Case IdCbte$
        Case "FCA": TIPO$ = "001"
        Case "NDA": TIPO$ = "002"
        Case "NCA": TIPO$ = "003"
        Case "FCB": TIPO$ = "006"
        Case "NDB": TIPO$ = "007"
        Case "NCB": TIPO$ = "008"
        Case "FCE": TIPO$ = "019"
        Case "NDE": TIPO$ = "020"
        Case "NCE": TIPO$ = "021"
    End Select
    
    TipoComprobante$ = TIPO$
    
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
        Case Else: CodiMo$ = "000"
    End Select
    
    CodigoMoneda = CodiMo$
    
End Function
Private Sub Image2_Click()
   TXT$ = MESAN.TEXT
   If TXT$ = "" Then Exit Sub
   If MESANO$(TXT$, DES%, HAS%) = TXT$ Then
      If DERACCE%("CIEMEVEN", "Cierre y Re-Apertura Mensual Ventas") >= 2 Then
         If COMALTER%("Confirme Cierre del Período Mensual " + TXT$ + "\\Cancelar\Cerrar") = 2 Then
           Call GRACONTROL("", "CIEMEVEN", TRIM$(Str$(HAS%)))
           Image2.Visible = False
           Image1.Visible = True
         End If
      End If
   End If
End Sub

Private Sub Image1_Click()
   TXT$ = MESAN.TEXT
   If TXT$ = "" Then Exit Sub
   If MESANO$(TXT$, DES%, HAS%) = TXT$ Then
      If DERACCE%("CIEMEVEN", "Cierre y Re-Apertura Mensual Ventas") >= 2 Then
         If COMALTER%("Confirme Re-Apertura del Período Mensual\\Cancelar\Re-Abrir") = 2 Then
            Call GRACONTROL("", "CIEMEVEN", TRIM$(Str$(DES% - 1)))
            Image1.Visible = False
            Image2.Visible = True
         End If
      End If
   End If
End Sub

Private Sub impresora_Click()
     Option2.Value = True
     Call Command7_Click

End Sub

Private Sub impresora0_Click()
   Option2.Value = True
   Call Command5_Click

End Sub

Private Sub impresora1_Click()
     Option2.Value = True
     Call Command4_Click

End Sub

Private Sub impresora2_Click()
     Option2.Value = True
     Call Command2_Click

End Sub

Private Sub impresora4_Click()
     Option2.Value = True
     Call Command3_Click

End Sub

Private Sub impresora5_Click()
     Option2.Value = True
     Call Command6_Click

End Sub

Private Sub impresora6_Click()
   Option2.Value = True
   Call Command19_Click

End Sub

Private Sub impresora7_Click()
    Option2.Value = True
    Call Command30_Click

End Sub

Private Sub impresora8_Click()
   Option2.Value = True
   Call Command11_Click

End Sub

Private Sub maestrocientes_Click()
   Call Command9_Click
End Sub

Private Sub MESAN_DblClick()
   Call Command12_Click
End Sub

Private Sub MESAN_Change()
   TXT$ = MESAN.TEXT
   Image2.Visible = True
   Image1.Visible = False
   If TXT$ = "" Then Exit Sub
   If MESANO$(TXT$, DES%, HAS%) = TXT$ Then
     If XVALO(CONTROL("", "CIEMEVEN")) >= HAS% Then
       Image2.Visible = False
       Image1.Visible = True
     End If
   End If
End Sub

Sub CONASIDIA()
    '
    MESA$ = TRIM$(MESAN)
    MESB$ = MESANO$(MESA$, DES%, HAS%)
    If MESB$ <> MESA$ Then
      Exit Sub
    End If
    '
    Screen.MousePointer = 11
    J& = 0
    Call BLANARCH("!ASIDIAFI")
    TTXXY$ = Chr$(1) + Chr$(1) + MKI$(HAS%)
    For IL& = 1 To ULTREG&("ASIDIAR")
        x$ = REGLEIDO$("ASIDIAR", IL&)
        If Left$(x$, 4) = TTXXY$ Then
            CUENU& = CVI(Mid$(x$, 5, 2))
            IDEB# = CVD(Mid$(x$, 39, 8))
            If Abs(IDEB#) >= 0.005 Then
                IHAB# = 0
                If IDEB# < 0 Then
                    IHAB# = Abs(IDEB#)
                    IDEB# = 0
                End If
                CUES$ = Space$(12)
                If CUENU& > 0 Then
                    If CUENU& <= ULTREG&("CUENTAS") Then
                        CUES$ = Left$(REGLEIDO$("CUENTAS", CUENU&), 12)
                    End If
                End If
                DENOCUES$ = ECOARCH$("ECUECON", CUES$)
                '
                Z$ = REGBLAN$("ASIDIAFI")
                Call REPLA(Z$, CUES$, 1, 12)
                Call REPLA(Z$, DENOCUES$, 13, 30)
                Call REPLA(Z$, MKD$(IDEB#), 43, 8)
                Call REPLA(Z$, MKD$(IHAB#), 51, 8)
                '
                J& = J& + 1
                Call GRAREG("!ASIDIAFI", Z$, J&)
            End If
        End If
    Next IL&
    Call SETULTREG("!ASIDIAFI", J&)
    '
    For K1& = 1 To J&
        For K2& = K1& + 1 To J&
            X1$ = REGLEIDO$("!ASIDIAFI", K1&)
            X2$ = REGLEIDO$("!ASIDIAFI", K2&)
            If Left$(X1$, 12) > Left$(X2$, 12) Then
                X0$ = X1$
                X1$ = X2$
                X2$ = X0$
                Call GRAREG("!ASIDIAFI", X1$, K1&)
                Call GRAREG("!ASIDIAFI", X2$, K2&)
            End If
        Next K2&
    Next K1&
    Call CIERRARCH("!ASIDIAFI")
    '
    Screen.MousePointer = 1
    VTB% = VENTABU%("ASIDIAFI")
    '
End Sub
'
Function VALIMESAN%()
   VMA% = 0
   CORRESP$ = MESANO$(MESAN, DESDE%, HASTA%)
   If CORRESP$ <> "" Then
       VMA% = 1
       DESDAT = Int(CDbl(CVDAT(DESDE%)))
       HASDAT = Int(CDbl(CVDAT(HASTA%)))
     Else
       Call MENSERR(24, "Falta Elegir Período Mensual")
   End If
   VALIMESAN% = VMA%
End Function

Sub GENSUDIVEN()
    '
    Screen.MousePointer = 11
    '
    Dim IMASI#(32767)
    CUEMX% = 0
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
'    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
'    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
'    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUD') "
'    SQLX$ = SQLX$ + "Order by FechConta ASC, NumeComp ASC "
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where FechConta >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND FechConta <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "Order by FechConta ASC, NumeComp ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '\\\OT-05-0416-WAR-23/06/05///
    SQLX$ = "SELECT * FROM CAJABANC "
    'SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    'SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "where FechConta >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND FechConta <= " & HASDAT & " "
    '
    Set CABAMENS = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '\\\OT-05-0416-WAR-FIN///
    JJ& = 0: REBLA$ = REGBLAN$("SUDIVEN")
    HUBORECU% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      CANRE& = .RecordCount
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      .MoveFirst
      UKL& = 0
      Do While Not .EOF
        UKL& = UKL& + 1
        Call TRACE(20, UKL&, CANRE&)
        ZZ$ = REBLA$
        NUMEFA$ = !CODICOM_LAR
        NCLIE% = !NuClien
        FECHAD = Int(CDbl(!FechConta))
        ITOTA# = !ValAbsComp
        SIG% = Sgn(!IDEBECOMP - !IHABECOMP)
        MOBASI$ = !TipoMovim
        SUMABAL# = 0      ' SUMADOR PARA VERIFICAR BALANCE DEL COMPROBANTE
        '
        If Abs(!IDEBECOMP - ROUND#(!IDEBECOMP)) > 0.000001 Or Abs(!IHABECOMP - ROUND#(!IHABECOMP)) > 0.000001 Then
           HUBORECU% = 1
           Call RECUMOVEN(!CODICOM_LAR, !FechConta)
        End If
        
        If MOBASI$ = "FVEN" Then
          Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 1, 2)
          Call REPLA(ZZ$, !CODICOM_LAR, 3, 18)
          Call REPLA(ZZ$, RASOCLI$(NCLIE%), 21, 30)
          Call REPLA(ZZ$, CUITCLI$(NCLIE%), 51, 16)
          Call REPLA(ZZ$, MKD$(ITOTA# * SIG%), 67, 8)
          Call REPLA(ZZ$, CODCUE$(!CueContaI), 75, 12)
          Call REPLA(ZZ$, MKD$(!IDEBECOMP), 87, 8)
          Call REPLA(ZZ$, MKD$(!IHABECOMP), 95, 8)
          JJ& = JJ& + 1
          Call GRAREG("SUDIVEN", ZZ$, JJ&)
          SUMABAL# = SUMABAL# + !IDEBECOMP - !IHABECOMP
        End If
        '
        CUEII% = !CueContaI
        If CUEII% > 0 Then
          IMASI#(CUEII%) = IMASI#(CUEII%) + (!IDEBECOMP - !IHABECOMP)
          If CUEII% > CUEMX% Then CUEMX% = CUEII%
        End If
        '
        'SQLX$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
        SQLX$ = "FechConta = " & FECHAD & " "
        SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX$ = SQLX$ + "AND (NuClien = " & NCLIE% & " OR (NuClien = 0 AND NuProve = 0)) "
        SQLX$ = SQLX$ + "AND TipoMovim <> '" & MOBASI$ & "' "
        '
        With CABAMENS 'CajaBanc
          .FindFirst SQLX$
25        If .NoMatch = False Then
            ZZ$ = REBLA$
            '
            If Abs(!IDEBECOMP - ROUND#(!IDEBECOMP)) > 0.000001 Or Abs(!IHABECOMP - ROUND#(!IHABECOMP)) > 0.000001 Then
               HUBORECU% = 1
               Call RECUMOVEN(!CODICOM_LAR, !FechConta)
            End If
            '
            If MOBASI$ = "FVEN" Then
              Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 1, 2)
              Call REPLA(ZZ$, !CODICOM_LAR, 3, 18)
              Call REPLA(ZZ$, RASOCLI$(NCLIE%), 21, 30)
              Call REPLA(ZZ$, CUITCLI$(NCLIE%), 51, 16)
              Call REPLA(ZZ$, CODCUE$(!CueContaI), 75, 12)
              Call REPLA(ZZ$, MKD$(!IDEBECOMP), 87, 8)
              Call REPLA(ZZ$, MKD$(!IHABECOMP), 95, 8)
              JJ& = JJ& + 1
              Call GRAREG("SUDIVEN", ZZ$, JJ&)
              SUMABAL# = SUMABAL# + !IDEBECOMP - !IHABECOMP
            End If
            '
            CUEII% = !CueContaI
            If CUEII% > 0 Then
              IMASI#(CUEII%) = IMASI#(CUEII%) + (!IDEBECOMP - !IHABECOMP)
              If CUEII% > CUEMX% Then CUEMX% = CUEII%
            End If
            '
           .FindNext SQLX$
            GoTo 25
          End If
        End With
        '
        ' SUMADOR PARA VERIFICAR EL BALANCE DEL COMPROBANTE
        If Abs(SUMABAL#) >= 0.005 Then
           HUBORECU% = 1
           Call RECUMOVEN(!CODICOM_LAR, !FechConta)
        End If
        '
      .MoveNext
      Loop
    End With
    Call SETULTREG("SUDIVEN", JJ&)
    '
    Call HEADERS("SUDIVEN", "")
    Call HEADERS("SUDIVEN", "Corresponde a: " + TRIM$(CORRESP$))
    '
    TTXX$ = Chr$(1) + Chr$(1) + MKI$(HASTA%)
    V& = 0: FIN& = ULTREG&("ASIDIAR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      x$ = REGLEIDO$("ASIDIAR", U&)
      If Left$(x$, 4) <> TTXX$ Then
        V& = V& + 1
        Call GRAREG("ASIDIAR", x$, V&)
      End If
    Next U&
    '
    AGRECU% = 0
    For k% = 1 To CUEMX%
      If Abs(IMASI#(k%)) >= 0.005 Then
        x$ = REGBLAN$("ASIDIAR")
        Call REPLA(x$, TTXX$, 1, 4)
        Call REPLA(x$, MKI$(k%), 5, 2)
        Call REPLA(x$, MKD$(IMASI#(k%)), 39, 8)
        V& = V& + 1
        Call GRAREG("ASIDIAR", x$, V&)
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
    Call SETULTREG("ASIDIAR", V&)
    '
    If AGRECU% > 0 Then
      Call FILESORT("CUECON", "", 1, 1, "ASC")
    End If
    Call CIERRARCH("CUECON")
    '
    If HUBORECU% > 0 Then
      Call COMUNI("Errores de Generación de Subdiario de Ventas.\La Información Generada no es Confiable.\   \Vuelva a Pedir el Listado.")
      Call SETULTREG("SUDIVEN", 0)
    End If
    '
99  Screen.MousePointer = 1
    '
End Sub

Sub RECUMOVEN(NUMEFA$, FECONTAFA)
   '
   FECONTAD = Int(CDbl(FECONTAFA))
   
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where CodiCom_Lar = '" & NUMEFA$ & "' "
   SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) = " & FECONTAD & " "
   SQLX$ = SQLX$ + "Order by ABS(IDebeComp - IHabeComp) ASC "
   '
   Set MOFASEL = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   SUMACOMP# = 0
   On Error Resume Next
   With MOFASEL
     .MoveFirst
     If Err > 0 Then
       On Error GoTo 0
       Exit Sub
     End If
     Do While Not .EOF
       IDEBAS# = ROUND#(!IDEBECOMP, 2)
       IHABAS# = ROUND#(!IHABECOMP, 2)
       '
       .Edit
       !IDEBECOMP = IDEBAS#
       !IHABECOMP = IHABAS#
       !ValAbsComp = Abs(IDEBAS# - IHABAS#)
       .Update
       SUMACOMP# = SUMACOMP# + IDEBAS# - IHABAS#
     .MoveNext
     Loop
   End With
   On Error GoTo 0
   '
   If Abs(SUMACOMP#) >= 0.005 Then
     If Abs(SUMACOMP#) < 0.1 Then
       TTXX$ = "Diferencia de " + TRIM$(FORMATNUM$(SUMACOMP#, "F12.2")) + " en Comprobante " + TRIM$(NUMEFA$)
       TTXX$ = TTXX$ + "\   \Desea Recuperar Esta Diferencia Automáticamente ? \\Si, Recuperar\No, Dejar como está"
       If COMALTER%(TTXX$) = 1 Then
         MOFASEL.MoveLast
         MOFASEL.Edit
         If MOFASEL!IDEBECOMP >= 0.005 Then
             MOFASEL!IDEBECOMP = MOFASEL!IDEBECOMP - SUMACOMP#
           Else
             MOFASEL!IHABECOMP = MOFASEL!IHABECOMP + SUMACOMP#
         End If
         MOFASEL.Update
       End If
     End If
   End If
   '
End Sub

Sub GENIVAVEN()
    '
    Screen.MousePointer = 11
    '
    Dim RELIVA$(32)
    Dim IMASI#(32767)
    Dim NUFAK$(), TIMOK$(), VAMOK$(), IPAXK#(), BASEK#(), CUECK%(), OPMOK$()
    UBA& = 1024
    ReDim Preserve NUFAK$(UBA&), TIMOK$(UBA&), VAMOK$(UBA&), IPAXK#(UBA&), BASEK#(UBA&), CUECK%(UBA&), OPMOK$(UBA&)
    CUEMX% = 0
    '
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where FechEmisi >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND FechEmisi <= " & HASDAT & " "
    SQLX$ = SQLX$ + "Order by FechEmisi ASC, NumeComp ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0: JPER& = 0
    REBLA$ = REGBLAN$("LIVAVEN")
    REBLB$ = REGBLAN$("REIVAVEN")
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 50
      End If
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      U& = 0: JJU& = 0: VECOMP$ = ""
      '
      Do While Not .EOF
        U& = U& + 1
        Call TRACE(20, U&, FIN&)
        '
        INCLU$ = ""
        On Error Resume Next
        INCLU$ = TRIM$(UCase$(!IncSubIva))
        On Error GoTo 0
        If INCLU$ = "NO" Then GoTo 19
        '
        TIMO$ = TRIM$(UCase$(!TipoMovim))
        NUMEFA$ = TRIM$(!CODICOM_LAR)
        NCLIE% = !NuClien
        FECHAD = Int(CDbl(!FechConta))
        ITOTA# = !ValAbsComp
        SIG% = Sgn(!IDEBECOMP - !IHABECOMP)
        If TIMO$ = "FVEN" Then
              ZZ$ = REBLA$
              VECOMP$ = VECOMP$ + AJUSTI$(NUMEFA$, 18)
              Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 3, 2)
              Call REPLA(ZZ$, MKI$(CVINT(!FechEmisi)), 5, 2)
              Call REPLA(ZZ$, MKS$(!NumeComp), 7, 4)
              Call REPLA(ZZ$, NUMEFA$, 11, 18)
              Call REPLA(ZZ$, MKI$(NCLIE%), 31, 2)
              If NCLIE% = 0 And TRIM$(UCase$(!OPCIMOVIM)) = "ANUL" Then 'si esta anulado
                  Call REPLA(ZZ$, "COMPROBANTE ANULADO", 33, 30)
              Else
                  Call REPLA(ZZ$, RASOCLI$(NCLIE%), 33, 30)
              End If
              '
              Call REPLA(ZZ$, CUITCLI$(NCLIE%), 63, 16)
              If NCLIE% = 0 And TRIM$(UCase$(!OPCIMOVIM)) = "ANUL" Then 'si esta anulado
                  Call REPLA(ZZ$, "", 79, 12)
              Else
                  Call REPLA(ZZ$, POSIVAC$(NCLIE%), 79, 12)
              End If
              '
              If Abs(ITOTA#) < 0.005 Then 'si esta anulado
                  Call REPLA(ZZ$, "COMPROBANTE ANULADO", 33, 58)
              End If
              '
              Call REPLA(ZZ$, MKD$(ITOTA# * SIG%), 91, 8)
              '
              '\\\OT-05-0709-WAR-02/11/05///
              Pi% = PIVACLI%(NCLIE%)
              If Pi% = 5 Or Pi% = 7 Then
                Call REPLA(ZZ$, MKD$(ITOTA# * SIG%), 185, 8)
              End If
              '\\\OT-05-0709-WAR-FIN///
              If TRIM$(Mid$(ZZ$, 129, 8)) = "" Then
                Call REPLA(ZZ$, "NO GRAV.", 129, 8)
              End If
              '
              JJU& = JJU& + 1
              Call GRAREG("!LIVAVEN", ZZ$, JJU&)
            ElseIf InStr("IVA_GAN_IBRU", TIMO$) > 0 Then
              OPMO$ = TRIM$(!OPCIMOVIM)
              VAMO$ = TRIM$(!VARIMOVIM)
              IPAX# = Abs(ITOTA#)
              BASE# = 0
              On Error Resume Next
              BASE# = Abs(!BASEIMPO)
              On Error GoTo 0
              If UCase$(OPMO$) = "DEFI" Then
                If UCase$(VAMO$) = "RNI" Then
                  BASE# = 0
                End If
              End If
              '
              If Abs(IPAX#) >= 0.005 Or BASE# >= 0.005 Then
                JK& = JK& + 1
                If JK& > UBound(NUFAK$) Then
                  UBA& = 1024 + UBound(NUFAK$)
                  ReDim Preserve NUFAK$(UBA&), TIMOK$(UBA&), VAMOK$(UBA&), IPAXK#(UBA&), BASEK#(UBA&), CUECK%(UBA&), OPMOK$(UBA&)
                End If
                NUFAK$(JK&) = NUMEFA$
                TIMOK$(JK&) = TIMO$
                VAMOK$(JK&) = VAMO$
                IPAXK#(JK&) = IPAX#
                BASEK#(JK&) = BASE#
                CUECK%(JK&) = !CueContaI
                OPMOK$(JK&) = OPMO$
              End If
        End If
19    .MoveNext
      Loop
    End With
    '
    For U2& = 1 To JJU&
       Call TRACE(20, U2&, JJU&)
       RELIVA$(0) = REGLEIDO$("!LIVAVEN", U2&)
       NUMEFA$ = TRIM$(Mid$(RELIVA$(0), 11, 18))
       If InStr(VECOMP$, NUMEFA$) < 1 Then GoTo 49
       '
       IREL$ = Left$(RELIVA$(0), 10)
       For KU% = 1 To 32
         RELIVA$(KU%) = REBLA$
         Mid$(RELIVA$(KU%), 1, 10) = IREL$
       Next KU%
       '
       SIG% = Sgn(CVD(Mid$(RELIVA$(0), 91, 8)))
       ULIVA% = 0: PROIVA% = 0: HAYIVA% = 0
       '
       FECHCO% = CVI(Mid$(RELIVA$(0), 3, 2))
       NCLIE% = CVI(Mid$(RELIVA$(0), 31, 2))
       RACLI$ = Mid$(RELIVA$(0), 33, 30)
       CTTC$ = Mid$(RELIVA$(0), 63, 16)
       ITOTACO# = CVD(Mid$(RELIVA$(0), 91, 8))
       '
       INOGRAD# = ITOTACO#
       For JK1& = 1 To JK&
          If NUFAK$(JK1&) = NUMEFA$ Then
             TIMO$ = TIMOK$(JK1&)
             VAMO$ = VAMOK$(JK1&)
             IPAX# = IPAXK#(JK1&)
             BASE# = BASEK#(JK1&)
             CUECO% = CUECK%(JK1&)
             OPMO$ = OPMOK$(JK1&)
             '
             If TIMO$ = "IVA" And UCase$(VAMO$) = "PER-IVA" Then
                  ZZ$ = RELIVA$(0)
                  Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 161, 8)
                  RELIVA$(0) = ZZ$
                ElseIf (TIMO$ = "GAN" Or TIMO$ = "IVA") And UCase$(VAMO$) = "PER-GAN" Then
                  ZZ$ = RELIVA$(0)
                  Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 169, 8)
                  RELIVA$(0) = ZZ$
                ElseIf TIMO$ = "IBRU" And UCase$(VAMO$) = "PERC" Then
                  ZZ$ = RELIVA$(0)
                  Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 193, 8)
                  Call REPLA(ZZ$, MKI$(CUECO%), 121, 2)
                    JURIB$ = TRIM$(Mid$(OPMO$, 4, 1))
                  Call REPLA(ZZ$, JURIB$, 203, 1)
                  Call REPLA(ZZ$, ECOARCH$("PROVINC", JURIB$), 205, 24)
                  RELIVA$(0) = ZZ$
                  '
                  If Abs(IPAX#) >= 0.005 Then
                    ZPER$ = REGBLAN$("SUPERIB")
                    Call REPLA(ZPER$, MKI$(FECHCO%), 1, 2)
                    Call REPLA(ZPER$, NUMEFA$, 3, 18)
                    Call REPLA(ZPER$, RACLI$, 21, 30)
                    Call REPLA(ZPER$, NIBRCLI$(NCLIE%), 65, 14)
                    Call REPLA(ZPER$, CTTC$, 51, 14)
                    Call REPLA(ZPER$, MKD$(BASE# * SIG%), 79, 8)
                    Call REPLA(ZPER$, MKD$(IPAX# * SIG%), 87, 8)
                    JPER& = JPER& + 1
                    Call GRAREG("SUPERIB", ZPER$, JPER&)
                  End If
                  '
                ElseIf TIMO$ = "IVA" Then
                  ZZ$ = RELIVA$(PROIVA%)
                    If TRIM$(VAMO$) = "" Then
                       VAMO$ = "NO GRAV."
                    End If
                  Call REPLA(ZZ$, VAMO$, 129, 8)
                  Call REPLA(ZZ$, MKD$(BASE# * SIG%), 137, 8)
                  Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 145, 8)
                     If Abs(BASE#) >= 0.005 Then   ' PARA CONSIDERAR CASOS DE SOLO IVA
                        VAMO5$ = REPLACAR$(VAMO$, "%", " ")
                        If XVALO(VAMO5$) >= 0.1 Then
                          BASECALC# = 100 * IPAX# / (XVALO(VAMO5$))
                          INOGRAD1# = BASE# - BASECALC#
                          If Abs(INOGRAD1#) >= 1 Then
                               Call REPLA(ZZ$, MKD$(BASECALC# * SIG%), 137, 8)
                             Else
                               INOGRAD1# = 0
                          End If
                        End If
                     End If
                  BASE2# = SIG% * CVD(Mid$(ZZ$, 137, 8))
                  If Abs(BASE2#) > 0.005 Then
                    Call REPLA(ZZ$, MKD$(100 * IPAX# / BASE2#), 153, 8)
                  End If
                  RELIVA$(PROIVA%) = ZZ$
                  ULIVA% = PROIVA%
                  PROIVA% = PROIVA% + 1
                  HAYIVA% = 1
             End If
          End If
       Next JK1&
       '
       ' LO QUE SIGUE ES PARA OBTENER EL IMPORTE NO GRAVADO POR DIFERENCIA
       For KKI% = 0 To ULIVA%
          INOGRAD# = INOGRAD# - ROUND#(CVD(Mid$(RELIVA$(KKI%), 137, 8)))
          INOGRAD# = INOGRAD# - ROUND#(CVD(Mid$(RELIVA$(KKI%), 145, 8)))
          INOGRAD# = INOGRAD# - ROUND#(CVD(Mid$(RELIVA$(KKI%), 161, 8)))
          INOGRAD# = INOGRAD# - ROUND#(CVD(Mid$(RELIVA$(KKI%), 169, 8)))
          INOGRAD# = INOGRAD# - ROUND#(CVD(Mid$(RELIVA$(KKI%), 177, 8)))
          INOGRAD# = INOGRAD# - ROUND#(CVD(Mid$(RELIVA$(KKI%), 185, 8)))
          INOGRAD# = INOGRAD# - ROUND#(CVD(Mid$(RELIVA$(KKI%), 193, 8)))
       Next KKI%
       '
       If Abs(INOGRAD#) >= 0.005 Then
          ZZ$ = RELIVA$(PROIVA%)
             VAMO$ = "NO GRAV."
             INOGRAD# = INOGRAD# + CVD(Mid$(ZZ$, 137, 8))
          Call REPLA(ZZ$, VAMO$, 129, 8)
          Call REPLA(ZZ$, MKD$(INOGRAD#), 137, 8)
          Call REPLA(ZZ$, MKD$(0), 145, 8)
          RELIVA$(PROIVA%) = ZZ$
          ULIVA% = PROIVA%
          PROIVA% = PROIVA% + 1
          HAYIVA% = 1
       End If
       '
       For KU% = 0 To ULIVA%
         JJ& = JJ& + 1
         Call GRAREG("LIVAVEN", RELIVA$(KU%), JJ&)
       Next KU%
       '
49  Next U2&
    '
50  Call SETULTREG("LIVAVEN", JJ&)
    Call SETULTREG("SUPERIB", JPER&)
    '
    Call HEADERS("LIVAVEN", "")
    Call HEADERS("LIVAVEN", "Corresponde a: " + TRIM$(CORRESP$))
    '
    ' CARGA RESUMEN MENSUAL DE IVA - VENTAS
    Dim DECATIV$(128)
    CACATIV% = 0
    Dim VECREIVA$(256)
    UVECRE% = 0
    '
    MODOTANIT% = 0
    If EXISTE%("LIVENTAN.PRF") > 0 Then
      MODOTANIT% = 1
    End If
    '
    KUT& = 0
    For KUL& = 1 To JJ&
      Call TRACE(20, KUL&, JJ&)
      RECUE$ = REGLEIDO$("LIVAVEN", KUL&)
      If TRIM$(Mid$(RECUE$, 11, 18)) <> "" Then
        NUCO$ = TRIM$(Mid$(RECUE$, 11, 18))
        SIG% = Sgn(CVD(Mid$(RECUE$, 91, 8)))
      End If
      '
      CORUBRO% = 1: DERUBRO$ = "VENTA" '\\\OT-06-0079-WAR-03/03/06///
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
      If TRIM$(Mid$(RECUE$, 129, 8)) <> "" Then
         DECAT1$ = TRIM$(Mid$(RECUE$, 129, 8))
      End If
      If TRIM$(Mid$(RECUE$, 79, 12)) <> "" Then
         DECAT2$ = Mid$(RECUE$, 79, 12)
      End If
      DECATE$ = DECAT1$ + " - " + DECAT2$
      For KLI% = 1 To CACATIV%
        If DECATE$ = DECATIV$(KLI%) Then
          COCATE% = KLI%
          GoTo 20
        End If
      Next KLI%
      CACATIV% = CACATIV% + 1
      COCATE% = CACATIV%
      DECATIV$(CACATIV%) = DECATE$
      '
20    BASEIMPO# = CVD(Mid$(RECUE$, 137, 8)) + CVD(Mid$(RECUE$, 185, 8))
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
      ' PARA TANIT SEPARAR IMPORTE NO GRAVADO
      If MODOTANIT% = 1 Then
        RECUE1$ = Left$(RECUE$, 152) + String$(8, 0) + Mid$(RECUE$, 153)
        TANOMIX$ = TRIM$(REPLACAR$(Mid$(RECUE1$, 128, 8), "%", " "))
        TANOMI = XVALO(TANOMIX$)
        '
        ITONE# = CVD(Mid$(RECUE1$, 137, 8)) + CVD(Mid$(RECUE1$, 193, 8))
        INEGRA# = ITONE#
        IEXEN# = 0: IEXEN1# = 0: IEXEN2# = 0
        IMIVA# = CVD(Mid$(RECUE1$, 145, 8))
        If Abs(IMIVA#) < 0.005 Then
           IEXEN# = INEGRA#
           INEGRA# = 0
           GoTo 95
        End If
        If Abs(INEGRA# * TANOMI / 100 - IMIVA#) > 0.5 Then
           If Abs(INEGRA#) >= 0.005 Then
             If TANOMI > 0 Then
               INEGRA# = 100 * IMIVA# / TANOMI
               IEXEN# = ITONE# - INEGRA#
             End If
           End If
        End If
95      If UCase$(Mid$(RECUE1$, 11, 1)) = "F" Then
            IEXEN1# = IEXEN#
          Else
            IEXEN2# = IEXEN#
        End If
        '
'If TANOMIX$ = "RNI" Then
'AAA = BBB
'End If
        '
        If Abs(INEGRA#) < 0.005 Then
          If KUT& > 0 Then
            RECUE2$ = REGLEIDO$("LIVENTAN", KUT&)
            If Mid$(RECUE1$, 11, 80) = Mid$(REBLA$, 11, 80) Then
              'If Abs(CVD(Mid$(RECUE2$, 153, 8))) < 0.005 Then
                'If Abs(CVD(Mid$(RECUE2$, 161, 8))) < 0.005 Then
              IEXA0# = CVD(Mid$(RECUE2$, 145, 8))
              IEXA1# = CVD(Mid$(RECUE2$, 153, 8))
              IEXA2# = CVD(Mid$(RECUE2$, 161, 8))
              Call REPLA(RECUE2$, MKD$(IEXA0# + IMIVA#), 145, 8)
              Call REPLA(RECUE2$, MKD$(IEXA1# + IEXEN1#), 153, 8)
              Call REPLA(RECUE2$, MKD$(IEXA2# + IEXEN2#), 161, 8)
              Call GRAREG("LIVENTAN", RECUE2$, KUT&)
              GoTo 96
                'End If
              'End If
            End If
          End If
        End If
        '
        KUT& = KUT& + 1
        Call REPLA(RECUE1$, MKD$(INEGRA#), 137, 8)
        Call REPLA(RECUE1$, MKD$(IEXEN1#), 153, 8)
        Call REPLA(RECUE1$, MKD$(IEXEN2#), 161, 8)
        Call GRAREG("LIVENTAN", RECUE1$, KUT&)
      End If
      ' FIN PARA TANIT
      '
96  Next KUL&
    '
    If MODOTANIT% = 1 Then
      Call SETULTREG("LIVENTAN", KUT&)
    End If
    '
    JJK& = 0
    For UI1% = 1 To UVECRE%
      JJK& = JJK& + 1
      Call GRAREG("REIVAVEN", VECREIVA$(UI1%), JJK&)
    Next UI1%
    Call SETULTREG("REIVAVEN", JJK&)
    Call CIERRARCH("REIVAVEN")
    '
    'SE AGREGAN ESTAS 2 LINEAS PARA PASAR EL HEADERS A LIVENTAN
    Call HEADERS("LIVENTAN", "")
    Call HEADERS("LIVENTAN", "Corresponde a: " + TRIM$(CORRESP$))
    '
    Call HEADERS("REIVAVEN", "")
    Call HEADERS("REIVAVEN", "Corresponde a: " + TRIM$(CORRESP$))
    '
99  Screen.MousePointer = 1
    '
End Sub

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
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "AND NuClien = " & NC% & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "Order by CODICOM_LAR ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NCLIE% = !NuClien
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!CODICOM_LAR, 13, 6))
            If Len(!CODICOM_LAR) < 18 Then
              NPX& = XVALO(Right$(!CODICOM_LAR, 6))
            End If
            Y$ = MKS$(NPX&) + AJUSTI$(!CODICOM_LAR, 18) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUECO", Y$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("!SELCUECO", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUECO") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no Registra\Facturas de Venta que Puedan Seleccionarse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUECO")
    Call FRESHECHO("!SELCUECO")
    '
    Call SELECHO("!SELCUECO")
    If VALACT1$("!SELCUECO") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUECO"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = Left$(VALACT2$("!SELCUECO"), 18)
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    '
    COINC = 0
    Call APERBASDAT("CABAN")
    TTXX$ = "Libro de I.V.A. Ventas\  \Comprobante '" + TRIM$(DOCULA$) + "' \de Cuenta '" + TRIM$(RASOCLI$(NC%)) + "' ?"
    If COMALTER%(TTXX$ + "\\Incluye\Excluye de Listado") = 1 Then
      COINC = 1
    End If
'    With CajaBanc
'      .FindFirst "CodiCom_Lar = '" & DOCULA$ & "' AND NuClien = " & NC% & " "
'20     If .NoMatch = False Then
'        .Edit
'         If COINC = 1 Then
'           !IncSubIva = ""
'         Else
'           !IncSubIva = "NO"
'         End If
'        .Update
'        .FindNext "CodiCom_Lar = '" & DOCULA$ & "' AND NuClien = " & NC% & ""
'        GoTo 20
'      End If
'    End With
    '
    SQLX$ = "SELECT IncSubIva FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE CodiCom_Lar = '" & DOCULA$ & "' AND NuClien = " & NC%
    If COINC > 0 Then
       SQLX$ = SQLX$ + " AND (IncSubIva = 'NO' OR IncSubIva = '')"
    Else
       SQLX$ = SQLX$ + " AND IncSubIva = ''  "
    End If
    Set RS = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    With RS
        Do While Not .EOF
            .Edit
             If COINC = 1 Then
               !IncSubIva = ""
             Else
               !IncSubIva = "NO"
             End If
            .Update
             .MoveNext
        Loop
    End With

    GoTo 10
    '
End Sub

Sub ACOSREPVIE()
     '
     '//OT-07-0458//

     'RECORRE TODA LA TABLA STOCKS CARGANDO EL COSTO ACTUAL DEL
     'MAESTRO DE CADA ARTICULO EN EL CAMPO "CostoRep"
     'SI EL CODIGO ES 0 (material no codificado)
     'LE VA A GRABAR UN 0 EN EL CAMPO "CostoRep"
     '
     Call APERBASDAT("STOCKS")
     SQLX$ = "SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE ISNULL(CostoRep) = TRUE OR ISEMPTY(CostoRep) = true "
     SQLX$ = SQLX$ + "ORDER BY FechMov ;"
     '
     Set RECO = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
     With RECO
       On Error Resume Next
       .MoveLast
       If Err < 1 Then
          On Error GoTo 0
          FIN& = .RecordCount
          If FIN& > 1000 Then
             Call COMUNI("Actualizando los Costos de Reposición.\Este Proceso Puede Demandar Varios Minutos.\   \Pulse 'Aceptar' para Continuar.")
          End If
         .MoveFirst
         For i& = 1 To FIN&
            Call TRACE(20, i&, FIN&)
            CI% = !Cod_Inte
            .Edit
            If CI% = 0 Then
                !CostoRep = 0
              Else
                !CostoRep = COSUNI(CI%)
            End If
            .Update
            .MoveNext
         Next i&
       End If
     End With
     '
     '//OT-07-0458//
End Sub

Sub LISTADEBCRE()
'
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%) + MKI$(0)
    '
    OBX$ = OBJPLA$("DESHASFECHCLI", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    NCLI% = CVI(Mid$(OBX$, 5, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIODO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIODO$ = PERIODO$ + " - " + FECHATEX$(FEX)
    If PERIODO$ = "" Then Exit Sub
    DESDEX = CDbl(CVDAT(DES%))
    HASTAX = CDbl(CVDAT(HAS%))

    
    '
    Call SELECHO("IMPDECRE")
    MOTIVO$ = TRIM$(VALACT1$("IMPDECRE"))
    If MOTIVO$ = "" Then Exit Sub
    REGISTRO$ = FILTERGETRECORD$("IMPDECRE", 1, MKI$(XVALO(MOTIVO$)))
    CCONTAB% = CVI(Mid$(REGISTRO$, 43, 2))
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC"
    SQLX$ = SQLX$ + " WHERE CueContaI=" + Str(CCONTAB%)
    If NCLI% <> 0 Then
        SQLX$ = SQLX$ + " AND NuClien = " & NCLI% & " "
    End If
    SQLX$ = SQLX$ + " AND FechConta >= " & DESDEX & " "
    SQLX$ = SQLX$ + " AND FechConta <= " & HASTAX & " "
    SQLX$ = SQLX$ + " AND (OpciMovim='ND' OR OpciMovim='NC')"
    SQLX$ = SQLX$ + " AND TipoMovim='VENTA'"
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        Call MENSERR(24, "No Existen Registros Para las Opciones Seleccionadas")
        GoTo 99
      End If
         On Error GoTo 0
         Do While Not .EOF
            FECHA% = CVINT%(!FechConta)
            CLIENTE% = !NuClien
            COMPROBANTE$ = !CODICOM_LAR
            '
            XX$ = REGBLAN("LISDECRE")
            Call REPLA(XX$, MKI$(FECHA%), 1, 2)
            Call REPLA(XX$, MKI$(CLIENTE%), 3, 2)
            Call REPLA(XX$, COMPROBANTE$, 5, 18)
            CajaBanc.FindFirst "(CodiCom_Lar = '" + COMPROBANTE$ + "' AND TipoMovim = 'FVEN')"
            On Error Resume Next
            IMPORTE# = CajaBanc!ValAbsComp ' MUESTRA EL IMPORTE FINAL DEL COMPROBANTE
            On Error GoTo 0
            Call REPLA(XX$, MKD$(IMPORTE#), 23, 8)
            JJ& = JJ& + 1
            Call GRAREG("LISDECRE", XX$, JJ&)
            '
         .MoveNext
         Loop
         '
         Call SETULTREG("LISDECRE", JJ&)
         '
         Call FILESORT("LISDECRE", "", 2, 2, "ASC")
         Call FILESORT("LISDECRE", "", 1, 1, "ASC")
         '
         Call HEADERS("LISDECRE", "")
         Call HEADERS("LISDECRE", "Periodo: " + PERIODO$)
         Call HEADERS("LISDECRE", "Motivo: " & VALACT2$("IMPDECRE"))
         '
         Call FINAL("*LISDECRE")
         '
      'End If
      On Error GoTo 0
99  End With
    '
End Sub

Private Sub notadebito_Click()
   Call Command4_Click
End Sub

Private Sub notascredito_Click()
   Call Command5_Click
End Sub

Private Sub rankingporproveedorporrangodefechas_Click()
End Sub

Private Sub pantalla_Click()
     Option1.Value = True
     Call Command7_Click
End Sub

Private Sub pantalla0_Click()
   Option1.Value = True
   Call Command5_Click
End Sub

Private Sub pantalla1_Click()
     Option1.Value = True
     Call Command4_Click

End Sub


Private Sub pantalla2_Click()
     Option1.Value = True
     Call Command2_Click

End Sub

Private Sub pantalla4_Click()
     Option1.Value = True
     Call Command3_Click

End Sub

Private Sub pantalla5_Click()
    Option1.Value = True
     Call Command6_Click
End Sub

Private Sub pantalla6_Click()
   Option1.Value = True
   Call Command19_Click
End Sub

Private Sub pantalla7_Click()
    Option1.Value = True
    Call Command30_Click
End Sub

Private Sub pantalla8_Click()
   Option1.Value = True
   Call Command11_Click
End Sub

Private Sub reabrirmes_Click()
   Call Image1_Click
End Sub

Private Sub reimputaciondecomprobantesdeventas_Click()
    Call Command22_Click
End Sub

Private Sub resumenmensualivaventas_Click()
    Call Command13_Click
End Sub

Private Sub resumenprovinciaygrupos_Click()
   Call Command10_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub subdiariodeventas_Click()
   Call Command16_Click
End Sub

Private Sub subiarioivaventasporimpresora_Click()
   Call Command21_Click
End Sub

Private Sub varios_Click()
    Call Command8_Click
End Sub
