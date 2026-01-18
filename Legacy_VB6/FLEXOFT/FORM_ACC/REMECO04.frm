VERSION 5.00
Begin VB.Form REMECO04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Gestión - Resúmenes Mensuales de Compras"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture11 
      Height          =   5055
      Left            =   120
      ScaleHeight     =   4995
      ScaleWidth      =   9315
      TabIndex        =   62
      Top             =   1080
      Visible         =   0   'False
      Width           =   9375
      Begin VB.Image Image6 
         Height          =   5025
         Left            =   0
         Picture         =   "REMECO04.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9300
      End
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
         Picture         =   "REMECO04.frx":10205
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
      BackColor       =   &H00C0E0FF&
      Caption         =   "Libros de Compras"
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
      TabIndex        =   41
      Top             =   3720
      Width           =   4635
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2940
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
            Height          =   600
            Left            =   0
            Picture         =   "REMECO04.frx":1034F
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
            Height          =   600
            Left            =   1840
            Picture         =   "REMECO04.frx":10659
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
            Picture         =   "REMECO04.frx":10963
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
            Height          =   600
            Left            =   2400
            Picture         =   "REMECO04.frx":10C6D
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
            TabIndex        =   44
            Top             =   105
            Width           =   1485
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Gestión de Compras"
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
      TabIndex        =   31
      Top             =   1200
      Width           =   4635
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1470
         ScaleHeight     =   585
         ScaleWidth      =   2265
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
            Height          =   600
            Left            =   0
            Picture         =   "REMECO04.frx":10F77
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Facturación de Compras"
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
            Left            =   840
            TabIndex        =   34
            Top             =   105
            Width           =   1200
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   2265
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
            Height          =   600
            Left            =   -20
            Picture         =   "REMECO04.frx":11281
            Style           =   1  'Graphical
            TabIndex        =   36
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
            TabIndex        =   37
            Top             =   105
            Width           =   1410
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2265
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
            Height          =   600
            Left            =   0
            Picture         =   "REMECO04.frx":1158B
            Style           =   1  'Graphical
            TabIndex        =   39
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
            TabIndex        =   40
            Top             =   105
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0FF&
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
      TabIndex        =   27
      Top             =   105
      Width           =   4665
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
         Left            =   3465
         TabIndex        =   29
         ToolTipText     =   "Click Para Seleccionar Período Mensual"
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
         Left            =   1365
         TabIndex        =   28
         ToolTipText     =   "Doble-Click Para Rango de Fechas"
         Top             =   410
         Width           =   2115
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
         Index           =   10
         Left            =   -945
         TabIndex        =   30
         Top             =   450
         Width           =   2220
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   3840
         Picture         =   "REMECO04.frx":11895
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3885
         Picture         =   "REMECO04.frx":11B9F
         Top             =   315
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00004080&
      Height          =   7000
      Left            =   9480
      ScaleHeight     =   6945
      ScaleWidth      =   990
      TabIndex        =   21
      Top             =   -120
      Width           =   1050
      Begin VB.PictureBox Picture16 
         Height          =   6250
         Left            =   0
         ScaleHeight     =   6195
         ScaleWidth      =   1005
         TabIndex        =   61
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
         Begin VB.Image Image4 
            Height          =   6255
            Left            =   0
            Picture         =   "REMECO04.frx":11EA9
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1065
         End
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
         Height          =   690
         Left            =   105
         Picture         =   "REMECO04.frx":14241
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Archivos de Transferencia Ingresos Brutos / C.I.T.I. Compras/Perc.Ganancias"
         Top             =   2990
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   58
         Top             =   5500
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -315
            TabIndex        =   59
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command8 
         Caption         =   "AcDir"
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
         Picture         =   "REMECO04.frx":1454B
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3680
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "REMECO04.frx":14855
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   990
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "REMECO04.frx":14B5F
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   4370
         Width           =   650
      End
      Begin VB.CommandButton Command15 
         Caption         =   "A-B-M"
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
         Picture         =   "REMECO04.frx":14FA1
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Alrchivo Maestro de Proveedores"
         Top             =   2300
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "REMECO04.frx":152AB
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   300
         Width           =   650
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -315
         Picture         =   "REMECO04.frx":153F5
         Top             =   5775
         Width           =   1515
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0FF&
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
      TabIndex        =   8
      Top             =   5250
      Width           =   4635
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
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
         BackColor       =   &H00C0E0FF&
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
      BackColor       =   &H00C0E0FF&
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
      Top             =   105
      Width           =   4425
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Estadísticas de Compras"
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
         TabIndex        =   52
         Top             =   4620
         Width           =   4005
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3135
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
               Height          =   600
               Left            =   0
               Picture         =   "REMECO04.frx":17317
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
               Height          =   600
               Left            =   2580
               Picture         =   "REMECO04.frx":17621
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
               Height          =   600
               Left            =   2020
               Picture         =   "REMECO04.frx":1792B
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
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   720
               TabIndex        =   57
               Top             =   90
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
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
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2055
            TabIndex        =   17
            Top             =   350
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
               Picture         =   "REMECO04.frx":17D6D
               Style           =   1  'Graphical
               TabIndex        =   18
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
               TabIndex        =   19
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0E0FF&
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
         Height          =   1170
         Left            =   210
         TabIndex        =   11
         Top             =   3320
         Width           =   4005
         Begin VB.PictureBox Picture15 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3150
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
               Height          =   600
               Left            =   2600
               Picture         =   "REMECO04.frx":181AF
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
               Height          =   600
               Left            =   0
               Picture         =   "REMECO04.frx":182F9
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
               Height          =   600
               Left            =   2020
               Picture         =   "REMECO04.frx":18603
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
               TabIndex        =   15
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ordenes de Pago y Ajustes"
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
         TabIndex        =   1
         Top             =   315
         Width           =   4005
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2370
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
               Height          =   600
               Left            =   -20
               Picture         =   "REMECO04.frx":1890D
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Detalle por Or- denes de Pago"
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
               TabIndex        =   4
               Top             =   105
               Width           =   1695
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2370
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
               Height          =   600
               Left            =   0
               Picture         =   "REMECO04.frx":18C17
               Style           =   1  'Graphical
               TabIndex        =   6
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
               TabIndex        =   7
               Top             =   0
               Width           =   1590
            End
         End
      End
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   4680
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu edicion 
      Caption         =   "Edición"
      Begin VB.Menu reimputaciondecomprobantesdecompras 
         Caption         =   "Re-Imputación de Comprobantes de Compras"
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
      End
      Begin VB.Menu notadebito 
         Caption         =   "Notas de Dédito (Del Mes Activo)"
      End
      Begin VB.Menu facturasdecompra 
         Caption         =   "Facturas de Compra  (Del Mes Activo)"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu librosdecompras 
         Caption         =   "Libros de Compra "
         Begin VB.Menu subdiariodecompras 
            Caption         =   "Subdiario de Compras por Impresora"
         End
         Begin VB.Menu subiarioivacomprasporimpresora 
            Caption         =   "Subdiario de I.V.A.- Compras por Impresora"
         End
         Begin VB.Menu resumenmensualivacompras 
            Caption         =   "Resumen Mensual I.V.A. Compras"
         End
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu resumenmensualporcuenta 
         Caption         =   "Resumen Mensual por Cuenta (Saldo de Proveedores al Comenzar y al Terminar el mes)"
      End
      Begin VB.Menu sss 
         Caption         =   "-"
      End
      Begin VB.Menu listadoporcuentacontable 
         Caption         =   "Consulta / Listado por Cuenta Contable"
      End
      Begin VB.Menu consultadeasientocontableporpantalla 
         Caption         =   "Consulta de Asiento Contable por Pantalla"
      End
      Begin VB.Menu estadisticodecompras 
         Caption         =   "Listado Estadístico de Facturación de Compras"
      End
      Begin VB.Menu rankingporproveedorporrangodefechas 
         Caption         =   "Ranking de Compras por Proveedor por Rango de Fechas"
      End
      Begin VB.Menu rankingdecomprasporarticuloyfecha 
         Caption         =   "Ranking de Compras por Articulo por Rango de Fechas"
      End
   End
   Begin VB.Menu ordenesdepagoyajustes 
      Caption         =   "Ordenes de Pago y Ajustes"
      Begin VB.Menu ajustesyotrosmovimientos 
         Caption         =   "Ajustes y Otros Movimientos"
         Begin VB.Menu pantalla 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu detalledeordenesdepago 
         Caption         =   "Detalle de Ordenes de Pago"
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
      Begin VB.Menu maestroproveedores 
         Caption         =   "Maestro de Proveedores"
      End
   End
   Begin VB.Menu configuracion 
      Caption         =   "Configuración"
      Begin VB.Menu config 
         Caption         =   "Configuración"
      End
   End
End
Attribute VB_Name = "REMECO04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DESDE%, HASTA%, CORRESP$
Dim DESDAT As Double
Dim HASDAT As Double

Private Sub cerrarmes_Click()
    Call Image2_Click
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
    DES% = DESDE%: DESDEX = CDbl(CVDAT(DES%))
    HAS% = HASTA%: HASTAX = CDbl(CVDAT(HAS%))
    '
10  Call SELECHO("PROVED1")
    NC% = XVALO(VALACT1$("PROVED1"))
    If NC% < 1 Then GoTo 99
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FCOM' OR TIPOMOVIM = 'AJUP') "
    SQLX$ = SQLX$ + "AND NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND FechConta >= " & DESDEX & " "
    SQLX$ = SQLX$ + "AND FechConta <= " & HASTAX & " "
    SQLX$ = SQLX$ + "Order by FechConta ASC "
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
            NPX& = XVALO(Mid$(!CodiCom_Lar, 11, 8))
            Y$ = MKS$(JJ& + 1) + AJUSTI$(!CodiCom_Lar, 18) + IMPOS$
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
20  ATRI$ = DOCULA$ + " - $ " + TRIM$(Mid$(VALACT2$("!SELCUERE"), 19, 14)) + " - " + TRIM$(RASOCLI$(NC%))
    '
    Call BLANARCH("!REIMPUTA")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(FechConta)) >= " & DESDEX & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASTAX & " "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & DOCULA$ & "' "
    SQLX$ = SQLX$ + "AND NuProve = " & NC% & " "
    '
    Set CABASEL = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
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
        CUE% = !CueContaI
        'If CUE% > 0 Then
           IPA# = !Idebecomp - !IHABECOMP
           Y$ = REGBLAN$("!REIMPUTA")
           Call REPLA(Y$, CODCUE$(CUE%), 1, 12)
           Call REPLA(Y$, MKD$(IPA#), 13, 8)
           '\\\OT-06-0124-WAR-03/04/06///
           CC1$ = "": If IsNull(!CENCOSTO) = False Then CC1$ = !CENCOSTO
           Call REPLA(Y$, CC1$, 21, 8)
           '\\\OT-06-0124-WAR-FIN///
           JJR& = JJR& + 1
           Call GRAREG("!REIMPUTA", Y$, JJR&)
           SUIMPU# = SUIMPU# + IPA#
        'End If
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
        Y$ = REGLEIDO$("!REIMPUTA", U&)
        CUENUE% = CUEINT%(Left$(Y$, 12))
        CC1$ = Mid$(Y$, 21, 8)
        '
        If CUENUE% > 0 Then
          With CABASEL
          .Edit
          !CueContaI = CUENUE%
          !CENCOSTO = CC1$
          If !TipoMovim = "COMPRA" Then
             ' SOLICITADO POR Claudia Matsusaka 2l 23/2/07 que
             ' cambie la descripción acorde con la nueva cuenta
             ' contable
             DEMOVX$ = DENOCUE$(CUENUE%)
             NC1% = (-1) * Abs(NC%)
             !Descrimov = Left$(TRIM$(DEMOVX$) + " - " + RASOCLI$(NC1%), 48)
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
   Command11.Enabled = False
   '\\\OT-05-0275-WAR-11/05/05///
   JJ0% = COMALTER%("Opciones de Trabajo:\\Ingresos Brutos\C.I.T.I. Compras\Percepciones de Ganancias\Regimen Informativo de Compras")
   If JJ0% < 1 Or JJ0% > 4 Then GoTo 99
   If JJ0% = 1 Then
     JJ1% = COMALTER%("Opciones de Trabajo:\\Percepciones\Retenciones")
     If JJ1% < 1 Or JJ1% > 2 Then GoTo 99
     '\\\OT-05-0275-WAR-FIN///
     If JJ1% = 2 Then
       'RETENCIONES
       OPX% = COMALTER%("Opciones de Trabajo:\\Subdiario I.Brutos\Archivo Transferencia")
       If OPX% < 1 Or OPX% > 2 Then GoTo 99
       '
       Call DESHASFECHA(DES%, HAS%, PERIO$)
       If DES% < 1 Then GoTo 99
       '
       If OPX% = 1 Then
         CORRESP$ = PERIO$
         Call HEADERS("SUBRETI", "")
         Call HEADERS("SUBRETI", "Corresponde a: " + CORRESP$)
         J& = 0: FIN& = ULTREG("SURETIB")
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           XX$ = REGLEIDO$("SURETIB", U&)
           FEMIX% = CVI(Left$(XX$, 2))
           If FEMIX% >= DES% Then
             If FEMIX% <= HAS% Then
               J& = J& + 1
               Call GRAREG("SUBRETI", XX$, J&)
             End If
           End If
         Next U&
         Call SETULTREG("SUBRETI", J&)
         Call FINAL("*SUBRETI")
       ElseIf OPX% = 2 Then
  ''       CommonDialog1.CancelError = True
  ''       CommonDialog1.DialogTitle = "Crear Archivo de Transferencia"
  ''       'deshabilitado porque si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
  ''       CommonDialog1.Filter = "Todos los Archivos (*.*)|*.*|"
  ''       CommonDialog1.FilterIndex = 1
  ''       On Error GoTo 99
  ''       CommonDialog1.ShowOpen
  ''       On Error GoTo 0
         '
         ANCHIMPOR% = 11
         EPAC$ = TRIM$(UCase$(EMPREAC$))
         If InStr(EPAC$, "DOSIVA") > 0 Then ANCHIMPOR% = 10
         '
  ''       ASALID$ = UCase$(CommonDialog1.filename)
         ASALID$ = LUCOM$ + "TRANIBRU.TXT"
         NX% = FILEOPEN%(ASALID$, 2, 128)
         ASALID1$ = LUCOM$ + "DDJJRDET.TXT"
         NX1% = FILEOPEN%(ASALID1$, 2, 128)
         '
10       cuitagen$ = CONTROL$("CUITAGEN", "CUITAGEN")
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
         PL11 = CVDAT(HAS%)
         PL1$ = Format(PL11, "YYYYMM")
         Call REPLA(REBLA$, PL1$, 14, 6)
         Call REPLA(REBLA$, "S", 73, 1)
         PL22 = CVDAT(HAS%)
         PL2$ = Format(PL22, "YYYYMMDD")
         Call REPLA(REBLA$, PL2$, 144, 8)
         Call REPLA(REBLA$, "OTA0001", 152, 7)
         Call REPLA(REBLA$, "PSO", 167, 3)
         Call REPLA(REBLA$, "2.5000", 188, 6)
         Call REPLA(REBLA$, "S", 210, 1)
         '\\\OT-05-0285-WAR-FIN///
         CAREGI% = 0
         For U& = 1 To ULTREG&("SURETIB")
20         XX$ = REGLEIDO$("SURETIB", U&)
           FEMIX% = CVI(Left$(XX$, 2))
           If FEMIX% >= DES% Then
             If FEMIX% <= HAS% Then
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
                 While Len(IMPOX$) < ANCHIMPOR%: IMPOX$ = "0" + IMPOX$: Wend
               Call REPLA(TTXX$, IMPOX$, 36, ANCHIMPOR%)
               Print #NX%, TTXX$
               '\\\OT-05-0285-WAR-12/05/05///
               TTYY$ = REBLA$
               Call REPLA(TTYY$, CUITTX$, 20, 13)
               NC% = CVI(Mid$(XX$, 95, 2))
               RS1$ = RASOCLI(NC% * (-1))
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
               If CVD(Mid$(XX$, 87, 8)) < 0.005 Then GoTo 29 ' SALTEA REGISTROS SIN RETENCION
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
29       Next U&
         ' PONE MARCA DE FIN DE ARCHIVO SEGúN MUESTRA RECIBIDA
         Print #NX1%, Chr$(26)
         ' FIN MARCA DE FIN DE ARCHIVO
         Close #NX%
         Close #NX1%
         ASALID3$ = REPLACAR(ASALID$, "\", "/")
         ASALID4$ = REPLACAR(ASALID1$, "\", "/")
         Call COMUNI("Archivos de Transferencia\ '" + ASALID3$ + "' y '" + ASALID4$ + "' \Generados con " + TRIM$(Str$(CAREGI%)) + " Registros.")
       End If
     End If
      '\\\OT-05-0275-WAR-11/05/05///
     If JJ1% = 1 Then
        'PERCEPCIONES
21      Call DESHASFECHA(DES%, HAS%, PERIO$)
        If DES% < 1 Then GoTo 99
        DES1# = Int(CDbl(CVDAT(DES%)))
        HAS1# = Int(CDbl(CVDAT(HAS%)))
        Call APERBASDAT("CABAN")
        SQLX$ = "SELECT * FROM CAJABANC "
        SQLX$ = SQLX$ + "WHERE INT(CDBL(FECHCONTA)) >= " & DES1# & ""
        SQLX$ = SQLX$ + "AND INT(CDBL(FECHCONTA)) <= " & HAS1# & " "
        SQLX$ = SQLX$ + "AND TIPOMOVIM = 'IBRU' "
        SQLX$ = SQLX$ + "AND VARIMOVIM = 'PERC' "
        SQLX$ = SQLX$ + "AND NUPROVE > 0 "
        Set PERLISTMP = CueCorri.OpenRecordset(SQLX$, 4)
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
          Do While .EOF = False
            Call REPLA(ZZ$, MKI(CVINT(!FechConta)), 1, 2)
            Call REPLA(ZZ$, !CodiCom_Lar, 3, 22)
            NC% = !Nuprove
            CTP$ = CODICLI$((-1) * NC%)
            Call REPLA(ZZ$, CTP$, 25, 12)
            IMPER# = !Idebecomp - !IHABECOMP
            Call REPLA(ZZ$, MKD(IMPER#), 37, 8)
            CPR$ = TRIM$(Mid$(!OpciMovim, 4))
            CPR1$ = ECOARCH$("PROVINC", CPR$)
            Call REPLA(ZZ$, CPR$, 45, 1)
            JJ& = JJ& + 1
            Call GRAREG("DEPERIB", ZZ$, JJ&)
            .MoveNext
          Loop
          Call SETULTREG("DEPERIB", JJ&)
          Call FILESORT("DEPERIB", "DEPERIB", 1, 1, "ASC") ' POR FECHA
          Call FILESORT("DEPERIB", "DEPERIB", 6, 6, "ASC") ' POR JURISDICCIóN
          Call FINAL("*DEPERIB")
          '
        End With
      End If
    End If
      '\\\OT-05-0275-WAR-FIN///
    '\\\OT-06-0271-WAR-07/07/06///
    If JJ0% = 2 Then 'C.I.T.I. Compras
      Call CITICOMP
    End If
    '\\\OT-06-0271-WAR-FIN///
    '
    ''\\\OT-08-0589-OD-08/08/08///
    If JJ0% = 3 Then 'Ganancias
        'SELECCION DE FECHAS
25      Call DESHASFECHA(DES%, HAS%, PERIO$)
        If PERIO$ = "" Then GoTo 99
        DESDAT = Int(CDbl(CVDAT(DES%)))
        HASDAT = Int(CDbl(CVDAT(HAS%)))
        '
        Screen.MousePointer = 11
        JX& = 0
        Call APERBASDAT("CABAN")
        SQLX$ = "SELECT * FROM CAJABANC"
        SQLX$ = SQLX$ + " WHERE TipoMovim = 'GAN' "
        SQLX$ = SQLX$ + " AND   VARIMOVIM= 'Per-Gan'"
        SQLX$ = SQLX$ + " AND   CDBL(FechConta) >= " & DESDAT & ""
        SQLX$ = SQLX$ + " AND   CDBL(FechConta) <= " & HASDAT & ""
        Set GananTemp = CueCorri.OpenRecordset(SQLX$, 4)
        With GananTemp
          On Error Resume Next
          .MoveFirst
          If Err Then
            Call COMUNI("No existen Percepciones en el Rango de Fechas Elegido")
            On Error GoTo 25
            GoTo 99
          End If
          On Error GoTo 0
          '
          CAREG1% = 0
          Do While Not .EOF
            '
            FECHA = !FechEmisi
            COMPROBANTE$ = !CodiCom_Lar
            CERTIFICADO$ = Mid$(!Descrimov, 29, 12)
            CLIENTE% = !Nuprove
            CUIT$ = CUITCLI$(!Nuprove * (-1))
            IMPORTE = !Idebecomp
            '\\\OT-08-0916-OD-01/10/08///
            CODIPERGAN$ = SAFETEXT$(!CodPerGan)
            If CODIPERGAN$ = "" Then CODIPERGAN$ = "19"
            '\\\OT-08-0916-OD-FIN//
            '
            Call REPLA(TXXX$, MKI$(CVINT(FECHA)), 1, 2)
            Call REPLA(TXXX$, COMPROBANTE$, 3, 18)
            Call REPLA(TXXX$, CERTIFICADO$, 21, 12)
            Call REPLA(TXXX$, MKI$(CLIENTE%), 33, 2)
            Call REPLA(TXXX$, CUIT$, 35, 13)
            Call REPLA(TXXX$, MKS(IMPORTE), 48, 4)
            Call REPLA(TXXX$, CODIPERGAN$, 52, 3)
            '
            JX& = JX& + 1
            Call GRAREG("!PERETGAN", TXXX$, JX&)
            .MoveNext
          Loop
        End With
        Call SETULTREG("!PERETGAN", JX&)
        '\\\OT-08-0916-OD-01/10/08///
        VTB% = VENTABU%("PERETGA2")
        If VTB% >= 0 Then
           For U& = 1 To ULTREG&("!PERETGAN")
           XX$ = REGLEIDO$("!PERETGAN", U&)
           Call GRAREG("PERETGAN", XX$, U&)
           '
           ABUS$ = Mid$(XX$, 3, 18)
           FEC% = CVI(Mid$(XX$, 1, 2))
           FECCHH = Int(CDbl(CVDAT(FEC%)))
           NuProvv% = CVI(Mid$(XX$, 33, 2))
           XODPERGAN$ = TRIM$(Mid$(XX$, 52, 3))
           '
           'que coincida documento ,fecha y proveedor
           SQLX$ = "UPDATE CAJABANC SET CodPerGan = '" & XODPERGAN$ & "'"
           SQLX$ = SQLX$ + " WHERE CodiCom_Lar = '" & ABUS$ & "'"
           SQLX$ = SQLX$ + " AND INT(CDBL(FechEmisi)) = " & FECCHH & ""
           SQLX$ = SQLX$ + " AND Nuprove = " & NuProvv% & ""
           CueCorri.Execute (SQLX$)
          Next U&
          Call SETULTREG("PERETGAN", JX&)
         Else
          GoTo 25
        End If
        '\\\OT-08-0916-OD-FIN///
        '
        OPPCC% = COMALTER%("Opciones de Trabajo:\\Subdiarios de Ganancias\Archivo de Transferencia")
        If OPPCC% < 1 Or OPPCC% > 2 Then GoTo 25
        If OPPCC% = 1 Then
            Call FILESORT("PERETGAN", "", 2, 2, "ASC")
            Call FILESORT("PERETGAN", "", 1, 1, "ASC")
            '
            Call HEADERS("PERETGAN", "")
            Call HEADERS("PERETGAN", "Periodo: " & PERIO$)
            Call FINAL("*PERCGAN")
        Else
            Dim XXX As String * 52
            ASALID$ = LUCOM$ + "PERGAN.TXT"
            NX% = FILEOPEN%(ASALID$, 2, 128)
            FIN& = ULTREG("PERETGAN")
            Screen.MousePointer = 11
            '
            For I& = 1 To FIN&
               TTXX$ = REGLEIDO$("PERETGAN", I&)
               XXX = ""
                 CUIT$ = Mid$(TTXX$, 35, 13)
                 NuCuit$ = REPLACAR(CUIT$, "-", "")
               Call REPLA(XXX, "2", 1, 1)
               Call REPLA(XXX, NuCuit$, 2, 11)
                 FECH% = CVI(Mid$(TTXX$, 1, 2))
               Call REPLA(XXX, FECHATEXLAR(FECH%), 13, 10)
                REGIMEN$ = Mid$(TTXX$, 52, 3)
                'REGIMEN$ = "19"
               Call REPLA(XXX, REGIMEN$, 23, 3)
                 IMPO# = CVS(Mid$(TTXX$, 48, 4))
               Call REPLA(XXX, FORMATNUM(IMPO#, "F15.2"), 26, 15)
                 CERTIFICADO$ = AJUSTI$(TRIM$(Mid$(TTXX$, 21, 12)), 12)
               Call REPLA(XXX, CERTIFICADO$, 41, 12)
               Print #NX%, XXX
            Next I&
            Close #NX%
            ASALID11$ = REPLACAR(ASALID$, "\", "/") 'INVERTIMOS LAS BARRAS PARA QUE _
            NO LO PRESENTE EN DIFERENTES RENGLONES EL CALL COMUNI

            Screen.MousePointer = 1
            Call COMUNI("Archivo de Transferencia '" + ASALID11$ + "'\Generado con " + TRIM$(Str$(FIN&)) + " Registros.")
        End If
    End If
    
    If JJ0% = 4 Then
        If VALIMESAN% < 1 Then GoTo 99
        Call GENIVACOM
        Call RegimenInformativoCompras
    End If
    '\\\OT-08-0589-OD-FIN///
    '
99  Command11.Enabled = True
    Screen.MousePointer = 1
    On Error GoTo 0
End Sub

Private Sub RegimenInformativoCompras()
    
    
    On Error Resume Next
    Call TRALOCAL("LIVACOM", "")
    On Error GoTo 0
    
    If Err > 0 Then
        Call MENSERR(24, "Imposible Copiar Archivo con Datos del Período")
        Exit Sub
    End If
    
    Call ABRECABAN
    
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
    
    Dim IdNumeroDespacho%
    
    INDICE% = 1: GoSub InicializarCamposVGD
    INDICE% = 2: GoSub InicializarCamposVGD
    INDICE% = 3: GoSub InicializarCamposVGD
    
    FIN& = ULTREG&("!LIVACOM")
    If FIN& = 0 Then GoTo 99
    
    Screen.MousePointer = 11
    
    For I& = 1 To ULTREG&("!LIVACOM")
        Call TRACE(24, I&, ULTREG&("!LIVACOM"))
        
        REGISTRO$ = REGLEIDO$("!LIVACOM", I&)
        NumeroComprobante$ = Mid$(REGISTRO$, 11, 18)
        
        CuitProveedor$ = TRIM$(Mid$(REGISTRO$, 63, 16))
        If TRIM$(NumeroComprobante$) = "" Then
            NumeroComprobante$ = UltimoNumeroComprobante$
            CuitProveedor$ = UltimoCuitProveedor$
        End If
       
        If (NumeroComprobante$ <> UltimoNumeroComprobante$) Or ((NumeroComprobante$ = UltimoNumeroComprobante$) And (CuitProveedor$ <> UltimoCuitProveedor$)) Then
            '
            RegistroAnterior$ = REGISTRO$
            
            FechaComprobante$ = FECHATEXLAR$(CVI(Mid$(REGISTRO$, 3, 2)))
            FechaComprobante$ = Mid$(FechaComprobante$, 7, 4) & Mid$(FechaComprobante$, 4, 2) & Mid$(FechaComprobante$, 1, 2)
            
            If UltimoNumeroComprobante$ <> "" And FlagError% = 0 Then
                If ImporteExImportacion# > 0.05 Then
                    CantAlicuotas% = CantAlicuotas% + 1
                    CantidadAlicuotas$ = CStr(CantAlicuotas%)
                    Mid$(RegistroRIV$, 238, 1) = CantidadAlicuotas$
                End If
                If TRIM$(RegistroRIV$) <> "" Then
                    Print #NX%, RegistroRIV$
                    RegistroRIV$ = ""
                End If
                If ImporteExImportacion# > 0 Then
                    ImporteNeto$ = RIVFormatoNumero(ImporteExImportacion#, 15)
                    Alicuota$ = "0003": ImporteIva0$ = RIVFormatoNumero(0, 15)
                    RegistroRIV_Importaciones$ = NumeroDespachoImportacion$ & ImporteNeto$ & Alicuota$ & ImporteIva0$
                    Print #NX2%, RegistroRIV_Importaciones$
                    ImporteExImportacion# = 0
                End If
            End If
            FlagError% = 0
            EsRegimenSimplificado% = 0
            
25          NumeroProveedor% = CVI(Mid$(REGISTRO$, 31, 2)) * (-1)
            If NumeroProveedor% = 0 Then
                ListaDeErrores$ = ListaDeErrores$ & "No Se Pudo Encontrar Numero de Proveedor para el Comprobante '" & NumeroComprobante$ & "'"
                FlagError% = 1: GoTo 35
            End If

            TipoCbte$ = TipoComprobante$(NumeroComprobante$)
            
            PuntoVta$ = PuntoVenta$(NumeroComprobante$)
            ' Es obligatorio que el punto de venta sea > 0
            If PuntoVta$ = "00000" Then PuntoVta$ = "00001"
            
            NumeroCbt$ = NumeroCbte$(NumeroComprobante$)

            ' Verifica si se trata de una factura de importación común o simplificada
                Despa$ = SAFETEXT$(VALOBADAMDB("CAJABANC", "DespaImpo", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%), "NOMESS"))
            NumeroDespachoImportacion$ = AJUSTI$(Despa$, 16)
            EsFacturaDespachante# = 0
            EsFacturaDespachante# = XVALO(VALOBADAMDB("CAJABANC", "ValAbsComp", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%) & " AND Varimovim='Aduana'"))
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
                TipoCbte$ = "066": PuntoVta$ = "00000": NumeroCbt$ = "00000000000000000000"
            End If
            
            CodigoDocComprador$ = CodigoDocumento$(PIVACLI%(NumeroProveedor%))
            
            NumeroIdentificacionComprador$ = SAFETEXT$(VALOBADAMDB("CAJABANC", "Cuit", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%), "NOMESS"))
            If TRIM$(NumeroIdentificacionComprador$) = "" Or VALICUIT%(NumeroIdentificacionComprador$) = 0 Then NumeroIdentificacionComprador$ = TRIM$(Mid$(REGISTRO$, 63, 16))
            If NumeroIdentificacionComprador$ = "" Or VALICUIT%(NumeroIdentificacionComprador$) = 0 Then
                NumeroIdentificacionComprador$ = CUITCLI$(NumeroProveedor%)
            End If
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
            RazonSocial$ = AJUSTI$(SAFETEXT$(VALOBADAMDB("CAJABANC", "RefeCom", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%), "NOMESS")), 30)
            If TRIM$(RazonSocial$) = "" Then RazonSocial$ = AJUSTI$(RASOCLI$(NumeroProveedor%), 30)
            LargoRS% = InStr(1, RazonSocial$, "//")
            If LargoRS% > 0 Then RazonSocial$ = AJUSTI$(Mid(RazonSocial$, 1, LargoRS% - 1), 30)
            If PIVACLI%(NumeroProveedor%) = 0 Then
                RazonSocial$ = AJUSTI$("CONSUMIDOR FINAL", 30)
            End If
            
            ImporteTotal$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 91, 8)), 15)
            
            ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
            ValorIva$ = TRIM$(Mid$(REGISTRO$, 129, 8))
            If InStr(1, UCase(ValorIva$), "NO GRAV") > 0 Then
                ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
            End If
            
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
                
                MONE% = XVALO(VALOBADAMDB("CAJABANC", "Mone_Emis", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
            CodMoneda$ = CodigoMoneda(MONE%)
            If CodMoneda$ = "000" Then CodMoneda$ = "PES"
            
                TCambio# = XVALO(VALOBADAMDB("CAJABANC", "Tipo_Cam", "CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
            TipoCambio$ = RIVFormatoNumero$(TCambio#, 10, 6)
            If CodMoneda$ = "PES" And TCambio# > 1 Then
                    ListaDeErrores$ = ListaDeErrores$ & "Error en Tipo de Cambio en el Proveedor '" & CStr(Abs(NumeroProveedor%)) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                    FlagError% = 1: GoTo 35
            End If
            If XVALO(TipoCambio$) = 0 And CodMoneda$ = "PES" Then TipoCambio$ = RIVFormatoNumero(1, 10, 6)
                                
                CantAlicuotas% = 1
            If InStr(1, "006;007;008;009;082;013;012;011;015;013;111", TipoCbte$) > 0 Then CantAlicuotas% = 0 ' Si es Comprobante "C" la cantidad debe ser cero
            CantidadAlicuotas$ = CStr(CantAlicuotas%)
            
            Percepciones# = XVALO(ImportePercepcionesNacionales$) / 100 + XVALO(ImportePercepcionIB$) / 100 + XVALO(ImportePercepcionPagosACuenta$) / 100
            If PIVACLI%(NumeroProveedor%) = 7 Then
                CodigoOperacion$ = "Z"
            ElseIf PIVACLI%(NumeroProveedor%) = 5 Or (TRIM$(NumeroDespachoImportacion$) <> "" Or EsFacturaDespachante# > 0.005) Then
                CodigoOperacion$ = "X"
            ElseIf InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 Or (TRIM$(Mid$(REGISTRO$, 129, 8)) = "" And Percepciones# > 0.005) Then
                CodigoOperacion$ = "N"
            ElseIf PIVACLI%(NumeroProveedor%) = 4 Or (InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "EXENTO") > 0 And EsRegimenSimplificado% = 0) Then
                CodigoOperacion$ = "E"
            ElseIf CantAlicuotas% = 0 Then
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
            RegistroRIV$ = FechaComprobante$ & TipoCbte$ & PuntoVta$ & NumeroCbt$ & NumeroDespachoImportacion$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & RazonSocial$ _
                               & ImporteTotal$ & ImporteNoGravado$ & ImporteOperacionExenta$ & ImportePercepcionPagosACuenta$ & ImportePercepcionesNacionales$ _
                               & ImportePercepcionIB$ & ImportePercepcionMunicipales$ & ImporteImpuestosInternos$ & CodMoneda$ & TipoCambio$ & CantidadAlicuotas$ _
                               & CodigoOperacion$ & ImporteIva$ & OtrosTributos$ & CuitEmisor$ & DenominacionEmisor$ & IvaComision$
                
            UltimoNumeroComprobante$ = NumeroComprobante$
            UltimoCuitProveedor$ = CuitProveedor$
            
            If XVALO(CantidadAlicuotas$) > 0 Then
            
                ' Tratamiento para los comprobantes con parte no gravada o exenta
                If TRIM$(CodigoOperacion$) <> "" Then
                    ImporteNeto$ = RIVFormatoNumero(0, 15)
                    Alicuota$ = "0003"
                    ImporteIva0$ = RIVFormatoNumero(0, 15)
                    
                    If (CodigoOperacion$ = "N") Then
                        ConIva# = XVALO(VALOBADAMDB("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                        Percepciones# = XVALO(ImportePercepcionesNacionales$) / 100 + XVALO(ImportePercepcionIB$) / 100 + XVALO(ImportePercepcionPagosACuenta$) / 100
                        If ConIva# < 0.01 Or (Percepciones# > 0.005 And ConIva# < 0.01) Then
                            ' cuando el comprobante es totalmente no gravado, dicho importe se escribe en el archivo de alicuotas
                            ImporteNeto$ = ImporteTotal$
                            If XVALO(Percepciones#) > 0.005 Then ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                            ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
                            Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                        Else
                            ' Cuando el comprobante tiene parte gravada y parte que no lo esta, el no gravado se escribe
                            ' en el archivo de comprobantes y los importes gravados en el de alicuotas
                            CantAlicuotas% = 0
                            CantidadAlicuotas$ = CStr(CantAlicuotas%)
                            Call REPLA(RegistroRIV$, " ", 239, 1)
                            Call REPLA(RegistroRIV$, CantidadAlicuotas$, 238, 1)
                            GoTo 35
                        End If
                    End If
                    If CodigoOperacion$ = "E" Then
                        ConIva# = XVALO(VALOBADAMDB("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                        Percepciones# = XVALO(ImportePercepcionesNacionales$) / 100 + XVALO(ImportePercepcionIB$) / 100 + XVALO(ImportePercepcionPagosACuenta$) / 100
                        If ConIva# < 0.01 Or (Percepciones# > 0.005 And ConIva# < 0.01) Then
                            ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                            ImporteOperacionExenta$ = RIVFormatoNumero$(0, 15)
                            Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 135, 15)
                        Else
                            CantAlicuotas% = 0
                            CantidadAlicuotas$ = CStr(CantAlicuotas%)
                            Call REPLA(RegistroRIV$, " ", 239, 1)
                            Call REPLA(RegistroRIV$, CantidadAlicuotas$, 238, 1)
                            ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                            Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                            GoTo 35
                        End If
                    End If
                    If CodigoOperacion$ = "X" And (TRIM$(NumeroDespachoImportacion$) <> "" Or EsFacturaDespachante# > 0.005) Then
                        If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") > 0 Then
                            OtrosTributos$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
                            ImporteIva$ = RIVFormatoNumero(0, 15)
                            Call REPLA(RegistroRIV$, ImporteIva$, 240, 15)
                            CantAlicuotas% = 0
                            CantidadAlicuotas$ = CStr(CantAlicuotas%)
                            GoTo 35
                        ElseIf InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "EXENTO") > 0 Then
                            ConIva# = XVALO(VALOBADAMDB("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and VariMovim<>'Aduana' and VariMovim<>'Per-Iva' and VariMovim<>'Per-Gan' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                            If ConIva# > 0.005 Then
                                ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                CantAlicuotas% = 0
                                CantidadAlicuotas$ = CStr(CantAlicuotas%)
                                GoTo 35
                            Else
                                ImporteNoGravado$ = RIVFormatoNumero(0, 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                                ImporteExImportacion# = XVALO(ImporteNeto$) / 100
                                CantAlicuotas% = 0
                                CantidadAlicuotas$ = CStr(CantAlicuotas%)
                                GoTo 35
                            End If
                        ElseIf InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 Then
                            ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuProve=" & Abs(NumeroProveedor%)))
                            If ConIva# > 0.005 Then
                                ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                CantAlicuotas% = 0
                                CantidadAlicuotas$ = CStr(CantAlicuotas%)
                                GoTo 35
                            Else
                                ImporteNoGravado$ = RIVFormatoNumero(0, 15)
                                Call REPLA(RegistroRIV$, ImporteNoGravado$, 120, 15)
                                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                            End If
                        ElseIf InStr(1, "0004;0005;0006;0008;0009", CodigoAlicuota(Mid$(REGISTRO$, 129, 8))) > 0 Then
                            ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                            ImporteIva0$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
                            Alicuota$ = CodigoAlicuota(ValorIva$)
                        End If
                    End If
                Else
                    ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                    ImporteIva0$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
                    Alicuota$ = CodigoAlicuota(ValorIva$)
                End If
                
                If (TRIM$(NumeroDespachoImportacion$) = "" And EsFacturaDespachante# < 0.005) Or EsRegimenSimplificado% = 1 Then
                
                    RegistroRIV_Alicuotas$ = TipoCbte$ & PuntoVta$ & NumeroCbt$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & ImporteNeto$ & Alicuota$ & ImporteIva0$
                    Print #NX1%, RegistroRIV_Alicuotas$
                Else
                    
                    RegistroRIV_Importaciones$ = NumeroDespachoImportacion$ & ImporteNeto$ & Alicuota$ & ImporteIva0$
                    Print #NX2%, RegistroRIV_Importaciones$
                End If
            End If
        Else
            If FlagError% = 1 Then GoTo 35
            
            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 Then
                ImporteNGAnterior = XVALO(Mid$(RegistroRIV$, 120, 15)) / 100
                ImporteNGActual = CVD(Mid$(REGISTRO$, 137, 8))
                'If Not (ImporteNGActual < 0 And (Mid$(NumeroComprobante$, 1, 2) = "FC" Or Mid$(NumeroComprobante$, 1, 2) = "CO")) Then
                    ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior + ImporteNGActual, 15)
                'Else

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
                                
                IMPO# = XVALO(Mid$(RegistroRIV$, 240, 15)) / 100
                ImporteIvaActual# = Abs(CVD(Mid$(REGISTRO$, 145, 8)))
            ImporteIva$ = RIVFormatoNumero(IMPO# + ImporteIvaActual#, 15)

                CantAlicuotas% = XVALO(CantidadAlicuotas$)
            ImporteNeto0# = CVD(Mid$(REGISTRO$, 137, 8))
            ValorIva$ = TRIM$(REPLACAR$(Mid$(REGISTRO$, 129, 8), "%", ""))
            CantidadAlicuotaNuevoRegistro% = 0
            If (Not (ImporteNeto0# < 0 And (Mid$(NumeroComprobante$, 1, 2) = "FC" Or Mid$(NumeroComprobante$, 1, 2) = "CO"))) And XVALO(ImporteIvaActual#) > 0.005 And InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") = 0 Then CantidadAlicuotaNuevoRegistro% = 1
            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 And CodigoOperacion$ = "X" Then
                ImporteNGAnterior = XVALO(ImporteNoGravado$) / 100
                ImporteNGActual = CVD(Mid$(REGISTRO$, 137, 8))
                If ImporteNGActual > 0 Then
                    CantidadAlicuotaNuevoRegistro% = 1
                    ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior + ImporteNGActual, 15)
                Else
                    If (ImporteNGAnterior + ImporteExImportacion#) > Abs(ImporteNGActual) Then
                        ImporteExImportacion# = ImporteExImportacion# + ImporteNGAnterior - Abs(ImporteNGActual)
                    End If
                End If
            End If
            CantidadAlicuotas$ = CStr(CantAlicuotas% + CantidadAlicuotaNuevoRegistro%)
            
            If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "ADUANA") > 0 Then
                IMPO# = XVALO(Mid$(RegistroRIV$, 255, 15)) / 100
                OtAtributos# = CVD(Mid$(REGISTRO$, 145, 8))
                OtrosTributos$ = RIVFormatoNumero(IMPO# + OtAtributos#, 15)
                IMPO# = XVALO(ImporteIva$) / 100
                ImporteIva$ = RIVFormatoNumero(IMPO# - OtAtributos#, 15)
            End If

            RegistroRIV$ = FechaComprobante$ & TipoCbte$ & PuntoVta$ & NumeroCbt$ & NumeroDespachoImportacion$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & RazonSocial$ _
                           & ImporteTotal$ & ImporteNoGravado$ & ImporteOperacionExenta$ & ImportePercepcionPagosACuenta$ & ImportePercepcionesNacionales$ _
                           & ImportePercepcionIB$ & ImportePercepcionMunicipales$ & ImporteImpuestosInternos$ & CodMoneda$ & TipoCambio$ & CantidadAlicuotas$ _
                           & CodigoOperacion$ & ImporteIva$ & OtrosTributos$ & CuitEmisor$ & DenominacionEmisor$ & IvaComision$
            
            If (CantidadAlicuotaNuevoRegistro% > 0) Then
                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                ImporteIva0$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
                Alicuota$ = CodigoAlicuota(ValorIva$)
                If InStr(1, UCase(Mid$(REGISTRO$, 129, 8)), "NO GRAV") > 0 And CodigoOperacion$ = "X" Then
                    Alicuota$ = "0003"
                End If
                   
                If (TRIM$(NumeroDespachoImportacion$) = "" And EsFacturaDespachante# < 0.005) Or EsRegimenSimplificado% = 1 Then
                
                    RegistroRIV_Alicuotas$ = TipoCbte$ & PuntoVta$ & NumeroCbt$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & ImporteNeto$ & Alicuota$ & ImporteIva0$
                    Print #NX1%, RegistroRIV_Alicuotas$
                Else
                    
                    RegistroRIV_Importaciones$ = NumeroDespachoImportacion$ & ImporteNeto$ & Alicuota$ & ImporteIva0$
                    Print #NX2%, RegistroRIV_Importaciones$
                End If
            End If
        End If
35  Next I&
    
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

Function CodigoAlicuota(Porcentaje$) As String

    Dim codigo$, Porce$
    Porce$ = REPLACAR$(Porcentaje$, "%", "")
    If (XVALO(Porce$) = 10.5) Then
        codigo$ = "0004"
    ElseIf (XVALO(Porce$) = 21) Then
        codigo$ = "0005"
    ElseIf (XVALO(Porce$) = 27) Then
        codigo$ = "0006"
    ElseIf (XVALO(Porce$) = 5) Then
        codigo$ = "0008"
    ElseIf (XVALO(Porce$) = 2.5) Then
        codigo$ = "0009"
    ElseIf (UCase(Porce$) = "EXENTO") Or (UCase(Porce$) = "ADUANA") Then
        codigo$ = "0003"
    ElseIf (UCase(Porce$) = "NO GRAV") Then
        codigo$ = "0001"
    End If
    
    CodigoAlicuota = codigo$
    
End Function


Function RIVFormatoNumero(NUMERO#, Optional Longitud%, Optional Decimales%) As String
    
    If Longitud <= 0 Then Longitud = 15
    If Decimales% <= 0 Then Decimales% = 2
    'VALOR$ = TRIM$(FORMATNUM$(Abs(NUMERO#), "F" & CStr(Longitud%) & "." & Decimales%))
    VALOR$ = TRIM$(FORMATNUM$(Abs(NUMERO#), "F" & CStr(Longitud%) & "." & Decimales%))
    VALOR$ = REPLACAR$(VALOR$, ".", "")
    If VALOR$ = "" Then VALOR$ = "0"
    
    If XVALO(VALOR$) >= 0 Then
        Do While Len(VALOR$) < Longitud%
            VALOR$ = "0" & VALOR$
        Loop
    Else
        VALOR$ = Mid$(VALOR$, 2)
        Do While Len(VALOR$) < (Longitud% - 1)
            VALOR$ = "0" & VALOR$
        Loop
        VALOR$ = "-" & VALOR$
    End If
    
    RIVFormatoNumero = VALOR$
    
End Function

Function TipoComprobante(NumeroComprobante$) As String

    IdCbte$ = Mid$(NumeroComprobante$, 1, 2) & Mid$(NumeroComprobante$, 4, 1)
    
    TipoCbte$ = ""
    Select Case IdCbte$
        Case "FCA": TipoCbte$ = "001"
        Case "NDA": TipoCbte$ = "002"
        Case "NCA": TipoCbte$ = "003"
        Case "FCB": TipoCbte$ = "006"
        Case "NDB": TipoCbte$ = "007"
        Case "NCB": TipoCbte$ = "008"
        Case "FCE": TipoCbte$ = "019"
        Case "NDE": TipoCbte$ = "020"
        Case "NCE": TipoCbte$ = "021"
        Case "FCC": TipoCbte$ = "011"
        Case "NDC": TipoCbte$ = "012"
        Case "NCC": TipoCbte$ = "013"
        Case "RHA": TipoCbte$ = "004"
        Case "RHB": TipoCbte$ = "009"
        Case "RHC": TipoCbte$ = "015"
        Case "RHM": TipoCbte$ = "054"
        Case "CPA": TipoCbte$ = "002"
        Case "CPB": TipoCbte$ = "007"
        Case "CPC": TipoCbte$ = "012"
        Case "CPM": TipoCbte$ = "052"
        Case "FCM": TipoCbte$ = "051"
        Case "NDM": TipoCbte$ = "052"
        Case "NCM": TipoCbte$ = "053"
        Case "TFA": TipoCbte$ = "081"
        Case "TFB": TipoCbte$ = "082"
        Case "TFC": TipoCbte$ = "111"
        Case "TFM": TipoCbte$ = "118"
        Case "COA": TipoCbte$ = "001"
        Case "COB": TipoCbte$ = "006"
        Case "COE": TipoCbte$ = "019"
        Case "COC": TipoCbte$ = "011"
        Case "COM": TipoCbte$ = "051"
        Case "DPA": TipoCbte$ = "003"
        Case "DPB": TipoCbte$ = "008"
        Case "DPC": TipoCbte$ = "013"
        Case "DPM": TipoCbte$ = "053"
    End Select
    
    ' Presume que es un asiento contable
    If TipoCbte$ = "" And Mid$(NumeroComprobante$, 7, 1) = "-" Then TipoCbte$ = "039"
    
    TipoComprobante$ = TipoCbte$
    
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

Function NumeroCbte(NumeroComprobante$) As String

    NroComprobante$ = REPLACAR$(NumeroComprobante$, "-", "")
    NroComprobante$ = SOLONUM$(Mid$(NroComprobante$, 8, 8))
    
    If Len(TRIM$(NumeroComprobante$)) < 18 Then NroComprobante$ = TRIM$(REPLACAR$(NumeroComprobante$, "-", ""))
    
    Do While Len(NroComprobante$) < 20
        NroComprobante$ = "0" & NroComprobante$
    Loop
    
    NumeroCbte = NroComprobante$
    
End Function

Function PuntoVenta(NumeroComprobante$) As String

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
AB1$ = VALIMESAN%
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   If DERACCE%("SETUPACR", "Configuración de Funcionamiento ACREEDORES") = 2 Then
      FOLIVACO04.Show 1
   End If
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCHIG04/AMAPRO", "Archivos Maestros")
   Command15.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   If VALIMESAN% < 1 Then GoTo 99
   OPX% = ALTERNA%("Base Subdiario\Analisis por Cuenta\Analisis por Proveedor\Analisis por Centro de Costos")
   If OPX% < 1 Or OPX% > 4 Then GoTo 99
   '
   Call GENSUDICOM
   '
   Select Case OPX%
   Case 1
     If Option1.Value = True Then
          Call CONECRUN("*COSUCOM", "Consulta por Cuenta Contable")
       ElseIf Option2.Value = True Then
          Call CONECRUN("*SORANAI", "Listado Analitico de Imputacion de Compras")
     End If
   Case 2
     If Option1.Value = True Then
          Call CONECRUN("*ANAICOM", "Analisis por Cuenta Contable")
       Else
          Call CONECRUN("*SORANAI", "Listado Analitico de Imputacion de Compras")
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
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "AND OPCIMOVIM = 'OP' "
    SQLX$ = SQLX$ + "Order by FechConta ASC, CodiCom_Lar ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    J1& = 0: J2& = 0: REBLA$ = REGBLAN$("DETPAGC")
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
        NUMEFA$ = !CodiCom_Lar
        NCLIE% = !Nuprove
        FECHAD = Int(CDbl(!FechConta))
        INETO# = !ValAbsComp
        ITOTA# = !ValAbsComp
        '
        Call REPLA(ZZ$, Left$(TRIM$(!TipoMovim), 2), 1, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 3, 2)
        Call REPLA(ZZ$, !CodiCom_Lar, 5, 18)
        Call REPLA(ZZ$, MKI$(!Nuprove), 23, 2)
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
      .MoveNext
      Loop
    End With
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
'

Private Sub Command20_Click()
   '
   Command20.Enabled = False
   If VALIMESAN% < 1 Then GoTo 99
   '
   Call GENSUDICOM
   '
   JJ& = 0
   TTXXY$ = Chr$(1) + Chr$(2) + MKI$(HASTA%)
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
    RESP1% = ALTERNA%("Selección de Comprobantes a Listar.\Listado Libro de IVA - Compras.")
     If RESP1% < 1 Or RESP1% > 2 Then
       GoTo 99
     End If
    End If
    If RESP1% = 2 Then
        '\\\OT-05-0264-WAR-FIN///
        Call GENIVACOM
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
      Else
        Call EXCLUIVA
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
    DESTIDOC% = 0
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

Private Sub CITICOMP() ' ESTA VERSION AGRUPA LOS COMPROBANTES MENORES A $ 500.-
                       ' Y SUPRIME LOS CRÉDITOS
     '
     Dim TIDOCU$(2048), NUDOCU$(2048), FEDOCU$(2048), CUIPROX$(2048), DENPROX$(2048), IMPOIVAX#(2048)
     
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
       ' la condicion anterior se suprimio para tener en cuenta las notas de credito
       ' y debitos propios y elcaso de facturas con parte gravada y parte no gravada
       ' Luego al 8 de octubre de 2007 Claudia Matsusaka pidio que lo volvamos a poner
       ' como antes, que solo toma las facturas y debitos, no las notas de credito
       'C.U.I.T. DEL PROVEEDOR
       CUIPRO1$ = Mid$(XX$, 63, 16)
       CUIPRO2$ = REPLACAR$(CUIPRO1$, "-", "") 'CUIT DEL VENDEDOR
       CUIPRO3$ = TRIM$(CUIPRO2$)
       CUIEM3$ = CUIPRO3$
       'DENOMINACION DEL PROVEEDOR
       DENO2$ = Mid$(XX$, 33, 30)
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
25     KI% = KI% + 1
       If U& + KI% <= URE& Then
         YY$ = REGLEIDO$("LIVACOM", U& + KI%)
         If TRIM$(Mid$(YY$, 1, 99)) = "" Then
           TOTIVA# = TOTIVA# + CVD(Mid$(YY$, 145, 8))
           GoTo 25
         End If
       End If
       IMPOL3# = TOTIVA#
       '
       If IMPOL3# > 0 Then ' SI IMPORTE I.V.A. es positivo para suprimir los creditos
         
         If IMPOL3# >= 500 Then
            ' nunca se agrupan los mayores a 500 $
            CAREGI% = CAREGI% + 1
            TIDOCU$(CAREGI%) = TIPO3$
            NUDOCU$(CAREGI%) = NUMEFAC3$
            FEDOCU$(CAREGI%) = FEXU3$
            CUIPROX$(CAREGI%) = CUIPRO3$
            DENPROX$(CAREGI%) = DENO3$
            IMPOIVAX#(CAREGI%) = IMPOL3#
            GoTo 29
         End If
         ' los menores a 500 $ se agrupan
         For CARE1% = 1 To CAREGI%
           If CUIPROX$(CARE1%) = CUIPRO3$ Then
              ' encontro un registro anterior
              If NUDOCU$(CARE1%) <> String$(24, "0") Then
                 If IMPOIVAX#(CARE1%) >= 500 Then
                    GoTo 27     ' BUSCA EL SIGUENTE
                 End If
              End If
              IMPOIVAX#(CARE1%) = IMPOIVAX#(CARE1%) + IMPOL3#
              GoTo 29
           End If
27       Next CARE1%
         ' no encontro registro anterior
28       CAREGI% = CAREGI% + 1
         TIDOCU$(CAREGI%) = "00"
         NUDOCU$(CAREGI%) = String$(24, "0")
         FEDOCU$(CAREGI%) = "00000000"
         CUIPROX$(CAREGI%) = CUIPRO3$
         DENPROX$(CAREGI%) = DENO3$
         IMPOIVAX#(CAREGI%) = IMPOL3#
         '
       End If
29   Next U&
     '
     For CARE2% = 1 To CAREGI%
       If Abs(IMPOIVAX#(CARE2%)) >= 0.005 Then
         '
         TTXX$ = Space$(130)
         Call REPLA(TTXX$, TIDOCU$(CARE2%), 1, 2)
         Call REPLA(TTXX$, NUDOCU$(CARE2%), 3, 24)
         Call REPLA(TTXX$, FEDOCU$(CARE2%), 27, 8)
         Call REPLA(TTXX$, CUIPROX$(CARE2%), 35, 11)
         Call REPLA(TTXX$, DENPROX$(CARE2%), 46, 25)
           IMPOL# = Abs(IMPOIVAX#(CARE2%))
           IMPOX1$ = FORMATNUM$(IMPOL#, "F12.2")
           IMPOX2$ = REPLACAR$(IMPOX1$, ".", "")
           IMPOX3$ = TRIM$(IMPOX2$)
           While Len(IMPOX3$) < 12: IMPOX3$ = "0" + IMPOX3$: Wend
         Call REPLA(TTXX$, IMPOX3$, 71, 12)
         Print #NX%, TTXX$
         '
       End If
     Next CARE2%
     '
     SUMAT# = 0
     For KKU% = 1 To CAREGI%
       SUMAT# = SUMAT# + IMPOIVAX#(KKU%)
     Next KKU%
     '
     Close #NX%
     ASALID1$ = REPLACAR$(ASALID$, "\", "/")
     Call COMUNI("Archivo de Transferencia '" + ASALID1$ + "'\Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.\El Importe Total del IVA es de " + TRIM$(FORMATNUM$(SUMAT#, "F12.2")))
99   Screen.MousePointer = 1
     '
End Sub

Private Sub CITICOMP_VIE1() 'C.I.T.I.COMPRAS
  'ESTA VERSION SUPRIME CREDITOS Y EXPRESA LOS MENORES DE 500 $ SIN AGRUPAR
  ' SI ES UN ÚNICO COMPROBANTE. BOCHADO POR CLAUDIA MATSUSAKA EL 9/10/07
     '
     Dim TIDOCU$(2048), NUDOCU$(2048), FEDOCU$(2048), CUIPROX$(2048), DENPROX$(2048), IMPOIVAX#(2048)
     
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
       If CVD(Mid$(XX$, 145, 8)) > 0 Then ' SI IMPORTE I.V.A. es positivo
       ' la condicion anterior se suprimio para tener en cuenta las notas de credito
       ' y debitos propios y elcaso de facturas con parte gravada y parte no gravada
       ' Luego al 8 de octubre de 2007 Claudia Matsusaka pidio que lo volvamos a poner
       ' como antes, que solo toma las facturas y debitos, no las notas de credito
         'C.U.I.T. DEL PROVEEDOR
         CUIPRO1$ = Mid$(XX$, 63, 16)
         CUIPRO2$ = REPLACAR$(CUIPRO1$, "-", "") 'CUIT DEL VENDEDOR
         CUIPRO3$ = TRIM$(CUIPRO2$)
         CUIEM3$ = CUIPRO3$
         'DENOMINACION DEL PROVEEDOR
         DENO2$ = Mid$(XX$, 33, 30)
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
25       KI% = KI% + 1
         If U& + KI% <= URE& Then
           YY$ = REGLEIDO$("LIVACOM", U& + KI%)
           If TRIM$(Mid$(YY$, 1, 99)) = "" Then
             TOTIVA# = TOTIVA# + CVD(Mid$(YY$, 145, 8))
             GoTo 25
           End If
         End If
         IMPOL3# = TOTIVA#
         '
         If TIPO3$ = "03" Or TIPO3$ = "08" Then GoTo 28 ' NUNCA AGRUPA CRÉDITOS - 05/10/07 CON CLAUDIA MATSUSAKA
         ' GoTo 28   ' PEDIDO POR CLAUDIA MATSUSAKA EL 7/10/07 para nunca agrupar facturas
         ' Luego el 8 de octubre de 2007 C.M. pidió suprimir y volver a agrupar
         '
         If IMPOL3# < 500 Then
           For CARE1% = 1 To CAREGI%
             If TIDOCU$(CARE1%) = "03" Or TIDOCU$(CARE1%) = "08" Then GoTo 27
             '  If NUDOCU$(CARE1%) = String$(24, "0") Then
             If CUIPROX$(CARE1%) = CUIPRO3$ Then
                ' encontro un registro anterior
                If NUDOCU$(CARE1%) <> String$(24, "0") Then
                   If IMPOIVAX#(CARE1%) >= 500 Then
                      GoTo 27     ' BUSCA EL SIGUENTE
                   End If
                End If
                IMPOIVAX#(CARE1%) = IMPOIVAX#(CARE1%) + IMPOL3#
                TIDOCU$(CARE1%) = "00"
                NUDOCU$(CARE1%) = String$(24, "0")
                FEDOCU$(CARE1%) = "00000000"
                GoTo 29
             End If
27         Next CARE1%
         End If
         '
28       CAREGI% = CAREGI% + 1
         TIDOCU$(CAREGI%) = TIPO3$
         NUDOCU$(CAREGI%) = NUMEFAC3$
         FEDOCU$(CAREGI%) = FEXU3$
         CUIPROX$(CAREGI%) = CUIPRO3$
         DENPROX$(CAREGI%) = DENO3$
         IMPOIVAX#(CAREGI%) = IMPOL3#
       End If
29   Next U&
     '
     For CARE2% = 1 To CAREGI%
       If Abs(IMPOIVAX#(CARE2%)) >= 0.005 Then
         '
         TTXX$ = Space$(130)
         Call REPLA(TTXX$, TIDOCU$(CARE2%), 1, 2)
         Call REPLA(TTXX$, NUDOCU$(CARE2%), 3, 24)
         Call REPLA(TTXX$, FEDOCU$(CARE2%), 27, 8)
         Call REPLA(TTXX$, CUIPROX$(CARE2%), 35, 11)
         Call REPLA(TTXX$, DENPROX$(CARE2%), 46, 25)
           IMPOL# = Abs(IMPOIVAX#(CARE2%))
           IMPOX1$ = FORMATNUM$(IMPOL#, "F12.2")
           IMPOX2$ = REPLACAR$(IMPOX1$, ".", "")
           IMPOX3$ = TRIM$(IMPOX2$)
           While Len(IMPOX3$) < 12: IMPOX3$ = "0" + IMPOX3$: Wend
         Call REPLA(TTXX$, IMPOX3$, 71, 12)
         Print #NX%, TTXX$
         '
       End If
     Next CARE2%
     '
     SUMAT# = 0
     For KKU% = 1 To CAREGI%
       SUMAT# = SUMAT# + IMPOIVAX#(KKU%)
     Next KKU%
     '
     Close #NX%
     ASALID1$ = REPLACAR$(ASALID$, "\", "/")
     Call COMUNI("Archivo de Transferencia '" + ASALID1$ + "'\Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.\El Importe Total del IVA es de " + TRIM$(FORMATNUM$(SUMAT#, "F12.2")))
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
    Case "TF"
      If VAR1$ = "A" Then TICITICOM$ = "81"
  End Select
  '
End Function
'\\\OT-06-0271-WAR-FIN///

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
    Command3.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "AND OpciMovim = 'FC' "
    SQLX$ = SQLX$ + "Order by FechConta ASC "
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    SQLY$ = "SELECT * FROM CAJABANC "
    SQLY$ = SQLY$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLY$ = SQLY$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    Set CABUSMEN = CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot)
    '
    JJ& = 0: REBLA$ = REGBLAN$("DETFACC")
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
        KUL& = KUL& + 1
        Call TRACE(20, KUL&, FIN&)
        ZZ$ = REBLA$
        NUMEFA$ = !CodiCom_Lar
        NCLIE% = !Nuprove
        FECHAD = Int(CDbl(!FechConta))
        '
        INETO# = 0
        IIVA# = 0
        IVADU# = 0
        RETEN# = 0
        IOTRO# = 0
        ITOTA# = !ValAbsComp
        '
        SQLX$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
        SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX$ = SQLX$ + "AND NuProve = " & NCLIE% & " "
        '
        With CABUSMEN
          .FindFirst SQLX$
25        If .NoMatch = False Then
            TIMOV$ = TRIM$(UCase$(!TipoMovim))
            IPARD# = !ValAbsComp
            If TIMOV$ = "COMPRA" Then
                INETO# = INETO# + IPARD#
              ElseIf TIMOV$ = "IVA" And (UCase$(!VariMovim) = "ADUANA" Or UCase$("!VariMovim") = "IMPORT") Then
                IVADU# = IVADU# + IPARD#
              ElseIf TIMOV$ = "IVA" And UCase$(!VariMovim) = "PER-GAN" Then
                RETEN# = RETEN# + IPARD#
              ElseIf TIMOV$ = "IVA" Then
                IIVA# = IIVA# + IPARD#
            End If
            .FindNext SQLX$
            GoTo 25
          End If
        End With
        '
        IOTRO# = ITOTA# - INETO# - IIVA# - IVADU# - RETEN#
        '
        Call REPLA(ZZ$, Left$(TRIM$(!TipoMovim), 2), 1, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 3, 2)
        Call REPLA(ZZ$, !CodiCom_Lar, 5, 18)
        Call REPLA(ZZ$, MKI$(!Nuprove), 23, 2)
        Call REPLA(ZZ$, MKD$(INETO#), 25, 8)
        Call REPLA(ZZ$, MKD$(IIVA#), 33, 8)
        Call REPLA(ZZ$, MKD$(IVADU#), 41, 8)
        Call REPLA(ZZ$, MKD$(RETEN#), 49, 8)
        Call REPLA(ZZ$, MKD$(IOTRO#), 57, 8)
        Call REPLA(ZZ$, MKD$(ITOTA#), 65, 8)
        JJ& = JJ& + 1
        Call GRAREG("DETFACC", ZZ$, JJ&)
        '
      .MoveNext
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
99  Command3.Enabled = True
    Screen.MousePointer = 1
End Sub

Sub Command4_Click()
    Command4.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "AND OpciMovim = 'ND' "
    SQLX$ = SQLX$ + "Order by FechConta ASC "
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    SQLY$ = "SELECT * FROM CAJABANC "
    SQLY$ = SQLY$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLY$ = SQLY$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    Set CABUSMEN = CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot)
    '
    JJ& = 0: REBLA$ = REGBLAN$("DETFACC")
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
        NUMEFA$ = !CodiCom_Lar
        NCLIE% = !Nuprove
        FECHAD = Int(CDbl(!FechConta))
        '
        INETO# = 0
        IIVA# = 0
        IVADU# = 0
        RETEN# = 0
        IOTRO# = 0
        ITOTA# = !ValAbsComp
        '
        SQLX$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
        SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX$ = SQLX$ + "AND NuProve = " & NCLIE% & " "
        '
        With CABUSMEN
          .FindFirst SQLX$
25        If .NoMatch = False Then
            TIMOV$ = TRIM$(UCase$(!TipoMovim))
            IPARD# = !ValAbsComp
            If TIMOV$ = "COMPRA" Then
                INETO# = INETO# + IPARD#
              ElseIf TIMOV$ = "IVA" And (UCase$(!VariMovim) = "ADUANA" Or UCase$("!VariMovim") = "IMPORT") Then
                IVADU# = IVADU# + IPARD#
              ElseIf TIMOV$ = "IVA" And UCase$(!VariMovim) = "PER-GAN" Then
                RETEN# = RETEN# + IPARD#
              ElseIf TIMOV$ = "IVA" Then
                IIVA# = IIVA# + IPARD#
            End If
            .FindNext SQLX$
            GoTo 25
          End If
        End With
        '
        IOTRO# = ITOTA# - INETO# - IIVA# - IVADU# - RETEN#
        '
        Call REPLA(ZZ$, Left$(TRIM$(!TipoMovim), 2), 1, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 3, 2)
        Call REPLA(ZZ$, !CodiCom_Lar, 5, 18)
        Call REPLA(ZZ$, MKI$(!Nuprove), 23, 2)
        Call REPLA(ZZ$, MKD$(INETO#), 25, 8)
        Call REPLA(ZZ$, MKD$(IIVA#), 33, 8)
        Call REPLA(ZZ$, MKD$(IVADU#), 41, 8)
        Call REPLA(ZZ$, MKD$(RETEN#), 49, 8)
        Call REPLA(ZZ$, MKD$(IOTRO#), 57, 8)
        Call REPLA(ZZ$, MKD$(ITOTA#), 65, 8)
        JJ& = JJ& + 1
        Call GRAREG("DETFACC", ZZ$, JJ&)
        '
      .MoveNext
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
99  Command4.Enabled = True
    Screen.MousePointer = 1
End Sub


Sub Command5_Click()
    Command5.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "AND OpciMovim = 'NC' "
    SQLX$ = SQLX$ + "Order by FechConta ASC "
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    SQLY$ = "SELECT * FROM CAJABANC "
    SQLY$ = SQLY$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLY$ = SQLY$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    Set CABUSMEN = CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot)
    '
    JJ& = 0: REBLA$ = REGBLAN$("DETFACC")
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
        NUMEFA$ = !CodiCom_Lar
        NCLIE% = !Nuprove
        FECHAD = Int(CDbl(!FechConta))
        '
        INETO# = 0
        IIVA# = 0
        IVADU# = 0
        RETEN# = 0
        IOTRO# = 0
        ITOTA# = !ValAbsComp
        '
        SQLX$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
        SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX$ = SQLX$ + "AND NuProve = " & NCLIE% & " "
        '
        With CABUSMEN
          .FindFirst SQLX$
25        If .NoMatch = False Then
            TIMOV$ = TRIM$(UCase$(!TipoMovim))
            IPARD# = !ValAbsComp
            If TIMOV$ = "COMPRA" Then
                INETO# = INETO# + IPARD#
              ElseIf TIMOV$ = "IVA" And (UCase$(!VariMovim) = "ADUANA" Or UCase$("!VariMovim") = "IMPORT") Then
                IVADU# = IVADU# + IPARD#
              ElseIf TIMOV$ = "IVA" And UCase$(!VariMovim) = "PER-GAN" Then
                RETEN# = RETEN# + IPARD#
              ElseIf TIMOV$ = "IVA" Then
                IIVA# = IIVA# + IPARD#
            End If
            .FindNext SQLX$
            GoTo 25
          End If
        End With
        '
        IOTRO# = ITOTA# - INETO# - IIVA# - IVADU# - RETEN#
        '
        Call REPLA(ZZ$, Left$(TRIM$(!TipoMovim), 2), 1, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 3, 2)
        Call REPLA(ZZ$, !CodiCom_Lar, 5, 18)
        Call REPLA(ZZ$, MKI$(!Nuprove), 23, 2)
        Call REPLA(ZZ$, MKD$(INETO#), 25, 8)
        Call REPLA(ZZ$, MKD$(IIVA#), 33, 8)
        Call REPLA(ZZ$, MKD$(IVADU#), 41, 8)
        Call REPLA(ZZ$, MKD$(RETEN#), 49, 8)
        Call REPLA(ZZ$, MKD$(IOTRO#), 57, 8)
        Call REPLA(ZZ$, MKD$(ITOTA#), 65, 8)
        JJ& = JJ& + 1
        Call GRAREG("DETFACC", ZZ$, JJ&)
        '
      .MoveNext
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
99  Command5.Enabled = True
    Screen.MousePointer = 1
End Sub

Sub Command6_Click()
    Command6.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Screen.MousePointer = 11
    Dim SALINI#(32767), SUFAC#(32767), SUPAG#(32767)
    '
    Call APERBASDAT("CABAN")
    TABORI$ = "CAJABANC"
    '
    CLIMAX% = 0
    '
10  SQLX$ = "SELECT * FROM " + TABORI$ + " "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "or TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUP' "
    '
    On Error Resume Next
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    If Err Then GoTo 60
    On Error GoTo 0
    '
    JU& = 0
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
      Do While Not .EOF
        JU& = JU& + 1
        Call TRACE(20, JU&, FIN&)
        FEMII% = CVINT(!FechConta)
        If FEMII% <= HASTA% Then
          NCLIE% = !Nuprove
          If NCLIE% > 0 Then
            If TICUEPRO((-1) * NCLIE%) = 0 Then ' solo cuentas corrientes
              If NCLIE% > CLIMAX% Then CLIMAX% = NCLIE%
              If TRIM$(!TipoMovim) = "FCOM" Or TRIM$(!TipoMovim) = "AJUP" Then
                  IFAC# = !IHABECOMP - !Idebecomp
                  IPAG# = 0
                Else
                  IFAC# = 0
                  IPAG# = !Idebecomp - !IHABECOMP
              End If
              '
              If FEMII% < DESDE% Then
                  SALINI#(NCLIE%) = SALINI#(NCLIE%) + IFAC# - IPAG#
                Else
                  SUFAC#(NCLIE%) = SUFAC#(NCLIE%) + IFAC#
                  SUPAG#(NCLIE%) = SUPAG#(NCLIE%) + IPAG#
              End If
            End If
          End If
        End If
      .MoveNext
      Loop
    End With
    '
60  On Error GoTo 0
    If TABORI$ = "CAJABANC" Then
      TABORI$ = "CABANVIE"
      GoTo 10
    End If
    '
    JJ& = 0
    For K% = 1 To CLIMAX%
      If Abs(SALINI#(K%)) >= 0.005 Or Abs(SUFAC#(K%)) >= 0.005 Or Abs(SUPAG#(K%)) >= 0.005 Then
        X$ = REGBLAN$("RESUCCP")
        Call REPLA(X$, MKI$(K%), 1, 2)
        Call REPLA(X$, MKD$(SALINI#(K%)), 3, 8)
        Call REPLA(X$, MKD$(SUFAC#(K%)), 11, 8)
        Call REPLA(X$, MKD$(SUPAG#(K%)), 19, 8)
        Call REPLA(X$, MKD$(SALINI#(K%) + SUFAC#(K%) - SUPAG#(K%)), 27, 8)
        JJ& = JJ& + 1
        Call GRAREG("RESUCCP", X$, JJ&)
       End If
    Next K%
    '
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
99 Command6.Enabled = True
   Screen.MousePointer = 1
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'AJUP' "
    'SQLX$ = SQLX$ + "AND OpciMovim = 'FC' "
    'SQLX$ = SQLX$ + "Order by FechConta ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
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
        '
        ZZ$ = REBLA$
        '
        IIVA# = 0
        IVADU# = 0
        RETEN# = 0
        IOTRO# = 0
        INETO# = !IHABECOMP - !Idebecomp
        ITOTA# = INETO#
        '
        Call REPLA(ZZ$, Left$(!TipoMovim, 2), 1, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 3, 2)
        Call REPLA(ZZ$, !CodiCom_Lar, 5, 18)
        Call REPLA(ZZ$, MKI$(!Nuprove), 23, 2)
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

Private Sub config_Click()
   Call Command14_Click
End Sub

Private Sub consultadeasientocontableporpantalla_Click()
    Call Command20_Click
End Sub

Private Sub estadisticodecompras_Click()
   Call Command24_Click
End Sub

Private Sub facturasdecompra_Click()
    Call Command3_Click
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.Picture16.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame7.Width = 9250
      Me.Frame5.Visible = False
      Me.Frame7.BackColor = &HFCD7B6
      Me.Picture11.Visible = True
'      Me.Height = 1800
      MARCOBARRA.Top = 5800
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If


    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Function VALIMESAN%()
   VMA% = 0
   '
   If Len(MESAN) = 19 Then
     DESDE% = FECHANUM(Left$(MESAN, 8))
     If DESDE% > 0 Then
       HASTA% = FECHANUM(Mid$(MESAN, 12, 8))
       If HASTA% >= DESDE% Then
         VMA% = 1
         DESDAT = Int(CDbl(CVDAT(DESDE%)))
         HASDAT = Int(CDbl(CVDAT(HASTA%)))
         CORRESP$ = MESAN
         GoTo 98
       End If
     End If
   End If
   '
   CORRESP$ = MESANO$(MESAN, DESDE%, HASTA%)
   If CORRESP$ <> "" Then
       VMA% = 1
       DESDAT = Int(CDbl(CVDAT(DESDE%)))
       HASDAT = Int(CDbl(CVDAT(HASTA%)))
     Else
       Call MENSERR(24, "Falta Elegir Período Mensual")
   End If
   '
98 VALIMESAN% = VMA%
   '
End Function

Private Sub Image2_Click()
   TXT$ = MESAN.TEXT
   If TXT$ = "" Then Exit Sub
   If MESANO$(TXT$, DES%, HAS%) = TXT$ Then
      If DERACCE%("CIEMECOM", "Cierre y Re-Apertura Mensual Compras") >= 2 Then
         If COMALTER%("Confirme Cierre del Período Mensual " + TXT$ + "\\Cancelar\Cerrar") = 2 Then
           Call GRACONTROL("", "CIEMECOM", TRIM$(Str$(HAS%)))
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
      If DERACCE%("CIEMECOM", "Cierre y Re-Apertura Mensual Compras") >= 2 Then
         If COMALTER%("Confirme Re-Apertura del Período Mensual\\Cancelar\Re-Abrir") = 2 Then
            Call GRACONTROL("", "CIEMECOM", TRIM$(Str$(DES% - 1)))
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

Private Sub impresora2_Click()
   Option2.Value = True
   Call Command2_Click

End Sub

Private Sub listadoporcuentacontable_Click()
    Call Command19_Click
End Sub

Private Sub maestroproveedores_Click()
   Call Command15_Click
End Sub

Private Sub MESAN_Change()
   TXT$ = MESAN.TEXT
   Image2.Visible = True
   Image1.Visible = False
   If TXT$ = "" Then Exit Sub
   If MESANO$(TXT$, DES%, HAS%) = TXT$ Then
     If XVALO(CONTROL("", "CIEMECOM")) >= HAS% Then
       Image2.Visible = False
       Image1.Visible = True
     End If
   End If
End Sub

Private Sub MESAN_DblClick()
   'Call Command12_Click
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   MESAN = PERIO$
   '
End Sub

Sub GENSUDICOM()
    '
    Screen.MousePointer = 11
    '
    Dim IMASI#(32767)
    CUEMX% = 0
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + "Order by FechConta ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    '
    Set CABAMENS = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Screen.MousePointer = 1
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
        NUMEFA$ = !CodiCom_Lar
        NCLIE% = !Nuprove
        FECHAD = Int(CDbl(!FechConta))
        ITOTA# = !ValAbsComp
        SIG% = Sgn(!IHABECOMP - !Idebecomp)
        MOBASI$ = TRIM$(!TipoMovim)
        SUMACO# = 0
        '
        ' LO QUE SIGUE ES PARA CONSIDERAR CLIENTES VARIOS
        RSCL$ = TRIM$(RASOCLI$((-1) * NCLIE%))
        CTCL$ = CUITCLI$((-1) * NCLIE%)
        If UCase$(Left$(RSCL$, 6)) = "VARIOS" Then
          REFEMO$ = !Refecom
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
          Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 1, 2)
          Call REPLA(ZZ$, !CodiCom_Lar, 3, 18)
          Call REPLA(ZZ$, RSCL$, 21, 30)
          Call REPLA(ZZ$, CTCL$, 51, 16)
          Call REPLA(ZZ$, MKD$(ITOTA# * SIG%), 67, 8)
          Call REPLA(ZZ$, CODCUE$(!CueContaI), 75, 12)
          Call REPLA(ZZ$, MKD$(!Idebecomp), 87, 8)
          Call REPLA(ZZ$, MKD$(!IHABECOMP), 95, 8)
          JJ& = JJ& + 1
          Call GRAREG("SUDICOM", ZZ$, JJ&)
        End If
        '
        SUMACO# = SUMACO# + !Idebecomp - !IHABECOMP
        '
        CUEII% = !CueContaI
        If CUEII% > 0 Then
          IMASI#(CUEII%) = IMASI#(CUEII%) + (!Idebecomp - !IHABECOMP)
          If CUEII% > CUEMX% Then CUEMX% = CUEII%
        End If
        '
        SQLX$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
        SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX$ = SQLX$ + "AND NuProve = " & NCLIE% & " "
        SQLX$ = SQLX$ + "AND TipoMovim <> '" & MOBASI$ & "' "
        SQLX$ = SQLX$ + "AND TipoMovim <> 'OPAG' "    ' PARA EVITAR PAGOS CONTADO
        SQLX$ = SQLX$ + "AND TipoMovim <> 'CHEMI' "   ' PARA EFITAR PAGOS CONTADO
        '
        With CABAMENS        ' CajaBanc
          .FindFirst SQLX$
25        If .NoMatch = False Then
            ZZ$ = REBLA$
            '
            If MOBASI$ = "FCOM" Then
              Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 1, 2)
              Call REPLA(ZZ$, !CodiCom_Lar, 3, 18)
              Call REPLA(ZZ$, RSCL$, 21, 30)
              Call REPLA(ZZ$, CTCL$, 51, 16)
              Call REPLA(ZZ$, CODCUE$(!CueContaI), 75, 12)
              Call REPLA(ZZ$, MKD$(!Idebecomp), 87, 8)
              Call REPLA(ZZ$, MKD$(!IHABECOMP), 95, 8)
              JJ& = JJ& + 1
              Call GRAREG("SUDICOM", ZZ$, JJ&)
            End If
            '
            SUMACO# = SUMACO# + !Idebecomp - !IHABECOMP
            '
            Z2$ = REBL2$
            Call REPLA(Z2$, MKI$(!CueContaI), 1, 2)
            Call REPLA(Z2$, MKI$(NCLIE%), 3, 2)
            Call REPLA(Z2$, !CodiCom_Lar, 5, 18)
            Call REPLA(Z2$, MKI$(CVINT(!FechConta)), 23, 2)
            Call REPLA(Z2$, MKD$(!Idebecomp - !IHABECOMP), 25, 8)
            J2& = J2& + 1
            Call GRAREG("ANAICOM", Z2$, J2&)
            '
            Z3$ = Space$(8) + Z2$
            CENCO$ = "": If IsNull(!CENCOSTO) = False Then CENCO$ = !CENCOSTO
            Call REPLA(Z3$, CENCO$, 1, 8)
            '\\\OT-06-0238-WAR-09/06/06///
            If TRIM$(CENCO$) <> "" Then
              J3& = J3& + 1
              Call GRAREG("ANCENCOS", Z3$, J3&)
            End If
            '\\\OT-06-0238-WAR-FIN///
            '
            CUEII% = !CueContaI
            If CUEII% > 0 Then
              IMASI#(CUEII%) = IMASI#(CUEII%) + (!Idebecomp - !IHABECOMP)
              If CUEII% > CUEMX% Then CUEMX% = CUEII%
            End If
            '
           .FindNext SQLX$
            GoTo 25
          End If
        End With
        '
        If Abs(SUMACO#) >= 0.005 Then
          Call MENSERR(24, "No Balancea Comprobante " + NUMEFA$ + "\de " + TRIM$(RASOCLI$((-1) * NCLIE%)) + "\  \Debe Anularse y Pasarse de Nuevo.")
        End If
        '
      .MoveNext
      Loop
    End With
    Call SETULTREG("SUDICOM", JJ&)
    Call SETULTREG("ANAICOM", J2&)
    '\\\OT-06-0238-WAR-09/06/06///'AGREGO CENTRO DE COSTOS POR OP
    
    SQLX1$ = "TipoMovim = 'OPAG' "
    SQLX1$ = SQLX1$ + "AND VariMovim = 'CGAST' "
    '
    With CABAMENS        ' CajaBanc
      .FindFirst SQLX1$
33    If .NoMatch = False Then
        Z3$ = REBL3$
        CENCO$ = ""
        If IsNull(!CENCOSTO) = False Then
          If TRIM$(!CENCOSTO) <> "" Then
            CENCO$ = !CENCOSTO
          End If
        End If
        Call REPLA(Z3$, CENCO$, 1, 8)
        If TRIM$(CENCO$) <> "" Then
          '
          Call REPLA(Z3$, MKI$(!CueContaI), 9, 2)
          Call REPLA(Z3$, MKI$(!Nuprove), 11, 2)
          Call REPLA(Z3$, !CodiCom_Lar, 13, 18)
          Call REPLA(Z3$, MKI$(CVINT(!FechConta)), 31, 2)
          Call REPLA(Z3$, MKD$(!Idebecomp - !IHABECOMP), 33, 8)
          '
          J3& = J3& + 1
          Call GRAREG("ANCENCOS", Z3$, J3&)
        End If
        .FindNext SQLX1$
        GoTo 33
      End If
    End With
    '
    ' AGREGA LOS ASIENTOS DE DIARIOS CON CENTRO DE COSTO
    Call APERBASDAT("PLANCUEN")
    SQLX$ = "SELECT * FROM DETAASIEN "
    SQLX$ = SQLX$ + "where INT(CDBL(FechAsien)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechAsien)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND STATPASE > 0" ' SE INCLUYE PARA NO CONSIDERAR LOS ASIENTOS ANULADOS
    Set ASIENDIAR = Contable.OpenRecordset(SQLX$, dbOpenSnapshot)
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
                 Call REPLA(Z3$, !NuIdAsien, 13, 18)
                 Call REPLA(Z3$, MKI$(CVINT(!FECHASIEN)), 31, 2)
                 Call REPLA(Z3$, MKD$(!IDebePase - !IHabePase), 33, 8)
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

    Call HEADERS("SUDICOM", "")
    Call HEADERS("SUDICOM", "Corresponde a: " + TRIM$(CORRESP$))
    Call HEADERS("ANAICOM", "")
    Call HEADERS("ANAICOM", "Corresponde a: " + TRIM$(CORRESP$))
    Call HEADERS("ANAICOL", "")
    Call HEADERS("ANAICOL", "Corresponde a: " + TRIM$(CORRESP$))
    Call HEADERS("ANCENCOS", "")
    Call HEADERS("ANCENCOS", "Corresponde a: " + TRIM$(CORRESP$))
    '
    Call FILESORT("ANCENCOS", "", 1, 5, "ASC")
    '
    TTXX$ = Chr$(1) + Chr$(2) + MKI$(HASTA%)
    V& = 0: FIN& = ULTREG&("ASIDIAR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("ASIDIAR", U&)
      If Left$(X$, 4) <> TTXX$ Then
        V& = V& + 1
        Call GRAREG("ASIDIAR", X$, V&)
      End If
    Next U&
    '
    AGRECU% = 0
    For K% = 1 To CUEMX%
      If Abs(IMASI#(K%)) >= 0.005 Then
        X$ = REGBLAN$("ASIDIAR")
        Call REPLA(X$, TTXX$, 1, 4)
        Call REPLA(X$, MKI$(K%), 5, 2)
        Call REPLA(X$, MKD$(IMASI#(K%)), 39, 8)
        V& = V& + 1
        Call GRAREG("ASIDIAR", X$, V&)
        '
        If ECOARCH$("CUECON", MKI$(K%)) = "" Then
          Z3$ = REGBLAN$("CUECON")
          Call REPLA(Z3$, MKI$(K%), 1, 2)
          Call REPLA(Z3$, DENOCUE$(K%), 3, 30)
          Call REPLA(Z3$, CODCUE$(K%), 33, 12)
          Call REPLA(Z3$, CODCUE$(K%), 45, 12)
          Call REGAPP("CUECON", Z3$)
          AGRECU% = 1
        End If
      End If
    Next K%
    '
    If AGRECU% > 0 Then
      Call FILESORT("CUECON", "", 1, 1, "ASC")
    End If
    Call CIERRARCH("CUECON")
    '
99  Screen.MousePointer = 1
    '
End Sub

Sub GENIVACOMVIE()
    Screen.MousePointer = 11
    '
    Dim RELIVA$(32)
    Dim IMASI#(32767)
    CUEMX% = 0
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'FCOM' "
    '\\\OT-05-0264-WAR-09/05/05///
    SQLX$ = SQLX$ + "AND (IncSubIva <> 'NO' "
    SQLX$ = SQLX$ + "OR isnull(IncSubIva) = True OR isempty(IncSubIva) = True) "
    '\\\OT-05-0264-WAR-FIN///
    SQLX$ = SQLX$ + "Order by FechConta ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    REBLA$ = REGBLAN$("LIVACCAI")
    REBLB$ = REGBLAN$("REIVACOM")
    '\\\OT-05-0710-WAR-07/11/05///
      Call SETULTREG("LIVACCAI", 0)
      Call SETULTREG("REIVACOM", 0)
    '\\\OT-05-0710-WAR-FIN///
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
        NUMEFA$ = !CodiCom_Lar
        NCLIE% = !Nuprove
        FECHAD = Int(CDbl(!FechConta))
        ITOTA# = !ValAbsComp
        SIG% = Sgn(!IHABECOMP - !Idebecomp)
        MOBASI$ = TRIM$(!TipoMovim)
        '
        Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 3, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!FechEmisi)), 5, 2)
        Call REPLA(ZZ$, MKS$(!NumeComp), 7, 4)
        Call REPLA(ZZ$, !CodiCom_Lar, 11, 18)
        Call REPLA(ZZ$, MKI$(NCLIE%), 31, 2)
        Call REPLA(ZZ$, RASOCLI$((-1) * NCLIE%), 33, 30)
        Call REPLA(ZZ$, CUITCLI$((-1) * NCLIE%), 63, 16)
        Call REPLA(ZZ$, POSIVAC$((-1) * NCLIE%), 79, 12)
        '\\\OT-05-0264-WAR-09/05/05///
        If IsNull(!NROCAI) Then
           NCAI$ = ""
          Else
           NCAI$ = !NROCAI
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
          If UCase$(Mid$(ZZ$, 33, 6)) = "VARIOS" Then
            REFEMO$ = !Refecom
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
        '
        '
        RELIVA$(0) = ZZ$
        For KU% = 1 To 32
          RELIVA$(KU%) = REBLA$
        Next KU%
        '
        SQLX$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
        SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX$ = SQLX$ + "AND NuProve = " & NCLIE% & " "
        SQLX$ = SQLX$ + "AND TipoMovim <> '" & MOBASI$ & "' "
        '
        PROIVA% = 0: ULIVA% = 0
        With CajaBanc
          .FindFirst SQLX$
25        If .NoMatch = False Then
            TIMO$ = UCase$(TRIM$(!TipoMovim))
            OPMO$ = !OpciMovim
            VAMO$ = !VariMovim
            IPAX# = Abs(!ValAbsComp)
            BASE# = 0
            On Error Resume Next
            BASE# = Abs(!BASEIMPO)
            On Error GoTo 0
            '
            If Abs(IPAX#) >= 0.005 Or BASE# >= 0.005 Then
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
                  Call REPLA(ZZ$, MKI$(!CueContaI), 121, 2)
                    JURIB$ = TRIM$(Mid$(OPMO$, 4, 1))
                  Call REPLA(ZZ$, JURIB$, 203, 1)
                  Call REPLA(ZZ$, ECOARCH$("PROVINC", JURIB$), 205, 24)
                  RELIVA$(0) = ZZ$
                ElseIf TIMO$ = "IVA" Then
                  ZZ$ = RELIVA$(PROIVA%)
                  Call REPLA(ZZ$, VAMO$, 129, 8)
                  Call REPLA(ZZ$, MKD$(BASE# * SIG%), 137, 8)
                  Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 145, 8)
                  If BASE# > 0.005 Then
                    Call REPLA(ZZ$, MKD$(100 * IPAX# / BASE#), 153, 8)
                  End If
                  RELIVA$(PROIVA%) = ZZ$
                  ULIVA% = PROIVA%
                  PROIVA% = PROIVA% + 1
              End If
            End If
            '
            .FindNext SQLX$
            GoTo 25
            '
          End If
        End With
        '
        For KU% = 0 To ULIVA%
          JJ& = JJ& + 1
          Call GRAREG("LIVACOM", RELIVA$(KU%), JJ&)
          Call GRAREG("LIVACCAI", RELIVA$(KU%), JJ&)
        Next KU%
        '
      .MoveNext
      Loop
    End With
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
      DECATE$ = TRIM$(Mid$(RECUE$, 129, 8)) + " - " + Mid$(RECUE$, 79, 12)
      For KLI% = 1 To CACATIV%
        If DECATE$ = DECATIV$(KLI%) Then
          COCATE% = KLI%
          GoTo 20
        End If
      Next KLI%
      CACATIV% = CACATIV% + 1
      DECATIV$(CACATIV%) = DECATE$
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
84       BASEIMPO# = BASEIMPO# + CVD(Mid$(VCC$, 65, 8))
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
    Call HEADERS("REIVACOM", "")
    Call HEADERS("REIVACOM", "Corresponde a: " + TRIM$(CORRESP$))
    '
99  Screen.MousePointer = 1
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
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "AND NuProve = " & NC% & " "
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
          NCLIE% = !Nuprove
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!CodiCom_Lar, 13, 6))
            If Len(!CodiCom_Lar) < 18 Then
              NPX& = XVALO(Right$(!CodiCom_Lar, 6))
            End If
            Y$ = MKS$(NPX&) + AJUSTI$(!CodiCom_Lar, 18) + IMPOS$
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
       Call MENSERR(24, "La Cuenta Elegida no Registra\Facturas de Compra que Puedan Seleccionarse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUECO")
    Call FRESHECHO("!SELCUECO")
    '
15  Call SELECHO("!SELCUECO")
    If VALACT1$("!SELCUECO") = "" And VALACT2$("!SELCUECO") = "" Then GoTo 10
    'REGCUEC& = XVALO(VALACT1$("!SELCUECO"))
    'If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = Left$(VALACT2$("!SELCUECO"), 18)
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    '
    Call APERBASDAT("CABAN")
    TTXX$ = "Libro de I.V.A. Compras\  \Comprobante '" + TRIM$(DOCULA$) + "' \de Cuenta '" + TRIM$(RASOCLI$((-1) * NC%)) + "' ?"
    If COMALTER%(TTXX$ + "\\Incluye\Excluye de Listado") = 1 Then
      COINC = 1
    End If
    With CajaBanc
      .FindFirst "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
20     If .NoMatch = False Then
        .Edit
         If COINC = 1 Then
           !IncSubIva = ""
         Else
           !IncSubIva = "NO"
         End If
        .Update
        .FindNext "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & ""
        GoTo 20
      End If
    End With
    '
    GoTo 15
    '
End Sub

Sub RANKPRO()
Dim TOTCOMPRA#(32767)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
20  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    VDEF$ = OBX$
    '
    Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTCOMPRA#(U&) = 0
    Next U&
    '
    DESDAT = Int(CDbl(CVDAT(DES%)))
    HASDAT = Int(CDbl(CVDAT(HAS%)))
    '
    Call APERBASDAT("COMPRAS")
    SQLX$ = "SELECT * FROM OCOMENCA "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(Femi_Ocom)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(Femi_Ocom)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
    '
    Set CABATEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
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
         NPRO% = !Npro_Ocom
         If NPRO% > 0 Then
           If NPRO% > PROMAX% Then PROMAX% = NPRO%
           TOTCOMPRA#(NPRO%) = TOTCOMPRA#(NPRO%) + (!TOBRU_OCOM * TICAMONE(!Mone_Emis) / 1)
         End If
         .MoveNext
      Loop
    End With
    '
    For I% = 1 To PROMAX
      TOTAGEN# = TOTAGEN# + TOTCOMPRA#(I%)
    Next I%
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
        Call GRAREG("RANKCLI", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKCLI", J&)
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
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
20  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    VDEF$ = OBX$
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
    SQLX$ = SQLX$ + "AND TipoMovim = 'COMPRA' "
    '
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
         NCLIE% = !Nuprove
         If NCLIE% > 0 Then
           If NCLIE% > CLIMAX% Then CLIMAX% = NCLIE%
           TOTVENTA#(NCLIE%) = TOTVENTA#(NCLIE%) + (!Idebecomp - !IHABECOMP)
         End If
         .MoveNext
      Loop
    End With
    '
    For I% = 1 To CLIMAX
      TOTAGEN# = TOTAGEN# + TOTVENTA#(I%)
    Next I%
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
        Call GRAREG("RANKCLI", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKCLI", J&)
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
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    '
    Screen.MousePointer = 11
    '
    DESDAT = Int(CDbl(CVDAT(DES%)))
    HASDAT = Int(CDbl(CVDAT(HAS%)))
    '
'2.- CONSULTA SQL
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC"
    SQLX$ = SQLX$ + " WHERE TipoMovim = 'COMPRA' "
    SQLX$ = SQLX$ + " AND   CDBL(FechConta) >= " & DESDAT & ""
    SQLX$ = SQLX$ + " AND   CDBL(FechConta) <= " & HASDAT & ""
    Set LIFACOTMP = CueCorri.OpenRecordset(SQLX$, 4)
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
        NPX1% = !Nuprove
          FEX = CVINT(!FechConta)
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
    '
    Screen.MousePointer = 1
    '4.-AQUI VA A GENERAR ARCHIVO
    Y$ = REGBLAN$("LIESFACO")
    JJ& = 0
    For U% = 1 To CAREG1%
      Call REPLA(Y$, PERI1$(U%), 1, 8)
      Call REPLA(Y$, MKI$(NPROV1%(U%)), 9, 2)
      Call REPLA(Y$, MKD$(TOTIMP#(U%)), 11, 8)
      JJ& = JJ& + 1
      Call GRAREG("LIESFACO", Y$, JJ&)
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
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
20  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    VDEF$ = OBX$
    '
    Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTART#(U&) = 0
    Next U&
    '
    DESDAT = Int(CDbl(CVDAT(DES%)))
    HASDAT = Int(CDbl(CVDAT(HAS%)))
    '
    Call APERBASDAT("COMPRAS")
    SQLX$ = "SELECT * FROM OCOMDETA "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(Femi_Ocom)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(Femi_Ocom)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
    '
    Set CABATEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
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
         NART% = !COD_INTE
         If NART% > 0 Then
           If NART% > PROMAX% Then PROMAX% = NART%
           TOTART#(NART%) = TOTART#(NART%) + (!INet_Item * TICAMONE(!Mone_Emis) / 1)
         End If
         .MoveNext
      Loop
    End With
    '
    For I% = 1 To PROMAX
      TOTAGEN# = TOTAGEN# + TOTART#(I%)
    Next I%
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
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call GRAREG("RANKPIE", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKPIE", J&)
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
    Call APERBASDAT("PLANCUEN")
    SQLX$ = "SELECT * FROM DetaAsien "
    SQLX$ = SQLX$ + " WHERE CDBL(Fechasien) >= " & DESDAT
    SQLX$ = SQLX$ + " AND CDBL(Fechasien) <= " & HASDAT
    SQLX$ = SQLX$ + " AND statpase >0  "
    SQLX$ = SQLX$ + "ORDER BY Fechasien "
    Set DETACONTA = Contable.OpenRecordset(SQLX$, 4)
    '
    With DETACONTA
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        GoTo 19
      End If
      FIN& = .RecordCount
      .MoveFirst
      On Error GoTo 0
      U& = 0
      Do While Not .EOF
        U& = U& + 1
        Call TRACE(20, U&, FIN&)
        NUMEFA$ = AJUSTI$(!NuIdAsien, 16)
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
      Encasien.FindFirst "NuidAsien = '" & NUMEFA$ & "' AND StatAsien > 0"
      If Encasien.NoMatch = False Then
        If Encasien!TipoAsien < 1 Then
          CANASIP% = CANASIP% + 1
          NUMEASI$(CANASIP%) = NUMEASI$(IXX%)
          FECHASI%(CANASIP%) = FECHASI%(IXX%)
        End If
      End If
    Next IXX%
    CANASIX% = CANASIP%
    '
22  Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    Set CABAFETMP = CueCorri.OpenRecordset(SQLX$, 4)
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'FCOM' "
    '\\\OT-05-0264-WAR-09/05/05///
    SQLX$ = SQLX$ + "AND (IncSubIva <> 'NO' "
    SQLX$ = SQLX$ + "OR isnull(IncSubIva) = True OR isempty(IncSubIva) = True) "
    '\\\OT-05-0264-WAR-FIN///
    '\\\OT-07-0063-RG-15/02/07///
    SQLX$ = SQLX$ + "Order by FechConta ASC, FechAcred ASC, FeHoReal ASC "
    '\\\OT-07-0063-RG-FIN///
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
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
      Do While Not .EOF
        U& = U& + 1
        Call TRACE(20, U&, FIN&)
        FECHAD = Int(CDbl(!FechConta))
          GoSub INCLUASDIAR
        FECHAD = Int(CDbl(!FechConta))
        NUMEFA$ = !CodiCom_Lar
        NCLIE% = !Nuprove
        ITOTA# = !ValAbsComp
        SIG% = Sgn(!IHABECOMP - !Idebecomp)
        MOBASI$ = TRIM$(!TipoMovim)
        '
        ZZ$ = REBLA$
        Call REPLA(ZZ$, MKI$(CVINT(!FechConta)), 3, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!FechAcred)), 5, 2)
        Call REPLA(ZZ$, MKS$(!NumeComp), 7, 4)
        Call REPLA(ZZ$, !CodiCom_Lar, 11, 18)
        Call REPLA(ZZ$, MKI$(NCLIE%), 31, 2)
        Call REPLA(ZZ$, RASOCLI$((-1) * NCLIE%), 33, 30)
        Call REPLA(ZZ$, CUITCLI$((-1) * NCLIE%), 63, 16)
        Call REPLA(ZZ$, POSIVAC$((-1) * NCLIE%), 79, 12)
        If PIVACLI%((-1) * NCLIE%) = 0 Then
             CUITT$ = TRIM$(SAFETEXT$(!CUIT))
             If CUITT$ = "" Then Call REPLA(ZZ$, CUITT$, 63, 16)
        End If
        '\\\OT-05-0264-WAR-09/05/05///
        If IsNull(!NROCAI) Then
           NCAI$ = ""
          Else
           NCAI$ = !NROCAI
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
          If UCase$(Mid$(ZZ$, 33, 6)) = "VARIOS" Then
            REFEMO$ = !Refecom
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
        PI% = PIVACLI%(NCLIE% * (-1))
        If PI% = 5 Then
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
        SQLX$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
        SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX$ = SQLX$ + "AND NuProve = " & NCLIE% & " "
        SQLX$ = SQLX$ + "AND TipoMovim <> '" & MOBASI$ & "' "
        '
        PROIVA% = 0: ULIVA% = 0: PROIBRU% = 0: ULIBRU% = 0
        With CABAFETMP
          .FindFirst SQLX$
25        If .NoMatch = False Then
            TIMO$ = UCase$(TRIM$(!TipoMovim))
            OPMO$ = !OpciMovim
            VAMO$ = !VariMovim
            IPAX# = Abs(!ValAbsComp)
            BASE# = 0
            On Error Resume Next
            BASE# = Abs(!BASEIMPO)
            On Error GoTo 0
            '
            If Abs(IPAX#) >= 0.005 Or BASE# >= 0.005 Then
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
                ElseIf TIMO$ = "IBRU" And UCase$(VAMO$) = "PERC" Then
'                  ZZ$ = RELIVA$(PROIBRU%)
'                  Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 193, 8)
'                  RESTOBRU# = RESTOBRU# - IPAX# * SIG%
'                  Call REPLA(ZZ$, MKI$(!CueContaI), 121, 2)
'                    JURIB$ = TRIM$(Mid$(OPMO$, 4))
'                  Call REPLA(ZZ$, JURIB$, 203, 1)
'                  Call REPLA(ZZ$, ECOARCH$("PROVINC", JURIB$), 205, 24)
'                  RELIVA$(PROIBRU%) = ZZ$
'                  ULIBRU% = PROIBRU%
'                  PROIBRU% = PROIBRU% + 1
                  ZZ$ = RELIVA$(0)
                  ACUANTEX# = CVD(Mid$(ZZ$, 193, 8))
                  Call REPLA(ZZ$, MKD$(ACUANTEX# + IPAX# * SIG%), 193, 8)
                  RESTOBRU# = RESTOBRU# - IPAX# * SIG%
                  Call REPLA(ZZ$, MKI$(!CueContaI), 121, 2)
                    JURIB$ = TRIM$(Mid$(OPMO$, 4))
                  Call REPLA(ZZ$, JURIB$, 203, 1)
                  Call REPLA(ZZ$, ECOARCH$("PROVINC", JURIB$), 205, 24)
                  If Abs(ACUANTEX#) >= 0.005 Then
                    Call REPLA(ZZ$, " ", 203, 1)
                    Call REPLA(ZZ$, "JURISD. VARIAS", 205, 24)
                  End If
                  RELIVA$(0) = ZZ$
                ElseIf TIMO$ = "IVA" Then
                  ZZ$ = RELIVA$(PROIVA%)
                  Call REPLA(ZZ$, VAMO$, 129, 8)
                  Call REPLA(ZZ$, MKD$(BASE# * SIG%), 137, 8)
                  Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 145, 8)
                  RESTOBRU# = RESTOBRU# - IPAX# * SIG% - BASE# * SIG%
                  If BASE# > 0.005 Then
                    Call REPLA(ZZ$, MKD$(100 * IPAX# / BASE#), 153, 8)
                  End If
                  RELIVA$(PROIVA%) = ZZ$
                  ULIVA% = PROIVA%
                  PROIVA% = PROIVA% + 1
              End If
            End If
            '
            .FindNext SQLX$
            GoTo 25
            '
          End If
        End With
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
          ZZ$ = RELIVA$(ULIVA%)
          Call REPLA(ZZ$, "No Grav.", 129, 8)
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
      .MoveNext
      Loop
    End With
    '
    FECHAD = CDbl(CDate("31/12/2050"))
    GoSub INCLUASDIAR
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
          SQLX$ = SQLX$ + " WHERE CDBL(Fechasien) >= " & DESDAT
          SQLX$ = SQLX$ + " AND CDBL(Fechasien) <= " & HASDAT
          SQLX$ = SQLX$ + " AND statpase >0  "
          SQLX$ = SQLX$ + " AND NuIdAsien = '" & NUMEFA$ & "' "
          SQLX$ = SQLX$ + "ORDER BY OrNuPase "
          Set DETACONTA = Contable.OpenRecordset(SQLX$, 4)
          '
          With DETACONTA
            On Error Resume Next
            .MoveFirst
            If Err < 1 Then
              On Error GoTo 0
              Do While Not .EOF
                icuexi% = !CUECOINT
                For IXX% = 1 To CACUEIVA%
                  If icuexi% = CUEIVA%(IXX%) Then
                    ICUEN#(IXX%) = ICUEN#(IXX%) + !IDebePase - !IHabePase
                    REPAS$ = !RefePase
                    GoTo 192
                  End If
                Next IXX%
                If NCLIE% < 1 Then
                  Call APERBASDAT("PROVE")
                  Proved12.FindFirst "CuMa_Cli = " & icuexi%
                  If Proved12.NoMatch = False Then
                    NCLIE% = Proved12!Nume_Cli
                  End If
                End If
192           .MoveNext
              Loop
            End If
          End With
          '
          ZZ$ = REBLA$
          Call REPLA(ZZ$, MKI$(FECHASI%(KUI%)), 3, 2)
          Call REPLA(ZZ$, MKI$(FECHASI(KUI%)), 5, 2)
          Call REPLA(ZZ$, MKS$(0), 7, 4)
          Call REPLA(ZZ$, NUMEFA$, 11, 18)
          Call REPLA(ZZ$, MKI$(NCLIE%), 31, 2)
          Call REPLA(ZZ$, RASOCLI$((-1) * NCLIE%), 33, 30)
          Call REPLA(ZZ$, CUITCLI$((-1) * NCLIE%), 63, 16)
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
              BASE# = 0
              If XVALO(TASIVANOM$(IXX%)) > 0.5 Then
                BASE# = 100 * ICUEN#(IXX%) / XVALO(TASIVANOM$(IXX%))
              End If
              Call REPLA(ZZ$, VAMO$, 129, 8)
              Call REPLA(ZZ$, MKD$(BASE#), 137, 8)
              Call REPLA(ZZ$, MKD$(ICUEN#(IXX%)), 145, 8)
              ITOTAX# = ITOTAX# + BASE# + ICUEN#(IXX%)
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
    '
End Sub

Private Sub notadebito_Click()
   Call Command4_Click
End Sub

Private Sub notascredito_Click()
    Call Command5_Click
End Sub

Private Sub pantalla_Click()
   Option1.Value = True
   Call Command7_Click
End Sub

Private Sub pantalla2_Click()
   Option1.Value = True
   Call Command2_Click

End Sub

Private Sub rankingdecomprasporarticuloyfecha_Click()
    Call Command22_Click
End Sub

Private Sub rankingporproveedorporrangodefechas_Click()
   Call Command17_Click
End Sub

'\\\OT-05-0559-WAR-FIN///
Private Sub reabrirmes_Click()
   Call Image1_Click
End Sub

Private Sub reimputaciondecomprobantesdecompras_Click()
   Call Command10_Click
End Sub

Private Sub resumenmensualivacompras_Click()
   Call Command13_Click
End Sub

Private Sub resumenmensualporcuenta_Click()
    Call Command6_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub subdiariodecompras_Click()
   Call Command16_Click
End Sub


Private Sub subiarioivacomprasporimpresora_Click()
    Call Command21_Click
End Sub


