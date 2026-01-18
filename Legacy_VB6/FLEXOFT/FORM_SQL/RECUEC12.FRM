VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form RECUEC07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes de Cuenta Corriente de Deudores"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C0FFE0&
      Height          =   375
      Left            =   100
      ScaleHeight     =   315
      ScaleWidth      =   9660
      TabIndex        =   66
      Top             =   7470
      Width           =   9720
      Begin VB.Line Line15 
         X1              =   0
         X2              =   9710
         Y1              =   305
         Y2              =   305
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7400
         TabIndex        =   70
         Top             =   10
         Width           =   1905
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   4225
         TabIndex        =   69
         Top             =   10
         Width           =   1870
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total General"
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
         TabIndex        =   68
         Top             =   20
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   7640
         TabIndex        =   67
         Top             =   20
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GRID2 
      Height          =   765
      Left            =   1920
      TabIndex        =   64
      Top             =   0
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
      Left            =   240
      TabIndex        =   65
      Top             =   0
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
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00006060&
      Height          =   7995
      Left            =   10000
      ScaleHeight     =   7935
      ScaleWidth      =   1755
      TabIndex        =   53
      ToolTipText     =   "Re-Aplicación de Comprobantes"
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   195
         Left            =   240
         TabIndex        =   71
         Top             =   6600
         Visible         =   0   'False
         Width           =   495
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "RECUEC12.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Impresión de Cuenta Corriente"
         Top             =   2415
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   60
         Top             =   6930
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -315
            TabIndex        =   61
            Top             =   0
            Width           =   12000
         End
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "RECUEC12.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Notas de Clientes"
         Top             =   3045
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "RECUEC12.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Re-Aplicacion de Creditos y Pagos a Cuenta"
         Top             =   4970
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
         Height          =   635
         Left            =   105
         Picture         =   "RECUEC12.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   955
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
         Picture         =   "RECUEC12.frx":0DC8
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Finaliza la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         Caption         =   "D-C"
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
         Picture         =   "RECUEC12.frx":0F12
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Listado Aplicación Diferencias de Cambio"
         Top             =   3885
         Width           =   640
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
         Height          =   640
         Left            =   120
         Picture         =   "RECUEC12.frx":121C
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Re-Aplicacion Manual de Créditos"
         Top             =   5760
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -280
         Picture         =   "RECUEC12.frx":1366
         Top             =   7370
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2730
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   195
      Left            =   8925
      TabIndex        =   47
      Top             =   105
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   12
      Left            =   12000
      TabIndex        =   46
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2750
      Index           =   10
      Left            =   9500
      TabIndex        =   45
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2750
      Index           =   8
      Left            =   16000
      TabIndex        =   44
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2750
      Index           =   7
      Left            =   2310
      TabIndex        =   43
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2750
      Index           =   9
      Left            =   3150
      TabIndex        =   42
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2750
      Index           =   11
      Left            =   4305
      TabIndex        =   41
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2750
      Index           =   13
      Left            =   6300
      TabIndex        =   40
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2750
      Index           =   15
      Left            =   7455
      TabIndex        =   39
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   0
      Left            =   1290
      TabIndex        =   38
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   1
      Left            =   3500
      TabIndex        =   37
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   2
      Left            =   4340
      TabIndex        =   36
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   3
      Left            =   5910
      TabIndex        =   35
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   4
      Left            =   7690
      TabIndex        =   34
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   5
      Left            =   9500
      TabIndex        =   33
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   6
      Left            =   12000
      TabIndex        =   32
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Crédito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   7350
      TabIndex        =   21
      Top             =   210
      Width           =   2520
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   735
         TabIndex        =   28
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1575
         Width           =   1590
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   735
         TabIndex        =   26
         Tag             =   "TBOX"
         Text            =   " "
         ToolTipText     =   "Doble Click para Listar Detalle"
         Top             =   1170
         Width           =   1590
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   735
         TabIndex        =   24
         Tag             =   "TBOX"
         Top             =   360
         Width           =   1590
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   735
         TabIndex        =   22
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   765
         Width           =   1590
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disp:"
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
         Index           =   1
         Left            =   60
         TabIndex        =   29
         Top             =   1635
         Width           =   645
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cheq:"
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
         Index           =   0
         Left            =   60
         TabIndex        =   27
         Top             =   1230
         Width           =   645
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L.Cre:"
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
         Index           =   2
         Left            =   60
         TabIndex        =   25
         Top             =   420
         Width           =   645
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.Cte:"
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
         Index           =   3
         Left            =   60
         TabIndex        =   23
         Top             =   825
         Width           =   645
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   9870
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   5460
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   2430
      Left            =   105
      ScaleHeight     =   2370
      ScaleWidth      =   9660
      TabIndex        =   14
      Top             =   2415
      Width           =   9725
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0F0FF&
         BorderStyle     =   0  'None
         Height          =   5580
         Left            =   0
         TabIndex        =   15
         Top             =   -150
         Width           =   10515
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00C0E0FF&
            Height          =   380
            Left            =   -120
            ScaleHeight     =   315
            ScaleWidth      =   9840
            TabIndex        =   50
            Top             =   120
            Width           =   9900
            Begin VB.Label Label8 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "  Fecha    Documento           Nro         Debe          Haber         Saldo"
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   0
               TabIndex        =   49
               Top             =   0
               Width           =   10410
            End
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2085
            IntegralHeight  =   0   'False
            Left            =   0
            TabIndex        =   16
            Top             =   480
            Width           =   9720
         End
         Begin VB.Line Line7 
            X1              =   1155
            X2              =   1155
            Y1              =   1470
            Y2              =   3255
         End
      End
      Begin VB.Line Line22 
         X1              =   0
         X2              =   6405
         Y1              =   2415
         Y2              =   2415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Datos de la Cuenta Activa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   105
      TabIndex        =   2
      Top             =   210
      Width           =   7100
      Begin VB.TextBox Text9 
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
         Left            =   1365
         TabIndex        =   11
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1560
         Width           =   3690
      End
      Begin VB.TextBox Text1 
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
         Left            =   210
         TabIndex        =   0
         Text            =   " "
         Top             =   800
         Width           =   855
      End
      Begin VB.CommandButton Command2 
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
         Left            =   1050
         TabIndex        =   1
         Top             =   800
         Width           =   180
      End
      Begin VB.TextBox Text2 
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
         Left            =   1365
         TabIndex        =   8
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   420
         Width           =   3690
      End
      Begin VB.TextBox Text3 
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
         Left            =   1365
         TabIndex        =   7
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   800
         Width           =   3690
      End
      Begin VB.TextBox Text4 
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
         Left            =   1365
         TabIndex        =   6
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1180
         Width           =   750
      End
      Begin VB.TextBox Text5 
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
         Left            =   2205
         TabIndex        =   5
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1155
         Width           =   2850
      End
      Begin VB.TextBox Text8 
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
         Left            =   5250
         TabIndex        =   4
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text7 
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
         Left            =   5250
         TabIndex        =   3
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   800
         Width           =   1695
      End
      Begin VB.Label Label26 
         BackColor       =   &H00EEEEEE&
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
         Height          =   300
         Left            =   210
         TabIndex        =   30
         Top             =   1155
         Width           =   1035
      End
      Begin VB.Label label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax-Nbr:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   5250
         TabIndex        =   13
         Top             =   1260
         Width           =   960
      End
      Begin VB.Label label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto:"
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
         Index           =   0
         Left            =   105
         TabIndex        =   12
         Top             =   1575
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         TabIndex        =   10
         ToolTipText     =   "Doble Click para Refrescar"
         Top             =   420
         Width           =   960
      End
      Begin VB.Label label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   5250
         TabIndex        =   9
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   2445
      Left            =   105
      ScaleHeight     =   2385
      ScaleWidth      =   9660
      TabIndex        =   18
      Top             =   5040
      Width           =   9725
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00C0FFE0&
         Height          =   360
         Left            =   -30
         ScaleHeight     =   300
         ScaleWidth      =   9840
         TabIndex        =   51
         Top             =   -30
         Width           =   9900
         Begin VB.Label Label9 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Documento           Nro   Emision         Importe   Vencim.        Saldo($)     Sdo (U$)    "
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            TabIndex        =   52
            Top             =   0
            Width           =   15000
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFE0&
         BorderStyle     =   0  'None
         Height          =   5630
         Left            =   0
         TabIndex        =   19
         Top             =   315
         Width           =   10515
         Begin VB.ListBox List2 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            IntegralHeight  =   0   'False
            Left            =   -30
            TabIndex        =   20
            Top             =   0
            Width           =   9725
         End
         Begin VB.Line Line14 
            X1              =   0
            X2              =   9660
            Y1              =   2100
            Y2              =   2100
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "RECUEC12.frx":3288
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "RECUEC12.frx":34BC
      TabIndex        =   63
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label MONEMIS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pesos"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005500&
      Height          =   330
      Left            =   7500
      TabIndex        =   31
      ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
      Top             =   150
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Line Line23 
      X1              =   105
      X2              =   9825
      Y1              =   4840
      Y2              =   4840
   End
   Begin VB.Line Line21 
      X1              =   9825
      X2              =   9825
      Y1              =   5040
      Y2              =   7820
   End
   Begin VB.Line Line20 
      X1              =   9825
      X2              =   9825
      Y1              =   2415
      Y2              =   4830
   End
   Begin VB.Line Line13 
      X1              =   4830
      X2              =   6090
      Y1              =   3570
      Y2              =   3990
   End
End
Attribute VB_Name = "RECUEC07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODOMONEI%
Dim MONRESU%
'jandy sql
Dim ctrl As Object
Sub ClearControls()
    For Each ctrl In Me.Controls
        If ctrl.Tag = "TBOX" Then ctrl.TEXT = ""
    Next
    Label26 = ""
End Sub



Private Sub recuperaapli()
    Call ABRECONEXION
    'NO LO MODIFICO  MAS POR QUE NO SE USA ,
    
    FF% = FECHANUM("01/11/10")
    FECHA$ = FETEXCOR1$(FF%)
    For II% = 1 To 32767
        JJ& = II%
        Call TRACE(24, JJ&, 32767)
        NCLI$ = CStr(II%)
        SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE NuClien = " & NCLI$ & " AND IHABECOMP >0 AND TIPOMOVIM='FVEN' AND FECHCONTA>'" & FECHA$ & "'"
        Set RST = READSET(SQLX$)
        With RST
            Do While Not .EOF
                DOCRED$ = SAFETEXT$(!codicom_lar)
                DOCREDCOR$ = SAFETEXT$(!CODICOM_COR)
                If CantRegistros("APLICOBRA", "CoDocAp_Lar='" & DOCRED$ & "'") > 0 Then GoTo 20
                '
                IMPO = XVALO(!IHABECOMP)
                'FF% = CVINT(!FECHCONTA)
                CONDI$ = "NUCLIEN=" & NCLI$ & " AND TIPOMOVIM='FVEN' AND OPCIMOVIM='FC' AND ABS((IDEBECOMP-" & TRIM$(FORMATNUM(IMPO, "F12.2")) & "))<0.005"
                DOCDEB$ = SAFETEXT$(VALOBADA("CAJABANC", "CODICOM_LAR", CONDI$))
                DOCDEBCOR$ = SAFETEXT$(VALOBADA("CAJABANC", "CODICOM_COR", CONDI$))
                'Call COMUNI(CONDI$)
                'Call COMUNI(DOCDEB$)
                If TRIM$(DOCDEB$) <> "" Then
                    'OPX% = COMALTER%("Nota Credito: '" & DOCRED$ & "'   Factura: '" & DOCDEB$ & "'\\No, Cancelar\Si, Continuar")
                    'If OPX% = 1 Then GoTo 20
                    '
                    Dim IMPUMANU As ADODB.Recordset
                    Set IMPUMANU = New ADODB.Recordset
25                  On Error Resume Next
                    IMPUMANU.Open FILTSQL$("APLICOBRA"), FLEXCONN, adOpenKeyset, adLockPessimistic
                    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                      On Error GoTo 0
                      Call CapturaErrorOpen
                      GoTo 25
                    End If
                    On Error GoTo 0
                    '
                    CI% = Int(NCLI$)
                    RAZON$ = RASOCLI(CI%)
                    NUCLI$ = CODICLI$(CI%)
                    '
                    With IMPUMANU
                    .AddNew   ' YA CONVERTIDO
                    On Error Resume Next
                    !CodiEmpr = CodiEmp%
                    !CodiEmpr = 0
                    !CodiSucu = 0
                    !NUCLIEN = CI%
                    !CODICLIE = CODICLI$(CI%)
                    !RasoClie = Mid(RAZON$, 1, 24)
                    !codocap_lar = DOCRED$
                    !CODOCAP_COR = DOCREDCOR$
                    !IMPOAPLI = IMPO
                    !refeapli = "Reaplicación de Comprobantes"
                    !codocor_lar = DOCDEB$
                    !CODOCOR_COR = DOCDEBCOR$
                    !ItotOrig = 0
                    !IIVAORIG = 0
                    .Update
                    On Error GoTo 0
                    '
                    End With
                Else
                    Call COMUNI("Fallo Nota Credito: '" & DOCRED$ & "'   Factura: '" & DOCDEB$ & "'")
                End If
20              .MoveNext
            Loop
        End With
    Next II%
End Sub

'Sub MUESTRASTSCREEN(NC%, NUDOCUM$, FE%, OK%)
'    '
'    ' OJO  SE GENERA ESTA RUTINA PARA MOSTRAR LAS FACTURAS CON EL STDSCREEN
'    ' SE LLAMAN A LOS SQR COMO ESTAN EN MEDITEA SE VA A TENER QUE VER SI
'    'SE VAN A CONFIGURAR PARA QUE NO QUEDE PERSONALIZADO
'
'    'VALIDAR SI ES EXISTE
''    If EXISTE%("VENGES09.EXE") < 1 Then
''       Exit Sub
''    End If
'    '
'    FEX = FETEXCOR1$(FE%)
'    CONDI$ = "TIPOMOVIM = 'FVEN' AND fechemisi = '" & FEX & "' AND NUCLIEN = " & NC% & "  AND  codicom_lar like '%" & NUDOCUM$ & "%'"
'    FACTURAX$ = VALOBADA("CAJABANC", "CODICOM_LAR", CONDI$, "NOMESS")
'    '
'    If TRIM$(FACTURAX$) <> "" Then
'        TIPOCOM$ = Mid$(FACTURAX$, 1, 2)
'        LETR$ = Right(FACTURAX$, 1)
'        '
'        CONDI$ = "DOSECLAR='" & FACTURAX$ & "' AND NUME_CLIE= " & NC% & " AND COD_INTE > 0"
'        CANTREG& = XVALO(CantRegistros("MOVISTO", CONDI$, "NOMESS"))
'        EMP$ = "1-MDT": If CANTREG& = 0 Then EMP$ = "T-MDT"
'        ARCSQR$ = TIPOCOM$ & LETR$ & EMP$
'         '
'        If EXISTE%(LUDAT$ + ARCSQR$ + ".SQR") = 0 Then
'            Call COMUNI("Falta Archivo " & ARCSQR$ + ".SQR")
'            Exit Sub
'        End If
'       '
'       OK% = 1
'       FILTX$ = TRIM$(FACTURAX$) & ";" & TRIM$(Str$(NC%))
'       Call STDFORM(ARCSQR$, FILTX$, MODO$)
'    End If
'
'End Sub
'
Private Sub Command6_Click()
   Command6.Enabled = False
   Call OPVARFA07.Command2_Click
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   NC% = XVALO(Text1)
   If NC% < 1 Then Exit Sub
   If REGICLI&(NC%) < 1 Then Exit Sub
   '
   Dim DOCABX$(2048), SALPES#(2048), SALDOL#(2048), REDOCA&(2048)
   '
   Screen.MousePointer = 11
   CADOCPE% = 0
   RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
   For KUL& = 1 To Len(RFG$) Step 4
     RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
     XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
     DOCABY$ = Mid$(XXC$, 13, 18)
     SALPESY# = CVD(Mid$(XXC$, 65, 8)) - CVD(Mid$(XXC$, 73, 8))
     SALDOLY# = CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8))
     '
     For KKI% = 1 To CADOCPE%
       If DOCABX$(KKI%) = DOCABY$ Then GoTo 20
     Next KKI%
     CADOCPE% = CADOCPE% + 1
     DOCABX$(CADOCPE%) = DOCABY$
     REDOCA&(CADOCPE%) = RECABIN&
     KKI% = CADOCPE%
     '
20   SALPES#(KKI%) = SALPES#(KKI%) + SALPESY#
     SALDOL#(KKI%) = SALDOL#(KKI%) + SALDOLY#
     '
   Next KUL&
   '
   Call BLANARCH("SELDOCAP")
   For KKI% = 1 To CADOCPE%
     'If SALDOL#(KKI%) <= 0.005 Then
     '  If SALPES#(KKI%) > 0 Then
         XXC$ = REGLEIDO$("CARBIMON", REDOCA&(KKI%))
         FEMI = CVI(Mid$(XXC$, 1, 2))
         Z$ = Space$(60) + MKS$(0)
         Call REPLA(Z$, DOCABX$(KKI%), 1, 18)
         Call REPLA(Z$, FECHATEX$(FEMI), 19, 8)
         Call REPLA(Z$, "$" + FORMATNUM$(SALPES#(KKI%), "F11.2"), 29, 12)
         Call REPLA(Z$, "u$s" + FORMATNUM$(SALDOL#(KKI%), "F11.2"), 44, 14)
         Call REPLA(Z$, MKS$(REDOCA&(KKI%)), 61, 4)
         Call REGAPP("SELDOCAP", Z$)
     '  End If
     'End If
   Next KKI%
   Call CIERRARCH("SELDOCAP")
   '
   Screen.MousePointer = 1
   If ULTREG&("SELDOCAP") < 1 Then
     Call COMUNI("No Aparecen Créditos Pendientes\para la Cuenta Elegida.")
     Exit Sub
   End If
   '
   Call SELECHO("SELDOCAP")
   DOCABY$ = Left$(VALACT1$("SELDOCAP"), 18)
   If TRIM$(DOCABY$) = "" Then Exit Sub
   '
End Sub

Private Sub Command8_Click()
10  SELECHO ("DEUDOR1")
    NCLI$ = TRIM$(VALACT1("DEUDOR1"))       'Muestra Clientes
    If NCLI$ <> "" Then
       Call COPYSTRU("SELDOCAP", "SELDOCA1")
       Call CARGALISEL("SELDOCA1", "CAJABANC", "CodiCom_Lar/A24;FechEmisi/D8;IHABECOMP/F12.2", "NuClien = " & NCLI$ & " AND IHABECOMP >0 AND (TIPOMOVIM='FVEN' OR TIPOMOVIM='RCOB')")
15     Call SELECHO("SELDOCA1")             'Muestra Docu de Credito
       DOCRED$ = TRIM$(VALACT1("SELDOCA1"))
       If DOCRED$ <> "" Then
          Call COPYSTRU("SELDOCAP", "SELDOCA2")
          Call CARGALISEL("SELDOCA2", "CAJABANC", "CodiCom_Lar/A24;FechEmisi/D8;IDEBECOMP/F12.2", "NuClien = " & NCLI$ & " AND IDEBECOMP >0 AND TIPOMOVIM='FVEN'")
          Call SELECHO("SELDOCA2")          ' Muestra Docu de Debito
          DOCDEB$ = TRIM$(VALACT1("SELDOCA2"))
          If DOCDEB$ <> "" Then
             XX1$ = Space$(53)
             Call REPLA(XX1$, DOCRED$, 1, 24)
             Call REPLA(XX1$, DOCDEB$, 25, 24)
20           RESP$ = OBJPLA$("INPUMANU", XX1$)      'Usuario ingresa imputacion manual
             INPUMANUAL# = CVS(Mid(RESP$, 49, 4))
             If INPUMANUAL# > 0.0005 Then
                'Se obtienen los datos para grabar registro en APLICOBRA
                CONDI$ = "CodiCom_Lar='" & DOCRED$ & "' AND NuClien=" & NCLI$ & " AND IHABECOMP >0 AND (TIPOMOVIM='FVEN' OR TIPOMOVIM='RCOB')"
                IMPOCRE# = VALOBADA("CAJABANC", "IHABECOMP", CONDI$)
                DOCREDCOR$ = VALOBADA("CAJABANC", "CODICOM_COR", CONDI$)
                '
                CONDI$ = "CodiCom_Lar='" & DOCDEB$ & "' AND NuClien=" & NCLI$ & " AND IDEBECOMP >0 AND TIPOMOVIM='FVEN'"
                ImpoDeb# = VALOBADA("CAJABANC", "IDEBECOMP", CONDI$)
                DOCDEBCOR$ = VALOBADA("CAJABANC", "CODICOM_COR", CONDI$)
                TIPOCAMBIO# = VALOBADA("CAJABANC", "TIPO_CAM", CONDI$)
                
                'Se extrae el iva del documento de debito
                CONDI$ = "CodiCom_Lar='" & DOCDEB$ & "' AND NuClien=" & NCLI$ & " AND TIPOMOVIM='IVA'"
                IMPIVADEBE$ = VALOBADA("CAJABANC", "IHABECOMP", CONDI$)
                
                '
                If (INPUMANUAL# >= IMPOCRE#) Or (INPUMANUAL# >= ImpoDeb#) Then
                   Call COMUNI("Importe Ingresado Supera al Importe de los Comprobantes.")
                   GoTo 20
                 End If
                 '
                 Call BORRAREGISTROS("APLICOBRA", "CODOCAP_LAR='" & DOCRED$ & "'", REG&)
                 Call BORRAREGISTROS("APLICOBRA", "CODOCOR_LAR='" & DOCDEB$ & "'", REG&)
                 '
                 Dim IMPUMANU As ADODB.Recordset
                 Set IMPUMANU = New ADODB.Recordset
25               On Error Resume Next
                 IMPUMANU.Open FILTSQL$("APLICOBRA"), FLEXCONN, adOpenKeyset, adLockPessimistic
                 If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                   On Error GoTo 0
                   Call CapturaErrorOpen
                   GoTo 25
                 End If
                 On Error GoTo 0
                 '
                 CI% = XVALO(NCLI$)
                 RAZON$ = RASOCLI(CI%)
                 NUCLI$ = CODICLI$(CI%)
                 '
                 With IMPUMANU
                 .AddNew   ' YA CONVERTIDO
                 On Error Resume Next
                 !CodiEmpr = CodiEmp%
                 On Error GoTo 0
                 !CodiEmpr = 0
                 !CodiSucu = 0
                 !NUCLIEN = NCLI$
                 !CODICLIE = NUCLI$
                 !RasoClie = RAZON$
                 !codocap_lar = DOCRED$
                 !CODOCAP_COR = DOCREDCOR$
                 !IMPOAPLI = INPUMANUAL#
                 !refeapli = "Reaplicación de Comprobantes"
                 !codocor_lar = DOCDEB$
                 !CODOCOR_COR = DOCDEBCOR$
                 !ItotOrig = ImpoDeb#
                 !IIVAORIG = IMPIVADEBE$
                 .Update
                 '
                 End With
                 Text1 = NCLI$
                 Call Command4_Click
             End If
           Else
             GoTo 15
           End If
        Else
          GoTo 10
       End If
    End If
       
End Sub

Private Sub Command9_Click()
Call recuperaapli
End Sub

Private Sub GRID1_DblClick()
   NGRID = GRID1.Row
   NLIST = List1.ListCount
   If (NLIST) >= NGRID - 1 Then
     List1.ListIndex = NGRID - 1
     List1_DblClick
   End If
End Sub

Private Sub GRID1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
      PVER = 0: PVER2 = 0: KKF% = GRID1.TopRow
      For KKII% = 0 To KKF% - 1
         PVER = PVER + GRID1.RowHeight(KKII%) + GRID1.GridLineWidth + 7
      Next KKII%
      KKII% = KKF% - 1
      While PVER2 < y - GRID1.RowHeight(KKII%)
         KKII% = KKII% + 1
         PVER = PVER + GRID1.RowHeight(KKII%) + GRID1.GridLineWidth + 7
         PVER2 = PVER2 + GRID1.RowHeight(KKII%) + GRID1.GridLineWidth + 7
      Wend
      NGRID = KKII%
      GRID1.Row = NGRID
      NLIST = List1.ListCount
      If (NLIST) >= NGRID - 1 Then
         List1.ListIndex = NGRID - 1
         Call List1_MouseDown(Button, Shift, x, y)
      End If
   End If
End Sub

Private Sub Label1_DblClick()
   Call Text1_Change
End Sub

Private Sub Label5_Change()
   Text10 = TRIM$(Label5)
End Sub

Private Sub MONEMIS_DblClick()
   '
   Static CAMON%, DMONE$(), TCMONE()
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
   If Abs(XVALO(Text1)) < 32767 Then
     NCX% = XVALO(Text1)
     If NCX% > 0 Then
       If RASOCLI$(NCX%) <> "" Then
         GoTo 20
       End If
     End If
   End If
   Exit Sub
   '
20 MONEMIAR% = MONRESU%
   MONRESU% = MONRESU% + 1
   If MONRESU% > CAMON% Then MONRESU% = 1
   MONEMIS = DMONE$(MONRESU%)
   MONEMIS.ForeColor = RGB(0, 0, 255)
   If MONRESU% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
   End If
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONRESU%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   Call CARESCUE
   '
End Sub

Private Sub Text10_GotFocus()
    Command1.SetFocus
End Sub
Private Sub Text11_GotFocus()
    Command1.SetFocus
End Sub

Private Sub Text12_GotFocus()
    Command1.SetFocus
End Sub

Private Sub Text2_gotfocus()
'\\\OT-06-0554-RG-13/12/06///
    Text1.SetFocus
'\\\OT-06-0554-RG-FIN///
End Sub


Private Sub Text3_GotFocus()
'\\\OT-06-0554-RG-13/12/06///
    Text1.SetFocus
'\\\OT-06-0554-RG-FIN///
End Sub

Private Sub Text4_GotFocus()
'\\\OT-06-0554-RG-13/12/06///
    Text1.SetFocus
'\\\OT-06-0554-RG-FIN///
End Sub

Private Sub Text5_GotFocus()
'\\\OT-06-0554-RG-13/12/06///
    Text1.SetFocus
'\\\OT-06-0554-RG-FIN///
End Sub

Private Sub Text6_GotFocus()
    Command1.SetFocus
End Sub


Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Command2_Click()
    '
    Call SELECHO("DEUDOR1")
    VSELE$ = TRIM$(VALACT1$("DEUDOR1"))
    If VSELE$ <> "" Then
       NC% = XVALO(VSELE$)
       Text1.TEXT = TRIM$(Str$(NC%))
       'Call CARESCUE
    End If
    '
End Sub

Sub Command3_Click()
    '
    Screen.MousePointer = 11
    '
    If MONRESU% > 1 Then
      Call MENSERR(24, "Proceso Imposible - Cambie Moneda a '" + TRIM$(DEMONECO$(1)) + "'")
      Screen.MousePointer = 1
      Exit Sub
    End If
    '
    If ULTREG&("DEUDOR1") > 0 Then
        x$ = REGLEIDO$("DEUDOR1", 1)      ' fuerza la
        Call GRAREG("DEUDOR1", x$, 1)     ' reindexacion
        Call CIERRARCH("DEUDOR1")         ' del archivo
    End If
    '
    Screen.MousePointer = 11
    '
    If Int(XVALO(Text1.TEXT)) > 0 Then
        If Int(XVALO(Text1.TEXT)) <= 32767 Then
            NCMIN% = Int(XVALO(Text1.TEXT))
            NCMAX% = NCMIN%
            GoTo 10
        End If
    End If
    '
    NCMIN% = 32767: NCMAX% = 0
    FIN& = ULTREG&("DEUDOR1")
    For U& = 1 To FIN&
        NCX% = CVI(Left$(REGLEIDO$("DEUDOR1", U&), 2))
        If NCX% > 0 Then
            If NCX% > NCMAX% Then NCMAX% = NCX%
            If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
    Next U&
    '
10  Screen.MousePointer = 1
    With RESCUC07
      .Text10 = TRIM$(Str$(NCMIN%))
      .Text9 = RASOCLI$(NCMIN%)
      .Text8 = TRIM$(Str$(NCMAX%))
      .Text7 = RASOCLI$(NCMAX%)
      '
      HOO% = HOY(HO$)
      .Text6 = "01" + Mid$(HO$, 3)
      .Text5 = HO$
    End With
    '
    RESCUC07.Show 1
    '
End Sub

Sub Command4_Click()
    '
    If MONRESU% > 1 Then
      Call MENSERR(24, "Proceso Imposible - Cambie Moneda a '" + TRIM$(DEMONECO$(1)) + "'")
      Exit Sub
    End If
    '
    MONEMIS = DEMONECO$(MONRESU%)
    MONEMIS.ForeColor = RGB(0, 0, 255)
    If MONRESU% > 1 Then
       MONEMIS.ForeColor = RGB(255, 0, 0)
    End If
    '
    NC% = XVALO(Text1.TEXT)
    If NC% < 1 Or RASOCLI$(NC%) = "" Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Text1.SetFocus
       Exit Sub
    End If
    '
    If CONTROLGRABA% = 1 Then GoTo 80
    If COMALTER%("El Proceso Invocado Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Créditos y Cobranzas\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar") <> 2 Then
       Exit Sub
    End If
    '
80  Call REAPLICRE(NC%)
    Text1.TEXT = TRIM$(Str$(NC%))
    List1.Clear
    List2.Clear
    Call REMPLAGRID(Me, GRID1, Picture1, Label8, List1)
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
    Call CARESCUE
    DoEvents
    '
End Sub

Private Sub Command5_Click()
    NC% = XVALO(TRIM$(Text1.TEXT))
    If NC% > 0 And RASOCLI$(NC%) <> "" Then
        Command5.Enabled = False
        Call NOTACLI(NC%, 2)
        Command5.Enabled = True
      Else
        Call MENSERR(24, "Falta Ingresar Número de Cliente")
    End If
End Sub

Private Sub Form_Activate()
    Refresh
    If XVALO(Text1.TEXT) > 0 Then
        NC% = XVALO(Text1.TEXT)
      ElseIf NC% > 0 Then
        Text1.TEXT = NC%
    End If
    'List1.Clear
    'List2.Clear
    'If NC% > 0 Then
    '  If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
    '    Call CARESCUE
    '  End If
    'End If
End Sub

Private Sub Form_Load()
    '

'    EPAC$ = TRIM$(EMPREAC$)
'    If EPAC$ <> "" Then
'        Caption = Caption + "  -  " + EPAC$
'    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    
    MONRESU% = 1
    MONEMIS = DEMONECO$(MONRESU%)
    MONEMIS.ForeColor = RGB(0, 0, 255)
    If MONRESU% > 1 Then
       MONEMIS.ForeColor = RGB(255, 0, 0)
    End If
    MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONRESU%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
       '
       Text11.FontName = "Lucida Console"
       Text11.FontSize = 9
       Text11.FontBold = True
       Text11.Height = 270
       Text11.Top = 300
       '
       Text10.FontName = "Lucida Console"
       Text10.FontSize = 9
       Text10.FontBold = True
       Text10.Height = 270
       Text10.Top = 620
       '
       Text6.FontName = "Lucida Console"
       Text6.FontSize = 9
       Text6.FontBold = True
       Text6.Height = 270
       Text6.Top = 940
       '
       Text12.FontName = "Lucida Console"
       Text12.FontSize = 9
       Text12.FontBold = True
       Text12.Height = 270
       Text12.Top = 1260
       '
       label2(2).Top = 400
       label2(3).Top = 720
       label(0).Top = 1040
       label(1).Top = 1360
       '
       Frame6.Height = 1670
       Frame6.Top = 600
       '
       MONEMIS.Visible = True
       '
       Label8.Caption = Label8.Caption + "  ME   Cotiz."
       Label8.FontName = "Lucida Console"
       Label8.FontSize = 9
       Label8.FontBold = False
       Label8.Top = 90
       List1.FontName = "Lucida Console"
       List1.FontSize = 9
       List1.FontBold = False
       '
       Label9.FontName = "Lucida Console"
       Label9.FontSize = 9
       Label9.FontBold = False
       Label9.Top = 90
       List2.FontName = "Lucida Console"
       List2.FontSize = 9
       List2.FontBold = False
       '
       Frame7(0).Left = 1140
       Frame7(1).Left = 3030
       Frame7(2).Left = 3785
       Frame7(3).Left = 5230
       Frame7(4).Left = 6740
       Frame7(5).Left = 8200
       Frame7(6).Left = 8620
       Frame7(7).Left = 2100
       Frame7(9).Left = 2750
       Frame7(11).Left = 3800
       Frame7(13).Left = 5500
       Frame7(15).Left = 6550
       Frame7(8).Left = 8220
       Frame7(10).Left = 9540
       Frame7(12).Left = 9540
       '
       Label6.Left = 3560
       Label5.Left = 6290
       Label4.Visible = True
   End If
   Call APLICACIONSKINS(Me, Picture3)
   Call REMPLAGRID(Me, GRID1, Picture1, Label8, List1)
   Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    On Error Resume Next
    Hide
    On Error GoTo 0
End Sub

Private Sub Label26_DblClick()
   Call SELECHO("INDICLIE")
   CPROX$ = TRIM$(VALACT1$("INDICLIE"))
   If CPROX$ <> "" Then
     NCXI% = NUMECLI%(CPROX$)
     If NCXI% > 0 Then
       If REGICLI&(NCXI%) > 0 Then
         Label26 = CPROX$
         NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NC%))
       End If
     End If
   End If
End Sub

Private Sub List1_DblClick()
   '
   OPX% = 0: TIPOCOM% = 0
   RELIS$ = List1.TEXT
   FE% = FECHANUM(Mid$(RELIS$, 1, 8))
   TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 16))
   If TIDOCUM$ = "Recibo Cobranza" Then TIDOCUM$ = "Recibo"
   If TIDOCUM$ = "Ajuste de Saldo" Then TIDOCUM$ = "Ajuste Saldo Cliente"
   If TIDOCUM$ = "Comp.Ctas." Then TIDOCUM$ = "Compensación de Cuentas"
   NUDOCU$ = " " + TRIM$(Mid$(RELIS$, 27, 8)) + " "
   NUDOCU1$ = Str$(XVALO(NUDOCU$))
   If InStr(UCase$(TIDOCUM$), "FACTURA") > 0 Or InStr(UCase$(TIDOCUM$), "NOTA DE") > 0 Then
     NUDOCU$ = TRIM$(NUDOCU$)
     While Len(NUDOCU$) < 8: NUDOCU$ = "0" + NUDOCU$: Wend
     If TRIM$(Mid$(RELIS$, 257, 18)) <> "" Then
        NUDOCU$ = TRIM$(Mid$(RELIS$, 257, 18))
        NUDOCU2$ = Mid$(NUDOCU$, 4, 13)
     End If
     TIPOCOM% = 1
   ElseIf InStr(UCase$(TIDOCUM$), "AJUSTE SALDO") > 0 Or InStr(UCase$(TIDOCUM$), "COMPENSACIÓN DE CUENTAS") > 0 Then
     If TRIM$(Mid$(RELIS$, 257, 18)) <> "" Then
        NUDOCU$ = TRIM$(Mid$(RELIS$, 257, 18))
        NUDOCU1$ = ""
        NUDOCU2$ = NUDOCU$
     End If
     TIPOCOM% = 2
   ElseIf InStr(UCase$(TIDOCUM$), "RECIBO") > 0 Then
     NUDOCU1$ = NUDOCU$
     NUDOCU2$ = NUDOCU$
     TIPOCOM% = 3
   End If
   OPX% = COMALTER%("Opciones de Trabajo:\\Documento Digital\Re-Emitir Comprobante")
   If OPX% = 2 Then
      Call REPRIFAC(NUDOCU$)
      Exit Sub
   End If
   '
Screen.MousePointer = 11
EPAC$ = TRIM$(UCase$(EMPREAC$))
If InStr(EPAC$, "FERVI") > 0 Then
   If OPX% > 0 And TIPOCOM% = 1 Then
      If FE% < FECHANUM("14/01/13") Then
         NUDOQUI& = FACTUFERVI&(TIDOCUM$, TRIM$(NUDOCU1$), FE%)
         If NUDOQUI& > 0 Then Call MUESTRADOC(NUDOQUI&, "", TIDOCUM$, NUDOCU1$, NUDOCU2$): Exit Sub
      End If
   End If
   
End If

   DOCUNU& = DOCID&(TIDOCUM$, NUDOCU1$, NUDOCU2$, FECHATEX$(FE%))
   If DOCUNU& < 1 Then
        If OPX% > 0 And TIPOCOM% = 1 Then ' SI OPX% ES MAYOR A 0 POR QUE ES FACTURA O NOTA DE ..
            Call MUESTRASTSCREEN(XVALO(Text1.TEXT), NUDOCU$)
        Else
            Call COMUNI("Documento No Encontrado")
        End If
   Else
        OPPXX$ = ""
        If NUDOCU$ = "00000032" Then OPPXX$ = "RECUP"
        Call MUESTRADOC(DOCUNU&, OPPXX$, TIDOCUM$, NUDOCU1$, NUDOCU2$)
   End If
   Screen.MousePointer = 1

'   RACLI$ = TRIM$(Left$(RASOCLI$(XVALO(Text1.TEXT)), 8))
'   Screen.MousePointer = 11
'   VUELTA% = 0
'10 For UI& = ULTREG&("PDOCLST") To 1 Step -1
'     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 66)
'     FEMIX% = CVI(Mid$(XX$, 65, 2))
'     If FEMIX% > 8224 Then
'        If FEMIX% < FE% - 32 Then
'          GoTo 12
'        End If
'     End If
'     If InStr(XX$, NUDOCU$) > 4 Then
'       If InStr(XX$, TIDOCUM$) > 4 Then
'         'SE OMITE LA COMPARACION POR RAZON SOCIAL, ASUMIENDO QUE EXISTE
'         'LA POSIBILIDAD DE QUE ESTA CAMBIE
'         If InStr(XX$, RACLI$) > 16 Or VUELTA% > 0 Then
'           DOCUNU& = CVS(Left$(XX$, 4))
'           GoTo 5
'         End If
'       End If
'     End If
'   Next UI&
'12 VUELTA% = VUELTA% + 1
'   If VUELTA% < 2 Then GoTo 10
'   Screen.MousePointer = 1
'   '
'   If OPX% > 0 Then ' SI OPX% ES MAYOR A 0 POR QUE ES FACTURA O NOTA DE ..
'      Call MUESTRASTSCREEN(XVALO(Text1.TEXT), NUDOCU$, FE%, OK%)
'   End If
'
'   If OK% < 1 Then
'   Call COMUNI("Imposible Abrir Documento Desde esta Aplicación")
'   End If
'   Exit Sub
'   '
'5  OPPXX$ = ""
'   If NUDOCU$ = "00000032" Then OPPXX$ = "RECUP"
'   Call MUESTRADOC(DOCUNU&, OPPXX$)
'   '
End Sub
'
Sub REPRIFAC(NUDOCU$)
   Call COMUNI("No Disponible por el Momento")
End Sub
'
Private Sub Text1_Change()
   List1.Clear
   List2.Clear
   'jandy SQL
   Call ClearControls
   Call REMPLAGRID(Me, GRID1, Picture1, Label8, List1)
   Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
   '
   If Abs(XVALO(Text1)) <= 999999 Then
     NCX% = XVALO(Text1)
     If NCX% > 0 Then
       If RASOCLI$(NCX%) <> "" Then
         NC% = NCX%
         Call CARESCUE
         '23/03/07 (OT-07-0121)
         LMCR# = LIMICRE(NC%)  'MUESTRA QUE LA CUENTA DEL CLIENTE ESTA SUSPENDIDA
         If LMCR# < 0 Then
            Text11 = "Suspendida"
         End If
        '23/03/07 (FIN)
       End If
     End If
   End If
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("DEUDOR1")
    VSELE$ = TRIM$(VALACT1$("DEUDOR1"))
    If VSELE$ <> "" Then
       NC% = XVALO(VSELE$)
       Text1.TEXT = TRIM$(Str$(NC%))
       'Call CARESCUE
    End If
    '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       'If XVALO(Text1.TEXT) > 0 Then
       '   If XVALO(Text1.TEXT) <= 32767 Then
       '      NC% = XVALO(Text1.TEXT)
       '      If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
       '         Call CARESCUE
       '         Exit Sub
       '      End If
       '   End If
       'End If
       '
       'Beep
       'Text1.TEXT = ""
       '
    End If
    '
End Sub

Sub Text6_DBLCLICK()
   Screen.MousePointer = 11
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 32767 Then
        NC% = XVALO(Text1.TEXT)
        'SELCHE04.LICHESEL.Style = 0
        SELCHE07.Caption = "Cheques Pendientes de Acreditaciòn: " + RASOCLI$(NC%)
        SELCHE07.List1.Clear
        SELCHE07.LICHESEL.Clear
        '
        FEXI% = DIANTE(DIANTE(HOY(HOS$)))
        'AYER = CDbl(CVDAT(FEXI%))
        AYER = FETEXCOR1$(FEXI%)
                       
15      Call APERBASDAT("CABAN")
        Dim CABATEMP As ADODB.Recordset
        Set CABATEMP = New ADODB.Recordset
        CABATEMP.CursorLocation = adUseClient
        
        SQLX$ = "SELECT * FROM CAJABANC "
        SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
        SQLX$ = SQLX$ + "AND NuClien = " & NC% & " "
        SQLX$ = SQLX$ + "AND Status >0 "
        SQLX$ = SQLX$ + "AND FechAcred >= '" & AYER & "'"
        SQLX$ = SQLX$ + "ORDER BY FechAcred ASC;"
        '
25      On Error Resume Next
        CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
        End If
        On Error GoTo 0
        '
        JJ& = 0: TRANSPOR# = 0
        FORMALTER.LINOSORT.Clear
        With CABATEMP
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
            On Error GoTo 0
            Do While Not .EOF
              If TRIM(!OpciMovim) <> "CH" Then GoTo 19 ' no es un cheque
              If IsNull(!FECHENTRE) = False Then
                If CVINT(!FECHENTRE) > 0 Then
                  GoTo 19    ' salio de cartera
                End If
              End If
              '
              FEACRED% = CVINT(!fechacred)
              BANUFE$ = !DescriMov
              ICHEQ# = !ValAbsComp
              '
              If TRIM$(Left$(BANUFE$, 12)) = "" Then
                Call REPLA(BANUFE$, "Cheque Nro:", 1, 11)
              End If
              RECHE& = !NUMECOMP
              '
              Z$ = "A" + Space$(80)
              Call REPLA(Z$, CSTRING$(MKI$(FEACRED%), 1, 5, 0, 2), 2, 5)
              Call REPLA(Z$, TRIM$(Left$(BANUFE$, 40)), 9, 40)
              Call REPLA(Z$, CSTRING$(MKD$(ICHEQ#), 4, 12, 2, 2), 49, 12)
              Call REPLA(Z$, CSTRING$(MKS$(RECHE&), 3, 8, 0, 2), 73, 8)
              SELCHE07.List1.AddItem Z$
              '
19          .MoveNext
            Loop
          End If
          On Error GoTo 0
        End With
        '
        SUCHETO# = 0
        For I& = 1 To SELCHE07.List1.ListCount
           Z$ = Mid$(SELCHE07.List1.LIST(I& - 1), 9)
           SELCHE07.LICHESEL.AddItem Z$
           SUCHETO# = SUCHETO# + XVALO(Mid$(Z$, 41, 12))
        Next I&
        '
        SELCHE07.Label4.Caption = CSTRING$(MKD$(SUCHETO#), 4, 11, 2, 2)
        SELCHE07.Label4.Refresh
        '
        Screen.MousePointer = 1
        SELCHE07.Show 1
        DoEvents
        '
     End If
   End If
   '
   Screen.MousePointer = 1
   '
End Sub

Sub CARESCUE()
Static VUELTA%
    '
    If ATENTI% = 1 Then
       Exit Sub
    End If
    '
10  Screen.MousePointer = 11
    NCX7% = XVALO(Text1.TEXT)
    If NCX7% > 0 Then
        List1.Clear
        List1.Refresh
        List2.Clear
        List2.Refresh
        Text2.TEXT = RASOCLI$(NCX7%)
        Text3.TEXT = DOMICLI$(NCX7%)
        Text4.TEXT = CPOSCLI$(NCX7%)
        Text5.TEXT = LOCACLI$(NCX7%)
        Text7.TEXT = TEL1CLI$(NCX7%)
        Label26 = CODICLI$(NCX7%)
        Label5.Caption = ""
        Text11.TEXT = CSTRING$(MKS$(LIMICRE(NCX7%)), 4, 11, 2, 2)
        Refresh
        Text6.TEXT = CSTRING$(MKD$(CHEPEACRE#(NCX7%)), 4, 11, 2, 2)
        Refresh
    End If
    '
100 List1.Clear
    List1.Refresh
    List2.Clear
    List2.Refresh
    Call REMPLAGRID(Me, GRID1, Picture1, Label8, List1)
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
    SALDO# = 0: SARRAS# = 0
    '
    RFF$ = RECFILT$("CUECOCLI", 1, MKI$(NCX7%))
    For KU& = 1 To Len(RFF$) Step 4
      RELE& = CVS(Mid$(RFF$, KU&, 4))
      XX$ = REGLEIDO$("CUECOCLI", RELE&)
      TIPOS$ = Mid$(XX$, 5, 16)
      IMPO# = ROUND(CVD(Mid$(XX$, 25, 8)))
      SALDO# = SALDO# + IMPO#
      FEX = CVI(Mid$(XX$, 3, 2))
      NROL& = CVS(Mid$(XX$, 21, 4))
      '
      Z$ = Space$(96)
      Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
      Call REPLA(Z$, TIPOS$, 11, 16)
      Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 27, 8)
      IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 13, 2, 2)
      SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
      If IMPO# >= 0.005 Then
          Call REPLA(Z$, IMPTEX$, 35, 13)
        ElseIf IMPO# <= (-0.005) Then
          Call REPLA(Z$, IMPTEX$, 50, 13)
        Else
          Call REPLA(Z$, "  ...... comp.anulado ...... ", 35, 29)
      End If
      Call REPLA(Z$, SALTEX$, 63, 14)
      List1.AddItem Z$
    Next KU&
    DoEvents
    Call REMPLAGRID(Me, GRID1, Picture1, Label8, List1)
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX7% & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "ORDER BY FechEmisi, fehoreal ASC;"
    'Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim CuecoTemp As ADODB.Recordset
    Set CuecoTemp = New ADODB.Recordset
    CuecoTemp.CursorLocation = adUseClient
25  On Error Resume Next
    CuecoTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    If NCX7% <> XVALO(Text1) Then GoTo 10  ' PARA ATAJAR ERROR CUANDO SE  INGRESA NUMERO DE CUENTA A MANO
    '
    With CuecoTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        On Error GoTo 0
        .MoveFirst
        Do While Not .EOF
          COCOLA$ = !codicom_lar
          TIX$ = Left$(!codicom_lar, 2)
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "AJ" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "RB" Then TIPOS$ = "Recibo Cobranza"
                              If TIX$ = "VC" Then TIPOS$ = "Venta Contado"
                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                              If TIX$ = "DC" Then TIPOS$ = "Descuento Caja"
                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito"
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1 '\\\OT-06-0112-WAR-30/03/06///
          IMPO# = !IDEBECOMP - !IHABECOMP
          '
          If MONRESU% > 1 Then
            MONEM% = ROUND(!Mone_Emis, 0)
            If MONEM% <> MONRESU% Then
              Z$ = Space$(128)
              FEX = CVINT%(!fechemisi)
              NROL& = !NUMECOMP
              Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
              Call REPLA(Z$, TIPOS$, 11, 16)
              Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 27, 8)
                TTX1$ = " Moneda: " + TRIM$(DEMONECO$(MONEM%)) + " "
                While Len(TTX1$) < 26: TTX1$ = "." + TTX1$ + ".": Wend
              Call REPLA(Z$, "  " + Left$(TTX1$, 26) + " ", 35, 29)
              GoTo 895
            End If
            TICAM = ROUND(!Tipo_Cam, 5)
            If TICAM < 0.05 Then TICAM = 1
            IMPO# = IMPO# / TICAM
          End If
          '
890       IMPO1# = ROUND(IMPO#)    '(Int(CDbl(100 * IMPO# + 0.5))) / 100
          SALDO# = SALDO# + IMPO1#  '(Int(100 * IMPO# + 0.5)) / 100
          FEX = CVINT%(!fechemisi)
          NROL& = !NUMECOMP
          '
          Z$ = Space$(128)
          Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
          Call REPLA(Z$, TIPOS$, 11, 16)
          Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 27, 8)
          IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 13, 2, 2)
          SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
          If IMPO# >= 0.005 Then
              Call REPLA(Z$, IMPTEX$, 35, 13)
            ElseIf IMPO# <= (-0.005) Then
              Call REPLA(Z$, IMPTEX$, 50, 13)
            Else
              Call REPLA(Z$, "  ...... comp.anulado ...... ", 35, 29)
          End If
          Call REPLA(Z$, SALTEX$, 63, 14)
          '
          If MODOMONEI% = 2 Then
             MONEM% = ROUND(!Mone_Emis, 0) Mod 6   ' enjuague gramy
             TICAM = ROUND(!Tipo_Cam, 5)
             If TICAM < 0.05 Then TICAM = 1
             '
             Call REPLA(Z$, SIMBOLPES$(MONEM%), 79, 2)
             If MONEM% > 1 Then
                Call REPLA(Z$, FORMATNUM$(TICAM, "F9.4"), 81, 9)
             End If
             Call REPLA(Z$, FORMATNUM$(MONEM%, "I2"), 103, 2)
             Call REPLA(Z$, COCOLA$, 105, 24)
             '
          End If
          'SE AGREGA EL NUMERO COMPLETO DEL AJUSTE PARA QUE SE PUEDA DIFERENCIAR SI
          'SE REALIZA UN AJUSTE EL MISMO DIA.
895       If TIX$ = "FC" Or TIX$ = "ND" Or TIX$ = "NC" Or TIX$ = "AS" Or TIX$ = "CC" Then
              Z$ = Z$ + Space$(128) + COCOLA$ ' Mid$(COCOLA$, 4, 13)
          End If
          List1.AddItem Z$
          '
         .MoveNext
        Loop
      End If
      On Error GoTo 0
990 End With
    CuecoTemp.Close
    Set CuecoTemp = Nothing
    Call REMPLAGRID(Me, GRID1, Picture1, Label8, List1)
    '
991 SARRAS# = SALDO#
    SUMADEU# = 0
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh
    '
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX7% & " "
    SQLX$ = SQLX$ + "AND (ABS(SaldAcre) >= 0.005 OR ABS(SaldDebe) >= 0.005 "
    If MODOMONEI% = 2 Then
         SQLX$ = SQLX$ + "OR ABS(SaldAcre_Dola) >= 0.005 OR ABS(SaldDebe_Dola) >= 0.005) "
       Else
         SQLX$ = SQLX$ + ") "       ' PARA CERRAR EL PARENTESIS
    End If
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    'Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim FacPenTemp As ADODB.Recordset
    Set FacPenTemp = New ADODB.Recordset
    FacPenTemp.CursorLocation = adUseClient
26  On Error Resume Next
    FacPenTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
    End If
    On Error GoTo 0

    '
    If NCX7% <> XVALO(Text1) Then GoTo 10   ' PARA ATAJAR ERROR CUANDO SE  INGRESA NUMERO DE CUENTA A MANO
    '
    SUDEUDOL# = 0
    With FacPenTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        On Error GoTo 0
        .MoveFirst
        Do While Not .EOF
          TIX$ = UCase$(Left$(!codicom_lar, 2))
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "AJ" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "RB" Then TIPOS$ = "Recibo Cobranza"
                              If TIX$ = "VC" Then TIPOS$ = "Venta Contado"
                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                              If TIX$ = "DC" Then TIPOS$ = "Descuento Caja"
                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito"
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1 ''\\\OT-06-0112-WAR-30/03/06///
                              '
          IMPO# = (-1) * ROUND(!ibrucomp) * SG%
          SUMAFAC# = SUMAFAC# + IMPO#
          '
          SALDO# = ROUND(!salddebe - !saldacre)
          SUMADEU# = SUMADEU# + SALDO#    ' (Int(100 * SALDO# + 0.5)) / 100
          '
          NROL& = XVALO(Mid$(!CODICOM_COR, 4, 6))
          NROX$ = CSTRING$(MKS$(NROL&), 3, 8, 0, 2)
          FEMI = CVINT(!fechemisi)
          FEVE = CVINT(!FEVENCIM)
          '
          IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 14, 2, 2)
          SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
          SALTEXDOL$ = ""
          '
          COCOLAR$ = !codicom_lar
          '
          If MODOMONEI% = 2 Then
             SALDOLA# = ROUND(!SaldDebe_Dola - !SaldAcre_Dola)
             SALTEXDOL$ = FORMATNUM$(SALDOLA#, "F13.2")
             GoTo 20
          End If
          '
          ' rutina incluida para procesamiento bimonetario
          RECABIN& = 0
          If Abs(Abs(IMPO#) - Abs(SALDO#)) < 0.005 Then ' NO ESTA CANCELADA
             MONEMX% = 0
             On Error Resume Next   ' por si no está el campo
             MONEMX% = !Mone_Emis
             On Error GoTo 0
             If MONEMX% < 1 Then
                COCOLAR$ = TRIM$(!codicom_lar)
                RFG$ = RECFILT$("CARBIMON", 2, MKI$(NCX7%))
                For KUL& = Len(RFG$) - 3 To 1 Step -4
                   RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
                   XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
                   If TRIM$(Mid$(XXC$, 45, 18)) = COCOLAR$ Then
                      MONEMX% = CVI(Mid$(XXC$, 31, 2))
                      GoTo 15
                   End If
                Next KUL&
15              If MONEMX% < 1 Then MONEMX% = 1
             End If
             '
             If MONEMX% > 1 Then
               ' LO QUE SIGUE ES UN ENJUAGUE PARA GRAMY PARA RECUPERAR LA MONEDA DE EMISIÓN
               If MONEMX% > ULTREG&("TAMONED") Then
                 Call COMUNI("Moneda no Válida en Comprobante '" + COCOLAR$ + "'.\Debe Seleccionar la Moneda de Emisión.")
                 Call SELECHO("TAMONED")
                 If XVALO(VALACT1$("TAMONED")) > 0 Then
                   MONEMX% = XVALO(VALACT1$("TAMONED"))
                   If RECABIN& > 0 Then
                     XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
                     Call REPLA(XXC$, MKI$(MONEMX%), 31, 2)
                     Call GRAREG("CARBIMON", XXC$, RECABIN&)
                   End If
                 End If
               End If
               '
               RFG$ = RECFILT$("CARBIMON", 2, MKI$(NCX7%))
               For KUL& = Len(RFG$) - 3 To 1 Step -4
                 RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
                 XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
                 If TRIM$(Mid$(XXC$, 45, 18)) = COCOLAR$ Then
                   SALDOPES# = ROUND(CVD(Mid$(XXC$, 65, 8)) - CVD(Mid$(XXC$, 73, 8)))
                   If Abs(SALDOPES# - SALDO#) <= 0.005 Then
                     SALTEX$ = FORMATNUM$(CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8)), "F14.2")
                     SBPES$ = TRIM$(SIMBOLPES$(MONEMX%))
                     If SBPES$ <> "" Then
                       Call REPLA(SALTEX$, SBPES$, 1, Len(SPPES$))
                     End If
                     GoTo 20
                   End If
                 End If
               Next KUL&
             End If
          End If
         
          ' fin rutina para procesamiento bimonetario
          '
20        Z$ = Space$(80)
          Call REPLA(Z$, TIPOS$, 1, 16)
          Call REPLA(Z$, NROX$, 17, 8)
          Call REPLA(Z$, FECHATEX$(FEMI), 27, 8)
          Call REPLA(Z$, IMPTEX$, 37, 14)
          Call REPLA(Z$, FECHATEX$(FEVE), 53, 8)
          Call REPLA(Z$, SALTEX$, 63, 14)
          Call REPLA(Z$, SALTEXDOL$, 77, 13)
          SUDEUDOL# = SUDEUDOL# + SALDOLA#
          '
          List2.AddItem Z$
          '
        .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
    '
    If MONRESU% <= 1 Then
       If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
          ATENTI% = 1
          Call REAPLICRE(NCX7%)
          If InStr(UCase$(EMPREAC$), "FERVI") > 0 Then ATENTI% = 0: GoTo 899 'provisorio FFFERVI
          GoTo 100
          '
          ' SE COMENTA PARA QUE SIEMPRE ACTUALICE LA COMPOSICION DE DEUDA. AL PREGUNTAR SE DEJABA LA POSIBILIDAD
          ' DE QUE EL USUARIO OPTE POR NO ACTUALIZAR LA DEUDA Y DESPUES EL REPORTE DE CUENTAS POR COBRAR APARECIA
          ' DESACTUALIZADO.
          '
'          TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta Elegida.\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Pagos y Créditos\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
'          If COMALTER%(TTXX$) = 2 Then
'              Call REAPLICRE(NCX7%)
'              GoTo 100
'            Else
'              ATENTI% = 0
'              Call CIERRARCH("*.*")
'              Call FINAL("")
'          End If
       End If
    End If
    '
899 List2.ListIndex = List2.ListCount - 1
    List2.Refresh
    Label5.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label5.Refresh
    Label6.Caption = CSTRING$(MKD$(SUMAFAC#), 4, 14, 2, 2)
    Label6.Refresh
    Label4.Caption = CSTRING$(MKD$(SUDEUDOL#), 4, 14, 2, 2)
    Label4.Refresh
    '
999 Screen.MousePointer = 1
    '
    If InStr(UCase$(EMPREAC$), "FERVI") > 0 Then GoTo 900 'provisorio FFFERVI

    If MONRESU% <= 1 Then
       If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
          ATENTI% = 1
          TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta"
          If Visible = True Then
              TTXX$ = TTXX$ + " Elegida."
            Else
              TTXX$ = TTXX$ + ":\   \'" + TRIM$(RASOCLI$(NCX7%)) + "'"
          End If
          TTXX$ = TTXX$ + "\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Créditos y Cobranzas\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
          '
          If COMALTER%(TTXX$) = 2 Then
              Call REAPLICRE(NCX7%)
              GoTo 100
            Else
              ATENTI% = 0
              Call CIERRARCH("*.*")
              Call FINAL("")
          End If
       End If
    End If
    '
900 DISPONI = XVALO(Text11.TEXT) - XVALO(Text10.TEXT) - XVALO(Text6.TEXT)
    Text12.ForeColor = &H80000008
    If DISPONI < 0 Then Text12.ForeColor = &HFF&
    Text12.TEXT = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
    '
    Screen.MousePointer = 1
    '
End Sub
'
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   '
   ' RESUMEN - OD-15/12/11
   ' SI EL SISTEMA NO ES MULTIMONEDA Y SE PRESIONO BOTON DERECHO
   ' PRESENTA LOS DOCUMENTOS APLICADOS A LA FACTURA
   '
   ' SI EL SISTEMA ES MULITIMONEDA Y SE PRESIONO BOTON DERECHO SE PRESENTAN OPCIONES
   ' PARA AJUSTE DE TIPO DE CAMBIO DE DOCUMENTOS Y MOSTRAR DOCUMENTOS APLICADOS
   '
   If Button = 2 Then
       If MODOMONEI% <> 2 Then
           OPX% = 1
         Else
           OPX% = COMALTER%("Opciones de Trabajo\\Documentos Aplicados\Ajuste de Cotización")
           If OPX% < 1 Or OPX% > 2 Then Exit Sub
       End If
       '
       If OPX% = 1 Then
            TTXX1$ = List1.TEXT
            Call MUESTRADOCUAPLI(TTXX1$)
          Else
            If MONRESU% > 1 Then
               Call COMUNI("Imposible Ajustar Tipo de Cambio\en Resumen en Moneda Extranjera.")
               Exit Sub
            End If
            If DERACCE%("AJUTICAM", "Ajuste de Tipo de Cambio Documentos Emitidos") < 2 Then
               Exit Sub
            End If
            '
            Call AJUCOTIZ
       End If
   End If
   '
End Sub

Sub AJUCOTIZ()
    '
    List1.ListIndex = List1.TopIndex + y / 190   ' OJO QUE EL 190 DEPENDE DE LA FUENTE
    TTXX$ = List1.TEXT
    COCOLA$ = TRIM$(Mid$(TTXX$, 105, 24))
    If COCOLA$ = "" Then
       Call MENSERR(24, "Imposible Ajustar Parámetros.")
       Exit Sub
    End If
    MONEM% = XVALO(Mid$(TTXX$, 103, 2))
    IMPESO# = XVALO(TRIM$(Mid$(TTXX$, 35, 28)))
    TICAM = XVALO(Mid$(TTXX$, 83, 9))
    If TICAM < 0.01 Then TICAM = 1
    VDEF$ = AJUSTI$(COCOLA$, 24) + Chr$(MONEM%) + Chr$(0) + MKD$(IMPESO#) + MKS$(TICAM)
20    OBX$ = OBJPLA$("AJUCOTFAC", VDEF$)
    If OBX$ = "" Then Exit Sub
    '
    MONEM% = Asc(Mid$(OBX$, 25, 1))
    If MONEM% < 1 Then GoTo 20
    TICAM = CVS(Mid$(OBX$, 35, 4))
    If MONEM% = 1 Then TICAM = 1
    If TICAM < 0.01 Then GoTo 20
    '
    NCX7% = XVALO(Text1)
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX7% & " "
    SQLX$ = SQLX$ + "AND CODICOM_LAR = '" & COCOLA$ & "' "
    'Set CUECO = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    Dim CUECO As ADODB.Recordset
    Set CUECO = New ADODB.Recordset
25  On Error Resume Next
    CUECO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 25
    End If
    On Error GoTo 0
    '
    With CUECO
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          IPESO# = ROUND(!IDEBECOMP - !IHABECOMP)
          '.Edit
          !Mone_Emis = MONEM%
          !Tipo_Cam = TICAM
          !Impo_Dola = ROUND#(IPESO# / TICAM)
          .Update
        .MoveNext
        Loop
      End If
    End With
    Set CUECO = Nothing
    '
    TOPI = List1.TopIndex
    LINDI = List1.ListIndex
    Call CARESCUE
    On Error Resume Next
    List1.TopIndex = TOPI
    List1.ListIndex = LINDI
    '
End Sub

Sub REAPLICRE(NCX%)
    '
    'If InStr(UCase$(EMPREAC$), "FERVI") > 0 Then Exit Sub 'provisorio FFFERVI
    Dim DOCULAR$(), DOCUCOR$(), IMPO#(), IORI#(), FFDOC%(), FVENC%(), STATDOC%(), CODIEM%(), CODISU%(), SALDOLA#()
    RAMAX& = 0
    ATENTI% = 1
    '
100 Screen.MousePointer = 11
    '
    NUDOCU% = 0
    RFF$ = RECFILT$("CUECOCLI", 1, MKI$(NCX%))
    For KU& = 1 To Len(RFF$) Step 4
      RELE& = CVS(Mid$(RFF$, KU&, 4))
      XX$ = REGLEIDO$("CUECOCLI", RELE&)
      NUDOCU% = NUDOCU% + 1
      If NUDOCU% > RAMAX& Then
        RAMAX& = RAMAX& + 128
        ReDim Preserve DOCULAR$(RAMAX&), DOCUCOR$(RAMAX&), IORI#(RAMAX&), IMPO#(RAMAX&), FFDOC%(RAMAX&), FVENC%(RAMAX&), STATDOC%(RAMAX&), CODIEM%(RAMAX&), CODISU%(RAMAX&), SALDOLA#(RAMAX&)
      End If
      DOCULAR$(NUDOCU%) = TRIM$(Mid$(XX$, 33, 24))
      DOCUCOR$(NUDOCU%) = TRIM$(Mid$(XX$, 57, 12))
      IMPO#(NUDOCU%) = ROUND(CVD(Mid$(XX$, 25, 8)))
      IORI#(NUDOCU%) = Abs(IMPO#(NUDOCU%))
      FFDOC%(NUDOCU%) = CVI(Mid$(XX$, 3, 2))
      CODIEM%(NUDOCU%) = CVI(Mid$(XX$, 69, 2))
      CODISU%(NUDOCU%) = CVI(Mid$(XX$, 71, 2))
      STATDOC%(NUDOCU%) = CVI(Mid$(XX$, 73, 2))
      SALDOLA#(NUDOCU%) = 0
      '
      NRODOC& = XVALO(Mid$(DOCULAR$(NUDOCU%), 11))
      For KKU& = 1 To List2.ListCount
        TTXY$ = List2.LIST(KKU& - 1)
        If XVALO(Mid$(TTXY$, 17, 8)) = NRODOC& Then
          If FECHANUM(Mid$(TTXY$, 27, 8)) = FFDOC%(NUDOCU%) Then
            If Abs(Abs(XVALO(Mid$(TTXY$, 37, 14))) - IORI#(NUDOCU%)) < 0.005 Then
              FVENC%(NUDOCU%) = FECHANUM(Mid$(TTXY$, 53, 8))
              GoTo 219
            End If
          End If
        End If
      Next KKU&
      FVENC%(NUDOCU%) = 0
219 Next KU&
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX% & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUD' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'RCOB') "
    SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC;"
    'Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim CuecoTemp As ADODB.Recordset
    Set CuecoTemp = New ADODB.Recordset
    CuecoTemp.CursorLocation = adUseClient
25  On Error Resume Next
    CuecoTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    SALDO# = 0: SARRAS# = 0
    With CuecoTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        On Error GoTo 0
        .MoveFirst
        Do While Not .EOF
          NUDOCU% = NUDOCU% + 1
          If NUDOCU% > RAMAX& Then
            RAMAX& = RAMAX& + 128
            ReDim Preserve DOCULAR$(RAMAX&), DOCUCOR$(RAMAX&), IORI#(RAMAX&), IMPO#(RAMAX&), FFDOC%(RAMAX&), FVENC%(RAMAX&), STATDOC%(RAMAX&), CODIEM%(RAMAX&), CODISU%(RAMAX&), SALDOLA#(RAMAX&)
          End If
          DOCULAR$(NUDOCU%) = TRIM$(!codicom_lar)   ' ANTES ESTABA SIN TRIM - 24/03/07
          DOCUCOR$(NUDOCU%) = !CODICOM_COR
          IMPO#(NUDOCU%) = ROUND(!IDEBECOMP - !IHABECOMP)
          IORI#(NUDOCU%) = !ValAbsComp
          FFDOC%(NUDOCU%) = CVINT(!fechemisi)
          CODIEM%(NUDOCU%) = !CodiEmpr
          CODISU%(NUDOCU%) = !CodiSucu
          SALDOLA#(NUDOCU%) = 0
          If IsNull(!Status) = False Then STATDOC%(NUDOCU%) = !Status
          '
          NRODOC& = XVALO(Mid$(DOCULAR$(NUDOCU%), 11))
          For KKU& = 1 To List2.ListCount
            TTXY$ = List2.LIST(KKU& - 1)
            If XVALO(Mid$(TTXY$, 17, 8)) = NRODOC& Then
              If FECHANUM(Mid$(TTXY$, 27, 8)) = FFDOC%(NUDOCU%) Then
                If Abs(Abs(XVALO(Mid$(TTXY$, 37, 14))) - IORI#(NUDOCU%)) < 0.005 Then
                  FVENC%(NUDOCU%) = FECHANUM(Mid$(TTXY$, 53, 8))
                  GoTo 119
                End If
              End If
            End If
          Next KKU&
          FVENC%(NUDOCU%) = 0
          '
119     .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
    ' AQUI SE BLANQUEABA SI PASABA POR CERO. vIGENTE HASTA EL 14/02/08
    ' SE SUPRIMIÓ PARA PRIMERO PRECESAR LAS APLICACIONES EXPLÍCITAS
    ' Y RECIÉN DESPUÉS LAS IMPLÍCITAS. RECLAMADO POR GUIDI CT-100330-FL
    '
'    ' blanquea si pasa por cero
'    FIN& = NUDOCU%
'    SUMAPEN# = 0
'    For KKI% = 1 To NUDOCU%
'      SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
'      If Abs(SUMAPEN#) < 0.005 Then
'        For KKJ% = 1 To KKI%
'          IMPO#(KKJ%) = 0
'        Next KKJ%
'      End If
'    Next KKI%
    '
    SQLX$ = "SELECT * FROM APLICOBRA "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX% & " "
    'Set ApliTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim ApliTemp As ADODB.Recordset
    Set ApliTemp = New ADODB.Recordset
    ApliTemp.CursorLocation = adUseClient
26  On Error Resume Next
    ApliTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
    End If
    On Error GoTo 0
    '
    With ApliTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          IAPLI# = !IMPOAPLI
          RECI$ = TRIM$(!codocap_lar)
          FACU$ = TRIM$(!codocor_lar)
          '
          DOCUAP% = 0: DOCUOR% = 0
          For KKI% = 1 To NUDOCU%
            If DOCULAR$(KKI%) = RECI$ And DOCUAP% < 1 Then
                DOCUAP% = KKI%
            End If
            If DOCULAR$(KKI%) = FACU$ And DOCUOR% < 1 Then
                DOCUOR% = KKI%
            End If
          Next KKI%

AA1$ = DOCULAR$(DOCUAP%)
aa2 = DOCULAR$(DOCUOR%)
          If DOCUAP% > 0 And DOCUOR% > 0 Then
            SIGAP% = 1: If IAPLI# < 0 Then SIGAP% = (-1)
            If SIGAP% * IAPLI# > SIGAP% * IMPO#(DOCUOR%) Then IAPLI# = IMPO#(DOCUOR%)
            'If SIGAP% * IAPLI# > (-1) * SIGAP% * IMPO#(DOCUAP%) Then
            '   IAPLI# = (-1) * IMPO#(DOCUAP%)   ' QUEDA POR INVESTIGAR SI
            'End If                              ' ESTA CONDICION CORRECTA - EPB 11/10/08
            IMPO#(DOCUOR%) = IMPO#(DOCUOR%) - IAPLI#
            IMPO#(DOCUAP%) = IMPO#(DOCUAP%) + IAPLI#
          End If
          '
'          IRECIMA# = 0: IFACMAX# = 0
'          For KKI% = 1 To NUDOCU%
'             If DOCULAR$(KKI%) = RECI$ Then
'               If IMPO#(KKI%) < 0 Then
'                 IRECIMA# = IRECIMA# - IMPO#(KKI%)
'               End If
'             End If
'             If Mid$(DOCULAR$(KKI%), 9, 8) = Mid$(FACU$, 9, 8) Then
'               If IMPO#(KKI%) > 0 Then
'                 IFACMAX# = IFACMAX# + IMPO#(KKI%)
'               End If
'             End If
'          Next KKI%
'          '
'          If IAPLI# > IRECIMA# Then IAPLI# = IRECIMA#
'          If IAPLI# > IFACMAX# Then IAPLI# = IFACMAX#
'          '
' If IAPLI# < 1 Then
' AAA = BBB
' End If
'          IRECIMA# = IAPLI#
'          IFACMAX# = IAPLI#
'          '
'          If IAPLI# > 0 Then
'             For KKI% = 1 To NUDOCU%
'                If DOCULAR$(KKI%) = RECI$ Then
'                   If IMPO#(KKI%) < 0 Then
'                      IAP1# = IRECIMA#
'                      If IAP1# > Abs(IMPO#(KKI%)) Then IAP1# = Abs(IMPO#(KKI%))
'                      IMPO#(KKI%) = IMPO#(KKI%) + IAP1#
'                      IRECIMA# = IRECIMA# - IAP1#
'                   End If
'                End If
'                If Mid$(DOCULAR$(KKI%), 9, 8) = Mid$(FACU$, 9, 8) Then
'                  If IMPO#(KKI%) > 0 Then
'                      IAP2# = IFACMAX#
'                      If IAP2# > Abs(IMPO#(KKI%)) Then IAP2# = Abs(IMPO#(KKI%))
'                      IMPO#(KKI%) = IMPO#(KKI%) - IAP2#
'                      IFACMAX# = IFACMAX# - IAP2#
'                   End If
'                End If
'             Next KKI%
'          End If
          '
        .MoveNext
        Loop
      End If
    End With
    '
    SUMAPEN# = 0
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
      U& = KKI%: Call TRACE(20, U&, FIN&)
      SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
      If Abs(SUMAPEN#) < 0.005 Then
        For KKJ% = 1 To KKI%
          IMPO#(KKJ%) = 0
        Next KKJ%
      End If
    Next KKI%
    '
    For KKI% = 1 To NUDOCU%
      For KKJ% = KKI% + 1 To NUDOCU%
        If (Abs(IMPO#(KKI%) + IMPO#(KKJ%))) < 0.005 Then
          IMPO#(KKI%) = 0
          IMPO#(KKJ%) = 0
        End If
      Next KKJ%
    Next KKI%
          
    KKM% = 0
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
      If STATDOC%(KKI%) = 8 Then
        SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
      End If
    Next KKI%
    '
    If Abs(SUMAPEN#) < 0.005 Then
      For KKI% = 1 To NUDOCU%
        If STATDOC%(KKI%) = 8 Then
          IMPO#(KKI%) = 0
        End If
      Next KKI%
      GoTo 19
    End If
    '
    For KKI% = NUDOCU% To 1 Step -1
      If STATDOC%(KKI%) = 8 Then
        If Sgn(IMPO#(KKI%)) <> Sgn(SUMAPEN#) Then
            'SUMAPEN# = SUMAPEN# - IMPO#(KKI%)
            IMPO#(KKI%) = 0
          Else
            If Abs(SUMAPEN#) >= Abs(IMPO#(KKI%)) Then
                SUMAPEN# = SUMAPEN# - IMPO#(KKI%)
                'IMPO#(KKI%) = 0
              Else
                IMPO#(KKI%) = SUMAPEN#
                SUMAPEN# = 0
            End If
        End If
      End If
    Next KKI%
    '
19  If MODOMONEI% = 2 Then       ' cargar saldo en dolares de carbimon.dat
       RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCX%))
       FIN& = Len(RFF$)
       For KKU& = 1 To FIN& Step 4
          Call TRACE(20, KKU&, FIN&)
          U& = CVS(Mid$(RFF$, KKU&, 4))
          XX$ = REGLEIDO$("CARBIMON", U&)
          NUDOCX$ = TRIM$(Mid$(XX$, 13, 18))
          DEBEDOL# = ROUND#(CVD(Mid$(XX$, 97, 8)))
          HABEDOL# = ROUND#(CVD(Mid$(XX$, 105, 8)))
          '
          For JJ% = 1 To NUDOCU%
            If Left$(DOCULAR$(JJ%), 16) = Left$(NUDOCX$, 16) Then
               SALDOLA#(JJ%) = SALDOLA#(JJ%) + DEBEDOL# - HABEDOL#
               GoTo 191
            End If
          Next JJ%
          '
          NUDOCU% = NUDOCU% + 1
          If NUDOCU% > RAMAX& Then
            RAMAX& = RAMAX& + 128
            ReDim Preserve DOCULAR$(RAMAX&), DOCUCOR$(RAMAX&), IORI#(RAMAX&), IMPO#(RAMAX&), FFDOC%(RAMAX&), FVENC%(RAMAX&), STATDOC%(RAMAX&), CODIEM%(RAMAX&), CODISU%(RAMAX&), SALDOLA#(RAMAX&)
          End If
          DOCULAR$(NUDOCU%) = NUDOCX$
          DOCUCOR$(NUDOCU%) = Left$(NUDOCX$, 3) + Right$(NUDOCX$, 8)
          IMPO#(NUDOCU%) = 0
          IORI#(NUDOCU%) = CVD(Mid$(XX$, 65, 8))
          FFDOC%(NUDOCU%) = CVI(Mid$(XX$, 33, 2))
          CODIEM%(NUDOCU%) = 0
          CODISU%(NUDOCU%) = 0
          STATDOC%(NUDOCU%) = 1
          SALDOLA#(NUDOCU%) = DEBEDOL# - HABEDOL#
          '
191    Next KKU&
    End If
    '
    For JJ% = 1 To NUDOCU%
      For JJ1% = JJ% + 1 To NUDOCU%
        If Abs(IMPO#(JJ%) + IMPO#(JJ1%)) < 0.005 Then
          If Abs(SALDOLA#(JJ%) + SALDOLA#(JJ1%)) < 0.005 Then
            IMPO#(JJ%) = 0
            SALDOLA#(JJ%) = 0
            IMPO#(JJ1%) = 0
            SALDOLA#(JJ1%) = 0
          End If
        End If
      Next JJ1%
    Next JJ%
    '
    If CONTROLGRABA% = 1 Then GoTo 24
    Screen.MousePointer = 1
    TTXX$ = "Se ha Completado la Reconstrucción\de la Composición de la Deuda de\   \'" + TRIM$(RASOCLI$(NCX%))
    'TTXX$ = "Se ha Completado la Reconstrucción\de la Composición de la Deuda de\   \'" + TRIM$(RASOCLI$((-1) * NCX%))
    '
    ' SE COMENTA PARA QUE SIEMPRE ACTUALICE LA COMPOSICION DE DEUDA. AL PREGUNTAR SE DEJABA LA POSIBILIDAD
    ' DE QUE EL USUARIO OPTE POR NO ACTUALIZAR LA DEUDA Y DESPUES EL REPORTE DE CUENTAS POR COBRAR APARECIA
    ' DESACTUALIZADO.
    '
'    If COMALTER%(TTXX$ + "'\   \Desea Grabar la Nueva Composición de la Deuda ?\\Cancelar\Grabar") <> 2 Then
'       GoTo 999
'    End If
    '
24  Screen.MousePointer = 11
    SQLX$ = "DELETE FROM SaDeuPen WHERE NuClien = " & NCX% & " "
     FLEXCONN.Execute (FILTSQL$(SQLX$))
    '
    ' AQUI ABRIR UN RECORDSET PARA GRABAR EN SADEUPEN
    Dim SaDeuPen As ADODB.Recordset
    Set SaDeuPen = New ADODB.Recordset
    SaDeuPen.CursorType = adOpenDynamic
    SaDeuPen.LockType = adLockPessimistic
    SQLX$ = "Select *  from SaDeuPen WHERE NuClien = " & NCX% & " "
27  On Error Resume Next
    SaDeuPen.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 27
    End If
    On Error GoTo 0
    '
    HOII% = HOY(HO$)
    FIN& = NUDOCU%
    LU$ = LUDAT$
    Call APERBASDAT("CABAN")
    NOCLOSE% = 1
    For U& = 1 To NUDOCU%
      Call TRACE(20, U&, FIN&)
      If Abs(IMPO#(U&)) >= 0.005 Or Abs(SALDOLA#(U&)) >= 0.005 Then
        IMORIGDOL# = 0: MONEMIOR% = 0
        If MODOMONEI% = 2 Then
          ' BUSCAR MONEDA E IMPORTE ORIGINAL DEL DOCUMENTO
          Call APERBASDAT("CABAN")
          SQLX$ = " SELECT * FROM CAJABANC "
          SQLX$ = SQLX$ + " WHERE CodiCom_Lar = '" & DOCULAR$(U&) & "' "
          Dim CAJJJABANC As ADODB.Recordset
          Set CAJJJABANC = New ADODB.Recordset
28        On Error Resume Next
          CAJJJABANC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 28
          End If
          On Error GoTo 0
          With CAJJJABANC
            'CajaBanc.FindFirst "CodiCom_Lar = '" & DOCULAR$(U&) & "' "
            'If CajaBanc.NoMatch = False Then
            On Error Resume Next
            If Not (.EOF And .BOF) Then
                On Error GoTo 0
                IMORIGDOL# = !Impo_Dola
                MONEMIOR% = !Mone_Emis Mod 6
                TICAMO = !Tipo_Cam
                VALACOM1# = !ValAbsComp
                On Error GoTo 0
                If MONEMIOR% > 1 Then
                  If Abs(IMORIGDOL#) < 0.005 Then
                    If TICAMO > 0.1 Then
                      IMORIGDOL# = ROUND(VALACOM1# / TICAMO)
                      '.Edit
                      !Impo_Dola = IMORIGDOL#
                      .Update
                    End If
                  End If
                End If
            End If
            On Error GoTo 0
          End With
          CAJJJABANC.Close
          Set CAJJJABANC = Nothing
        End If
        '
        CoComCor$ = DOCUCOR$(U&)
        CoComLar$ = DOCULAR$(U&)
'        CodiEmp% = CODIEM%(U&)
'        CodiSuc% = CODISU%(U&)
        NumeClie% = NCX%
        NumeProv% = 0
        FechEmi = CVDAT(FFDOC%(U&))
        IBruTot# = IORI#(U&)
        IDolTot# = IMORIGDOL#
        '
        If DOCULAR$(U&) <> DOCULARANT$ Then NuorVen% = 0
        NuorVen% = NuorVen% + 1
        '
        FVECIX% = FVENC%(U&)
        If FVECIX% < 1 Then
           COPAGI% = CPAGCLI%(NCX%)
           FEDOCU% = FFDOC%(U&)
           Call BUCONPA(LU$, COPAGI%, COPAG$, FREF%, PLA$)
           aaaa = FEDOCU%
           BBBB = FEDOCU%
           PLAPA% = 0: If Len(PLA$) > 0 Then PLAPA% = Asc(Left$(PLA$, 1))
           While DIENTRE%(aaaa, BBBB) < PLAPA%
             BBBB = DIPOST(BBBB)
           Wend
           FVECIX% = BBBB
        End If
        '
        FeVenc = CVDAT(FVECIX%)
        ImVenc# = IMPO#(U&)
        ImVDol# = SALDOLA#(U&)
        DOCULARANT$ = DOCULAR$(U&)
        '
        GoSub GRADEURECU
      End If
    Next U&
    SaDeuPen.Close
    NOCLOSE% = 0
    '
999 Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    ATENTI% = 0
    Exit Sub
    '
    '
GRADEURECU:
   SQLX$ = "SELECT * FROM SADEUPEN"
   Dim SADDDEUPEN As ADODB.Recordset
   Set SADDDEUPEN = New ADODB.Recordset
29 On Error Resume Next
   SADDDEUPEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 29
   End If
   On Error GoTo 0

   With SADDDEUPEN
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !NUCLIEN = NumeClie%
      !NuProve = NumeProv%
      !CodCliPro = Left$(CODICLI$(NumeClie% - NumeProv%), 16)
      !RaSoCliPro = Left$(RASOCLI$(NumeClie% - NumeProv%), 48)
      !CODICOM_COR = Left$(CoComCor$, 12)
      !codicom_lar = Left$(CoComLar$, 24)
      !fechemisi = FechEmi
        On Error Resume Next
      !Mone_Emis = MONEMIOR%
        On Error GoTo 0
        SIGX% = 1
        If Abs(ImVenc#) > 0 Then SIGX% = Sgn(ImVenc#)
      !ibrucomp = IBruTot# * SIGX%
      !IIvaComp = IIvaTot#
      !Ibru_Dola = IDolTot# * SIGX%
      !OrVencim = NuorVen%
      !FEVENCIM = FeVenc
      If ImVenc > 0 Then
          !salddebe = ImVenc#
          !saldacre = 0
        Else
          !saldacre = Abs(ImVenc#)
          !salddebe = 0
      End If
      If ImVDol > 0 Then
          !SaldDebe_Dola = ImVDol#
          !SaldAcre_Dola = 0
        Else
          !SaldAcre_Dola = Abs(ImVDol#)
          !SaldDebe_Dola = 0
      End If
      !Status = 0
     .Update
   End With
   SADDDEUPEN.Close
   Set SADDDEUPEN = Nothing
Return
'
End Sub

Private Sub Text7_GotFocus()
'\\\OT-06-0554-RG-13/12/06///
    Text1.SetFocus
'\\\OT-06-0554-RG-FIN///
End Sub

Private Sub Text8_GotFocus()
'\\\OT-06-0554-RG-13/12/06///
    Text1.SetFocus
'\\\OT-06-0554-RG-FIN///
End Sub

Private Sub Text9_GotFocus()
'\\\OT-06-0554-RG-13/12/06///
    Text1.SetFocus
'\\\OT-06-0554-RG-FIN///
End Sub

Sub MUESTRADOCUAPLI(RELIS$)
    '
    'MUESTRA LOS DOCUMENTOS APLICADOS A LA FACTURA  O NOTA DE DEBITO SELECCIONADA
    Screen.MousePointer = 11
    'LEO LOS IMPORTES Y FECHAS
    IMPODEBE# = XVALO(Mid$(RELIS$, 35, 13))
    IMPOHABE# = XVALO(Mid$(RELIS$, 50, 13))
    FECH% = FECHANUM(TRIM$(Mid$(RELIS$, 1, 8)))
    FEX = FETEXCOR1$(FECH%)
    NC% = XVALO(Text1)
    TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 16))

    'CASO FACTURAS, N.DEBITO. AJUSTES EN POSITIVO
    If IMPODEBE# > 0.005 Then
      '
      NUDOCUM$ = TRIM$(Mid$(RELIS$, 257, 13))
      TIMOV$ = "FVEN"
      If TIDOCUM$ = "Ajuste de Saldo" Then TIMOV$ = "AJUD"
      CONDI$ = "TIPOMOVIM =  '" & TIMOV$ & "'  AND fechemisi = '" & FEX & "' AND NUCLIEN = " & NC% & "  AND  codicom_lar like '%" & NUDOCUM$ & "%'"
      NUDOCU$ = VALOBADA("CAJABANC", "CODICOM_LAR", CONDI$, "NOMESS")
      '
      CONDI$ = "codocor_lar = '" & NUDOCU$ & "' AND NUCLIEN = " & NC% & ""
      '
      Call CARGALISEL("DOCUMAPL", "APLICOBRA", "codocap_lar/A24;IMPOAPLI/F14.2", CONDI$, "ORDER BY codocor_lar")
      '
    'CASO DE NOTAS DE CREDITO , RECIBOS , AJUSTES NEGATIVOS
    ElseIf IMPOHABE# > 0.005 Then
      NUDOCUM$ = TRIM$(Mid$(RELIS$, 27, 8))
      
      TIMOV$ = "FVEN"
      If TIDOCUM$ = "Recibo Cobranza" Then TIMOV$ = "RCOB"
      If TIDOCUM$ = "Ajuste de Saldo" Then TIMOV$ = "AJUD"
      CONDI$ = "TIPOMOVIM =  '" & TIMOV$ & "'  AND fechemisi = '" & FEX & "' AND NUCLIEN = " & NC% & "  AND  codicom_lar like '%" & NUDOCUM$ & "%'"
      NUDOCU$ = VALOBADA("CAJABANC", "CODICOM_LAR", CONDI$, "NOMESS")
      '
      If TIDOCUM$ = "Nota de Credito" Then
         CONDI$ = "CODOCOR_LAR = '" & NUDOCU$ & "' AND NUCLIEN = " & NC% & ""
         Call CARGALISEL("DOCUMAPL", "APLICOBRA", " CODOCAP_LAR/A24;IMPOAPLI/F14.2", CONDI$, "ORDER BY codocor_lar")
      Else
         CONDI$ = "CODOCAP_LAR = '" & NUDOCU$ & "' AND NUCLIEN = " & NC% & ""
         Call CARGALISEL("DOCUMAPL", "APLICOBRA", " CODOCOR_LAR/A24;IMPOAPLI/F14.2", CONDI$, "ORDER BY codocor_lar")
      End If
    End If
    
90  Screen.MousePointer = 1
    ENCAB$ = "Comprobante: " & NUDOCU$
    Call SELECHO("DOCUMAPL/" & ENCAB$)
    '
End Sub




