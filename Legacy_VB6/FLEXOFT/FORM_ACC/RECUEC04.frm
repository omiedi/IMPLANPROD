VERSION 5.00
Begin VB.Form RECUEC04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes de Cuenta Corriente de Deudores"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Height          =   7900
      Left            =   9840
      ScaleHeight     =   7845
      ScaleWidth      =   1110
      TabIndex        =   68
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
      Begin VB.Image Image3 
         Height          =   7980
         Left            =   0
         Picture         =   "RECUEC04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   975
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2450
      Index           =   12
      Left            =   12000
      TabIndex        =   61
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   10
      Left            =   9500
      TabIndex        =   59
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   8
      Left            =   16000
      TabIndex        =   58
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   7
      Left            =   2310
      TabIndex        =   57
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   9
      Left            =   3150
      TabIndex        =   56
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   11
      Left            =   4305
      TabIndex        =   55
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   13
      Left            =   6300
      TabIndex        =   54
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   15
      Left            =   7455
      TabIndex        =   53
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2450
      Index           =   0
      Left            =   1290
      TabIndex        =   52
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2450
      Index           =   1
      Left            =   3500
      TabIndex        =   51
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2450
      Index           =   2
      Left            =   4300
      TabIndex        =   50
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2450
      Index           =   3
      Left            =   5910
      TabIndex        =   49
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2450
      Index           =   4
      Left            =   7690
      TabIndex        =   48
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2450
      Index           =   5
      Left            =   9500
      TabIndex        =   47
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2450
      Index           =   6
      Left            =   12000
      TabIndex        =   46
      Top             =   2415
      Width           =   15
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00006060&
      Height          =   7995
      Left            =   9975
      ScaleHeight     =   7935
      ScaleWidth      =   1755
      TabIndex        =   34
      ToolTipText     =   "Re-Aplicación de Comprobantes"
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton Command8 
         Caption         =   "(2)"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   6120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Caption         =   "(1)"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   5760
         Visible         =   0   'False
         Width           =   615
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
         Picture         =   "RECUEC04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Listado Aplicación Diferencias de Cambio"
         Top             =   3885
         Width           =   640
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
         Picture         =   "RECUEC04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Finaliza la Aplicación"
         Top             =   315
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
         Picture         =   "RECUEC04.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   955
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
         Picture         =   "RECUEC04.frx":2AF6
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Re-Aplicacion de Creditos y Pagos a Cuenta"
         Top             =   4970
         Width           =   650
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
         Picture         =   "RECUEC04.frx":2C40
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Notas de Clientes"
         Top             =   3045
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   36
         Top             =   6930
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -315
            TabIndex        =   37
            Top             =   0
            Width           =   12000
         End
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
         Picture         =   "RECUEC04.frx":2F4A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Impresión de Cuenta Corriente"
         Top             =   2415
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -280
         Picture         =   "RECUEC04.frx":35B4
         Top             =   7370
         Width           =   1515
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Crédito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   7350
      TabIndex        =   25
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
         TabIndex        =   32
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
         TabIndex        =   30
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
         TabIndex        =   28
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
         TabIndex        =   26
         Text            =   " "
         Top             =   765
         Width           =   1590
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disp:"
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
         Index           =   1
         Left            =   60
         TabIndex        =   33
         Top             =   1635
         Width           =   645
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cheq:"
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
         Index           =   0
         Left            =   60
         TabIndex        =   31
         Top             =   1230
         Width           =   645
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L.Cre:"
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
         Index           =   2
         Left            =   60
         TabIndex        =   29
         Top             =   420
         Width           =   645
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.Cte:"
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
         Index           =   3
         Left            =   60
         TabIndex        =   27
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
      TabIndex        =   13
      Top             =   2415
      Width           =   9725
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   330
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   10515
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   43
            Top             =   50
            Width           =   10410
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0F0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   5580
         Left            =   0
         TabIndex        =   15
         Top             =   -1050
         Width           =   10515
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
            Left            =   -30
            TabIndex        =   16
            Top             =   1365
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   120
      TabIndex        =   2
      Top             =   210
      Width           =   7100
      Begin VB.CommandButton Command25 
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
         Left            =   1995
         TabIndex        =   64
         Top             =   1560
         Width           =   160
      End
      Begin VB.TextBox Text20 
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
         TabIndex        =   63
         Top             =   1560
         Width           =   645
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   8
         Text            =   " "
         Top             =   420
         Width           =   3690
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   7
         Text            =   " "
         Top             =   800
         Width           =   3690
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   6
         Text            =   " "
         Top             =   1180
         Width           =   750
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   5
         Text            =   " "
         Top             =   1180
         Width           =   2850
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00E0E0E0&
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
         Left            =   5250
         TabIndex        =   4
         Text            =   " "
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00E0E0E0&
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
         Left            =   5250
         TabIndex        =   3
         Text            =   " "
         Top             =   800
         Width           =   1695
      End
      Begin VB.Label Label27 
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
         Left            =   2265
         TabIndex        =   65
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label26 
         BackColor       =   &H00EEEEEE&
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
         TabIndex        =   42
         Top             =   1155
         Width           =   1035
      End
      Begin VB.Label label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax-Nbr:"
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
         Index           =   1
         Left            =   5250
         TabIndex        =   12
         Top             =   1260
         Width           =   960
      End
      Begin VB.Label label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal:"
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
         Index           =   0
         Left            =   210
         TabIndex        =   11
         Top             =   1575
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   315
         TabIndex        =   10
         Top             =   420
         Width           =   960
      End
      Begin VB.Label label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
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
         Index           =   0
         Left            =   5250
         TabIndex        =   9
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   2800
      Left            =   105
      ScaleHeight     =   2745
      ScaleWidth      =   9660
      TabIndex        =   18
      Top             =   5040
      Width           =   9725
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFE0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   5630
         Left            =   0
         TabIndex        =   20
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
            TabIndex        =   21
            Top             =   0
            Width           =   9725
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Left            =   7665
            TabIndex        =   60
            Top             =   2130
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Line Line15 
            X1              =   0
            X2              =   9710
            Y1              =   2420
            Y2              =   2420
         End
         Begin VB.Line Line14 
            X1              =   0
            X2              =   9660
            Y1              =   2100
            Y2              =   2100
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
            Left            =   25
            TabIndex        =   24
            Top             =   2130
            Width           =   1695
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Left            =   4330
            TabIndex        =   23
            Top             =   2130
            Width           =   1695
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Left            =   7455
            TabIndex        =   22
            Top             =   2130
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFE0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   330
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   10515
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   44
            Top             =   50
            Width           =   10410
         End
      End
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
      TabIndex        =   45
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
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu notacliente 
         Caption         =   "Nota de Cliente"
      End
      Begin VB.Menu Reaplicacion 
         Caption         =   "Re-Aplicacion de Creditos y Pagos a Cuenta"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu imprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu listado 
      Caption         =   "Listados"
      Begin VB.Menu ListadoAplic 
         Caption         =   "Listado Aplicación Diferencias de Cambio"
      End
   End
End
Attribute VB_Name = "RECUEC04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODOMONEI%
Dim MONRESU%

Private Sub Command6_Click()
   Command6.Enabled = False
   Call OPVARFA04.Command2_Click
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   '
    Call ABRECABAN
    '
    Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% = 0 Then Exit Sub
    '
    'SQLX$ = "delete from aplicobra where nuclien=" & NC%
    'CueCorri.Execute SQLX$
    'Exit Sub
    Dim comprofalta(40, 1) As String
    Dim ix As Integer
    '
    SQLX$ = "select * from cajabanc where nuclien=" & NC% & " and mone_emis=2 and (tipomovim='fven')"
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    'CueCorri.Execute ("update cajabanc set mone_emis=1 where codicom_lar='rb-x-0001-00001516' and nuclien=" & NC%)
    'CueCorri.Execute ("update cajabanc set mone_emis=1 where codicom_lar='rb-x-0001-00002064' and nuclien=" & NC%)
    'CueCorri.Execute ("update cajabanc set mone_emis=1 where codicom_lar='rb-x-0001-00002141' and nuclien=" & NC%)
    '
    With CABATEMP
        Do While Not .EOF
            docuor$ = SAFETEXT$(!CODICOM_LAR)
            impoor# = XVALO(!IDEBECOMP)
            RFF$ = RECFILT$("carbimon", 8, docuor$)
            If RFF$ <> "" Then
                impode# = 0: impoha# = 0: saldope# = 0: ix = 0
                SQLX$ = "delete from aplicobra where codocor_lar='" & docuor$ & "' and nuclien=" & NC%
                CueCorri.Execute SQLX$
                For II% = 1 To Len(RFF$) Step 4
                    REGI& = CVS(Mid$(RFF$, II%, 4))
                    XX$ = REGLEIDO$("carbimon", REGI&)
                    impoha# = ROUND#(CVD(Mid$(XX$, 73, 8)), 2)
                    CoComLar = Mid$(XX$, 45, 18)
                    MONE% = CVI(Mid$(XX$, 63, 2))
                    If Mid$(CoComLar, 1, 2) = "AS" And InStr(1, CoComLar, "2239") > 0 Then
                        SQLX$ = "delete from cajabanc where codicom_lar='" & CoComLar & "' and nuclien=" & NC%
                        CueCorri.Execute SQLX$
                    End If
                    If Mid$(CoComLar, 1, 2) = "AS" And InStr(1, CoComLar, "2210") > 0 Then
                        SQLX$ = "delete from cajabanc where codicom_lar='" & CoComLar & "' and nuclien=" & NC%
                        CueCorri.Execute SQLX$
                    End If
                    If Mid$(CoComLar, 1, 2) = "AS" And InStr(1, CoComLar, "2205") > 0 Then
                        SQLX$ = "delete from cajabanc where codicom_lar='" & CoComLar & "' and nuclien=" & NC%
                        CueCorri.Execute SQLX$
                    End If
                    If impoha# > 0.005 Then
                        Debug.Print docuor$
                        CoComCor$ = SAFETEXT$(VALOBADAMDB("cajabanc", "codicom_cor", "codicom_lar='" & CoComLar & "' and nuclien=" & NC%))
                        DocOrLar$ = docuor$
                        DocOrCor$ = SAFETEXT$(VALOBADAMDB("cajabanc", "codicom_cor", "codicom_lar='" & DocOrLar$ & "' and nuclien=" & NC%))
                        CodiEmp% = 0
                        CodiSuc% = 0
                        NumeClie% = NC%
                        ImpApli# = ROUND#(impoha#, 2)
                        '
                        If saldope# > 0.005 Then
                            TotOrig# = ROUND#(impode#, 2)

                            If (ImpApli# - saldope#) > 0.005 Then ImpApli# = saldope#

                            IvaOrig# = 0
                            saldope# = saldope# - ImpApli#

                            Call GRAPLICOB
                        End If
                    Else
                        If CoComLar = docuor$ Then
                            impode# = ROUND#(CVD(Mid$(XX$, 65, 8)), 2)
                            saldope# = impode#
                        Else
                            ix = ix + 1
                            comprofalta(ix, 0) = CoComLar
                            comprofalta(ix, 1) = Str$(CVD(Mid$(XX$, 65, 8)))
                        End If
                    End If

                Next II%
                '
                If ix > 0 Then
                    For iii% = 1 To ix
                        docacancelar$ = comprofalta(iii%, 0)
                        acancelar# = ROUND#(XVALO(comprofalta(iii%, 1)), 2)
                        SQLX$ = "delete from aplicobra where codocor_lar='" & docacancelar$ & "' and nuclien=" & NC%
                        CueCorri.Execute SQLX$

                        For II% = 1 To Len(RFF$) Step 4
                            REGI& = CVS(Mid$(RFF$, II%, 4))
                            XX$ = REGLEIDO$("carbimon", REGI&)
                            impoha# = ROUND#(CVD(Mid$(XX$, 73, 8)), 2)
                            CoComLar = Mid$(XX$, 45, 18)
                            '
                            If impoha# > 0.005 Then
                                saldocance# = ROUND#(XVALO(VALOBADAMDB("aplicobra", "sum(impoapli)", "CoDocAp_Lar='" & CoComLar & "' and codocor_lar='" & docuor$ & "'")), 2)
                                saldopen# = impoha# - saldocance#
                                If saldopen# > 0.005 Then
                                    ImpApli# = acancelar#
                                    If acancelar# - saldopen# > 0.005 Then ImpApli# = saldopen#

                                    CoComCor$ = SAFETEXT$(VALOBADAMDB("cajabanc", "codicom_cor", "codicom_lar='" & CoComLar & "' and nuclien=" & NC%))
                                    DocOrLar$ = docacancelar$
                                    DocOrCor$ = SAFETEXT$(VALOBADAMDB("cajabanc", "codicom_cor", "codicom_lar='" & DocOrLar$ & "' and nuclien=" & NC%))
                                    CodiEmp% = 0
                                    CodiSuc% = 0
                                    NumeClie% = NC%
                                    TotOrig# = acancelar#
                                    IvaOrig# = 0
                                    Call GRAPLICOB
                                End If
                            End If
                        Next II%
                    Next iii%
                End If
                ix = 0
            End If
20            .MoveNext
        Loop
    End With
    '
    Call COMUNI("Proceso Finalizado")
    '
End Sub



Private Sub Command8_Click()

                SQLX$ = "delete from cajabanc where codicom_cor='AS-002205-01' and nuclien=105"
                CueCorri.Execute SQLX$


End Sub

Private Sub imprimir_Click()
   Call Command3_Click
End Sub

Private Sub Label5_Change()
   Text10 = TRIM$(Label5)
End Sub

Private Sub ListadoAplic_Click()
   Call Command6_Click
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
       If ECOARCH$("DEUDOR1", MKI$(NCX%)) <> "" Then
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

Private Sub notacliente_Click()
    Command5_Click
End Sub

Private Sub Reaplicacion_Click()
   Call Command4_Click
End Sub

Private Sub salir_Click()
  Call Command1_Click
End Sub

Private Sub Text10_Change()
AAA = bbb
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
        X$ = REGLEIDO$("DEUDOR1", 1)      ' fuerza la
        Call GRAREG("DEUDOR1", X$, 1)     ' reindexacion
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
    With RESCUC04
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
    RESCUC04.Show 1
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
    If NC% < 1 Or ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
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
    Call CARESCUE
    DoEvents
    '
End Sub

Private Sub Command5_Click()
    NC% = XVALO(TRIM$(Text1.TEXT))
    If NC% > 0 And ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
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
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
          Me.BackColor = &HFCD7B6
          Picture4.Visible = True
          Me.Frame1.BackColor = &HFCD7B6
          Me.Frame2.BackColor = &HFCD7B6
          Me.Frame3.BackColor = &HFCD7B6
          Me.Frame4.BackColor = &HFCD7B6
          Me.Frame5.BackColor = &HFCD7B6
          Me.Frame6.BackColor = &HFCD7B6
    End If
    If CONTROL("FACTURA", "SISTMANT") = "SI" Then
       Reaplicacion.Visible = False
       ListadoAplic.Enabled = False
    End If
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
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

Private Sub List1_DBLClick()
   '
   RELIS$ = List1.TEXT
   TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 16))
   If TIDOCUM$ = "Recibo Cobranza" Then TIDOCUM$ = "Recibo"
   If TIDOCUM$ = "Ajuste de Saldo" Then TIDOCUM$ = "Ajuste Saldo Cliente"
   NUDOCU$ = " " + TRIM$(Mid$(RELIS$, 27, 8)) + " "
   If InStr(UCase$(TIDOCUM$), "FACTURA") > 0 Or InStr(UCase$(TIDOCUM$), "NOTA DE") > 0 Then
     'NUDOCU$ = TRIM$(NUDOCU$)
     'While Len(NUDOCU$) < 8: NUDOCU$ = "0" + NUDOCU$: Wend
     NUDOCU$ = TRIM$(NUDOCU$)
     While Len(NUDOCU$) < 8: NUDOCU$ = "0" + NUDOCU$: Wend
     If TRIM$(Mid$(RELIS$, 257, 18)) <> "" Then
        NUDOCU2$ = TRIM$(Mid$(RELIS$, 257, 18))
        NUDOCU$ = Mid$(NUDOCU2$, 4, 13)
     End If
   End If
   RACLI$ = TRIM$(Left$(RASOCLI$(XVALO(Text1.TEXT)), 8))
   FECHADOC$ = Mid$(RELIS$, 1, 5)
   Screen.MousePointer = 11
   For UI& = ULTREG&("PDOCLST") To 1 Step -1
     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
     If InStr(XX$, TIDOCUM$) > 4 Then
       If InStr(XX$, NUDOCU$) > 4 Then
         'If InStr(XX$, RACLI$) > 16 Or InStr(XX$, FECHADOC$) > 0 Then
         'If InStr(XX$, FECHADOC$) > 0 Then
           On Error Resume Next
           DOCUNU& = CVS(Left$(XX$, 4))
           If Err Then
              NUERRO = Err.Number
              On Error GoTo 0
              Call MENSERR(24, "Se ha Producido Error " & NUERRO & " (1).\Imposible Abrir Documento en este Momento.\Base de Datos de Documentos Digitalizados\Probablemente Dañada. Consulte.")
              Exit Sub
           End If
           On Error GoTo 0
           '
           GoTo 5
         'End If
       End If
     End If
   Next UI&
   Screen.MousePointer = 1
   Call COMUNI("Imposible Abrir Documento Desde esta Aplicación")
   Exit Sub
   '
5  For KKI% = 0 To 5
     EDITFORM.Picture1(KKI%).Cls
     EDITFORM.Picture1(KKI%).DrawWidth = 3
     EDITFORM.Picture1(KKI%).DrawStyle = 0
     EDITFORM.Picture1(KKI%).FillStyle = 1
   Next KKI%
   '
   REDOCU$ = REGACT$("PDOCLST")
   URE& = ULTREG&("PDOCSPL")
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   On Error Resume Next
   DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
   If Err Then
      NUERRO = Err.Number
      On Error GoTo 0
      Call MENSERR(24, "Se ha Producido Error " & NUERRO & " (2)\al Intentar Leer el Registro " & L& & " de 'PDOCSPL'.\Imposible Abrir Documento en este Momento.\Base de Datos de Documentos Digitalizados\Probablemente Dañada. Consulte.")
      Exit Sub
   End If
   On Error GoTo 0
   '
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Screen.MousePointer = 1
   Call MENSERR(24, "Documento no Registrado")
   Exit Sub
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     On Error Resume Next
     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
     If Err Then
        NUERRO = Err.Number
        On Error GoTo 0
        Call MENSERR(24, "Se ha Producido Error " & NUERRO & " (3)\al Intentar Leer el Registro " & DCC& & " de 'PDOCSPL'.\Imposible Abrir Documento en este Momento.\Base de Datos de Documentos Digitalizados\Probablemente Dañada. Consulte.")
        Exit Sub
     End If
     On Error GoTo 0
     '
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  CAPAGINAS% = 0
   For KKL& = L& To URE&
     TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
     End If
   Next KKL&
   '
95 SPOOLSTRING$ = TPSP$
   For KKI% = 0 To 5
     EDITFORM.Picture1(KKI%).Cls
     EDITFORM.Picture1(KKI%).Top = 0
     EDITFORM.Picture1(KKI%).Refresh
     EDITFORM.Picture1(KKI%).Height = 7170
     EDITFORM.Picture1(KKI%).Width = 11700
     EDITFORM.Picture1(KKI%).DrawWidth = 3
     EDITFORM.Picture1(KKI%).DrawStyle = 0
     EDITFORM.Picture1(KKI%).FillStyle = 1
   Next KKI%
   ALTUPAGINA = 7170
   TOPEPAGINA = 0
   PAGINACTIVA% = 0
   CAPAGINAS% = 0
   '
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
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
   Screen.MousePointer = 1
   PAGINACTIVA% = 0
   For KKI% = 0 To 5
     EDITFORM.Picture1(KKI%).Visible = False
     If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
   Next KKI%
   EDITFORM.Picture1(PAGINACTIVA%).Refresh
   EDITFORM.VScroll1.Min = 0
   EDITFORM.VScroll1.Value = 0
   EDITFORM.VScroll1.Max = 1
   If ALTUPAGINA > EDITFORM.Frame1.Height Then
      EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
   End If
   EDITFORM.Command3.Enabled = False
   'EDITFORM.Command1.Enabled = False
   EDITFORM.Show 1
   '
End Sub

Private Sub Text1_Change()
   List1.Clear
   List2.Clear
   Text20 = "0"
   If Abs(XVALO(Text1)) < 32767 Then
     NCX% = XVALO(Text1)
     If NCX% > 0 Then
       If ECOARCH$("DEUDOR1", MKI$(NCX%)) <> "" Then
         NC% = NCX%
         If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
            Call FILTERFILE("SUCENTRE", "", 5, MKI$(NCX%))
         End If
         Call CARESCUE
         LMCR# = LIMICRE(NC%)  'MUESTRA QUE LA CUENTA DEL CLIENTE ESTA SUSPENDIDA
         If LMCR# < 0 Then
            Text11 = "Suspendida"
         End If
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
        SELCHE04.Caption = "Cheques Pendientes de Acreditaciòn: " + RASOCLI$(NC%)
        SELCHE04.List1.Clear
        SELCHE04.LICHESEL.Clear
        '
        FEXI% = DIANTE(DIANTE(HOY(HOS$)))
        AYER = CDbl(CVDAT(FEXI%))
15      Call APERBASDAT("CABAN")
        SQLX$ = "SELECT * FROM CAJABANC "
        SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
        SQLX$ = SQLX$ + "AND NuClien = " & NC% & " "
        SQLX$ = SQLX$ + "AND Status >0 "
        SQLX$ = SQLX$ + "AND CDBL(FechAcred) >= " & AYER & " "
        SQLX$ = SQLX$ + "ORDER BY FechAcred ASC;"
        '
        Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
        '
        JJ& = 0: TRANSPOR# = 0
        FORMALTER.LINOSORT.Clear
        With CABATEMP
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
            On Error GoTo 0
            Do While Not .EOF
              If TRIM(!OPCIMOVIM) <> "CH" Then GoTo 19 ' no es un cheque
              If IsNull(!FECHENTRE) = False Then
                If CVINT(!FECHENTRE) > 0 Then
                  GoTo 19    ' salio de cartera
                End If
              End If
              '
              FEACRED% = CVINT(!FechAcred)
              BANUFE$ = !Descrimov
              ICHEQ# = !ValAbsComp
              '
              If TRIM$(Left$(BANUFE$, 12)) = "" Then
                Call REPLA(BANUFE$, "Cheque Nro:", 1, 11)
              End If
              RECHE& = !NumeComp
              '
              Z$ = "A" + Space$(80)
              Call REPLA(Z$, CSTRING$(MKI$(FEACRED%), 1, 5, 0, 2), 2, 5)
              Call REPLA(Z$, TRIM$(Left$(BANUFE$, 40)), 9, 40)
              Call REPLA(Z$, CSTRING$(MKD$(ICHEQ#), 4, 12, 2, 2), 49, 12)
              Call REPLA(Z$, CSTRING$(MKS$(RECHE&), 3, 8, 0, 2), 73, 8)
              SELCHE04.List1.AddItem Z$
              '
19          .MoveNext
            Loop
          End If
          On Error GoTo 0
        End With
        '
        SUCHETO# = 0
        For i& = 1 To SELCHE04.List1.ListCount
           Z$ = Mid$(SELCHE04.List1.List(i& - 1), 9)
           SELCHE04.LICHESEL.AddItem Z$
           SUCHETO# = SUCHETO# + XVALO(Mid$(Z$, 41, 12))
        Next i&
        '
        SELCHE04.Label4.Caption = CSTRING$(MKD$(SUCHETO#), 4, 11, 2, 2)
        SELCHE04.Label4.Refresh
        '
        Screen.MousePointer = 1
        SELCHE04.Show 1
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
    SALDO# = 0: SARRAS# = 0
    SUCUSEL& = XVALO(Text20)
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
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX7% & " "
    If SUCUSEL& > 0 Then SQLX$ = SQLX$ + " AND Sucu_Clie = " & SUCUSEL& & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "ORDER BY FechEmisi, fehoreal ASC;"
    Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
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
          COCOLA$ = !CODICOM_LAR
          TIX$ = Left$(!CODICOM_LAR, 2)
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
              FEX = CVINT%(!FechEmisi)
              NROL& = !NumeComp
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
          FEX = CVINT%(!FechEmisi)
          NROL& = !NumeComp
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
          If !TipoMovim = "RCOB" Then
            NTL% = 0
            On Error Resume Next
            NTL% = (XVALO(!NUMETALO) Mod 10)
            On Error GoTo 0
            If NTL% > 0 Then
              Call REPLA(Z$, "-" + TRIM$(Str$(NTL%)), 35, 2)
            End If
          End If
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
895       If TIX$ = "FC" Or TIX$ = "ND" Or TIX$ = "NC" Or TIX$ = "AS" Or TIX$ = "CC" Then
              Z$ = Z$ + Space$(128) + COCOLA$ ' Mid$(COCOLA$, 4, 13)
          End If
          List1.AddItem Z$
          '
898     .MoveNext
        Loop
      End If
      On Error GoTo 0
990 End With
    '
991 SARRAS# = SALDO#
    SUMADEU# = 0
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh
    '
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX7% & " "
    If SUCUSEL& > 0 Then SQLX$ = SQLX$ + " AND Sucu_Clie = " & SUCUSEL& & " "
    SQLX$ = SQLX$ + "AND (ABS(SaldAcre) >= 0.005 OR ABS(SaldDebe) >= 0.005 "
    SQLX$ = SQLX$ + "OR ABS(SaldAcre_Dola) >= 0.005 OR ABS(SaldDebe_Dola) >= 0.005) "
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
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
          TIX$ = UCase$(Left$(!CODICOM_LAR, 2))
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
          IMPO# = (-1) * ROUND(!IBruComp) * SG%
          SUMAFAC# = SUMAFAC# + IMPO#
          '
          SALDO# = ROUND(!SaldDebe - !SaldAcre)
          SUMADEU# = SUMADEU# + SALDO#    ' (Int(100 * SALDO# + 0.5)) / 100
          '
          NROL& = XVALO(Mid$(!CodiCom_Cor, 4, 6))
          NROX$ = CSTRING$(MKS$(NROL&), 3, 8, 0, 2)
          FEMI = CVINT(!FechEmisi)
          FEVE = CVINT(!FEVENCIM)
          '
          IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 14, 2, 2)
          SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
          SALTEXDOL$ = ""
          '
          COCOLAR$ = !CODICOM_LAR
          '
          If MODOMONEI% = 2 Then
             SALDOLA# = ROUND(!SaldDebe_Dola - !SaldAcre_Dola)
             SALTEXDOL$ = FORMATNUM$(SALDOLA#, "F13.2")
             
             NUFACTU$ = SAFETEXT$(!CODICOM_LAR): NCLIENTE% = XVALO(!NuClien)
             MONEMX% = XVALO(VALOBADAMDB("CAJABANC", "Mone_Emis", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'FVEN' AND NUCLIEN = " & CStr(NCLIENTE%), "NOMESS"))
             If MONEMX% > 2 Then
                SALDOLA# = 0
                SALTEXDOL$ = FORMATNUM$(SALDOLA#, "F13.2")
             End If
             
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
                COCOLAR$ = TRIM$(!CODICOM_LAR)
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
20        Z$ = Space$(90)
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
998     .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
    If MONRESU% <= 1 Then
      If SUCUSEL& < 1 Then
        If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
          ATENTI% = 1
          TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta Elegida.\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Pagos y Créditos\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
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
    End If
    '
    List2.ListIndex = List2.ListCount - 1
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
    If MONRESU% <= 1 Then
      If SUCUSEL& < 1 Then
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
    End If
    '
    DISPONI = XVALO(Text11.TEXT) - XVALO(Text10.TEXT) - XVALO(Text6.TEXT)
    Text12.ForeColor = &H80000008
    If DISPONI < 0 Then Text12.ForeColor = &HFF&
    Text12.TEXT = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
    '
    Screen.MousePointer = 1
    '
End Sub
'
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If MODOMONEI% <> 2 Then Exit Sub
   If MONRESU% > 1 Then Exit Sub
   If DERACCE%("AJUTICAM", "Ajuste de Tipo de Cambio Documentos Emitidos") < 2 Then
      Exit Sub
   End If
   '
   If Button = 2 Then
      List1.ListIndex = List1.TopIndex + Y / 190   ' OJO QUE EL 190 DEPENDE DE LA FUENTE
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
      Set CUECO = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
      '
      With CUECO
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          On Error GoTo 0
          Do While Not .EOF
            IPESO# = ROUND(!IDEBECOMP - !IHABECOMP)
            .Edit
            !Mone_Emis = MONEM%
            !Tipo_Cam = TICAM
            !Impo_Dola = ROUND#(IPESO# / TICAM)
            .Update
          .MoveNext
          Loop
        End If
      End With
      '
      TOPI = List1.TopIndex
      LINDI = List1.ListIndex
      Call CARESCUE
      On Error Resume Next
      List1.TopIndex = TOPI
      List1.ListIndex = LINDI
      '
   End If
End Sub

Sub RECUTELE()
    '
'    call APERBASDAT("CABAN")
'    SQLY$ = "NuClien = 136 "
'    SQLY$ = SQLY$ + "AND CodiCom_Cor = 'FC-001361-A' "
'    With CajaBanc
'      .FindFirst SQLY$
'895   If .NoMatch = False Then
'        .Edit
'          VANUE = !ValAbsComp / 2.896
'        !ValAbsComp = VANUE
'          VANUE = !IDEBECOMP / 2.896
'        !IDEBECOMP = VANUE
'          VANUE = !IHABECOMP / 2.896
'        !IHABECOMP = VANUE
'        .Update
'        .FindNext SQLY$
'        GoTo 895
'      End If
'    End With
    '
    SQLY$ = "NuClien = 144 "
    SQLY$ = SQLY$ + "AND CodiCom_Cor = 'FC-001382-A' "
    With CajaBanc
      .FindFirst SQLY$
895   If .NoMatch = False Then
        .Edit
          VANUE = !ValAbsComp / 2.865
        !ValAbsComp = VANUE
          VANUE = !IDEBECOMP / 2.865
        !IDEBECOMP = VANUE
          VANUE = !IHABECOMP / 2.865
        !IHABECOMP = VANUE
          VANUE = !BASEIMPO / 2.865
        !BASEIMPO = VANUE
        .Update
        .FindNext SQLY$
        GoTo 895
      End If
    End With
    '
'    SQLY$ = "NuClien = 136 "
'    SQLY$ = SQLY$ + "AND CodiCom_Cor = 'FC-001361-A' "
'    With CajaBanc
'      .FindFirst SQLY$
'900   If .NoMatch = False Then
'        .Edit
'          VANUE = !BASEIMPO / 2.896
'        !BASEIMPO = VANUE
'        .Update
'        .FindNext SQLY$
'        GoTo 900
'      End If
'    End With
    '
End Sub

Sub REAPLICRE(NCX%)
    '
    Dim DOCULAR$(), DOCUCOR$(), IMPO#(), IORI#(), FFDOC%(), FVENC%(), STATDOC%(), CODIEM%(), CODISU%(), SALDOLA#(), MOEMI%()
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
        TTXY$ = List2.List(KKU& - 1)
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
    SQLX$ = SQLX$ + "ORDER BY FeHoReal ASC;"
    Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
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
            ReDim Preserve DOCULAR$(RAMAX&), DOCUCOR$(RAMAX&), IORI#(RAMAX&), IMPO#(RAMAX&), FFDOC%(RAMAX&), FVENC%(RAMAX&), STATDOC%(RAMAX&), CODIEM%(RAMAX&), CODISU%(RAMAX&), SALDOLA#(RAMAX&), MOEMI%(RAMAX&)
          End If
          DOCULAR$(NUDOCU%) = TRIM$(!CODICOM_LAR)   ' ANTES ESTABA SIN TRIM - 24/03/07
          If InStr(1, DOCULAR$(NUDOCU%), "1026") > 0 Then
          A = B
          End If
          DOCUCOR$(NUDOCU%) = !CodiCom_Cor
          IMPO#(NUDOCU%) = ROUND(!IDEBECOMP - !IHABECOMP)
          IORI#(NUDOCU%) = !ValAbsComp
          FFDOC%(NUDOCU%) = CVINT(!FechEmisi)
          CODIEM%(NUDOCU%) = !CodiEmpr
          CODISU%(NUDOCU%) = !CodiSucu
          SALDOLA#(NUDOCU%) = 0
          MOEMI%(NUDOCU%) = XVALO(!Mone_Emis)
          
If Left$(DOCULAR$(NUDOCU%), 2) = "NC" Then
MONEM% = ROUND(!Mone_Emis, 0) Mod 6   ' enjuague gramy
TICAM = ROUND(!Tipo_Cam, 5)
If MONEM% > 1 Then
  If TICAM > 0.1 Then
    SALDOLA#(NUDOCU%) = IMPO#(NUDOCU%) / TICAM
  End If
End If
End If
          If IsNull(!Status) = False Then STATDOC%(NUDOCU%) = !Status
          '
          NRODOC& = XVALO(Mid$(DOCULAR$(NUDOCU%), 11))
          For KKU& = 1 To List2.ListCount
            TTXY$ = List2.List(KKU& - 1)
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
    SQLX$ = "SELECT * FROM APLICOBRA "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX% & " "
    Set ApliTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With ApliTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          IAPLI# = !ImpoApli
          RECI$ = TRIM$(!CoDocAp_Lar)
          FACU$ = TRIM$(!CoDocOr_Lar)
          If InStr(1, FACU$, "1026") > 0 Or InStr(1, RECI$, "750") > 0 Then
          A = B
          End If
          'CONDI$ = "CODICOM_LAR='" & FACU$ & "'"
          CONDI$ = "CODICOM_LAR='" & RECI$ & "' and ihabecomp>0"
          MONE% = XVALO(VALOBADAMDB("CAJABANC", "Mone_Emis", CONDI$))
          OPCI$ = VALOBADAMDB("CAJABANC", "OPCIMOVIM", CONDI$)
          TICAM = VALOBADAMDB("CAJABANC", "Tipo_Cam", CONDI$)
          '
          DOCUAP% = 0: docuor% = 0
          For KKI% = 1 To NUDOCU%
            If DOCULAR$(KKI%) = RECI$ And DOCUAP% < 1 Then
                DOCUAP% = KKI%
            End If
            If DOCULAR$(KKI%) = FACU$ And docuor% < 1 Then
                If IMPO#(KKI%) > 0 Then docuor% = KKI%
            End If
          Next KKI%
          '
          If MONE% = 2 Then
            A = Ob
          End If
          If DOCUAP% > 0 And docuor% > 0 Then
            If IAPLI# > IMPO#(docuor%) Then IAPLI# = IMPO#(docuor%)
            If IAPLI# > (-1) * IMPO#(DOCUAP%) Then
               IAPLI# = (-1) * IMPO#(DOCUAP%)   ' QUEDA POR INVESTIGAR SI
            End If
            ' ESTA CONDICION CORRECTA - EPB 11/10/08
            If MONE% = 2 Then
            A = B
            End If
            IMPO#(docuor%) = IMPO#(docuor%) - IAPLI#
            IMPO#(DOCUAP%) = IMPO#(DOCUAP%) + IAPLI#
            If (OPCI$ = "NC" Or OPCI$ = "AS") And MONE% > 1 Then
                If IAPLI# > 0 Then
                    A = B
                    IAPLI# = (-1) * IAPLI#
                Else
                    c = d
                End If
                If TICAM > 0.05 Then SALDOLA#(DOCUAP%) = SALDOLA#(DOCUAP%) - (IAPLI# / TICAM)
            End If
          End If
          '
'          IRECIMA# = 0: IFACMAX# = 0
'          For KKI% = 1 To NUDOCU%
'             If DOCULAR$(KKI%) = RECI$ Then
'               If IMPO#(KKI%) < 0 Then
'                 IRECIMA# = IRECIMA# - IMPO#(KKI%)
'               End If
'             End If
'             If DOCULAR$(KKI%) = FACU$ Then
'               If IMPO#(KKI%) > 0 Then
'                 IFACMAX# = IFACMAX# + IMPO#(KKI%)
'               End If
'             End If
'          Next KKI%
'          '
'          If IAPLI# > IRECIMA# Then IAPLI# = IRECIMA#
'          If IAPLI# > IFACMAX# Then IAPLI# = IFACMAX#
'          '
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
'                If DOCULAR$(KKI%) = FACU$ Then
'                   If IMPO#(KKI%) > 0 Then
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
            If MOEMI%(KKI%) <= 2 Then
                IMPO#(KKI%) = 0
                IMPO#(KKJ%) = 0
            End If
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
            If DOCULAR$(JJ%) = NUDOCX$ Then
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
    If COMALTER%(TTXX$ + "'\   \Desea Grabar la Nueva Composición de la Deuda ?\\Cancelar\Grabar") <> 2 Then
       GoTo 999
    End If
    '
24  Screen.MousePointer = 11
    YAPRIRE% = 0
    ABUS$ = "NuClien = " & NCX% & " "
    With SaDeuPen
25    If YAPRIRE% = 0 Then
          .FindFirst ABUS$
          YAPRIRE% = 1
        Else
          .FindNext ABUS$
      End If
      If .NoMatch = False Then
        .Delete
        GoTo 25
      End If
    End With
    '
    HOII% = HOY(HO$)
    FIN& = NUDOCU%
    LU$ = LUDAT$
    Call APERBASDAT("CABAN")
    NOCLOSE% = 1
    For U& = 1 To NUDOCU%
      Call TRACE(20, U&, FIN&)
      If Abs(IMPO#(U&)) >= 0.005 Or (Abs(SALDOLA#(U&)) >= 0.005 And MOEMI%(U&) <= 2) Then
        IMORIGDOL# = 0: MONEMIOR% = 0
        Debug.Print DOCULAR$(U&)
        If MODOMONEI% = 2 Then
          ' BUSCAR MONEDA E IMPORTE ORIGINAL DEL DOCUMENTO
          Call APERBASDAT("CABAN")
          CajaBanc.FindFirst "CodiCom_Lar = '" & DOCULAR$(U&) & "' "
          If CajaBanc.NoMatch = False Then
             On Error Resume Next
             IMORIGDOL# = CajaBanc!Impo_Dola
             MONEMIOR% = CajaBanc!Mone_Emis Mod 6
             TICAMO = CajaBanc!Tipo_Cam
             VALACOM1# = CajaBanc!ValAbsComp
             On Error GoTo 0
             If MONEMIOR% > 1 Then
               If Abs(IMORIGDOL#) < 0.005 Then
                 If TICAMO > 0.1 Then
                   IMORIGDOL# = ROUND(VALACOM1# / TICAMO)
                   CajaBanc.Edit
                   CajaBanc!Impo_Dola = IMORIGDOL#
                   CajaBanc.Update
                 End If
               End If
             End If
                   
          End If
        End If
        '
        CoComCor$ = DOCUCOR$(U&)
        CoComLar$ = DOCULAR$(U&)
        CodiEmp% = CODIEM%(U&)
        CodiSuc% = CODISU%(U&)
        NumeClie% = NCX%
        NumeProv% = 0
        FechEmi = CVDAT(FFDOC%(U&))
        IBruTot# = IORI#(U&)
        IDolTot# = IMORIGDOL#
        '
        If DOCULAR$(U&) <> DOCULARANT$ Then NuorVen% = 0
        NuorVen% = NuorVen% + 1
        '
        If IMPO#(U&) < 0 Then FVENC%(U&) = FFDOC%(U&)
        FVECIX% = FVENC%(U&)
        If FVECIX% < 1 Then
           COPAGI% = CPAGCLI%(NCX%)
           FEDOCU% = FFDOC%(U&)
           Call BUCONPA(LU$, COPAGI%, CoPag$, FREF%, PLA$)
           AAAA = FEDOCU%
           BBBB = FEDOCU%
           PLAPA% = 0: If Len(PLA$) > 0 Then PLAPA% = Asc(Left$(PLA$, 1))
           While DIENTRE%(AAAA, BBBB) < PLAPA%
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
   With SaDeuPen
     .AddNew
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !NuClien = NumeClie%
      !NuProve = NumeProv%
      !CodCliPro = Left$(CODICLI$(NumeClie% - NumeProv%), 16)
      !RaSoCliPro = Left$(RASOCLI$(NumeClie% - NumeProv%), 48)
      !CodiCom_Cor = Left$(CoComCor$, 12)
      !CODICOM_LAR = Left$(CoComLar$, 24)
      !FechEmisi = FechEmi
        On Error Resume Next
      !Mone_Emis = MONEMIOR%
        On Error GoTo 0
        SIGX% = 1
        If Abs(ImVenc#) > 0 Then SIGX% = Sgn(ImVenc#)
      !IBruComp = IBruTot# * SIGX%
      !IIvaComp = IIvaTot#
      !Ibru_Dola = IDolTot# * SIGX%
      !OrVencim = NuorVen%
      !FEVENCIM = FeVenc
      If ImVenc > 0 Then
          !SaldDebe = ImVenc#
          !SaldAcre = 0
        Else
          !SaldAcre = Abs(ImVenc#)
          !SaldDebe = 0
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

Private Sub Text20_Change()
  NSUC& = XVALO(Text20)
  Label27 = "Sin Clasificar"
  If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
    If NSUC& > 0 Then
      If NSUC& <= ULTREG&("!SUCENTRE") Then
        XX$ = REGLEIDO$("!SUCENTRE", NSUC&)
        Label27 = TRIM$(Mid$(XX$, 3, 62))
      End If
    End If
  End If
  Call CARESCUE
End Sub

Private Sub Text20_DblClick()
   Call Command25_Click
End Sub

Private Sub Command25_Click()
   If EXISTE%(LUDAT$ + "SUCENTRE.RFS") > 0 Then
      Call CommSel("!SUCENTRE", Command25, Text20)
   End If
End Sub


