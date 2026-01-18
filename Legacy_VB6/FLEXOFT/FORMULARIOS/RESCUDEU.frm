VERSION 5.00
Begin VB.Form RESCUDEU 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes de Cuenta Corriente de Deudores"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7245
      TabIndex        =   35
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
         Height          =   300
         Left            =   735
         TabIndex        =   42
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
         TabIndex        =   40
         Text            =   " "
         ToolTipText     =   "Doble Click para Listar Detalle"
         Top             =   1170
         Width           =   1590
      End
      Begin VB.TextBox Text11 
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
         TabIndex        =   38
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
         TabIndex        =   36
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
         TabIndex        =   43
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
         TabIndex        =   41
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
         TabIndex        =   39
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
         TabIndex        =   37
         Top             =   825
         Width           =   645
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   9975
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   33
      Top             =   5040
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Re-apli"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9975
      Picture         =   "RESCUDEU.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Re-Aplicacion de Creditos "
      Top             =   3990
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9975
      Picture         =   "RESCUDEU.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Resúmenes de Cuenta Corriente por Impresora"
      Top             =   2415
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9975
      Picture         =   "RESCUDEU.frx":07B4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   2800
      Left            =   105
      ScaleHeight     =   2745
      ScaleWidth      =   9615
      TabIndex        =   23
      Top             =   5040
      Width           =   9675
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFE0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   5630
         Left            =   0
         TabIndex        =   26
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
            Height          =   2085
            Left            =   -30
            TabIndex        =   27
            Top             =   0
            Width           =   9675
         End
         Begin VB.Line Line19 
            X1              =   7350
            X2              =   7350
            Y1              =   2100
            Y2              =   2415
         End
         Begin VB.Line Line18 
            X1              =   6195
            X2              =   6195
            Y1              =   2100
            Y2              =   2415
         End
         Begin VB.Line Line17 
            X1              =   4200
            X2              =   4200
            Y1              =   2100
            Y2              =   2415
         End
         Begin VB.Line Line16 
            X1              =   3045
            X2              =   3045
            Y1              =   2100
            Y2              =   2415
         End
         Begin VB.Line Line15 
            X1              =   0
            X2              =   9660
            Y1              =   2420
            Y2              =   2420
         End
         Begin VB.Line Line14 
            BorderWidth     =   2
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
            TabIndex        =   30
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
            TabIndex        =   29
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
            TabIndex        =   28
            Top             =   2130
            Width           =   1695
         End
         Begin VB.Line Line12 
            X1              =   2205
            X2              =   2205
            Y1              =   1575
            Y2              =   3360
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFE0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   330
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   10515
         Begin VB.Line Line11 
            X1              =   6195
            X2              =   6195
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line10 
            X1              =   4200
            X2              =   4200
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line9 
            X1              =   3045
            X2              =   3045
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line8 
            X1              =   2205
            X2              =   2205
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   " Documento                     Nro      Emision                    Importe      Vencim.                      Saldo              "
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
            TabIndex        =   25
            Top             =   50
            Width           =   10410
         End
         Begin VB.Line Line2 
            X1              =   7350
            X2              =   7350
            Y1              =   0
            Y2              =   1050
         End
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   9870
      Sorted          =   -1  'True
      TabIndex        =   22
      Top             =   5460
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   2430
      Left            =   105
      ScaleHeight     =   2370
      ScaleWidth      =   9615
      TabIndex        =   17
      Top             =   2415
      Width           =   9675
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   330
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   10515
         Begin VB.Line Line3 
            X1              =   7560
            X2              =   7560
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "   Fecha       Documento                      Nro                Debe                   Haber                  Saldo              "
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
            TabIndex        =   19
            Top             =   50
            Width           =   10410
         End
         Begin VB.Line Line1 
            X1              =   1155
            X2              =   1155
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line4 
            X1              =   3360
            X2              =   3360
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line5 
            X1              =   4200
            X2              =   4200
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line6 
            X1              =   5775
            X2              =   5775
            Y1              =   0
            Y2              =   1050
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0F0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   5580
         Left            =   0
         TabIndex        =   20
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
            Left            =   -30
            TabIndex        =   21
            Top             =   1365
            Width           =   9670
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
      Left            =   105
      TabIndex        =   5
      Top             =   210
      Width           =   7050
      Begin VB.TextBox Text9 
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
         TabIndex        =   14
         Text            =   " "
         Top             =   1560
         Width           =   3690
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Text            =   " "
         Top             =   800
         Width           =   1695
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
         TabIndex        =   16
         Top             =   1260
         Width           =   960
      End
      Begin VB.Label label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto:"
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
         Left            =   105
         TabIndex        =   15
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   480
         Width           =   960
      End
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
      Height          =   855
      Left            =   9975
      Picture         =   "RESCUDEU.frx":0ABE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1365
      Width           =   855
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
      Height          =   855
      Left            =   9975
      Picture         =   "RESCUDEU.frx":0DC8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Finaliza la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Line Line24 
      X1              =   9765
      X2              =   10920
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line23 
      X1              =   105
      X2              =   9765
      Y1              =   4840
      Y2              =   4840
   End
   Begin VB.Line Line21 
      X1              =   9775
      X2              =   9775
      Y1              =   5040
      Y2              =   7820
   End
   Begin VB.Line Line20 
      X1              =   9775
      X2              =   9775
      Y1              =   2415
      Y2              =   4830
   End
   Begin VB.Line Line13 
      X1              =   4830
      X2              =   6090
      Y1              =   3570
      Y2              =   3990
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   9390
      Picture         =   "RESCUDEU.frx":0F12
      Top             =   7290
      Width           =   2010
   End
End
Attribute VB_Name = "RESCUDEU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
    If PROLLAMA$ = "MODICLI" Then
       PROLLAMA$ = ""
       MODICLI.Show
       Hide
       Exit Sub
    End If
    '
    Call CIERRARCH("*.*")
    For K% = 0 To Forms.Count - 1
        ASSS$ = Forms(K%).Name
        If Forms(K%).Name = PROLLAMA$ Then
            On Error Resume Next
            Forms(K%).Show
            On Error GoTo 0
            PROLLAMA$ = ""
            Hide
            Exit Sub
          ElseIf Forms(K%).Name = "COGEVENT" Then
            On Error Resume Next
            COGEVENT.Show
            On Error GoTo 0
            Hide
            Exit Sub
          ElseIf Forms(K%).Name = "AMACLI1" Then
            AMACLI1.Show
            Hide
            Exit Sub
          ElseIf Forms(K%).Name = "FORMFAC" Then
            FORMFAC.Show
            Hide
            Exit Sub
          ElseIf Forms(K%).Name = "FORMRAC" Then
            FORMRAC.Show
            Hide
            Exit Sub
          ElseIf Forms(K%).Name = "FORFACNEU" Then
            FORFACNEU.Show
            Hide
            Exit Sub
          ElseIf Forms(K%).Name = "FOREMIFAC" Then
            FOREMIFAC.Show
            Hide
            Exit Sub
        End If
    Next K%
End Sub


Private Sub Command2_Click()
    '
    Call SELECHO("DEUDOR1")
    VSELE$ = TRIM$(VALACT1$("DEUDOR1"))
    If VSELE$ <> "" Then
       NC% = Val(VSELE$)
       Text1.TEXT = TRIM$(Str$(NC%))
       'Call CARESCUE
    End If
    '
End Sub

Private Sub Command3_Click()
    '
    Screen.MousePointer = 11
    '
    If ULTREG&("DEUDOR1") > 0 Then
        X$ = REGLEIDO$("DEUDOR1", 1)      ' fuerza la
        Call GRAREG("DEUDOR1", X$, 1)     ' reindexacion
        Call CIERRARCH("DEUDOR1")         ' del archivo
    End If
    '
    Screen.MousePointer = 11
    '
    If Int(Val(Text1.TEXT)) > 0 Then
        If Int(Val(Text1.TEXT)) <= 32767 Then
            NCMIN% = Int(Val(Text1.TEXT))
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
    RESCUECLI.Text10.TEXT = TRIM$(Str$(NCMIN%))
    RESCUECLI.Text9.TEXT = RASOCLI$(NCMIN%)
    RESCUECLI.Text8.TEXT = TRIM$(Str$(NCMAX%))
    RESCUECLI.Text7.TEXT = RASOCLI$(NCMAX%)
    '
    HOO% = HOY(HO$)
    RESCUECLI.Text6.TEXT = "01" + Mid$(HO$, 3)
    RESCUECLI.Text5.TEXT = HO$
    '
    RESCUECLI.Show 1
    '
End Sub

Private Sub Command4_Click()
    '
    NC% = Val(Text1.TEXT)
    If NC% < 1 Or ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Text1.SetFocus
       Exit Sub
    End If
    '
    If COMALTER%("El Proceso Invocado Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Créditos y Cobranzas\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar") <> 2 Then
       Exit Sub
    End If
    '
    Call REAPLICRE(NC%)
    Text1.TEXT = TRIM$(Str$(NC%))
    List1.Clear
    List2.Clear
    Call CARESCUE
    DoEvents
    '
End Sub

Private Sub Command5_Click()
    NC% = Val(TRIM$(Text1.TEXT))
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
    If Val(Text1.TEXT) > 0 Then
        NC% = Val(Text1.TEXT)
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
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    On Error Resume Next
    Hide
    On Error GoTo 0
    For K% = 0 To Forms.Count - 1
        If Forms(K%).Name = PROLLAMA$ Then
            Forms(K%).Show
            PROLLAMA$ = ""
            Exit Sub
          ElseIf Forms(K%).Name = "COGEVENT" Then
            On Error Resume Next
            COGEVENT.Show
            On Error GoTo 0
            Exit Sub
          ElseIf Forms(K%).Name = "AMACLI1" Then
            AMACLI1.Show
            Exit Sub
          ElseIf Forms(K%).Name = "FORMFAC" Then
            FORMFAC.Show
            Exit Sub
          ElseIf Forms(K%).Name = "FORMRAC" Then
            FORMRAC.Show
            Exit Sub
          ElseIf Forms(K%).Name = "FORFACNEU" Then
            FORFACNEU.Show
            Exit Sub
          ElseIf Forms(K%).Name = "FOREMIFAC" Then
            FOREMIFAC.Show
            Exit Sub
        End If
    Next K%
End Sub


Private Sub List1_DblClick()
   '
   RELIS$ = List1.TEXT
   TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 16))
   If TIDOCUM$ = "Ajuste de Saldo" Then TIDOCUM$ = "Ajuste Saldo Cliente"
   NUDOCU$ = " " + TRIM$(Mid$(RELIS$, 27, 8)) + " "
   If InStr(UCase$(TIDOCUM$), "FACTURA") > 0 Or InStr(UCase$(TIDOCUM$), "NOTA DE") > 0 Then
     NUDOCU$ = TRIM$(NUDOCU$)
     While Len(NUDOCU$) < 8: NUDOCU$ = "0" + NUDOCU$: Wend
   End If
   RACLI$ = TRIM$(Left$(RASOCLI$(Val(Text1.TEXT)), 8))
   Screen.MousePointer = 11
   For UI& = ULTREG&("PDOCLST") To 1 Step -1
     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
     If InStr(XX$, TIDOCUM$) > 4 Then
       If InStr(XX$, NUDOCU$) > 4 Then
         If InStr(XX$, RACLI$) > 16 Then
           DOCUNU& = CVS(Left$(XX$, 4))
           GoTo 5
         End If
       End If
     End If
   Next UI&
   Screen.MousePointer = 1
   Call COMUNI("Imposible Abrir Documento Desde esta Aplicación")
   Exit Sub
   '
   '16/03/07 (OT-07-0113)
5  Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
   '16/03/07 (FIN)
'5  For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Cls
'     EDITFORM.Picture1(KKI%).DrawWidth = 3
'     EDITFORM.Picture1(KKI%).DrawStyle = 0
'     EDITFORM.Picture1(KKI%).FillStyle = 1
'   Next KKI%
'   '
'   REDOCU$ = REGACT$("PDOCLST")
'   URE& = ULTREG&("PDOCSPL")
'   LI& = 0: LS& = URE&
'6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'   DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'   GoTo 8
'   '
'7  Screen.MousePointer = 1
'   Call MENSERR(24, "Documento no Registrado")
'   Exit Sub
'   '
'8  TPSP$ = ""
'   Screen.MousePointer = 11
'   While L& > 1
'     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'     If DCC& < DOCUNU& Then GoTo 9
'     L& = L& - 1
'   Wend
'   '
'9  CAPAGINAS% = 0
'   For KKL& = L& To URE&
'     TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'     End If
'   Next KKL&
'   '
'95 SPOOLSTRING$ = TPSP$
'   For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Cls
'     EDITFORM.Picture1(KKI%).Top = 0
'     EDITFORM.Picture1(KKI%).Refresh
'     EDITFORM.Picture1(KKI%).Height = 7170
'     EDITFORM.Picture1(KKI%).Width = 11700
'     EDITFORM.Picture1(KKI%).DrawWidth = 3
'     EDITFORM.Picture1(KKI%).DrawStyle = 0
'     EDITFORM.Picture1(KKI%).FillStyle = 1
'   Next KKI%
'   ALTUPAGINA = 7170
'   TOPEPAGINA = 0
'   PAGINACTIVA% = 0
'   CAPAGINAS% = 0
'   '
'10 PPXX% = InStr(TPSP$, Chr$(255))
'   If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 10
'      '
'   End If
'   '
'   Screen.MousePointer = 1
'   PAGINACTIVA% = 0
'   For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Visible = False
'     If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
'   Next KKI%
'   EDITFORM.Picture1(PAGINACTIVA%).Refresh
'   EDITFORM.VScroll1.Min = 0
'   EDITFORM.VScroll1.Value = 0
'   EDITFORM.VScroll1.Max = 1
'   If ALTUPAGINA > EDITFORM.Frame1.Height Then
'      EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'   End If
'   EDITFORM.Command3.Enabled = False
'   EDITFORM.Command5.Enabled = False
'   EDITFORM.Show
   '
End Sub

Private Sub Text1_Change()
   List1.Clear
   List2.Clear
   If Abs(Val(Text1)) < 32767 Then
     NCX% = Val(Text1)
     If NCX% > 0 Then
       If ECOARCH$("DEUDOR1", MKI$(NCX%)) <> "" Then
         NC% = NCX%
         Call CARESCUE
       End If
     End If
   End If
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("DEUDOR1")
    VSELE$ = TRIM$(VALACT1$("DEUDOR1"))
    If VSELE$ <> "" Then
       NC% = Val(VSELE$)
       Text1.TEXT = TRIM$(Str$(NC%))
       'Call CARESCUE
    End If
    '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       'If Val(Text1.TEXT) > 0 Then
       '   If Val(Text1.TEXT) <= 32767 Then
       '      NC% = Val(Text1.TEXT)
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


Private Sub Text6_DblClick()
   Text6.SelStart = 0
   Text6.SelLength = 0
   Screen.MousePointer = 11
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NC% = Val(Text1.TEXT)
       If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
         Call COPYSTRU("CHECAR", "LICHECLI")
         Call HEADERS("LICHECLI", "")
         Call HEADERS("LICHECLI", "Cliente: " + RASOCLI$(NC%))
         '
         JJL& = 0
         RFF$ = RECFILT$("CHECAR", 3, MKI$(NC%))
         HOII% = HOY(HOS$)
         For UI& = 1 To Len(RFF$) Step 4
            RECHE& = CVS(Mid$(RFF$, UI&, 4))
            If RECHE& > 0 Then
               If RECHE& <= ULTREG&("CHECAR") Then
                  XX$ = REGLEIDO$("CHECAR", RECHE&)
                  If CVI(Mid$(XX$, 3, 2)) = NC% Then
                     If CVI(Mid$(XX$, 41, 2)) >= HOII% Then
                        JJL& = JJL& + 1
                        Call GRAREG("LICHECLI", XX$, JJL&)
                     End If
                  End If
               End If
            End If
         Next UI&
         Call SETULTREG("LICHECLI", JJL&)
         Call CIERRARCH("LICHECLI")
         Call CIERRARCH("CHECAR")
         '
         Call CONECRUN("*LICHECLI", "Listado de Cheques por Cliente/WAIT")
         '
       End If
     End If
   End If
   '
   Screen.MousePointer = 1
   '
End Sub

Sub CARESCUE()
    '
    If ATENTI% = 1 Then
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    NC% = Val(RESCUDEU.Text1.TEXT)
    If NC% > 0 Then
        RESCUDEU.List1.Clear
        RESCUDEU.List1.Refresh
        RESCUDEU.List2.Clear
        RESCUDEU.List2.Refresh
        RESCUDEU.Text2.TEXT = RASOCLI$(NC%)
        RESCUDEU.Text3.TEXT = DOMICLI$(NC%)
        RESCUDEU.Text4.TEXT = CPOSCLI$(NC%)
        RESCUDEU.Text5.TEXT = LOCACLI$(NC%)
        RESCUDEU.Text7.TEXT = TEL1CLI$(NC%)
        RESCUDEU.Text10.TEXT = CSTRING$(MKD$(SALDCLI#(NC%)), 4, 11, 2, 2)
        RESCUDEU.Text11.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 4, 11, 2, 2)
        RESCUDEU.Refresh
        RESCUDEU.Text6.TEXT = CSTRING$(MKD$(CHEPEACRE#(NC%)), 4, 11, 2, 2)
        DISPONI = Val(RESCUDEU.Text11.TEXT) - Val(RESCUDEU.Text10.TEXT) - Val(RESCUDEU.Text6.TEXT)
        RESCUDEU.Text12.ForeColor = &H80000008
        If DISPONI < 0 Then RESCUDEU.Text12.ForeColor = &HFF&
        RESCUDEU.Text12.TEXT = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
        RESCUDEU.Refresh
    End If
    '
    '
100 MAXDIAC& = ULTREG&("COBRAN")
    MAXE& = ULTREG&("FACTUR")
    '
    RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
    RFC$ = RECFILT$("COBRAN", 1, MKI$(NC%))
    '
    RESCUDEU.LINDICE.Clear
    For U& = 1 To Len(RFF$) Step 4
        REG& = CVS(Mid$(RFF$, U&, 4))
        X$ = REGLEIDO$("FACTUR", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("FACTUR", 1)
            Call GRAREG("FACTUR", X$, 1)
            Call CIERRARCH("FACTUR")
            GoTo 100
        End If
        FF% = CVI(Mid$(X$, 7, 2))
        Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2)
        RESCUDEU.LINDICE.AddItem Z$
    Next U&
    '
    For U& = 1 To Len(RFC$) Step 4
        REG& = CVS(Mid$(RFC$, U&, 4))
        X$ = REGLEIDO$("COBRAN", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("COBRAN", 1)
            Call GRAREG("COBRAN", X$, 1)
            Call CIERRARCH("COBRAN")
            GoTo 100
        End If
        FF% = CVI(Mid$(X$, 7, 2))
        Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$((-1) * REG&), 4, 8, 0, 2)
        RESCUDEU.LINDICE.AddItem Z$
    Next U&
    '
    SALDO# = 0: SARRAS# = 0
    For U& = 1 To RESCUDEU.LINDICE.ListCount
        RESCUDEU.LINDICE.ListIndex = U& - 1
        REG& = Val(Mid$(RESCUDEU.LINDICE.TEXT, 7, 8))
        If REG& > 0 Then
            If REG& > MAXE& Then GoTo 990
            X$ = REGLEIDO$("FACTUR", REG&)
            TI% = Asc(Mid$(X$, 13, 1))
            TIPOS$ = "Factura": If TI% = 0 Then TIPOS$ = "Ajuste de Saldo"
                               If TI% = 1 Then TIPOS$ = "Venta Contado"
                               If TI% = 3 Then TIPOS$ = "Nota de Debito"
                               If TI% = 5 Then TIPOS$ = "Descuento Caja"
                               If TI% = 6 Then TIPOS$ = "Nota de Credito"
                               If TI% = 7 Then TIPOS$ = "Ajuste de Saldo"
                               '
            NROL& = CVS(Mid$(X$, 3, 4))
            If NROL& < 1 Then
               If TI% = 0 Or TI% = 7 Then
                  Call REPLA(X$, MKS$(REG&), 3, 4)
                  Call GRAREG("FACTUR", X$, REG&)
               End If
            End If
            '
            SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
            IMPO# = CVD(Mid$(X$, 15, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 23, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 31, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 49, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 57, 8))
            IMPO# = SG% * IMPO#
            GoTo 890
            '
         ElseIf REG& < 0 Then
            REG1& = Abs(REG&)
            If REG1& > MAXDIAC& Then GoTo 990
            X$ = REGLEIDO$("COBRAN", REG1&)
            TI% = 4
            TIPOS$ = "Recibo"
            SG% = (-1)
            IMPO# = (-1) * CVD(Mid$(X$, 15, 8))
            GoTo 890
        End If
        GoTo 990
        '
890     SALDO# = SALDO# + (Int(100 * IMPO# + 0.5)) / 100
        FEX = CVI(Mid$(X$, 7, 2))
        NROL& = CVS(Mid$(X$, 3, 4))
        Z$ = Space$(80)
        '
        Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
        Call REPLA(Z$, TIPOS$, 11, 16)
        Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 27, 8)
        IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 13, 2, 2)
        SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
        If IMPO# > 0 Then
            Call REPLA(Z$, IMPTEX$, 35, 13)
          ElseIf IMPO# < 0 Then
            Call REPLA(Z$, IMPTEX$, 50, 13)
        End If
        Call REPLA(Z$, SALTEX$, 63, 14)
        RESCUDEU.List1.AddItem Z$
        '
990 Next U&
    '
991 SARRAS# = SALDO#
    RESCUDEU.List1.ListIndex = RESCUDEU.List1.ListCount - 1
    RESCUDEU.List1.Refresh
    '
    MAXDEU& = ULTREG&("SVD202")
    RFD$ = RECFILT$("SVD202", 1, MKI$(NC%))
    '
    RESCUDEU.LINDICE.Clear
    For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        X$ = REGLEIDO$("SVD202", REG&)
        NCII% = CVI(Left$(X$, 2))
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("SVD202", 1)
            Call GRAREG("SVD202", X$, 1)
            Call CIERRARCH("SVD202")
            GoTo 100
        End If
        IMPO# = CVD(Mid$(X$, 21, 8))
        SALDO# = CVD(Mid$(X$, 11, 8))
        If Abs(SALDO#) < 0.005 Then
          If Abs(SALDO#) > 0.0005 Then
            SALDO# = 0
            Call REPLA(X$, MKD$(0), 11, 8)
            Call GRAREG("SVD202", X$, REG&)
          End If
        End If
        '
        If Abs(SALDO#) >= 0.005 Then
          If Sgn(IMPO#) <> Sgn(SALDO#) Or Abs(SALDO#) > Abs(IMPO#) + 0.005 Then
            ABUS$ = Left$(X$, 8)
            For UU& = 1 To Len(RFD$) Step 4
                REG& = CVS(Mid$(RFD$, UU&, 4))
                X$ = REGLEIDO$("SVD202", REG&)
                If Left$(X$, 8) = ABUS$ Then
                    Call GRAREG("SVD202", String$(32, 0), REG&)
                End If
            Next UU&
            GoTo 100
          End If
        End If
        '
        If Abs(SALDO#) > 0 Then
            FF% = CVI(Mid$(X$, 9, 2))
            Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2)
            RESCUDEU.LINDICE.AddItem Z$
        End If
    Next U&
    '
    SUMADEU# = 0: SUMAFAC# = 0
    For U& = 1 To RESCUDEU.LINDICE.ListCount
        RESCUDEU.LINDICE.ListIndex = U& - 1
        REG& = Val(Mid$(RESCUDEU.LINDICE.TEXT, 7, 8))
        If REG& > 0 Then
            If REG& <= MAXDEU& Then
                X$ = REGLEIDO$("SVD202", REG&)
                TI% = Asc(Mid$(X$, 3, 1))
                TIPOS$ = "Factura": If TI% = 0 Then TIPOS$ = "Ajuste de Saldo"
                                   If TI% = 1 Then TIPOS$ = "Venta Contado"
                                   If TI% = 3 Then TIPOS$ = "Nota de Debito"
                                   If TI% = 4 Then TIPOS$ = "Anticipo"
                                   If TI% = 5 Then TIPOS$ = "Descuento Caja"
                                   If TI% = 6 Then TIPOS$ = "Nota de Credito"
                                   If TI% = 7 Then TIPOS$ = "Ajuste de Saldo"
                                   '
                SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
                IMPO# = CVD(Mid$(X$, 21, 8))
                If TI% <> 0 Then IMPO# = Abs(IMPO#)
                IMPO# = SG% * IMPO#
                SUMAFAC# = SUMAFAC# + IMPO#
                '
                SALDO# = CVD(Mid$(X$, 11, 8))
                If TI% <> 0 Then SALDO# = Abs(SALDO#)
                SALDO# = SG% * SALDO#
                SUMADEU# = SUMADEU# + SALDO#
                '
                NROL& = CVS(Mid$(X$, 5, 4))
                FEMI = CVI(Mid$(X$, 19, 2))
                FEVE = CVI(Mid$(X$, 9, 2))
                '
1234            FEVEX$ = FECHATEX$(FEVE)
                If Val(Mid$(FEVEX$, 4, 2)) = 2 Then
                  If Val(Mid$(FEVEX$, 1, 2)) > 28 Then
                    FEVE = DIPOST(FEVE)
                    GoTo 1234
                  End If
                End If
                If CVI(Mid$(X$, 9, 2)) <> FEVE Then
                  FEVEI% = FEVE
                  Call REPLA(X$, MKI$(FEVEI%), 9, 2)
                  Call GRAREG("SVD202", X$, REG&)
                End If
                '
                IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 14, 2, 2)
                SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
                '
If Abs(Abs(IMPO#) - Abs(SALDO#)) < 0.005 Then ' NO ESTA CANCELADA
  MONEMX% = Asc(Mid$(X$, 31, 1))
  If MONEMX% > 2 Then
    MONEMX% = 0
    X$ = REGLEIDO$("SVD202", REG&)
    Call REPLA(X$, Chr$(0), 31, 1)
    Call GRAREG("SVD202", X$, REG&)
  End If
  If MONEMX% <= 1 Then
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
      If CVS(Mid$(XXC$, 39, 4)) = NROL& Then
        If Asc(Mid$(XXC$, 37, 1)) = TI% Then
          MONEMX% = Asc(Mid$(XXC$, 31, 1))
          GoTo 15
        End If
      End If
    Next KUL&
15  If MONEMX% <> 2 Then MONEMX% = 1
    X$ = REGLEIDO$("SVD202", REG&)
    Call REPLA(X$, Chr$(MONEMX%), 31, 1)
    Call GRAREG("SVD202", X$, REG&)
  End If
  If MONEMX% = 2 Then
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
      If CVS(Mid$(XXC$, 39, 4)) = NROL& Then
        If Asc(Mid$(XXC$, 37, 1)) = TI% Then
          SALDOPES# = CVD(Mid$(XXC$, 65, 8)) - CVD(Mid$(XXC$, 73, 8))
          If Abs(SALDOPES# - SALDO#) <= 0.005 Then
            SALTEX$ = "U$" + FORMATNUM$(CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8)), "F12.2")
            GoTo 20
          End If
        End If
      End If
    Next KUL&
  End If
End If
20              If TRIM$(SALTEX$) <> "" Then
                  Z$ = Space$(80)
                  Call REPLA(Z$, TIPOS$, 1, 16)
                  Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 17, 8)
                  Call REPLA(Z$, FECHATEX$(FEMI), 27, 8)
                  Call REPLA(Z$, IMPTEX$, 37, 14)
                  Call REPLA(Z$, FECHATEX$(FEVE), 53, 8)
                  Call REPLA(Z$, SALTEX$, 63, 14)
                  '
                  RESCUDEU.List2.AddItem Z$
                End If
                '
            End If
        End If
        '
    Next U&
    '
    RESCUDEU.List2.ListIndex = RESCUDEU.List2.ListCount - 1
    RESCUDEU.List2.Refresh
    RESCUDEU.Label5.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    RESCUDEU.Label5.Refresh
    RESCUDEU.Label6.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    RESCUDEU.Label6.Refresh
    '
    If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
       ATENTI% = 1
       TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta"
       If RESCUDEU.Visible = True Then
           TTXX$ = TTXX$ + " Elegida."
         Else
           TTXX$ = TTXX$ + ":\   \'" + TRIM$(RASOCLI$(NC%)) + "'"
       End If
       TTXX$ = TTXX$ + "\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Créditos y Cobranzas\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
       '
       If COMALTER%(TTXX$) = 2 Then
           RESCUDEU.List1.Refresh
           RESCUDEU.List2.Refresh
           Call REAPLICRE(NC%)
           RESCUDEU.List1.Clear
           RESCUDEU.List2.Clear
           GoTo 100
         Else
           ATENTI% = 0
           Call CIERRARCH("*.*")
           Call FINAL("")
       End If
    End If
    '
    If Abs(SARRAS# - SALDCLI#(NC%)) >= 0.005 Then
        RGXC& = REGICLI&(NC%)
        XC$ = REGLEIDO$("DEUDOR2", RGXC&)
        Call REPLA(XC$, MKD$(SARRAS#), 109, 8)
        Call GRAREG("DEUDOR2", XC$, RGXC&)
        RESCUDEU.Text10.TEXT = CSTRING$(MKD$(SARRAS#), 4, 11, 2, 2)
        RESCUDEU.Refresh
    End If
    Screen.MousePointer = 1
    '
End Sub
'

Sub REAPLICRE(NC%)
    '
    ATENTI% = 1
    Dim DOCU$(8192), IMPO#(8192), REGDO&(8192), FFDOC%(8192), IORI#(8192)
    '
100 Screen.MousePointer = 11
    '
    SUMARRAS# = 0
    If Left$(App.EXEName, 7) = "AMACLI0" Then
         If AMACLIEN.List3.ListCount > 0 Then
           SUMARRAS# = Val(Mid$(AMACLIEN.List3.List(AMACLIEN.List3.ListCount - 1), 63, 14))
         End If
         SUMADEUP# = Val(AMACLIEN.Label28.Caption)
         CADOCPEN% = AMACLIEN.List2.ListCount
      Else
         If RESCUDEU.List1.ListCount > 0 Then
           SUMARRAS# = Val(Mid$(RESCUDEU.List1.List(RESCUDEU.List1.ListCount - 1), 63, 14))
         End If
         SUMADEUP# = Val(RESCUDEU.Label5.Caption)
         CADOCPEN% = RESCUDEU.List2.ListCount
    End If
    '
    MAXDIAC& = ULTREG&("COBRAN")
    MAXE& = ULTREG&("FACTUR")
    '
    RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
    RFC$ = RECFILT$("COBRAN", 1, MKI$(NC%))
    '
    NUDOCU% = 0
    FIN& = Len(RFF$)
    For U& = 1 To Len(RFF$) Step 4
        Call TRACE(20, U&, FIN&)
        REG& = CVS(Mid$(RFF$, U&, 4))
        X$ = REGLEIDO$("FACTUR", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("FACTUR", 1)
            Call GRAREG("FACTUR", X$, 1)
            Call CIERRARCH("FACTUR")
            GoTo 100
        End If
110     NUDO& = CVS(Mid$(X$, 3, 4))
        TI% = Asc(Mid$(X$, 13, 1))
        If NUDO& < 1 Then
           If TI% = 0 Or TI% = 7 Then
              Call REPLA(X$, MKS$(REG&), 3, 4)
              Call GRAREG("FACTUR", X$, REG&)
              GoTo 110
           End If
        End If
        VA% = Asc(Mid$(X$, 14, 1))
        SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
115     IMPO#(NUDOCU%) = CVD(Mid$(X$, 15, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 23, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 31, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 49, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 57, 8))
        IMPO#(NUDOCU%) = SG% * (Int(100 * IMPO#(NUDOCU%) + 0.5)) / 100
        NUSUCUVE% = CVI(Mid$(X$, 47, 2))
        DOCU$(NUDOCU%) = Left$(FORMDOC$(NUDO&, TI%, VA%, NUSUCUVE%), 2) + "-" + TRIM$(Str$(NUDO&))
        REGDO&(NUDOCU%) = REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(X$, 7, 2))
        '
        If TI% > 0 Then
          If TI% < 7 Then
            If CVD(Mid$(X$, 15, 8)) < (-0.005) Then
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 15, 8)))), 15, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 23, 8)))), 23, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 31, 8)))), 31, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 49, 8)))), 49, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 57, 8)))), 57, 8)
              Call GRAREG("FACTUR", X$, REG&)
              GoTo 115
            End If
          End If
        End If
    Next U&
    '
    FIN& = Len(RFC$)
    For U& = 1 To Len(RFC$) Step 4
        Call TRACE(20, U&, FIN&)
        REG& = CVS(Mid$(RFC$, U&, 4))
        X$ = REGLEIDO$("COBRAN", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("COBRAN", 1)
            Call GRAREG("COBRAN", X$, 1)
            Call CIERRARCH("COBRAN")
            GoTo 100
        End If
        NUDO& = CVS(Mid$(X$, 3, 4))
        TI% = 4
        VA% = 1
        SG% = (-1)
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
        IMPO#(NUDOCU%) = CVD(Mid$(X$, 15, 8))
        IMPO#(NUDOCU%) = SG% * (Int(100 * IMPO#(NUDOCU%) + 0.5)) / 100
        DOCU$(NUDOCU%) = "RB-" + TRIM$(Str$(NUDO&))
        REGDO&(NUDOCU%) = (-1) * REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(X$, 7, 2))
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    If NUDOCU% >= 1 Then
       If Left$(DOCU$(1), 2) = "AS" Then
          FIN& = ULTREG&("SALINID")
          For U& = 1 To FIN&
             X$ = REGLEIDO$("SALINID", U&)
             If CVI(Left$(X$, 2)) = NC% Then
               TI% = Asc(Mid$(X$, 3, 1))
               VA% = Asc(Mid$(X$, 4, 1))
               NUSUCUVE% = CVI(Mid$(X$, 29, 2))
               SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
               IPEN# = SG% * CVD(Mid$(X$, 11, 8))
               If Sgn(IPEN#) = Sgn(IMPO#(1)) Then
                 If Abs(IPEN#) <= Abs(IMPO#(1)) Then
                   IMPO#(1) = IMPO#(1) - IPEN#
                   NUDO& = CVS(Mid$(X$, 5, 4))
                   NUDOCU% = NUDOCU% + 1
                   If NUDOCU% > 8192 Then
                      Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
                      Call CIERRARCH("*.*")
                      Call FINAL("")
                   End If
                   IMPO#(NUDOCU%) = IPEN#
                   DOCU$(NUDOCU%) = Left$(FORMDOC$(NUDO&, TI%, VA%, NUSUCUVE%), 2) + "-" + TRIM$(Str$(NUDO&))
                   REGDO&(NUDOCU%) = 900000 + U&
                   FFDOC%(NUDOCU%) = CVI(Mid$(X$, 19, 2))
                 End If
               End If
             End If
          Next U&
       End If
    End If
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
' HASTA AQUI O.K
    KKI% = 0: SUMAPEX# = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          SUMAPEX# = SUMAPEX# + IMPO#(KKI%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    FIN& = ULTREG&("APLICOB")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("APLICOB", U&)
       If CVI(Mid$(X$, 3, 2)) = NC% Then
          RECI$ = Mid$(X$, 5, 18)
          RECI$ = Left$(RECI$, 2) + "-" + TRIM$(Val(Mid$(RECI$, 11)))
          FACU$ = Mid$(X$, 55, 18)
          PINI% = 11: If Mid$(FACU$, 17, 1) = "-" Then PINI% = 9
          FACU$ = Left$(FACU$, 2) + "-" + TRIM$(Val(Mid$(FACU$, PINI%, 8)))
          '
210       IAPLI# = CVD(Mid$(X$, 23, 8))
          IRECIMA# = 0: IFACMAX# = 0
          For KKI% = 1 To NUDOCU%
             If DOCU$(KKI%) = RECI$ Or (RECI$ = "AS-0" And Left$(DOCU$(KKI%), 2) = "AS") Then
               If IMPO#(KKI%) < 0 Then
                 IRECIMA# = IRECIMA# - IMPO#(KKI%)
               End If
             End If
             If DOCU$(KKI%) = FACU$ Then
               If IMPO#(KKI%) > 0 Then
                 IFACMAX# = IFACMAX# + IMPO#(KKI%)
               End If
             End If
          Next KKI%
          '
          If IAPLI# > IRECIMA# Then IAPLI# = IRECIMA#
          If IAPLI# > IFACMAX# Then IAPLI# = IFACMAX#
          '
          IRECIMA# = IAPLI#
          IFACMAX# = IAPLI#
          '
          If IAPLI# > 0 Then
             For KKI% = 1 To NUDOCU%
                If DOCU$(KKI%) = RECI$ Or (RECI$ = "AS-0" And Left$(DOCU$(KKI%), 2) = "AS") Then
                   If IMPO#(KKI%) < 0 Then
                      IAP1# = IRECIMA#
                      If IAP1# > Abs(IMPO#(KKI%)) Then IAP1# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) + IAP1#
                      IRECIMA# = IRECIMA# - IAP1#
                   End If
                End If
                If DOCU$(KKI%) = FACU$ Then
                   If IMPO#(KKI%) > 0 Then
                      IAP2# = IFACMAX#
                      If IAP2# > Abs(IMPO#(KKI%)) Then IAP2# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) - IAP2#
                      IFACMAX# = IFACMAX# - IAP2#
                   End If
                End If
             Next KKI%
          End If
          '
       End If
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Or (FFDOC%(KKI%) = FFDOC%(KKJ%) And Sgn(REGDO&(KKI%)) = Sgn(REGDO&(KKJ%)) > 0 And Abs(REGDO&(KKI%)) > Abs(REGDO&(KKJ%))) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    KKI% = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    For KKI% = 1 To NUDOCU%
       For KKJ% = KKI% + 1 To NUDOCU%
          If Abs(IMPO#(KKI%) + IMPO#(KKJ%)) < 0.005 Then
             IMPO#(KKI%) = 0
             IMPO#(KKJ%) = 0
          End If
       Next KKJ%
    Next KKI%
    '
    KKI% = 0: SUMAPEX# = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          SUMAPEX# = SUMAPEX# + IMPO#(KKJ%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    Screen.MousePointer = 1
    If COMALTER%("Se ha Completado la Reconstrucción\de la Composición de la Deuda de\   \'" + TRIM$(RASOCLI$(NC%)) + "'\   \Desea Grabar la Nueva Composición de la Deuda ?\\Cancelar\Grabar") <> 2 Then
       GoTo 999
    End If
    '
    ' para no enquilombar las cuentas que mas o menos estan
    If Abs(SUMARRAS# - SUMADEUP#) < 0.005 Then
      If NUDOCU% >= CADOCPEN% Then
        GoTo 999
      End If
    End If
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("SVD202")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("SVD202", U&)
       If CVI(Left$(X$, 2)) = NC% Then
          Call REPLA(X$, String$(30, 0), 3, 30)
          Call GRAREG("SVD202", X$, U&)
       End If
    Next U&
    '
    LU$ = LUDAT$
    ARCHIDEU$ = "SVD202"
    RDX% = 1
    FIN& = NUDOCU%
    YABLO% = 0
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       If Abs(IMPO#(KKI%)) >= 0.005 Then
          IORIG# = IMPO#(KKI%)
          If REGDO&(KKI%) > 0 Then
             If REGDO&(KKI%) > 900000 Then
                X$ = REGLEIDO$("SALINID", REGDO&(KKI%) - 900000)
                NRO = CVS(Mid$(X$, 5, 4))
                TIPO% = Asc(Mid$(X$, 3, 1))
                VA% = Asc(Mid$(X$, 4, 1))
                FFAC% = CVI(Mid$(X$, 9, 2))
                COPAGI% = 0
                GoTo 290
             End If
             X$ = REGLEIDO$("FACTUR", REGDO&(KKI%))
             NRO = CVS(Mid$(X$, 3, 4))
             TIPO% = Asc(Mid$(X$, 13, 1))
             VA% = Asc(Mid$(X$, 14, 1))
             FFAC% = CVI(Mid$(X$, 7, 2))
             COPAGI% = CPAGCLI%(NC%)
             If TIPO% < 1 Or TIPO% > 3 Then COPAGI% = 0
             SG% = 1: If TIPO% >= 4 Then If TIPO% <= 6 Then SG% = -1
             IORIG# = CVD(Mid$(X$, 15, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 23, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 31, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 49, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 57, 8))
             IORIG# = SG% * IORIG#
           Else
             X$ = REGLEIDO$("COBRAN", (-1) * REGDO&(KKI%))
             NRO = CVS(Mid$(X$, 3, 4))
             TIPO% = 4
             VA% = 0
             FFAC% = CVI(Mid$(X$, 7, 2))
             COPAGI% = 0
          End If
290       TOTFAC# = IMPO#(KKI%)
          If YABLO% = 0 Then
             YABLO% = 1
             Call BLOQARCH(ARCHIDEU$)
          End If
          ARDEX$ = ARCHIDEU$ + "/" + TRIM$(CSTRING$(MKD$(IORIG#), 4, 12, 2, 2))
          Call GRADEU(LU$, ARDEX$, RDX%, TIPO%, VA%, NRO, NC%, FFAC%, COPAGI%, TOTFAC#)
       End If
    Next KKI%
    '
999 Call CIERRARCH(ARCHIDEU$)
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    ATENTI% = 0
End Sub


