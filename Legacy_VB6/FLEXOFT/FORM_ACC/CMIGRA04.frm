VERSION 5.00
Begin VB.Form CMIGRA04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0E0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presupuestación - Migración de Datos Preexistentes"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11070
   Icon            =   "CMIGRA04.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   3800
      Left            =   8280
      TabIndex        =   19
      Top             =   1600
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3800
      Left            =   2040
      TabIndex        =   18
      Top             =   1600
      Width           =   15
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   3810
      Left            =   200
      ScaleHeight     =   3750
      ScaleWidth      =   9750
      TabIndex        =   43
      Top             =   1600
      Width           =   9810
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         IntegralHeight  =   0   'False
         ItemData        =   "CMIGRA04.frx":0442
         Left            =   -30
         List            =   "CMIGRA04.frx":0444
         TabIndex        =   46
         Top             =   360
         Width           =   9820
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   330
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   10515
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   $"CMIGRA04.frx":0446
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
            Left            =   0
            TabIndex        =   45
            Top             =   80
            Width           =   10995
         End
      End
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
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
      Left            =   8295
      TabIndex        =   41
      Top             =   5670
      Width           =   1650
   End
   Begin VB.TextBox TICAMBIS 
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
      Height          =   300
      Left            =   5365
      MaxLength       =   14
      TabIndex        =   38
      Top             =   6060
      Width           =   1065
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
      Height          =   285
      Left            =   5115
      TabIndex        =   33
      Top             =   840
      Width           =   870
   End
   Begin VB.TextBox Text18 
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
      Height          =   285
      Left            =   5115
      TabIndex        =   32
      Top             =   210
      Width           =   3300
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
      Height          =   285
      Left            =   5115
      TabIndex        =   31
      Top             =   525
      Width           =   3300
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
      Height          =   285
      Left            =   6060
      TabIndex        =   30
      Top             =   840
      Width           =   2340
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Abrir"
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
      Left            =   3675
      Picture         =   "CMIGRA04.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Acceso a Listas de Precios de Venta"
      Top             =   5730
      Width           =   650
   End
   Begin VB.CommandButton CmdLiPre 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3285
      TabIndex        =   27
      Top             =   5730
      Width           =   180
   End
   Begin VB.TextBox TEXT21 
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
      Left            =   2865
      TabIndex        =   26
      Text            =   " "
      Top             =   5730
      Width           =   435
   End
   Begin VB.TextBox TxtLiPre 
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
      Height          =   285
      Left            =   210
      TabIndex        =   25
      Text            =   " "
      Top             =   6090
      Width           =   3270
   End
   Begin VB.TextBox Text38 
      Alignment       =   1  'Right Justify
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
      Left            =   8295
      TabIndex        =   23
      Top             =   6105
      Width           =   1650
   End
   Begin VB.TextBox Text6 
      Height          =   330
      Left            =   6510
      TabIndex        =   22
      Text            =   "Text6"
      Top             =   4200
      Width           =   2220
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
      Height          =   285
      Left            =   9300
      TabIndex        =   20
      Top             =   840
      Width           =   1650
   End
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
      Height          =   285
      Left            =   9820
      TabIndex        =   10
      Top             =   210
      Width           =   1125
   End
   Begin VB.TextBox Text11 
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
      Height          =   285
      Left            =   9975
      TabIndex        =   9
      Top             =   8595
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.TextBox Text10 
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
      Height          =   285
      Left            =   9820
      TabIndex        =   8
      Top             =   525
      Width           =   1125
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   7845
      Left            =   10185
      ScaleHeight     =   7785
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   1260
      Width           =   1275
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
         Left            =   120
         Picture         =   "CMIGRA04.frx":07E4
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   3525
         Width           =   650
      End
      Begin VB.CommandButton Command25 
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
         Height          =   690
         Left            =   105
         Picture         =   "CMIGRA04.frx":0C26
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Imprime Cotización de Cámara"
         Top             =   1860
         Width           =   650
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Copy"
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
         Picture         =   "CMIGRA04.frx":1290
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Copia Interactiva de Rubro"
         Top             =   2535
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   6
         Top             =   4440
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   7
            Top             =   105
            Width           =   12000
         End
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "CMIGRA04.frx":13DA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8100
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Picture         =   "CMIGRA04.frx":16E4
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exit"
         Top             =   105
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
         Height          =   690
         Left            =   105
         Picture         =   "CMIGRA04.frx":182E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Ayuda en Línea"
         Top             =   735
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "CMIGRA04.frx":1B38
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   1410
      End
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
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
      Left            =   6825
      TabIndex        =   40
      Top             =   5775
      Width           =   1380
   End
   Begin VB.Line Line5 
      X1              =   6705
      X2              =   6705
      Y1              =   6540
      Y2              =   5580
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "T-C:"
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
      Left            =   4905
      TabIndex        =   39
      Top             =   6135
      Width           =   435
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5115
      TabIndex        =   37
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Line Line4 
      X1              =   4560
      X2              =   4560
      Y1              =   6540
      Y2              =   5580
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
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
      Left            =   3960
      TabIndex        =   36
      Top             =   305
      Width           =   1110
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Localidad:"
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
      Left            =   4170
      TabIndex        =   35
      Top             =   935
      Width           =   900
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Domicilio:"
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
      Left            =   3990
      TabIndex        =   34
      Top             =   620
      Width           =   1110
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de Precios de Venta:"
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
      Index           =   0
      Left            =   135
      TabIndex        =   28
      Top             =   5805
      Width           =   2640
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Migración:"
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
      Left            =   6825
      TabIndex        =   24
      Top             =   6030
      Width           =   1380
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   11025
      Y1              =   5565
      Y2              =   5565
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11025
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Resp.:"
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
      Left            =   8720
      TabIndex        =   21
      Top             =   935
      Width           =   585
   End
   Begin VB.Label Label7 
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
      Height          =   610
      Left            =   1155
      TabIndex        =   17
      Top             =   525
      Width           =   2910
   End
   Begin VB.Label Label31 
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
      Height          =   285
      Left            =   1155
      TabIndex        =   16
      Top             =   210
      Width           =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. de Presupuesto:"
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
      Left            =   0
      TabIndex        =   15
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Denom.:"
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
      Left            =   -210
      TabIndex        =   14
      Top             =   630
      Width           =   1350
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   9045
      TabIndex        =   13
      Top             =   305
      Width           =   690
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Validez:"
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
      Left            =   8925
      TabIndex        =   12
      Top             =   620
      Width           =   810
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Resp.:"
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
      Left            =   9450
      TabIndex        =   11
      Top             =   8700
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Line Line1 
      X1              =   -420
      X2              =   11580
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "CMIGRA04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%, MODIFIC%
Dim GRABATODO%
Dim ARCHIPRE$



Private Sub Command1_Click()
    '
    If Text6.BackColor = 15132390 Then
        Exit Sub
    End If
    '
    Call SELECHO("MASTER")
    CODD$ = VALACT1$("MASTER")
    If CODD$ = "" Then
        Exit Sub
    End If
    COD$ = CODD$
    Text6.TEXT = COD$
    '
End Sub

Private Sub AyuCli_Click()
  Call Command3_Click
End Sub

Private Sub CmdLiPre_Click()
    '
    If TRIM$(Label31) = "" Then
       On Error Resume Next
       Call MENSERR(24, "Imposible Continuar.\Número de Presupuesto no válido")
       On Error GoTo 0
       Exit Sub
    End If
    '
    ARCHEC$ = "LISTAS": LOCA% = 2
    ARGU$ = CBIN$(TRIM$(Text21.TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       Text21.TEXT = VDEVU$
    End If
    '
End Sub

Private Sub CoInLiMa_Click()
  Call Command22_Click
End Sub



Private Sub Command14_Click()
  Command14.Enabled = False
    OPCFPRE04.Show 1
  Command14.Enabled = True
End Sub

Private Sub Command22_Click()
  Command22.Enabled = False
    Call COACCE
  Command22.Enabled = True
End Sub

Private Sub Command4_Click()
   Call HELPONLINE("AMACLI")
End Sub

Private Sub ConPagCli_Click()
With OPMASCL04
  .Option1(6).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub CuCoCli_Click()
With OPMASCL04
  .Option1(8).Value = True
  Call .Command16_Click
End With
End Sub


Private Sub DeLiCreCli_Click()
    Call AMACLIS04.Command3_Click
End Sub

Private Sub EditCli_Click()
   'Call Command9_Click
End Sub

Private Sub ConGenCli_Click()
  '
  Call Command14_Click
  '
End Sub

Private Sub Exit_Click()
   Call Command3_Click
End Sub
'
Private Sub Form_Load()
  '
  Call CENTRAFORM(Me)
  '
  EPAC$ = TRIM$(EMPREAC$)
  If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
  End If
  Call GRATITFOR(Caption)
  'LEO LA LISTA DE PRECIOS DE ACCESORIOS
  Text21 = CONTROL("LIPREVEN", "ACCESO")
  '
End Sub
'
Private Sub Form_Activate()
  WindowState = 0
  OPENFORMS = DoEvents
  Call CARGODETPRESU
End Sub
'
Private Sub Form_Resize()
  If WindowState = 1 Then
    Call CIERRARCH("*.*")
    Call BAJALDISCO
  End If
End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

'\\\OT-05-0396-WAR-13/06/05///
'Private Sub COMMAND15_CLICK()
'    '
'    If DERACCE%("GRACOTCA", "Grabar Cotización") >= 2 Then
'      NUPRESU1$ = TRIM(Label31)
'      'RUBRO =2
'      For I = 0 To List1.ListCount - 1
'          Y$ = List1.List(I)
'          CODEXT1$ = TRIM(Mid$(Y$, 1, 16))
'          CI% = CODINT(CODEXT1$)
'          PREU1# = Val(Mid$(Y$, 65, 12))
'          CANTI1# = Val(Mid$(Y$, 78, 7))
'          IMPITEM# = Val(Mid$(Y$, 86, 12))
'          Call ABREPRECAM
'          With DetaPresu
'            SQLX$ = " Nro_Presu = '" & NUPRESU1$ & "' "
'            SQLX$ = SQLX$ + " AND  Rubro_Det = 3 "
'            'SQLX$ = SQLX$ + " AND  SubRub_Det = 1 "
'            SQLX$ = SQLX$ + " AND CodINte_Det = " & CI% & " "
'            .FindFirst SQLX$
'10              If .NoMatch = False Then
'              .Edit
'                !PreUni_Det = PREU1#
'                !Cantid_Det = CANTI1#
'                !Impo_Det = IMPITEM#
'              .Update
'              .FindNext SQLX$
'              GoTo 10
'            End If
'          End With
'      Next I
'    End If
'End Sub
'\\\OT-05-0396-WAR-FIN///
Sub CAMGRIS()
   '
   COLOBAK = RGB(230, 230, 230)
   Text6.BackColor = COLOBAK
   Text13.BackColor = COLOBAK
   Text38.BackColor = COLOBAK
   TxtLiPre.BackColor = COLOBAK
   TICAMBIS.BackColor = COLOBAK
   '
End Sub

Private Sub Command25_Click()
   Call PRILISMAT
  ' PRINTFORM
End Sub

Private Sub Command3_Click()
   '
   List1.Clear
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
   '
End Sub

Private Sub Command5_Click()
   '
   Command5.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "LISPRE04"
   Call CONECRUN(ACONEC$, "Listas de Precios de Venta")
99 Command5.Enabled = True
   '
End Sub

Private Sub LiCuInCli_Click()
    Call AMACLIS04.Command5_Click
End Sub

Private Sub Frame3_DBLCLICK()
'    Call List1_DblClick
End Sub

Private Sub Label17_dblClick()
    'MUESTRA AL O/C
     NPX& = Val(Mid$(Label17, 4, 7))
     If NPX& > 0 Then
        Call SHOWOCOM(NPX&)
     End If
    '
End Sub
Sub SHOWOCOM(NPX&)
    '
    TIDOCUM$ = "Nota de Pedido"
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
    Exit Sub
    '
5   NUPIC% = EDITFORM.Picture1.Count
    For KKI% = 0 To NUPIC% - 1
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    Exit Sub
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    NUPIC% = EDITFORM.Picture1.Count
    For KKI% = 0 To NUPIC% - 1
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
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
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    NUPIC% = EDITFORM.Picture1.Count
    For KKI% = 0 To NUPIC% - 1
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command1.Enabled = False
    EDITFORM.Show 1
    '
End Sub

Private Sub Label18_dblClick()
    '
    NPX& = Val(Mid$(Label18, 4, 7))
    If NPX& > 0 Then
     Call SHOWBORE(NPX&)
    End If
    '
End Sub
Sub SHOWBORE(NPX&)
    '
    TIDOCUM$ = "Boleta de Recepción"
    NUDOCU$ = TRIM$(Str$(NPX&))
    NUDOCU$ = " " + NUDOCU$ + " -"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    Exit Sub
    '
5   NUPIC% = EDITFORM.Picture1.Count
    For KKI% = 0 To NUPIC% - 1
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    Exit Sub
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    NUPIC% = EDITFORM.Picture1.Count
    For KKI% = 0 To NUPIC% - 1
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
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
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    NUPIC% = EDITFORM.Picture1.Count
    For KKI% = 0 To NUPIC% - 1
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command1.Enabled = False
    EDITFORM.Show 1
    '
End Sub




Private Sub Label31_Change()
   '
   NUOTR$ = Label31
   '
   With REPRFLEX04
     Label31 = .Text1
     Label7 = .Text3
     Text9 = .Text9
     Text10 = .Text10
     Text8 = .Text11
     Text18 = .Text18
     Text5 = .Text5
     Text7 = .Text7
     Text4 = .Text8
     Label17 = .Label2
     Label17.ForeColor = .Label2.ForeColor
     Label17.FontUnderline = .Label2.FontUnderline
     TICAMBIS = .TICAMBIS
   End With
   '
10 NUPRESU$ = TRIM$(Label31)
   Call ABREPRECAM
   SQLX$ = "SELECT * FROM DETAPRESU "
   SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
   SQLX$ = SQLX$ + "AND RUBRO_DET = 2 "
   SQLX$ = SQLX$ + "ORDER BY CodExte_Det ASC "
   Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
   J& = 0
   With DETPRETMP
     '
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do Until .EOF = True
         Y$ = REGBLAN("!DECOTPIS")
         '
         CODINTE% = !CODINTE_DET
         Call REPLA(Y$, MKI$(CODINTE), 1, 2)
         PREUNI1# = !PreUni_Det
         Call REPLA(Y$, MKD$(PREUNI1#), 3, 8)
         CANTI1# = !Cantid_Det
         Call REPLA(Y$, MKS$(CANTI1#), 11, 4)
         J& = J& + 1
         Call GRAREG("!DECOTPIS", Y$, J&)
       .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   '
   Call SETULTREG("!DECOTPIS", J&)
   Call CIERRARCH("!DECOTPIS")
   '
   List1.Clear
   J& = ULTREG("!DECOTPIS")
   TOTPIS# = 0
   For KKIL& = 1 To J&
      XX$ = REGLEIDO$("!DECOTPIS", KKIL&)
      ITE% = CVI(Mid$(XX$, 1, 2))
      'VALIDACION SI codigo VALE 0
      If ITE% > 0 Then
        CODI1$ = CODEXT(ITE%)
        DESCRI1$ = DESCRIT0(ITE%)
        PREUNI1# = CVD(Mid$(XX$, 3, 8))
        CANTI1# = CVS(Mid$(XX$, 11, 4))
        PREUNI2$ = FORMATNUM(PREUNI1#, "F12.4")
        CANTI2$ = FORMATNUM(CANTI1#, "F7.1")
        Y$ = Space$(256)
        Call REPLA(Y$, CODI1$, 1, 16)
        Call REPLA(Y$, DESCRI1$, 19, 45)
        Call REPLA(Y$, PREUNI2$, 65, 12)
        Call REPLA(Y$, CANTI2$, 78, 7)
        'CALCULO EL TOTAL DEL MATERIAL
        IMPOIT1# = PREUNI1# * CANTI1#
        IMPOIT2$ = FORMATNUM(IMPOIT1#, "F12.2")
        TOTPIS# = TOTPIS# + IMPOIT1#
       Call REPLA(Y$, IMPOIT2$, 86, 12)
        List1.AddItem Y$
      End If
   Next KKIL&
   '
   Text38 = FORMATNUM(TOTPIS#, "F12.2")
   Text13 = FORMATNUM(TOTPIS#, "F12.2")
End Sub

Private Sub Label9_Click()
  Call List1_DBLClick
End Sub

Private Sub LiGeClaCli_Click()
 Call Command25_Click
End Sub

Private Sub SuspOper_Click()
    Call Command3_Click
End Sub

Sub GENEPLACO() 'GENERAR CAPLACO.DAT
    '
    NUOTRAB$ = TRIM$(Text1.TEXT)
    Call ABREORTRA
    SQLX$ = "SELECT * FROM PlanCob "
    SQLX$ = SQLX$ + " WHERE Nro_Otrab= '" & NUOTRAB$ & "' "
    Set PLACOTMP = OTrabajo.OpenRecordset(SQLX$, 4)
    
    J& = 0
    With PLACOTMP
      '
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do Until .EOF = False
          Y$ = REGBLAN("CAPLACO")
          Call REPLA(Y$, !Nro_Otrab, 1, 16)
          Call REPLA(Y$, !NuOrItem, 17, 2)
          Call REPLA(Y$, !Refe_Pcob, 19, 64)
          Call REPLA(Y$, !Impo_Pcob, 83, 8)
          Call REPLA(Y$, !FeVen_Pcob, 91, 2)
          Call REPLA(Y$, !NroFac_Pcob, 93, 16)
          Call REPLA(Y$, !MONEDA, 109, 2)
          Call REPLA(Y$, !Coti_Pcob, 111, 8)
          J& = J& + 1
          Call GRAREG("!CAPLACO", Y$, J&)
        .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
    Call SETULTREG("!CAPLACO", J&)
    Call CIERRARCH("!CAPLACO")
    
End Sub


Private Sub List1_DBLClick()
    'GENERO EL DAT
    If TRIM$(Label31) <> "" Then
     If DERACCE%("GRACOTCA", "Grabar Cotización") >= 2 Then
10    NUPRESU$ = TRIM$(Label31)
      Call ABREPRECAM
      SQLX$ = "SELECT * FROM DETAPRESU "
      SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
      SQLX$ = SQLX$ + "AND RUBRO_DET = 2 "
      SQLX$ = SQLX$ + "ORDER BY CodExte_Det ASC "
      Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
      J& = 0
      With DETPRETMP
       '
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
        On Error GoTo 0
        Do Until .EOF = True
          Y$ = REGBLAN("!DECOTPIS")
          '
          CODINTE% = !CODINTE_DET
          Call REPLA(Y$, MKI$(CODINTE), 1, 2)
          PREUNI1# = !PreUni_Det
          Call REPLA(Y$, MKD$(PREUNI1#), 3, 8)
          CANTI1# = !Cantid_Det
          Call REPLA(Y$, MKS$(CANTI1#), 11, 4)
          J& = J& + 1
          Call GRAREG("!DECOTPIS", Y$, J&)
        .MoveNext
        Loop
       End If
       On Error GoTo 0
      End With
      '
      Call SETULTREG("!DECOTPIS", J&)
      Call CIERRARCH("!DECOTPIS")
    '
    PRF$ = "DECOTLIC"
    ATRI$ = "/NC/ARCHECO(1)=" + ARCHIPRE$
    ATRI$ = ATRI$ + "/TITUPAN=Detalle de Cotización - Accesorios/"
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture2.Left))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture2.Top))
    VTB% = VENTABU%(PRF$ + ATRI$)
    If VTB% >= 0 Then
      'BORRA LOS REGISTROS///
      With DetaPresu
20      .FindFirst "Nro_PRESU = '" & NUPRESU$ & "' AND RUBRO_DET = 2 "
        If .NoMatch = False Then
          .Delete
          GoTo 20
        End If
      End With
      '
      'ACA VA RUTINA DE GRABACION
      List1.Clear
      J& = ULTREG("!DECOTPIS")
      TOTPIS# = 0
      For KKIL& = 1 To J&
         XX$ = REGLEIDO$("!DECOTPIS", KKIL&)
         ITE% = CVI(Mid$(XX$, 1, 2))
         'VALIDACION SI codigo VALE 0
         If ITE% <= 0 Then GoTo 7
          CODI1$ = CODEXT(ITE%)
          DESCRI1$ = DESCRIT0(ITE%)
          PREUNI1# = CVD(Mid$(XX$, 3, 8))
          CANTI1# = 1    ' CVS(Mid$(XX$, 11, 4))
          PREUNI2$ = FORMATNUM(PREUNI1#, "F12.2")
          'CANTI2$ = FORMATNUM(CANTI1#, "F7.1")
          Y$ = Space$(256)
          Call REPLA(Y$, CODI1$, 1, 16)
          Call REPLA(Y$, DESCRI1$, 19, 56)
          'Call REPLA(Y$, PREUNI2$, 79, 12)
          'Call REPLA(Y$, CANTI2$, 78, 7)
          'CALCULO EL TOTAL DEL MATERIAL
          IMPOIT1# = PREUNI1# * CANTI1#
          IMPOIT2$ = FORMATNUM(IMPOIT1#, "F12.2")
          Call REPLA(Y$, IMPOIT2$, 78, 12)
          List1.AddItem Y$
         '
          With DetaPresu
            NUPRESU$ = TRIM(Label31)
            .AddNew
              !Nro_PRESU$ = NUPRESU$
              !rubro_det = 2
              !SUBRUB_DET = 0
              !CODINTE_DET = ITE%
              !CODEXTE_DET = CODEXT$(ITE%)
              !Cantid_Det = CANTI1#
              !PreUni_Det = PREUNI1#
              !MONE_DET = 1 'POR AHORA
              !Impo_Det = PREUNI1# * CANTI1#
            .Update
           End With
           TOTPIS# = TOTPIS# + IMPOIT1#
           Text38 = FORMATNUM(TOTPIS#, "F12.2")
           Text13 = FORMATNUM(TOTPIS#, "F12.2")
7         Next KKIL&
     End If
    End If
  End If
    '
End Sub

Private Sub mnuagrefi_Click()
  Call List1_DBLClick
End Sub

Private Sub mnudel_Click()
    Call List1_DBLClick
End Sub

Private Sub mnuEdit_Click()
  Call List1_DBLClick
End Sub

Private Sub mnuIns_Click()
  Call List1_DBLClick
End Sub

Private Sub NuMat_Click()
  Call List1_DBLClick
End Sub
Private Sub Text1_GotFocus()
  '
  Command3.SetFocus
  '
End Sub


Private Sub Text10_GOTFOCUS()
  '
  Command3.SetFocus
  '
End Sub

Private Sub Text12_GotFocus()
  '
  Command3.SetFocus
  '
End Sub
Private Sub Text13_GOTFOCUS()
  '
  Command3.SetFocus
  '
End Sub
Private Sub Text18_GOTFOCUS()
  '
  Command3.SetFocus
  '
End Sub

Private Sub Text2_GotFocus()
  '
  Command3.SetFocus
  '
End Sub

Private Sub Text21_Change()
    '
    If Val(Text21.TEXT) < 0 Or Val(Text21.TEXT) > 255 Then
       Text21.TEXT = ""
    End If
    '
    If Val(Text21) < 1 Then
        TxtLiPre.TEXT = "CONCERTADOS FIJOS"
        ARCHIPRE$ = "MASTER"
      Else
        TxtLiPre.TEXT = ECOARCH$("LISTAS", MKI$(Val(Text21.TEXT)))
        ARCHIPRE$ = TRIM$(Text21)
        While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
        ARCHIPRE$ = "LIPRE" + ARCHIPRE$
    End If
    '
End Sub


Sub BLANFORMU()
   '
   Text2 = ""
   Text3 = ""
   Text24 = ""

  Text4.TEXT = ""

  Text19.TEXT = ""

     Text23 = ""
     Text35 = ""
     Text36 = ""
     Text37 = ""
   
     Text3 = ""
     
     Text18 = ""
     Text5 = ""
     Text7 = ""
     Text8 = ""
     '
     Text9 = ""
     Text10 = ""
     Text25 = ""
     Text24 = ""
     Text11 = ""
     Text12 = ""
     Text17 = ""

   List3.Clear

   '
   YACAR% = 0
   '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
   '
   Unload Me
   '
End Sub

Private Sub GruCli_Click()
With OPMASCL04
  .Option1(9).Value = True
  Call .Command16_Click
End With
End Sub



Private Sub LiPreCli_Click()
With OPMASCL04
  .Option1(7).Value = True
  Call .Command16_Click
End With
End Sub




Private Sub Print_Click()
   Call Command25_Click
End Sub

Private Sub PrZoVeCli_Click()
With OPMASCL04
  .Option1(5).Value = True
  Call .Command16_Click
End With
End Sub

Sub PRILISMAT()
    '
    'ACA VA CODIGO PARA IMPRIMIR OBSERVACIONES
    'Call SETULTREG("!OBSERVOF", 0)
    '
    ' IMPRESION DE LA LISTA DE MATERIALES
    '
    FORIPRE$ = CONTROL$("", "LISTAMAT")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
' VER ESTO CON WALTER
' TODA LA PARTE DE LOS PARAMETROS TENES QUE HACERLA VOS, FLAVIO
      FECHDOC$ = Text9
      NUMEDOC$ = TRIM$(Label31) + " / " + TRIM$(Label2)
      TIPODOC$ = "Lista de Materiales "
      '
      DESTIDOC% = Val(Label22) 'CLIENTE
      '
      CODPARAM$(1) = "DOCUORIG" 'NUMERO DE OBRA
      CODPARAM$(2) = "REF-MAT1" 'CODIGO DE GRUPO
      CODPARAM$(3) = "RASO-CLI" 'CLIENTE
      CODPARAM$(4) = "FECH-EMI" 'FECHA DE EMISION
      CODPARAM$(5) = "REF-MAT2" 'DENOMINACION DE LA OBRA
      CODPARAM$(6) = "REF-MAT3" 'DENOMINACION DEL GRUPO
      CODPARAM$(7) = "TELE-CLI" 'RESPONSABLE
      
      CAPARDOC% = 7
      
      DOCPARAM(1) = Label31
      DOCPARAM(2) = Label2
      DOCPARAM(3) = Label5
      DOCPARAM(4) = Text9
      DOCPARAM(5) = Label7
      DOCPARAM(6) = Label8
      DOCPARAM(7) = Text11
      '
 'A PARTIR DE AQUI LO MODIFIQUE YO - EPB
      CAITAB1% = 0
      CAITCRU% = 0
      '
 'AQUI PONER LOS CODIGOS DE COLUMNAS A LISTAR
      CODTABU1$(1) = "ORD-ITEM"
      CODTABU1$(2) = "CANTIDAD"
      CODTABU1$(3) = "DES-MAT"
      CODTABU1$(4) = "N-PLANO"
      CODTABU1$(5) = "MEDIDA-1"
      CODTABU1$(6) = "MEDIDA-2"
      CODTABU1$(7) = "COD-MAT"
      CODTABU1$(8) = "MEDIDA-3"
      
      CACOLTAB1% = 8
      '
      For U& = 1 To List3.ListCount
        '
        TTXX$ = List3.List(U& - 1)
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = Left$(TTXX$, 2)
        TETABU1$(2, CAITAB1%) = Mid$(TTXX$, 4, 6)
        TETABU1$(3, CAITAB1%) = Mid$(TTXX$, 11, 36)
        TETABU1$(4, CAITAB1%) = Mid$(TTXX$, 48, 14)
        TETABU1$(5, CAITAB1%) = Mid$(TTXX$, 63, 5)
        TETABU1$(6, CAITAB1%) = Mid$(TTXX$, 69, 12)
        TETABU1$(7, CAITAB1%) = Mid$(TTXX$, 82, 16)
        TETABU1$(8, CAITAB1%) = Mid$(TTXX$, 98, 2)
        '
        'AQUI AGREGA LA DESCRIPCION LARGA
        DELARGA$ = Mid$(TTXX$, 129)
        Call FRATEXTO$(DELARGA$, 36)
        For UI% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = ""
          TETABU1$(2, CAITAB1%) = ""
          TETABU1$(3, CAITAB1%) = RETEX$(UI%)
          TETABU1$(4, CAITAB1%) = ""
          TETABU1$(5, CAITAB1%) = ""
          TETABU1$(6, CAITAB1%) = ""
          TETABU1$(7, CAITAB1%) = ""
          TETABU1$(8, CAITAB1%) = ""
        Next UI%
        '
        'ESTO ES PARA QUE EL PROXIMO EMPIECE EN UN NUEVO RENGLON TRIPLE
        While (CAITAB1% Mod 3) <> 0
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = ""
          TETABU1$(2, CAITAB1%) = ""
          TETABU1$(3, CAITAB1%) = ""
          TETABU1$(4, CAITAB1%) = ""
          TETABU1$(5, CAITAB1%) = ""
          TETABU1$(6, CAITAB1%) = ""
          TETABU1$(7, CAITAB1%) = ""
          TETABU1$(8, CAITAB1%) = ""
        Wend
        '
      Next U&
      '
      NOCONFIRM% = 1
      Call PRINTDOC("LISTAMAT")                   ' LISTA DE MATERIALES
      '
      'LIMPIA LA LISTA
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
    End If
    '
199 Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1

End Sub
'
Sub CARGODETPRESU() 'CARGO LISTA DE DETALLE DE UN PRESU
    If TRIM$(Label31) <> "" Then
      NUPRESU$ = TRIM$(Label31)
      Call ABREPRECAM
      SQLX$ = "SELECT * FROM DETAPRESU "
      SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
      SQLX$ = SQLX$ + "AND RUBRO_DET = 2 "
      SQLX$ = SQLX$ + "ORDER BY CodExte_Det ASC "
      Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
      J& = 0
      With DETPRETMP
       '
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
        On Error GoTo 0
        List1.Clear
        TOTPIS# = 0
        Do Until .EOF = True
          CODI1$ = CODEXT(!CODINTE_DET)
          DESCRI1$ = DESCRIT0(!CODINTE_DET)
          PREUNI1# = !PreUni_Det
          CANTI1# = !Cantid_Det
          'PREUNI2$ = FORMATNUM(PREUNI1#, "F12.4")
          'CANTI2$ = FORMATNUM(CANTI1#, "F7.1")
          Y$ = Space$(256)
          Call REPLA(Y$, CODI1$, 1, 16)
          Call REPLA(Y$, DESCRI1$, 19, 56)
          'Call REPLA(Y$, PREUNI2$, 65, 12)
          'Call REPLA(Y$, CANTI2$, 78, 7)
          IMPOIT1# = PREUNI1# * CANTI1#
15        IMPOIT2$ = FORMATNUM(IMPOIT1#, "F12.2")
          Call REPLA(Y$, IMPOIT2$, 78, 12)
          List1.AddItem Y$
          TOTPIS# = TOTPIS# + IMPOIT1#
          Text2 = FORMATNUM(TOTPIS#, "F12.2")
          .MoveNext
        Loop
       End If
      End With
      'CALCULA EL TOTAL DE LA ACCESORIOS.
      TOTCAM# = TOTPIS#
      Text38 = FORMATNUM(TOTCAM#, "F12.2") 'ACCESORIOS
      Text13 = FORMATNUM(TOTCAM#, "F12.2") 'ACCESORIOS
    End If
    '
End Sub

Private Sub Text3_GotFocus()
  '
  Command3.SetFocus
  '
End Sub

Private Sub text38_gotfocus()
  '
  Command3.SetFocus
  '
End Sub
Private Sub text38_change()
  '
  TOTCAM1# = Val(Text38)
  REPRFLEX04.Label24 = FORMATNUM(TOTCAM1#, "F11.2")
  '
End Sub

Private Sub Text4_GOTFOCUS()
  '
  Command3.SetFocus
  '
End Sub

Private Sub Text5_GOTFOCUS()
  '
  Command3.SetFocus
  '
End Sub

Private Sub Text7_GOTFOCUS()
  '
    Command3.SetFocus
  '
End Sub

Private Sub Text8_GOTFOCUS()
  '
  Command3.SetFocus
  '
End Sub

Private Sub Text9_GOTFOCUS()
  '
  Command3.SetFocus
  '
End Sub

Private Sub TICAMBIS_CLICK()
  '
  Command3.SetFocus
  '
End Sub
Sub COACCE()
    '\\\OT-05-0393-WAR-10/06/05///
  Command22.Enabled = False
  If DERACCE%("COPYRUB", "Generar Rubros por Copia") < 2 Then
      GoTo 99
  End If
  '
9 Screen.MousePointer = 1
  '
  VDEF$ = Space$(64) + String$(6, 0)
  '
  OBX$ = OBJPLA$("DEACOPRE", VDEF$)
  If OBX$ = "" Then GoTo 99
  VDEF$ = OBX$
  '
  PRESU1$ = TRIM$(Mid$(OBX$, 1, 16))
  PRESU2$ = TRIM$(Mid$(OBX$, 17, 48))
  '//////////////VALIDACIONES////////////
  If TRIM$(PRESU1$) = "" Then
    Call COMUNI("Número de Presupuesto no valido")
    GoTo 9
  End If
  '
  If TRIM$(PRESU2$) = "" Then
    Call COMUNI("Número de Presupuesto no valido.")
    GoTo 9
  End If
  '
  If ECOARCH("IPRESCAM", PRESU1$) = "" Then
    Call MENSERR(24, "Presupuesto no valido.")
    GoTo 9
  End If
  '
  If ECOARCH("IPRESCAM", PRESU2$) = "" Then
    Call MENSERR(24, "Presupuesto no valido.")
    GoTo 9
  End If
  '
  If TRIM$(PRESU1$) = TRIM$(PRESU2$) Then
    Call MENSERR(24, "Copia Imposible.\Coinciden Presupuesto Origen y Destino.")
    GoTo 99
  End If
  'validación de que en el presupuesto
  'no haya detalle con este rubro.
  NRUB% = 2 'ACCESORIOS
  Call ABREPRECAM
  SQLX$ = " Nro_Presu= '" & PRESU2$ & "' "
  SQLX$ = SQLX$ + " AND Rubro_Det = " & NRUB% & " "
  DetaPresu.FindFirst SQLX$
  If DetaPresu.NoMatch = False Then
    If COMALTER%("El Presupuesto de Cámara " & PRESU2$ & " \ya Contiene Detalle en Rubro.\¿Desea Sobre-Escribirlo ?\\Si, Sobre-escribir.\\No, Cancelar.") = 2 Then
      On Error GoTo 0
      Screen.MousePointer = 1
      Command22.Enabled = True
      Exit Sub
    Else
13    DetaPresu.FindFirst SQLX$
      If DetaPresu.NoMatch = False Then
        DetaPresu.Delete
        GoTo 13
      End If
    End If
  End If
  '
17  SQLX$ = " SELECT * FROM detapresu "
    SQLX$ = SQLX$ + " WHERE Nro_Presu= '" & PRESU1$ & "' "
    SQLX$ = SQLX$ + " AND Rubro_Det = " & NRUB% & ""
    Set DETAPRESUTMP = PresuCam.OpenRecordset(SQLX$, 4)
    '
    With DETAPRESUTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
     Else
        'VALIDA LA MONEDA EN ENCABEZADO DE PRESUP. DESTINO
        With EncaPresu
          .FindFirst " Nro_Presu = '" & PRESU1$ & "' "
          MONEDET% = !Mone_Presu
          .FindFirst " Nro_Presu = '" & PRESU2$ & "' "
          MONDEST% = !Mone_Presu
        End With
          TC1# = 1
          TC2# = 1
        If MONDEST% <> MONEDET% Then
          TC1# = TICAMONE(MONEDET%)
          TC2# = TICAMONE(MONDEST%)
        End If
        '
       Do While .EOF = False
        NURUB% = !rubro_det
        SUBRUB% = !SUBRUB_DET
        CODET% = !CODINTE_DET
        CODEXTET$ = !CODEXTE_DET
        CANMAT# = !Cantid_Det
        PREUMAT# = !PreUni_Det
        '
        MED1 = !MED1_DET
        MED2 = !Med2_Det
        MED3 = !Med3_Det
        IMPODEP# = !Impo_Det
        DESCRIDET$ = !DESCRI_DET
        HORADET = !Hora_Det
        OPERADET% = !Opera_Det
        VALOHORDET# = !ValoHor_Det
        OBSGRU$ = !OBSERVA_GRUP
        With DetaPresu
          .AddNew
            !Nro_PRESU = PRESU2$
            !rubro_det = NURUB%
            !SUBRUB_DET = SUBRUB%
            !CODINTE_DET = CODET%
            !CODEXTE_DET = CODEXTET$
            !Cantid_Det = CANMAT#
            !PreUni_Det = PREUMAT# * TC1# / TC2#
            !MONE_DET = MONEDET%
            !MED1_DET = MED1
            !Med2_Det = MED2
            !Med3_Det = MED3
            !Impo_Det = (PREUMAT# * TC1# / TC2#) * CANMAT#
            !DESCRI_DET = DESCRIDET$
            !Hora_Det = HORADET
            !Opera_Det = OPERADET%
            !ValoHor_Det = VALOHORDET#
            !OBSERVA_GRUP = OBSGRU$
          .Update
        End With
        .MoveNext
       Loop
       'COPIA EXITOSA
       FORPRE04.Text1 = PRESU2$
       Label31 = PRESU2$
       Call Form_Activate
       '
       End If
    End With
    '
99  Screen.MousePointer = 1
    Command22.Enabled = True
    '\\\OT-05-0393-WAR-FIN///
  '
End Sub
