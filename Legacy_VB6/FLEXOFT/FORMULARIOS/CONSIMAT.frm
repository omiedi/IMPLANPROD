VERSION 5.00
Begin VB.Form CONSIMAT 
   BackColor       =   &H00E0FFD0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Solicitudes por Número"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Opcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   10920
      TabIndex        =   40
      Top             =   3360
      Width           =   875
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
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
         Left            =   95
         Picture         =   "CONSIMAT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Listado de SIM por Obra"
         Top             =   1510
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
         Height          =   600
         Left            =   95
         Picture         =   "CONSIMAT.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Listado de SIM por Artículo"
         Top             =   910
         Width           =   650
      End
      Begin VB.CommandButton Command23 
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
         Left            =   95
         Picture         =   "CONSIMAT.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Listado de SIM por Número"
         Top             =   315
         Width           =   650
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Otros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1550
      Left            =   6510
      TabIndex        =   34
      Top             =   210
      Width           =   4215
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
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
         Index           =   4
         Left            =   -945
         TabIndex        =   39
         Top             =   1155
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   2
         Left            =   -945
         TabIndex        =   38
         Top             =   375
         Width           =   1905
      End
      Begin VB.Label LENTREGA 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   1050
         TabIndex        =   37
         Top             =   315
         Width           =   2970
      End
      Begin VB.Label LENTREGA 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   1050
         TabIndex        =   36
         Top             =   720
         Width           =   2970
      End
      Begin VB.Label LENTREGA 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   1050
         TabIndex        =   35
         Top             =   1125
         Width           =   2970
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   10970
      Picture         =   "CONSIMAT.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Consultas y Re-Impresión"
      Top             =   2520
      Width           =   750
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Solicitud Interna de Materiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1550
      Left            =   105
      TabIndex        =   18
      Top             =   210
      Width           =   6315
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2205
         TabIndex        =   20
         Top             =   315
         Width           =   180
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1155
         TabIndex        =   33
         Top             =   1095
         Width           =   4965
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Parte:"
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
         Index           =   3
         Left            =   3045
         TabIndex        =   32
         Top             =   800
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3990
         TabIndex        =   26
         Top             =   735
         Width           =   2130
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Obra:"
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
         TabIndex        =   25
         Top             =   765
         Width           =   750
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1155
         TabIndex        =   24
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Left            =   3675
         TabIndex        =   23
         Top             =   375
         Width           =   1275
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1155
         TabIndex        =   22
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5040
         TabIndex        =   21
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
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
         Left            =   -840
         TabIndex        =   19
         Top             =   380
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   105
      TabIndex        =   7
      Top             =   6300
      Width           =   10620
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1175
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   735
         Width           =   7155
      End
      Begin VB.CommandButton Command3 
         Caption         =   "O.K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   745
         Left            =   8360
         Picture         =   "CONSIMAT.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1140
         Width           =   750
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         Height          =   425
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   10515
         TabIndex        =   9
         Top             =   315
         Width           =   10580
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            Height          =   15
            Left            =   0
            TabIndex        =   11
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   10
            Top             =   75
            Width           =   10200
         End
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   10665
         Y1              =   1900
         Y2              =   1900
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
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
         Left            =   8400
         TabIndex        =   29
         Top             =   840
         Width           =   4000
      End
      Begin VB.Line Line7 
         X1              =   10580
         X2              =   10580
         Y1              =   315
         Y2              =   1900
      End
      Begin VB.Line Line6 
         X1              =   8295
         X2              =   10605
         Y1              =   1130
         Y2              =   1130
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Registrar Anotación"
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
         Left            =   9345
         TabIndex        =   28
         Top             =   1310
         Width           =   960
      End
      Begin VB.Line Line5 
         X1              =   8295
         X2              =   8295
         Y1              =   735
         Y2              =   1995
      End
      Begin VB.Label Label9 
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
         Height          =   1165
         Left            =   0
         TabIndex        =   12
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label10 
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
         Height          =   1170
         Left            =   8310
         TabIndex        =   13
         Top             =   735
         Width           =   2260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Anotaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4450
      Left            =   105
      TabIndex        =   5
      Top             =   1800
      Width           =   10620
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   3795
         Left            =   10290
         TabIndex        =   31
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4005
         Left            =   8085
         TabIndex        =   17
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4110
         Left            =   1155
         TabIndex        =   16
         Top             =   630
         Width           =   15
      End
      Begin VB.ListBox List1 
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
         Height          =   3660
         Left            =   0
         TabIndex        =   15
         Top             =   735
         Width           =   10580
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10515
         TabIndex        =   6
         Top             =   315
         Width           =   10580
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   27
            Top             =   105
            Width           =   10200
         End
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10605
         Y1              =   4410
         Y2              =   4410
      End
      Begin VB.Line Line1 
         X1              =   10580
         X2              =   10580
         Y1              =   320
         Y2              =   6225
      End
   End
   Begin VB.CommandButton Command26 
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
      Height          =   700
      Left            =   10970
      Picture         =   "CONSIMAT.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Configurar Comprobante de SIM"
      Top             =   6825
      Width           =   750
   End
   Begin VB.CommandButton Command2 
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
      Height          =   700
      Left            =   10970
      Picture         =   "CONSIMAT.frx":1374
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1785
      Width           =   750
   End
   Begin VB.CommandButton Command9 
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
      Height          =   700
      Left            =   10970
      Picture         =   "CONSIMAT.frx":167E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Sale de la Aplicación"
      Top             =   315
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   10970
      Picture         =   "CONSIMAT.frx":17C8
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Graba las Modificaciones"
      Top             =   1050
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10340
      Picture         =   "CONSIMAT.frx":1AD2
      Top             =   7710
      Width           =   2010
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
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
      Left            =   -105
      TabIndex        =   4
      Top             =   2625
      Width           =   1695
   End
End
Attribute VB_Name = "CONSIMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub COMMAND21_CLICK()
'   Command21.Enabled = False
'   Call COPYSTRU("INDIPEDI", "PEPENDES")
'   RFF$ = RECFILT$("DESPAPED", 7, Chr$(1))
'   NPXA& = 0: JJ& = 0: FIN& = Len(RFF$)
'   For U& = 1 To FIN& Step 4
'     Call TRACE(20, U&, FIN&)
'     REDE& = CVS(Mid$(RFF$, U&, 4))
'     XX$ = REGLEIDO$("DESPAPED", REDE&)
'     NPX& = CVS(Left$(XX$, 4))
'     If NPX& <> NPXA& Then
'       NPXA& = NPX&
'       NCLIE% = CVI(Mid$(XX$, 7, 2))
'       NUOCO$ = TRIM$(Mid$(XX$, 9, 16))
'       FEPE = CVI(Mid$(XX$, 5, 2))
'       REPAR$ = "": If Mid$(XX$, 51, 1) = "R" Then REPAR$ = " (Reparación)"
'       YY$ = MKS$(NPX&) + TRIM$(RASOCLI$(NCLIE%))
'       If NUOCO$ <> "" Then YY$ = YY$ + " - " + NUOCO$
'       YY$ = YY$ + " - " + FECHATEX$(FEPE) + REPAR$
'       JJ& = JJ& + 1
'       Call GRAREG("PEPENDES", YY$, JJ&)
'     End If
'   Next U&
'   '
'   Call SETULTREG("PEPENDES", JJ&)
'   If JJ& < 1 Then
'     Call COMUNI("No hay Pedidos\Pendientes de Despacho.")
'     GoTo 99
'   End If
'   '
'3  Call SELECHO("PEPENDES/Indice de Pedidos Pendientes de Despacho)")
'   NPX& = Val(VALACT1$("PEPENDES"))
'   If NPX& < 1 Then GoTo 99
'   '
'   Label7 = TRIM$(Str$(NPX&))
'   '
'99 Command21.Enabled = True
'End Sub

'Private Sub Command23_Click()
'   Command23.Enabled = False
'   OPLIDESP.Show 1
'   Call FRESHALL
'   DoEvents
'   If OPLIDESP.OPIMPRE = "" Then GoTo 99
'   '
'   NEPREPA% = 0
'   If OPLIDESP.Check1(1).Value = 1 Then NEPREPA% = 1
'   '
'   Screen.MousePointer = 11
'   Dim ACTI%(32767), STACT(32767), OFPEN(32767)
'   Dim CAPEPE%, NROPEPE&(4096), REGPEPE&(4096)
'   RFF$ = RECFILT$("DESPAPED", 7, Chr$(1))
'   NPXA& = 0: JJ& = 0: REBLA$ = REGBLAN$("CONPEDIP")
'   For U& = 1 To Len(RFF$) Step 4
'     REDE& = CVS(Mid$(RFF$, U&, 4))
'     XX$ = REGLEIDO$("DESPAPED", REDE&)
'     If Mid$(XX$, 51, 1) <> "R" Then
'       If NEPREPA% = 1 Then
'         IPEDIT$ = TRIM$(Mid$(XX$, 113, 10))
'         If IPEDIT$ <> "" Then
'           RFG$ = RECFILT$("PEDIPREP", 1, IPEDIT$)
'           If Len(RFG$) >= 4 Then
'             CAYAPRE = 0
'             For KU& = 1 To Len(RFG$) Step 4
'               REPRE& = CVS(Mid$(RFG$, KU&, 4))
'               YY$ = REGLEIDO$("PEDIPREP", REPRE&)
'               CAYAPRE = CAYAPRE + CVS(Mid$(YY$, 91, 4))
'             Next KU&
'             CANPE = CVS(Mid$(XX$, 75, 4)) - CAYAPRE
'             If CANPE < 0.05 Then GoTo 9
'           End If
'         End If
'       End If
'       '
'       CAPEPE% = CAPEPE% + 1
'       NROPEPE&(CAPEPE%) = CVS(Left$(XX$, 4))
'       REGPEPE&(CAPEPE%) = REDE&
'     End If
'9  Next U&
'   '
'   For U& = 1 To CAPEPE%
'     For V& = U& + 1 To CAPEPE%
'       If NROPEPE&(U&) > NROPEPE&(V&) Then
'         NROPX& = NROPEPE&(U&)
'         REPEX& = REGPEPE&(U&)
'         NROPEPE&(U&) = NROPEPE&(V&)
'         REGPEPE&(U&) = REGPEPE&(V&)
'         NROPEPE&(V&) = NROPX&
'         REGPEPE&(V&) = REPEX&
'       End If
'     Next V&
'   Next U&
'   '
'   NPEDXA& = 0
'   For UI% = 1 To CAPEPE%
'     REDE& = REGPEPE&(UI%)
'     XX$ = REGLEIDO$("DESPAPED", REDE&)
'     If Mid$(XX$, 51, 1) <> "R" Then
'       CAPEN = CVS(Mid$(XX$, 75, 4))
'       '
'       If NEPREPA% = 1 Then
'         IPEDIT$ = TRIM$(Mid$(XX$, 113, 10))
'         If IPEDIT$ <> "" Then
'           RFG$ = RECFILT$("PEDIPREP", 1, IPEDIT$)
'           If Len(RFG$) >= 4 Then
'             CAYAPRE = 0
'             For KU& = 1 To Len(RFG$) Step 4
'               REPRE& = CVS(Mid$(RFG$, KU&, 4))
'               ZZ$ = REGLEIDO$("PEDIPREP", REPRE&)
'               CAYAPRE = CAYAPRE + CVS(Mid$(ZZ$, 91, 4))
'             Next KU&
'             CAPEN = CAPEN - CAYAPRE
'           End If
'         End If
'       End If
'       '
'       YY$ = REBLA$
'       Call REPLA(YY$, Left$(XX$, 70), 1, 70)
'       Call REPLA(YY$, MKS$(CAPEN), 71, 4)
'       Call REPLA(YY$, String$(54, 0), 75, 54)
'       NPEDX& = CVS(Left$(YY$, 4))
'       If NPEDX& <> NPEDXA& Then STAPEA% = 1
'       '
'       CIXI% = CVI(Mid$(XX$, 33, 2))
'       If ACTI%(CIXI%) < 1 Then
'         ACTI%(CIXI%) = 1
'         STACT(CIXI%) = STOCKACT(CIXI%)
'         '
'         If NEPREPA% = 1 Then
'           'MsgBox "OJO Instruccion Deshabilitada"
'           STACT(CIXI%) = STACT(CIXI%) - STOPREP(CIXI%)
'         End If
'         '
'         If STACT(CIXI%) < 0 Then STACT(CIXI%) = 0
'         OFPEN(CIXI%) = STORFAB(CIXI%)
'       End If
'       '
'       Call REPLA(YY$, MKS$(STACT(CIXI%)), 75, 4)
'       Call REPLA(YY$, MKS$(OFPEN(CIXI%)), 79, 4)
'       FALTA = CAPEN - STACT(CIXI%) - OFPEN(CIXI%)
'       If FALTA < 0 Then FALTA = 0
'       Call REPLA(YY$, MKS$(FALTA), 83, 4)
'       JJ& = JJ& + 1
'       Call GRAREG("CONPEDIP", YY$, JJ&)
'       '
'       STAPE% = 1
'       If CAPEN > CVS(Mid$(YY$, 75, 4)) Then STAPE% = 2
'       If FALTA > 0 Then STAPE% = 3
'       If STAPE% > STAPEA% Then STAPEA% = STAPE%
'       '
'       For KK& = JJ& To 1 Step -1
'         ZZ$ = REGLEIDO$("CONPEDIP", KK&)
'         If CVS(Left$(ZZ$, 4)) <> NPEDX& Then GoTo 15
'         Call REPLA(ZZ$, MKI$(STAPEA%), 127, 2)
'         Call GRAREG("CONPEDIP", ZZ$, KK&)
'       Next KK&
'       '
'15     NPEDXA& = NPEDX&
'       If CAPEN <= STACT(CIXI%) Then
'           STACT(CIXI%) = STACT(CIXI%) - CAPEN
'           CAPEN = 0
'         Else
'           CAPEN = CAPEN - STACT(CIXI%)
'           STACT(CIXI%) = 0
'       End If
'       '
'       If CAPEN <= OFPEN(CIXI%) Then
'           OFPEN(CIXI%) = OFPEN(CIXI%) - CAPEN
'           CAPEN = 0
'         Else
'           CAPEN = CAPEN - OFPEN(CIXI%)
'           OFPEN(CIXI%) = 0
'       End If
'       '
'     End If
'   Next UI%
'   '
'   Call SETULTREG("CONPEDIP", JJ&)
'   Screen.MousePointer = 1
'   If JJ& < 1 Then
'     Call COMUNI("No hay Pedidos\Pendientes de Despacho.")
'     GoTo 99
'   End If
'   '
'   Call HEADERS("CONPEDIP", "")
'   MAGRU% = Val(OPLIDESP.AGRUPA)
'   If MAGRU% = 1 Then
'       OPX% = Val(OPLIDESP.OPIMPRE) - 1
'       If OPX% > 3 Then GoTo 99
'       '
'       If OPX% < 1 Then
'            SITU$ = "Todos los Pendientes": If NEPREPA% = 1 Then SITU$ = SITU$ + " - Netos de Preparados"
'            Call HEADERS("CONPEDIP", "Situación: " + SITU$)
'          Else
'            KK& = 0
'            For U& = 1 To JJ&
'              YY$ = REGLEIDO$("CONPEDIP", U&)
'              If CVI(Mid$(YY$, 127, 2)) = OPX% Then
'                KK& = KK& + 1
'                Call GRAREG("CONPEDIP", YY$, KK&)
'              End If
'            Next U&
'            Call SETULTREG("CONPEDIP", KK&)
'            Call CIERRARCH("CONPEDIP")
'            If OPX% = 1 Then
'                 SITU$ = "Listos para Despachar": If NEPREPA% = 1 Then SITU$ = SITU$ + " - Netos de Preparados"
'                 Call HEADERS("CONPEDIP", "Situación: " + SITU$)
'              ElseIf OPX% = 2 Then
'                 SITU$ = "En Preparación / Fabricación": If NEPREPA% = 1 Then SITU$ = SITU$ + " - Netos de Preparados"
'                 Call HEADERS("CONPEDIP", "Situación: " + SITU$)
'              ElseIf OPX% = 3 Then
'                 SITU$ = "Para Fabricar": If NEPREPA% = 1 Then SITU$ = SITU$ + " - Netos de Preparados"
'                 Call HEADERS("CONPEDIP", "Situación: " + SITU$)
'            End If
'       End If
'       Call FINAL("*LIDESPE")
'     ElseIf MAGRU% = 2 Then
'       Call FILESORT("CONPEDIP", "", 7, 7, "ASC")
'       Call FINAL("*LIDESQU")
'   End If
'   '
'99 Command23.Enabled = True
'   '
End Sub

Private Sub Command24_Click()
'   Command24.Enabled = False
'   Call DESPAMAT
'   Command24.Enabled = True
End Sub

Private Sub Command23_Click()
   OPSIMAT.Option1 = False
   OPSIMAT.Option2 = True
   OPSIMAT.Command4_Click
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   CFGSIMAT.Show 1
   Command26.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   NPX& = Val(Label7)
   If NPX& > 0 Then
     If TRIM$(Text1) <> "" Then
       Call FRATEXTO(Text1, 56)
       For KKI% = 1 To CARETEX%
         TTXX$ = Space$(76)
         Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTXX$, HOS$, 1, 8)
           Call REPLA(TTXX$, USERTER$, 69, 24)
         End If
         Call REGAPP("ANOTASIM", Left$(MKS$(NPX&) + TTXX$, 96))
       Next KKI%
       Call REGAPP("ANOTASIM", MKS$(NPX&) + String$(88, "="))
       Call CIERRARCH("ANOTASIM")
       Text1 = ""
       '
       List1.Clear
       RFF$ = RECFILT$("ANOTASIM", 1, MKS$(NPX&))
       '
       For U& = 1 To Len(RFF$) Step 4
         RENOTA& = CVS(Mid$(RFF$, U&, 4))
         XX$ = REGLEIDO$("ANOTASIM", RENOTA&)
         List1.AddItem Mid$(XX$, 5, 92)
       Next U&
       '
     End If
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   command1.Enabled = False
3  Call SELECHO("INDISIM/Indice General de Solicitudes Internas de Materiales")
   NPX& = Val(VALACT1$("INDISIM"))

'3  Call SELECHO("BDSIMAT/Indice General de Solicitudes Internas de Materiales")
'   NPX& = Val(VALACT1$("BDSIMAT"))
   If NPX& < 1 Then GoTo 99
   '
   Label7 = TRIM$(Str$(NPX&))
   '
99 command1.Enabled = True
End Sub

Private Sub Command5_Click()
   OPSIMAT.Option1 = False
   OPSIMAT.Option2 = True
   OPSIMAT.Command9_Click
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   NPX& = Val(Label7)
   '
   If NPX& > 0 Then
      TIDOCUM$ = "Solicitud Interna de Materiales"
'      If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
'        TIDOCUM$ = "Orden de Despacho"
'      End If
      NUDOCU$ = TRIM$(Str$(NPX&))
      While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
      NUDOCU$ = NUDOCU$ + " "
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
      '
      Screen.MousePointer = 1
      Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a esta Solicitud")
      On Error Resume Next
      OPPEDIDO.Command46.SetFocus
      On Error GoTo 0
      OPPEDIDO.Show 1
      GoTo 99
      '
5     For KKI% = 0 To 5
        EDITFORM.Picture1(KKI%).Cls
        EDITFORM.Picture1(KKI%).DrawWidth = 3
        EDITFORM.Picture1(KKI%).DrawStyle = 0
        EDITFORM.Picture1(KKI%).FillStyle = 1
      Next KKI%
      '
      REDOCU$ = REGACT$("PDOCLST")
      URE& = ULTREG&("PDOCSPL")
      LI& = 0: LS& = URE&
6     E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
      L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then LI& = L&: GoTo 6
      If DCC& > DOCUNU& Then LS& = L&: GoTo 6
      GoTo 8
      '
7     Screen.MousePointer = 1
      Call MENSERR(24, "Documento no Registrado")
      GoTo 99
      '
8     TPSP$ = ""
      Screen.MousePointer = 11
      While L& > 1
        DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
        If DCC& < DOCUNU& Then GoTo 9
        L& = L& - 1
      Wend
      '
9     CAPAGINAS% = 0
      For KKL& = L& To URE&
        TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
        If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
        If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
          TPSP$ = TPSP$ + Mid$(TBUF$, 5)
        End If
      Next KKL&
      '
95    SPOOLSTRING$ = TPSP$
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
10    PPXX% = InStr(TPSP$, Chr$(255))
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
      Screen.MousePointer = 1
      PAGINACTIVA% = 0
      For KKI% = 0 To 5
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
      EDITFORM.Command3.Enabled = False
      EDITFORM.Command5.Enabled = False
      EDITFORM.Show 1
    End If
    '
99  Command6.Enabled = True
    '
End Sub

Private Sub Command7_Click()
'   Command7.Enabled = False
'   PREPAPED.Show 1
'   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   OPSIMAT.Option1 = False
   OPSIMAT.Option2 = True
   OPSIMAT.Command8_Click
End Sub

Private Sub Command9_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Form_Load()
    Call CENTRAFORM(Me)
    '
    DoEvents
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Screen.MousePointer = 1
    HOII% = HOY(HOS$)
    Label9 = " " + HOS$
    Label12 = Left$(USERNAME, 17)
    '
End Sub


Private Sub Label7_Change()
   List1.Clear
   If Val(Label7) < 1 Then Exit Sub
   '
   NPX& = Val(Label7)
   RFF$ = RECFILT$("BDSIMAT", 1, MKS$(NPX&))
   If Len(RFF$) >= 4 Then
     REPED& = CVS(Left$(RFF$, 4))
     XX$ = REGLEIDO$("BDSIMAT", REPED&)
     FEX = CVI(Mid$(XX$, 5, 2))
     Label8 = FECHATEX$(FEX)
     Label11 = Mid$(XX$, 7, 8)
     Label16 = ECOARCH$("INDIOBR", Label11)
     Label14 = Mid$(XX$, 15, 24)
     NCLIE% = CVI(Mid$(XX$, 39, 2))
     LENTREGA(0) = RASOCLI$(NCLIE%)
     LENTREGA(1) = DOMICLI$(NCLIE%)
     LENTREGA(2) = Mid$(XX$, 43, 22)
     'Label18 = TRIM$(Str$(Asc(Mid$(XX$, 267, 1))))
     '
'     LENTREGA(0) = Mid$(XX$, 73, 32)
'     LENTREGA(1) = Mid$(XX$, 105, 32)
'     LENTREGA(2) = Mid$(XX$, 137, 32)
   End If
   '
'   Label14.ForeColor = RGB(0, 0, 255)
'   Label14 = "Venta"
'   RFQ$ = RECFILT$("PEDDETA", 1, MKS$(NPX&))
'   For UKKK& = 1 To Len(RFQ$) Step 4
'     UU& = CVS(Mid$(RFQ$, UKKK&, 4))
'     XX$ = REGLEIDO$("PEDDETA", UU&)
'     If CVS(Left$(XX$, 4)) = NPX& Then
'       If UCase$(Mid$(XX$, 51, 1)) = "R" Then
'         Label14.ForeColor = RGB(255, 0, 0)
'         Label14 = "Reparación"
'       End If
'       GoTo 20
'     End If
'   Next UKKK&
'   '
20 RFF$ = RECFILT$("ANOTASIM", 1, MKS$(NPX&))
   For U& = 1 To Len(RFF$) Step 4
     RENOTA& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("ANOTASIM", RENOTA&)
     List1.AddItem Mid$(XX$, 5, 92)
   Next U&
   '
End Sub

Sub DESPAMAT()
'   '
'   Call COPYSTRU("INDIPEDI", "PEPENDES")
'   RFF$ = RECFILT$("DESPAPED", 7, Chr$(1))
'   NPXA& = 0: JJ& = 0: FIN& = Len(RFF$)
'   For U& = 1 To FIN& Step 4
'     Call TRACE(20, U&, FIN&)
'     REDE& = CVS(Mid$(RFF$, U&, 4))
'     XX$ = REGLEIDO$("DESPAPED", REDE&)
'     NPX& = CVS(Left$(XX$, 4))
'     If NPX& <> NPXA& Then
'       NPXA& = NPX&
'       NCLIE% = CVI(Mid$(XX$, 7, 2))
'       NUOCO$ = TRIM$(Mid$(XX$, 9, 16))
'       FEPE = CVI(Mid$(XX$, 5, 2))
'       REPAR$ = "": If Mid$(XX$, 51, 1) = "R" Then REPAR$ = " (Reparación)"
'       YY$ = MKS$(NPX&) + TRIM$(RASOCLI$(NCLIE%))
'       If NUOCO$ <> "" Then YY$ = YY$ + " - " + NUOCO$
'       YY$ = YY$ + " - " + FECHATEX$(FEPE) + REPAR$
'       JJ& = JJ& + 1
'       Call GRAREG("PEPENDES", YY$, JJ&)
'     End If
'   Next U&
'   '
'   Call SETULTREG("PEPENDES", JJ&)
'   If JJ& < 1 Then
'     Call COMUNI("No hay Pedidos\Pendientes de Despacho.")
'     GoTo 99
'   End If
'   '
'   If Val(Label7) > 0 Then
'     NPX& = Val(Label7)
'     RFF$ = RECFILT$("PEPENDES", 1, MKS$(NPX&))
'     If Len(RFF$) >= 4 Then GoTo 5
'   End If
'   '
'3  Call SELECHO("PEPENDES/Indice de Pedidos Pendientes de Despacho)")
'   NPX& = Val(VALACT1$("PEPENDES"))
'   If NPX& < 1 Then GoTo 99
'   '
'5  Label7 = TRIM$(Str$(NPX&))
'   '
'   FORIPRE$ = TRIM$(CONTROL$("", "CDESPACH"))
'   If FORIPRE$ <> "" Then
'     LU$ = LUDAT$
'     If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
'       RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
'       Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito en Consignación:\  \'" + RMCC$ + "'.")
'       Exit Sub
'     End If
'   End If
'   '
'   DEPOCON% = DEPOEXPE%
'   DEPORIG% = DEPOCON%
'   DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
'   '
'   Call CIERRARCH("*.*")
'   Call BLANARCH("!BOLREDET")
'   '
'   Screen.MousePointer = 11
'   NPED& = Val(Label7)
'   REBLA$ = REGBLAN$("DESPAPED")
'   RFF$ = RECFILT$("DESPAPED", 1, MKS$(NPED&))
'   JJ& = 0
'   '
'   For U& = 1 To Len(RFF$) Step 4
'     REDE& = CVS(Mid$(RFF$, U&, 4))
'     XX$ = REGLEIDO$("DESPAPED", REDE&)
'     STATU% = Asc(Mid$(XX$, 27, 1))
'     If STATU% < 2 Then
'       REPARX% = 0: If Mid$(XX$, 51, 1) = "R" Then REPARX% = 1
'       CI% = CVI(Mid$(XX$, 33, 2))
'       NCLIE% = CVI(Mid$(XX$, 7, 2))
'       CAPENE# = CVS(Mid$(XX$, 75, 4))
'       FENTRESO% = CVI(Mid$(XX$, 25, 2))
'       '
'       Y$ = REGBLAN$("!BOLREDET")
'       Call REPLA(Y$, MKI$(CI%), 83, 2)
'       Call REPLA(Y$, UNIMED$(CI%), 85, 4)
'       Call REPLA(Y$, MKD$(CAPENE#), 89, 8)
'       Call REPLA(Y$, MKS$(NPED&), 151, 4)
'       Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
'       Call REPLA(Y$, MKD$(CAPENE#), 163, 8)
'       Call REPLA(Y$, MKS$(REDE&), 189, 4)
'       Call REGAPP("!BOLREDET", Y$)
'     End If
'   Next U&
'   '
'   VTB% = VENTABU%("DESPAPED/NC/TITUPAN=Despacho Materiales Vendidos - " + TRIM$(RASOCLI$(NCLIE%)))
'   If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
'     Exit Sub
'   End If
'   '
'   For YX& = 1 To ULTREG&("!BOLREDET")
'     YY$ = REGLEIDO$("!BOLREDET", YX&)
'     CI% = CVI(Mid$(YY$, 83, 2))
'     CAN# = CVD(Mid$(YY$, 163, 8))
'     If CI% > 0 Then
'       If CI% <= NUMAS% Then
'         If CAN# > 0.005 Then
'           GoTo 15
'         End If
'       End If
'     End If
'   Next YX&
'   Call MENSERR(24, "Despacho Cancelado !!!\Faltan Cantidades o Códigos.")
'   Exit Sub
'   '
'15 CONTA% = ULTREG&("!BOLREDET")
'   If CONTA% > 0 Then
'     Call CIERRARCH("MOVISTO")
'     Call BLOQARCH("MOVISTO")
'     Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
'     '
'     TTXXX$ = "Materiales Vendidos"
'     If REPARX% = 1 Then
'        TTXXX$ = "Reparación"
'        NREP& = 0
'        For KU& = ULTREG&("PRESREPA") To 1 Step -1
'          XYX$ = REGLEIDO$("PRESREPA", KU&)
'          If CVS(Mid$(XYX$, 125, 4)) = NPED& Then
'            NREP& = CVS(Left$(XYX$, 4))
'            GoTo 16
'          End If
'        Next KU&
'     End If
'     '
'16   TTXYZ$ = Space$(76)
'     HOII% = HOY(HOS$)
'     Call REPLA(TTXYZ$, HOS$, 1, 8)
'     Call REPLA(TTXYZ$, "Despacho de " + TTXXX$ + ":", 11, 56)
'     Call REPLA(TTXYZ$, USERTER$, 69, 24)
'     Call REGAPP("ANOTAPED", MKS$(NPED&) + TTXYZ$)
'     Call REGAPP("ANOTAPED", MKS$(NPED&) + Space$(88))
'     If REPARX% = 1 Then
'       Call REGAPP("ANOTAREP", MKS$(NREP&) + TTXYZ$)
'       Call REGAPP("ANOTAREP", MKS$(NREP&) + Space$(88))
'     End If
'     '
'     AINDI% = 0
'     For YX& = 1 To ULTREG&("!BOLREDET")
'       YY$ = REGLEIDO$("!BOLREDET", YX&)
'       CI% = CVI(Mid$(YY$, 83, 2))
'       CAN# = CVD(Mid$(YY$, 163, 8))
'       LOTE$ = Mid$(YY$, 69, 12)
'       If CI% > 0 Then
'         If CI% <= NUMAS% Then
'           If CAN# > 0.005 Then
'             NUOC& = CVS(Mid$(YY$, 151, 4))
'             REDE& = CVS(Mid$(YY$, 189, 4))
'             DOSEC$ = ""
'             CMO$ = "PL"
'             If REPARX% = 1 Then CMO$ = "*D"
'             REMI$ = "PL." + TRIM$(Str$(NPED&))
'             FF% = HOY(HOS$)
'             '
'             CANTI = (-1) * CAN#
'             Call MOVISTO(FF%, CI%, LOTE$, CMO$, REMI$, DOSEC$, RASOCLI$(NCLIE%), 0, "$", NC1%, DEPORIG%, CANTI)
'             '
'             ZZ$ = REGLEIDO$("DESPAPED", REDE&)
'             CAENT = CAN#
'             Call REPLA(ZZ$, MKS$(CAENT), 95, 4)
'             Call REPLA(ZZ$, Chr$(2), 27, 1) ' PASA A STATUS ENTREGADO
'             Call GRAREG("DESPAPED", ZZ$, REDE&)
'             '
'             TTXX$ = Space$(76)
'             CIXI% = CI%
'             Call REPLA(TTXX$, "- " + CODEXT$(CIXI%), 11, 18)
'             Call REPLA(TTXX$, DESCRIT0$(CIXI%), 29, 38)
'             Call REPLA(TTXX$, CSTRING$(MKS$(CAENT), 3, 10, 1, 2), 57, 10)
'             Call REGAPP("ANOTAPED", Left$(MKS$(NPED&) + TTXX$, 96))
'             If REPARX% = 1 Then
'               Call REGAPP("ANOTAREP", MKS$(NREP&) + TTXX$)
'             End If
'             '
'             CAPEN = CVS(Mid$(ZZ$, 75, 4)) - CAENT
'             If CAPEN > 0.05 Then
'               Call REPLA(ZZ$, MKS$(CAPEN), 75, 4)
'               Call REPLA(ZZ$, MKS$(0), 95, 4)
'               Call REPLA(ZZ$, Chr$(1), 27, 1)
'               Call REGAPP("DESPAPED", ZZ$)
'             End If
'             '
'             IPEDIT$ = Mid$(ZZ$, 113, 10)
'             RFG$ = RECFILT$("PEDIPREP", 1, IPEDIT$)
'             For KU& = 1 To Len(RFG$) Step 4
'               REPRE& = CVS(Mid$(RFG$, KU&, 4))
'               YY$ = REGLEIDO$("PEDIPREP", REPRE&)
'               CAPREP = CVS(Mid$(YY$, 91, 4))
'               If CAPREP <= CAENT + 0.005 Then
'                    Call GRAREG("PEDIPREP", String$(128, 0), REPRE&)
'                    CAENT = CAENT - CAPREP
'                  Else
'                    Call REPLA(YY$, MKS$(CAPREP - CAENT), 91, 4)
'                    Call GRAREG("PEDIPREP", YY$, REPRE&)
'                    CAENT = 0
'               End If
'             Next KU&
'             '
'             JJ& = 0
'             FIN& = ULTREG&("PEDIPREP")
'             For U& = 1 To FIN&
'               ZZ$ = REGLEIDO$("PEDIPREP", U&)
'               If CVS(Mid$(ZZ$, 17, 4)) > 0 Then
'                 JJ& = JJ& + 1
'                 Call GRAREG("PEDIPREP", ZZ$, JJ&)
'               End If
'             Next U&
'             Call SETULTREG("PEDIPREP", JJ&)
'             Call CIERRARCH("PEDIPREP")
'             '
'           End If
'         End If
'       End If
'     Next YX&
'     '
'     Call REGAPP("ANOTAPED", Left$(MKS$(NPED&) + Space$(10) + String$(56, "-"), 96))
'     Call REGAPP("ANOTAPED", MKS$(NPED&) + String$(88, "="))
'     If REPARX% = 1 Then
'       Call REGAPP("ANOTAREP", Left$(MKS$(NREP&) + Space$(10) + String$(56, "-"), 96))
'       Call REGAPP("ANOTAREP", MKS$(NREP&) + String$(88, "="))
'     End If
'     '
'     Call CIERRARCH("ANOTAPED")
'     Call CIERRARCH("ANOTAREP")
'     Call CIERRARCH("MOVISTO")
'     Call CIERRARCH("MOVIREPA")
'     Call CIERRARCH("DESPAPED")
'     '
'     Label7 = ""
'     Label7 = TRIM$(Str$(NPED&))
'     '
'     Screen.MousePointer = 1
'     '
'     If TRIM$(FORIPRE$) <> "" Then
'       '
'       FEX = FF%
'       FECHDOC$ = FECHATEX$(FEX)
'       NUMEDOC$ = TRIM$(REMI$)
'       TIPODOC$ = "Despacho de Materiales"
'       DESTIDOC% = NCLIE%
'       '
'       CODPARAM$(1) = "UBI-MPRI": DOCPARAM$(1) = TRIM$(Label14)
'       CAPARDOC% = 1
'       '
'       CODTABU1$(1) = "COD-MAT"
'       CODTABU1$(2) = "DES-MAT"
'       CODTABU1$(3) = "CANTIDAD"
'       CODTABU1$(4) = "CAKILOS"
'       CODTABU1$(5) = "UNIMED"
'       CACOLTAB1% = 5
'       '
'       CAITAB1% = 0
'       '
'       For YX& = 1 To ULTREG&("!BOLREDET")
'         YY$ = REGLEIDO$("!BOLREDET", YX&)
'         CI% = CVI(Mid$(YY$, 83, 2))
'         CAN# = CVD(Mid$(YY$, 89, 8))
'         If CI% > 0 Then
'           If CI% <= NUMAS% Then
'             If CAN# > 0.005 Then
'               UNID$ = TRIM$(Mid$(YY$, 85, 4))
'               If UNID$ = "" Then UNID$ = UNIMED$(CI%)
'               CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
'               KILO$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 12, 3, 2)
'               NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
'               NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
'               '
'               CAITAB1% = CAITAB1% + 1
'               TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
'               TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
'               TETABU1$(3, CAITAB1%) = CANS$
'               TETABU1$(4, CAITAB1%) = KILO$
'               TETABU1$(5, CAITAB1%) = UNID$
'               '
'             End If
'           End If
'         End If
'       Next YX&
'       '
'       CANCELPRINT% = 0
'       If CONTROL$("CDESPACH", "SUPRINT") = "SI" Then CANCELPRINT% = 1
'       Call PRINTDOC("CDESPACH")   ' Comprobante de Despacho
'       '
'     End If
'   End If
'   '
'99 Call CIERRARCH("*.*")
'   '
End Sub
