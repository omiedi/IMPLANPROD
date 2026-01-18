VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form REMDET09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle - Remuneraciones Ordinarias"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   840
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command10 
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
      Height          =   330
      Left            =   1930
      TabIndex        =   42
      Top             =   120
      Width           =   175
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   40
      Top             =   480
      Width           =   4665
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3900
      MaxLength       =   8
      TabIndex        =   38
      Top             =   120
      Width           =   2070
   End
   Begin VB.Frame Frame1 
      Height          =   5070
      Left            =   200
      TabIndex        =   3
      Top             =   840
      Width           =   5865
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   1140
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Text            =   "REMDET09.frx":0000
         Top             =   3720
         Width           =   4560
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         MaxLength       =   3
         TabIndex        =   29
         Top             =   3240
         Width           =   405
      End
      Begin VB.CommandButton Command8 
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
         Left            =   1545
         Picture         =   "REMDET09.frx":0002
         TabIndex        =   28
         Top             =   3240
         Width           =   175
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         MaxLength       =   3
         TabIndex        =   25
         Top             =   2760
         Width           =   405
      End
      Begin VB.CommandButton Command7 
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
         Left            =   1545
         Picture         =   "REMDET09.frx":030C
         TabIndex        =   24
         Top             =   2760
         Width           =   175
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         MaxLength       =   3
         TabIndex        =   21
         Top             =   2280
         Width           =   405
      End
      Begin VB.CommandButton Command6 
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
         Left            =   1545
         Picture         =   "REMDET09.frx":0616
         TabIndex        =   20
         Top             =   2280
         Width           =   175
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         MaxLength       =   3
         TabIndex        =   17
         Top             =   1800
         Width           =   405
      End
      Begin VB.CommandButton Command5 
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
         Left            =   1545
         Picture         =   "REMDET09.frx":0920
         TabIndex        =   16
         Top             =   1800
         Width           =   175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         MaxLength       =   3
         TabIndex        =   13
         Top             =   1320
         Width           =   405
      End
      Begin VB.CommandButton Command2 
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
         Left            =   1545
         Picture         =   "REMDET09.frx":0C2A
         TabIndex        =   12
         Top             =   1320
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         MaxLength       =   3
         TabIndex        =   9
         Top             =   840
         Width           =   405
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1545
         Picture         =   "REMDET09.frx":0F34
         TabIndex        =   8
         Top             =   840
         Width           =   175
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         MaxLength       =   3
         TabIndex        =   5
         Top             =   360
         Width           =   405
      End
      Begin VB.CommandButton Command19 
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
         Left            =   1545
         Picture         =   "REMDET09.frx":123E
         TabIndex        =   4
         Top             =   360
         Width           =   175
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formula:"
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
         Left            =   -855
         TabIndex        =   33
         Top             =   3780
         Width           =   1905
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Recalculo:"
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
         Left            =   -15
         TabIndex        =   31
         Top             =   3345
         Width           =   1110
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   30
         Top             =   3240
         Width           =   3855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Quincena:"
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
         Left            =   -15
         TabIndex        =   27
         Top             =   2865
         Width           =   1110
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   26
         Top             =   2760
         Width           =   3855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Retención:"
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
         Left            =   -15
         TabIndex        =   23
         Top             =   2385
         Width           =   1110
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   22
         Top             =   2280
         Width           =   3855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Carácter:"
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
         Left            =   -15
         TabIndex        =   19
         Top             =   1905
         Width           =   1110
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   18
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Modo:"
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
         Left            =   -735
         TabIndex        =   15
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   14
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
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
         Left            =   -15
         TabIndex        =   11
         Top             =   945
         Width           =   1110
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   10
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
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
         Left            =   -15
         TabIndex        =   7
         Top             =   465
         Width           =   1110
      End
      Begin VB.Label Label23 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   6
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   6135
      Left            =   6240
      ScaleHeight     =   6075
      ScaleWidth      =   1335
      TabIndex        =   0
      Top             =   0
      Width           =   1395
      Begin VB.CommandButton Command21 
         Caption         =   "New"
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
         Left            =   90
         Picture         =   "REMDET09.frx":1548
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Alta de Nuevo Item de Remuneraciones Ordinarias"
         Top             =   1785
         Width           =   650
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Save"
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
         Left            =   90
         Picture         =   "REMDET09.frx":1692
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Guardar "
         Top             =   2730
         Width           =   640
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
         Left            =   90
         Picture         =   "REMDET09.frx":199C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exit"
         Top             =   120
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
         Left            =   90
         Picture         =   "REMDET09.frx":1AE6
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ayuda en Línea"
         Top             =   810
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -330
         Picture         =   "REMDET09.frx":1DF0
         Top             =   5400
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1995
      OleObjectBlob   =   "REMDET09.frx":3D12
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   330
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "REMDET09.frx":3F46
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Abreviatura:"
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
      Left            =   1905
      TabIndex        =   39
      Top             =   225
      Width           =   1905
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
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
      Left            =   120
      TabIndex        =   36
      Top             =   585
      Width           =   1110
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   120
      TabIndex        =   35
      Top             =   240
      Width           =   1110
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   34
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "REMDET09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '
    Call SELECHO("TATIREM")
    VDEVU$ = TRIM$(VALACT1$("TATIREM"))
    If VDEVU$ <> "" Then
       Text1.TEXT = VALACT1$("TATIREM")
    End If
    '
End Sub

Private Sub Command10_Click()
  '
  Call SELECHO("ICONCEP")
  COD$ = TRIM$(VALACT1$("ICONCEP"))
  If COD$ = "" Then Exit Sub
  '
  JJ& = Val(COD$)
  ZZ$ = REGLEIDO$("CONCEP", JJ&)
  DESCRI$ = TRIM$(Mid$(ZZ$, 1, 30))
  ABRE$ = TRIM$(Mid$(ZZ$, 121, 8))
  PROP$ = Mid$(ZZ$, 31, 8)
  FORMU$ = Mid$(ZZ$, 49, 72)
  Text8 = ABRE$ 'ABREVIATURA
  Text7 = FORMU$ 'FORMULA
    EST% = Asc(Mid$(PROP$, 1, 1))
    ET$ = ESTADO$(EST%)
  Text16 = ET$ 'ESTADO
  Text1 = Asc(Mid$(PROP$, 2, 1)) 'TIPO
  Text2 = Asc(Mid$(PROP$, 3, 1)) 'MODO
  Text3 = Asc(Mid$(PROP$, 4, 1)) 'CARACTER
  Text4 = Asc(Mid$(PROP$, 5, 1)) 'RETENCION
  Text5 = Asc(Mid$(PROP$, 6, 1)) 'QUINCENAL
  Text6 = Asc(Mid$(PROP$, 7, 1)) 'RECALCULO
  Label15 = COD$
  Text9 = DESCRI$
  '
End Sub

Private Sub Command19_Click()
    '
    Call SELECHO("TASTATUS")
    VDEVU$ = TRIM$(VALACT1$("TASTATUS"))
    If VDEVU$ <> "" Then
       Text16.TEXT = VALACT1$("TASTATUS")
    End If
    '
End Sub

Private Sub Command2_Click()
    '
    Call SELECHO("TAMOCREM")
    VDEVU$ = TRIM$(VALACT1$("TAMOCREM"))
    If VDEVU$ <> "" Then
       Text2.TEXT = VALACT1$("TAMOCREM")
    End If
    '
End Sub

Private Sub Command21_Click()
 '
  Command21.Enabled = False
  'VDEF$ TIENE QUE DAR EL PROX. Nº DE REMUNERACIÓN
  VDEF$ = Space(32)
  PRONU% = Int(ULTREG&("CONCEP"))
  If PRONU% > 99 Then
    Call MENSERR(24, "Imposible Agregar Nuevo Concepto.\  \Debe Re-Definir Alguno de los Pre-Existentes.")
    Exit Sub
  End If
  '
  Call REPLA(VDEF$, MKI$(PRONU% + 1), 1, 2)
9 OBJ$ = OBJPLA$("NUEREM", VDEF$)
  If OBJ$ = "" Then GoTo 99
  VDEF$ = OBJ$
  '
  DESCRI$ = TRIM$(Mid$(OBJ$, 3, 30))
  If TRIM$(DESCRI$) = "" Then
    Call MENSERR(24, "Descripción no Válida.")
    GoTo 9
  End If
  '
  ZZ$ = Space(64)
  NUME = CVI(Mid$(OBJ$, 1, 2))
  NUMER$ = FORMATNUM$(NUME, "F5")
  NUME1$ = AJUSTD$(Val(NUMER$), 7)
  Call REPLA(ZZ$, NUME1$, 1, 7)
  Call REPLA(ZZ$, DESCRI$, 10, 30)
  'List1.AddItem ZZ$
  '
  ZZ1$ = Space(128)
  Call REPLA(ZZ1$, DESCRI$, 1, 30) 'DESCRIPCION
  Call REPLA(ZZ1$, Chr$(1), 31, 1) 'ESTADO ACTIVO
  Call REPLA(ZZ1$, Chr$(1), 32, 1) 'TIPO NORMAL
  Call REPLA(ZZ1$, Chr$(0), 33, 1) 'MODO CALCULO INTERNO
  Call REPLA(ZZ1$, Chr$(0), 34, 1) 'CARACTER OPCIONAL
  Call REPLA(ZZ1$, Chr$(1), 35, 1) 'RETENCION GRAVADO
  Call REPLA(ZZ1$, Chr$(0), 36, 1) 'QUINCENA TODAS
  Call REPLA(ZZ1$, Chr$(0), 37, 1) 'RECALCULO NO
  Call REPLA(ZZ1$, Chr$(0), 38, 1) 'FILLER
  Call REPLA(ZZ1$, " ", 39, 10) 'FILLER 1
  Call REPLA(ZZ1$, " ", 49, 72) 'FORMULA
  Call REPLA(ZZ1$, " ", 121, 8) 'ABREVIATURA
  Call REGAPP("CONCEP", ZZ1$)
  Call CIERRARCH("CONCEP")
  Call FRESHECHO("CONCEP")
  '
99 Command21.Enabled = True
  '
End Sub

Private Sub Command3_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command5_Click()
    '
    Call SELECHO("TACARAC")
    VDEVU$ = TRIM$(VALACT1$("TACARAC"))
    If VDEVU$ <> "" Then
       Text3.TEXT = VALACT1$("TACARAC")
    End If
    '
End Sub

Private Sub Command6_Click()
    '
    Call SELECHO("TASURET")
    VDEVU$ = TRIM$(VALACT1$("TASURET"))
    If VDEVU$ <> "" Then
       Text4.TEXT = VALACT1$("TASURET")
    End If
    '
End Sub

Private Sub Command7_Click()
    '
    Call SELECHO("TAQUINLI")
    VDEVU$ = TRIM$(VALACT1$("TAQUINLI"))
    If VDEVU$ <> "" Then
       Text5.TEXT = VALACT1$("TAQUINLI")
    End If
    '
End Sub

Private Sub Command8_Click()
    '
    Call SELECHO("TARECALC")
    VDEVU$ = TRIM$(VALACT1$("TARECALC"))
    If VDEVU$ <> "" Then
       Text6.TEXT = VALACT1$("TARECALC")
    End If
    '
End Sub

Private Sub Command9_Click()
  Command9.Enabled = False
  Screen.MousePointer = 11
  'VALIDACIONES
  If TRIM$(Text9) = "" Then
    Call MENSERR(24, "Descripción no Válida.")
    GoTo 99
  End If
  '
  If TRIM(RECFILT("TATIREM", 1, MKI$(Val(Text1)))) = "" Then
    Call MENSERR(24, "Tipo de Remuneración no Válido.")
    GoTo 99
  End If
  '
  If TRIM(RECFILT("TAMOCREM", 1, MKI$(Val(Text2)))) = "" Then
    Call MENSERR(24, "Modo de Calculo de Remuneración no Válido.")
    GoTo 99
  End If
  '
  If TRIM(RECFILT("TACARAC", 1, MKI$(Val(Text3)))) = "" Then
    Call MENSERR(24, "Carácter no Válido.")
    GoTo 99
  End If
  '
  If TRIM(RECFILT("TASURET", 1, MKI$(Val(Text4)))) = "" Then
    Call MENSERR(24, "Sujeto a Retención no Válido.")
    GoTo 99
  End If
  '
  If TRIM(RECFILT("TAQUINLI", 1, MKI$(Val(Text5)))) = "" Then
    Call MENSERR(24, "Quincena no Válida.")
    GoTo 99
  End If
  '
  If TRIM(RECFILT("TARECALC", 1, MKI$(Val(Text6)))) = "" Then
    Call MENSERR(24, "Modo de Recalculo no Válido.")
    GoTo 99
  End If
  '
  If TRIM(RECFILT("TASTATUS", 1, Text16)) = "" Then
    Call MENSERR(24, "Estado no Válido.")
    GoTo 99
  End If
  REGI& = Val(Label15)
  ZZ1$ = REGLEIDO("CONCEP", REGI&)
  '
  DESCRI$ = TRIM$(Text9)
  ABREV$ = TRIM$(Text8)
  FORMU$ = Text7
  ESTA1% = ESTADO1%(TRIM(Text16))
  '
  Call REPLA(ZZ1$, DESCRI$, 1, 30) 'DESCRIPCION
  Call REPLA(ZZ1$, Chr$(ESTA1%), 31, 1) 'ESTADO
  Call REPLA(ZZ1$, Chr$(Val(Text1)), 32, 1) 'TIPO NORMAL
  Call REPLA(ZZ1$, Chr$(Val(Text2)), 33, 1) 'MODO CALCULO INTERNO
  Call REPLA(ZZ1$, Chr$(Val(Text3)), 34, 1) 'CARACTER OPCIONAL
  Call REPLA(ZZ1$, Chr$(Val(Text4)), 35, 1) 'RETENCION GRAVADO
  Call REPLA(ZZ1$, Chr$(Val(Text5)), 36, 1) 'QUINCENA TODAS
  Call REPLA(ZZ1$, Chr$(Val(Text6)), 37, 1) 'RECALCULO NO
  Call REPLA(ZZ1$, Chr$(0), 38, 1) 'FILLER
  Call REPLA(ZZ1$, " ", 39, 10) 'FILLER 1
  Call REPLA(ZZ1$, FORMU$, 49, 72) 'FORMULA
  Call REPLA(ZZ1$, ABREV$, 121, 8) 'ABREVIATURA
  Call GRAREG("CONCEP", ZZ1$, REGI&)
  Call SETULTREG("CONCEP", ULTREG&("CONCEP"))
  Call CARICONC
  '
99 Screen.MousePointer = 1
  Command9.Enabled = True
  '
End Sub

Private Sub Form_Load()
  '
  Call APLICACIONSKINS(Me)
  Call CENTRAFORM(Me)
  Call CARICONC
  '
End Sub

Private Sub Text1_Change()
  '
  Label1 = ECOARCH$("TATIREM", MKI$(Val(Text1)))
  '
End Sub

Private Sub Text16_Change()
  '
  Text16 = UCase(Text16)
  Label23 = ECOARCH$("TASTATUS", Text16)
  '
End Sub

Private Sub Text2_Change()
  '
  Label3 = ECOARCH$("TAMOCREM", MKI(Val(Text2)))
  '
End Sub

Private Sub Text3_Change()
  '
  Label5 = ECOARCH$("TACARAC", MKI$(Val(Text3)))
  '
End Sub

Private Sub Text4_Change()
  Label7 = ECOARCH$("TASURET", MKI$(Val(Text4)))
End Sub

Private Sub Text5_Change()
  Label9 = ECOARCH$("TAQUINLI", MKI(Val(Text5)))
End Sub

Private Sub Text6_Change()
  Label11 = ECOARCH$("TARECALC", MKI(Val(Text6)))
End Sub

Function ESTADO1%(ESTA$) 'DEVUELVE EL VALOR EN $ DEL ESTADO
  Select Case ESTA$
    Case "S"
      ESTADO1% = 0 'SUSPENDIDO
    Case "A"
      ESTADO1% = 1 'ACTIVO
    Case "F"
      ESTADO1% = 2 'FANTASMA
  End Select
End Function

Function ESTADO$(EST%) 'DEVUELVE EL VALOR EN $ DEL ESTADO
  Select Case EST%
    Case 0
      ESTADO$ = "S" 'SUSPENDIDO
    Case 1
      ESTADO$ = "A" 'ACTIVO
    Case 2
      ESTADO$ = "F" 'FANTASMA
  End Select
End Function

Sub CARICONC()
   JJ& = 0
   REBLA$ = REGBLAN$("ICONCEP")
   For U& = 1 To ULTREG&("CONCEP")
     XX$ = REGLEIDO$("CONCEP", U&)
     CICO% = U&
     DECO$ = TRIM$(Left$(XX$, 30))
     YY$ = REBLA$
     Call REPLA(YY$, MKI$(CICO%), 1, 2)
     Call REPLA(YY$, DECO$, 3, 30)
     JJ& = JJ& + 1
     Call GRAREG("ICONCEP", YY$, JJ&)
   Next U&
   Call CIERRARCH("ICONCEP")
   Call FRESHECHO("ICONCEP")
End Sub
