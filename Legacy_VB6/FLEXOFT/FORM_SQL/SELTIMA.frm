VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SELTIMA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Tipo de Material"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6585
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2730
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   23
      Top             =   630
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sin Filtro por Tipo de Material"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   21
      Top             =   3080
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CommandButton BOTREC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   5640
      Picture         =   "SELTIMA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Aprueba la Selección"
      Top             =   4080
      Width           =   885
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Filtro de Fechas Tardías"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   105
      TabIndex        =   12
      ToolTipText     =   "Doble-Click para Editar Rango de Fechas"
      Top             =   3480
      Width           =   5430
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sin Filtro por Fechas Tardias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   22
         Top             =   960
         Value           =   1  'Checked
         Width           =   3735
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
         Height          =   285
         Left            =   4935
         TabIndex        =   18
         Top             =   420
         Width           =   175
      End
      Begin VB.CommandButton Command9 
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
         Left            =   2310
         TabIndex        =   15
         Top             =   420
         Width           =   175
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   5400
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Material Reque- rido antes de:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2715
         TabIndex        =   17
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   3885
         TabIndex        =   16
         ToolTipText     =   "Se listarán los Materiales que Deberán Recibirse ANTES de la Fecha Indicada"
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label123 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O/C's a Emi- tir antes de:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         TabIndex        =   14
         Top             =   300
         Width           =   960
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1260
         TabIndex        =   13
         ToolTipText     =   "Se listarán los Requerimientos cuya O/C debe Emitirse ANTES de la Fecha Indicada"
         Top             =   420
         Width           =   1065
      End
   End
   Begin VB.ListBox LICOPEN 
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
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Tilde los Vendedores a Considerar"
      Top             =   900
      Width           =   5385
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   5325
      TabIndex        =   7
      Top             =   495
      Width           =   5385
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Nro.      Tipo de Material"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   370
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   10755
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   4020
      Left            =   5640
      ScaleHeight     =   3960
      ScaleWidth      =   1455
      TabIndex        =   2
      Top             =   0
      Width           =   1515
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
         Height          =   650
         Left            =   120
         Picture         =   "SELTIMA.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   1560
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   3240
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton BOTEXIT 
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
         Left            =   120
         Picture         =   "SELTIMA.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "SELTIMA.frx":075E
         Top             =   3570
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "SELTIMA.frx":2680
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "SELTIMA.frx":28B4
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5520
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   120
      Y1              =   3000
      Y2              =   3360
   End
   Begin VB.Label Label5 
      Height          =   435
      Left            =   4200
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label2 
      Height          =   435
      Left            =   4830
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Marque (tilde) los Tipo de Materiales que desea incluir: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   105
      TabIndex        =   10
      Top             =   0
      Width           =   5400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5500
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   5505
      X2              =   5505
      Y1              =   495
      Y2              =   3360
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   7440
      TabIndex        =   9
      Top             =   1140
      Width           =   1695
   End
End
Attribute VB_Name = "SELTIMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
   LICOPEN.Clear
   Unload Me
End Sub
'

Private Sub BOTREC_Click()
   ' VALIDA FECHAS
   FECHA1% = FECHANUM(Label7)
   FECHA2% = FECHANUM(Label3)
   If FECHA2% < FECHA1% Then
      Call MENSERR(24, "Rango de Fechas No Procesable")
      Exit Sub
   End If
   ' VALIDA LA SELECCION DE ALGUN VENDEDOR
   CANSELE% = 0
   For i = 1 To LICOPEN.ListCount
     If LICOPEN.Selected(i - 1) = True Then
       CANSELE% = CANSELE% + 1
     End If
   Next i
   Label5 = TRIM$(Str$(CANSELE%))
   'Call MENSERR(24, "Ningun Vendedor Ha Sido Elegido")
   'Exit Sub
   ' SI NO SELECCIONA NINGUNO, EMITE TODOS JUNTOS
   '
99 Label2 = "OK"
   Hide
End Sub


Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LICOPEN.TopIndex
   For U& = 1 To LICOPEN.ListCount
     If LICOPEN.Selected(U& - 1) = True Then
         LICOPEN.Selected(U& - 1) = False
       Else
         LICOPEN.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LICOPEN.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True
End Sub

Private Sub Command2_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   Label3 = VALACT1$("SELFECHA")
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   Label7 = VALACT1$("SELFECHA")
End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture2)
    ' CARGA LISTA CON TIPOS DE MATERIALES
    For U& = 1 To ULTREG&("TATIMAT")
       X$ = REGLEIDO$("TATIMAT", U&)
       NRO$ = Asc(Mid$(X$, 1, 1))
       nombre$ = Mid$(X$, 2, 30)
       NRO1$ = Space(3 - Len(NRO$)) & NRO$
       TTXX$ = NRO1$ & Space(7) & nombre$
       LICOPEN.AddItem TTXX$
    Next U&
    '
    ' FECHAS DE CORTE POR OMISION
    Label7 = FECHATEX(COMESACT)
    Label3 = FECHATEX(HOY(HO$))
    '
End Sub

Private Sub Frame6_DblClick()
   '
   DES1% = FECHANUM(Label7)
   HAS1% = FECHANUM(Label3)
   Call DESHASFECHA(DES1%, HAS1%, PERIO1$)
   If PERIO1$ <> "" Then
       Label7 = FECHATEX$(DES1%)
       Label3 = FECHATEX$(HAS1%)
   End If
   '
End Sub

Private Sub Label3_DblClick()
   Call Command2_Click
End Sub

Private Sub Label7_DblClick()
   Call Command9_Click
End Sub
