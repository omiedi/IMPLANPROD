VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form SHOWPLA07 
   BackColor       =   &H00DAE4EB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generador de Reportes FLEXOFT"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   495
      Left            =   5000
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   873
      ButtonWidth     =   794
      ButtonHeight    =   714
      Appearance      =   1
      ImageList       =   "imlToolbarPics"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Search"
            Object.ToolTipText     =   "Busqueda por Texto"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PRINT"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "EXIT"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin ComctlLib.Toolbar Toolbar2 
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      ButtonWidth     =   794
      ButtonHeight    =   714
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Exportar a Excel"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grafic"
            Object.ToolTipText     =   "Acceso a Gráficos"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "FullScreen"
            Object.ToolTipText     =   "Pantalla Completa"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRID2 
      Height          =   6090
      Left            =   105
      TabIndex        =   24
      Top             =   2115
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   10742
      _Version        =   393216
      Rows            =   28
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   5197685
      ForeColorSel    =   12582912
      GridColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H009BBBC6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8760
      ScaleHeight     =   495
      ScaleWidth      =   3135
      TabIndex        =   23
      Top             =   50
      Width           =   3135
      Begin VB.Image Image1 
         Height          =   525
         Left            =   120
         Picture         =   "SHOWPLA07.frx":0000
         ToolTipText     =   "Exporta Reporte a Excel"
         Top             =   0
         Width           =   540
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   660
         Picture         =   "SHOWPLA07.frx":0F06
         ToolTipText     =   "Maximiza / Minimiza Pantalla"
         Top             =   0
         Width           =   570
      End
      Begin VB.Image Image3 
         Height          =   570
         Left            =   1200
         Picture         =   "SHOWPLA07.frx":1E3C
         ToolTipText     =   "Ajuste de Columnas"
         Top             =   0
         Width           =   585
      End
      Begin VB.Image Image4 
         Height          =   570
         Left            =   1800
         Picture         =   "SHOWPLA07.frx":304E
         ToolTipText     =   "Buscar"
         Top             =   0
         Width           =   465
      End
      Begin VB.Image Image5 
         Height          =   660
         Left            =   2280
         Picture         =   "SHOWPLA07.frx":3ED0
         ToolTipText     =   "Imprime Reporte"
         Top             =   0
         Width           =   645
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      Appearance      =   0  'Flat
      BackColor       =   &H002B9973&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   10200
      ScaleHeight     =   165
      ScaleWidth      =   1455
      TabIndex        =   21
      Top             =   720
      Width           =   1485
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H0048CC9D&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   -120
         TabIndex        =   22
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   6300
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DAE4EB&
      Caption         =   "CONDICIONES DE SELECCION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   150
      TabIndex        =   4
      Top             =   960
      Width           =   5010
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
         Height          =   945
         Left            =   4310
         Picture         =   "SHOWPLA07.frx":55C2
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   705
      End
      Begin VB.TextBox FilText1 
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
         Index           =   0
         Left            =   1200
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1185
         TabIndex        =   7
         Top             =   420
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
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
         Index           =   0
         Left            =   3735
         TabIndex        =   6
         ToolTipText     =   "Lista de Valores Posibles"
         Top             =   420
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label FiltEco1 
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
         Index           =   0
         Left            =   3780
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro 1:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   435
         TabIndex        =   8
         Top             =   465
         Visible         =   0   'False
         Width           =   660
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      IntegralHeight  =   0   'False
      ItemData        =   "SHOWPLA07.frx":5C66
      Left            =   8820
      List            =   "SHOWPLA07.frx":5C68
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.ListBox LSTMOSTARCMP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      ItemData        =   "SHOWPLA07.frx":5C6A
      Left            =   6090
      List            =   "SHOWPLA07.frx":5C6C
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   5880
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      IntegralHeight  =   0   'False
      ItemData        =   "SHOWPLA07.frx":5C6E
      Left            =   840
      List            =   "SHOWPLA07.frx":5C70
      TabIndex        =   1
      Top             =   6300
      Visible         =   0   'False
      Width           =   4950
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   10185
      TabIndex        =   9
      Top             =   525
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Teprueba 
      AutoSize        =   -1  'True
      BackColor       =   &H00DAE4EB&
      Caption         =   "TEPRUEBA"
      Height          =   240
      Left            =   6360
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   1245
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6510
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483644
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":5C72
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1B2A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1C2FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1C614
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Index           =   0
      Left            =   10815
      TabIndex        =   14
      Top             =   1260
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   10815
      TabIndex        =   13
      Top             =   945
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad de Registros"
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
      Left            =   8715
      TabIndex        =   12
      Top             =   945
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
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
      Index           =   0
      Left            =   10290
      TabIndex        =   11
      Top             =   1575
      Visible         =   0   'False
      Width           =   1380
   End
   Begin ComctlLib.ImageList imlToolbarPics 
      Left            =   7140
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483644
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1C92E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1CC48
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1CF62
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1D27C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1D596
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1D8B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1DBCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1DEE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1E1FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1E518
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1E832
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1EA0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1EB1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWPLA07.frx":1EE38
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo de la Consulta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   210
      TabIndex        =   5
      Top             =   210
      Width           =   2160
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   1995
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label9 
      BackColor       =   &H009BBBC6&
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "SHOWPLA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim YAENTRO%
Dim CANCELA%
'
Dim ENCACOLU$(), CANCOLS%
Dim ENCAFILA$(), CANFILS%
Dim SENTIOR%(256)
Dim ARBUS$, ABUSQUE$
'
Dim HAYTOTA%
Dim AMAXI%(256)
Dim SUMA#(256), SUBTOT#(256)
Dim TIESUM%(256), FORSUMA$(256)
Dim COLUVALOR
'
Sub REPLA_CONSULTA()
    For i& = 1 To Combo1.Count
       If Combo1(i& - 1).Visible = True Or Combo1(i& - 1).TEXT <> "" Or FilText1(i& - 1) <> "" Then
          TREPLA$ = Combo1(i& - 1).TEXT
          TBUS$ = "¿" + Label1(i& - 1)
5         PPXX% = InStr(TXT_CONSULTA, TBUS$)
          If PPXX% > 0 Then
            IZQUI$ = Left$(TXT_CONSULTA, PPXX% - 1)
            Resto$ = Mid$(TXT_CONSULTA, PPXX% + 1)
            PPXX% = InStr(Resto$, "?")
            If PPXX% < 1 Then PPXX% = 1 + Len(Resto$)
            DERE$ = Mid$(Resto$, PPXX% + 1)
            MEDIO$ = Left$(Resto$, PPXX% - 1)
            TREP$ = TREPLA$
            If InStr(UCase$(MEDIO$), "/MASTER") > 0 Then
               If CODINT%(TREPLA$) > 0 Then
                  'TREP$ = TRIM$(Str$(CODINT%(TREPLA$)))
               End If
            End If
            TXT_CONSULTA = IZQUI$ + TRIM$(TREP$) + DERE$
            ' AL APLICAR 'UNION ALL', CADA CONDICION A REEMPLAZAR EN LA CONSULTA
            ' SE DUPLICA PARA LA OTRA TABLA
            If UNIONALL$ <> "" And VUELTA% = 0 Then
                VUELTA% = 1
                GoTo 5
            End If
          End If
          
          VUELTA% = 0
55        PPXX% = InStr(TXT_CAREGIST, TBUS$)
          If PPXX% > 0 Then
            IZQUI$ = Left$(TXT_CAREGIST, PPXX% - 1)
            Resto$ = Mid$(TXT_CAREGIST, PPXX% + 1)
            PPXX% = InStr(Resto$, "?")
            If PPXX% < 1 Then PPXX% = 1 + Len(Resto$)
            DERE$ = Mid$(Resto$, PPXX% + 1)
            MEDIO$ = Left$(Resto$, PPXX% - 1)
            TREP$ = TREPLA$
            If InStr(UCase$(MEDIO$), "/MASTER") > 0 Then
               If CODINT%(TREPLA$) > 0 Then
                  'TREP$ = TRIM$(Str$(CODINT%(TREPLA$)))
               End If
            End If
            TXT_CAREGIST = IZQUI$ + TRIM$(TREP$) + DERE$
            ' AL APLICAR 'UNION ALL', CADA CONDICION A REEMPLAZAR EN LA CONSULTA
            ' SE DUPLICA PARA LA OTRA TABLA
            If UNIONALL$ <> "" And VUELTA% = 0 Then
                VUELTA% = 1
                GoTo 55
            End If
          End If
       End If
    Next i&
    '

'    For i& = 1 To Combo1.Count
'       If Combo1(i& - 1).Visible = True Or Combo1(i& - 1).TEXT <> "" Then
'          TREPLA$ = Combo1(i& - 1).TEXT
'          TBUS$ = "¿" + Label1(i& - 1)
'          PPXX% = InStr(TXT_CONSULTA, TBUS$)
'          If PPXX% > 0 Then
'            IZQUI$ = Left$(TXT_CONSULTA, PPXX% - 1)
'            RESTO$ = Mid$(TXT_CONSULTA, PPXX% + 1)
'            PPXX% = InStr(RESTO$, "?")
'            If PPXX% < 1 Then PPXX% = 1 + Len(RESTO$)
'            DERE$ = Mid$(RESTO$, PPXX% + 1)
'            TXT_CONSULTA = IZQUI$ + TRIM$(TREPLA$) + DERE$
'          End If
'          PPXX% = InStr(TXT_CAREGIST, TBUS$)
'          If PPXX% > 0 Then
'            IZQUI$ = Left$(TXT_CAREGIST, PPXX% - 1)
'            RESTO$ = Mid$(TXT_CAREGIST, PPXX% + 1)
'            PPXX% = InStr(RESTO$, "?")
'            If PPXX% < 1 Then PPXX% = 1 + Len(RESTO$)
'            DERE$ = Mid$(RESTO$, PPXX% + 1)
'            TXT_CAREGIST = IZQUI$ + TRIM$(TREPLA$) + DERE$
'          End If
'       End If
'    Next i&
End Sub
'

Sub EJECUTACONSULTA()
    '
    For KKI& = 1 To 256
       AMAXI%(KKI&) = 0
       SUMA#(KKI&) = 0
       SUBTOT#(KKI&) = 0
       TIESUM%(KKI&) = 0
       FORSUMA$(KKI&) = ""
    Next KKI&
    Label7 = "0"
    HAYTOTA% = 0
    '
    Call ARMA_CONSULTA
    Call REPLA_CONSULTA
    '
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    '
    Screen.MousePointer = 11
    On Error Resume Next
    Set RS = FLEXCONN.Execute(FILTSQL$(TXT_CAREGIST$))
    If Err Then
       On Error GoTo 0
       If APLINVO$ = "" Then Call MENSERR(24, "Imposible Ejecutar la Consulta")
       GoTo 99
    End If
    On Error GoTo 0
    FIN& = RS!CANREGIS
    '
    Set RS = New ADODB.Recordset
    On Error Resume Next
    Set RS = FLEXCONN.Execute(FILTSQL$(TXT_CONSULTA))
    '
    IMPOSEJEC% = 0
    If Err Then
       On Error GoTo 0
       If APLINVO$ = "" Then Call MENSERR(24, "Imposible Ejecutar la Consulta")
       IMPOSEJEC% = 1
       GoTo 99
    End If
    On Error GoTo 0
    '
    COLUFECHA = 0  ' XVALO(Left$(ComboFecha.TEXT, 3)) - 1
    COLUCLASE = 1  ' XVALO(Left$(ComboClase.TEXT, 3)) - 1
    COLUVALOR = 2  ' XVALO(Left$(ComboValor.TEXT, 3)) - 1
    COLUVALO2 = 3  ' XVALO(Left$(ComboValor.TEXT, 3)) - 1
    '
    CONECO% = 0: TECO$ = ""
    If TABLAECO$(2) <> "" Then
       If FORMAECO$(2) <> "" Then
          CONECO% = 1
          TECO$ = Left$(TRIM$(UCase$(TABLAECO$(2))), 5)
       End If
    End If
    '
    RWW = 0: CANFILS% = 0
    With RS
       Do While Not .EOF
          RWX& = RWX& + 1
          Call TRACE(20, RWX&, FIN&)
          'VALFECHA$ = RS.Fields(COLUFECHA).Value
          FF% = CVINT(RS.Fields(COLUFECHA).Value)
          VALFECHA$ = FECHATEXLAR$(FF%)
          MESAX$ = TRIM$(Mid$(VALFECHA$, 4))
          For JJ% = 1 To CANCOLS%
             If ENCACOLU$(JJ%) = MESAX$ Then GoTo 10
          Next JJ%
          CANCOLS% = CANCOLS% + 1
          ReDim Preserve ENCACOLU$(CANCOLS%)
          ENCACOLU$(CANCOLS%) = MESAX$
          '
10        VALCLASE = RS.Fields(COLUCLASE).Value
          VTP% = VarType(VALCLASE)
          ENCAFI$ = VALCLASE
          '
          For JJ% = 1 To CANFILS%
             If UCase$(TRIM$(ENCAFILA$(JJ%))) = UCase$(TRIM$(ENCAFI$)) Then GoTo 20
          Next JJ%
          CANFILS% = CANFILS% + 1
          ReDim Preserve ENCAFILA$(CANFILS%)
          '
          If VTP% >= 2 And VTP% <= 6 Then
               ENCAFILA$(CANFILS%) = AJUSTD$(ENCAFI$, 8)
             Else
               ENCAFILA$(CANFILS%) = TRIM$(ENCAFI$)
          End If
          '
20     .MoveNext
       Loop
    End With
    '
    For i& = 1 To CANCOLS%
      For J& = i& + 1 To CANCOLS%
        A = FECHANUM("01/" + ENCACOLU$(i&))
        B = FECHANUM("01/" + ENCACOLU$(J&))
        If FECHANUM("01/" + ENCACOLU$(i&)) > FECHANUM("01/" + ENCACOLU$(J&)) Then
           Call SWAP(ENCACOLU$(i&), ENCACOLU$(J&))
        End If
      Next J&
    Next i&
    '
    ORDEX$ = TRIM$(ORDCAMPO$(2))
    If ORDEX$ = "ASC" Then
         For i& = 1 To CANFILS%
           For J& = i& + 1 To CANFILS%
             If UCase$(ENCAFILA$(i&)) > UCase$(ENCAFILA$(J&)) Then
                Call SWAP(ENCAFILA$(i&), ENCAFILA$(J&))
             End If
           Next J&
         Next i&
      ElseIf ORDEX$ = "DESC" Then
         For i& = 1 To CANFILS%
           For J& = i& + 1 To CANFILS%
             If UCase$(ENCAFILA$(i&)) < UCase$(ENCAFILA$(J&)) Then
                Call SWAP(ENCAFILA$(i&), ENCAFILA$(J&))
             End If
           Next J&
         Next i&
    End If
    '
    Set RS = New ADODB.Recordset
    On Error Resume Next
    Set RS = FLEXCONN.Execute(FILTSQL$(TXT_CONSULTA))
    '
    If Err Then
       On Error GoTo 0
       If APLINVO$ = "" Then Call MENSERR(24, "Imposible Ejecutar la Consulta")
       GoTo 99
    End If
    On Error GoTo 0
    '
    GRID2.Rows = CANFILS% + 1
    Label7 = TRIM$(Str$(CANFILS%))
    GRID2.Cols = CANCOLS% + 1
    GRID2.Clear
    '
    ' PONE LOS ENCABEZADOS DE LAS FILAS
    GRID2.ColWidth(0) = 0
    GRID2.ColAlignment(0) = 1
    '
    ' FORMATO PARA CODIGO DE ARTICULO
    ESCODI% = 0
    If CANFILS% > 0 Then
       If Left$(UCase$(FORCAMPO$(2)), 1) = "C" Then
          ESCODI% = 1
          LOTEX = XVALO(Mid$(FORCAMPO$(2), 2))
          Label5.FontName = GRID2.CellFontName
          Label5.FontSize = GRID2.CellFontSize
          Label5.FontBold = GRID2.CellFontBold
          Label5.Caption = String$(LOTEX + 1, "A")
          GRID2.ColWidth(0) = Label5.Width
          JI& = 1
15        If JI& >= Label8.Count Then
             NUPIC% = Label8.Count
             Load Label8(NUPIC%)
             GoTo 15
          End If
          Label8(JI&) = Str$(Len(Label5) - 1)
       End If
    End If
    '
    For FILA% = 1 To CANFILS%
       GRID2.TextMatrix(FILA%, 0) = ENCAFILA$(FILA%)
       '
       If ESCODI% = 0 Then
          Label5.FontName = GRID2.CellFontName
          Label5.FontSize = GRID2.CellFontSize
          Label5.FontBold = GRID2.CellFontBold
          Label5.Caption = ENCAFILA$(FILA%) + "*"
          If GRID2.ColWidth(0) < Label5.Width Then
             GRID2.ColWidth(0) = Label5.Width
             JI& = 1
16           If JI& >= Label8.Count Then
                NUPIC% = Label8.Count
                Load Label8(NUPIC%)
                GoTo 16
             End If
             Label8(JI&) = Str$(Len(Label5) - 1)
          End If
       End If
       '
       'AQUI AGREGAR COLUMNA ECO
       If TECO$ <> "" Then
          TXVALO$ = TRIM$(ENCAFILA$(FILA%))
          valor = XVALO(TXVALO$)
          If TECO$ = "DEUDO" Or TECO$ = "CLIEN" Then
                VALOECO = rasocli$(XVALO(valor))
              ElseIf TECO$ = "PROVE" Or TECO$ = "ACRED" Then
                VALOECO = rasocli$((-1) * XVALO(valor))
              ElseIf TECO$ = "MASTE" Then
                   If Len(TXVALO$) > 4 Then
                        CIXI% = CODINT%(TXVALO$)
                      Else
                        CIXI% = valor
                   End If
                VALOECO = DESCRIT0$(CIXI%)
              ElseIf EXISTE%(LUDAT$ + TABLAECO$(2) + ".RFS") Then
                LFF% = LENFIELD%(TABLAECO$(2), 1)
                TF1% = TIPFIELD%(TABLAECO$(2), 1)
                If TF1% < 3 Then
                     ABUS$ = MKI$(XVALO(valor))
                  ElseIf TF1% < 5 Then
                     ABUS$ = MKS$(XVALO(valor))
                  ElseIf TF1% = 6 Then
                     ABUS$ = valor
                End If
                On Error Resume Next
                If LFF% = 2 Then ABUS$ = MKI$(Int(XVALO(valor)))
                If LFF% = 1 Then ABUS$ = Chr$(Int(XVALO(valor)))
                On Error GoTo 0
                VALOECO = ECOARCH$(TABLAECO$(2), ABUS$)
          End If
          JI& = 2
          If JI& > GRID2.Cols Then GRID2.Cols = JI&
          On Error Resume Next
          GRID2.TextMatrix(FILA%, JI& - 1) = FORMATNUM$(VALOECO, FORMAECO$(2))
          On Error GoTo 0
       End If
    Next FILA%
    '
    '
    'PONE ENCABEZADO A LA COLUMNA 1
    GRID2.TextMatrix(0, 0) = TITCAMPO(2)
    '
    '
45  'PONE ENCABEZADO DE LA COLUMNA ECO
    If TECO$ <> "" Then
       FECO$ = TRIM$(UCase$(FORMAECO$(2)))
       If FECO$ <> "" Then
          JI& = 2
          If JI& > GRID2.Cols Then GRID2.Cols = JI&
          AMAXI%(JI&) = 0
          GRID2.ColAlignment(JI& - 1) = 1
          ENCABE$ = "Referencia"
          If TECO$ = "DEUDO" Or TECO$ = "CLIEN" Then
               ENCABE$ = "R.Social Cliente"
            ElseIf TECO$ = "PROVE" Or TECO$ = "ACRED" Then
               ENCABE$ = "R.Social Proveedor"
            ElseIf TECO$ = "MASTE" Then
               ENCABE$ = "Descripcion"
            ElseIf TECO$ = "SUTOT" Then
               ENCABE$ = ">Sub-Total"
            ElseIf TECO$ = "COUNT" Then
               ENCABE$ = ">Cant"
          End If
          If Left$(ENCABE$, 1) = ">" Then
             ENCABE$ = Mid$(ENCABE$, 2)
             GRID.ColAlignment(JI& - 1) = 6
          End If
          Label5.FontName = GRID2.CellFontName
          Label5.FontSize = GRID2.CellFontSize
          Label5.FontBold = GRID2.CellFontBold
          Label5.Caption = TEXTBIN$("1.234567", FORMAECO$(2)) + "*"
          GRID2.ColWidth(JI& - 1) = Label5.Width
25        If JI& >= Label8.Count Then
             NUPIC% = Label8.Count
             Load Label8(NUPIC%)
             GoTo 25
          End If
          Label8(JI&) = Str$(Len(Label5) - 1)
          '
          GRID2.TextMatrix(0, JI& - 1) = ENCABE$
          '
       End If
    End If
    '
    ' PONE LOS ENCABEZADOS DE LAS RESTANTES COLUMNAS
    For ICOLU% = 1 To CANCOLS%
35     If ICOLU% + CONECO% >= GRID2.Cols Then
          GRID2.Cols = GRID2.Cols + 1
          GoTo 35
       End If
       GRID2.TextMatrix(0, ICOLU% + CONECO%) = ENCACOLU$(ICOLU%)
       '
       JI& = JI& + 1
36     If JI& >= Label8.Count Then
          NUPIC% = Label8.Count
          Load Label8(NUPIC%)
          GoTo 36
       End If
       Label5.FontName = GRID2.CellFontName
       Label5.FontSize = GRID2.CellFontSize
       Label5.FontBold = GRID2.CellFontBold
       Label5.Caption = TEXTBIN$("1.234567", FORCAMPO$(COLUVALOR + 1)) + "*"
       If GRID2.ColWidth(ICOLU% + CONECO%) < Label5.Width Then
          GRID2.ColWidth(ICOLU% + CONECO%) = Label5.Width
       End If
       Label8(JI&) = Str$(Len(Label5) - 1)
       '
    Next ICOLU%
    '
    CACOLTO% = CANCOLS%
    If COLUTOTA% > 0 Then
       CACOLTO% = CACOLTO% + 1
       If CACOLTO% + CONECO% >= GRID2.Cols Then GRID2.Cols = GRID2.Cols + 1
       COLUTOTA% = CACOLTO%
       GRID2.ColAlignment(CACOLTO% + CONECO%) = 6
       GRID2.TextMatrix(0, COLUTOTA% + CONECO%) = "Total"
       GRID2.ColWidth(COLUTOTA% + CONECO%) = Label5.Width
       JI& = JI& + 1
37     If JI& >= Label8.Count Then
          NUPIC% = Label8.Count
          Load Label8(NUPIC%)
          GoTo 37
       End If
       Label8(JI&) = Str$(Len(Label5) - 1)
    End If
    '
    If COLUMINI% > 0 Then
       CACOLTO% = CACOLTO% + 1
       If CACOLTO% + CONECO% >= GRID2.Cols Then GRID2.Cols = GRID2.Cols + 1
       COLUMINI% = CACOLTO%
       GRID2.ColAlignment(CACOLTO% + CONECO%) = 6
       GRID2.TextMatrix(0, COLUMINI% + CONECO%) = "Minimo"
       GRID2.ColWidth(COLUMINI% + CONECO%) = Label5.Width
       JI& = JI& + 1
38     If JI& >= Label8.Count Then
          NUPIC% = Label8.Count
          Load Label8(NUPIC%)
          GoTo 38
       End If
       Label8(JI&) = Str$(Len(Label5) - 1)
    End If
    '
    If COLUMAXI% > 0 Then
       CACOLTO% = CACOLTO% + 1
       If CACOLTO% + CONECO% >= GRID2.Cols Then GRID2.Cols = GRID2.Cols + 1
       COLUMAXI% = CACOLTO%
       GRID2.ColAlignment(CACOLTO% + CONECO%) = 6
       GRID2.TextMatrix(0, COLUMAXI% + CONECO%) = "Maximo"
       GRID2.ColWidth(COLUMAXI% + CONECO%) = Label5.Width
       JI& = JI& + 1
39     If JI& >= Label8.Count Then
          NUPIC% = Label8.Count
          Load Label8(NUPIC%)
          GoTo 39
       End If
       Label8(JI&) = Str$(Len(Label5) - 1)
    End If
    '
    If COLUPROM% > 0 Then
       CACOLTO% = CACOLTO% + 1
       If CACOLTO% + CONECO% >= GRID2.Cols Then GRID2.Cols = GRID2.Cols + 1
       COLUPROM% = CACOLTO%
       GRID2.ColAlignment(CACOLTO% + CONECO%) = 6
       GRID2.TextMatrix(0, COLUPROM% + CONECO%) = "Promedio"
       GRID2.ColWidth(COLUPROM% + CONECO%) = Label5.Width
       JI& = JI& + 1
40     If JI& >= Label8.Count Then
          NUPIC% = Label8.Count
          Load Label8(NUPIC%)
          GoTo 40
       End If
       Label8(JI&) = Str$(Len(Label5) - 1)
    End If
    '
    HAYTOTA% = 1
    GRID2.Top = Frame1.Top + Frame1.Height + 150
    If HAYTOTA% = 1 Then
       GRID2.Height = Height - 300 - GRID2.Top - 500
       PPX& = GRID2.Left
       For KI& = 1 To GRID2.Cols
          If KI& > Label4.Count Then
             NUPIC% = Label4.Count
             Load Label4(NUPIC%)
          End If
          Label4(KI& - 1).Caption = ""
          Label4(KI& - 1).Visible = False
          Label4(KI& - 1).Top = GRID2.Top + GRID2.Height + 100
          Label4(KI& - 1).Left = PPX& + 30
          Label4(KI& - 1).Width = GRID2.ColWidth(KI& - 1)
          If KI& > 1 + CONECO% Then
             TIESUM%(KI&) = 1
             Label4(KI& - 1).Visible = True
          End If
          PPX& = PPX& + GRID2.ColWidth(KI& - 1)
       Next KI&
    End If
    '
    Screen.MousePointer = 11
    RWX& = 0
    With RS
       Do While Not .EOF
          RWX& = RWX& + 1
          Call TRACE(20, RWX&, FIN&)
          VALOVALOR = 0: If COLUVALOR < RS.Fields.Count Then VALOVALOR = RS.Fields(COLUVALOR).Value
          If TRIM$(ORDCAMPO$(3)) = "-" Then VALOVALOR = (-1) * VALOVALOR
          VALOVALO2 = 0: If COLUVALO2 < RS.Fields.Count Then VALOVALO2 = RS.Fields(COLUVALO2).Value
          If TRIM$(ORDCAMPO$(4)) = "-" Then VALOVALO2 = (-1) * VALOVALO2
                    
          VALFECHA$ = RS.Fields(COLUFECHA).Value
          MESAX$ = TRIM$(Mid$(VALFECHA$, 4))
          For JJ% = 1 To CANCOLS%
             If ENCACOLU$(JJ%) = MESAX$ Then GoTo 110
          Next JJ%
          '
110       VALCLASE = RS.Fields(COLUCLASE).Value
          VTP% = VarType(VALCLASE)
          ENCAFI$ = VALCLASE
          '
          For II% = 1 To CANFILS%
             If UCase$(TRIM$(ENCAFILA$(II%))) = UCase$(TRIM$(ENCAFI$)) Then GoTo 120
          Next II%
          '
120       VALOANTE = XVALO(GRID2.TextMatrix(II%, JJ% + CONECO%))
          GRID2.TextMatrix(II%, JJ% + CONECO%) = FORMATNUM(VALOANTE + VALOVALOR + VALOVALO2, FORCAMPO$(COLUVALOR + 1))
          SUMA#(JJ% + CONECO% + 1) = SUMA#(JJ% + CONECO% + 1) + VALOVALOR + VALOVALO2
          '
       RS.MoveNext
       Loop
    End With
    '
    If SALACUMU% > 0 Then
       For II% = 1 To CANFILS%
         VALOAN# = 0
         For JJ% = 1 To CANCOLS%
            VALOAN# = VALOAN# + XVALO(GRID2.TextMatrix(II%, JJ% + CONECO%))
            GRID2.TextMatrix(II%, JJ% + CONECO%) = FORMATNUM(VALOAN#, FORCAMPO$(COLUVALOR + 1))
         Next JJ%
       Next II%
    End If
    '
    ' obtiene minimo,maximo,promedio,total
    For II% = 1 To CANFILS%
      SUFIL# = 0: VAMAX# = (-999999999999999#): VAMIN# = 999999999999999#
      For JJ% = 1 To CANCOLS%
         VACOL# = XVALO(GRID2.TextMatrix(II%, JJ% + CONECO%))
         If Abs(VACOL#) > 0.000005 Then
            If VACOL# > VAMAX# Then VAMAX# = VACOL#
            If VACOL# < VAMIN# Then VAMIN# = VACOL#
         End If
         SUFIL# = SUFIL# + VACOL#
      Next JJ%
      '
      If VAMIN# > VAMAX# Then
         VAMIN# = 0: VAMAX# = 0
      End If
      '
      If COLUMINI% > 0 Then GRID2.TextMatrix(II%, COLUMINI% + CONECO%) = FORMATNUM(VAMIN#, FORCAMPO$(COLUVALOR + 1))
      If COLUMAXI% > 0 Then GRID2.TextMatrix(II%, COLUMAXI% + CONECO%) = FORMATNUM(VAMAX#, FORCAMPO$(COLUVALOR + 1))
      If COLUTOTA% > 0 Then GRID2.TextMatrix(II%, COLUTOTA% + CONECO%) = FORMATNUM(SUFIL#, FORCAMPO$(COLUVALOR + 1))
      If COLUPROM% > 0 Then GRID2.TextMatrix(II%, COLUPROM% + CONECO%) = FORMATNUM(SUFIL# / CANCOLS%, FORCAMPO$(COLUVALOR + 1))
    Next II%
    '
    If ESCODI% = 1 Then
       For FILA% = 1 To CANFILS%
          VALCLASE = GRID2.TextMatrix(FILA%, 0)
          GRID2.TextMatrix(FILA%, 0) = FORMATNUM$(VALCLASE, FORCAMPO$(2))
       Next FILA%
    End If
    '
    For KI& = 1 To GRID2.Cols
       Label4(KI& - 1).Caption = TRIM$(FORMATNUM$(SUMA#(KI&), FORCAMPO$(COLUVALOR + 1)))
       Label4(KI& - 1).Refresh
    Next KI&
    '
    GRID2.Visible = True
    GRID2.Refresh
    Screen.MousePointer = 1
    '
99  Call COLOREAR_GRILLASQL(GRID2, &HDAE4EB, 1)
    Screen.MousePointer = 1
    '
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then
      For I1& = Index + 1 To Combo1.Count - 1
        If Combo1(I1&).Visible = True Then
          Combo1(I1&).SetFocus
          Exit Sub
        End If
      Next I1&
      Call Command1_Click
   End If
End Sub

Private Sub Command1_Click()
   Call EJECUTACONSULTA
End Sub

Private Sub Command2_Click(Index As Integer)
   '
   If FilText1(Index).ToolTipText <> "" Then
      ASELEC$ = TRIM$(FilText1(Index).ToolTipText)
      Select Case ASELEC$
        Case "PROVED"
            FilText1(Index).TEXT = ListaSeleccionProveedores(RetornaSegundoCampoEco$)
            Combo1(Index).TEXT = FilText1(Index).TEXT
            FiltEco1(Index).Caption = RetornaSegundoCampoEco$
        Case "DEUDOR"
            FilText1(Index).TEXT = ListaSeleccionClientes(RetornaSegundoCampoEco$)
            Combo1(Index).TEXT = FilText1(Index).TEXT
            FiltEco1(Index).Caption = RetornaSegundoCampoEco$
        Case "MASTER"
            FilText1(Index).TEXT = ListaSeleccionArticulos(RetornaSegundoCampoEco$)
            Combo1(Index).TEXT = FilText1(Index).TEXT
            FiltEco1(Index).Caption = RetornaSegundoCampoEco$
        Case Else
            Call SELECHO(ASELEC$)
            VEMOFREN% = 11
            FilText1(Index).TEXT = VALACT1$(ASELEC$)
      End Select
      Exit Sub
   End If
   
   CAMCONDI$ = CAMPOCOND$(Index + 1)
   If CAMCONDI$ <> "" Then
      Screen.MousePointer = 11
      SQLX$ = "SELECT DISTINCT " & CAMCONDI$ & " AS VALOR FROM " & TABSELE$ & " WITH(NOLOCK)"
      SQLX$ = SQLX$ + " ORDER BY " & CAMCONDI$ & " ASC"
      '
      Combo1(Index).Clear
      Dim VAPOSI As ADODB.Recordset
      Set VAPOSI = New ADODB.Recordset
      Set VAPOSI = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With VAPOSI
         Do While Not .EOF
            On Error Resume Next
            Combo1(Index).AddItem !valor
            On Error GoTo 0
         .MoveNext
         Loop
      End With
      Screen.MousePointer = 1
   End If
   '
End Sub

Private Sub Command3_Click()
   GRABAR07.MSChart1.ColumnCount = 3
   CAMESA% = GRID2.Cols - 2
   GRABAR07.MSChart1.RowCount = CAMESA%
   '
   For UJ% = 1 To CAMESA%
     GRABAR07.MSChart1.Row = UJ%
     For UK% = 1 To 3
       GRABAR07.MSChart1.Column = UK%
       GRABAR07.MSChart1.Data = 0
     Next UK%
   Next UJ%
   '
   GRABAR07.MSChart1.Repaint = True
   GRABAR07.MSChart1.RowCount = CAMESA%
   For KKI% = 1 To CAMESA%
     GRABAR07.MSChart1.Row = KKI%
     GRABAR07.MSChart1.RowLabel = GRID2.TextMatrix(0, KKI% + 1)
   Next KKI%
   '
   For UJ% = 1 To CAMESA%
      GRABAR07.MSChart1.Row = UJ%
      '
      GRABAR07.MSChart1.Column = 1
      COVEN = XVALO(Label4(UJ% + 1)) / 1000
      GRABAR07.MSChart1.Data = COVEN
   Next UJ%
   '
   GRABAR07.Show 1
   '
End Sub

Private Sub FilText1_Change(Index As Integer)
   If FilText1(Index).ToolTipText <> "" Then
      ASELEC$ = TRIM$(FilText1(Index).ToolTipText)
      If InStr(1, "MASTER_DEUDOR_PROVED", ASELEC$, vbTextCompare) > 0 Then
        If ASELEC$ = "MASTER" Then
            CodigoInterno% = CODINT%(FilText1(Index))
            DatoEco$ = DESCRIT0$(CodigoInterno%)
        End If
        FiltEco1(Index).Caption = DatoEco$
        Combo1(Index).TEXT = FilText1(Index).TEXT
        Exit Sub
      End If
      If EXISTE%(LUDAT$ + ASELEC$ + ".RFS") > 0 Then
         ABUS$ = FilText1(Index).TEXT
         If TIPFIELD%(ASELEC$, 1) = 1 Or TIPFIELD%(ASELEC$, 1) = 2 Then
             On Error Resume Next: ABUS$ = MKI$(XVALO(ABUS$)): On Error GoTo 0
           ElseIf TIPFIELD%(ASELEC$, 1) = 3 Or TIPFIELD%(ASELEC$, 1) = 4 Then
             On Error Resume Next: ABUS$ = MKS$(XVALO(ABUS$)): On Error GoTo 0
           ElseIf TIPFIELD%(ASELEC$, 1) = 5 Then
             On Error Resume Next: ABUS$ = MKI$(Int(FECHANUM(ABUS$))): On Error GoTo 0
         End If
         FiltEco1(Index).Caption = ECOARCH$(ASELEC$, ABUS$)
      End If
   End If
   Combo1(Index).TEXT = FilText1(Index).TEXT
End Sub

Private Sub Form_Activate()
    Timer2.Enabled = True
End Sub

Private Sub Form_LostFocus()
    Timer2.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   CANCELA% = 1
   Cancel = 0
End Sub

Private Sub GRID2_Click()
   If GRID2.Row <= 1 Then
      ' para apagar el color azul de la seleccion
      On Error Resume Next
      GRID2.RowSel = GRID2.Row
      On Error GoTo 0
   End If
End Sub

Private Sub GRID2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
       If ARBUS$ <> "" Then
          J& = GRID2.Row

          h& = 0: If GRID2.COL < GRID2.Cols Then h& = GRID2.COL + 1
10        For U& = J& To GRID2.Rows - 1
             HI& = h&: If U& > J& Then HI& = 0
             For i& = HI& To GRID2.Cols - 1
              If InStr(UCase$(GRID2.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
                GRID2.Row = U&: GRID2.COL = i&
                On Error Resume Next
                 VISIROWS = (GRID2.Height / GRID2.RowHeight(J&)) - 4
                If Err > 0 Then
                 VISIROWS = 1
                End If
                On Error GoTo 0
                 If VISIROWS > 0 Then
                    If GRID2.TopRow < GRID2.Row - VISIROWS Then
                      GRID2.TopRow = GRID2.Row - VISIROWS
                    End If
                 End If
                GRID2.SetFocus
                Exit Sub
              End If
              Next i&
          Next U&
          '
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If

End Sub

Private Sub GRID2_Scroll()
   ' AA1 = GRID2.LeftCol
   If HAYTOTA% = 1 Then
     Call MUESTRATOTS
   End If
End Sub

Private Sub Image1_Click()

    Call EXPORTEXCEL
    
End Sub

Private Sub Image2_Click()

    Call FULLSCREEN

End Sub


Private Sub Image3_Click()

    Call AUTOADJUST
    
End Sub


Private Sub Image4_Click()

    Call Search

End Sub


Private Sub Image5_Click()

    Call IMPRIMIR

End Sub


Private Sub Label3_DblClick()
    MODICONSU% = 1
    Hide
End Sub

Private Sub Label2_Change()
    If TRIM$(Label2.Caption) <> "" Then
        GRID2.Clear
        GRID2.Cols = 1: CANCOLS% = 0
        GRID2.Rows = 1: CANFILS% = 0
        NUPIC% = Label4.Count
        For KKI% = 1 To NUPIC%
           Label4(KKI% - 1) = ""
           Label4(KKI% - 1).Visible = False
        Next KKI%
        '
        If Label2.Caption = "@NEW" Then
           FGENPLA07.Label1.Caption = ""
           FGENPLA07.Show 1
           Exit Sub
        End If
        Call FILTROVAR
      Else
        GRID2.Clear
        For i& = 1 To Label1.Count
            Label1(i& - 1).Visible = False
            Combo1(i& - 1).Visible = False
            Command2(i& - 1).Visible = False
        Next i&
        Frame1.Visible = False
    End If
End Sub

Sub FILTROVAR()
    '
    CACONFIL% = 0
    ULTIFIL& = 1: ULTIENT& = 0
    Frame1.Width = 2000
    Frame1.Visible = True
    '
    For i& = 1 To Label1.Count
        Label1(i& - 1).Visible = False
        Combo1(i& - 1).Visible = False
        Command2(i& - 1).Visible = False
        Command2(i& - 1).Width = 330
        FilText1(i& - 1).Visible = False
        FilText1(i& - 1).ToolTipText = ""
        FiltEco1(i& - 1).Visible = False
    Next i&
    
    If CACONDI% < 1 Then
       Frame1.Visible = False
       Frame1.Height = 0
       GoTo 90
       'Hide
       Exit Sub
    End If
    '
    PVER = 600
    JK& = 0
    TAMAXI = 0
    For i& = 1 To CACONDI%
       VAREFE$ = VAREFCOND$(i&)
       VALORCOND$(i&) = VAREFE$
       IZX$ = "": If Left$(VAREFE$, 1) = "'" Then IZX$ = "'": VAREFE$ = Mid$(VAREFE$, 2)
       DZX$ = "": If Right$(VAREFE$, 1) = "'" Then DZX$ = "'": VAREFE$ = Left$(VAREFE$, Len(VAREFE$) - 1)
10     If Left$(VAREFE$, 1) = "¿" Then
          If Right$(VAREFE$, 1) = "?" Then
             LABPRE$ = Mid$(VAREFE$, 2, Len(VAREFE$) - 2)
             JK& = JK& + 1
             '
             ASELEC$ = ""
             PPXX% = InStr(LABPRE$, "/")
             '
             If PPXX% > 1 Then
                ASELEC$ = Mid$(LABPRE$, PPXX% + 1)
                LABPRE$ = Left$(LABPRE$, PPXX% - 1)
             End If
             If TRIM$(UCase$(ASELEC$)) = "FECHA" Then ASELEC$ = "SELFECHA"
             '
20           If i& > Label1.Count Then
                NUPIC% = Label1.Count
                Load Label1(NUPIC%)
                Load Combo1(NUPIC%)
                Load Command2(NUPIC%)
                Load FilText1(NUPIC%)
                Load FiltEco1(NUPIC%)
                GoTo 20
             End If
             Label1(JK& - 1).Top = PVER + 400 * (JK& - 1)
             Combo1(JK& - 1).Top = PVER - 50 + 400 * (JK& - 1)
             Command2(JK& - 1).Top = PVER - 50 + 400 * (JK& - 1)
             Label1(JK& - 1) = LABPRE$
             Combo1(JK& - 1).Clear
             If Label1(JK& - 1).Width > TAMAXI Then TAMAXI = Label1(JK& - 1).Width
             Label1(JK& - 1).Visible = True
             Combo1(JK& - 1).Visible = True
             Command2(JK& - 1).Visible = True
             FilText1(JK& - 1).Visible = False
             FilText1(JK& - 1).ToolTipText = ASELEC$
             FiltEco1(JK& - 1).Visible = False
             CAMPOSEL$(JK& - 1) = CAMPOCOND$(i&)
             ULTIFIL& = JK&
             '
             If APLINVO$ <> "" Then
               If FILTX$ <> "" Then
               
                 PPXX% = InStr(FILTX$, ";")
                 If PPXX% < 1 Then PPXX% = InStr(FILTX$, ")")
                 If PPXX% < 1 Then PPXX% = 1 + Len(FILTX$)
                 If PPXX% > 1 Then
                    Combo1(JK& - 1).TEXT = Left$(FILTX$, PPXX% - 1)
                    FilText1(JK& - 1).TEXT = Left$(FILTX$, PPXX% - 1)
                    FILTX$ = Mid$(FILTX$, PPXX% + 1)
                 End If
               End If
             End If
             '
          End If
       End If
    Next i&
    '
    CACONFIL% = JK&
    If CACONFIL% < 1 Then
        Frame1.Visible = False
        Frame1.Height = 0
        GoTo 90
        Exit Sub
    End If
    '
    For i& = 1 To CACONFIL%
30     If i& > Label1.Count Then
          NUPIC% = Label1.Count
          Load Label1(NUPIC%)
          Load Combo1(NUPIC%)
          Load Command2(NUPIC%)
          Load FilText1(NUPIC%)
          Load FiltEco1(NUPIC%)
          GoTo 30
       End If
       '
       Label1(i& - 1).Left = 500 + TAMAXI - Label1(i& - 1).Width
       Combo1(i& - 1).Left = Label1(i& - 1).Left + Label1(i& - 1).Width + 250
       Command2(i& - 1).Left = Combo1(i& - 1).Left + Combo1(i& - 1).Width + 150
       
       FilText1(i& - 1).Left = Label1(i& - 1).Left + Label1(i& - 1).Width + 250
       FiltEco1(i& - 1).Left = Command2(i& - 1).Left + Command2(i& - 1).Width + 100
       FiltEco1(i& - 1).Top = FilText1(i& - 1).Top
       '
       If FilText1(i& - 1).ToolTipText <> "" Then
            Combo1(i& - 1).Visible = False
            FilText1(i& - 1).Visible = True
            ASELEC$ = FilText1(i& - 1).ToolTipText
            FiltEco1(i& - 1).Visible = False
            If EXISTE%(LUDAT$ + ASELEC$ + ".RFS") > 0 Then
               FiltEco1(i& - 1).Visible = True
            End If
          Else
            Combo1(i& - 1).Visible = True
            FilText1(i& - 1).Visible = False
            FiltEco1(i& - 1).Visible = False
       End If
       Call AJUSTANCHO(i&)
       '
    Next i&
    '
    ' ACOMODAR LOS FILTROS
    For i& = 2 To CACONFIL%
       If CAMPOSEL$(i& - 1) = CAMPOSEL$(i& - 2) Then
          If FilText1(i& - 1).ToolTipText <> "" Then
               EXTRE = FilText1(i& - 2).Left + FilText1(i& - 2).Width
               If FiltEco1(i& - 2).Left + FiltEco1(i& - 2).Width > EXTRE Then EXTRE = FiltEco1(i& - 2).Left + FiltEco1(i& - 2).Width
               ANECE = FilText1(i& - 1).Width + FiltEco1(i& - 1).Width
             Else
               EXTRE = Combo1(i& - 2).Left + Combo1(i& - 2).Width
               ANECE = Combo1(i& - 1).Width
          End If
          If Command2(i& - 2).Left + Command2(i& - 2).Width > EXTRE Then EXTRE = Command2(i& - 2).Left + Command2(i& - 2).Width
          ANECE = EXTRE + 500 + Label1(i& - 1).Width + ANECE + Command2(i& - 1).Width
          If ANECE < Width - 2 * Frame1.Left - Command1.Width - 200 Then
              Label1(i& - 1).Top = Label1(i& - 2).Top
              Label1(i& - 1).Left = 500 + EXTRE
              DIFLAB = (Label1(i& - 1).Left + Label1(i& - 1).Width) - (Label1(i& - 2).Left + Label1(i& - 2).Width)
              Combo1(i& - 1).Top = Combo1(i& - 2).Top
              Combo1(i& - 1).Left = Combo1(i& - 2).Left + DIFLAB
              FilText1(i& - 1).Top = FilText1(i& - 2).Top
              FilText1(i& - 1).Left = FilText1(i& - 2).Left + DIFLAB
              Command2(i& - 1).Top = Command2(i& - 2).Top
              Command2(i& - 1).Left = Command2(i& - 2).Left + DIFLAB
              Call AJUSTANCHO(i&)
          End If
       End If
    Next i&
    '
    'Frame1.Width = Command2(ULTIFIL& - 1).Left + Command2(ULTIFIL& - 1).Width + 450
    Frame1.Height = Command2(ULTIFIL& - 1).Top + 600
    '
    Command1.Top = 130
    Command1.Height = Frame1.Height - 150
    Command1.Left = Frame1.Width - Command1.Width
    '
90  GRID2.Top = Frame1.Top + Frame1.Height + 250
    GRID2.Height = Height - GRID2.Top - 500
    '
    If CACONFIL% < 1 Then
       Call Command1_Click
       Exit Sub
    End If
    '
    If APLINVO$ <> "" Then
      For i& = 1 To ULTIFIL&
        If Combo1(i& - 1).Visible = True Then
          If TRIM$(Combo1(i& - 1).TEXT) = "" Then Exit Sub
        End If
      Next i&
      Call Command1_Click
      Exit Sub
    End If
End Sub

Sub AJUSTANCHO(i&)
   '
   TASELE$ = TRIM$(UCase$(FilText1(i& - 1).ToolTipText))
   If TASELE$ = "MASTER" Then TASELE$ = "ECOMAST"
   If TASELE$ = "FECHA" Then TASELE$ = "SELFECHA"
   FiltEco1(i& - 1).Width = 0
   If Combo1(i& - 1).Visible = True Or TASELE$ = "" Then GoTo 80
   '
   QTR& = CantRegistros("SQL_LISELEC", "NomLiSele = '" & TASELE$ & "'")
   If QTR& > 0 Then                            ' TOMA DE LA BASE DE DATOS
        ACOLU1% = XVALO(VALOBADA("SQL_LISELEC", "Acol_1", "NomLiSele = '" & TASELE$ & "'"))
        ACOLU2% = XVALO(VALOBADA("SQL_LISELEC", "Acol_2", "NomLiSele = '" & TASELE$ & "'"))
     Else
        TIVA1% = TIPFIELD%(TASELE$, 1)
        LOVA1% = LENFIELD%(TASELE$, 1)
        If TIVA1% = 1 Or TIVA1% = 2 Then
            If LOVA1% < 2 Then ACOLU1% = 3 Else ACOLU1% = 5
          ElseIf TIVA1% = 3 Or TIVA1% = 4 Then
            If LOVA1% < 8 Then ACOLU1% = 9 Else ACOLU1% = 14
          ElseIf TIVA1% = 5 Then
            ACOLU1% = 8
          ElseIf TIVA1% = 7 Then
            ACOLU1% = 16
          Else
            ACOLU1% = LOVA1%
        End If
        '
        TIVA2% = TIPFIELD%(TASELE$, 2)
        LOVA2% = LENFIELD%(TASELE$, 2)
        If TIVA2% = 1 Or TIVA2% = 2 Then
            If LOVA2% < 2 Then ACOLU2% = 3 Else ACOLU2% = 5
          ElseIf TIVA2% = 3 Or TIVA2% = 4 Then
            If LOVA2% < 8 Then ACOLU2% = 9 Else ACOLU2% = 14
          ElseIf TIVA2% = 5 Then
            ACOLU2% = 8
          ElseIf TIVA2% = 7 Then
            ACOLU2% = 16
          Else
            ACOLU2% = LOVA2%
        End If
   End If
   '
   Teprueba.Font = FilText1(i& - 1).Font
   Teprueba.FontSize = FilText1(i& - 1).FontSize
   Teprueba.FontBold = FilText1(i& - 1).FontBold
   Teprueba.FontItalic = FilText1(i& - 1).FontItalic
   Teprueba.Caption = String$(ACOLU1%, "O")
   FilText1(i& - 1).FontSize = 1
   FilText1(i& - 1).Width = Teprueba.Width
   FilText1(i& - 1).FontSize = Teprueba.FontSize
   '
   FilText1(i& - 1).Top = Combo1(i& - 1).Top
   Command2(i& - 1).Width = 175
   Command2(i& - 1).Left = FilText1(i& - 1).Left + FilText1(i& - 1).Width
   '
   Teprueba.Font = FiltEco1(i& - 1).Font
   Teprueba.FontSize = FiltEco1(i& - 1).FontSize
   Teprueba.FontBold = FiltEco1(i& - 1).FontBold
   Teprueba.FontItalic = FiltEco1(i& - 1).FontItalic
   Teprueba.Caption = String$(ACOLU2%, "O")
   FiltEco1(i& - 1).FontSize = 1
   FiltEco1(i& - 1).Width = Teprueba.Width
   FiltEco1(i& - 1).FontSize = Teprueba.FontSize
   '
   FiltEco1(i& - 1).Left = Command2(i& - 1).Left + Command2(i& - 1).Width + 100
   FiltEco1(i& - 1).Top = FilText1(i& - 1).Top
   '
   If FiltEco1(i& - 1).Width + FiltEco1(i& - 1).Left + Frame1.Left > ScaleWidth - 600 Then
      FiltEco1(i& - 1).Width = ScaleWidth - FiltEco1(i& - 1).Left - Frame1.Left - 600
   End If
   '
80 AMAX0 = 0: AMAX1 = 0: AMAX2 = 0
   If TASELE$ = "" Then
         AMAX0 = Command2(i& - 1).Left + Command2(i& - 1).Width + 150 + Command1.Width
      Else
         AMAX1 = Command2(i& - 1).Left + Command2(i& - 1).Width + 150 + Command1.Width
         If FiltEco1(i& - 1).Width > 0 Then
            AMAX2 = FiltEco1(i& - 1).Left + FiltEco1(i& - 1).Width + 150 + Command1.Width
         End If
   End If
   If AMAX1 > AMAX0 Then AMAX0 = AMAX1
   If AMAX2 > AMAX0 Then AMAX0 = AMAX2
   If Frame1.Width < AMAX0 Then
      Frame1.Width = AMAX0
      Command1.Left = Frame1.Width - Command1.Width
      Frame1.Refresh
   End If
   '
End Sub


Private Sub Timer2_Timer()
   Static ACOLGRI(256)
   For KI% = 0 To GRID2.Cols - 1
      If ACOLGRI(KI%) < 1 Then ACOLGRI(KI%) = GRID2.ColWidth(KI%)
      If GRID2.ColWidth(KI%) <> ACOLGRI(KI%) Then
         Call MUESTRATOTS
         ACOLGRI(KI%) = GRID2.ColWidth(KI%)
      End If
   Next KI%
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   Select Case UCase$(Button.Key)
      Case "PRINT"
        Call IMPRIMIR
      Case "SEARCH"
        Call Search
      Case "EXIT"
        'YAENTRO% = 0
        'GRID.Clear
        Hide
   End Select
   '
End Sub

Sub IMPRIMIR()
   If XVALO(Label7) > 0 Then
      Call PRINREP07.CARGALIMPRE
      PRINREP07.Show 1
   End If
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)
   Select Case UCase$(Button.Key)
      Case "EXCEL"
        Call EXPORTEXCEL
      Case "GRAFIC"
        Call GENERGRAF
      Case "FULLSCREEN"
        Call FULLSCREEN
   End Select
End Sub

Sub GENERGRAF()
    '
    GRAMEN07.Show 1
    '
End Sub

Sub EXPORTEXCEL()
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
      Call GENEXCEL(AREXPO$)
   End If
   '
99 End Sub
'
Sub GENEXCEL(NOMBREDEARCHIVO$)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   For i& = 1 To GRID2.Rows
       RENG1& = i&
       For COLU1& = 1 To GRID2.Cols
          VALORX = GRID2.TextMatrix(i& - 1, COLU1 - 1)
          XLSHEET.Cells(RENG1&, COLU1&) = VALORX
       Next COLU1&
   Next i&
   '
   Close #N1%
   On Error GoTo 30
   XLSHEET.SaveAs NOMBREDEARCHIVO$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NOMBREDEARCHIVO$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub

Private Sub GRID2_DblClick()
    If GRID2.Row <= 1 Then
       Screen.MousePointer = 11
       '
       COLUSOR& = GRID2.COL + 1
       If SENTIOR%(COLUSOR&) < 1 Then
            SENTIOR%(COLUSOR&) = 1    ' ASCENDENTE
          Else
            SENTIOR%(COLUSOR&) = (-1) ' DESCENDENTE
       End If
       '
       Dim VALORES
       Dim ARTEMP
       URETRA& = XVALO(Label7)
       FILAS& = URETRA&
       COLUS& = GRID2.Cols
       ReDim VALORES(FILAS&, COLUS&)
       ReDim ARTEMP(COLUS&)
       ESPORFECHA% = 0
       For i& = 1 To FILAS&
         For J& = 1 To COLUS&
           VALORES(i&, J&) = GRID2.TextMatrix(i&, J& - 1)
           If J& = COLUSOR& Then
             VLX$ = VALORES(i&, J&)
             If Len(VLX$) = 8 Then
               If Mid$(VLX$, 3, 1) = "/" And Mid$(VLX$, 6, 1) = "/" Then
                 ESPORFECHA% = 1
               End If
             End If
           End If
         Next J&
       Next i&
       If ESPORFECHA% = 1 Then
         For i& = 1 To FILAS&
           VLX$ = VALORES(i&, COLUSOR&)
           VALORES(i&, COLUSOR&) = FECHANUM(VLX$)
         Next i&
       End If
       '
       ' ORDENAMIENTO SHELL
       SIG% = SENTIOR%(COLUSOR&)
       INCRE& = FILAS& \ 2
       Do Until INCRE& < 1
         For UI& = INCRE& + 1 To FILAS&
            For KKI& = 1 To COLUS&: ARTEMP(KKI&) = VALORES(UI&, KKI&): Next KKI&
            For UJ& = UI& - INCRE& To 1 Step -INCRE&
              If SIG% > 0 Then
                  If ARTEMP(COLUSOR&) >= VALORES(UJ&, COLUSOR&) Then Exit For
                 Else
                  If ARTEMP(COLUSOR&) <= VALORES(UJ&, COLUSOR&) Then Exit For
              End If
              '
              For KKI& = 1 To COLUS&: VALORES(UJ& + INCRE&, KKI&) = VALORES(UJ&, KKI&): Next KKI&
            Next UJ&
            For KKI& = 1 To COLUS&:  VALORES(UJ& + INCRE&, KKI&) = ARTEMP(KKI&): Next KKI&
         Next UI&
         INCRE& = INCRE& \ 2
       Loop
       ' FIN ORDENAMIENTO SHELL
       '
       For i& = 1 To FILAS&
         For J& = 1 To COLUS&
           If J& = COLUSOR& Then
             If ESPORFECHA% = 1 Then
               VLY = VALORES(i&, J&)
               GRID.TextMatrix(i&, J& - 1) = FECHATEX$(VLY)
               GoTo 29
             End If
           End If
           '
           GRID2.TextMatrix(i&, J& - 1) = VALORES(i&, J&)
29       Next J&
       Next i&
       Screen.MousePointer = 1
       '
     Else
       If GRID2.COL <= 1 Then
         TTXX$ = GRID2.TEXT
         For KKI1% = 1 To GRAMEN07.Combo4.Count
           VALOSE$ = UCase$(TRIM$(GRAMEN07.Combo4(KKI1% - 1).TEXT))
           If VALOSE$ = "" Or VALOSE$ = UCase$(TRIM$(TTXX$)) Then
             GRAMEN07.Combo4(KKI1% - 1).TEXT = TTXX$
             Exit For
           End If
         Next KKI1%
       End If
    End If
End Sub
'
Sub FULLSCREEN()
    '
    If SHOWPLA07.WindowState = 0 Then
     SHOWPLA07.WindowState = 2
    Else
     SHOWPLA07.WindowState = 0
    End If
    '
    GRID2.Width = SHOWPLA07.Width - 300
    GRID2.Height = (SHOWPLA07.Height - GRID2.Top) - 500
    Toolbar1.Left = SHOWPLA07.Width - Toolbar1.Width - 50
    MARCOBARRA.Left = Toolbar1.Left + (Toolbar1.Width - MARCOBARRA.Width) / 2
    Toolbar2.Left = Toolbar1.Left - Toolbar2.Width - 50
    Picture1.Left = SHOWPLA07.Width - Picture1.Width - 50
    '
    Label9.Width = Me.Width
    '
    If HAYTOTA% = 1 Then
       GRID2.Height = ScaleHeight - GRID2.Top - 500
       Call MUESTRATOTS
    End If
    
'    If SHOWPLA07.WindowState = 0 Then
'        SHOWPLA07.WindowState = 2
'      Else
'        SHOWPLA07.WindowState = 0
'    End If
'    '
'    GRID2.Width = SHOWPLA07.Width - 300
'    GRID2.Height = (SHOWPLA07.Height - GRID2.Top) - 700
'    Toolbar1.Left = SHOWPLA07.Width - 1960
'    Toolbar2.Left = (SHOWPLA07.Width - 1960) - 1500
'    '
'    If HAYTOTA% = 1 Then
'       GRID2.Height = Height - 300 - GRID2.Top - 700
'       Call MUESTRATOTS
'    End If

    '
End Sub

Sub Search()

   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", , ARBUS$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     For U& = 0 To GRID2.Rows - 1
      For i& = 0 To GRID2.Cols - 1
       If InStr(UCase$(GRID2.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
        GRID2.Row = U&: GRID2.COL = i&
        If GRID2.TopRow > GRID2.Row Then GRID2.TopRow = GRID2.Row
        On Error Resume Next ' PARA EVITAR ERROR DIVISION POR 0
        VISIROWS = (GRID2.Height / GRID2.RowHeight(J&)) - 4
        If Err > 0 Then
         VISIROWS = 1
        End If
        On Error GoTo 0
         If VISIROWS > 0 Then
            If GRID2.TopRow < GRID2.Row - VISIROWS Then
              GRID2.TopRow = GRID2.Row - VISIROWS
            End If
         End If
        GRID2.SetFocus
        Exit Sub
       End If
       Next i&
     Next U&
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If

End Sub

Sub MUESTRATOTS()
    '
    PPX& = GRID2.Left + GRID2.ColPos(0)
    For KI& = 1 To GRID2.Cols
       If KI& > Label4.Count Then
          NUPIC% = Label4.Count
          Load Label4(NUPIC%)
       End If
       Label4(KI& - 1).Caption = ""
       Label4(KI& - 1).Visible = False
       Label4(KI& - 1).Top = GRID2.Top + GRID2.Height + 100
       Label4(KI& - 1).Left = PPX& + 30
       Label4(KI& - 1).Width = GRID2.ColWidth(KI& - 1)
       If TIESUM%(KI&) = 1 Then
          Label4(KI& - 1).Visible = True
          Label4(KI& - 1).Caption = TRIM$(FORMATNUM$(SUMA#(KI&), FORCAMPO$(COLUVALOR + 1)))
          Label4(KI& - 1).Refresh
       End If
       PPX& = PPX& + GRID2.ColWidth(KI& - 1)
    Next KI&
    '
End Sub

Sub AUTOADJUST()
   '
   Teprueba.Font = GRID2.Font
   Teprueba.FontSize = GRID2.FontSize
   Teprueba.FontBold = GRID2.FontBold
   '
   For COLUI& = 1 To GRID2.Cols
      ACOLU = 0
      For RENGI& = 1 To GRID2.Rows
         Teprueba.Caption = TRIM$(GRID2.TextMatrix(RENGI& - 1, COLUI& - 1)) + "."
         If Teprueba.Width > ACOLU Then ACOLU = Teprueba.Width
      Next RENGI&
      GRID2.ColWidth(COLUI& - 1) = ACOLU
   Next COLUI&
   '
End Sub
