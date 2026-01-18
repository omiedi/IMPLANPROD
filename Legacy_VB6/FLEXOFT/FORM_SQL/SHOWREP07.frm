VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.ocx"
Begin VB.Form SHOWREP07 
   BackColor       =   &H00DAE4EB&
   BorderStyle     =   1  'Fixed Single
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
      Left            =   10005
      TabIndex        =   10
      Top             =   600
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
      Left            =   7800
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   873
      ButtonWidth     =   794
      ButtonHeight    =   714
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
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
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "AutoAdjust"
            Object.ToolTipText     =   "Auto Ajustar Ancho de Columnas"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Excel"
      Height          =   255
      Left            =   6960
      TabIndex        =   27
      ToolTipText     =   "Exportador Ràpido a Excel"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRID 
      Height          =   6210
      Left            =   120
      TabIndex        =   25
      Top             =   2160
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   10954
      _Version        =   393216
      Rows            =   28
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   14737632
      ForeColorSel    =   12582912
      GridColor       =   12632256
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
      TabIndex        =   24
      Top             =   30
      Width           =   3135
      Begin VB.Image Image5 
         Height          =   660
         Left            =   2280
         Picture         =   "SHOWREP07.frx":0000
         ToolTipText     =   "Imprime Reporte"
         Top             =   0
         Width           =   645
      End
      Begin VB.Image Image4 
         Height          =   570
         Left            =   1800
         Picture         =   "SHOWREP07.frx":16F2
         ToolTipText     =   "Buscar"
         Top             =   0
         Width           =   465
      End
      Begin VB.Image Image3 
         Height          =   570
         Left            =   1200
         Picture         =   "SHOWREP07.frx":2574
         ToolTipText     =   "Ajuste de Columnas"
         Top             =   0
         Width           =   585
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   660
         Picture         =   "SHOWREP07.frx":3786
         ToolTipText     =   "Maximiza / Minimiza Pantalla"
         Top             =   0
         Width           =   570
      End
      Begin VB.Image Image1 
         Height          =   525
         Left            =   120
         Picture         =   "SHOWREP07.frx":46BC
         ToolTipText     =   "Exporta Reporte a Excel"
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      Appearance      =   0  'Flat
      BackColor       =   &H002B9973&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   10280
      ScaleHeight     =   165
      ScaleWidth      =   1455
      TabIndex        =   22
      Top             =   680
      Width           =   1485
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H0048CC9D&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   -120
         TabIndex        =   23
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   5400
      Top             =   3960
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   1275
      TabIndex        =   19
      ToolTipText     =   "Click con Boton Derecho para Cambiar la Imagen"
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   5355
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6960
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DAE4EB&
      Caption         =   "CONDICIONES DE SELECCIÓN"
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
      Left            =   105
      TabIndex        =   4
      Top             =   840
      Width           =   5010
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
         Height          =   945
         Left            =   4320
         Picture         =   "SHOWREP07.frx":55C2
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
         TabIndex        =   17
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
         Index           =   0
         Left            =   3720
         TabIndex        =   6
         ToolTipText     =   "Lista de Valores Posibles"
         Top             =   420
         Visible         =   0   'False
         Width           =   210
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
         TabIndex        =   18
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
      ItemData        =   "SHOWREP07.frx":5C66
      Left            =   8820
      List            =   "SHOWREP07.frx":5C68
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
      ItemData        =   "SHOWREP07.frx":5C6A
      Left            =   6090
      List            =   "SHOWREP07.frx":5C6C
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
      ItemData        =   "SHOWREP07.frx":5C6E
      Left            =   840
      List            =   "SHOWREP07.frx":5C70
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRID2 
      Height          =   6210
      Left            =   120
      TabIndex        =   26
      Top             =   2160
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   10954
      _Version        =   393216
      Rows            =   28
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   5197685
      ForeColorSel    =   12582912
      GridColor       =   12632256
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Label Teprueba 
      AutoSize        =   -1  'True
      BackColor       =   &H00DAE4EB&
      Caption         =   "TEPRUEBA"
      Height          =   240
      Left            =   6240
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   1245
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8820
      Top             =   735
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483644
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":5C72
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":D2B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":DC62
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":DF7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":E296
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
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      Left            =   9450
      Top             =   735
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
            Picture         =   "SHOWREP07.frx":E5B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":E8CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":EBE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":EEFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":F218
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":F532
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":F84C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":FB66
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":FE80
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":1019A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":104B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":1068E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":107A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SHOWREP07.frx":10ABA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   7920
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DAE4EB&
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
      TabIndex        =   21
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "SHOWREP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VEMOFREN%               ' VECES QUE SE MOVIO AL FRENTE
Dim CANCELA%
Dim MULTILIN%
Dim SENTIOR%(256)
Dim ARBUS$, ABUSQUE$
'
Dim RS As ADODB.Recordset
'
Dim HAYTOTA%
Dim AMAXI%(256), ACAMPO$(256)
Dim SUMA#(256), SUBTOT#(256)
Dim TIESUM%(256), FORSUMA$(256)
Public NO_ORDENAR%
'
Dim RWW&

Dim ReporteConFiltro%, ReporteConsultaSql$

Function ARMA_ENCABEZADO() As String

    Dim tituloColumna As String
    Dim formatoColumna As String
    Dim cantidadColumnasEco As Integer
    
    ARMA_ENCABEZADO = ""
    For i = 1 To UBound(TITCAMPO)
        tituloColumna = TITCAMPO(i)
        formatoColumna = FORCAMPO(i)
        If TRIM$(tituloColumna) <> "" And TRIM$(formatoColumna) <> "" Then
            ARMA_ENCABEZADO = ARMA_ENCABEZADO & tituloColumna & "/" & formatoColumna & ";"
            TECO$ = Left$(TRIM$(UCase$(TABLAECO$(i))), 5)
            If TECO$ <> "" Then
                cantidadColumnasEco = cantidadColumnasEco + 1
                tituloColumna = GRID.TextMatrix(0, i + cantidadColumnasEco - 1)
                formatoColumna = TRIM$(UCase$(FORMAECO$(i)))
                
                ARMA_ENCABEZADO = ARMA_ENCABEZADO & tituloColumna & "/" & formatoColumna & ";"
            End If
        End If
    Next i
    If Len(ARMA_ENCABEZADO) > 1 Then
        ARMA_ENCABEZADO = Mid$(ARMA_ENCABEZADO, 1, Len(ARMA_ENCABEZADO) - 1)
    End If

    
End Function

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
               If CODINT%(TREPLA$) > 0 And InStr(1, "codint;cod_inte", TRIM$(CAMPOSEL$(i& - 1)), vbTextCompare) > 0 Then
                  TREP$ = TRIM$(Str$(CODINT%(TREPLA$)))
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
                  TREP$ = TRIM$(Str$(CODINT%(TREPLA$)))
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
End Sub
'

Sub EJECUTACONSULTA()
    '
    For i& = 1 To UBound(TABLAECO$)
      EXO$ = TRIM$(UCase$(TABLAECO$(i&)))
      If EXO$ <> "" Then
          EXOTRUNCO$ = Left$(TRIM$(UCase$(TABLAECO$(i&))), 5)
          If EXOTRUNCO$ = "MASTE" Then
             Call VectArticulos.cargarvector
          ElseIf EXOTRUNCO$ = "DEUDO" Or EXOTRUNCO$ = "CLIEN" Then
            Call VectCliente.cargarvector
          ElseIf EXOTRUNCO$ = "PROVE" Or EXOTRUNCO$ = "ACRED" Then
            Call VectProveedor.cargarvector
          Else
            Call VectTablasDat.CarVectGeneral(EXO$)
         End If
      End If
    Next i&
    
    Dim VECTORSTATUS$(10)
    VECTORSTATUS$(0) = ""
    VECTORSTATUS$(1) = "ACTIVO"
    VECTORSTATUS$(3) = "CERRADO"
    VECTORSTATUS$(8) = "ANULADO"
    VECTORSTATUS$(9) = "ANULADO"
    
    For KKI% = 0 To 256
       AMAXI%(KKI%) = 0
       SUMA#(KKI%) = 0
       SUBTOT#(KKI%) = 0
       TIESUM%(KKI%) = 0
       FORSUMA$(KKI%) = ""
       ACAMPO$(KKI%) = ""
    Next KKI%
    Label7 = "0"
    '
    VUELTANUMERO% = 0
10  Call ARMA_CONSULTA(VUELTANUMERO%)
    Call REPLA_CONSULTA
    ' MAS ADELANTE SE EJECUTA LA CONSULTA CONTENIDA EN 'TXT_CONSULTA' CON VALOBADAS INCLUIDOS. EL SQL NO ARROJA ERROR EN UNA CONSULTA SIMPLE
    ' AUNQUE LOS VALOBADAS NO EXISTAN COMO CAMPOS DE LA TABLA, PERO SIN EMBARGO FALLA CUANDO LA CONSULTA INCLUYE AL OPERADOR
    ' 'UNION ALL'. PARA EVITAR LO ANTERIOR SE REEMPLAZAN DICHOS VALOBADA POR EL CAMPO CODIEMPR QUE SE ENCUENTRA EN TODAS LAS TABLAS
    ' Y NO AFECTA A LA OBTENCION DEL DATO YA QUE LAS CONDICIONES PARA LOS VALOBADA ESTAN EN VECTORES (VER MAS ABAJO)
    For idValobada& = 1 To 20 ' SE ASUME UN MAXIMO DE 10 VALOBADA
        Call REPLATEXT(TXT_CONSULTA, "VALOBADA_" & CStr(idValobada&), "CodiEmpr")
    Next idValobada&
    '
    Set RS = New ADODB.Recordset
    '
    Screen.MousePointer = 11
    On Error Resume Next
    Set RS = FLEXCONN.Execute(FILTSQL$(TXT_CAREGIST$))
    If Err Then
       On Error GoTo 0
       FIN& = 1
       If APLINVO$ = "" Then Call MENSERR(24, "Imposible Determinar Cantidad de Registros")
       GoTo 20       '  99
    End If
    On Error GoTo 0
'    Set RS = READSET(TXT_CAREGIST$)

    If VUELTANUMERO% > 0 And CONSIDERAR_MOVIVIE% > 0 Then
        FIN& = FIN& + RS!CANREGIS
    Else
        ' SE INCLUYE ON ERROR PORQUE AL TRABAJAR CON LA CLAUSULA 'UNION ALL'
        ' NO EXISTE 'CANREGIS' Y POR ENDE SE PRODUCE UN ERROR. PASO SIGUIENTE
        ' SE USA EL INDICE CERO PARA OBTENER EL RESULTADO DEL COUNT
        On Error Resume Next
        FIN& = RS!CANREGIS
        If Err Then FIN& = RS(0)
        On Error GoTo 0
    End If
    '
20  Set RS = New ADODB.Recordset
'    Set RS = READSET(TXT_CONSULTA)
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
    HAYTOTA% = 0: MULTILIN% = 0: FORMATX% = 0
    If VUELTANUMERO% < 1 And CONSIDERAR_MOVIVIE% < 1 Then
      GRID.Clear
    End If
    If VUELTANUMERO% > 0 And CONSIDERAR_MOVIVIE% > 0 Then GoTo 30

    On Error Resume Next
    GRID.Rows = FIN& + 1
    If Err Then
       On Error GoTo 0
       ' suprimido - se captura en ejecucion
       'Call MENSERR(24, "Memoria Insuficiente para Ejecutar Consulta.")
       'GoTo 99
    End If
    On Error GoTo 0
    '
    GRID.Cols = RS.Fields.Count
    JI& = 0
    For i& = 1 To RS.Fields.Count
      If CANCELA% > 0 Then GoTo 99
      JI& = JI& + 1
      If JI& > GRID.Cols Then
         On Error Resume Next
         GRID.Cols = JI&
         If Err Then
            ERDE$ = Err.Number & " - " & Err.Description
            On Error GoTo 0
            MsgBox "Error Fatal: " + ERDE$
            Call FINAL("")
         End If
         On Error GoTo 0
      End If
      '

      GRID.COL = JI& - 1: ACAMPO$(JI&) = SELCAMPO$(i&)
      AMAXI%(JI&) = 0
      GRID.ColAlignment(JI& - 1) = 1
      ENCABE$ = TITCAMPO$(i&)
      If Left$(ENCABE$, 1) = ">" Then
            ENCABE$ = Mid$(ENCABE$, 2)
            GRID.ColAlignment(JI& - 1) = 7      ' derecha y centrado vertical
         ElseIf Left$(ENCABE$, 1) = "<" Then
            ENCABE$ = Mid$(ENCABE$, 2)
      End If
      '
      If TOTCAMPO$(i&) = "1" Then
         HAYTOTA% = 1
         TIESUM%(JI&) = 1
         FORSUMA$(JI&) = FORCAMPO$(i&)
      End If
      '
      If FORCAMPO$(i&) = "" Then
         TICAMPO% = RS.Fields(i& - 1).Type
         If TICAMPO% = 2 Then
              FORCAMPO$(i&) = "F7.0"             ' INTEGER
           ElseIf TICAMPO% = 3 Then
              FORCAMPO$(i&) = "F7.0"             ' INTEGER
           ElseIf TICAMPO% = 5 Then
              FORCAMPO$(i&) = "F12.2"            ' NUMERICO
           ElseIf TICAMPO% = 202 Then
              FORCAMPO$(i&) = "A64"              ' TEXTO
           ElseIf TICAMPO% = 135 Then
              FORCAMPO$(i&) = "D8"               ' FECHA
           Else
              MsgBox "Falta Formato para Tipo de Dato " & TICAMPO%
              FORCAMPO$(i&) = "A16"
         End If
      End If
      '
      ' ESTO ES PARA QUE EN CAMPOS NUMERICOS EL ENCABEZADO APAREZCA ALINEADO A DERECHA
      FFCC$ = UCase$(Left$(FORCAMPO$(i&), 1))
      If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
          GRID.ColAlignment(JI& - 1) = 6      ' derecha y centrado vertical
          GRID.Row = 0: GRID.COL = JI& - 1: GRID.CellAlignment = 6
      End If
      '
      Label5.Font = GRID.Font
      Label5.FontName = GRID.CellFontName
      Label5.FontSize = GRID.CellFontSize
      Label5.FontBold = GRID.CellFontBold
         FCPXX$ = REPLACAR$(FORCAMPO$(i&), "X", "A")
         FCPYY$ = Left$(FCPXX$, 3)
         If FCPYY$ = "JPG" Or FCPYY$ = "BMP" Or FCPYY$ = "GIF" Then   ' agregado para prever formato grafico
            FCPXX$ = "A" + Mid$(FCPXX$, 4)                            ' EPB 16/01/2011
         End If
      'Label5.Caption = TEXTBIN$("1.234567", FCPXX$) + "*"
      Label5.Caption = FORMATNUM$("9999999.9999999", FCPXX$) + "*"    ' PROVISORIAMENTE A VER SI FUNCIONA
      GRID.ColWidth(JI& - 1) = Label5.Width
      If OCUCAMPO$(i&) = "1" Then
         GRID.ColWidth(JI& - 1) = 0
      End If
      '
15    If JI& >= Label8.Count Then
         NUPIC% = Label8.Count
         Load Label8(NUPIC%)
         GoTo 15
      End If
      Label8(JI&) = Str$(Len(Label5) - 1)
      '
      If VUELTANUMERO% < 1 Then GRID.Row = 0
      GRID.TEXT = ENCABE$
      '
      ' ENCABEZADO DE COLUMNA ECO
      TECO$ = Left$(TRIM$(UCase$(TABLAECO$(i&))), 5)
      If TECO$ <> "" Then
        FECO1$ = TRIM$(UCase$(FORMAECO$(i&)))
        If FECO1$ <> "" Then
          JI& = JI& + 1
          
          On Error Resume Next
          If JI& > GRID.Cols Then GRID.Cols = JI&
          If Err.Number = 30006 Then
            Call COMUNI("Imposible Mostrar Consulta !!!\Cantidad de Registros Supera la Permitida\Si la Consulta Poesee Filtro Intente Acotar el Rango de Búsqueda\O Bien, Consulte a su Soporte de Sistemas")
            Screen.MousePointer = 1
            Exit Sub
          End If
          On Error GoTo 0
          
          GRID.COL = JI& - 1: ACAMPO$(JI&) = ""   ' SELCAMPO$(I&)
          AMAXI%(JI&) = 0
          GRID.ColAlignment(JI& - 1) = 1
          ENCABE$ = "Referencia"
          If TECO$ = "DEUDO" Or TECO$ = "CLIEN" Then
               ENCABE$ = "R.Social Cliente"
            ElseIf TECO$ = "PROVE" Or TECO$ = "ACRED" Then
               ENCABE$ = "R.Social Proveedor"
            ElseIf TECO$ = "MASTE" Then
               ENCABE$ = "Descripcion"
            ElseIf TECO$ = "SUTOT" Then
               ENCABE$ = ">Sub-Total"
               If XVALO(Mid$(FORMAECO$(i&), 2)) < 9 Then
                  ENCABE$ = ">S-Tot"
               End If
            ElseIf TECO$ = "COUNT" Then
               ENCABE$ = ">Cant"
               If XVALO(Mid$(FORMAECO$(i&), 2)) < 9 Then
                  ENCABE$ = ">Qt"
               End If
          End If
          If Left$(ENCABE$, 1) = ">" Then
             ENCABE$ = Mid$(ENCABE$, 2)
             GRID.ColAlignment(JI& - 1) = 7      ' derecha y centrado vertical
          End If
          Label5.FontName = GRID.CellFontName
          Label5.FontSize = GRID.CellFontSize
          Label5.FontBold = GRID.CellFontBold
          ' Label5.Caption = TEXTBIN$("1.234567", FORMAECO$(I&)) + "*"
          Label5.Caption = FORMATNUM$("9999999.9999999", FORMAECO$(i&)) + "*"    ' PROVISORIAMENTE A VER SI FUNCIONA
          GRID.ColWidth(JI& - 1) = Label5.Width
25        If JI& >= Label8.Count Then
             NUPIC% = Label8.Count
             Load Label8(NUPIC%)
             GoTo 25
          End If
          Label8(JI&) = Str$(Len(Label5) - 1)
          '
          GRID.Row = 0
          GRID.TEXT = ENCABE$
          '
        End If
      End If
      '
    Next i&
    '
    If HAYTOTA% = 1 Then
       ' GRID.Height = Height - 300 - GRID.Top - 500
       GRID.Height = ScaleHeight - GRID.Top - 500
       GRID2.Height = ScaleHeight - GRID.Top - 500
       Call MUESTRATOTS
    End If
    '
    RWW& = 0: VANTE = "": MAXROW& = 0: ICOLU% = 0
    On Error Resume Next
    Visible = True
    On Error GoTo 0
    MARCOBARRA.Visible = True
30  With RS
       Do While Not .EOF
       Call TRACE(20, RWW&, FIN&)
       If CANCELA% = 1 Then GoTo 99
       Openforms = DoEvents
       JI& = 0
       For i& = 1 To RS.Fields.Count
         '
         COLSU% = 0
         TECO$ = Left$(TRIM$(UCase$(TABLAECO$(i&))), 5)
         If TECO$ = "SUTOT" Or TECO$ = "COUNT" Then
            PPXX% = InStr(TABLAECO$(i&), "/")
            If PPXX% > 0 Then
               COLSU% = XVALO(Mid$(TABLAECO$(i&), PPXX% + 1))
            End If
            If COLSU% > 0 Then
              GRID.Row = RWW&: GRID.COL = COLSU% - 1: VANTE = GRID.TEXT
              GRID.Row = RWW& + 1
              If GRID.TEXT <> VANTE Then
                  SUBTOT#(JI& + 1) = 0
                Else
                  If JI& > GRID.Cols - 2 Then GRID.Cols = JI& + 2
                  GRID.TextMatrix(RWW&, JI& + 1) = ""
              End If
            End If
         End If
         '
         If RWW& > GRID.Rows - 2 Then
            On Error Resume Next
            GRID.Rows = RWW& + 2
            If Err Then
               On Error GoTo 0
               OPX% = COMALTER%("Memoria Insuficiente para Ejecutar Consulta Completa.\   \¿ Desea Ver Reporte PARCIAL Generado ?\\Si, Ver Parcial\No, Cancelar")
               If OPX% = 1 Then
                 RWW = RWW& - 1
                 GoTo 42
               End If
               GoTo 99
            End If
            On Error GoTo 0
         End If
         JI& = JI& + 1
         If JI& > GRID.Cols Then
            GRID.Cols = JI&
         End If
         On Error Resume Next
         valor = .Fields(i& - 1).Value
         '
         ' calculo algebraico
         If TRIM$(CONVALO$(i&)) <> "" Then
             'If UCase$(Left$(SELCAMPO$(I&), 8)) <> "VALOBADA" Then
             CONDIX$ = TEXREPLA$(CONVALO$(i&))
             If TRIM$(TABVALO$(i&)) = "" Then   ' no es un valobada
                valor = RESUCALC#(CONDIX$)
                GoTo 35
             End If
             'End If
         End If
         '
         If UCase$(Left$(SELCAMPO$(i&), 8)) = "VALOBADA" Then
            TABLAX$ = TABVALO$(i&)
            Campox$ = CAMVALO$(i&)
            CONDIX$ = CONVALO$(i&)
17          PPXX1% = InStr(CONDIX$, "{")
            If PPXX1% > 0 Then
               PPXX2% = InStr(CONDIX$, "}")
               If PPXX2% > PPXX1% + 1 Then
                  CONDIXA$ = CONDIX$
                  VAREPLA$ = Mid$(CONDIX$, PPXX1% + 1, PPXX2% - PPXX1% - 1)
                  IZQUI$ = Left$(CONDIX$, PPXX1% - 1)
                  DEREX$ = Mid$(CONDIX$, PPXX2% + 1)
                  '
                  ' ESTO SE PUSO PARA POR VALOBADA PODER REFERENCIAR CAMPOS ANTERIORES
                  For JJ& = 1 To GRID.Cols
                     If UCase$(ACAMPO$(JJ&)) = UCase$(VAREPLA$) Then
                        VAREPLA$ = "'" & TRIM$(GRID.TextMatrix(RWW& + 1, JJ& - 1)) & "'"
                        GoTo 29
                     End If
                  Next JJ&
                  '
                  ' ESTO ES LA MODALIDAD ANTERIOR QUE NO REFERENCIABA CAMPOS VALOBADA ANTERIORES
22                For JJ& = 1 To .Fields.Count
                     If UCase$(SELCAMPO$(JJ&)) = UCase$(VAREPLA$) Then
                        VAREPLA$ = "'" & .Fields(JJ& - 1).Value & "'"
                        GoTo 29
                     End If
                  Next JJ&
                  '
29                CONDIX$ = IZQUI$ + VAREPLA$ + DEREX$
                  If CONDIX$ <> CONDIXA$ Then GoTo 17
                  ' VALOR = VALOBADA(TABLAX$, CAMPOX$, CONDIX$)
               End If
            End If
            
170         PPXX1% = InStr(Campox$, "{")
            If PPXX1% > 0 Then
               PPXX2% = InStr(Campox$, "}")
               If PPXX2% > PPXX1% + 1 Then
                  CONDIXA$ = Campox$
                  VAREPLA$ = Mid$(Campox$, PPXX1% + 1, PPXX2% - PPXX1% - 1)
                  IZQUI$ = Left$(Campox$, PPXX1% - 1)
                  DEREX$ = Mid$(Campox$, PPXX2% + 1)
                  '
                  ' ESTO SE PUSO PARA POR VALOBADA PODER REFERENCIAR CAMPOS ANTERIORES
                  For JJ& = 1 To GRID.Cols
                     If UCase$(ACAMPO$(JJ&)) = UCase$(VAREPLA$) Then
                        For JJ0& = 1 To GRID.Cols
                            If UCase$(SELCAMPO$(JJ0&)) = UCase$(VAREPLA$) Then Exit For
                        Next JJ0&
                            
                        VAREPLA$ = "'" & TRIM$(GRID.TextMatrix(RWW& + 1, JJ& - 1)) & "'"
                        GoTo 290
                     End If
                  Next JJ&
                  '
                  ' ESTO ES LA MODALIDAD ANTERIOR QUE NO REFERENCIABA CAMPOS VALOBADA ANTERIORES
220               For JJ& = 1 To .Fields.Count
                     If UCase$(SELCAMPO$(JJ&)) = UCase$(VAREPLA$) Then
                        VAREPLA$ = "'" & .Fields(JJ& - 1).Value & "'"
                        GoTo 290
                     End If
                  Next JJ&
                  '
290               If Mid$(FORCAMPO$(JJ0&), 1, 1) = "D" Then
                    VAREPLA$ = TRIM$(REPLACAR$(VAREPLA$, "'", ""))
                    If VAREPLA$ = "" Then VAREPLA$ = "01/01/80"
                    FechaNumerica% = FECHANUM(VAREPLA$)
                    VAREPLA$ = "'" & FETEXCOR1$(FechaNumerica%) & "'"
                  End If
  
                  Campox$ = IZQUI$ + VAREPLA$ + DEREX$
                  If Campox$ <> CONDIXA$ Then GoTo 170
                  ' VALOR = VALOBADA(TABLAX$, CAMPOX$, CONDIX$)
               End If
            End If
            
            
            TIMEOUTANT = FLEXCONN.CommandTimeout
            FLEXCONN.CommandTimeout = TIMEOUTANT * 2
            valor = VALOBADA(TABLAX$, Campox$, CONDIX$)
            FLEXCONN.CommandTimeout = TIMEOUTANT
         End If
         '
35       ' aqui se reconoce y procesa el formato grafico
         FORCAX$ = Left$(TRIM$(UCase$(FORCAMPO$(i&))), 3)
         If FORCAX$ = "JPG" Or FORCAX$ = "BMP" Or FORCAX$ = "GIF" Then
             Call PONEIMAGENENCELDA(GRID, RWW& + 1, JI& - 1, FORCAX$, valor)
             
             ' Call CARGAIMAGENENCELDA(RWW& + 1, JI& - 1, VALOR, FORCAX$)
             GoTo 36
         End If
         '
         TEXTOCONFORMATO$ = FORMATNUM$(valor, FORCAMPO$(i&))
         If Mid$(FORCAMPO$(i&), 1, 1) = "D" Then
            If TRIM$(TEXTOCONFORMATO) = "01/01/80" Then TEXTOCONFORMATO$ = ""
         End If
         VAFORMA$ = TEXTOCONFORMATO$
         If RWW& < 1 Then
            ICOLU% = ICOLU% + 1
            FORCOLUM$(ICOLU%) = FORCAMPO$(i&)
         End If
         If Left$(UCase$(FORCAMPO$(i&)), 1) <> "X" Then
              GRID.TextMatrix(RWW& + 1, JI& - 1) = VAFORMA$
              If RWW& + 1 > MAXROW& Then MAXROW = RWW& + 1
            Else
              Call PONETEXMULINENCELDA(GRID, RWW& + 1, JI& - 1)  ' Modificado EPB 16/01/2011 para mostrar como imagen
'              FORMATX% = 1
'              For KIKI% = 1 To CARETEX%
'                 If RWW& + KIKI% > GRID.Rows - 1 Then
'                   GRID.Rows = RWW& + KIKI% + 1
'                 End If
'                 GRID.TextMatrix(RWW& + KIKI%, JI& - 1) = RETEX$(KIKI%)
'                 If RWW& + KIKI% > MAXROW& Then MAXROW = RWW& + KIKI%
'                 If KIKI% > 1 Then MULTILIN% = 1
'              Next KIKI%
         End If
         If FORCAMPO$(i&) = "D20" Then
            GRID.TextMatrix(RWW& + 1, JI& - 1) = Format$(CDbl(valor), "DD/MM/YY - hh:mm:ss")
         End If
         On Error GoTo 0
         If TIESUM%(JI&) = 1 Then SUMA#(JI&) = SUMA#(JI&) + XVALO(valor)
         ' VASUTO = XVALO(VALOR)    ' este reordenamiento acelera el procesamiento - EPB 16/01/11
         If VarType(valor) < 2 Or (VarType(valor) > 6 And VarType(valor) < 14) Or VarType(valor) > 17 Then
              VASUTO = 1
            Else
              VASUTO = XVALO(valor)
         End If
         On Error Resume Next
         SUBTOT#(JI&) = SUBTOT#(JI&) + VASUTO
         On Error GoTo 0
         '
         SUBTOTI% = 0
         If TECO$ <> "" Then
                  If valor < 0 Then
                     A = 1
                  End If

           FECO1$ = TRIM$(UCase$(FORMAECO$(i&)))
           If FECO1$ <> "" Then
             VALOECO = Null: SUTOTA% = 0
             If TECO$ = "DEUDO" Or TECO$ = "CLIEN" Then
                   If valor < 0 Then
                     A = 1
                   End If
                   VALOECO = VectCliente.nombre(XVALO(valor))
                   'VALOECO = rasocli$(XVALO(VALOR))
                 ElseIf TECO$ = "PROVE" Or TECO$ = "ACRED" Then
                   VALOECO = VectProveedor.nombre(XVALO(valor))
                   'VALOECO = rasocli$((-1) * XVALO(VALOR))
                 ElseIf TECO$ = "MASTE" Then
                   TXVALO$ = SAFETEXT$(valor)
                     If Len(TXVALO$) > 4 Then
                          CIXI% = CODINT%(TXVALO$)
                          If CIXI% < 1 Then CIXI% = XVALO(valor)
                        Else
                          CIXI% = XVALO(valor)
                     End If
                     VALOECO = VectArticulos.Descripcion(CIXI%)
                   'VALOECO = DESCRIT0$(CIXI%)
                 ElseIf TECO$ = "SUTOT" Or TECO$ = "COUNT" Then
                   VALOECO = SUBTOT#(JI&)
                   If Abs(VALOECO) > 0.000005 Then SUBTOTI% = 1
                 ElseIf TECO$ = "USERS" Then
                    NroUsuario% = XVALO(valor)
                    VUSER$ = LOADFILE$("USER.DAT")
                    F0$ = Mid$(VUSER$, 64 * NroUsuario% - 63, 64)
                    Ident$ = Mid$(F0$, 31, 12)
                    If Left$(Ident$, 1) = "@" Then Ident$ = DESENCRI$(Mid$(F0$, 32, 11)) + " "
                    Call CIERRARCH("USER")
                    VALOECO = Ident$
'                 ElseIf EXISTE%(LUDAT$ + TABLAECO$(i&) + ".RFS") Then
                 ElseIf TECO$ = "MONED" Then
                    VALOECO = DEMONECO$(XVALO(valor))
                 ElseIf TECO$ = "STATU" Then
                    VALOECO = VECTORSTATUS$(XVALO(valor))
                 ElseIf TRIM$(TABLAECO$(i&)) <> "" Then 'SI ES UNA CAMPO ECO
                   ABUS$ = SAFETEXT$(valor)
                   VALOECO = TRIM$(VectTablasDat.ValorTblDat(TABLAECO$(i&), ABUS$)) 'LO BUSCA EN EL VECTOR DE CLASE
                   If VectTablasDat.RetornoOk < 0 Then  'SI NO LO ENCUENTRA RETORNA CON OK 0 -1
                     If EXISTE%(LUDAT$ + TABLAECO$(i&) + ".RFS") Then 'EN ESE CASO LO BUSCA DE LA MANERA TRADICIONAL
                        TFF% = TIPFIELD%(TABLAECO$(i&), 1)
                        LFF% = LENFIELD%(TABLAECO$(i&), 1)
                        If TFF% = 6 Then
                          ABUS$ = SAFETEXT$(valor)
                          GoTo 34
                        End If
                        ABUS$ = MKS$(XVALO(valor))
                        On Error Resume Next
                        If LFF% = 2 Then ABUS$ = MKI$(Int(XVALO(valor)))
                        If LFF% = 1 Then ABUS$ = Chr$(Int(XVALO(valor)))
                        On Error GoTo 0
34                      VALOECO = ECOARCH$(TABLAECO$(i&), ABUS$)
                     End If
                   End If
             End If
             JI& = JI& + 1
             If JI& > GRID.Cols Then GRID.Cols = JI&
             If SUBTOTI% = 1 Then
                GRID.COL = JI& - 1
                If RWW <= 1 Then ACAMPO$(JI&) = SELCAMPO$(i&)
                
                GRID.CellFontBold = True
                GRID.CellFontUnderline = True
             End If
             '
             If RWW& < 1 Then
               ICOLU% = ICOLU% + 1
               FORCOLUM$(ICOLU%) = FORMAECO$(i&)
             End If
             '
             On Error Resume Next
             GRID.TextMatrix(RWW& + 1, JI& - 1) = FORMATNUM$(VALOECO, FORMAECO$(i&))
             On Error GoTo 0
           End If
         End If
36     Next i&
       RWW& = RWW& + 1
       If RWW& < MAXROW& Then RWW& = MAXROW&
       If FORMATX% > 0 Then
          RWW& = RWW& + 1
          If RWW& > GRID.Rows - 2 Then GRID.Rows = RWW& + 2
          For JI& = 1 To GRID.Cols
             GRID.TextMatrix(RWW&, JI& - 1) = String$(128, "-")
          Next JI&
       End If
       .MoveNext
       Loop
    End With
    '
    If VUELTANUMERO% < 1 And CONSIDERAR_MOVIVIE% > 0 Then
       VUELTANUMERO% = 1
       GoTo 10
    End If
42  Label7 = Str$(RWW&)
    GRID.Rows = RWW& + 1
    '
    If HAYTOTA% = 1 Then
      For KI& = 1 To GRID.Cols
        If TIESUM%(KI&) = 1 Then
          If FORSUMA$(KI&) <> "" Then
            Label4(KI& - 1).Caption = TRIM$(FORMATNUM$(SUMA#(KI&), FORSUMA$(KI&)))
            Label4(KI& - 1).Refresh
          End If
        End If
      Next KI&
    End If
    '
    ' ESTO QUE SIGUE ES PARA FORZAR EL ORDENAMIENTO ASCENDENTE POR CODIGO EXTERNO
    ' CUANDO EL PRIMER CAMPO ORDENATRIZ ES UN CÓDIGO O SEA FORMATO 'Cnn'
    For KI& = 1 To CANFI
      If ORDCAMPO$(KI&) <> "" Then
        If UCase$(Left$(FORCAMPO$(KI&), 1)) = "C" Or UCase$(Left$(SELCAMPO$(KI&), 8)) = "VALOBADA" Then
          GRID.COL = KI& - 1
          GRID.Row = 0
          SENTIOR%(KI&) = 1
          If ORDCAMPO$(KI&) = "ASC" Then SENTIOR%(KI&) = (-1)
          Call GRID_DblClick
        End If
        GoTo 99     ' CONSIDERA SOLO EL PRIMER CAMPO ORDENATRIZ
      End If
    Next KI&
    '
99  'Call COLOREAR_GRILLASQL(GRID, &HC0C0C0, 1)
'    Call COLOREAR_GRILLASQL2(GRID, &HDAE4EB, 1)
    
    MARCOBARRA.Visible = False
    Screen.MousePointer = 1
    
    ' CHEQUEA SI SE ESTA UTILIZANDO UNA TABLA TEMPORAL
    ' PARA COMPLETAR LA CONSULTA. EN CASO AFIRMATIVO, LA ELIMINA LUEGO DE HABERLA UTILIZADO.
    ' NO ELIMINA LAS TABLAS CUANDO EL REPORTE TIENE FILTROS. SE QUITAN JUNTO CON EL CIERRE DE GENREP
    If InStr(1, TXT_CONSULTA, "FROM TEMP_") > 0 And ReporteConFiltro% = 0 Then
        n% = InStr(1, TXT_CONSULTA$, "TEMP_")
        L% = InStr(n%, TXT_CONSULTA$, " ")
        TablaTemporal$ = Mid$(TXT_CONSULTA, n%, L% - n%)
        On Error GoTo Err_BorrarTablaTemporal
        FLEXCONN.Execute "DROP TABLE " & TablaTemporal$
        On Error GoTo 0
    ElseIf ReporteConFiltro% = 1 Then
        ReporteConsultaSql$ = TXT_CONSULTA$
    End If
    '
    Exit Sub
    '
    '
Err_BorrarTablaTemporal:
    Call MENSERR(24, "Imposible Eliminar Tabla Temporal\" & Err.Description)
    Screen.MousePointer = 1
    Exit Sub
    
End Sub

Function TEXREPLA$(TEXORI$)
     '
     TREP$ = TEXORI$
     '
10   PPXX1% = InStr(TREP$, "{")
     If PPXX1% > 0 Then      ' BUSCA ENCIERRO ENTRE LLAVES
        PPXX2% = InStr(TREP$, "}"): If PPXX2% <= PPXX1% Then PPXX2% = 1 + Len(TREP$)
        IZQUI$ = Left$(TREP$, PPXX1% - 1)
        DEREX$ = Mid$(TREP$, PPXX2% + 1)
        VAREPLA$ = Mid$(TREP, PPXX1% + 1, PPXX2% - PPXX1% - 1)
        '
        'For JJ& = 1 To RS.Fields.Count
        For JJ& = 1 To GRID.Cols
           'If UCase$(SELCAMPO$(JJ&)) = UCase$(VAREPLA$) Then
           If UCase$(ACAMPO$(JJ&)) = UCase$(VAREPLA$) Then
              ' VALORR = RS.Fields(JJ& - 1).Value
              VALORR = GRID.TextMatrix(RWW& + 1, JJ& - 1)
              If IsNull(VALORR) = True Then VALORR = GRID.TextMatrix(RWW& + 1, JJ& - 1)
              VAREPLA$ = TRIM$(FORMATNUM$(VALORR, "F20.6"))
              If IsDate(VALORR) Then VAREPLA$ = TRIM$(CStr(VALORR))
              If VAREPLA$ = "" Then VAREPLA$ = "0.000000"
              GoTo 29
           End If
        Next JJ&
        '
29      TREP$ = IZQUI$ + comilla$ + VAREPLA$ + comilla$ + DEREX$
        GoTo 10
     End If
     '
80   TEXREPLA = TREP$
     '
End Function
'
Private Sub Combo1_Click(Index As Integer)
   VEMOFREN% = 11
End Sub
'
Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
   VEMOFREN% = 11
   If KeyAscii = 13 Or KeyAscii = 9 Then
      For I1& = Index + 1 To Combo1.Count - 1
        If Combo1(I1&).Visible = True Then
          Combo1(I1&).SetFocus
          Exit Sub
        End If
      Next I1&
      Call Command3_Click
   End If
End Sub

Private Sub Command1_Click()
  Command1.Enabled = False
  Call CopyToExcel(GRID, "REPORTE", "")
  Command1.Enabled = True
End Sub

Private Sub Command2_Click(Index As Integer)
   '
   VEMOFREN% = 11
   '
   If FilText1(Index).ToolTipText <> "" Then
      ASELEC$ = FilText1(Index).ToolTipText
      Call SELECHO(ASELEC$)
      VEMOFREN% = 11
      FilText1(Index).TEXT = VALACT1$(ASELEC$)
      Exit Sub
   End If
   '
   CAMCONDI$ = CAMPOSEL$(Index)
   If CAMCONDI$ <> "" Then
      Screen.MousePointer = 11
      ' ESTA INTERVENCION PARA PODER SELECCIONAR POR TABLA SECUNDARIA - 26/08/09
      TAPRIN$ = TABSELE$
      SQLY$ = "SELECT DISTINCT " & CAMCONDI$ & " AS VALOR FROM " & TAPRIN$ & " WITH(NOLOCK)"
      '
      TASEC$ = TABSECU$
      If TASEC$ <> "" Then
         CAMPRIN$ = CAMPOMA1$
         If CAMPRIN$ <> "" Then
            CAMSEC$ = CAMPOMA2$
            If CAMSEC$ <> "" Then
               SQLY$ = SQLY$ + " INNER JOIN " + TASEC$
               SQLY$ = SQLY$ + " ON " + TAPRIN$ + "." + CAMPRIN$
               SQLY$ = SQLY$ + " = " + TASEC$ + "." + CAMSEC$
            End If
         End If
      End If
      
      If UNIONALL$ <> "" Then
        PrimeraConsulta$ = SQLY$
        SegundaConsulta$ = "SELECT DISTINCT " & CAMCONDI$ & " AS VALOR FROM " & UNIONALL$ & " WITH(NOLOCK)"
        SQLY$ = PrimeraConsulta$ & " UNION " & SegundaConsulta$
      End If
      '
      ' AGREGADO EL 25/ENE/2010 PARA QUE EN LISTA DE SELECCION CONSIDERE
      ' LAS RESTANTES CONDICIONES
      If CACONDI% > 0 Then
        JJ5& = 0
        For I5& = 1 To CACONDI%
          If TRIM$(VAREFCOND$(I5&)) <> "" Then
            If InStr(VAREFCOND$(I5&), "¿") < 1 Then
'              If JJ5& < 1 Then
'                  SQLY$ = SQLY$ + " WHERE "
'                Else
'                  SQLY$ = SQLY$ + " AND "
'              End If
              'INTERVENCION OR
              If JJ5& < 1 Then
                    SQLY$ = SQLY$ + " WHERE "
              ElseIf UCase$((Left$(TRIM$(CAMPOCOND$(I5&)), 3))) = "OR " Then
                    SQLY$ = SQLY$ + " "
              Else
                     SQLY$ = SQLY$ + " AND "
              End If
              SQLY$ = SQLY$ + CAMPOCOND$(I5&) + " "
              SQLY$ = SQLY$ + OPERACOND$(I5&) + " "
              SQLY$ = SQLY$ + VAREFCOND$(I5&)
              JJ5& = JJ5& + 1
            End If
          End If
        Next I5&
      End If
      '
      SQLY$ = SQLY$ + " ORDER BY " & CAMCONDI$ & " ASC"
      '
      ESFECHA% = 0
      If TIPODECAMPO(TAPRIN$, CAMCONDI$) = 58 Then ESFECHA% = 1
      '
      Combo1(Index).Clear
      UVALO$ = ""
      Dim VAPOSI As ADODB.Recordset
      Set VAPOSI = New ADODB.Recordset
      Set VAPOSI = FLEXCONN.Execute(FILTSQL$(SQLY$))
      With VAPOSI
         Do While Not .EOF
            On Error Resume Next
            If ESFECHA% = 0 Then
                Combo1(Index).AddItem !valor
              Else
                AVALO$ = FETEXCOR1$(CVINT(!valor))
                If AVALO$ <> UVALO$ Then
                   UVALO$ = AVALO$
                   Combo1(Index).AddItem UVALO$
                End If
            End If
            On Error GoTo 0
         .MoveNext
         Loop
      End With
      Screen.MousePointer = 1
   End If
   '
End Sub

Private Sub Command3_Click()
   VEMOFREN% = 11
   CANCELA% = 0
   Call EJECUTACONSULTA
   
End Sub

Private Sub FilText1_Change(Index As Integer)
   If FilText1(Index).ToolTipText <> "" Then
      ASELEC$ = FilText1(Index).ToolTipText
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

Private Sub FilText1_Click(Index As Integer)
   VEMOFREN% = 11
End Sub

Private Sub FilText1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index < FilText1.UBound Then
         On Error Resume Next
         If FilText1(Index + 1).Visible = True Then
              FilText1(Index + 1).SetFocus
            ElseIf Combo1(Index + 1).Visible = True Then
              Combo1(Index + 1).SetFocus
         End If
         On Error GoTo 0
       Else
         Call Command3_Click
      End If
   End If
End Sub

Private Sub Form_Activate()
   'VEMOFREN% = 0
   Call Timer1_Timer            ' traer pantalla al frente
   Timer2.Enabled = True
End Sub

Private Sub Form_Load()
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
      Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
   End If
   Call GRATITFOR(Caption)
   '
   ' APLICA FUENTE ANCHO FIJO
   Call AplicaFuenteAnchoFijo(GRID, , 8)
   '
   Timer3.Enabled = True
   '
End Sub

Private Sub Form_LostFocus()
   Timer2.Enabled = False
End Sub

Private Sub Form_Resize()
    
    GRID.Width = SHOWREP07.Width - 300
    GRID.Height = (SHOWREP07.Height - GRID.Top) - 500
    GRID2.Width = SHOWREP07.Width - 300
    GRID2.Height = (SHOWREP07.Height - GRID.Top) - 500
    Toolbar1.Left = SHOWREP07.Width - Toolbar1.Width - 50
    MARCOBARRA.Left = Toolbar1.Left + (Toolbar1.Width - MARCOBARRA.Width) / 2
    Toolbar2.Left = Toolbar1.Left - Toolbar2.Width - 50
    '
    If HAYTOTA% = 1 Then
       GRID.Height = ScaleHeight - GRID.Top - 500
       GRID2.Height = ScaleHeight - GRID.Top - 500
       
       Call MUESTRATOTS
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CANCELA% = 1
    Cancel = 0
   
    ' CHEQUEA SI SE ESTA UTILIZANDO UNA TABLA TEMPORAL
    ' EN CASO AFIRMATIVO, LA ELIMINA LUEGO DE HABERLA UTILIZADO.
    ' ACA SON ELIMINADAS LAS TABLAS UTILIZADAS EN REPORTE CON FILTROS
    ' EN EJECUTA_CONSULTA SE ELIMINAN LAS QUE CORRESPONDEN A REPORTES SIN FILTRO
    Screen.MousePointer = 11
    If InStr(1, ReporteConsultaSql$, "FROM TEMP_") > 0 And ReporteConFiltro% = 1 Then
        n% = InStr(1, ReporteConsultaSql$, "TEMP_")
        L% = InStr(n%, ReporteConsultaSql$, " ")
        TablaTemporal$ = Mid$(ReporteConsultaSql$, n%, L% - n%)
        On Error GoTo Err_BorrarTablaTemporal
        FLEXCONN.Execute "DROP TABLE " & TablaTemporal$
        On Error GoTo 0
    End If
    Screen.MousePointer = 1
    
    Exit Sub
    '
Err_BorrarTablaTemporal:
    Call MENSERR(24, "Imposible Eliminar Tabla Temporal\" & Err.Description)
    Screen.MousePointer = 1
    Exit Sub
    
End Sub

Private Sub GRID_Click()
   If GRID.Row <= 1 Then
      ' para apagar el color azul de la seleccion
      On Error Resume Next
      GRID.RowSel = GRID.Row
      On Error GoTo 0
   End If
End Sub

Private Sub GRID_DblClick()
    If NO_ORDENAR% > 0 Then Exit Sub
    If GRID.Row <= 1 Then
       '
       If MULTILIN% > 0 Then
          Call MENSERR(24, "Imposible Ordenar")
          Exit Sub
       End If
       '
       Screen.MousePointer = 11
       '
       COLUSOR& = GRID.COL + 1
       If SENTIOR%(COLUSOR&) < 1 Then
            SENTIOR%(COLUSOR&) = 1    ' ASCENDENTE
          Else
            SENTIOR%(COLUSOR&) = (-1) ' DESCENDENTE
       End If
       '
       Dim VALORES
       Dim ARTEMP
       URETRA& = XVALO(Label7)
       filas& = URETRA&
       COLUS& = GRID.Cols
       ReDim VALORES(filas&, COLUS&)
       ReDim ARTEMP(COLUS&)
       '
       ESPORFECHA% = 0
       For i& = 1 To filas&
         For J& = 1 To COLUS&
           LEFOR$ = UCase$(Left$(FORCOLUM$(J&), 1))
           If LEFOR$ = "A" Or LEFOR$ = "C" Or LEFOR$ = "D" Or LEFOR$ = "H" Or LEFOR$ = "T" Or LEFOR$ = "X" Then
                 VALORES(i&, J&) = GRID.TextMatrix(i&, J& - 1)
               Else
                 VALORES(i&, J&) = XVALO(GRID.TextMatrix(i&, J& - 1))
           End If
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
         For i& = 1 To filas&
           VLX$ = VALORES(i&, COLUSOR&)
           VALORES(i&, COLUSOR&) = FECHANUM(VLX$)
         Next i&
       End If
       '
       ' ORDENAMIENTO SHELL
       SIG% = SENTIOR%(COLUSOR&)
       INCRE& = filas& \ 2
       Do Until INCRE& < 1
         For UI& = INCRE& + 1 To filas&
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
       For i& = 1 To filas&
         For J& = 1 To COLUS&
           If J& = COLUSOR& Then
             If ESPORFECHA% = 1 Then
               VLY = VALORES(i&, J&)
               GRID.TextMatrix(i&, J& - 1) = FECHATEX$(VLY)
               GoTo 29
             End If
           End If
           '
           LEFOR$ = UCase$(Left$(FORCOLUM$(J&), 1))
           If LEFOR$ = "A" Or LEFOR$ = "C" Or LEFOR$ = "D" Or LEFOR$ = "H" Or LEFOR$ = "T" Or LEFOR$ = "X" Then
                GRID.TextMatrix(i&, J& - 1) = VALORES(i&, J&)
              Else
                GRID.TextMatrix(i&, J& - 1) = FORMATNUM$(VALORES(i&, J&), FORCOLUM$(J&))
           End If
29       Next J&
       Next i&
       Call RESUTOT
       Screen.MousePointer = 1
       '
    End If
End Sub

Sub RESUTOT()
    '
    CANSU% = 0
    OFCOL% = 0
    '
    For i& = 1 To GRID.Cols
       COLSU% = 0
       TECO$ = Left$(TRIM$(UCase$(TABLAECO$(i&))), 5)
       If TECO$ <> "" Then
          OFCOL% = OFCOL% + 1
       End If
       '
       If TECO$ = "SUTOT" Or TECO$ = "COUNT" Then
          OFCOL% = OFCOL% - 1
          PPXX% = InStr(TABLAECO$(i&), "/")
          If PPXX% > 0 Then
             COLSU% = XVALO(Mid$(TABLAECO$(i&), PPXX% + 1))
          End If
          If COLSU% > 0 Then
            ' CALCULA LOS SUBTOTALES
            SUTOT# = 0
            If i& + CANSU% + 1 > GRID.Cols - 1 Then Exit Sub
            For RWW& = 1 To GRID.Rows - 1
               GRID.TextMatrix(RWW&, i& + OFCOL%) = ""
               If TECO$ = "SUTOT" Then
                    SUTOT# = SUTOT# + XVALO(GRID.TextMatrix(RWW&, i& + OFCOL% - 1))
                  ElseIf TECO$ = "COUNT" Then
                    SUTOT# = SUTOT# + 1
               End If
               valor = GRID.TextMatrix(RWW&, COLSU% - 1)
               VALSIG = "": If RWW& < GRID.Rows - 1 Then VALSIG = GRID.TextMatrix(RWW& + 1, COLSU% - 1)
               GRID.Row = RWW&
               GRID.COL = i& + OFCOL%
               If VALSIG <> valor Or RWW& = GRID.Rows - 1 Then
                    GRID.TextMatrix(RWW&, i& + OFCOL%) = FORMATNUM$(SUTOT#, FORMAECO$(i&))
                    SUTOT# = 0
                    GRID.CellFontBold = True
                    GRID.CellFontUnderline = True
                  Else
                    GRID.CellFontBold = False
                    GRID.CellFontUnderline = False
               End If
            Next RWW&
            CANSU% = CANSU% + 1
          End If
          OFCOL% = OFCOL% + 1
       End If
    Next i&
    '
End Sub

Sub RESUTOT_vie()
    '
    CANSU% = 0
    '
    For i& = 1 To GRID.Cols
       COLSU% = 0
       TECO$ = Left$(TRIM$(UCase$(TABLAECO$(i&))), 5)
       If TECO$ = "SUTOT" Or TECO$ = "COUNT" Then
          PPXX% = InStr(TABLAECO$(i&), "/")
          If PPXX% > 0 Then
             COLSU% = XVALO(Mid$(TABLAECO$(i&), PPXX% + 1))
          End If
          If COLSU% > 0 Then
            ' CALCULA LOS SUBTOTALES
            SUTOT# = 0
            If i& + CANSU% + 1 >= GRID.Cols - 1 Then Exit Sub
            GRID.COL = i& + CANSU% + 1
            For RWW& = 1 To GRID.Rows - 1
               GRID.TextMatrix(RWW&, i& + CANSU% + 1) = ""
               If TECO$ = "SUTOT" Then
                    SUTOT# = SUTOT# + XVALO(GRID.TextMatrix(RWW&, i& + CANSU%))
                  ElseIf TECO$ = "COUNT" Then
                    SUTOT# = SUTOT# + 1
               End If
               valor = GRID.TextMatrix(RWW&, COLSU% - 1)
               VALSIG = "": If RWW& < GRID.Rows - 1 Then VALSIG = GRID.TextMatrix(RWW& + 1, COLSU% - 1)
               GRID.Row = RWW&
               If VALSIG <> valor Or RWW& = GRID.Rows - 1 Then
                    'GRID.TextMatrix(RWW&, I& + CANSU% + 1) = FORMATNUM$(SUTOT#, FORCAMPO$(I& + CANSU% + 1))
                    GRID.TextMatrix(RWW&, i& + CANSU% + 1) = FORMATNUM$(SUTOT#, FORMAECO$(i&))
                    SUTOT# = 0
                    GRID.CellFontBold = True
                    GRID.CellFontUnderline = True
                  Else
                    GRID.CellFontBold = False
                    GRID.CellFontUnderline = False
               End If
            Next RWW&
            CANSU% = CANSU% + 1
          End If
       End If
    Next i&
    '
End Sub

Private Sub GRID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
       If ARBUS$ <> "" Then
          VUELTA% = 0
          J& = GRID.Row

          h& = 0: If GRID.COL < GRID.Cols Then h& = GRID.COL + 1
10        For U& = J& To GRID.Rows - 1
             HI& = h&: If U& > J& Then HI& = 0
             For i& = HI& To GRID.Cols - 1
              If InStr(UCase$(GRID.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
                GRID.Row = U&: GRID.COL = i&
                On Error Resume Next
                 VISIROWS = (GRID.Height / GRID.RowHeight(J&)) - 4
                If Err > 0 Then
                 VISIROWS = 1
                End If
                On Error GoTo 0
                 If VISIROWS > 0 Then
                    If GRID.TopRow < GRID.Row - VISIROWS Then
                      GRID.TopRow = GRID.Row - VISIROWS
                    End If
                 End If
                GRID.SetFocus
                Exit Sub
              End If
              Next i&
          Next U&
          '
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If

End Sub

Private Sub GRID_Scroll()
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
    '
    If APLINVO$ <> "" Then    ' MUESTRA EL NOMBRE DEL REPORTE
       Timer1.Enabled = False
       MsgBox "Reporte: " + UCase$(APLINVO$) + " - " + TRIM$(Label3), vbSystemModal
       Exit Sub
    End If
    '
    MODICONSU% = 1
    Hide
End Sub

Private Sub Label2_Change()
    If TRIM$(Label2.Caption) <> "" Then
        If Label2.Caption = "@NEW" Then
           'FGENREP07.Combo2.TEXT = ""
           FGENREP07.Label1.Caption = ""
           FGENREP07.Show 1
           Exit Sub
        End If
        Call FILTROVAR
        If SUBAPLI$ <> "" Then
           Call EJEC_SUBAP
        End If
    Else
        GRID.Clear
        For i& = 1 To Label1.Count
            Label1(i& - 1).Visible = False
            Combo1(i& - 1).Visible = False
            Command2(i& - 1).Visible = False
        Next i&
        For i& = 1 To Label4.Count
            Label4(i& - 1).Visible = False
            Label4(i& - 1).Caption = ""
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
    ' BANDERA QUE SE UTILIZA PARA MANTENER ACTIVA LAS TABLAS TEMPORALES
    ' CREADAS DESDE LOS MODULOS JUNTOS CON EL REPORTE, CUANDO EL REPORTE
    ' CONTIENE FITROS.
    If CACONDI% > 0 Then ReporteConFiltro% = 1
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
    '
    If CACONDI% < 1 Then
       Frame1.Visible = False
       Frame1.Height = 0
       GoTo 90
       'Hide
       Exit Sub
    End If
    '
    PVER = 600
    i& = 0
    TAMAXI = 0
    For IKI& = 1 To CACONDI%
       VAREFE$ = VAREFCOND$(IKI&)
       VALORCOND$(IKI&) = VAREFE$
       IZX$ = "": If Left$(VAREFE$, 1) = "'" Then IZX$ = "'": VAREFE$ = Mid$(VAREFE$, 2)
       DZX$ = "": If Right$(VAREFE$, 1) = "'" Then DZX$ = "'": VAREFE$ = Left$(VAREFE$, Len(VAREFE$) - 1)
10     If Left$(VAREFE$, 1) = "¿" Then
          If Right$(VAREFE$, 1) = "?" Then
             i& = i& + 1
             LABPRE$ = Mid$(VAREFE$, 2, Len(VAREFE$) - 2)
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
             Label1(i& - 1).Top = PVER + 400 * (i& - 1)
             Combo1(i& - 1).Top = PVER - 50 + 400 * (i& - 1)
             Command2(i& - 1).Top = PVER - 50 + 400 * (i& - 1)
             Command2(i& - 1).Width = 330
             Label1(i& - 1) = LABPRE$
             Combo1(i& - 1).Clear
             If Label1(i& - 1).Width > TAMAXI Then TAMAXI = Label1(i& - 1).Width
             Label1(i& - 1).Visible = True
             Combo1(i& - 1).Visible = True
             Command2(i& - 1).Visible = True
             FilText1(i& - 1).Visible = False
             FilText1(i& - 1).ToolTipText = ASELEC$
             FiltEco1(i& - 1).Visible = False
             CAMPOSEL$(i& - 1) = CAMPOCOND$(IKI&)
             ULTIFIL& = i&
             '
             If APLINVO$ <> "" Then
               If FILTX$ <> "" Then
                 PPXX% = InStr(FILTX$, ";")
                 If PPXX% < 1 Then PPXX% = InStr(FILTX$, ")")
                 If PPXX% < 1 Then PPXX% = 1 + Len(FILTX$)
                 If PPXX% > 1 Then
                    Combo1(i& - 1).TEXT = Left$(FILTX$, PPXX% - 1)
                    FilText1(i& - 1).TEXT = Left$(FILTX$, PPXX% - 1)
                    FILTX$ = Mid$(FILTX$, PPXX% + 1)
                 End If
               End If
             End If
             '
          End If
       End If
    Next IKI&
    '
    CACONFIL% = i&
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
          GoTo 20
       End If
       '
       Label1(i& - 1).Left = 500 + TAMAXI - Label1(i& - 1).Width
       Combo1(i& - 1).Left = Label1(i& - 1).Left + Label1(i& - 1).Width + 250
       FilText1(i& - 1).Left = Label1(i& - 1).Left + Label1(i& - 1).Width + 250
       Command2(i& - 1).Left = Combo1(i& - 1).Left + Combo1(i& - 1).Width + 150
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
          If ANECE < Width - 2 * Frame1.Left - Command3.Width - 200 Then
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
    Frame1.Height = Command2(ULTIFIL& - 1).Top + 600
    '
    Command3.Top = 130
    Command3.Height = Frame1.Height - 150
    Command3.Left = Frame1.Width - Command3.Width
    '
90  GRID.Top = Frame1.Top + Frame1.Height + 250
    GRID2.Top = Frame1.Top + Frame1.Height + 250
    GRID.Height = Height - GRID.Top - 500
    GRID2.Height = Height - GRID.Top - 500
    If CACONFIL% < 1 Then
       Call Command3_Click
       Exit Sub
    End If
    '
    If APLINVO$ <> "" Then
      For i& = 1 To ULTIFIL&
        If Combo1(i& - 1).Visible = True Then
          If TRIM$(Combo1(i& - 1).TEXT) = "" Then Exit Sub
        End If
      Next i&
      Call Command3_Click
      Exit Sub
    End If
    '
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
         AMAX0 = Command2(i& - 1).Left + Command2(i& - 1).Width + 150 + Command3.Width
      Else
         AMAX1 = Command2(i& - 1).Left + Command2(i& - 1).Width + 150 + Command3.Width
         If FiltEco1(i& - 1).Width > 0 Then
            AMAX2 = FiltEco1(i& - 1).Left + FiltEco1(i& - 1).Width + 150 + Command3.Width
         End If
   End If
   If AMAX1 > AMAX0 Then AMAX0 = AMAX1
   If AMAX2 > AMAX0 Then AMAX0 = AMAX2
   If Frame1.Width < AMAX0 Then
      Frame1.Width = AMAX0
      Command3.Left = Frame1.Width - Command3.Width
      Frame1.Refresh
   End If
   '
End Sub

'
Private Sub Timer1_Timer()
   If INTERPRE% < 1 Then
     If VEMOFREN% < 10 Then          ' que lo haga 10 veces
       VEMOFREN% = VEMOFREN% + 1
       On Error Resume Next
       Me.ZOrder (0)
       AppActivate Caption
       SendKeys Chr$(0)
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub Timer2_Timer()
'   Static ACOLGRI(256)
'   If INTERPRE% < 1 Then
'     For KI% = 0 To GRID.Cols - 1
'       If ACOLGRI(KI%) < 1 Then ACOLGRI(KI%) = GRID.ColWidth(KI%)
'       If GRID.ColWidth(KI%) <> ACOLGRI(KI%) Then
'         Call MUESTRATOTS
'         ACOLGRI(KI%) = GRID.ColWidth(KI%)
'       End If
'     Next KI%
'   End If
End Sub

Private Sub Timer3_Timer()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
                 SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    
    SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, _
                 SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Timer3.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   Select Case UCase$(Button.key)
      Case "PRINT"
        Call IMPRIMIR
      Case "SEARCH"
        Call Search
      Case "EXIT"
        GRID.Clear
        Unload Me
   End Select
   '
End Sub

Sub IMPRIMIR()
   If XVALO(Label7) > 0 Then
      Call PRINREP07.CARGALIMPRE
      If PRINREP07.Visible = True Then
         PRINREP07.Hide
      End If
      PRINREP07.Show 1
   End If
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)
   Select Case UCase$(Button.key)
      Case "EXCEL"
        Call EXPORTEXCEL
      Case "GRAFIC"
        Call GENERGRAF
      Case "FULLSCREEN"
        Call FULLSCREEN
      Case "AUTOADJUST"
        Call AUTOADJUST
   End Select
End Sub

Sub AUTOADJUST()
   '
   Teprueba.Font = GRID.Font
   Teprueba.FontSize = GRID.FontSize
   Teprueba.FontBold = GRID.FontBold
   '
   For COLUI& = 1 To GRID.Cols
      ACOLU = 0
      For RENGI& = 1 To GRID.Rows
         Teprueba.Caption = TRIM$(GRID.TextMatrix(RENGI& - 1, COLUI& - 1)) + "."
         If Teprueba.Width > ACOLU Then ACOLU = Teprueba.Width
      Next RENGI&
      GRID.ColWidth(COLUI& - 1) = ACOLU
   Next COLUI&
   '
End Sub
'
Sub GENERGRAF()
    '
    DEFGRAF07.Label11 = Label2
    DEFGRAF07.Show 1
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
      'Call GENEXCEL(AREXPO$)
      FormatoColumnas$ = ARMA_ENCABEZADO()
      Call GeneraExcelDeGrilla(AREXPO$, GRID, Label3, , FormatoColumnas$)
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
   Dim FORMATONUMERO As String
   Dim IDIOMA As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TIPODATO% = 0: COLUM& = 0
   For i& = 1 To GRID.Rows
       RENG1& = i&: COLUM& = 0: CampoConEco% = 0: ColumnaReal = 0: VieneDeUnEco% = 0
       For COLU1& = 1 To GRID.Cols
          If GRID.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          ColumnaReal = ColumnaReal + 1
          '
          VALORX = GRID.TextMatrix(i& - 1, COLU1 - 1)
          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
             FEXU$ = VALORX
             VALORX = FETEXCOR2$(Int(FECHANUM(FEXU$)))
'             If IDIOMA$ = "" Then
'                IDIOMA$ = "ESP"
'                XLSHEET.Cells(RENG1&, COLUM&) = FEXU$
'                XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
'                '
'                FF0$ = XLSHEET.Cells(RENG1&, COLUM&)
'                FF1% = CVINT(CDate(FF0$))
'                If FECHATEX$(FF1%) <> FEXU$ Then IDIOMA$ = "ING"
'             End If
'             If IDIOMA$ = "ING" Then VALORX = FETEXCOR2$(Int(FECHANUM(FEXU$)))
             TIPODATO% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          If CampoConEco% = 0 Then
            FFCC$ = UCase$(Left$(FORCAMPO$(ColumnaReal), 1))
          Else
            ColumnaReal = ColumnaReal - 1
            FFCC$ = UCase$(Left$(FORMAECO$(ColumnaReal), 1)) ' UCase$(Left$(FORCAMPO$(ColumnaFormato), 1))
            CampoConEco% = 0: VieneDeUnEco% = 1
          End If
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then VALORX = 0
             
             FORMATONUMERO = "0" ' FORMATO ENTERO
             If FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
                MASCARA$ = Mid$(FORCAMPO$(ColumnaReal), 2)
                NN% = InStr(1, MASCARA$, ".")
                PARTEDECIMAL$ = Mid$(MASCARA$, NN% + 1)
                
                JJ% = 1
                If FFCC$ = "G" Or FFCC$ = "$" Then
                    FORMATONUMERO = "#,##0."
                    VALORX = XVALO(REPLACAR(CStr(VALORX), ".", ""))
                    If XVALO(PARTEDECIMAL$) = 0 Then FORMATONUMERO = "#,##0"
                ElseIf FFCC$ = "F" Then
                    FORMATONUMERO = "#0."
                    If XVALO(PARTEDECIMAL$) = 0 Then FORMATONUMERO = "0"
                End If
                While JJ% <= XVALO(PARTEDECIMAL$)
                    FORMATONUMERO = FORMATONUMERO + "0"
                    JJ% = JJ% + 1
                Wend
             End If
             TIPODATO% = 2
          End If
          If FFCC$ = "C" Or FFCC$ = "A" Then VALORX = "'" & VALORX
          '
          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          If TIPODATO% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = FORMATONUMERO
          'If TIPODATO% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TIPODATO% = 0
          '
          'If TRIM$(TABLAECO(COLU1&)) <> "" Then CampoConEco% = 1
          If TRIM$(TABLAECO(ColumnaReal)) <> "" And VieneDeUnEco% = 0 Then CampoConEco% = 1
          VieneDeUnEco% = 0
          '
10     Next COLU1&
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


Sub FULLSCREEN()
   '
    If SHOWREP07.WindowState = 0 Then
     SHOWREP07.WindowState = 2
    Else
     SHOWREP07.WindowState = 0
    End If
    '
    GRID.Width = SHOWREP07.Width - 300
    GRID.Height = (SHOWREP07.Height - GRID.Top) - 500
    GRID2.Width = SHOWREP07.Width - 300
    GRID2.Height = (SHOWREP07.Height - GRID.Top) - 500

    Toolbar1.Left = SHOWREP07.Width - Toolbar1.Width - 50
    MARCOBARRA.Left = Toolbar1.Left + (Toolbar1.Width - MARCOBARRA.Width) / 2
    Toolbar2.Left = Toolbar1.Left - Toolbar2.Width - 50
    Picture1.Left = SHOWREP07.Width - Picture1.Width - 50
    '
    Label9.Width = Me.Width
    '
    If HAYTOTA% = 1 Then
       ' GRID.Height = Height - 300 - GRID.Top - 500
       GRID.Height = ScaleHeight - GRID.Top - 500
       GRID2.Height = ScaleHeight - GRID.Top - 500
       Call MUESTRATOTS
'       PPX& = GRID.Left
'       For KI& = 1 To GRID.Cols
'          If KI& > Label4.Count Then
'             NUPIC% = Label4.Count
'             Load Label4(NUPIC%)
'          End If
'          Label4(KI& - 1).Caption = ""
'          Label4(KI& - 1).Visible = False
'          Label4(KI& - 1).Top = GRID.Top + GRID.Height + 100
'          Label4(KI& - 1).Left = PPX& + 30
'          Label4(KI& - 1).Width = GRID.ColWidth(KI& - 1)
'          If TIESUM%(KI&) = 1 Then
'            If FORSUMA$(KI&) <> "" Then
'              Label4(KI& - 1).Visible = True
'              Label4(KI& - 1).Caption = TRIM$(FORMATNUM$(SUMA#(KI&), FORSUMA$(KI&)))
'              Label4(KI& - 1).Refresh
'            End If
'          End If
'          PPX& = PPX& + GRID.ColWidth(KI& - 1)
'       Next KI&
    End If
End Sub
Sub Search()

   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", , ARBUS$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     For U& = 1 To GRID.Rows - 1
      For i& = 0 To GRID.Cols - 1
       If InStr(UCase$(GRID.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
        GRID.Row = U&: GRID.COL = i&
        If GRID.TopRow > GRID.Row Then GRID.TopRow = GRID.Row
        On Error Resume Next
        VISIROWS = (GRID.Height / GRID.RowHeight(J&)) - 4
        If Err > 0 Then
         VISIROWS = 1
        End If
        On Error GoTo 0
         If VISIROWS > 0 Then
            If GRID.TopRow < GRID.Row - VISIROWS Then
              GRID.TopRow = GRID.Row - VISIROWS
            End If
         End If
        GRID.SetFocus
        Exit Sub
       End If
       Next i&
     Next U&
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If

End Sub

Sub FILTROVAR_VIE()
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
        FilText1(i& - 1).Visible = False
        FilText1(i& - 1).ToolTipText = ""
        FiltEco1(i& - 1).Visible = False
    Next i&
    '
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
             JK& = JK& + 1
             LABPRE$ = Mid$(VAREFE$, 2, Len(VAREFE$) - 2)
             ASELEC$ = ""
             PPXX% = InStr(LABPRE$, "/")
             '
             If PPXX% > 1 Then
                ASELEC$ = Mid$(LABPRE$, PPXX% + 1)
                LABPRE$ = Left$(LABPRE$, PPXX% - 1)
             End If
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
             Label1(i& - 1).Top = PVER + 400 * (JK& - 1)
             Combo1(i& - 1).Top = PVER - 50 + 400 * (JK& - 1)
             Command2(i& - 1).Top = PVER - 50 + 400 * (JK& - 1)
             Label1(i& - 1) = LABPRE$
             Combo1(i& - 1).Clear
             If Label1(i& - 1).Width > TAMAXI Then TAMAXI = Label1(i& - 1).Width
             Label1(i& - 1).Visible = True
             Combo1(i& - 1).Visible = True
             Command2(i& - 1).Visible = True
             FilText1(i& - 1).Visible = False
             FilText1(i& - 1).ToolTipText = ASELEC$
             FiltEco1(i& - 1).Visible = False
             ULTIFIL& = i&
             '
             If APLINVO$ <> "" Then
               If FILTX$ <> "" Then
                 PPXX% = InStr(FILTX$, ";")
                 If PPXX% < 1 Then PPXX% = InStr(FILTX$, ")")
                 If PPXX% < 1 Then PPXX% = 1 + Len(FILTX$)
                 If PPXX% > 1 Then
                    Combo1(i& - 1).TEXT = Left$(FILTX$, PPXX% - 1)
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
    For i& = 1 To CACONDI%
30     If i& > Label1.Count Then
          NUPIC% = Label1.Count
          Load Label1(NUPIC%)
          Load Combo1(NUPIC%)
          Load Command2(NUPIC%)
          Load FilText1(NUPIC%)
          Load FiltEco1(NUPIC%)
          GoTo 20
       End If
       '
       Label1(i& - 1).Left = 500 + TAMAXI - Label1(i& - 1).Width
       Combo1(i& - 1).Left = Label1(i& - 1).Left + Label1(i& - 1).Width + 250
       FilText1(i& - 1).Left = Label1(i& - 1).Left + Label1(i& - 1).Width + 250
       Command2(i& - 1).Left = Combo1(i& - 1).Left + Combo1(i& - 1).Width + 150
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
    'Frame1.Width = Command2(ULTIFIL& - 1).Left + Command2(ULTIFIL& - 1).Width + 450
    Frame1.Height = Command2(ULTIFIL& - 1).Top + 600
    '
    Command3.Top = 130
    Command3.Height = Frame1.Height - 150
    Command3.Left = Frame1.Width - Command3.Width
    '
90  GRID.Top = Frame1.Top + Frame1.Height + 250
    GRID.Height = Height - GRID.Top - 500
    If CACONFIL% < 1 Then
       Call Command3_Click
       Exit Sub
    End If
    '
    If APLINVO$ <> "" Then
      For i& = 1 To CACONDI%
        If Combo1(i& - 1).Visible = True Then
          If TRIM$(Combo1(i& - 1).TEXT) = "" Then Exit Sub
        End If
      Next i&
      Call Command3_Click
      Exit Sub
    End If
End Sub

Sub MUESTRATOTS()
    '
    PPX& = GRID.Left + GRID.ColPos(0)
    For KI& = 1 To GRID.Cols
       If KI& > Label4.Count Then
          NUPIC% = Label4.Count
          Load Label4(NUPIC%)
       End If
       Label4(KI& - 1).Caption = ""
       Label4(KI& - 1).Visible = False
       Label4(KI& - 1).Top = GRID.Top + GRID.Height + 100
       Label4(KI& - 1).Left = PPX& + 30
       Label4(KI& - 1).Width = GRID.ColWidth(KI& - 1)
       If TIESUM%(KI&) = 1 Then
         If FORSUMA$(KI&) <> "" Then
           Label4(KI& - 1).Visible = True
           Label4(KI& - 1).Caption = TRIM$(FORMATNUM$(SUMA#(KI&), FORSUMA$(KI&)))
           Label4(KI& - 1).Refresh
         End If
       End If
       PPX& = PPX& + GRID.ColWidth(KI& - 1)
    Next KI&
    '
End Sub

Sub EJEC_SUBAP()
    '
    If SUBAPLI$ = "PRINT" Then
          Call IMPRIMIR
       Else
          Call ABRETACON
          CONDI$ = "COD_QUERY = '" & APLINVO$ & "' "
          CONDI$ = CONDI$ & " AND Cod_Graph = '" & SUBAPLI$ & "'"
          TEXGRA$ = VALOTACO("SQL_GRAFICS", "PAR_GRAPH", CONDI$, "NOMESS")
          '
          If TEXGRA$ <> "" Then   ' ES UN GRAFICO
               DEFGRAF07.Label11 = APLINVO$
               DEFGRAF07.Combo6.TEXT = SUBAPLI$
               Call DEFGRAF07.Combo6_Click
'''               Call DEFGRAF07.EJECUTAR_Click
             Else
               TTXX$ = "Error de Invocación.\No Existe Gráfico '" + TRIM$(SUBAPLI$) + "'\en Reporte '" + APLINVO$ + "'."
               Call MENSERR(24, TTXX$)
               Call FINAL("")
          End If
    End If
    '
End Sub

'''Sub CARGAIMAGENENCELDA(FILA&, COLU&, TIMAGEN, FORCAX$)
'''    '
'''    ARIMAG$ = LUCOM$ + "CARPETMP\MG" + RANDSTRING$(-6) + "." + FORCAX$
'''    TIMAG$ = TIMAGEN
'''    Call SAVEFILE(ARIMAG$, TIMAG$)
'''    '
'''    If ARFOR$ <> "" Then
'''       GRID.TextMatrix(FILA&, COLU&) = ARIMAG$
'''       Exit Sub
'''    End If
'''    '
'''    FORMALTER.Image3.Picture = LoadPicture("")
'''    FORMALTER.Image3.Picture = LoadPicture(ARIMAG$)
'''    HIMAGEN = FORMALTER.Image3.Picture.Width
'''    VIMAGEN = FORMALTER.Image3.Picture.Height
'''    '
'''    HCALCU = GRID.ColWidth(COLU) * VIMAGEN / HIMAGEN
'''    If HCALCU > GRID.RowHeight(FILA) Then
'''       GRID.RowHeight(FILA) = HCALCU
'''    End If
'''    GRID.Col = COLU&
'''    GRID.Row = FILA&
'''    Picture2.Width = GRID.CellWidth
'''    Picture2.Height = GRID.CellHeight
'''    HCUADRO = Picture2.Width
'''    VCUADRO = Picture2.Height
'''    '
'''    REDUH = HIMAGEN / HCUADRO
'''    REDUV = VIMAGEN / VCUADRO
'''    REDUC = REDUH: If REDUV < REDUH Then REDUC = REDUV
'''    '
'''    Picture2.Cls
'''    Picture2.PaintPicture FORMALTER.Image3.Picture, 0, 0, HIMAGEN / REDUC, VIMAGEN / REDUC ' Picture2.Height, Picture2.Height
'''    Set GRID.CellPicture = Picture2.Image
'''    '
'''    On Error Resume Next
'''    Kill ARIMAG$
'''    On Error GoTo 0
'''    '
'''End Sub

Sub PONEIMAGENENCELDA(NGRILLA, Fila&, COLU&, FORCAX$, TIMAGEN)
    '
    If TIMAGEN = "" Then
       If EXISTE%(FORCAX$) > 0 Then
          ARIMAG$ = FORCAX$
          TIMAG$ = LOADFILE$(ARIMAG$)
          GoTo 10
       End If
    End If
    '
    EXTE$ = FORCAX$
5   PPXX% = InStr(EXTE$, ".")
    EXTE$ = Mid$(EXTE$, PPXX% + 1)
    If InStr(EXTE$, ".") > 0 Then GoTo 5
    '
    On Error Resume Next
    MkDir LUCOM$ + "CARPETMP"
    On Error GoTo 0
    ARIMAG$ = LUCOM$ + "CARPETMP\MG" + RANDSTRING$(-6) + "." + EXTE$
    TIMAG$ = TIMAGEN
    Call SAVEFILE(ARIMAG$, TIMAG$)
    '
10  If CARGAIMAGENCOMOLINK% = 1 Then
       NGRILLA.TextMatrix(Fila&, COLU&) = Space$(1024) + ARIMAG$
       Exit Sub
    End If
    '
    FORMALTER.Image3.Picture = LoadPicture("")
    On Error Resume Next
    FORMALTER.Image3.Picture = LoadPicture(ARIMAG$)
    If Err Then
       On Error GoTo 0
       Exit Sub
    End If
    On Error GoTo 0
    HIMAGEN = FORMALTER.Image3.Picture.Width
    VIMAGEN = FORMALTER.Image3.Picture.Height
    '
    HCALCU = NGRILLA.ColWidth(COLU&) * VIMAGEN / HIMAGEN
    If HCALCU > NGRILLA.RowHeight(Fila&) Then
       NGRILLA.RowHeight(Fila) = HCALCU
    End If
    NGRILLA.COL = COLU&
    NGRILLA.Row = Fila&
    FORMALTER.Picture2.Width = NGRILLA.CellWidth
    FORMALTER.Picture2.Height = NGRILLA.CellHeight
    HCUADRO = FORMALTER.Picture2.Width
    VCUADRO = FORMALTER.Picture2.Height
    '
    REDUH = HIMAGEN / HCUADRO
    REDUV = VIMAGEN / VCUADRO
    REDUC = REDUH: If REDUV < REDUH Then REDUC = REDUV
    '
    FORMALTER.Picture2.Cls
    FORMALTER.Picture2.PaintPicture FORMALTER.Image3.Picture, 0, 0, HIMAGEN / REDUC, VIMAGEN / REDUC ' Picture2.Height, Picture2.Height
    Set NGRILLA.CellPicture = FORMALTER.Picture2.Image
    '
    On Error Resume Next
    Kill ARIMAG$
    On Error GoTo 0
    '
End Sub

'''Sub CARGATEXTOMULTILINEA(FILA&, COLU&)
'''    '
'''    GRID.Col = COLU&
'''    GRID.Row = FILA&
'''    Picture2.Width = GRID.CellWidth
'''    Picture2.Height = GRID.CellHeight
'''    '
'''10  Picture2.Cls
'''    Picture2.CurrentY = 0
'''    Picture2.FontName = GRID.CellFontName
'''    Picture2.FontSize = GRID.CellFontSize
'''    Picture2.FontBold = GRID.CellFontBold
'''    Picture2.BackColor = RGB(255, 255, 255) ' BLANCO
'''    '
'''    TTTXX$ = ""
'''    For KIKI% = 1 To CARETEX%
'''       Picture2.Print RETEX$(KIKI%)
'''       If Picture2.CurrentY > Picture2.Height Then
'''          Picture2.Height = Picture2.CurrentY
'''       End If
'''       If TTTXX$ > "" Then TTTXX$ = TTTXX$ + vbCrLf
'''       TTTXX$ = TTTXX$ + RETEX$(KIKI%)
'''       ' If KIKI% > 1 Then MULTILIN% = 1
'''    Next KIKI%
'''    '
'''    If Picture2.Height > GRID.RowHeight(FILA&) Then
'''       GRID.RowHeight(FILA&) = Picture2.Height
'''       GoTo 10
'''    End If
'''    '
'''    GRID.TextMatrix(FILA&, COLU&) = Space$(1024) + TTTXX$   ' SPACE$(1024) delante para que no sea visible
'''    Picture2.Refresh
'''    Set GRID.CellPicture = Picture2.Image
'''    '
'''    '
'''End Sub

Sub PONETEXMULINENCELDA(NGRILLA, Fila&, COLU&)
    '
    NGRILLA.COL = COLU&
    NGRILLA.Row = Fila&
    FORMALTER.Picture2.Width = NGRILLA.CellWidth
    FORMALTER.Picture2.Height = NGRILLA.CellHeight
    '
10  FORMALTER.Picture2.Cls
    FORMALTER.Picture2.CurrentY = 0
    FORMALTER.Picture2.FontName = NGRILLA.CellFontName
    FORMALTER.Picture2.FontSize = NGRILLA.CellFontSize
    FORMALTER.Picture2.FontBold = NGRILLA.CellFontBold
    FORMALTER.Picture2.BackColor = RGB(255, 255, 255) ' BLANCO
    '
    TTTXX$ = ""
    For KIKI% = 1 To CARETEX%
       FORMALTER.Picture2.Print RETEX$(KIKI%)
       If FORMALTER.Picture2.CurrentY > FORMALTER.Picture2.Height Then
          FORMALTER.Picture2.Height = FORMALTER.Picture2.CurrentY
       End If
       If TTTXX$ > "" Then TTTXX$ = TTTXX$ + vbCrLf
       TTTXX$ = TTTXX$ + RETEX$(KIKI%)
       ' If KIKI% > 1 Then MULTILIN% = 1
    Next KIKI%
    '
    If FORMALTER.Picture2.Height > NGRILLA.RowHeight(Fila&) Then
       NGRILLA.RowHeight(Fila&) = FORMALTER.Picture2.Height
       GoTo 10
    End If
    '
    NGRILLA.TextMatrix(Fila&, COLU&) = Space$(1024) + "///TXT =" + TTTXX$ ' SPACE$(1024) delante para que no sea visible
    FORMALTER.Picture2.Refresh
    Set NGRILLA.CellPicture = FORMALTER.Picture2.Image
    '
End Sub

Sub IMPRIMELAIMAGEN(NGRILLA, Fila&, COLU&)
    '
    CURRX = Printer.CurrentX
       CURRX = CURRX + Printer.TextWidth(" ") ' PARA DEJAR UNA COLUMNA EN BLANCO
    CURRY = Printer.CurrentY
    FORMALTER.Image3.Picture = LoadPicture("")
    NGRILLA.Row = Fila&
    NGRILLA.COL = COLU&
    '
    Set FORMALTER.Image3.Picture = NGRILLA.CellPicture
    HIMAGEN = FORMALTER.Image3.Picture.Width
    VIMAGEN = FORMALTER.Image3.Picture.Height
    '
    FORMALTER.Picture2.Width = NGRILLA.CellWidth
    FORMALTER.Picture2.Height = NGRILLA.CellHeight
'      COEFI = Printer.TextWidth(Space$(ACOLUM%(COLU& + 1))) / FORMALTER.Picture2.Width
      coefi = NGRILLA.ColWidth(COLU&) * COEFPRI(NGRILLA) / FORMALTER.Picture2.Width
    HCUADRO = FORMALTER.Picture2.Width * coefi  ' OJO QUE NO SE SI ESTO NO EMPEORA
    VCUADRO = FORMALTER.Picture2.Height * coefi
    '
    If VCUADRO > ALTUREMAX Then
        ALTUREMAX = VCUADRO
    End If

    REDUC = HIMAGEN / HCUADRO    ' LA VERTICAL SE EXTIENDE LO NECESARIO PARA MANTENER LA PROPORCION - EPB
    If REDUC > 0 Then
        Printer.PaintPicture FORMALTER.Image3.Picture, CURRX, CURRY, HIMAGEN / REDUC, VIMAGEN / REDUC ' FORMALTER.PICTURE2.Height, FORMALTER.PICTURE2.Height
    End If
    '
    Printer.CurrentY = CURRY
'    Printer.CurrentX = CURRX + Printer.TextWidth(Space$(1 + ACOLUM%(COLU& + 1)))
    Printer.CurrentX = CURRX + NGRILLA.ColWidth(COLU&) * COEFPRI(NGRILLA)

End Sub

Sub IMPRIMEMULIN(NGRILLA, Fila&, COLU&)
    '
    CURRX = Printer.CurrentX
       CURRX = CURRX + Printer.TextWidth(" ")       ' PARA DEJAR UNA COLUMNA EN BLANCO
    CURRY = Printer.CurrentY
    NGRILLA.Row = Fila&
    NGRILLA.COL = COLU&
    TTXX$ = LTrim$(Mid$(NGRILLA.TEXT, 1025))
    ' MsgBox TTXX$
    If Left$(TTXX$, 8) = "///TXT =" Then TTXX$ = Mid$(TTXX$, 9)
    Call FRATEXTO(TTXX$, ACOLUM%(COLU& + 1))
    For KKII% = 1 To CARETEX%
       Printer.CurrentX = CURRX
       Printer.Print RETEX$(KKII%)
       If Printer.CurrentY - CURRY > ALTUREMAX Then
          ALTUREMAX = Printer.CurrentY - CURRY
       End If
    Next KKII%
    '
    Printer.CurrentY = CURRY
    Printer.CurrentX = CURRX + Printer.TextWidth(Space$(ACOLUM%(COLU& + 1)))
    '
End Sub

Function COEFPRI(NGRILLA)
    '
    APRIX = Printer.Width
    AGRID = 0
    For KKI% = 1 To NGRILLA.Cols
       AGRID = AGRID + NGRILLA.ColWidth(KKI% - 1)
    Next KKI%
    COEFPRI = APRIX / AGRID
    If COEFPRI > 1 Then COEFPRI = 1           ' PARA QUE NO ESTIRE HORIZONTALMENTE
    '
End Function
Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
      strContenidoCelda$ = GRID.TextMatrix(GRID.MouseRow, GRID.MouseCol)
      If MsgBox(strContenidoCelda$, vbYesNo, "Copiar Contenido de Celda Seleccionada") = vbYes Then
       Clipboard.Clear
       Clipboard.SetText strContenidoCelda
      End If
    End If
End Sub

Sub CopyToExcel(InFlexGrid As MSHFlexGrid, Nome$, _
      ByVal TextoAdicional$)
  Dim R%, c%, Buf$, LstRow%, LstCol%
  Dim FormatMoney As Boolean
  Dim MyExcel As Excel.Application
  Dim wbExcel As Excel.Workbook
  Dim shExcel As Excel.Worksheet
  On Error Resume Next

  Set MyExcel = GetObject(, "Excel.Application")
  If Err.Number <> 0 Then
 Set MyExcel = CreateObject("Excel.Application")
  End If
  Set wbExcel = MyExcel.Workbooks.Add
  Set shExcel = wbExcel.Worksheets.Add
  shExcel.Name = Nome$
  shExcel.Activate
  LstCol% = 0
  For c% = 0 To InFlexGrid.Cols - 1
 InFlexGrid.COL = c%
 LstRow% = 0
 shExcel.Columns(Chr(Asc("A") + c%)).ColumnWidth = InFlexGrid.ColWidth(c%) / 72
 For R% = 0 To InFlexGrid.Rows - 1
   InFlexGrid.Row = R%
   Err.Clear
   Buf$ = InFlexGrid.TextMatrix(R%, c%)
   If Buf$ <> "" Then
  FormatMoney = False
  If InStr(Buf$, vbCrLf) Then
    Buf$ = StrTran(Buf$, vbCrLf, vbLf)
    Do While Right(Buf$, 1) = vbLf
   Buf$ = Left(Buf$, Len(Buf$) - 1)
    Loop
    shExcel.Range(Chr(Asc("A") + c%)).WrapText = True
  ElseIf Format(CDbl(Buf$), csFormatMoneyZero) = Buf$ Then
    If Err.Number = 0 Then
   Buf$ = Str(CDbl(Buf$))
   FormatMoney = True
    End If
  End If
  If Buf$ <> "" Then
    If InFlexGrid.MergeRow(R%) Then
   For LstCol% = c% To 1 Step -1
     If InFlexGrid.TextMatrix(R%, LstCol% - 1) <> InFlexGrid.TextMatrix(R%, c%) Then
    Exit For
     End If
   Next
   If LstCol% <> c% Then
     shExcel.Range(Chr(Asc("A") + LstCol%) & (R% + 1), _
         Chr(Asc("A") + c%) & (R% + 1)).MergeCells = True
     shExcel.Range(Chr(Asc("A") + LstCol%) & (R% + 1), _
         Chr(Asc("A") + c%) & (R% + 1)).BorderAround
   End If
    End If
    If InFlexGrid.MergeCol(c%) And LstRow% <> R% Then
   If InFlexGrid.TextMatrix(LstRow%, c%) = InFlexGrid.TextMatrix(R%, c%) Then
     shExcel.Range(Chr(Asc("A") + c%) & (LstRow% + 1), _
         Chr(Asc("A") + c%) & (R% + 1)).MergeCells = True
     shExcel.Range(Chr(Asc("A") + c%) & (LstRow% + 1), _
         Chr(Asc("A") + c%) & (R% + 1)).BorderAround
   Else
     LstRow% = R%
   End If
    End If
    shExcel.Range(Chr(Asc("A") + c%) & _
        (R% + 1)).Font.COLOR = InFlexGrid.CellForeColor
    If R% < InFlexGrid.FixedRows Or c% < InFlexGrid.FixedCols Then
   shExcel.Range(Chr(Asc("A") + c%) & _
       (R% + 1)).Font.Bold = True
    shExcel.Range(Chr(Asc("A") + c%) & _
        (R% + 1)).Font.BackColor = 40
    End If
    shExcel.Range(Chr(Asc("A") + c%) & (R% + 1)).Value = Buf$
    If FormatMoney Then
   shExcel.Range(Chr(Asc("A") + c%) & _
       (R% + 1)).NumberFormat = "#,##0.00;#,##0.00;#,##0.00"
    End If
  End If
   End If
 Next
  Next
  If TextoAdicional$ <> "" Then
 ' shExcel.Rows(Str(r%+2)).Delete (xlShiftUp)
 Do While Right(TextoAdicional$, 1) = vbLf
   TextoAdicional$ = Left(TextoAdicional$, _
         Len(TextoAdicional$) - 1)
 Loop
 shExcel.Range("A" & (R% + 2)).Value = TextoAdicional$
  End If
  MyExcel.Visible = True
  Set shExcel = Nothing
  Set wbExcel = Nothing
  Set MyExcel = Nothing
End Sub
Public Function StrTran(Cadena As String, BUSCAR As String, Sustituir As String, Optional Veces As Variant) As String
   Dim Contador As Integer
 
Dim Resultado As String
   Dim Cambios As Integer
 

   Resultado = ""
   Cambios = 0
 
   For Contador = 1 To Len(Cadena)
   If Mid(Cadena, Contador, Len(BUSCAR)) = BUSCAR Then
  
Resultado = Resultado & Sustituir
   If Len(BUSCAR) > 1 Then
     
Contador = Contador + Len(BUSCAR) - 1
   End If
  

   ' si se especifica un nº de cambios determinados
   If Not IsMissing(Veces) Then
     
Cambios = Cambios + 1
   If Cambios = Veces Then
    
Resultado = Resultado & Mid(Cadena, Contador + 1)
    
Exit For
     
End If
End If
   If Len(BUSCAR) > 1 Then
     
Contador = Contador + Len(BUSCAR) - 1
   End If
 
Else
   Resultado = Resultado & Mid(Cadena, Contador, 1)
   End If
   Next
 
   StrTran = Resultado
End Function

 

