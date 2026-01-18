VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MENUEXP09 
   BackColor       =   &H00C0C0C0&
   Caption         =   "FLEXOFT - Sistema de Carpetas de Exportacion"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11880
   Icon            =   "MENUEXP09.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   11880
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8600
      Left            =   11025
      ScaleHeight     =   8535
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   0
      Width           =   1275
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         Height          =   2115
         Left            =   0
         ScaleHeight     =   2115
         ScaleWidth      =   1275
         TabIndex        =   21
         Top             =   6195
         Width           =   1275
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   90
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   25
            Top             =   945
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   26
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.Image Image4 
            Height          =   390
            Left            =   -315
            Picture         =   "MENUEXP09.frx":74F2
            Top             =   1185
            Width           =   1515
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "00:00:00"
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
            Left            =   35
            TabIndex        =   24
            Top             =   1680
            Width           =   750
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "00/00/00"
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
            Left            =   35
            TabIndex        =   23
            Top             =   525
            Width           =   750
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Miercoles"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   22
            Top             =   210
            Width           =   795
         End
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "MENUEXP09.frx":9414
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Generador de Reportes - Consultas SQL"
         Top             =   4320
         Width           =   600
      End
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "MENUEXP09.frx":971E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cerrar Sesión - Cambiar Usuario / Entorno"
         Top             =   700
         Width           =   600
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "MENUEXP09.frx":9868
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2810
         Width           =   600
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "MENUEXP09.frx":9CAA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cambio de Clave Personal de Acceso"
         Top             =   2205
         Width           =   600
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "MENUEXP09.frx":9FB4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Manuales de Uso y Operacion"
         Top             =   1600
         Width           =   600
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "MENUEXP09.frx":A2BE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Terminar - Cerrar Sesion de Trabajo"
         Top             =   105
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H80000001&
      Height          =   8500
      Left            =   -10
      ScaleHeight     =   8445
      ScaleWidth      =   3090
      TabIndex        =   17
      Top             =   -50
      Width           =   3150
      Begin VB.PictureBox Picture19 
         Height          =   855
         Left            =   -240
         ScaleHeight     =   795
         ScaleWidth      =   3000
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   3060
         Begin VB.Image Image11 
            Height          =   1590
            Left            =   0
            Top             =   0
            Width           =   2010
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         Caption         =   "Frame1"
         Height          =   8415
         Left            =   3045
         TabIndex        =   19
         Top             =   -100
         Width           =   75
      End
      Begin ComctlLib.TreeView Treeview1 
         Height          =   2400
         Left            =   -30
         TabIndex        =   18
         Top             =   2040
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   4233
         _Version        =   327682
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imlSmallIcons"
         Appearance      =   1
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
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Menu Principal"
         Height          =   260
         Left            =   0
         TabIndex        =   28
         Top             =   840
         Width           =   2850
      End
      Begin ComctlLib.ImageList imlIcons 
         Left            =   600
         Top             =   4185
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   1
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MENUEXP09.frx":A408
               Key             =   "book"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame1"
      ForeColor       =   &H00404040&
      Height          =   75
      Left            =   0
      TabIndex        =   30
      Top             =   7140
      Width           =   15000
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0080C0FF&
      Height          =   2000
      Left            =   3120
      ScaleHeight     =   1935
      ScaleWidth      =   15945
      TabIndex        =   29
      Top             =   5880
      Width           =   16000
      Begin VB.Image Image3 
         Height          =   480
         Index           =   0
         Left            =   210
         Picture         =   "MENUEXP09.frx":A722
         Top             =   210
         Width           =   480
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3465
      Top             =   3780
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4620
      Top             =   1785
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00C0C0C0&
      Height          =   1300
      Left            =   8925
      Picture         =   "MENUEXP09.frx":AB64
      ScaleHeight     =   1245
      ScaleWidth      =   1770
      TabIndex        =   1
      Top             =   105
      Visible         =   0   'False
      Width           =   1830
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   165
         TabIndex        =   7
         Top             =   420
         Width           =   225
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   780
         TabIndex        =   6
         Top             =   900
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1460
         TabIndex        =   5
         Top             =   470
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   800
         TabIndex        =   4
         Top             =   60
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   1380
      Left            =   3720
      ScaleHeight     =   1320
      ScaleWidth      =   2370
      TabIndex        =   20
      ToolTipText     =   "Click con Boton Derecho para Cambiar la Imagen"
      Top             =   360
      Visible         =   0   'False
      Width           =   2430
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
      Left            =   2520
      Picture         =   "MENUEXP09.frx":122F6
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Depósito Caotico SAPORITI"
      Top             =   6195
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   2625
      Pattern         =   "*.EXK"
      TabIndex        =   14
      Top             =   210
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   420
      Top             =   0
   End
   Begin VB.PictureBox Picture9 
      Height          =   2010
      Left            =   420
      ScaleHeight     =   1950
      ScaleWidth      =   2055
      TabIndex        =   12
      Top             =   210
      Visible         =   0   'False
      Width           =   2115
      Begin VB.Image Image1 
         Height          =   1590
         Left            =   0
         Top             =   0
         Width           =   2010
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5500
      Top             =   0
   End
   Begin VB.CommandButton Command12 
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
      Left            =   1785
      Picture         =   "MENUEXP09.frx":12440
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Control y Movimientos de Stocks"
      Top             =   5040
      Width           =   600
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
      Height          =   600
      Left            =   630
      Picture         =   "MENUEXP09.frx":1274A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Actualización y Consulta de Archivos Maestros"
      Top             =   3360
      Width           =   600
   End
   Begin VB.Image Image2 
      Height          =   5760
      Left            =   6120
      Picture         =   "MENUEXP09.frx":12A54
      Top             =   105
      Visible         =   0   'False
      Width           =   8250
   End
   Begin ComctlLib.ImageList imlSmallIcons 
      Left            =   5000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MENUEXP09.frx":1BCAA
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MENUEXP09.frx":1C1CC
            Key             =   "cylinder"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MENUEXP09.frx":1C6EE
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MENUEXP09.frx":1CC10
            Key             =   "open"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MENUEXP09.frx":1D132
            Key             =   "smlBook"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MENUEXP09.frx":1D794
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MENUEXP09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARRABOR%

Sub CONSULCOLIMO()
    Call FINAL("?LISMOVRANFE")
End Sub

Sub CONSULFKARDEX()
    FIKARDEX09.Show 1
End Sub

Sub CONSULGESINCO()
    Call CONECRUN("INVENT07", "Inventarios Físicos y Valorizados")
End Sub

Sub CONSULLIQUIDA()
    Call FINAL("?LIQUICOS")
End Sub

Sub CONSULLPROVENC()

    OPX% = ALTERNA%("Por Fecha de Vencimiento\Por Fecha de Retest")
    Select Case OPX%
       Case 1
          FF% = DIPOST(HOY(HO$) + 29)
          AA$ = FECHATEX(FF%)
          Call FINAL("?LOTEPORVENCE(" + TRIM$(AA$) + ")")
       Case 2
          FF% = DIPOST(HOY(HO$) + 29)
          AA$ = FECHATEX(FF%)
          Call FINAL("?LOTPORVENRET(" + TRIM$(AA$) + ")")
          Call FINAL("?MACONSXOCTTT(" & NOCOM$ & ")")
    End Select
End Sub

Sub CONSULVENPORCARP()
    '
    CONDI$ = "CodcarpImp <>'' AND STATUS < 9 AND MODO = 1 ORDER BY CodCarpImp"
    Call CARGALISEL("!INCARPIM", "CARPEIMP", "CodcarpImp/A12;FechAlta/d8;RefeCarpImp/A48", CONDI$, "NOMESS")
    Call SELECHO("!INCARPIM/INDICE DE CARPETAS DE IMPORTACION")
    '
    CODCARPE$ = TRIM$(VALACT1$("!INCARPIM"))
    If CODCARPE$ <> "" Then
        Call FINAL("?VENPORCARP(" & CODCARPE$ & ")")
    End If
End Sub

Sub CONSULPPSTOCK()
    Call FINAL("?UBICALOTES( )") ' Por defecto no carga ningun deposito
End Sub

Sub CONSULTCODCOM()
    Call FINAL("?REXIPORPROD")
End Sub

Sub INGREID()
     OPCOTEL.Show 1
End Sub

Sub INSPYAPROCOMPD()
    APRODES09.Show 1
End Sub

Sub LIQUIDAPORCCOSTO()
    
    CONDI$ = "CodcarpImp <>'' AND STATUS < 9 AND MODO = 1"
    CONDI$ = CONDI$ + "ORDER BY CodCarpImp"
    Call CARGALISEL("!INCARPIM", "CARPEIMP", "CodcarpImp/A12;FechAlta/d8;RefeCarpImp/A48", CONDI$, "NOMESS")
    Call SELECHO("!INCARPIM/INDICE DE CARPETAS DE IMPORTACION")
    '
    CODCARPE$ = TRIM$(VALACT1$("!INCARPIM"))
    '
    If CODCARPE$ <> "" Then
       CONDI$ = "CodcarpImp ='" & CODCARPE$ & "' AND STATUS < 9"
       If CantRegistros("CARPEIMP", CONDI$) > 0 Then
          NPRO# = XVALO(VALOBADA("CARPEIMP", "CodCenCos", CONDI$, "NOMESS"))
       End If
    End If
    '
    SQLX$ = "SELECT * FROM CAJABANC with(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE CenCosto = " & NPRO# & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'COMPRA' OR TipoMovim = 'OPAG')"
    '
    Dim LIQPORCCOSTO As New ADODB.Recordset
    Set LIQPORCCOSTO = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    CONDI$ = "CodCenCos ='" & NPRO# & "' AND ImpoOrig <> 0"
    Call BORRAREGISTROS("LIQUIDACOS", CONDI$, REG&, "NOMESS")
    '
    SQLX$ = "SELECT * FROM LIQUIDACOS"
    SQLX$ = SQLX$ + " WHERE CodCenCos = " & NPRO# & " "
    Set LIQUI = New ADODB.Recordset
    LIQUI.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    With LIQPORCCOSTO
    '
        Do While Not .EOF
            CCOS$ = SAFETEXT(!CenCosto)
            COMPROB$ = SAFETEXT$(!CodiCom_Lar)
            IMPO# = XVALO(!ValAbsComp)
            FECHA% = CVINT(!FechEmisi)
            CONDI$ = "CodCenCos ='" & CCOS$ & "' AND STATUS < 9"
            If CantRegistros("CARPEIMP", CONDI$) > 0 Then
                 MONERE% = XVALO(VALOBADA("CARPEIMP", "Moneda", CONDI$, "NOMESS"))
            End If
            MONEMIS% = XVALO(!Mone_Emis)
            IMPORCONV = 0: TIPOCAM# = 0
            If TICAMONE(MONERE%) > 0 Then
                IMPORCONV = (IMPO# * TICAMONE(MONEMIS%) / TICAMONE(MONERE%))
                TIPOCAM# = TICAMONE(MONERE%)
            End If
            With LIQUI
              'Grabacion en Tabla LIQUIDACOS
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next
              !CodiEmpr = CodiEmp%
              On Error GoTo 0
              !CODCENCOS = CCOS$
              !FECHA = CVDAT(FECHA%)
              !CONCEPTO = COMPROB$
              !MONEORIG = XVALO(MONEMIS%)
              !MONEREDU = XVALO(MONERE%)
              !TIPOCAM = TIPOCAM#
              !IMPOORIG = IMPO#
              !IMPOREDU = IMPORCONV
              .Update
            End With
           .MoveNext
        Loop
        LIQUI.Close
        Set LIQUI = Nothing
    End With
    LIQPORCCOSTO.Close
    Set LIQPORCCOSTO = Nothing
    '
    
    SQLX$ = "SELECT * FROM DETAASIEN with(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE CenCos = " & NPRO# & " "
    '
    Dim LIQUID As New ADODB.Recordset
    Set LIQUID = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    SQLX$ = "SELECT * FROM LIQUIDACOS"
    SQLX$ = SQLX$ + " WHERE CodCenCos = " & NPRO# & " "
    Set LIQUIDA = New ADODB.Recordset
    LIQUIDA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    With LIQUID
    '
        Do While Not .EOF
            CCOS$ = SAFETEXT(!CenCos)
            CONCEPTO$ = SAFETEXT$(!NuIdAsien) & " - " & SAFETEXT$(!RefePase)
            CONDI$ = "CodCenCos ='" & CCOS$ & "' AND STATUS < 9"
            If CantRegistros("CARPEIMP", CONDI$) > 0 Then
                 MONERE% = XVALO(VALOBADA("CARPEIMP", "Moneda", CONDI$, "NOMESS"))
            End If
            IDEBE# = XVALO(!IDEBEPASE)
            IHABE# = XVALO(!IHABEPASE)
            IMPO# = IDEBE# - IHABE#
            FECHA% = CVINT(!FECHASIEN)
            '
            MONEMIS% = 1
            IMPORCONV = 0: TIPOCAM# = 0
            If TICAMONE(MONERE%) > 0 Then
                IMPORCONV = (IMPO# * TICAMONE(MONEMIS%) / TICAMONE(MONERE%))
                TIPOCAM# = TICAMONE(MONERE%)
            End If
            '
            With LIQUIDA
              'Grabacion en Tabla LIQUIDACOS
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next
              !CodiEmpr = CodiEmp%
              On Error GoTo 0
              !CODCENCOS = CCOS$
              !FECHA = CVDAT(FECHA%)
              !CONCEPTO = CONCEPTO$
              !MONEORIG = XVALO(MONEMIS%)
              !MONEREDU = XVALO(MONERE%)
              !TIPOCAM = TIPOCAM#
              !IMPOORIG = IMPO#
              !IMPOREDU = IMPORCONV
              .Update
            End With
           .MoveNext
        Loop
        LIQUIDA.Close
        Set LIQUI = Nothing
    End With
    LIQUID.Close
    Set LIQUID = Nothing
    
    If CantRegistros("LIQUIDACOS", "ImpoOrig > 0") < 1 Then
       Call COMUNI("No Se Ha Encontrado Informacion de Gastos")
       Exit Sub
    End If
End Sub

'
Private Sub Command1_Click()
   '
   Call GRAVALCON(USNBR%, "POSIFORPRIN", "TOP", FORMATNUM$(Top, "F10.0"))
   Call COMPACTABD
   '
   Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
   
   ' If PROCACTI% < 1 Then        ' ESTO SE SUPRIME PORQUE ES SOLO PARA LOS MODULOS DE
   '    Call GENBACKUP            ' CONTROL GENERAL - FLGEST07 - FLEXO07 - FLEXO09
   '    Call CANCELUCOM
   '    Call FINAL("")
   ' End If
   '
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   If EXISTE%("IFABRI07.EXE") > 0 Then
     Call CONECRUN("IFABRI07", "Informes de Fabricación")
   Else
     Call COMUNI("Falta módulo IFABRI07")
   End If
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   If EXISTE%("ARCHIG07.EXE") > 0 Then
     Call CONECRUN("ARCHIG07", "Archivos Maestros")
   Else
     Call COMUNI("Falta módulo ARCHIG07")
   End If
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   If EXISTE%("CONSTO07.EXE") > 0 Then
     Call CONECRUN("CONSTO07", "Control de Stocks")
   Else
     Call COMUNI("Falta módulo CONSTO07")
   End If
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   If EXISTE%("PROMRP07.EXE") > 0 Then
     Call CONECRUN("PROMRP07", "Programación de Necesidades MRP")
   Else
     Call COMUNI("Falta módulo PROMRP07")
   End If
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call CONECRUN("SOLIMA07", "Solicitud Interna de Materiales")
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
'   If EXISTE%("ORFABR04.EXE") > 0 Then
'     Call CONECRUN("ORFABR04", "Sistema de Ordenes de Fabricación")
'   Else
'     Call COMUNI("Falta módulo ORFABR04")
'   End If
  '\\\OT-05-0415-WAR-22/06/05///
   If EXISTE%("ORCONF07.EXE") > 0 Then
       Call CONECRUN("ORCONF07", "Sistema de Ordenes de Fabricación y Envios de Prendas")
     ElseIf EXISTE%("ORFABR07.EXE") > 0 Then
       Call CONECRUN("ORFABR07", "Sistema de Ordenes de Fabricación")
     ElseIf EXISTE%("MINORF07.EXE") > 0 Then
       Call CONECRUN("MINORF07", "Sistema Basico de Fabricación")
     Else
       Call COMUNI("Falta módulo ORFABR07")
   End If
   '\\\OT-05-0415-WAR-FIN///
   Command15.Enabled = True
End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   If PROCACTI% < 1 Then
      Call CANCELUCOM
      ' verifica si es multiempresa
      Call RECONN_EMPRE(OKXX%)
      If OKXX% < 1 Then GoTo 80
      TCONFI$ = LOADFILE$(CurDir + "\" + "FLEXCFG.DAT")
      If TCONFI$ = "" Then Exit Sub
      '
      LUDAT$ = TRIM$(Mid$(TCONFI$, 51, 45))
      If Right$(LUDAT$, 1) <> "\" Then
         LUDAT$ = LUDAT$ + "\"
         Mid$(TCONFI$, 51, 45) = Left$(LUDAT$ + Space$(45), 45)
         Call SAVEFILE("FLEXCFG.DAT", TCONFI$)
      End If
      If Mid$(LUDAT$, 2, 1) <> ":" Then
         ULO$ = Left$(TRIM$(LOADFILE$(LUCOM$ + "ULOGDAT.DRV")), 2)
         If ULO$ = "" Then ULO$ = Left$(CurDir$, 2)
         LUDAT$ = ULO$ + LUDAT$
      End If
      '
      RSTRI$ = LOADFILE$(LUCOM$ + "FLEXRAM.DRV") + String$(512, 0)
      EMPREAC$ = TRIM$(Mid$(RSTRI$, 193, 30))
         Call GRARAM(99, AJUSTI$(LUDAT$, 46))
         Call GRARAM(193, AJUSTI$(EMPREAC$, 30))
         Call GRARAM(191, Chr$(CodiEmp%))
         Call BAJRAM
      EMPREAC$ = EMPREAC$ + " (" + TRIM$(Str$(CodiEmp%)) + ") "
      EMPREAC$ = EMPREAC$ + VERNUM$ + " - " + USERNAME$
      Screen.MousePointer = 1
      If OKEY% < 1 Then Call FINAL("")
      EPAC$ = TRIM$(EMPREAC$)
      If EPAC$ <> "" Then
         Caption = "Sistema Preconfigurado FLEXOFT" + " - " + EPAC$
         Me.Refresh
      End If
      Call GRATITFOR(Caption)
      Call PONELOGO
      Call UPDATES
GoTo 99

80    ' conecta con FLOGON.EXE
      On Error Resume Next
      RETOR$ = TRIM$(UCase$(App.EXEName))
      Shell "FLOGON.EXE " + RETOR$, vbNormalFocus
      If Err Then
        Call MERRFINAL("Falta Utilitario 'FLOGON.EXE'.")
      End If
      Call FINAL("")
   End If
   '
99 Command16.Enabled = True
End Sub

Sub UPDATES()
   '
   Screen.MousePointer = 11
   Label2 = 171: DoEvents
'   Call LEEDATSYS
   '
   YAMENSA% = 0
   INSTALAC% = 0
   '
   DIREPRO$ = App.Path
   If EXISTE%("F:\FLEXOFT\NOVEDAD\.") > 0 Then
       DIRENEW$ = "F:\FLEXOFT\NOVEDAD"
     ElseIf EXISTE%("C:\FLEXOFT\NOVEDAD\.") > 0 Then
       DIRENEW$ = "C:\FLEXOFT\NOVEDAD"
     Else
       DIRENEW$ = PARENTFOLDER$(DIREPRO$) + "\NOVEDAD"
   End If
   DIRNU$ = DIRENEW$
10 On Error Resume Next
   File1.Path = DIRNU$
   If Err < 1 Then
     On Error GoTo 0
     File1.Pattern = "*.*"
     '
     For KU& = 1 To File1.ListCount
       ANAME$ = TRIM$(File1.List(KU& - 1))
       PPXX% = InStr(ANAME$, ".")
       If PPXX% < 1 Then PPXX% = 1 + Len(ANAME$)
       exten$ = TRIM$(UCase$(Mid$(ANAME$, PPXX% + 1)))
       '
       If Len(exten$) = 3 Then
         '
         If exten$ = "RFS" Or exten$ = "FRM" Then
           FECHANUE# = 0
           On Error Resume Next
           FECHANUE# = FileDateTime(DIRENEW$ + "\" + ANAME$)
           On Error GoTo 0
           FECHAVIE# = 0
           If EXISTE%(LUDAT$ + ANAME$) > 0 Then
             FECHAVIE# = FileDateTime(LUDAT$ + ANAME$)
           End If
           '
           If FECHANUE# > FECHAVIE# + 0.0001 Then  ' ESO SON APROX 6 SEGUNDOS
             On Error Resume Next
             FileCopy DIRENEW$ + "\" + ANAME$, LUDAT$ + ANAME$
             If Err Then
               MsgBox "Imposible Instalar (3) " + ANAME$
             End If
             On Error GoTo 0
           End If
           GoTo 29         ' SALE DE INSTALACION DE DATOS
         End If
         '
         RUTACOM$ = App.Path + "\"
         If EXISTE%(RUTACOM$ + ANAME$) > 0 Or exten$ = "PRF" Then
           On Error Resume Next
           FECHANUE# = FileDateTime(DIRENEW$ + "\" + ANAME$)
           If Err Then
              On Error GoTo 0
              GoTo 29
           End If
           On Error GoTo 0
           FECHAVIE# = 0
           If EXISTE%(RUTACOM$ + ANAME$) > 0 Then
             FECHAVIE# = FileDateTime(RUTACOM$ + ANAME$)
           End If
           If FECHANUE# > FECHAVIE# + 0.0001 Then  ' ESO SON APROX 6 SEGUNDOS
             If YAMENSA% = 0 Then
               MESSA$ = "Hay Actualizaciones Listas para Instalarse, des-" + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "de la Carpeta " + DIRENEW$ + "." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "Para Instalar las Actualizaciones Disponibles" + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "Cierre las Aplicaciones FLEXOFT en TODAS las" + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "Estaciones de Trabajo." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "¿ Desea Ahora Actualizar su Sistema FLEXOFT ?"
               RESPU% = MsgBox(MESSA$, vbYesNoCancel)
               If RESPU% <> 6 Then
                 Screen.MousePointer = 1
                 Exit Sub     ' 6 es = SI
               End If
               YAMENSA% = 1
             End If
             '
             EXTECO$ = Left$(EXTE$, 2)
             ANAMCO$ = Left$(ANAME$, Len(ANAME$) - 1)
             On Error Resume Next
             Kill RUTACOM$ + ANAMCO$ + "0"
             Kill DIREVIE$ + ANAMCO$ + "9"
             On Error GoTo 0
             If EXISTE%(DIREVIE$ + ANAMCO$ + "9") > 0 Then
               'MsgBox "Imposible Instalar (1) " + ANAME$
               MsgBox "Imposible Instalar. No Pudieron Borrarse " + RUTACOM$ + ANAMCO$ + "0 y " + DIREVIE$ + ANAMCO$ + "9"
               On Error GoTo 0
               GoTo 29
             End If
             '
             On Error Resume Next
             For KI% = 8 To 0 Step -1
               Name DIREVIE$ + ANAMCO$ + TRIM$(Str$(KI%)) As DIREVIE$ + ANAMCO$ + TRIM$(Str$(KI% + 1))
             Next KI%
             On Error GoTo 0
             If EXISTE%(DIREVIE$ + ANAMCO$ + "0") > 0 Then
               'MsgBox "Imposible Instalar (2) " + ANAME$
               MsgBox "Imposible Instalar. No Pudo Renombrarse " + DIREVIE$ + ANAMCO$ + "0 como " + DIREVIE$ + ANAMCO$ + "1"
               On Error GoTo 0
               GoTo 29
             End If
             '
             If InStr(ANAME$, "MSTDRV04") < 1 Then  ' PARA QUE NO RESPALDE RECLAMO
               If EXISTE%(RUTACOM$ + ANAME$) < 1 Then GoTo 24
               On Error Resume Next
               Name RUTACOM$ + ANAME$ As RUTACOM$ + ANAMCO$ + "0"
               If Err Then
                 ' MsgBox "Imposible Instalar (4) " + ANAME$
                 MsgBox "Imposible Instalar. No Pudo Renombrarse " + RUTACOM$ + ANAME$ + " como " + RUTACOM$ + ANAMCO$ + "0"
                 On Error GoTo 0
                 GoTo 29
               End If
               On Error Resume Next
               Name RUTACOM$ + ANAMCO$ + "0" As DIREVIE$ + ANAMCO$ + "0"
               If Err Then
                 ' MsgBox "Imposible Instalar (5) " + ANAME$
                 MsgBox "Imposible Instalar. No Pudo Renombrarse " + RUTACOM$ + ANAMCO$ + "0 como " + DIREVIE$ + ANAMCO$ + "0"
                 On Error GoTo 0
                 GoTo 29
               End If
             End If
             '
24           On Error Resume Next
             FileCopy DIRENEW$ + "\" + ANAME$, RUTACOM$ + ANAME$
             If Err Then
               MsgBox "Imposible Instalar. No Pudo Copiarse " + DIRENEW$ + "\" + ANAME$ + " como " + RUTACOM$ + ANAME$
               On Error GoTo 0
               GoTo 29
             End If
             INSTALAC% = 1
             '
           End If
         End If
       End If
29   Next KU&
     '
   End If
   '
   ' esto es para explorar tambien el \FLEXOFT\NOVEDAD de la carpeta de trabajo
   On Error GoTo 0
   If DIRNU$ = "F:\FLEXOFT\NOVEDAD" Then
     DIRNU$ = PARENTFOLDER$(DIREPRO$) + "\NOVEDAD"
     If DIRNU$ <> "F:\FLEXOFT\NOVEDAD" Then
       GoTo 10
     End If
   End If
   '
   If INSTALAC% > 0 Then MsgBox "Actualizaciones Instaladas"
   Screen.MousePointer = 1
   '
End Sub

Sub RECONN_EMPRE(OKXX%)
   '
   OKXX% = 0
   If EXISTE%("FLEXMUL.DRV") = 1 Then
      SELUDAT2.List2.Clear
      SELUDAT2.List2.Visible = True
      '
      NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
      Dim FMX As String * 79
      '
      For I% = 1 To LOF(NX%) / 79
         Get #NX%, I%, FMX$
         SUCAR% = 0
         For J% = 1 To 77
            SUCAR% = SUCAR% + Asc(Mid$(FMX$, J%, 1))
         Next J%
         If CVI(Mid$(FMX$, 78, 2)) = SUCAR% Then
            RASO$ = DESENCRI$(Left$(FMX$, 30))
            PACCE$ = DESENCRI$(Mid$(FMX$, 32, 46))
            If EXISTE%(TRIM$(PACCE$) + "\.") > 0 Then
               Z$ = Space$(96)
               Call REPLA(Z$, RASO$, 1, 30)
               Call REPLA(Z$, PACCE$, 51, 46)
               SELUDAT2.List2.AddItem Z$
            End If
         End If
      Next I%
      Close #NX%
      '
      If SELUDAT2.List2.ListCount > 0 Then
         OKXX% = 1
         SELUDAT2.List2.ListIndex = 0
         Screen.MousePointer = 1
         If SELUDAT2.List2.ListCount = 1 Then
            Call SELUDAT2.List2_DblClick
            GoTo 25
         End If
         SELUDAT2.Show 1
         Screen.MousePointer = 11
         '
         If SELUDAT2.List2.ListCount > 1 Then
           If TRIM$(EMPREAC$) <> TRIM$(Left$(SELUDAT2.List2.List(0), 30)) Then
             If DERACCE%(EMPREAC$, "Acceso a " + EMPREAC$) < 2 Then
               Call FINAL("")                 ' NO ACCEDE
             End If
           End If
         End If
         '
      End If
   End If
   '
25
End Sub

Private Sub Command17_Click()
   Command17.Visible = False
   ACONEC$ = "TRANSA07"
   Call CONECRUN(ACONEC$, "Deposito Caotico SAPORITI")
   Command17.Visible = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   If EXISTE%("ORCOMP07.EXE") > 0 Then
     Call CONECRUN("ORCOMP07", "Sistema de Ordenes de Compra")
   Else
     Call COMUNI("Falta módulo ORCOMP07")
   End If
   Command18.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call MENSERR(24, "Ayuda en Linea no Disponible\en Esta Localización.")
   Command2.Enabled = True
End Sub

Private Sub Command23_Click()
    Command23.Enabled = False
    MODU$ = "REMITO07"
    Call CONECRUN(MODU$, "Emision de Remito de Facturacion")
    Command23.Enabled = True
End Sub

Private Sub Command28_Click()
   Command28.Visible = False
   ACONEC$ = "BRECEP07"
   Call CONECRUN(ACONEC$, "Recepcion de Materiales")
   Command28.Visible = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call CONECRUN("FSETUP04/CLAVUSER", "Cambio de Clave Personal de Acceso")
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   If EXISTE%(App.Path + "\FSETUP04.EXE") > 0 Then
      Call CONECRUN("FSETUP04", "Configuración de Funcionamiento")
      Call FINAL("")
   End If
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
' comentado porque aqui no se necesita la utilidad
'     NOMECOR$ = "SQLCTRL"
'     On Error Resume Next
'     BaDaConsul.Close
'     On Error GoTo 0
'     On Error Resume Next
'     Kill App.Path + "\" + NOMECOR$ + "1.MDB"
'     On Error GoTo 0
'60   On Error Resume Next
'     CompactDatabase App.Path + "\" + NOMECOR$ + ".MDB", App.Path + "\" + NOMECOR$ + "1.MDB"
'     If Err Then GoTo 75
'     FileCopy App.Path + "\" + NOMECOR$ + ".MDB", App.Path + "\" + NOMECOR$ + ".OLD"
'     If Err Then GoTo 75
'     FileCopy App.Path + "\" + NOMECOR$ + "1.MDB", App.Path + "\" + NOMECOR$ + ".MDB"
'     If Err Then
'75     TTXX$ = "Error (" & Err.Number & ") al Compactar " + NOMECOR$
'       If Err.Number = 3343 Then
'          On Error GoTo 0
'          TTXX$ = TTXX$ + "\  \Esta Base de Datos está Dañada."
'          TTXX$ = TTXX$ + "\¿ Desea Que FLEXOFT Intente Repararla ?"
'          TTXX$ = TTXX$ + "\\Si, Reparar\No, Cancelar"
'          If COMALTER%(TTXX$) = 1 Then
'            On Error Resume Next
'            DBEngine.RepairDatabase LUACC$ + NOMECOR$ + ".MDB"
'            GoTo 60
'          End If
'          GoTo 80
'       End If
'       MsgBox TTXX$
'     End If
'80   On Error GoTo 0
'     On Error Resume Next
'     Kill App.Path + "\" + NOMECOR$ + "1.MDB"
'     On Error GoTo 0
'     MsgBox "FIN"
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call CONECRUN("GENREP07", "Generador de Reportes SQL")
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   If EXISTE%("CARMAQ07.EXE") > 0 Then
     Call CONECRUN("CARMAQ07", "Cálculo de Capacidad y Secuencia de Máquina")
   Else
     Call COMUNI("Falta módulo CARMAQ07")
   End If
   Command7.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   If EXISTE%("CALCAP07.EXE") > 0 Then
     Call CONECRUN("CALCAP07", "Cálculo de Capacidad")
   Else
     Call COMUNI("Falta módulo CALCAP07")
   End If
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
   '
   Screen.MousePointer = 11
   Visible = False
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   Screen.MousePointer = 11
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   Picture4.BackColor = Picture7.BackColor
   '
   If DINFE& < 1 Or DINFE& > Height Then DINFE& = 1500
   Frame3.Top = Height - DINFE&
   '
   On Error Resume Next
   Picture3.Top = Frame3.Top + Frame3.Height
   Picture3.Height = Height - Picture3.Top
   On Error GoTo 0
   Screen.MousePointer = 11
   Call ACOMODATODO
   '
   Call ABRECONEXION
   Call CREACAMPO("COMPRAS", "CARPEIMP", "CodCarpImp", 10, 12, 1)
   Call CREACAMPO("COMPRAS", "CARPEIMP", "RefeCarpImp", 10, 256)
   Call CREACAMPO("COMPRAS", "CARPEIMP", "Descri_Carpi", 12, 0)
   Call CREACAMPO("COMPRAS", "CARPEIMP", "FechAlta", 8, 0)
   Call CREACAMPO("COMPRAS", "CARPEIMP", "Status", 3, 2, 1)
   Call CREACAMPO("COMPRAS", "CARPEIMP", "ValorEstimado", 7, 0)
   Call CREACAMPO("COMPRAS", "CARPEIMP", "CodCenCos", 10, 12, 1)
   Call CREACAMPO("COMPRAS", "CARPEIMP", "Npro_Ocom", 3, 2, 1)
   Call CREACAMPO("COMPRAS", "CARPEIMP", "Raso_Pro", 10, 64)
   Call CREACAMPO("COMPRAS", "CARPEIMP", "His_Carpi", 12, 0)
   Call CREACAMPO("COMPRAS", "CARPEIMP", "TipoImp", 10, 12)
   Call CREACAMPO("COMPRAS", "CARPEIMP", "Moneda", 3, 0)
   Call CREACAMPO("COMPRAS", "CARPEIMP", "Tipo", 10, 3)
   Call CREACAMPO("COMPRAS", "CARPEIMP", "Modo", 3, 0)
   Call CREACAMPO("COMPRAS", "CARPEIMP", "NClien_Oven", 3, 2, 1)
   '
'   Call CREACAMPO("COMPRAS", "CARPEEXP", "CodCarpImp", 10, 12, 1)
'   Call CREACAMPO("COMPRAS", "CARPEEXP", "RefeCarpImp", 10, 256)
'   Call CREACAMPO("COMPRAS", "CARPEEXP", "Descri_Carpi", 12, 0)
'   Call CREACAMPO("COMPRAS", "CARPEEXP", "FechAlta", 8, 0)
'   Call CREACAMPO("COMPRAS", "CARPEEXP", "Status", 3, 2, 1)
'   Call CREACAMPO("COMPRAS", "CARPEEXP", "ValorEstimado", 7, 0)
'   Call CREACAMPO("COMPRAS", "CARPEEXP", "CodCenCos", 10, 12, 1)
'   Call CREACAMPO("COMPRAS", "CARPEEXP", "Npro_Ocom", 3, 2, 1)
'   Call CREACAMPO("COMPRAS", "CARPEEXP", "Raso_Pro", 10, 64)
'   Call CREACAMPO("COMPRAS", "CARPEEXP", "His_Carpi", 12, 0)
'   Call CREACAMPO("COMPRAS", "CARPEEXP", "TipoImp", 10, 12)
'   Call CREACAMPO("COMPRAS", "CARPEEXP", "Moneda", 3, 0)
'   Call CREACAMPO("COMPRAS", "CARPEEXP", "Tipo", 10, 3)
'  Call CREACAMPO("COMPRAS", "CARPEEXP", "Modo", 3, 0)
   '
   Call CREACAMPO("COMPRAS", "LIQUIDACOS", "CodCenCos", 10, 12, 1)
   Call CREACAMPO("COMPRAS", "LIQUIDACOS", "Fecha", 8, 0)
   Call CREACAMPO("COMPRAS", "LIQUIDACOS", "Concepto", 10, 256)
   Call CREACAMPO("COMPRAS", "LIQUIDACOS", "MoneOrig", 3, 0)
   Call CREACAMPO("COMPRAS", "LIQUIDACOS", "ImpoOrig", 7, 0)
   Call CREACAMPO("COMPRAS", "LIQUIDACOS", "MoneRedu", 3, 0)
   Call CREACAMPO("COMPRAS", "LIQUIDACOS", "TipoCam", 7, 0)
   Call CREACAMPO("COMPRAS", "LIQUIDACOS", "ImpoRedu", 7, 0)
   Call CREACAMPO("COMPRAS", "LIQUIDACOS", "TipoGasto", 10, 64)
   '
'   Call CREACAMPO("COMPRAS", "CARPEIMP", "RepComer", 3, 0)
'   Call CREACAMPO("COMPRAS", "CARPEIMP", "Aprobador", 3, 0)
'   Call CREACAMPO("COMPRAS", "CARPEIMP", "TipoOportu", 3, 0)
'   Call CREACAMPO("COMPRAS", "CARPEIMP", "FunilVta", 3, 0)
'   Call CREACAMPO("COMPRAS", "CARPEIMP", "PersCtacto", 10, 64)
'   Call CREACAMPO("COMPRAS", "CARPEIMP", "PersDecisor", 10, 64)
   Call PoneEnCeroNull("CARPEIMP", "STATUS", RAFE&, "NOMESS")
   
   '
   ' PONE LOS ACCESOS DIRECTOS
   Screen.MousePointer = 11
   Call CARGACCDIR
   '
   Treeview1.Style = tvwTreelinesPlusMinusPictureText
   On Error Resume Next
   imlSmallIcons.ListImages.Add , "nueva", LoadPicture("C:\WINDOWS\ICONS\BOOKS05.ICO")
   On Error GoTo 0
   '
   Screen.MousePointer = 11
   Call CARGAMENU
   '
   Call COPYSTRU("INDOPORT", "INCARPIM")
   '
   Visible = True
   Screen.MousePointer = 1
   '
End Sub

Sub AGREGANODO(CODPA$, CODIJ$, DESIJ$)
    '
    On Error Resume Next
    If CODIJ$ = "ROOT" Then
        Set nodX = Treeview1.Nodes.Add(, , "ROOT", DESIJ$, "closed", "open")
      Else
        Set nodX = Treeview1.Nodes.Add(CODPA$, tvwChild, CODIJ$, DESIJ$, "closed", "open")
    End If
'    nodX.EnsureVisible       ' Muestra todos los nodos.
    If CODPA$ = "ROOT" Or CODIJ$ = "ROOT" Then
        nodX.EnsureVisible       ' Muestra el nodo.
    End If
    On Error GoTo 0
End Sub
'
Sub CARGAMENU()
    '
    Treeview1.Nodes.Clear
    '
    Call AGREGANODO("", "ROOT", "Sistema FLEXOFT de Carpetas de Importacion")
        '
        Call AGREGANODO("ROOT", "ARCHIVOS", "Archivos Maestros")
            Call AGREGANODO("ARCHIVOS", "AMASTO07", "Maestro de Items de Stock")
            Call AGREGANODO("ARCHIVOS", "AMAPRO", "Maestro de Proveedores")
            Call AGREGANODO("ARCHIVOS", "AMACLI", "Padron General de Clientes")
            Call AGREGANODO("ARCHIVOS", "TACONFIG", "Tablas Complementarias de Configuracion")
               Call AGREGANODO("TACONFIG", "PLANCUEN", "Plan General de Cuentas Contables")
               Call AGREGANODO("TACONFIG", "TACENCOS", "Tabla de Centros de Costos")
               '
        Call AGREGANODO("ROOT", "CARPEXPO", "Carpetas de Exportacion")
            Call AGREGANODO("CARPEXPO", "CARPENUE", "Nueva Carpeta de Exportacion")
            Call AGREGANODO("CARPEXPO", "CARPEXIS", "Abrir Carpeta Existente")
            Call AGREGANODO("CARPEXPO", "BORRCARP", "Borrar Carpeta de Exportacion")
            '
        Call AGREGANODO("ROOT", "PROTRAN", "Procesar Transacciones")
            Call AGREGANODO("PROTRAN", "ORDECOM", "Ordenes de Venta")
               Call AGREGANODO("ORDECOM", "NUEVAOVEN", "Nueva O/Venta")
               Call AGREGANODO("ORDECOM", "MODIOVEN", "Modificar O/Venta")
               Call AGREGANODO("ORDECOM", "CONLISV", "Consultas y Listados")
                 Call AGREGANODO("CONLISV", "CONLISVEN", "Generales")
                 Call AGREGANODO("CONLISV", "CONPORCAR", "Consulta por Carpeta de Importacion")
            '
            Call AGREGANODO("PROTRAN", "DESPAMAT", "Despacho de Materiales")
               Call AGREGANODO("DESPAMAT", "GEPACKLIST", "Generar Packing List")
'               Call AGREGANODO("DESPAMAT", "INSPYAPRO", "R.U.F.")
'               Call AGREGANODO("DESPAMAT", "INSPYAPROD", "R.U.F.")
'               '
'            Call AGREGANODO("PROTRAN", "FACTPROV", "Facturacion a Proveedores")
'            Call AGREGANODO("PROTRAN", "GASTOSNA", "Gastos de Exportacion")
'            Call AGREGANODO("PROTRAN", "COSCONT", "Costos Contables")
            '
        Call AGREGANODO("ROOT", "LISTADOS", "Reportes, Consultas y Listados")
            Call AGREGANODO("LISTADOS", "ANOTADOR", "Historial de Transacciones")
'            Call AGREGANODO("LISTADOS", "DETALLE", "Detalle de Costos de Exportacion")
'            Call AGREGANODO("LISTADOS", "DISTRIBUCION", "Distribucion de Costos")
'            Call AGREGANODO("LISTADOS", "LIQCCOS", "Liquidacion por Carpeta de Exportacion")
'               Call AGREGANODO("LIQCCOS", "RECALIQ", "Re-calculo de Liquidacion")
'               Call AGREGANODO("LIQCCOS", "CONLIQ", "Consulta de Liquidacion")
            '
        Call AGREGANODO("ROOT", "CONFIGUR", "Herramientas de Configuracion")
            Call AGREGANODO("CONFIGUR", "FSETUP04", "Configuracion de Funcionamiento")
            Call AGREGANODO("CONFIGUR", "CLAVEACC", "Cambio Clave Personal de Acceso")
            Call AGREGANODO("CONFIGUR", "COMPACBD", "Mantenimiento Base de Control")
            '
    Treeview1.Nodes.Item(1).Selected = True
    Call ACOMODATODO
    '
End Sub

Sub CARGACCDIR()
   CANADIR = XVALO(VALCONTROL(USNBR%, "ADIRECT", "CANADIR"))
   If CANADIR > 255 Then CANADIR = 255
   '
   For I& = 1 To CANADIR
      If I& > Image3.Count Then
         Load Image3(I& - 1)
         Image3(I& - 1).Top = Image3(I& - 2).Top
         Image3(I& - 1).Left = Image3(I& - 2).Left + Image3(I& - 2).Width + 100
         Image3(I& - 1).Visible = True
      End If
      '
      ADIR$ = "ACCDIR-" + TRIM$(Str$(I& - 1))
      ARCHICON$ = VALCONTROL(USNBR%, ADIR$, "ARCHICON")
      TOOTIPTE$ = VALCONTROL(USNBR%, ADIR$, "TOOTIPTE")
      On Error Resume Next
      Image3(I& - 1).Picture = LoadPicture(ARCHICON$)
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Imagen no Válida")
      End If
      On Error GoTo 0
      Image3(I& - 1).Visible = True
      Image3(I& - 1).ToolTipText = TOOTIPTE$
      Image3(I& - 1).Refresh
   Next I&
End Sub

Sub PONELOGO()
   '
   RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
   If Right$(RUTARGRA$, 1) <> "\" Then
      RUTARGRA$ = RUTARGRA$ + "\"
   End If
   '
   ISOLOGO$ = TRIM$(CONTROL$("", "ISOLOGO"))
   If ISOLOGO$ <> "" Then
      ARCHIMAG$ = RUTARGRA$ + ISOLOGO$
      If EXISTE%(ARCHIMAG$) = 1 Then
        Image11.Picture = LoadPicture("")
        On Error Resume Next
        Image11.Picture = LoadPicture(ARCHIMAG$)
        If Err Then
          On Error GoTo 0
          ARXX$ = REPLACAR$(ARCHIMAG$, "\", "/")
          Call MENSERR(24, "Archivo Imagen Isologo\" + ARXX$ + "\No Válido o Inexistente.")
        End If
        On Error GoTo 0
        Picture19.Visible = True
        IMAGEFILE$ = ARCHIMAG$     ' LUDAT$ + "LOGOEMPR.BMP"
        '
        FORMALTER.Image3.Picture = LoadPicture("")
        On Error Resume Next
        FORMALTER.Image3.Picture = LoadPicture(IMAGEFILE$)
        On Error GoTo 0
        '
12      HIMAGEN = FORMALTER.Image3.Picture.Width
        VIMAGEN = FORMALTER.Image3.Picture.Height
        '
        'Picture19.Width = Frame1.Left
        'Picture19.Height = 1000
        'Picture19.Left = 0
        HCUADRO = Frame1.Left   'Picture19.Width
        VCUADRO = 1000 ' Picture19.Height
        '
        REDUH = HIMAGEN / HCUADRO
        REDUV = VIMAGEN / VCUADRO
        REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
        If REDUC < 1 Then REDUC = 1
        '
        With Picture19
          .Height = VCUADRO * REDUV / REDUC
          .Width = HCUADRO * REDUH / REDUC
          '.Left = .Left + (HCUADRO - Picture19.Width) / 2
          .Left = (HCUADRO - Picture19.Width) / 2
            If .Left < 0 Then .Left = 0
          .Top = .Top - (VCUADRO - Picture19.Height) / 2
            If .Top < 0 Then .Top = 0
        End With
        '
        Image11.Width = Picture19.Width
        Image11.Height = Picture19.Height
        Image11.Stretch = True
        On Error Resume Next
        Image11.Picture = LoadPicture(IMAGEFILE) ' FORMALTER.IMAGE11.Picture ' , PH1 + BORIZQ, PV1 + BORSUP, ANCHCUA, ALTUCUA
        On Error GoTo 0
        '
        FORMALTER.Image3.Picture = LoadPicture("")
        '
      End If
   End If
   '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   '
   Cancel = False
   If UnloadMode = 0 Then     ' desde el botón x
      Cancel = True           ' cancela la descarga
      MsgBox "Debe Utilizar el Botón de Salida para Cerrar Correctamente la Sesión de Trabajo", vbExclamation
   End If
   '
End Sub
'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
   Picture10.Height = ScaleHeight
   Picture4.Top = ScaleHeight - Picture4.Height
   Command6.Top = Picture4.Top - Command6.Height
   Command6.Visible = True
   Picture4.Visible = True
   If Picture4.Top < Command6.Top + Command6.Height Then Command6.Top = Picture4.Top - Command6.Height
   If Command6.Top < Command4.Top + Command4.Height + 200 Then Command6.Visible = False
   If Picture4.Top < Command4.Top + Command4.Height Then Picture4.Visible = False
   Frame1.Height = ScaleHeight + 250
   Picture7.Height = ScaleHeight
   APIC7 = 2 * Command1.Left + Command1.Width
   Picture7.Left = ScaleWidth - APIC7 - 20
      On Error Resume Next
      DINFE& = XVALO(VALCONTROL$(USNBR%, "POSIFORPRIN", "FRAME3_TOP"))
      On Error GoTo 0
      If DINFE& < 1 Or DINFE& > Height Then DINFE& = 1500
   Frame3.Top = Height - DINFE&
   Picture3.Top = Frame3.Top + Frame3.Height
   Picture3.Height = Height - Picture3.Top
   Call CARGAIMAGEN
   Call ACOMODATODO
   Picture7.Refresh
   Picture8.Refresh
   Picture10.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call FINAL("")
End Sub


Private Sub frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 1
End Sub

Private Sub frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame1.MousePointer = 9
   If ARRABOR% = 1 Then
      If Frame1.Left + X > 20 Then
          Frame1.Left = Frame1.Left + X
          Call ACOMODATODO
      End If
   End If
End Sub

Private Sub frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 0
End Sub

Sub ACOMODATODO()
    '
    On Error Resume Next
    Picture10.Width = Frame1.Left + Frame1.Width + 30
    Label6.Width = Frame1.Left + 50
    Label7.Width = Frame1.Left + 50
    On Error GoTo 0
    Call CARGAIMAGEN
    Call PONELOGO
    On Error Resume Next
    Label6.Top = Picture19.Height
    Treeview1.Top = Label6.Top + Label6.Height - 10
    Treeview1.Height = Me.ScaleHeight - Treeview1.Top
    Treeview1.Width = 12000
    Treeview1.Width = Frame1.Left
    '
    Picture7.Refresh
    Picture8.Refresh
    Picture10.Refresh
    On Error GoTo 0
    '
End Sub

Private Sub frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 1
End Sub

Private Sub frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame2.MousePointer = 7
   If ARRABOR% = 1 Then
      Frame2.Top = Frame2.Top + Y
      Call ACOMODATODO
   End If
End Sub

Private Sub frame2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 0
End Sub

Private Sub frame3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 1
End Sub

Private Sub frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame3.MousePointer = 7
   If ARRABOR% = 1 Then
      Frame3.Top = Frame3.Top + Y
      On Error Resume Next
      Picture2.Height = Frame3.Top
      Picture3.Top = Frame3.Top + Frame3.Height
      Picture3.Height = Height - Picture3.Top
      On Error GoTo 0
      Picture3.Refresh
      Call CARGAIMAGEN
   End If
End Sub

Private Sub frame3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 0
   DINFE& = Height - Frame3.Top
   Call GRAVALCON(USNBR%, "POSIFORPRIN", "FRAME3_TOP", FORMATNUM$(DINFE&, "F10.0"))
End Sub

Private Sub Image3_DblClick(Index As Integer)
   AREXE$ = VALCONTROL$(USNBR%, "ACCDIR-" + TRIM$(Str$(Index)), "DESTAPLI")
   LINCOMMA$ = VALCONTROL$(USNBR%, "ACCDIR-" + TRIM$(Str$(Index)), "LINCOMMA")
   On Error Resume Next
   NX1 = Shell(AREXE$ + " " + LINCOMMA$, vbNormalFocus)
   If Err Then
      MsgBox "Imposible Ejecutar '" + AREXE$ + " " + LINCOMMA$ + "'"
   End If
   On Error GoTo 0
End Sub

Private Sub Image3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      FORCONSHK.Label7 = ""
      FORCONSHK.Label6 = TRIM$(Str$(Index))
      ADIR$ = "ACCDIR-" + TRIM$(Str$(Index))
      FORCONSHK.Text1 = VALCONTROL(USNBR%, ADIR$, "DESTAPLI")
      FORCONSHK.Text2 = VALCONTROL(USNBR%, ADIR$, "LINCOMMA")
      FORCONSHK.Text3 = VALCONTROL(USNBR%, ADIR$, "ARCHICON")
      FORCONSHK.Text4 = VALCONTROL(USNBR%, ADIR$, "TOOTIPTE")
      '
      FORCONSHK.Show 1
      If FORCONSHK.Label7 = "OK" Then
         DESTAPLI$ = FORCONSHK.Text1
         LINCOMMA$ = FORCONSHK.Text2
         ARCHICON$ = FORCONSHK.Text3
         TOOTIPTE$ = FORCONSHK.Text4
         '
         If FORCONSHK.Label6 = "" Then
             INDIA& = Image3.Count
           Else
             INDIA& = XVALO(FORCONSHK.Label6)
         End If
         Call GRAVALCON(USNBR%, "ADIRECT", "CANADIR", TRIM$(Str$(1 + INDIA&)))
         '
         ADIR$ = "ACCDIR-" + TRIM$(Str$(INDIA&))
         Call GRAVALCON(USNBR%, ADIR$, "DESTAPLI", DESTAPLI$)
         Call GRAVALCON(USNBR%, ADIR$, "LINCOMMA", LINCOMMA$)
         Call GRAVALCON(USNBR%, ADIR$, "ARCHICON", ARCHICON$)
         Call GRAVALCON(USNBR%, ADIR$, "TOOTIPTE", TOOTIPTE$)
      End If
      '
      Call CARGACCDIR
   End If
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    If Button = 2 Then
10      CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Archivo Origen de Imagen"
            ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
        CommonDialog1.Filter = "Todos los archivos (*.*)|*.*|Mapas de Bit" & _
        "(*.BMP)|*.BMP|Archivos Graficos JPEG (*.JPG)|*.JPG"
        ' Especificar el filtro predeterminado
        CommonDialog1.FilterIndex = 2
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
        '
        DoEvents
        ARIMAG$ = UCase$(CommonDialog1.filename)
        '
        If ARIMAG$ <> "" Then
           On Error Resume Next
           Image2.Picture = LoadPicture(ARIMAG$)
           If Not (Err) Then
              On Error GoTo 0
              Call GRAVALCON(USNBR%, "IMAGEN", "ARCHIMAGEN", ARIMAG$)
              Call CARGAIMAGEN
              Exit Sub
           End If
           On Error GoTo 0
           Call MENSERR(24, "Imposible Aplicar la Imagen Elegida.")
           GoTo 10
        End If
    End If
99  On Error GoTo 0
    '
End Sub

Private Sub Picture3_DblClick()
'URLopen "http://www.smn.gov.ar/"
'Inet1.Password = "I(3Lei#4"
'Inet1.USERNAME = "Juan Soto"
'Text1 = Inet1.OpenURL
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      FORCONSHK.Label7 = ""
      FORCONSHK.Show 1
      If FORCONSHK.Label7 = "OK" Then
         DESTAPLI$ = FORCONSHK.Text1
         LINCOMMA$ = FORCONSHK.Text2
         ARCHICON$ = FORCONSHK.Text3
         TOOTIPTE$ = FORCONSHK.Text4
         '
         If FORCONSHK.Label6 = "" Then
             INDIA& = Image3.Count
             Call GRAVALCON(USNBR%, "ADIRECT", "CANADIR", TRIM$(Str$(1 + INDIA&)))
           Else
             INDIA& = XVALO(FORCONSHK.Label6)
         End If
         '
         ADIR$ = "ACCDIR-" + TRIM$(Str$(INDIA&))
         Call GRAVALCON(USNBR%, ADIR$, "DESTAPLI", DESTAPLI$)
         Call GRAVALCON(USNBR%, ADIR$, "LINCOMMA", LINCOMMA$)
         Call GRAVALCON(USNBR%, ADIR$, "ARCHICON", ARCHICON$)
         Call GRAVALCON(USNBR%, ADIR$, "TOOTIPTE", TOOTIPTE$)
      End If
      '
      Call CARGACCDIR
   End If
End Sub

Private Sub Timer1_Timer()
   Static LACTIVA%
   For KKI% = 0 To 3
     Label1(KKI%).FontSize = 12
     Label1(KKI%).FontBold = False
     Label1(KKI%).ForeColor = RGB(0, 0, 0)
     If KKI% = LACTIVA% Then
       Label1(KKI%).FontSize = 12
       Label1(KKI%).FontBold = True
       Label1(KKI%).ForeColor = RGB(255, 0, 0)
     End If
   Next KKI%
   DoEvents
   LACTIVA% = (LACTIVA% + 1) Mod 4
   '
End Sub

Private Sub Timer2_Timer()
Exit Sub
   If EXISTE%("C:\FLEXOFT\PRUEBA.PDF") > 0 Then
     ACROL$ = FILEFIND$("C:", "ACRORD32.EXE", 0)
     If ACROL$ <> "" Then
       RUTOR$ = ACROL$ + " "
       TEXACT$ = "" '  "Adobe"
       RUTAR$ = "C:\FLEXOFT\"
       ANAME$ = RUTOR$ + RUTAR$ + "PRUEBA.PDF"
       RETURNVALUE = Shell(ANAME$, 1)
18     On Error Resume Next
       AppActivate TEXACT$ + " Reader - [PRUEBA.pdf]" ' Adobe Reader"
       'If Err Then
         'Resume 20
20       'On Error GoTo 0
         'TINI = Timer
         While Timer < TINI + 1: DoEvents: Wend
         'GoTo 18
       'End If
       On Error GoTo 0
       '
     End If
   End If
End Sub

Function PROCACTI%()
   '
   File1.Path = LUCOM$
   File1.Pattern = "*.EXK"
   File1.Refresh
   '
   For U& = 1 To File1.ListCount
      X$ = TRIM$(UCase$(File1.List(U& - 1)))
      If X$ = "ADFLESQL.EXK" Then GoTo 9         ' EXCEPCION
      If Left$(X$, 4) = "FLEX" Then GoTo 9      ' EXCEPCION
      If X$ <> "" Then
        If X$ <> App.EXEName + ".EXK" Then
          If FileLen(LUCOM$ + X$) > 0 Then
            On Error GoTo 10
            Kill LUCOM$ + X$
            On Error GoTo 0
          End If
        End If
      End If
9  Next U&
   PROCACTI% = 0
   Exit Function
   '
10 Resume 11
11 On Error GoTo 0
   PROCACTI% = 1
   X1$ = Left$(X$, Len(X$) - 1) + "E"
   TTXX$ = "Imposible Cerrar Normalmente la Sesión de Trabajo.\   \"
   TTXX$ = TTXX$ + "Se Encuentra Abierto y Ejecutando el Proceso\"
   TTXX$ = TTXX$ + "'" + X1$ + "'\  \"
   TTXX$ = TTXX$ + "Cierre este Proceso Antes de Cerrar Sesión."
   Call COMUNI(TTXX$)
   '
End Function

Private Sub Timer3_Timer()
    HOII = HOY(HOS$)
    Label4 = HOS$
    DIASE$ = DSEM$(HOII)
    Label5 = UCase$(Left$(DIASE$, 1)) + LCase$(Mid$(DIASE$, 2))
    Label3 = Left$(Time, 8)
End Sub

Private Sub Treeview1_NodeClick(ByVal Node As ComctlLib.Node)
   '
   CLANODO$ = TRIM$(Left$(Node.Key, 64))
   PPXX% = InStr(CLANODO$, "/"): If PPXX% < 1 Then PPXX% = 1 + Len(CLANODO$)
   EXECONN$ = Left$(CLANODO$, PPXX% - 1)
   '
   If EXECONN$ = "AMAPRO" Then
         Call CONECRUN("ARCHIG07/AMAPRO", "Padron Maestro de Proveedores")
       ElseIf EXECONN$ = "AMACLI" Then
         Call CONECRUN("ARCHIG07/AMACLI", "Padron Maestro de Clientes")
       ElseIf EXECONN$ = "PLANCUEN" Then
         Call CARCUEINGAS
       ElseIf EXECONN$ = "TACENCOS" Then
         Call CARTACENCOS
       ElseIf EXECONN$ = "FSETUP04" Then
         Call Command4_Click
       ElseIf EXECONN$ = "CLAVEACC" Then
         Call Command3_Click
       ElseIf EXECONN$ = "COMPACBD" Then
         Call Command5_Click
       ElseIf EXECONN$ = "GENREP07" Then
         Call Command6_Click
       ElseIf EXECONN$ = "CARPENUE" Then
         Call NUEVACARPIMP
       ElseIf EXECONN$ = "CARPEXIS" Then
         Call ABRECARPIMP
       ElseIf EXECONN$ = "BORRCARP" Then
         Call ANULCARP
       ElseIf EXECONN$ = "ANOTADOR" Then
         Call ANOTACARP
      ElseIf EXECONN$ = "NUEVAOVEN" Then
         Call NUEVAOVEN
      ElseIf EXECONN$ = "MODIOVEN" Then
         Call MODIFOVEN
      ElseIf EXECONN$ = "CONLISVEN" Then
         Call CONSULOVEN
      ElseIf EXECONN$ = "GEPACKLIST" Then
         Call GEPACKLIST
      ElseIf EXECONN$ = "CONPORCAR" Then
         Call CONSULVENPORCARP
'      ElseIf EXECONN$ = "INSPYAPRO" Then
'         Call INSPYAPROCOMP
'      ElseIf EXECONN$ = "INSPYAPROD" Then
'         Call INSPYAPROCOMPD
'      ElseIf EXECONN$ = "CODPRO" Then
'         Call CONSULCODCOM
'      ElseIf EXECONN$ = "TCODPRO" Then
'         Call CONSULTCODCOM
'      ElseIf EXECONN$ = "PPSTOCK" Then
'         Call CONSULPPSTOCK
'      ElseIf EXECONN$ = "LPROVENC" Then
'         Call CONSULLPROVENC
'      ElseIf EXECONN$ = "FKARDEX" Then
'         Call CONSULFKARDEX
'      ElseIf EXECONN$ = "GESINCO" Then
'         Call CONSULGESINCO
'      ElseIf EXECONN$ = "COLIMO" Then
'         Call CONSULCOLIMO
      ElseIf EXECONN$ = "FACTPROV" Then
         'Call FACTPROVCOMP
      ElseIf EXECONN$ = "GASTOSNA" Then
         'Call GASTOSNACOMP
      ElseIf EXECONN$ = "COSCONT" Then
         'Call COSCONTCOMP
       ElseIf EXECONN$ = "RECALIQ" Then
         'Call LIQUIDAPORCCOSTO
       ElseIf EXECONN$ = "CONLIQ" Then
         'Call CONSULLIQUIDA
       ElseIf EXISTE%(App.Path + "\" + EXECONN$ + ".EXE") > 0 Then
         Call CONECRUN(EXECONN$, "")
   End If
   '
End Sub

'
Sub ANOTACARP()
   ANOTACARPI.Show 1
End Sub

Sub CARGANOVED()
   '
21 Call SELECHO("INDIPROS/Indice Prospect")
   VDEVU$ = VALACT1$("INDIPROS")
   If VDEVU$ <> "" Then
      ANOTASEG09.Label26 = TRIM$(VDEVU$)
      ANOTASEG09.Show 1
   End If
   '
End Sub

Sub NUEVACARPIMP()
    '
    NUOPO$ = PROXICARP$
    If NUOPO$ <> "" Then
       CONDI$ = "CODCARPIMP = '" & NUOPO$ & "'"
       If CantRegistros("CARPEIMP", CONDI$) < 1 Then
          Call CreaRegistro("CARPEIMP", "CODCARPIMP", NUOPO$, "NOMESS")
          HOII% = HOY(HOS$): FEALTA = FETEXCOR1$(HOII%)
          Call ACTUREGISTRO("CARPEIMP", "FECHALTA", CONDI$, FEALTA, RAFE&, "NOMESS")
          Call ACTUREGISTRO("CARPEIMP", "STATUS", CONDI$, 0, RAFE&, "NOMESS")
          Call ACTUREGISTRO("CARPEIMP", "MODO", CONDI$, 1, RAFE&, "NOMESS")
10        Call STDFORM("CARPEXP", NUOPO$)
          
          If STDSCREEN.APRO = "OK" Then
            '
            NC% = XVALO(VALOFORM("TEXT-4"))
            If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
               Call MENSERR(24, "Cliente No Valido o Inexistente en la Base de Datos")
               GoTo 10
            End If
            '
            MONE% = XVALO(VALOFORM("TEXT-3"))
            If ECOARCH$("TAMONED", MKI$(MONE%)) = "" Then
               Call MENSERR(24, "Moneda Seleccionada Invalida")
               GoTo 10
            End If
            '
            CCOS$ = VALOFORM("TEXT-5")
            If ECOARCH$("TACENCOS", CCOS$) = "" Then
               Call MENSERR(24, "Centro de Costos Invalido")
               GoTo 10
            End If
            '
            Call ACTUREGISTRO("CARPEIMP", "TIPO", CONDI$, 1, RAFE&, "NOMESS")
            '
            Call REPLA(TTXX$, HOS$, 1, 8)
            Call REPLA(TTXX$, USERNAME$, 69, 24)
            Call REPLA(TTXX$, "Alta Carpeta de Exportacion", 11, 56)
            Call AGREGA_ANOTACARPI(NUOPO$, TTXX$)
            TTXX$ = Space(76)
            Call REPLA(TTXX$, "Ref: " + VALOFORM("PARAM-4"), 11, 56)
            Call AGREGA_ANOTACARPI(NUOPO$, TTXX$)
            Call REPLA(TTXX$, "Desc: " + VALOFORM("TEXT-1"), 11, 56)
            Call AGREGA_ANOTACARPI(NUOPO$, TTXX$)
            Call REPLA(TTXX$, "Tipo Exportacion: " + VALOFORM("TEXT-7"), 11, 56)
            Call AGREGA_ANOTACARPI(NUOPO$, TTXX$)
            Call REPLA(TTXX$, "Cliente: " + VALOFORM("PARAM-3"), 11, 56)
            Call AGREGA_ANOTACARPI(NUOPO$, TTXX$)
            Call REPLA(TTXX$, "Valor Estimado: " + FORMATNUM$(VALOFORM("TEXT-2"), "F10.2"), 11, 56)
            Call AGREGA_ANOTACARPI(NUOPO$, TTXX$)
            '
            Call REPLA(TTXX$, "Moneda: " + ECOARCH("TAMONED", MKI$(MONE%)), 11, 56)
            Call AGREGA_ANOTACARPI(NUOPO$, TTXX$)
            '
            Call REPLA(TTXX$, "Centro de Costo: " + VALOFORM("TEXT-5"), 11, 56)
            Call AGREGA_ANOTACARPI(NUOPO$, TTXX$)
            Call AGREGA_ANOTACARPI(NUOPO$, String$(88, "="))
          Else
            Call BORRAREGISTROS("CARPEIMP", CONDI$, REG&)
          End If
       End If
    End If
    '
End Sub

Sub ABRECARPIMP()
    '
    CONDI$ = "CodcarpImp <>'' AND STATUS < 9 AND MODO = 1"
    CONDI$ = CONDI$ + "ORDER BY CodCarpImp"
    Call CARGALISEL("!INCARPIM", "CARPEIMP", "CodcarpImp/A12;FechAlta/d8;RefeCarpImp/A48", CONDI$, "NOMESS")
    Call SELECHO("!INCARPIM/INDICE DE CARPETAS DE IMPORTACION/EXPORTACION")
    
    CODIOPO$ = TRIM$(VALACT1$("!INCARPIM"))
    If CODIOPO$ <> "" Then
       Call STDFORM("CARPEXP", CODIOPO$)
    End If
    '
End Sub

Sub ANULCARP()
    '
    OPX% = COMALTER%("Opciones de Trabajo:\\Anular Carpetas\Revertir Anulacion")
    If OPX% = 1 Then
        CONDI$ = "CodcarpImp <> '' AND STATUS < 9 AND MODO = 1"
        CONDI$ = CONDI$ + "ORDER BY CodCarpImp"
        Call CARGALISEL("!INCARPIM", "CARPEIMP", "CodcarpImp/A12;FechAlta/d8;RefeCarpImp/A48", CONDI$, "NOMESS")
        Call FRESHECHO("!INCARPIM")
        Call SELECHO("!INCARPIM/INDICE DE CARPETAS DE IMPORTACION/EXPORTACION")
        
        CODIOPO$ = TRIM$(VALACT1$("!INCARPIM"))
        If CODIOPO$ <> "" Then
           CONDI$ = "CODCARPIMP = '" & CODIOPO$ & "'"
           DESOPO$ = TRIM$(VALOBADA("CARPEIMP", "REFECARPIMP", CONDI$, "NOMESS"))
           TTXX$ = "Está Absolutamente Seguro de Anular\la Carpeta de Exportacion\'" + CODIOPO$ + " - " + DESOPO$ + "' ?"
           If COMALTER%(TTXX$ + "\\No, Conservar\Si, Anularla") = 2 Then
              Call ACTUREGISTRO("CARPEIMP", "STATUS", CONDI$, 9, RAFE&, "NOMESS")
           End If
        End If
     ElseIf OPX% = 2 Then
        CONDI$ = "CodcarpImp <>'' AND STATUS >= 9 AND MODO=1 ORDER BY CodCarpImp"
        Call CARGALISEL("!INCARPIM", "CARPEIMP", "CodcarpImp/A12;FechAlta/d8;RefeCarpImp/A48", CONDI$, "NOMESS")
        Call SELECHO("!INCARPIM/INDICE DE CARPETAS DE IMPORTACION/EXPORTACION")
        
        CODIOPO$ = TRIM$(VALACT1$("!INCARPIM"))
        If CODIOPO$ <> "" Then
           CONDI$ = "CODCARPIMP = '" & CODIOPO$ & "'"
           Call ACTUREGISTRO("CARPEIMP", "STATUS", CONDI$, 0, RAFE&, "NOMESS")
        End If
    End If
    '
End Sub

Sub CARGAIMAGEN()
    Picture2.Top = 0
    Picture2.Left = Picture10.Width
    On Error Resume Next
    Picture2.Width = Picture7.Left - Picture10.Width
    Picture2.Height = Frame3.Top
    On Error GoTo 0
    Picture2.Refresh
    Picture3.Top = Frame3.Top + Frame3.Height
    Picture3.Left = Frame1.Left + Frame1.Width
    '
    HIMAGEN = Image2.Picture.Width
    VIMAGEN = Image2.Picture.Height
    '
    HCUADRO = Picture2.Width
    VCUADRO = Picture2.Height
    '
    REDUH = HIMAGEN / HCUADRO
    REDUV = VIMAGEN / VCUADRO
    REDUC = REDUH: If REDUV < REDUH Then REDUC = REDUV
    '
    Picture2.PaintPicture Image2.Picture, 0, 0, HIMAGEN / REDUC, VIMAGEN / REDUC ' Picture2.Height, Picture2.Height
    Picture2.Visible = True
    Picture2.Refresh

End Sub

Function PROXICARP$()
   '
   If COTERMI$ = "" Then
      ITERMI$ = IDENTER$
10    COTERMI$ = CONTROL$(ITERMI$, "COTERMI")
      If Len(COTERMI$) <> 2 Then
        OPCOTEL.Show 1
        GoTo 10
      End If
   End If
   '
   OTX$ = "IX-______%"
   CAMPOSE$ = "MAX(CAST(SUBSTRING(CodCarpImp,4,6) AS Integer))"
   CONDI$ = "CODCARPIMP LIKE '" & OTX$ & "'"
   NUCONSU& = 1 + XVALO(VALOBADA("CARPEIMP", CAMPOSE$, CONDI$))
   '
20 NEXTNU$ = TRIM$(Str$(NUCONSU&))
   While Len(NEXTNU$) < 6: NEXTNU$ = "0" + NEXTNU$: Wend
   PROXICARP$ = "IX-" + NEXTNU$ + "-" + COTERMI$
   '
End Function
     
Sub NUEVAOVEN()
         
10       CONDI$ = "CodcarpImp <>'' AND STATUS < 9 AND MODO = 1"
         CONDI$ = CONDI$ + "ORDER BY CodCarpImp"
         Call CARGALISEL("!INCARPIM", "CARPEIMP", "CodcarpImp/A12;FechAlta/d8;RefeCarpImp/A48", CONDI$, "NOMESS")
         Call SELECHO("!INCARPIM/INDICE DE CARPETAS DE IMPORTACION\EXPORTACION")
         '
         CODCARPE$ = TRIM$(VALACT1$("!INCARPIM"))
         If CODCARPE$ <> "" Then
            CONDI$ = "CodcarpImp ='" & CODCARPE$ & "' AND STATUS < 9 AND MODO = 1"
            If CantRegistros("CARPEIMP", CONDI$) > 0 Then
               NCLI% = XVALO(VALOBADA("CARPEIMP", "NClien_Oven", CONDI$, "NOMESS"))
               If NCLI% >= 1 Then
                 Call CONECRUN("PEDIDO07/NUEVAOVEN;CODCARPE =" & CODCARPE$, "Ordenes de Ventas de Exportacion")
               Else
                 Call COMUNI("No Se Encontro Cliente Asociado a la Carpeta de Exportacion: '" + TRIM$(CODCARPE$) + "'.\  \Verifique y Vuelva a Intentar.")
                 GoTo 10
               End If
            End If
         End If

End Sub

Sub MODIFOVEN()

         CONDI$ = "CodcarpImp <>'' AND STATUS < 9 AND MODO = 1"
         CONDI$ = CONDI$ + "ORDER BY CodCarpImp"
         Call CARGALISEL("!INCARPIM", "CARPEIMP", "CodcarpImp/A12;FechAlta/d8;RefeCarpImp/A48", CONDI$, "NOMESS")
         Call SELECHO("!INCARPIM/INDICE DE CARPETAS DE IMPORTACION\EXPORTACION")
         '
         CODCARPE$ = TRIM$(VALACT1$("!INCARPIM"))
         If CODCARPE$ <> "" Then
            Call CONECRUN("PEDIDO07/MODPEDIDO;CODCARPE =" & CODCARPE$, "Ordenes de Ventas de Exportacion")
         End If
End Sub

Sub CONSULOVEN()

'         CONDI$ = "CodcarpImp <>'' AND STATUS < 9 ORDER BY CodCarpImp"
'         Call CARGALISEL("INCARPIM", "CARPEIMP", "CodcarpImp/A12;FechAlta/d8;RefeCarpImp/A48", CONDI$, "NOMESS")
'         Call SELECHO("INCARPIM/INDICE DE CARPETAS DE IMPORTACION")
'
'         CODCARPE$ = TRIM$(VALACT1$("INCARPIM"))
'         If CODCARPE$ <> "" Then
            Call CONECRUN("PEDIDO07/CONLISPED", "Ordenes de Ventas de Exportacion")
'         End If

End Sub

Sub GEPACKLIST()
         '
10       CONDI$ = "CodcarpImp <>'' AND STATUS < 9 AND MODO = 1"
         CONDI$ = CONDI$ + "ORDER BY CodCarpImp"
         Call CARGALISEL("!INCARPIM", "CARPEIMP", "CodcarpImp/A12;FechAlta/d8;RefeCarpImp/A48", CONDI$, "NOMESS")
         Call SELECHO("!INCARPIM/INDICE DE CARPETAS DE EXPORTACION")
         '
         CODCARPE$ = TRIM$(VALACT1$("!INCARPIM"))
         If CODCARPE$ <> "" Then
            Call CONECRUN("REMITO07/PACKLIST;CODCARPE =" & CODCARPE$, "Ordenes de Ventas de Exportacion")
         End If
End Sub
 
Sub CONSULCODCOM()

'         Call SELECHO("TACENCOS/Tabla de Centro de Costos")
'
'         CODICENCOS$ = TRIM$(VALACT1$("TACENCOS"))
'         If CODICENCOS$ <> "" Then
'            MENCAOT09.Show 1
'         End If
     Call FINAL("?DEXIPORPROD")
End Sub
     
Sub FACTPROVCOMP()

         CONDI$ = "CodcarpImp <>'' AND STATUS < 9 AND MODO = 1"
         CONDI$ = CONDI$ + "ORDER BY CodCarpImp"
         Call CARGALISEL("!INCARPIM", "CARPEIMP", "CodcarpImp/A12;FechAlta/d8;RefeCarpImp/A48", CONDI$, "NOMESS")
         Call SELECHO("!INCARPIM/INDICE DE CARPETAS DE EXPORTACION")
         '
         CODCARPE$ = TRIM$(VALACT1$("!INCARPIM"))
         If CODCARPE$ <> "" Then
            Call CONECRUN("ACRGES07/FACTURA;CODCARPE =" & CODCARPE$, "Facturacion de Proveedores")
         End If

End Sub
    
Sub GASTOSNACOMP()

         CONDI$ = "CodcarpImp <>'' AND STATUS < 9 AND MODO = 1"
         CONDI$ = CONDI$ + "ORDER BY CodCarpImp"
         Call CARGALISEL("!INCARPIM", "CARPEIMP", "CodcarpImp/A12;FechAlta/d8;RefeCarpImp/A48", CONDI$, "NOMESS")
         Call SELECHO("!INCARPIM/INDICE DE CARPETAS DE EXPORTACION")
         '
         CODCARPE$ = TRIM$(VALACT1$("!INCARPIM"))
         If CODCARPE$ <> "" Then
            Call CONECRUN("ACRGES07/ORPAGO;;CODCARPE =" & CODCARPE$, "Facturacion de Proveedores")
         End If
End Sub

Sub COSCONTCOMP()

'        Call SELECHO("TACENCOS/Tabla de Centro de Costos")
         '
         CONDI$ = "CodcarpImp <>'' AND STATUS < 9 ORDER BY CodCarpImp"
         Call CARGALISEL("!INCARPIM", "CARPEIMP", "CodcarpImp/A12;FechAlta/d8;RefeCarpImp/A48", CONDI$, "NOMESS")
         Call SELECHO("!INCARPIM/INDICE DE CARPETAS DE EXPORTACION")
         '
         CODCARPE$ = TRIM$(VALACT1$("!INCARPIM"))
         '
         If CODCARPE$ <> "" Then
            Call CONECRUN("CONTAB07/ASIEDIAR;CODCARPE =" & CODCARPE$, "Costos Contables")
         End If
End Sub

Sub SOLIMATCOMP()

         Call CONECRUN("SOLIMA07;CODCARPE =" & CODCARPE$, "Solicitud de Materiales")

End Sub
 
