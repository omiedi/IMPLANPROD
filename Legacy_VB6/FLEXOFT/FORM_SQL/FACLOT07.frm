VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FACLOT07 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturacion de Lotes"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   34
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   6120
      Left            =   90
      TabIndex        =   6
      Top             =   20
      Width           =   6375
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFC0&
         Height          =   1695
         Left            =   120
         TabIndex        =   21
         Top             =   4320
         Width           =   6080
         Begin VB.Label label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "+ IVA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   5160
            TabIndex        =   32
            Top             =   840
            Width           =   690
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Qt. Bobinas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   -240
            TabIndex        =   31
            Top             =   240
            Width           =   1770
         End
         Begin VB.Label label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Kilos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   0
            Left            =   1920
            TabIndex        =   30
            Top             =   240
            Width           =   1770
         End
         Begin VB.Label label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Importe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   5160
            TabIndex        =   29
            Top             =   240
            Width           =   690
         End
         Begin VB.Label label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo"
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
            Index           =   0
            Left            =   -240
            TabIndex        =   28
            Top             =   960
            Width           =   1770
         End
         Begin VB.Label label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Kilos"
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
            Index           =   0
            Left            =   1920
            TabIndex        =   27
            Top             =   960
            Width           =   1770
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
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
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
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
            Left            =   2280
            TabIndex        =   25
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   24
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
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
            Left            =   120
            TabIndex        =   23
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
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
            Left            =   2280
            TabIndex        =   22
            Top             =   1200
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Height          =   1575
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6080
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
            Left            =   3585
            TabIndex        =   12
            Top             =   240
            Width           =   2145
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
            Left            =   5715
            TabIndex        =   11
            Top             =   240
            Width           =   175
         End
         Begin VB.TextBox Text2 
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
            Left            =   1850
            TabIndex        =   10
            Top             =   1080
            Width           =   585
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
            Left            =   2400
            TabIndex        =   9
            Top             =   1080
            Width           =   175
         End
         Begin VB.Label Label7 
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
            Height          =   300
            Left            =   2760
            TabIndex        =   16
            Top             =   1080
            Width           =   3135
         End
         Begin VB.Label Label30 
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
            Height          =   300
            Left            =   240
            TabIndex        =   15
            Top             =   600
            Width           =   5655
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo"
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
            Left            =   1680
            TabIndex        =   14
            Top             =   300
            Width           =   1770
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lista de Precios"
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
            Index           =   1
            Left            =   0
            TabIndex        =   13
            Top             =   1120
            Width           =   1770
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   6080
         Begin VB.Frame Frame5 
            Caption         =   "Frame3"
            Height          =   1980
            Left            =   3600
            TabIndex        =   20
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame3"
            Height          =   1980
            Left            =   1920
            TabIndex        =   19
            Top             =   260
            Width           =   15
         End
         Begin VB.ListBox List2 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1740
            ItemData        =   "FACLOT07.frx":0000
            Left            =   240
            List            =   "FACLOT07.frx":0002
            MousePointer    =   1  'Arrow
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   17
            Top             =   480
            Width           =   5625
         End
         Begin VB.Label Label14 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "                   Kilos             Ancho(mm)      Lote"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   5625
         End
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   8400
      Left            =   6600
      ScaleHeight     =   8340
      ScaleWidth      =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   4
         Top             =   4440
         Width           =   690
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
      Begin VB.CommandButton Command10 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "FACLOT07.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   975
         Width           =   690
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
         Height          =   690
         Left            =   120
         Picture         =   "FACLOT07.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   225
         Width           =   690
      End
      Begin VB.CommandButton Command4 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "FACLOT07.frx":0458
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba Lotes Asignados"
         Top             =   5280
         Width           =   690
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "FACLOT07.frx":0762
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4680
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FACLOT07.frx":2684
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "FACLOT07.frx":28B8
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label15 
      Height          =   735
      Left            =   6360
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "FACLOT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '
   Call SELECHO("LISTAS")
   LISTPRE$ = TRIM$(VALACT1$("LISTAS"))
   If LISTPRE$ <> "" Then
     Text2.TEXT = LISTPRE$
     Label7 = ECOARCH$("LISTAS", MKI$(XVALO(Text2.TEXT)))
   End If
   '
End Sub

Private Sub Command11_Click()
   '
'   List2.Clear
'   Call LIMPILABEL
'   Label30 = ""
'   Call LIMPITEXT
'   Call CIERRARCH("*.*")
   Hide
   '
End Sub

Private Sub Command2_Click()
    '
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
     Label30 = DESCRIT$(CODINT(COD$))
   End If
   '
End Sub



Private Sub Command4_Click()
    Command4.Enabled = False
        If List2.SelCount < 1 Then
           FORFASQL.Text1 = ""
           Call Command11_Click
           GoTo 99
        End If
        '
        Call CARDETALOTE
        '''''''''''''''''''''
        ASIGLOT07.Label8 = Text1
        
        CANTI$ = Label8
        Call REPLA(TX$, Text1, 1, 16)
        Call REPLA(TX$, DESCRIT$(CODINT(Text1)), 18, 33)
        Call REPLA(TX$, FORMATNUM$(CANTI$, "F10.1"), 54, 12)
        Call REPLA(TX$, TRIM$(FORMATNUM$(1, "I5")), 80, 5)

         For KKU& = 1 To List2.ListCount
            If FACLOT07.List2.Selected(KKU& - 1) = True Then
               YY$ = List2.List(KKU& - 1)
               LOOTEE$ = LOOTEE$ + TRIM$(Mid$(YY$, 33, 16)) & ";"
            End If
        Next KKU&
        Call REPLA(TX$, LOOTEE$, 85, 120)
        ASIGLOT07.List1.AddItem (TX$)
        '
        FORFASQL.MARCONOTA.Visible = False
        Call FORFASQL.CALTOFAC
        Call FORFASQL.CARLISFAC
        'Unload Me
        FACLOT07.Hide
99  Command4.Enabled = True
End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me, Picture14)
End Sub

Private Sub List2_ItemCheck(Item As Integer)
   '
   If Text2 = "" Then
      Call MENSERR(24, "Debe Seleccionar una Lista de Precios")
      For KKU& = 1 To List2.ListCount
        List2.Selected(KKU& - 1) = False
      Next KKU&
      Exit Sub
   End If
   ARCHIPRE$ = TRIM$(Text2): LIPRACTI% = XVALO(Mid$(ARCHIPRE$, 6, 3))
   CCII% = CODINT%(Text1)
   Call LIMPILABEL
   '
   For KKU& = 1 To List2.ListCount
      If List2.Selected(KKU& - 1) = True Then
         KILOO = KILOO + XVALO(Mid$(List2.List(KKU& - 1), 1, 10))
         Label10 = FORMATNUM$(KILOO, "F10." & CNTDC$ & "")
         PREC = PRELISTA#(ARCHIPRE$, CCII%)
         Label11 = FORMATNUM$(KILOO * PREC, "F10.2")
         Label8 = XVALO(Label8) + 1
      End If
      TOTKILO = TOTKILO + XVALO(Mid$(List2.List(KKU& - 1), 1, 10))
      Label13 = FORMATNUM$(TOTKILO - KILOO, "F10.1")
   Next KKU&
   Label12 = List2.ListCount - XVALO(Label8)
   '
End Sub

Private Sub Text1_Change()
    '
    CCII% = CODINT%(TRIM$(Text1))
    List2.Clear
    Label30 = ""
    If CCII% > 0 Then
        Label30 = DESCRIT$(CCII%)
        CADECI = FORMATNUM$(FRACMIN(CCII%), "F10." & CNTDC$ & "")
        SQLX$ = "SELECT * FROM LOTINGRE "
        SQLX$ = SQLX$ + "WHERE COD_INTE = " & CCII%
        SQLX$ = SQLX$ + " AND (CANALTA - CANCONSU) > " & CADECI
        Dim PARDIATMP As ADODB.Recordset
        Set PARDIATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With PARDIATMP
        If .EOF = True Then
            Call MENSERR(24, "No Existen Lotes para el Articulo Seleccionado")
        End If
        Do While Not PARDIATMP.EOF
            '
30          TTXX$ = ""
            Call REPLA(TTXX$, FORMATNUM$(XVALO(!CanAlta), "F10." & CNTDC$ & ""), 1, 10)
            Call REPLA(TTXX$, Unimed$(CCII%), 12, 3)
            Call REPLA(TTXX$, FORMATNUM$(XVALO(!AnchoMili), "F10." & CNTDC$ & ""), 17, 10)
            Call REPLA(TTXX$, "mm", 28, 2)
            Call REPLA(TTXX$, SAFETEXT(!idenlote), 33, 16)
            List2.AddItem (TTXX$)
            '
50      PARDIATMP.MoveNext
        '
        Loop
        End With
        PARDIATMP.Close
    End If
    '
End Sub

Private Sub Text1_DblClick()
    '
    Call Command2_Click
    '
End Sub

Sub LIMPILABEL()
    '
    Label8 = ""
    Label10 = ""
    Label11 = ""
    Label12 = ""
    Label13 = ""
    '
End Sub

Sub LIMPITEXT()
    '
    Text1 = ""
    Text2 = ""
    '
End Sub

Private Sub Text2_Change()
    '
    Label7 = ECOARCH$("LISTAS", MKI$(XVALO(Text2.TEXT)))
    Call LIMPILABEL
    For KKU& = 1 To List2.ListCount
        List2.Selected(KKU& - 1) = False
    Next KKU&
    '
End Sub

Private Sub Text2_DblClick()
    '
    Call Command1_Click
    '
End Sub

Private Sub Text2_LostFocus()
    '
    If Label7 = "" Then
        Call MENSERR(24, "Lista de Precios Seleccionada no Valida")
    End If
    '
End Sub

Sub CARDETALOTE()
    '
    XX$ = ""
    CCII% = CODINT(Text1)
    CANTASI = XVALO(Label10)
    FORFASQL.TEXT21(0) = TRIM$(Text2)
    ARCHIPRE$ = TRIM$(Text2): LIPRACTI% = XVALO(Mid$(ARCHIPRE$, 6, 3))
    PREC = PRELISTA#(ARCHIPRE$, CCII%)
    Import# = XVALO(Label11)
    Call REPLA(XX$, MKS$(CANTASI), 1, 4)
    Call REPLA(XX$, MKI$(CCII%), 5, 2)
    Call REPLA(XX$, DESCRIT$(CCII%), 7, 64)
    Call REPLA(XX$, MKS$(PREC), 429, 4)
    Call REPLA(XX$, MKD$(Import#), 433, 8)
    '
    Call GRAREG("!CARDEFAC", XX$, 1)
    '
    Call SETULTREG("!CARDEFAC", 1)
    Call CIERRARCH("!CARDEFAC")
    '
End Sub
