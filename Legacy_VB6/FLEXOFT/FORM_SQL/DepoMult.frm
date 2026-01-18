VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DepoMult 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock x Depósitos Seleccionados"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox TXTDEPOS 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4320
      Width           =   5175
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Seleccionar Depósitos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Selección de Depósitos para visualizar su stock actual"
      Top             =   4680
      Width           =   2250
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   5160
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4650
      Left            =   5400
      ScaleHeight     =   4620
      ScaleWidth      =   900
      TabIndex        =   7
      Top             =   0
      Width           =   930
      Begin VB.CommandButton Command2 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   3720
         Width           =   660
      End
      Begin VB.CommandButton command1 
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
         Height          =   610
         Left            =   120
         Picture         =   "DepoMult.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "DepoMult.frx":014A
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4200
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Presentación de Costo y Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   3
      Top             =   7770
      Width           =   5160
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Calculado"
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
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Precio Base (C.Rep. + 15 %)"
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
         Left            =   3360
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Informado"
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
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Destino Archivo de Transferencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   0
      Top             =   6720
      Width           =   5160
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   735
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   420
         Width           =   4230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta:"
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
         Left            =   0
         TabIndex        =   2
         Top             =   525
         Width           =   645
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3585
      OleObjectBlob   =   "DepoMult.frx":206C
      Top             =   -105
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5040
      Top             =   6950
      Width           =   2010
   End
End
Attribute VB_Name = "DepoMult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub





Private Sub Command2_Click()
   Command2.Enabled = False
    'retorna el stock por cada deposito genera una columna para cada deposito seleccionado IA
   Campos$ = "ci/A0;" '1
   Campos$ = Campos$ + "Código/A20;" '2
   Campos$ = Campos$ + "Descripción/A40" '3
    
   Erase CADENATEX$
   If TRIM$(TXTDEPOS) <> "" Then
      Call TEXTOLOTES(TXTDEPOS, ",")
        For i& = 1 To UBound(CADENATEX$)
           Campos$ = Campos$ + ";Dep." & CADENATEX$(i&) & "/A12"
        Next i&
   End If
   
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    TX$ = "Stock x Depósitos (" & TXTDEPOS & ")"
  
    GRILLAGRAL.Caption = TX$

    SQLX$ = " SELECT COD_INTE,COD_EXTE,DESCRIP FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE (COD_EXTE <> '' OR COD_EXTE IS NOT NULL)"
    SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP "

    indDepo% = 0
    Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
             
                J& = J& + 1
                CI1% = XVALO(!cod_inte)
                GRILLAGRAL.GRID.Rows = J& + 1
                GRILLAGRAL.GRID.TextMatrix(J&, 1) = CI1%
                GRILLAGRAL.GRID.TextMatrix(J&, 2) = SAFETEXT$(!Cod_Exte)
                GRILLAGRAL.GRID.TextMatrix(J&, 3) = SAFETEXT$(!Descrip)
                For COL& = 4 To GRILLAGRAL.GRID.Cols - 1
                     indDepo% = indDepo% + 1
                     CantEnDeposito = STODEPOS(CI1%, Int(CADENATEX$(indDepo%))) ' Y quiere que lo presente
                    GRILLAGRAL.GRID.TextMatrix(J&, COL&) = FORMATNUM$(CantEnDeposito, "F12.1")
                Next COL&
                indDepo% = 0
'             End If
            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing

    '*******************************************
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    
    GRILLAGRAL.SinEfectoClickEnGrilla% = 1

    
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "INFOSCRP", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If
99 Command2.Enabled = True


End Sub



Private Sub Command15_Click()
   Command15.Enabled = False
   SELDEPOS.Show 1
   STRINGDEPOS$ = ""
   List1.Clear
   
   For U& = 1 To SELDEPOS.LISNUM.ListCount
       If SELDEPOS.LISNUM.Selected(U& - 1) = True Then
        TXTX$ = SELDEPOS.LISNUM.List(U& - 1)
        List1.AddItem TXTX$
        NDEP$ = TRIM$(Str$(Mid$(TXTX$, 1, 6)))
        STRINGDEPOS$ = STRINGDEPOS$ + NDEP$ + ","
      End If
     Next U&
     If TRIM$(STRINGDEPOS$) <> "" Then
        Me.TXTDEPOS = Mid$(STRINGDEPOS$, 1, Len(STRINGDEPOS$) - 1) 'para quitar la coma
     Else
        TXTDEPOS = ""
     End If

   Command15.Enabled = True
End Sub

