VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form InstrucSuc 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instrucciones de Sucursal"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   13620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2040
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7680
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13440
      Begin VB.CommandButton Command1 
         Caption         =   "Eliminar Instrucción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   14
         Top             =   7200
         Width           =   1620
      End
      Begin VB.TextBox TXTDESCRIPCION 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   4200
         Width           =   13020
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Grabar Instrucción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   11520
         TabIndex        =   7
         Top             =   7200
         Width           =   1620
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   3195
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   13035
         _ExtentX        =   22992
         _ExtentY        =   5636
         _Version        =   393216
         Cols            =   5
         ForeColor       =   0
         WordWrap        =   -1  'True
         AllowBigSelection=   -1  'True
         Enabled         =   -1  'True
         FocusRect       =   2
         ScrollBars      =   2
         AllowUserResizing=   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   12390
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Número deSucursal Seleccionada"
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
         Left            =   10440
         TabIndex        =   12
         Top             =   120
         Width           =   1860
      End
      Begin VB.Label Ltext2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   2400
         TabIndex        =   10
         Top             =   240
         Width           =   6735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Cuenta"
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
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   900
      End
      Begin VB.Label ltext1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   990
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         Caption         =   "TEPRUEBA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   4
         Top             =   3840
         Visible         =   0   'False
         Width           =   12975
         WordWrap        =   -1  'True
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "InstrucSuc.frx":0000
      Top             =   4800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4800
      OleObjectBlob   =   "InstrucSuc.frx":0234
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   135
      Left            =   6480
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "InstrucSuc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Command14_Click()
    Command14.Enabled = False
    If XVALO(Label4) < 1 Then
       Call COMUNI("No se Seleccionó una Sucursal de la Grilla")
       GoTo 99
    End If
    '
    CONDI$ = "nume_cli = " & XVALO(ltext1)
    CONDI$ = CONDI$ + " And Nume_Suc = " & XVALO(Label4)
    Call ACTUREGISTRO("SUCENTRE", "InstruccionSuc", CONDI$, Left$(TXTDESCRIPCION, 512), REGAF&, "NOMESS")
    Call CargarInstrucciones
    Label4 = ""
    Me.TXTDESCRIPCION = ""
99  Command14.Enabled = True

End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   If XVALO(Label4) < 1 Then
     Call COMUNI("No se Seleccionó una Sucursal de la Grilla")
     GoTo 99
   End If
   TXTDESCRIPCION.TEXT = ""
   Call Command14_Click
99 Command1.Enabled = True
   
End Sub

Private Sub Command22_Click()

End Sub





Private Sub Command3_Click()
       
End Sub

Private Sub Form_Load()
    '
    Label4 = ""
    Campos$ = "Sucursal/F8;Descripción/A32;Instrucción/A80"
    Call CARGA_ENCABEZADO(Me.MSF, Campos$, Me)
End Sub




Private Sub Text19_Change()

   Call CargarInstrucciones
End Sub


Sub CargarInstrucciones()

    Me.MSF.Rows = 1
    
    NC = XVALO(ltext1)
    
    If NC < 1 Then Exit Sub
    

    CONDI$ = "Nume_Cli = " & NC
    CONDI$ = CONDI$ + " AND (MARBORRA < 1 or MARBORRA is null) AND STATUS >=0"
    
    If MSF.Rows < 2 Then MSF.Rows = 2
    '
    SQLX$ = "SELECT nume_cli,Nume_Suc,deno_suc,InstruccionSuc FROM SUCENTRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        REG& = REG& + 1
        MSF.Rows = REG& + 1
        INSTRU$ = SAFETEXT$(!InstruccionSuc)
        Label22 = INSTRU$
        MSF.RowHeight(REG&) = Label22.Height + 100
        MSF.TextMatrix(REG&, 1) = SAFETEXT$(!Nume_Suc)
        MSF.TextMatrix(REG&, 2) = SAFETEXT$(!deno_suc)
        MSF.TextMatrix(REG&, 3) = SAFETEXT$(!InstruccionSuc)
       .MoveNext
      Loop
    End With

    RS.Close
    Set RS = Nothing
    Label22 = ""
End Sub



Private Sub Label4_Change()
    TXTDESCRIPCION.TEXT = ""
    If XVALO(Label4) > 0 Then
     TXTDESCRIPCION.TEXT = MSF.TextMatrix(XVALO(MSF.Row), 3)
    End If
End Sub

Private Sub ltext1_Change()
    Ltext2 = rasocli$(XVALO(ltext1))
    Call CargarInstrucciones
End Sub

Private Sub MSF_Click()
    If MSF.Row >= 1 Then
     Label4 = MSF.TextMatrix(MSF.Row, 1)
     Call SELECCIONARFILA(MSF, XVALO(MSF.Row))
     Me.TXTDESCRIPCION.SetFocus
     Call POSICIONARALFINALTXT(Me.TXTDESCRIPCION)
    End If

End Sub
