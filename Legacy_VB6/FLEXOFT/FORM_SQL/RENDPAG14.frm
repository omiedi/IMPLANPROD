VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form RENDPAG14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirmación de Pagos"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   10935
      Begin VB.Frame Frame3 
         Height          =   1695
         Left            =   4320
         TabIndex        =   17
         Top             =   240
         Width           =   6495
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   "Total Seleccionado:"
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1245
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "Imorte Total:"
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   840
            Width           =   1245
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   840
            Width           =   1365
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "Cantidad ItemsTotal:"
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "Cant. Item Selecc.:"
            ToolTipText     =   "Cantidad de Item Seleccionados"
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1365
         End
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
            Height          =   525
            Left            =   3240
            Picture         =   "RENDPAG14.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Refresca - Invierte Selección"
            Top             =   240
            Width           =   3165
         End
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         MaxLength       =   64
         TabIndex        =   16
         Top             =   6800
         Width           =   7410
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Referencia"
         Top             =   6800
         Width           =   975
      End
      Begin VB.Frame Frame5 
         Height          =   1695
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3015
         Begin VB.CommandButton Command4 
            Caption         =   "Listar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   14
            Top             =   1320
            Width           =   2680
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   960
            Width           =   1005
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "Fecha - Generación"
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   600
            Width           =   1850
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "Usuario:"
            Top             =   600
            Width           =   860
         End
         Begin VB.CommandButton Command2 
            Caption         =   "!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   240
            Width           =   195
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Text            =   "Listas Generadas"
            Top             =   240
            Width           =   1550
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1880
            TabIndex        =   7
            Top             =   240
            Width           =   930
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Registrar Pagos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8640
         TabIndex        =   0
         Top             =   6800
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   225
         Index           =   0
         Left            =   0
         OleObjectBlob   =   "RENDPAG14.frx":014A
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   4635
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   8176
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   16501405
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   375
      Index           =   0
      Left            =   2640
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "RENDPAG14.frx":01A8
      Top             =   120
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "RENDPAG14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()
   Command1.Enabled = False
   '
   If MSF1.Rows < 2 Then
      Call COMUNI("Falta Detalle de Lista de Impuestos a Pagar")
      GoTo 99
   End If
   '
   If XVALO(Text1) < 1 Then
      Call COMUNI("No Hay Ninguna Fila Seleccionada")
      GoTo 99
   End If
   '
   If TRIM$(Text5) = "" Then
     Call COMUNI("Falta Referencia de Lista")
     Call PINTATEXT(Text5)
     Text5.SetFocus
     GoTo 99
   End If
   '
   CONDI$ = "NUMLISTA > 0 "
   NUMLI& = XVALO(Text2)
   NUSER1% = USNBR% Mod 100
   On Error GoTo ERROR_GRABAR
   Dim RS1 As ADODB.Recordset
   Set RS1 = New ADODB.Recordset
25 On Error Resume Next
   SQLX$ = "SELECT * FROM LISTAIMPU  WHERE nuclien < 1  "
   RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
   FLEXCONN.BeginTrans
   With RS1
        For I& = 1 To MSF1.Rows - 1
         
          If MSF1.TextMatrix(I&, 0) = "*" Then
            NC% = XVALO(MSF1.TextMatrix(I&, 1))
            CONCEPTO$ = MSF1.TextMatrix(I&, 3)
            IMPOX# = XVALO(MSF1.TextMatrix(I&, 4))
            CODIMPU% = MSF1.TextMatrix(I&, 5)
            IDENTIF$ = MSF1.TextMatrix(I&, 6)
            '
            Call ACTUREGISTRO("CAJABANC", "STATUS", "ID = '" & IDENTIF$ & "' AND NUMLISTA = " & NUMLI&, 2, REGAF&, "NOMESS")
          End If
        Next I&
        '
    End With
    '
ERROR_GRABAR:
    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      MsgBox "SE HA DETECTADO UN FALLO AL GUARDAR ESTA LISTA,  NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      GoTo 99
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizada Exitosamente")
    End If
    RS1.Close
    Set RS1 = Nothing
    '
    FREP = HOY(HO$)
    FILTX$ = TRIM$(Str$(NUMLI&)) & ";" & TRIM$(FREP)
    Call STDFORM("LIST-GTS", FILTX$)
    '
    Text2 = ""
    Text5 = ""
    Text7 = ""
    MSF1.Rows = 1
    
99 Command1.Enabled = True
End Sub



Private Sub Command2_Click()
   Command2.Enabled = False
   CONDI$ = "NUMLISTA > 0 AND STATUS = 1"
   Call CARGALISEL("LISTEMIT1", "CAJABANC", "NUMLISTA/F6.0;REFERENCIA_LISTA/A32", CONDI$, "DISTINCT/NOMESS")
   Call SELECHO("LISTEMIT1/Listas Emitidas (C/O.Pago Generada)")
   LISTX& = XVALO(VALACT1$("LISTEMITI"))
   If LISTX& < 1 Then GoTo 99
   '
   Text2 = LISTX&
   

99 Command2.Enabled = True
End Sub



Private Sub Command3_Click()
   Command3.Enabled = False
   For I& = 1 To MSF1.Rows - 1
     
     Call COLOREAR_GRILLA_SOLAFILA(RENDPAG14.MSF1, &HFFFF80, I&)
    
   Next I&
   Call ACTUCANTISELECCION
   
   Command3.Enabled = True
   
End Sub
Sub ACTUCANTISELECCION()
   ACUMUCANTI = 0
   For I& = 1 To MSF1.Rows - 1
     If MSF1.TextMatrix(I&, 0) = "*" Then
       ACUMUCANTI = ACUMUCANTI + 1
     End If
   Next I&
   Text1 = ACUMUCANTI
End Sub


Private Sub Command4_Click()
   Text5 = ""
   Text11 = ""
   Text13 = ""
   TXTDESDE = ""
   TXTHASTA = ""
   Text7 = ""
   
   MSF1.Rows = 1
   
   LISTX& = XVALO(Text2)
   SQLX$ = "SELECT * FROM LISTAIMPU WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE NUMLISTA = " & LISTX&
  
   Set RS = READSET(SQLX$)
   With RS
    Do While Not .EOF
          I& = MSF1.Rows
          MSF1.Rows = MSF1.Rows + 1
          '
          NC% = XVALO(!NUCLIEN)
          IMPUESTO$ = SAFETEXT$(!DESCRI_IMPU)
          IMPOX# = XVALO(!IMPORTE)
          CONCEPTO$ = SAFETEXT$(!CONCEPTO)
          CODIMPU% = SAFETEXT$(!NUMEIMPU)
          'VER SI NO HABRIA QUE MOSTRAR USUARIO
          MSF1.TextMatrix(I&, 1) = SAFETEXT$(NC%)
          MSF1.TextMatrix(I&, 2) = RASOCLI$(NC%)
          MSF1.TextMatrix(I&, 3) = CONCEPTO$
          MSF1.TextMatrix(I&, 4) = FORMATNUM$(IMPOX#, "F12.2")
          MSF1.TextMatrix(I&, 5) = SAFETEXT$(CODIMPU%)
          Text5 = SAFETEXT$(!REFERENCIA_LISTA)
          Text11 = USERNAM$(XVALO(!NUMEUSER))
          Text13 = SAFETEXT$(!FECHEMISI)
          '
          .MoveNext
     Loop
   End With
   '
'   Call ACTUTOTALES
End Sub

Sub ACTUTOTALLISTA()
   '
   ACUMU# = 0
   For I& = 1 To MSF1.Rows - 1
      IMPOX# = ROUND(XVALO(MSF1.TextMatrix(I&, 4)), 2)
      ACUMU# = ACUMU# + IMPOX#
   Next I&
   TXTTOTAL = FORMATNUM$(ACUMU#, "F12.2")
   '
End Sub

Private Sub Form_Load()

    Call CENTRAFORM(Me)
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    Call CARGA_ENCABEZADO(RENDPAG14.MSF1, "Cta/F6;Nombre/A32;Concepto/A32;Importe/F12.2;Cod_Impuesto/A0", RENDPAG14)
    '
    Call CREACAMPO("", "LISTAIMPU", "MARBORRA", 3, 0)
    Call CREACAMPO("", "LISTAIMPU", "NUMEUSER", 3, 0)
    Call CREACAMPO("", "LISTAIMPU", "NUMEIMPU", 6, 0, 1)
    Call CREACAMPO("", "LISTAIMPU", "NUCLIEN", 6, 0, 1)
    Call CREACAMPO("", "LISTAIMPU", "IMPORTE", 7, 0)
    Call CREACAMPO("", "LISTAIMPU", "NUMLISTA", 6, 0, 1)
    Call CREACAMPO("", "LISTAIMPU", "FECHEMISI", 8, 0)
    Call CREACAMPO("", "LISTAIMPU", "CONCEPTO", 10, 64)
    Call CREACAMPO("", "LISTAIMPU", "DESCRI_IMPU", 10, 64)
    Call CREACAMPO("", "LISTAIMPU", "STATUS", 3, 0, 1)
    Call CREACAMPO("", "LISTAIMPU", "REFERENCIA_LISTA", 10, 64)
    '
End Sub
Private Sub MSF1_DblClick()
    I& = MSF1.MouseRow
    j& = MSF1.MouseCol
    If I& = 0 Or I& > MSF1.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF1.Row = I&
    MSF1.COL = j&
    On Error Resume Next

    REG& = I&
    If REG& < 1 Then
        Exit Sub
    End If
    Call COLOREAR_GRILLA_SOLAFILA(RENDPAG14.MSF1, &HFFFF80, REG&)
    Call ACTUCANTISELECCION
End Sub

