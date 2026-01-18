VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMCONSUMO 
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   14310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command43 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   6600
      Picture         =   "FRMCONSUMO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Refresh"
      Top             =   480
      Width           =   405
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5280
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   480
      Width           =   1290
   End
   Begin VB.TextBox TXTCODIGO 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   " "
      Top             =   10
      Width           =   3210
   End
   Begin VB.TextBox TXTDESCRI 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   " "
      Top             =   10
      Width           =   6210
   End
   Begin VB.TextBox TXTCODIN 
      Alignment       =   1  'Right Justify
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
      Height          =   240
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   810
   End
   Begin VB.CommandButton CMDOT 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4320
      TabIndex        =   8
      ToolTipText     =   "Lista de Selección"
      Top             =   10
      Width           =   375
   End
   Begin VB.TextBox TXTOT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   " "
      Top             =   0
      Width           =   2130
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2220
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3916
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   2220
      Left            =   7200
      TabIndex        =   1
      Top             =   720
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3916
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSF3 
      Height          =   4740
      Left            =   0
      TabIndex        =   2
      Top             =   3240
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   8361
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   4440
      TabIndex        =   14
      Top             =   480
      Width           =   825
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informe de Consumos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   4545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Por Consumo Informado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   7200
      TabIndex        =   6
      Top             =   480
      Width           =   2745
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Por Ficha Técnica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   480
      Width           =   1785
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDEN DE FABARICACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   2025
   End
End
Attribute VB_Name = "FRMCONSUMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MSF10_Click(index As Integer)

End Sub

Private Sub CMDOT_Click()
   NROOF$ = ORFASEL_ENVAPLAST$
   Me.TXTOT = NROOF$
End Sub

Private Sub Command43_Click()
    Command43.Enabled = False
    NORFA$ = Me.TXTOT
    Me.TXTOT = ""
    Me.TXTOT = NORFA$
99  Command43.Enabled = True
   
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call Command43_Click
   End If
End Sub


Private Sub Text1_LostFocus()
   Text1.TEXT = TRIM$(FORMATNUM$(XVALO(Text1), "f14.1"))
End Sub

Private Sub TXTOT_Change()
   NROOF$ = Me.TXTOT
   If NROOF$ <> "" Then
     SQLX$ = "SELECT REFEREN,NROPED, Cod_Inte   FROM ORDEFABR WITH(NOLOCK) WHERE IDSUBOR = '" & NROOF$ & "'"
     Set RS0 = READSET(SQLX$)
     CI1% = XVALO(RS0!COD_INTE)
     Me.TXTCODIN = CI1%
     Me.TXTCODIGO = CODEXT$(CI1%)
     TXTOT.TEXT = NROOF$
     NROPED& = XVALO(RS0!NROPED)
     TXTDESCRI = DESCRIT0$(CI1%)
     TXTECOOF = SAFETEXT$(RS0!Referen)
     '
     Call CARGA_ESTRUCTURA_FICHA(CI1%, XVALO(Text1))
     Call CARGA_CONSUMO_INFORMADO(NROOF$)
     Call CARGA_INFORMECONSUMO
   End If

End Sub
Sub CARGA_INFORMECONSUMO()
    Call CARGA_ENCABEZADO(Me.MSF3, "Código/A12;Descripción/A28;U./A2;Cant-Ficha/F12.1;Cant-Infor./F12.1;DIF.CONSUMO/F12.4;ULTFILA/F0", Me)
    'PRIMERO CARGO LA GRILLA
    For i& = 1 To MSF1.Rows - 1
      CODI$ = MSF1.TextMatrix(i&, 1)
      MSF3.Rows = i& + 1
      MSF3.TextMatrix(i&, 1) = CODI$
      MSF3.TextMatrix(i&, 2) = MSF1.TextMatrix(i&, 2)
      MSF3.TextMatrix(i&, 3) = MSF1.TextMatrix(i&, 3)
      MSF3.TextMatrix(i&, 4) = MSF1.TextMatrix(i&, 4)
      MSF3.TextMatrix(i&, 5) = FORMATNUM$(SUMA_CANT_SEGUNDATO(MSF2, CODI$, 1, 4), "F12.4") 'ACA AGREGO LA CANTIDAD INFORMADA
      MSF3.TextMatrix(MSF3.Rows - 1, 6) = FORMATNUM$(XVALO(MSF3.TextMatrix(MSF3.Rows - 1, 5)) - XVALO(MSF3.TextMatrix(MSF3.Rows - 1, 4)), "F12.4")

      'A LOS ITEMS QUE COINCIDEN
    Next i&
    '
    'AHORA MARCO LOS QUE COINCIDEN HACIENDOLE UN *
    For i& = 1 To MSF1.Rows - 1
      CODI$ = MSF1.TextMatrix(i&, 1)
      For j& = 1 To MSF2.Rows - 1
        If CODI$ = MSF2.TextMatrix(j&, 1) Then MSF2.TextMatrix(j&, 0) = "*"
      Next j&
    Next i&
    '
    'AHORA AGREGO LOS CONSUMOS INFORMADOS QUE NO COINCIDEN CON LOS DE LA FICHA TECNICA
    For i& = 1 To MSF2.Rows - 1
      If TRIM$(MSF2.TextMatrix(i&, 0)) <> "*" Then
        MSF3.Rows = MSF3.Rows + 1
        MSF3.TextMatrix(MSF3.Rows - 1, 1) = MSF2.TextMatrix(i&, 1)
        MSF3.TextMatrix(MSF3.Rows - 1, 2) = MSF2.TextMatrix(i&, 2)
        MSF3.TextMatrix(MSF3.Rows - 1, 3) = MSF2.TextMatrix(i&, 3)
        MSF3.TextMatrix(MSF3.Rows - 1, 4) = ""
        MSF3.TextMatrix(MSF3.Rows - 1, 5) = MSF2.TextMatrix(i&, 4)
        MSF3.TextMatrix(MSF3.Rows - 1, 6) = MSF3.TextMatrix(MSF3.Rows - 1, 5)

      End If
    Next i&

End Sub
Sub CARGA_ESTRUCTURA_FICHA(CI1%, CANTI)
    CODD$ = CODEXT$(CI1%)
    Call LEEEXPLO(CODD$, 1, 1)
    Call CARGA_ENCABEZADO(Me.MSF1, "Código/A12;Descripción/A28;U./A2;Cantidad/F12", Me)
    '
    If CANTI < 1 Then CANTI = 1
    MSF1.Rows = 1
    For i% = 1 To CAIT%
      TXTX$ = EXPLOLIN$(i%)
      MSF1.Rows = i% + 1
      MSF1.TextMatrix(i%, 1) = TRIM$(Mid$(TXTX$, 1, 16))
      MSF1.TextMatrix(i%, 2) = Mid$(TXTX$, 17, 28)
      MSF1.TextMatrix(i%, 3) = Mid$(TXTX$, 47, 2)
      MSF1.TextMatrix(i%, 4) = FORMATNUM$(XVALO(Mid$(TXTX$, 49, 12)) * XVALO(CANTI), "F12.4")
    Next i%
    
End Sub
Sub CARGA_CONSUMO_INFORMADO(NORFA$)
   Call CARGA_ENCABEZADO(Me.MSF2, "Código/A12;Descripción/A28;U./A2;Cantidad/F12", Me)
   Dim obj As New RECORXET
   Set RS1 = obj.RS_CONSUMOX3(NORFA$)
   With RS1
     j& = 0
     Do While Not .EOF
      If XVALO(!CONSUMOX) > 0 And XVALO(!CODINT) > 0 Then
        j& = j& + 1
        MSF2.Rows = j& + 1
        MSF2.TextMatrix(j&, 1) = SAFETEXT$(!Codigo)
        MSF2.TextMatrix(j&, 2) = SAFETEXT$(!Descripcion)
        MSF2.TextMatrix(j&, 3) = UNIMED$(XVALO(!CODINT))
        MSF2.TextMatrix(j&, 4) = FORMATNUM$(XVALO(!CONSUMOX), "F12.4")
      End If
      .MoveNext
     Loop
   End With
   Set obj = Nothing
   
   
   
End Sub


