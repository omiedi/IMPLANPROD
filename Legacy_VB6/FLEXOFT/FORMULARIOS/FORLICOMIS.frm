VERSION 5.00
Begin VB.Form FORLICOMIS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidación de Comisiones por Venta y Cobranza"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Más ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   9870
      Picture         =   "FORLICOMIS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Opciones ... Cuenta Corriente Vendedores"
      Top             =   3465
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Libro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   9870
      Picture         =   "FORLICOMIS.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Imprime Libro de Viajantes / Click Botón Derecho Ajusta Renglones por Hoja"
      Top             =   2685
      Width           =   855
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Modalidad de Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   210
      TabIndex        =   21
      Top             =   1740
      Width           =   4635
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Cobranza Neta de I.V.A"
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
         Left            =   420
         TabIndex        =   25
         Top             =   1575
         Width           =   3375
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Facturación Neta del Periodo"
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
         Left            =   420
         TabIndex        =   24
         Top             =   1260
         Width           =   3375
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Pedidos Netos y Cobranza"
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
         Left            =   420
         TabIndex        =   23
         Top             =   945
         Width           =   4005
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Cobranza Neta según Artículo"
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
         Left            =   420
         TabIndex        =   22
         Top             =   420
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Opciones de Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   210
      TabIndex        =   19
      Top             =   3950
      Width           =   4635
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Presenta F.C.P. por cada Producto"
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
         Left            =   420
         TabIndex        =   20
         Top             =   420
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   9870
      Picture         =   "FORLICOMIS.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1885
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Totales por Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2195
      Left            =   5040
      TabIndex        =   16
      Top             =   2625
      Width           =   4635
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   210
         TabIndex        =   17
         Top             =   380
         Width           =   4215
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H80000016&
      Height          =   1275
      Left            =   210
      ScaleHeight     =   1215
      ScaleWidth      =   4575
      TabIndex        =   7
      Top             =   210
      Width           =   4635
      Begin VB.TextBox MESAN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         TabIndex        =   10
         Text            =   " "
         Top             =   210
         Width           =   2220
      End
      Begin VB.TextBox Text4 
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
         Index           =   1
         Left            =   3045
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   750
      End
      Begin VB.TextBox Text4 
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
         Index           =   0
         Left            =   1575
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Período:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   10
         Left            =   -735
         TabIndex        =   15
         Top             =   315
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Generado el:"
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
         Index           =   13
         Left            =   -735
         TabIndex        =   14
         Top             =   840
         Width           =   2220
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hs"
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
         Index           =   12
         Left            =   3885
         TabIndex        =   13
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   225
         Left            =   2835
         TabIndex        =   12
         Top             =   840
         Width           =   15
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2730
         TabIndex        =   11
         Top             =   735
         Width           =   225
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Avance de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   5040
      TabIndex        =   2
      Top             =   210
      Width           =   4635
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   525
         ScaleHeight     =   75
         ScaleWidth      =   3525
         TabIndex        =   26
         Top             =   1785
         Width           =   3585
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cálculo de Comisiones"
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
         Left            =   525
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1365
         Width           =   2745
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Valorización por Producto"
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
         Left            =   525
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1050
         Width           =   2745
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Detalle de Facturas por Recibo"
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
         Left            =   525
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   735
         Width           =   3690
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Selección de Cobranza del Período"
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
         Left            =   525
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   420
         Width           =   3585
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   9870
      Picture         =   "FORLICOMIS.frx":0C7E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1100
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Height          =   750
      Left            =   9870
      Picture         =   "FORLICOMIS.frx":0F88
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   9345
      Picture         =   "FORLICOMIS.frx":10D2
      Top             =   4305
      Width           =   2010
   End
End
Attribute VB_Name = "FORLICOMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    Check1.Value = 0
End Sub

Private Sub Check2_Click()
    Check2.Value = 0
End Sub

Private Sub Check3_Click()
    Check3.Value = 0
End Sub

Private Sub Check4_Click()
    Check4.Value = 0
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    FORLICOMIS.Hide
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Command3.Enabled = False
    YALIQUI% = 0
    MODOCAL% = 0
    If Option2.Value = True Then MODOCAL% = 1
    If Option3.Value = True Then MODOCAL% = 2
    If Option4.Value = True Then MODOCAL% = 3
    If Option5.Value = True Then MODOCAL% = 4
    If MODOCAL% < 1 Then
        Call MENSERR(24, "Falta Seleccionar Modalidad de Cálculo")
        GoTo 99
    End If
    If MODOCAL% = 1 Then
        If InStr(UCase$(EMPREAC$), "NOVOESPU") > 0 Then
            Call CALCOMIS
          Else
            Call OPNOIN("")
        End If
        GoTo 99
    End If
    '
    If MODOCAL% > 1 Then
'        If InStr(UCase$(EMPREAC$), "SABO") > 0 Then
            Call LICOMIS
'          Else
'            Call OPNOIN("")
'        End If
        GoTo 99
    End If
    '
99  Command2.Enabled = True
    If YALIQUI% = 1 Then
        Command3.Enabled = True
        Command4.Enabled = True
    End If
    '
End Sub

Private Sub Command3_Click()
    Call PRILICOMIS
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call PRILIVIAJAN
    Command4.Enabled = True
End Sub

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    CAHOJLIA% = Val(CONTROL$("LIBROVIA", "CAHOJLI"))
    CAHOJLI% = InputBox$("Renglones por Hoja", "Libro de Viajantes", Str$(CAHOJLIA%))
    Call GRACONTROL("LIBROVIA", "CAHOJLI", Str$(CAHOJLI%))
    '
    J& = 0
    If CAHOJLI% > 0 Then
      If List1.ListIndex >= 0 Then
        TXT$ = List1.TEXT
        If Abs(Val(Mid$(TXT$, 23, 10))) >= 0.005 Then
          VENDE% = Val(Left$(TXT$, 3))
          Screen.MousePointer = 11
          FIN& = ULTREG("LIBROVIA")
          For U& = 1 To FIN&
            Call TRACE(20, U&, FIN&)
            YKU$ = REGLEIDO$("LIBROVIA", U&)
            If CVI(Mid$(YKU$, 127, 2)) = VENDE% Then
              J& = J& + 1
            End If
          Next U&
          HOJNECE& = 1 + Int((J& - 1) / CAHOJLI%)
          Screen.MousePointer = 1
          MsgBox TRIM$(Str$(J&)) + " Renglones - Se necesitan " + TRIM$(Str$(HOJNECE&)) + " Hojas"
        End If
      End If
    End If
    '
  End If
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
10 AMODO% = ALTERNA%("Cuenta Corriente Vendedores\Informar Pagos Y Descuentos\Otros - Accesos Directos")
   If AMODO% = 1 Then
       Call CONECRUN("*CUECOMIS", "Cuenta Corriente Vendedores")
     ElseIf AMODO% = 2 Then
       Call PAGOVEN
       Call CONECRUN("*CUECOMIS", "Cuenta Corriente Vendedores")
       GoTo 10
     ElseIf AMODO% = 3 Then
       Call ACCDIR("LICOMIS")
   End If
   Command5.Enabled = True
End Sub

Private Sub Form_Activate()
    If List1.ListCount < 1 Then
        Command3.Enabled = False
        Command4.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub List1_DblClick()
    AMODO% = ALTERNA%("Listado de Comisiones\Libro de Viajantes")
    If AMODO% = 1 Then
         Call PRILICOMIS
       ElseIf AMODO% = 2 Then
         Call PRILIVIAJAN
       Else
         Exit Sub
    End If
End Sub

Private Sub MESAN_GOTFOCUS()
    Command2.SetFocus
End Sub

Private Sub Option1_Click()
    If Option2.Value = False Then
        Option1.Value = False
    End If
End Sub

Private Sub Option2_Click()
    List1.Clear
    Command3.Enabled = False
    Command4.Enabled = False
End Sub

Private Sub Option3_Click()
    List1.Clear
    Command3.Enabled = False
    Command4.Enabled = False
    Option1.Value = False
End Sub

Private Sub Option4_Click()
    List1.Clear
    Command3.Enabled = False
    Command4.Enabled = False
    Option1.Value = False
End Sub

Private Sub Option5_Click()
    List1.Clear
    Command3.Enabled = False
    Command4.Enabled = False
    Option1.Value = False
End Sub

Private Sub Text4_GOTFOCUS(Index As Integer)
    Command2.SetFocus
End Sub

Sub PAGOVEN()
   '
5  VDEF$ = REGBLAN$("CUECOMIS")
   HOII% = HOY(HOS$)
   Call REPLA(VDEF$, MKI$(HOII%), 2, 2)
   Call REPLA(VDEF$, Chr$(2), 4, 1)
   '
10 OBX$ = OBJPLA$("PAGOCOMIS", VDEF$)
   If OBX$ <> "" Then
     VENDE% = Asc(Left$(OBX$, 1))
     If VENDE% < 1 Or ECOARCH("VENDEDOR", Chr$(VENDE%)) = "" Then
       VDEF$ = OBX$
       Call MENSERR(24, "Falta Vendedor")
       GoTo 10
     End If
     '
     REFE$ = TRIM$(Mid$(OBX$, 9, 32))
     If REFE$ = "" Then
       VDEF$ = OBX$
       Call MENSERR(24, "Falta Referencia")
       GoTo 10
     End If
     '
     IMPO# = CVD(Mid$(OBX$, 49, 8))
     If Abs(IMPO#) < 0.005 Then
       VDEF$ = OBX$
       Call MENSERR(24, "Falta Importe")
       GoTo 10
     End If
     '
     If IMPO# < 0 Then
       Call REPLA(VDEF$, MKD$(Abs(IMPO#)), 41, 8)
       Call REPLA(VDEF$, MKD$(0), 49, 8)
     End If
     '
     Call REGAPP("CUECOMIS", OBX$)
     Call CIERRARCH("CUECOMIS")
     '
     Call COMUNI("Pago Registrado !!!")
     GoTo 5
     '
   End If
   '
End Sub

