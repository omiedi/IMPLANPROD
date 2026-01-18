VERSION 5.00
Begin VB.Form FILICOS07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado Resumen de Costos Calculados"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8340
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
   ScaleHeight     =   3420
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   7455
      ScaleHeight     =   4080
      ScaleWidth      =   1035
      TabIndex        =   22
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   27
         Top             =   2880
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton command10 
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
         Height          =   690
         Left            =   120
         Picture         =   "filicos07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   200
         Width           =   650
      End
      Begin VB.CommandButton Command12 
         Caption         =   "O.K."
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
         Picture         =   "filicos07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   1800
         Width           =   650
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Help"
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
         Picture         =   "filicos07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   960
         Width           =   650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "FORMATO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   5460
      TabIndex        =   19
      Top             =   2160
      Width           =   1800
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   315
         TabIndex        =   21
         Top             =   300
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Condensado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   315
         TabIndex        =   20
         Top             =   770
         Width           =   1400
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "SELECCIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   105
      TabIndex        =   11
      Top             =   2160
      Width           =   3165
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Todos los Tipos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1785
         TabIndex        =   16
         Top             =   315
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   13
         Top             =   350
         Width           =   175
      End
      Begin VB.TextBox Text3 
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
         Height          =   285
         Left            =   735
         TabIndex        =   12
         Top             =   350
         Width           =   420
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Todos - Sin Clasificación"
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
         Height          =   285
         Left            =   210
         TabIndex        =   15
         Top             =   735
         Width           =   2715
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
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
         Left            =   -200
         TabIndex        =   14
         Top             =   400
         Width           =   750
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   8610
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   105
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CDDADA&
      Caption         =   "RANGO DE CÓDIGOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   120
      TabIndex        =   1
      Top             =   105
      Width           =   7155
      Begin VB.TextBox Text1 
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
         Left            =   4620
         TabIndex        =   5
         Top             =   315
         Width           =   2115
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
         Left            =   6720
         TabIndex        =   4
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text2 
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
         Left            =   4620
         TabIndex        =   3
         Text            =   " "
         Top             =   1140
         Width           =   2115
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
         Left            =   6720
         TabIndex        =   2
         Top             =   1140
         Width           =   175
      End
      Begin VB.Label Label2 
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
         Left            =   285
         TabIndex        =   10
         Top             =   1500
         Width           =   6600
      End
      Begin VB.Label Label1 
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
         Left            =   290
         TabIndex        =   9
         Top             =   670
         Width           =   6600
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
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
         Left            =   3645
         TabIndex        =   7
         Top             =   420
         Width           =   885
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Left            =   3570
         TabIndex        =   6
         Top             =   1245
         Width           =   960
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "FILTRADO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   3360
      TabIndex        =   0
      Top             =   2160
      Width           =   2010
      Begin VB.OptionButton Option5 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Art. con Precio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   200
         TabIndex        =   26
         Top             =   900
         Width           =   1700
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Pr.Terminados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   200
         TabIndex        =   18
         Top             =   640
         Width           =   1590
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Todos los Artículos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   200
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   1500
      End
   End
End
Attribute VB_Name = "FILICOS07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
   Text3.SetFocus
End Sub

Private Sub Check1_LostFocus()
   If Check1.Value = 1 Then
      Text3.TEXT = ""
      Label4.Caption = "Todos - Sin Clasificación"
   End If
End Sub

Private Sub Command1_Click()
    Call SELECHO("MASTER")
    COD1$ = TRIM$(VALACT1$("MASTER"))
    If COD1$ <> "" Then
       Text1.TEXT = COD1$
    End If
End Sub

Private Sub Command10_Click()
   Call CIERRARCH("*.*")
   Unload Me
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   '\\\OT-06-0299-WAR-17/07/06///
   'Call LILISCOS
   If Option2.Value = True Then
     Call LILISCOS
   ElseIf Option1.Value = True Then
     Call LICOSCON 'LISTADO CONDENSADO
   End If
   '\\\OT-06-0299-WAR-FIN///
   Command12.Enabled = True
End Sub

Private Sub Command2_Click()
    Call SELECHO("MASTER")
    COD2$ = TRIM$(VALACT1$("MASTER"))
    If COD2$ <> "" Then
       Text2.TEXT = COD2$
    End If
End Sub

Private Sub Command3_Click()
    Call SELECHO("TATIMAT")
    VDEV$ = VALACT1$("TATIMAT")
    If VDEV$ <> "" Then
       Text3.TEXT = VDEV$
       If TRIM$(Text3.TEXT) = "" Then
            Label4.Caption = "Todos - Sin Clasificación"
            Check1.Value = 1
          Else
            Label4.Caption = VALACT2$("TATIMAT")
            Check1.Value = 0
       End If
    End If
End Sub


Private Sub Form_Load()
    '
'    EPAC$ = TRIM$(EMPREAC$)
'    If EPAC$ <> "" Then
'        Caption = Caption + "  -  " + EPAC$
'    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    '
    Call APLICACIONSKINS(Me, Picture1)
    '
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16)
    Text1.TEXT = XX$
    CI1% = Abs(CODINT%(XX$))
    Label1.Caption = DESCRIT$(CI1%)
    '
    AAA& = ULTREG&("ECOMAST")
10  XX$ = Left$(REGLEIDO$("ECOMAST", AAA&), 16)
    If TRIM$(XX$) = "" Then
      If AAA& > 1 Then
        AAA& = AAA& - 1
        GoTo 10
      End If
    End If
    Text2.TEXT = XX$
    CI2% = Abs(CODINT%(XX$))
    Label2.Caption = DESCRIT$(CI2%)
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub


Private Sub Text1_DblClick()
    Call SELECHO("MASTER")
    COD1$ = TRIM$(VALACT1$("MASTER"))
    If COD1$ <> "" Then
       Text1.TEXT = COD1$
    End If
End Sub

Private Sub Text2_Change()
    CI2% = CODINT%(Text2.TEXT)
    If CI2% > 0 Then
         Label2.Caption = DESCRIT$(CI2%)
       Else
         Label2.Caption = ""
    End If
End Sub

Private Sub Text2_DblClick()
    Call SELECHO("MASTER")
    COD2$ = TRIM$(VALACT1$("MASTER"))
    If COD2$ <> "" Then
       Text2.TEXT = COD2$
    End If
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("TATIMAT")
    VDEV$ = VALACT1$("TATIMAT")
    If VDEV$ <> "" Then
       Text3.TEXT = VDEV$
       If TRIM$(Text3.TEXT) = "" Then
            Label4.Caption = "Todos - Sin Clasificación"
            Check1.Value = 1
          Else
            Label4.Caption = VALACT2$("TATIMAT")
            Check1.Value = 0
       End If
    End If
End Sub

Private Sub Text3_GotFocus()
   If TRIM$(Text3.TEXT) = "" Then
      Check1.Value = 1
      Label4.Caption = "Todos - Sin Clasificación"
   End If
End Sub

Private Sub Text3_LostFocus()
    If Text3.TEXT <> "" Then
      If Val(Text3.TEXT) < 256 Then
        VENDE% = Val(Text3.TEXT)
        VENDX$ = ECOARCH$("TATIMAT", Chr$(VENDE%))
        If VENDX <> "" Then
            Label4.Caption = VENDX$
            Check1.Value = 0
            Exit Sub
        End If
      End If
      Text3.TEXT = ""
      Label4.Caption = "Todos - Sin Clasificación"
      Check1.Value = 1
    End If
End Sub


Private Sub Text1_Change()
    CI1% = CODINT%(Text1.TEXT)
    If CI1% > 0 Then
         Label1.Caption = DESCRIT$(CI1%)
       Else
         Label1.Caption = ""
    End If
End Sub

Private Sub Text3_Change()
    If Val(Text3.TEXT) < 256 Then
        VENDE% = Val(Text3.TEXT)
        VENDX$ = ECOARCH$("TATIMAT", Chr$(VENDE%))
        If VENDX <> "" Then
            Label4.Caption = VENDX$
            Check1.Value = 0
            Exit Sub
          Else
            Label4.Caption = "Todos - Sin Clasificación"
            Check1.Value = 1
        End If
    End If
End Sub

Sub LILISCOS()
   '
   Static DP%(32767), DQ%(32767), FESTRU#
   '
   Screen.MousePointer = 11
   NMM% = NUMAS%
   '
   If Option4.Value = True Then
     ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
     Call ABRESTRUC
     'Set FORMULASTEMP = Articulos.OpenRecordset("SELECT * FROM Formulas", dbOpenSnapshot, dbReadOnly)
     SQLX$ = "SELECT * FROM Formulas"
     Dim FORMULASTEMP As ADODB.Recordset
     Set FORMULASTEMP = New ADODB.Recordset
     FORMULASTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

     With FORMULASTEMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Call MENSERR(24, "B.Datos de Fórmulas Aparece Vacía.")
         Exit Sub
       End If
       On Error GoTo 0
       '
       Do Until .EOF = True
         CI1% = !CODICONJ
         CI2% = !CODIELEM
         If CI1% > 0 Then
           If CI1% <= NMM% Then
             If !USOBRUTO > 0.00005 Then
               DP%(CI1%) = 1
             End If
           End If
         End If
         If CI2% > 0 Then
           If CI2% <= NMM% Then
             DQ%(CI2%) = 1
           End If
         End If
       .MoveNext
       Loop
     End With
     Set FORMULASTEMP = Nothing
   End If
   '
   Call COPYSTRU("RECOSCAL", "LICOSCAL")
   '
   On Error Resume Next
   Kill LUDAT$ + "LICOSCAL.DAT"
   On Error GoTo 0
   '
   COD1$ = TRIM$(Text1): COD2$ = TRIM$(Text2)
   J& = 0: URECO& = ULTREG&("RECOSCAL")
   FIN& = URECO&
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     CIXI% = U&
     I1$ = CODEXT$(CIXI%)
     If I1$ >= COD1$ Then
       If I1$ <= COD2$ Then
         If CIXI% > 0 Then
           If CIXI% <= NMM% Then
             If STATITEM%(CIXI%) > 0 Then
               If Check1.Value = 1 Or TIMATE%(CIXI%) = Val(Text3) Then
                 If Option3.Value = True Or (Option4.Value = True And DP%(CIXI%) > 0 And DQ%(CIXI%) = 0) Then
                   RELE& = CIXI%
                   If RELE& <= URECO& Then
                     X$ = REGLEIDO$("RECOSCAL", RELE&)
                     J& = J& + 1
                     Call GRAREG("LICOSCAL", X$, J&)
                   End If
                 ElseIf Option5.Value = True Then
                   RELE& = CIXI%
                   '
                   X$ = REGLEIDO$("RECOSCAL", RELE&)
                   PRECIOLISTA# = CVD(Mid$(X$, 105, 8)): LP% = 1 ' PRIMERO TOMA EL PRECIO DE LA LISTA 1
                   If PRECIOLISTA# < 0.01 Then
                        ' BUSCA EL PRECIO EN LA PRIMERA QUE ENCUENTRA
                        LP% = XVALO(VALOBADA("PREVENTA", "COD_LISTA", "COD_PIEZA=" & CIXI%))
                        PRECIOLISTA# = XVALO(VALOBADA("PREVENTA", "PRECIO_UNITARIO", "COD_PIEZA=" & CStr(CIXI%) & " AND COD_LISTA=" & CStr(LP%)))
                        PRECIOLISTA# = PRECIOLISTA# * TICAMONE(MONELISTA%(LP%))
                   End If
                   '
                   If PRECIOLISTA# > 0.005 Then
                        If RELE& <= URECO& Then
                          YY$ = REGBLAN$("LICOSCAF")
                          Call REPLA(YY$, X$, 1, 128)
                          Call REPLA(YY$, Chr$(LP%), 129, 1)
                          Call REPLA(YY$, MKD$(PRECIOLISTA#), 130, 8)
                          J& = J& + 1
                          Call GRAREG("LICOSCAF", YY$, J&)
                        End If
                   End If
                 End If
               End If
             End If
           End If
         End If
       End If
     End If
   Next U&
   '
   archivo$ = "LICOSCAL"
   If Option5.Value = True Then archivo$ = "LICOSCAF"
   Call SETULTREG(archivo$, J&)
   Call FILESORT(archivo$, "", 1, 1, "ASC")
   Call CIERRARCH(archivo$)
   '
   Call HEADERS(archivo$, "")
   On Error Resume Next
   FELICO$ = TRIM$(ANACOSTO.Label34(4)) + " - " + TRIM$(ANACOSTO.Label95)
   If Err Then
     FELICO$ = TRIM$(ANACOSTO07.Label34(4)) + " - " + TRIM$(ANACOSTO07.Label34(5))
   End If
   On Error GoTo 0
   Call HEADERS("LICOSCAL", "Cálculo del: " + FELICO$)
   Screen.MousePointer = 1
   '
   Call FINAL("*" & archivo$)
   '
End Sub
'\\\OT-06-0299-WAR-17/07/06///
Sub LICOSCON()
  '
   Static DP%(32767), DQ%(32767), FESTRU#
   '
   Screen.MousePointer = 11
   NMM% = NUMAS%
   '
   If Option4.Value = True Then
     ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
     Call ABRESTRUC
     'Set FORMULASTEMP = Articulos.OpenRecordset("SELECT * FROM Formulas", dbOpenSnapshot, dbReadOnly)
     SQLX$ = "SELECT * FROM Formulas"
     Dim FORMULASTEMP As ADODB.Recordset
     Set FORMULASTEMP = New ADODB.Recordset
     FORMULASTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

     With FORMULASTEMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Call MENSERR(24, "B.Datos de Fórmulas Aparece Vacía.")
         Exit Sub
       End If
       On Error GoTo 0
       '
       Do Until .EOF = True
         CI1% = !CODICONJ
         CI2% = !CODIELEM
         If CI1% > 0 Then
           If CI1% <= NMM% Then
             If !USOBRUTO > 0.00005 Then
               DP%(CI1%) = 1
             End If
           End If
         End If
         If CI2% > 0 Then
           If CI2% <= NMM% Then
             DQ%(CI2%) = 1
           End If
         End If
       .MoveNext
       Loop
     End With
     Set FORMULASTEMP = Nothing
   End If
   '
   Call COPYSTRU("RECOSCON", "LICOSCON")
   '
   On Error Resume Next
   Kill LUDAT$ + "LICOSCON.DAT"
   On Error GoTo 0
   '
   COD1$ = TRIM$(Text1): COD2$ = TRIM$(Text2)
   J& = 0: URECO& = ULTREG&("RECOSCON")
   FIN& = URECO&
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     CIXI% = U&
     I1$ = CODEXT$(CIXI%)
     If I1$ >= COD1$ Then
       If I1$ <= COD2$ Then
         If CIXI% > 0 Then
           If CIXI% <= NMM% Then
             If STATITEM%(CIXI%) > 0 Then
               If Check1.Value = 1 Or TIMATE%(CIXI%) = Val(Text3) Then
                 If Option3.Value = True Or (DP%(CIXI%) > 0 And DQ%(CIXI%) = 0) Then
                   RELE& = CIXI%
                   If RELE& <= URECO& Then
                     X$ = REGLEIDO$("RECOSCON", RELE&)
                     J& = J& + 1
                     Call GRAREG("LICOSCON", X$, J&)
                   End If
                 End If
               End If
             End If
           End If
         End If
       End If
     End If
   Next U&
   '
   Call SETULTREG("LICOSCON", J&)
   Call FILESORT("LICOSCON", "", 1, 1, "ASC")
   Call CIERRARCH("LICOSCON")
   '
   Call HEADERS("LICOSCON", "")
   On Error Resume Next
   FELICO$ = TRIM$(ANACOSTO.Label34(4)) + " - " + TRIM$(ANACOSTO.Label95)
   If Err Then
     FELICO$ = TRIM$(ANACOSTO07.Label34(4)) + " - " + TRIM$(ANACOSTO07.Label34(5))
   End If
   On Error GoTo 0
   Call HEADERS("LICOSCON", "Cálculo del: " + FELICO$)
   Screen.MousePointer = 1
   '
   Call FINAL("*LICOSCON")
   '
End Sub
'\\\OT-06-0299-WAR-FIN///
