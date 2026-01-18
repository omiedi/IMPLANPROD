VERSION 5.00
Begin VB.Form FILICOS04 
   BackColor       =   &H00D0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado Resumen de Costos Calculados"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   585
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
   ScaleHeight     =   3615
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   4110
      Left            =   7455
      ScaleHeight     =   4050
      ScaleWidth      =   1005
      TabIndex        =   22
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox Picture16 
         Height          =   3600
         Left            =   0
         ScaleHeight     =   3540
         ScaleWidth      =   1035
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   1100
         Begin VB.Image Image3 
            Height          =   3600
            Left            =   0
            Picture         =   "FILICOS04.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   26
         Top             =   2730
         Width           =   650
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
         Left            =   105
         Picture         =   "FILICOS04.frx":2398
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
         Left            =   105
         Picture         =   "FILICOS04.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   1890
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
         Left            =   105
         Picture         =   "FILICOS04.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   880
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "FILICOS04.frx":2AF6
         Top             =   3045
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0FFFF&
      Caption         =   "Formato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   5460
      TabIndex        =   19
      Top             =   2310
      Width           =   1800
      Begin VB.OptionButton Option2 
         BackColor       =   &H00D0FFFF&
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   315
         TabIndex        =   21
         Top             =   300
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0FFFF&
         Caption         =   "Condensado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   20
         Top             =   770
         Width           =   1400
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0FFFF&
      Caption         =   "Selección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   105
      TabIndex        =   11
      Top             =   2310
      Width           =   3165
      Begin VB.CheckBox Check1 
         BackColor       =   &H00D0FFFF&
         Caption         =   "Todos los Tipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Width           =   435
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Todos - Sin Clasificación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   15
         Top             =   735
         Width           =   2715
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
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
         Left            =   -105
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
      BackColor       =   &H00D0FFFF&
      Caption         =   "Rango de Códigos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   7155
      Begin VB.TextBox Text1 
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
         Height          =   360
         Left            =   6720
         TabIndex        =   4
         Top             =   315
         Width           =   175
      End
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
         Left            =   4620
         TabIndex        =   3
         Text            =   " "
         Top             =   1260
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
         Height          =   360
         Left            =   6720
         TabIndex        =   2
         Top             =   1260
         Width           =   175
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   10
         Top             =   1680
         Width           =   6720
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   9
         Top             =   735
         Width           =   6720
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
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
         Left            =   2205
         TabIndex        =   7
         Top             =   420
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta :"
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
         Left            =   2205
         TabIndex        =   6
         Top             =   1365
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00D0FFFF&
      Caption         =   "Filtrado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   3360
      TabIndex        =   0
      Top             =   2310
      Width           =   2010
      Begin VB.OptionButton Option4 
         BackColor       =   &H00D0FFFF&
         Caption         =   "Pr.Terminados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   18
         Top             =   770
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00D0FFFF&
         Caption         =   "Todos los Artículos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   315
         TabIndex        =   17
         Top             =   300
         Width           =   1380
      End
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "FILICOS04"
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
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Picture16.Visible = True
      Me.Picture16.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Option1.BackColor = &HFCD7B6
      Me.Option2.BackColor = &HFCD7B6
      Me.Option3.BackColor = &HFCD7B6
      Me.Option4.BackColor = &HFCD7B6
      Me.Check1.BackColor = &HFCD7B6
      MARCOBARRA.Top = 3200
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If

    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
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


Private Sub grabar_Click()
   Call Command12_Click
End Sub

Private Sub salir_Click()
   Call Command10_Click
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
      If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
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
    If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
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
     Set FORMULASTEMP = Articulos.OpenRecordset("SELECT * FROM Formulas", dbOpenSnapshot, dbReadOnly)
     With FORMULASTEMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         Call MENSERR(24, "B.Datos de Fórmulas Aparece Vacía.")
         Exit Sub
       End If
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
   End If
   '
   Call COPYSTRU("RECOSCAL", "LICOSCAL")
   '
   On Error Resume Next
   Kill LUDAT$ + "LICOSCAL.DAT"
   On Error GoTo 0
   '
   COD1$ = TRIM$(Text1): COD2$ = TRIM$(Text2)
   j& = 0: URECO& = ULTREG&("RECOSCAL")
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
               If Check1.Value = 1 Or TIMATE%(CIXI%) = XVALO(Text3) Then
                 If Option3.Value = True Or (DP%(CIXI%) > 0 And DQ%(CIXI%) = 0) Then
                   RELE& = CIXI%
                   If RELE& <= URECO& Then
                     X$ = REGLEIDO$("RECOSCAL", RELE&)
                     j& = j& + 1
                     Call GRAREG("LICOSCAL", X$, j&)
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
   Call SETULTREG("LICOSCAL", j&)
   Call FILESORT("LICOSCAL", "", 1, 1, "ASC")
   Call CIERRARCH("LICOSCAL")
   '
   Call HEADERS("LICOSCAL", "")
   On Error Resume Next
   FELICO$ = TRIM$(ANACOSTO.Label34(4)) + " - " + TRIM$(ANACOSTO.Label95)
   If Err Then
     FELICO$ = TRIM$(ANACOSTO04.Label34(4)) + " - " + TRIM$(ANACOSTO04.Label95)
   End If
   On Error GoTo 0
   Call HEADERS("LICOSCAL", "Cálculo del: " + FELICO$)
   Screen.MousePointer = 1
   '
   Call FINAL("*LICOSCAL")
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
     Set FORMULASTEMP = Articulos.OpenRecordset("SELECT * FROM Formulas", dbOpenSnapshot, dbReadOnly)
     With FORMULASTEMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         Call MENSERR(24, "B.Datos de Fórmulas Aparece Vacía.")
         Exit Sub
       End If
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
   End If
   '
   Call COPYSTRU("RECOSCON", "LICOSCON")
   '
   On Error Resume Next
   Kill LUDAT$ + "LICOSCON.DAT"
   On Error GoTo 0
   '
   COD1$ = TRIM$(Text1): COD2$ = TRIM$(Text2)
   j& = 0: URECO& = ULTREG&("RECOSCON")
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
               If Check1.Value = 1 Or TIMATE%(CIXI%) = XVALO(Text3) Then
                 If Option3.Value = True Or (DP%(CIXI%) > 0 And DQ%(CIXI%) = 0) Then
                   RELE& = CIXI%
                   If RELE& <= URECO& Then
                     X$ = REGLEIDO$("RECOSCON", RELE&)
                     j& = j& + 1
                     Call GRAREG("LICOSCON", X$, j&)
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
   Call SETULTREG("LICOSCON", j&)
   Call FILESORT("LICOSCON", "", 1, 1, "ASC")
   Call CIERRARCH("LICOSCON")
   '
   Call HEADERS("LICOSCON", "")
   On Error Resume Next
   FELICO$ = TRIM$(ANACOSTO.Label34(4)) + " - " + TRIM$(ANACOSTO.Label95)
   If Err Then
     FELICO$ = TRIM$(ANACOSTO04.Label34(4)) + " - " + TRIM$(ANACOSTO04.Label95)
   End If
   On Error GoTo 0
   Call HEADERS("LICOSCON", "Cálculo del: " + FELICO$)
   Screen.MousePointer = 1
   '
   Call FINAL("*LICOSCON")
   '
End Sub
'\\\OT-06-0299-WAR-FIN///
