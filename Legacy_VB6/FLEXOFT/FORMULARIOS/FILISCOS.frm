VERSION 5.00
Begin VB.Form FILISCOS 
   BackColor       =   &H00D0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado Resumen de Costos Calculados"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7920
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
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   6930
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   22
      Top             =   2655
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   12000
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
      Left            =   210
      TabIndex        =   14
      Top             =   2310
      Width           =   3585
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
         Height          =   240
         Left            =   210
         TabIndex        =   19
         Top             =   465
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
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
         Left            =   3150
         TabIndex        =   16
         Top             =   420
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
         Left            =   2730
         TabIndex        =   15
         Top             =   420
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
         TabIndex        =   18
         Top             =   735
         Width           =   3135
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
         Left            =   1890
         TabIndex        =   17
         Top             =   465
         Width           =   750
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   8610
      Sorted          =   -1  'True
      TabIndex        =   11
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
      Left            =   210
      TabIndex        =   4
      Top             =   105
      Width           =   6525
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
         Left            =   3990
         TabIndex        =   8
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
         Left            =   6090
         TabIndex        =   7
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
         Left            =   3990
         TabIndex        =   6
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
         Left            =   6090
         TabIndex        =   5
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
         TabIndex        =   13
         Top             =   1680
         Width           =   6090
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
         TabIndex        =   12
         Top             =   735
         Width           =   6090
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
         Left            =   1575
         TabIndex        =   10
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
         Left            =   1575
         TabIndex        =   9
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
      Left            =   3990
      TabIndex        =   3
      Top             =   2310
      Width           =   2745
      Begin VB.OptionButton Option4 
         BackColor       =   &H00D0FFFF&
         Caption         =   "Productos Terminados"
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
         TabIndex        =   21
         Top             =   735
         Value           =   -1  'True
         Width           =   2325
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
         Height          =   240
         Left            =   315
         TabIndex        =   20
         Top             =   420
         Width           =   2220
      End
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
      Height          =   750
      Left            =   6900
      Picture         =   "FILISCOS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   990
      Width           =   855
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
      Height          =   750
      Left            =   6900
      Picture         =   "FILISCOS.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Emite Reporte Solicitado"
      Top             =   1785
      Width           =   855
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
      Height          =   750
      Left            =   6900
      Picture         =   "FILISCOS.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   210
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   6350
      Picture         =   "FILISCOS.frx":075E
      Top             =   2950
      Width           =   2010
   End
End
Attribute VB_Name = "FILISCOS"
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
    cod1$ = TRIM$(VALACT1$("MASTER"))
    If cod1$ <> "" Then
       Text1.TEXT = cod1$
    End If
End Sub

Private Sub command10_Click()
   Call CIERRARCH("*.*")
   Unload FILISCOS
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call LILISCOS
   Command12.Enabled = True
End Sub

Private Sub Command2_Click()
    Call SELECHO("MASTER")
    cod2$ = TRIM$(VALACT1$("MASTER"))
    If cod2$ <> "" Then
       Text2.TEXT = cod2$
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


Private Sub Text1_DblClick()
    Call SELECHO("MASTER")
    cod1$ = TRIM$(VALACT1$("MASTER"))
    If cod1$ <> "" Then
       Text1.TEXT = cod1$
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
    cod2$ = TRIM$(VALACT1$("MASTER"))
    If cod2$ <> "" Then
       Text2.TEXT = cod2$
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
   NVI% = NUINV%
   '
   If Option4.Value = True Then
     ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
     FESTRUC# = FileDateTime(LUDAT$ + "ESTRUNUE.DAT")
     If FESTRUC# <> FESTRU# Then
       FESTRU# = FESTRUC#
       FIN& = ULTREG&("ESTRUNUE")
       For I& = 1 To FIN&
         Call TRACE(20, I&, FIN&)
         R0$ = REGLEIDO$("ESTRUNUE", I&)
         CI1% = CVI(Left$(R0$, 2))
         If CI1% > 0 Then
           If CI1% <= NMM% Then
             If CVS(Mid$(R0$, 5, 4)) > 0.00005 Then
               DP%(CI1%) = 1
             End If
             R1$ = REGLEIDO$("ESTRUDOS", I&)
             R2$ = Mid$(R0$, 3, 120) + R1$
             For KKI% = 1 To 376 Step 8
               CI2% = CVI(Mid$(R2$, KKI%, 2))
               If CI2% > 0 Then
                 If CI2% <= NMM% Then
                   DQ%(CI2%) = 1
                 End If
               End If
             Next KKI%
           End If
         End If
       Next I&
     End If
   End If
   '
   Call COPYSTRU("RECOSCAL", "LICOSCAL")
   '
   On Error Resume Next
   Kill LUDAT$ + "LICOSCAL.DAT"
   On Error GoTo 0
   '
   cod1$ = Text1: cod2$ = Text2
   J& = 0: URECO& = ULTREG&("RECOSCAL")
   FIN& = NUINV%
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      I0$ = REGLEIDO$("INVERT", U& + 1)
      I1$ = Left$(I0$, 16)
      If I1$ >= cod1$ Then
        If I1$ <= cod2$ Then
          CIXI% = CVI(Mid$(I0$, 17, 2))
          If CIXI% > 0 Then
            If CIXI% <= NMM% Then
              If Check1.Value = 1 Or TIMATE%(CIXI%) = Val(Text3) Then
                If Option3.Value = True Or (DP%(CIXI%) > 0 And DQ%(CIXI%) = 0) Then
                  RELE& = CIXI%
                  If RELE& <= URECO& Then
                    X$ = REGLEIDO$("RECOSCAL", RELE&)
                    J& = J& + 1
                    Call GRAREG("LICOSCAL", X$, J&)
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
   Next U&
   '
   Call SETULTREG("LICOSCAL", J&)
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
