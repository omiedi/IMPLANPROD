VERSION 5.00
Begin VB.Form MULITISEL09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección Multiple Asociada a un Comprobante Pendiente"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11205
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   10020
      TabIndex        =   20
      Top             =   105
      Width           =   10050
      Begin VB.Line Line2 
         X1              =   0
         X2              =   11580
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"MULITISEL09.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   10005
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4100
      Index           =   5
      Left            =   9915
      TabIndex        =   18
      Top             =   430
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4100
      Index           =   4
      Left            =   8670
      TabIndex        =   17
      Top             =   430
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4100
      Index           =   3
      Left            =   7425
      TabIndex        =   16
      Top             =   430
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4100
      Index           =   2
      Left            =   3930
      TabIndex        =   15
      Top             =   430
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4100
      Index           =   1
      Left            =   3090
      TabIndex        =   14
      Top             =   430
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4100
      Index           =   0
      Left            =   1185
      TabIndex        =   13
      Top             =   430
      Width           =   15
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0095C5E1&
      Height          =   1275
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   10050
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Roboto Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7980
         TabIndex        =   10
         Top             =   900
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Roboto Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7980
         TabIndex        =   9
         Top             =   525
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Diferencia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6920
         TabIndex        =   8
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe Mov. Elegido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6000
         TabIndex        =   7
         Top             =   600
         Width           =   1920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccionados:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6550
         TabIndex        =   6
         Top             =   225
         Width           =   1380
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Roboto Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7980
         TabIndex        =   5
         Top             =   165
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1590
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H002C3D46&
      ForeColor       =   &H80000008&
      Height          =   6330
      Left            =   10320
      ScaleHeight     =   6300
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   -120
      Width           =   1785
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "MULITISEL09.frx":00AF
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Búsqueda por Texto o Numero"
         Top             =   2160
         Width           =   680
      End
      Begin VB.CommandButton BOTEXIT 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "MULITISEL09.frx":1415
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salir sin Procesar"
         Top             =   360
         Width           =   680
      End
      Begin VB.CommandButton BOTREC 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "MULITISEL09.frx":171F
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Procesar Selección Múltiple"
         Top             =   1200
         Width           =   680
      End
   End
   Begin VB.ListBox LICHESEL 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      ItemData        =   "MULITISEL09.frx":1A29
      Left            =   120
      List            =   "MULITISEL09.frx":1A30
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   420
      Width           =   10050
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   3165
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2205
      Visible         =   0   'False
      Width           =   1065
   End
End
Attribute VB_Name = "MULITISEL09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Private Sub BOTREC_Click()
   BOTREC.Enabled = False
   If XVALO(Label9) <> 0 Then
      Call COMUNI("No se Puede Conciliar Con Diferencia de Valores")
      GoTo 99
   End If
   'LEO EL VALOR  DE LA FILA SELECCIONADO EN LA GRILLA 3
   For i& = 1 To CONCIMAN07.MSF3.Rows - 1
     If CONCIMAN07.MSF3.TextMatrix(i&, 0) = "*" Then
        REG3& = i&
        Exit For
     End If
   Next i&
  '
  'VALIDO QUE SE HAYA SELECCIONADO ALMENOS 1
  SELECCIONO% = 0
  For i& = 1 To LICHESEL.ListCount
     If LICHESEL.Selected(i& - 1) = True Then
        SELECCIONO% = 1
        Exit For
     End If
  Next i&
  If SELECCIONO% < 1 Then
     Call COMUNI("No se Ha Seleccionado Ningún Movimiento")
     GoTo 99
  End If
  '
  J& = 0
    For i& = 1 To LICHESEL.ListCount
     Q& = Q& + 1
     If LICHESEL.Selected(i& - 1) = True Then
       Z$ = LICHESEL.List(i& - 1)
       NUORD2& = XVALO(Mid$(Z$, 100, 5))
       'SUBO LA INFO A LISTA 1
       With CONCIMAN07.MSF2
          J& = CONCIMAN07.MSF1.Rows
          CONCIMAN07.MSF1.Rows = J& + 1
          CONCIMAN07.MSF1.TextMatrix(J&, 1) = CONCIMAN07.MSF3.TextMatrix(REG3&, 1)
          CONCIMAN07.MSF1.TextMatrix(J&, 2) = CONCIMAN07.MSF3.TextMatrix(REG3&, 2) 'COMPROBANTE QUE LO TIENE EN LA LISTA 3
          CONCIMAN07.MSF1.TextMatrix(J&, 3) = Mid$(Z$, 27, 8) 'FECHA DE ACREDITACION
          CONCIMAN07.MSF1.TextMatrix(J&, 4) = TRIM$(Mid$(Z$, 35, 32))
          CONCIMAN07.MSF1.TextMatrix(J&, 5) = FORMATNUM$(XVALO(Mid$(Z$, 67, 12)), "F12.2")
          CONCIMAN07.MSF1.TextMatrix(J&, 6) = FORMATNUM$(XVALO(Mid$(Z$, 79, 12)), "F12.2")
          CONCIMAN07.MSF1.TextMatrix(J&, 7) = XVALO(Mid$(Z$, 100, 5))
          CONCIMAN07.MSF1.TextMatrix(J&, 8) = TRIM$(Mid$(Z$, 105, 8))
          CONCIMAN07.MSF1.TextMatrix(J&, 9) = XVALO(Mid$(Z$, 113, 6))
          '
          Call CONCIMAN07.GRABA_COMPROBANTE_EN_CONCILIA(CONCIMAN07.MSF3.TextMatrix(REG3&, 2), CVINT(CONCIMAN07.MSF3.TextMatrix(REG3&, 3)), NUORD2&)
          Call CONCIMAN07.borrafila(CONCIMAN07.MSF2, Q&)
          Q& = Q& - 1
          '
       End With
     End If
    Next i&
    'AHORA BORRAR EL DE LA GRILLA 3
    Call CONCIMAN07.COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF3, &HCBEC86, REG3&, "*", 1)
    Call CONCIMAN07.borrafila(CONCIMAN07.MSF3, REG3&)
    Call CONCIMAN07.sortflex(CONCIMAN07.MSF1)
    '
    Call COMUNI("Conciliación Multiple Realizada")
    Unload Me
99  BOTREC.Enabled = True
    End Sub

Private Sub Command1_Click()
   Static ABUSQUE$
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", "Introduzca Argumento de Búsqueda", ABUSQUE$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     INI& = LICHESEL.ListIndex
10   For U& = INI& + 2 To LICHESEL.ListCount
       If InStr(Left$(UCase$(LICHESEL.List(U& - 1)), 100), ABUSQUE$) > 0 Then
         LICHESEL.ListIndex = U& - 1
         Exit Sub
       End If
     Next U&
     If INI& >= 0 Then
       INI& = -1
       GoTo 10
     End If
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If
End Sub



Private Sub Form_Load()

   Call APLICACIONSKINS(Me, Picture2)
   
   'LEO EL VALOR  DE LA FILA SELECCIONADO EN LA GRILLA 3
   For i& = 1 To CONCIMAN07.MSF3.Rows - 1
     If CONCIMAN07.MSF3.TextMatrix(i&, 0) = "*" Then
        REG3& = i&
        Exit For
     End If
   Next i&
   '
   With CONCIMAN07.MSF2
    .Redraw = False
    LICHESEL.Clear
    For i& = 1 To .Rows - 1
        FECHCO% = CVINT(.TextMatrix(i&, 1))
        COMPROB$ = TRIM$(.TextMatrix(i&, 2))
        FECHAC% = CVINT(.TextMatrix(i&, 3))
        REFE$ = TRIM$(.TextMatrix(i&, 4))
        DEBEX1# = XVALO(.TextMatrix(i&, 5))
        HABEX1# = XVALO(.TextMatrix(i&, 6))
        NUORD& = XVALO(.TextMatrix(i&, 7))
        TABLA$ = TRIM$(CONCIMAN07.MSF3.TextMatrix(REG3&, 8)) 'LA TABLA CORRESPONDIENTE DE PENDIENTES
        CTACBLEGASTOS% = XVALO(.TextMatrix(i&, 9))
        '
        Call REPLA(TX$, FECHATEX$(FECHCO%), 1, 8)
        Call REPLA(TX$, COMPROB$, 9, 18)
        Call REPLA(TX$, FECHATEX$(FECHAC%), 27, 8)
        Call REPLA(TX$, REFE$, 35, 32)
        Call REPLA(TX$, FORMATNUM$(DEBEX1#, "F12.2"), 67, 12)
        Call REPLA(TX$, FORMATNUM$(HABEX1#, "F12.2"), 79, 12)
        Call REPLA(TX$, TRIM$(Str$(NUORD&)), 100, 5)
        Call REPLA(TX$, TABLA$, 105, 8)
        Call REPLA(TX$, SAFETEXT$(CTACBLEGASTOS%), 113, 6)
        '
        LICHESEL.AddItem TX$
    Next i&
    .Redraw = True
   End With

End Sub

Private Sub LICHESEL_Click()
    SUCHESE# = 0
    For i& = 1 To LICHESEL.ListCount
       If LICHESEL.Selected(i& - 1) = True Then
         DEBEX# = DEBEX# + XVALO(Mid$(LICHESEL.List(i& - 1), 67, 12))
         HABEX# = HABEX# + XVALO(Mid$(LICHESEL.List(i& - 1), 79, 12))
       End If
    Next i&
    'ACUTALIZO LABEL
    Label2.Caption = FORMATNUM$(DEBEX# - HABEX#, "F12.2")
    Label9.Caption = FORMATNUM$(XVALO(Label8) - XVALO(Label2), "F12.2")
    Label2.Refresh
    Label9.Refresh
End Sub


