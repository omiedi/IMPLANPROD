VERSION 5.00
Begin VB.Form COLISIMT 
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "S.I.T. - Complementos, Consultas, Listados"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   720
      Left            =   5020
      Picture         =   "COLISIMT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1080
      Width           =   720
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
      Height          =   720
      Left            =   5020
      Picture         =   "COLISIMT.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cierra - Abandona la Aplicación"
      Top             =   315
      Width           =   720
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5020
      ScaleHeight     =   75
      ScaleWidth      =   660
      TabIndex        =   2
      Top             =   5280
      Width           =   720
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   3
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5835
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   4635
      Begin VB.Frame Frame2 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Exportación de Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   210
         TabIndex        =   9
         Top             =   3000
         Width           =   4215
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   19
            Top             =   945
            Width           =   2640
            Begin VB.CommandButton Command6 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "COLISIMT.frx":0454
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "Listado de Productos Exportados por Rango de Fechas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Detalle de Productos por Rango de Fechas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   21
               Top             =   45
               Width           =   1905
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   10
            Top             =   1785
            Width           =   2640
            Begin VB.CommandButton Command4 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "COLISIMT.frx":075E
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Listado de Componentes Tipificados Correspondientes a Exportaciones por R.Fechas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado Consumos Tipificados por Fecha"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   12
               Top             =   105
               Width           =   2010
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   13
            Top             =   420
            Width           =   2640
            Begin VB.CommandButton Command5 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "COLISIMT.frx":0A68
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Listado de Remitos de Exportación por Rango de Fechas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Detalle Remitos por Rango de Fechas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   14
               Top             =   45
               Width           =   1905
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Informes de Recepción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   210
         TabIndex        =   4
         Top             =   420
         Width           =   4215
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1080
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   25
            Top             =   1500
            Width           =   2640
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
               Height          =   600
               Left            =   0
               Picture         =   "COLISIMT.frx":0D72
               Style           =   1  'Graphical
               TabIndex        =   26
               ToolTipText     =   "Listado de IRMP's por Rango de Fechas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "I.R.M.P.'s por Codigo y Rango de Fechas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   27
               Top             =   100
               Width           =   1905
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   770
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   22
            Top             =   960
            Width           =   2640
            Begin VB.CommandButton Command2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "COLISIMT.frx":107C
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "Listado de IRMP's por Rango de Fechas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "IRMP's p/Proveedor  y Rango de Fechas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   720
               TabIndex        =   24
               Top             =   45
               Width           =   1905
            End
         End
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   720
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   7
            Top             =   3000
            Width           =   2640
            Begin VB.CommandButton Command9 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "COLISIMT.frx":1386
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Consulta por Número de IRMP"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Consulta por Número de I.R.Mat.Primas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   8
               Top             =   105
               Width           =   1905
            End
         End
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   480
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   5
            Top             =   420
            Width           =   2640
            Begin VB.CommandButton Command15 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "COLISIMT.frx":14D0
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "Listado de IRMP's por Rango de Fechas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado de I.R.M.P.'s por Rango de Fechas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   6
               Top             =   45
               Width           =   1905
            End
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4395
      Picture         =   "COLISIMT.frx":17DA
      Top             =   5520
      Width           =   2010
   End
End
Attribute VB_Name = "COLISIMT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call HELPONLINE("SIMPTEMP")
   Command11.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    '
5   Call SELECHO("PROVED1")
    PVXX% = Val(VALACT1$("PROVED1"))
    '
10  Call DESHASFECHA(DES%, HAS%, PERIO$)
    If DES% > 0 Then
      If HAS% >= DES% Then
        TTXX$ = "Puede Listar Solo Materias Primas\Solo Material de Reventa\o Todos los Materiales Ingresados"
        OPX% = COMALTER%(TTXX$ + "\\Mat.Primas\Reventa\Todo Junto")
        If OPX% < 1 Or OPX% > 3 Then GoTo 10
        '
        JJ& = 0
        REBLA$ = REGBLAN$("RECEPCO")
        Screen.MousePointer = 11
        If PVXX% > 0 Then
            RFF$ = RECFILT$("REMAPRI", 10, MKI$(PVXX%))
            FIN& = Len(RFF$)  ' ULTREG&("REMAPRI")
            For U& = 1 To FIN& Step 4
              Call TRACE(20, U&, FIN&)
              REMAP& = CVS(Mid$(RFF$, U&, 4))
              XX$ = REGLEIDO$("REMAPRI", REMAP&)
              GoSub 50
            Next U&
          Else
            FIN& = ULTREG&("REMAPRI")
            For U& = 1 To FIN& Step 4
              Call TRACE(20, U&, FIN&)
              XX$ = REGLEIDO$("REMAPRI", U&)
              GoSub 50
            Next U&
        End If
        '
        Call SETULTREG("RECEPCO", JJ&)
        Call CIERRARCH("RECEPCO")
        '
        If JJ& < 1 Then
          Call MENSERR(24, "Sin Movimientos en\Rango de Fechas Elegido.")
          GoTo 10
        End If
        '
        Call FILESORT("RECEPCO", "", 1, 4)
        '
        Call HEADERS("RECEPCO", "")
          PVYY$ = TRIM$(VALACT2$("PROVED1"))
          If PVXX% < 1 Then PVYY$ = "Todos"
        Call HEADERS("RECEPCO", "Proveedor: " + PVYY$ + " - " + PERIO$)
        Call CONECRUN("*LIREMP2", "Listado Recepciones por Rango Fechas")
        Screen.MousePointer = 1
        GoTo 10
        '
      End If
    End If
    GoTo 99
    '
50  FINGRE% = CVI(Mid$(XX$, 81, 2))
    If FINGRE% >= DES% Then
      If FINGRE% <= HAS% Then
        COMAT$ = TRIM$(UCase$(Mid$(XX$, 101, 2)))
        If OPX% = 3 Or (OPX% = 1 And COMAT$ = "MP") Or (OPX% = 2 And COMAT$ <> "MP") Then
          If Asc(Mid$(XX$, 98, 1)) <> 9 Then
              Call REPLA(XX$, Space$(8), 97, 8)
            Else
              Call REPLA(XX$, "Anulado", 97, 8)
          End If
          '
          YY$ = REBLA$
          Call REPLA(YY$, Mid$(XX$, 83, 6), 1, 6)
          Call REPLA(YY$, Mid$(XX$, 81, 2), 7, 2)
          Call REPLA(YY$, Mid$(XX$, 97, 8), 9, 8)
          Call REPLA(YY$, Mid$(XX$, 69, 12), 17, 12)
          Call REPLA(YY$, Mid$(XX$, 119, 16), 29, 16)
          Call REPLA(YY$, Mid$(XX$, 89, 8), 45, 8)
            If TRIM$(UCase$(Mid$(YY$, 9, 8))) = "ANULADO" Then
              Call REPLA(YY$, MKD$(0), 45, 8)
            End If
          Call REPLA(YY$, Mid$(XX$, 105, 2), 61, 2)
          Call REPLA(YY$, Mid$(XX$, 189, 4), 63, 4)
          Call REPLA(YY$, Mid$(XX$, 155, 8), 67, 8)
          Call REPLA(YY$, Mid$(XX$, 193, 44), 75, 44)
          JJ& = JJ& + 1
          Call GRAREG("RECEPCO", YY$, JJ&)
        End If
      End If
    End If
Return
   '
99 Command3.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call COPYSTRU("REMAPRI", "BOLREFIL")
    '
5   Call SELECHO("PROVED1")
    PVXX% = Val(VALACT1$("PROVED1"))
    If PVXX% < 1 Then GoTo 99
    '
10  Call DESHASFECHA(DES%, HAS%, PERIO$)
    If DES% > 0 Then
      If HAS% >= DES% Then
        TTXX$ = "Puede Listar Solo Materias Primas\Solo Material de Reventa\o Todos los Materiales Ingresados"
        OPX% = COMALTER%(TTXX$ + "\\Mat.Primas\Reventa\Todo Junto")
        If OPX% < 1 Or OPX% > 3 Then GoTo 10
        '
        Screen.MousePointer = 11
        RFF$ = RECFILT$("REMAPRI", 10, MKI$(PVXX%))
        JJ& = 0
        FIN& = Len(RFF$)  ' ULTREG&("REMAPRI")
        For U& = 1 To FIN& Step 4
          Call TRACE(20, U&, FIN&)
          REMAP& = CVS(Mid$(RFF$, U&, 4))
          XX$ = REGLEIDO$("REMAPRI", REMAP&)
          FINGRE% = CVI(Mid$(XX$, 81, 2))
          If FINGRE% >= DES% Then
            If FINGRE% <= HAS% Then
              COMAT$ = TRIM$(UCase$(Mid$(XX$, 101, 2)))
              If OPX% = 3 Or (OPX% = 1 And COMAT$ = "MP") Or (OPX% = 2 And COMAT$ <> "MP") Then
                If Asc(Mid$(XX$, 98, 1)) <> 9 Then
                     Call REPLA(XX$, Space$(8), 97, 8)
                   Else
                     Call REPLA(XX$, "Anulado", 97, 8)
                End If
                JJ& = JJ& + 1
                Call GRAREG("BOLREFIL", XX$, JJ&)
              End If
            End If
          End If
        Next U&
        Call SETULTREG("BOLREFIL", JJ&)
        Call CIERRARCH("BOLREFIL")
        '
        If JJ& < 1 Then
          Call MENSERR(24, "Sin Movimientos en\Rango de Fechas Elegido.")
          GoTo 10
        End If
        '
        Call HEADERS("BOLREFIL", "")
        Call HEADERS("BOLREFIL", "Proveedor: " + TRIM$(VALACT2$("PROVED1")) + " - " + PERIO$)
        Call CONECRUN("*LIREMP1", "Listado Recepciones por Rango Fechas")
        Screen.MousePointer = 1
        GoTo 10
        '
      End If
    End If
    GoTo 5
    '
99  Command2.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   MsgBox "Ultimo Cálculo: " + TRIM$(CONTROL$("CALCONIT", "FULTICAL"))
10 Call DESHASFECHA(DES%, HAS%, PERIO$)
   If DES% > 0 Then
     If HAS% >= DES% Then
       Screen.MousePointer = 11
       JJ& = 0: FIN& = ULTREG&("EXPOTIP")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("EXPOTIP", U&)
         FMXX% = CVI(Left$(XX$, 2))
         If FMXX% >= DES% Then
           If FMXX% <= HAS% Then
             JJ& = JJ& + 1
             Call GRAREG("EXPOTIPF", XX$, JJ&)
           End If
         End If
       Next U&
       Call SETULTREG("EXPOTIPF", JJ&)
       Call CIERRARCH("EXPOTIPF")
       '
       Call HEADERS("EXPOTIPF", "")
       Call HEADERS("EXPOTIPF", "Período: " + PERIO$)
       Call CONECRUN("*EXPOTIPF", "Listado Detalle de Consumos Re-Exportados")
       Screen.MousePointer = 1
       '
     End If
   End If
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   '
   Command5.Enabled = False
   MsgBox "Ultimo Cálculo: " + TRIM$(CONTROL$("CALCONIT", "FULTICAL"))
10 Call DESHASFECHA(DES%, HAS%, PERIO$)
   If DES% > 0 Then
     If HAS% >= DES% Then
       Screen.MousePointer = 11
       Call COPYSTRU("MOVISTO", "MOVSTEX1")
       JJ& = 0: FIN& = ULTREG&("MOVSTEX0")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("MOVSTEX0", U&)
         FMXX% = CVI(Left$(XX$, 2))
         If FMXX% >= DES% Then
           If FMXX% <= HAS% Then
             JJ& = JJ& + 1
             Call GRAREG("MOVSTEX1", XX$, JJ&)
           End If
         End If
       Next U&
       Call SETULTREG("MOVSTEX1", JJ&)
       Call CIERRARCH("MOVSTEX1")
       '
       Call HEADERS("MOVSTEX1", "")
       Call HEADERS("MOVSTEX1", "Período: " + PERIO$)
       Call CONECRUN("*LMOVEX1", "Listado Detalle de Exportación")
       Screen.MousePointer = 1
       '
     End If
   End If
   Command5.Enabled = True
   '
End Sub

Private Sub Command6_Click()
   '
   Command6.Enabled = False
   MsgBox "Ultimo Cálculo: " + TRIM$(CONTROL$("CALCONIT", "FULTICAL"))
10 Call DESHASFECHA(DES%, HAS%, PERIO$)
   If DES% > 0 Then
     If HAS% >= DES% Then
       Screen.MousePointer = 11
       JJ& = 0: FIN& = ULTREG&("MOVSTEX0")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("MOVSTEX0", U&)
         FMXX% = CVI(Left$(XX$, 2))
         If FMXX% >= DES% Then
           If FMXX% <= HAS% Then
             JJ& = JJ& + 1
             Call GRAREG("MOVSTEX1", XX$, JJ&)
           End If
         End If
       Next U&
       Call SETULTREG("MOVSTEX1", JJ&)
       Call CIERRARCH("MOVSTEX1")
       '
       Call FILESORT("MOVSTEX1", "", 2, 2)
       '
       Call HEADERS("MOVSTEX1", "")
       Call HEADERS("NOVSTEX1", "Período: " + PERIO$)
       Call CONECRUN("*LMOVEX2", "Detalle de Exportación por Producto")
       Screen.MousePointer = 1
       '
     End If
   End If
   Command6.Enabled = True
   '
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    Call OPNOIN("")
    Command9.Enabled = True
End Sub
'
Private Sub COMMAND15_CLICK()
    Command15.Enabled = False
    Call COPYSTRU("REMAPRI", "BOLREFIL")
    '
10  Call DESHASFECHA(DES%, HAS%, PERIO$)
    If DES% > 0 Then
      If HAS% >= DES% Then
        TTXX$ = "Puede Listar Solo Materias Primas\Solo Material de Reventa\o Todos los Materiales Ingresados"
        OPX% = COMALTER%(TTXX$ + "\\Mat.Primas\Reventa\Todo Junto")
        If OPX% < 1 Or OPX% > 3 Then GoTo 10
        '
        Screen.MousePointer = 11
        JJ& = 0: FIN& = ULTREG&("REMAPRI")
        For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          XX$ = REGLEIDO$("REMAPRI", U&)
          FINGRE% = CVI(Mid$(XX$, 81, 2))
          If FINGRE% >= DES% Then
            If FINGRE% <= HAS% Then
              COMAT$ = TRIM$(UCase$(Mid$(XX$, 101, 2)))
              If OPX% = 3 Or (OPX% = 1 And COMAT$ = "MP") Or (OPX% = 2 And COMAT$ <> "MP") Then
                If Asc(Mid$(XX$, 98, 1)) <> 9 Then
                     Call REPLA(XX$, Space$(8), 97, 8)
                   Else
                     Call REPLA(XX$, "Anulado", 97, 8)
                End If
                JJ& = JJ& + 1
                Call GRAREG("BOLREFIL", XX$, JJ&)
              End If
            End If
          End If
        Next U&
        Call SETULTREG("BOLREFIL", JJ&)
        Call CIERRARCH("BOLREFIL")
        '
        If JJ& < 1 Then
          Call MENSERR(24, "Sin Movimientos en\Rango de Fechas Elegido.")
          GoTo 10
        End If
        '98, 1
        Call CONECRUN("*LIREMP", "Listado Recepciones por Rango Fechas")
        Screen.MousePointer = 1
        '
      End If
    End If
    '
    Command15.Enabled = True
End Sub
'


