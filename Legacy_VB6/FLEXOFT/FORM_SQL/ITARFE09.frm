VERSION 5.00
Begin VB.Form ITARFE09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Actividad por Fecha y Tecnico"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   11025
      Picture         =   "ITARFE09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Agrega Nuevo Registro"
      Top             =   6135
      Width           =   750
   End
   Begin VB.CommandButton Command2 
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
      Height          =   820
      Left            =   11025
      Picture         =   "ITARFE09.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6870
      Width           =   750
   End
   Begin VB.PictureBox Picture3 
      Height          =   6870
      Left            =   105
      ScaleHeight     =   6810
      ScaleWidth      =   10755
      TabIndex        =   12
      Top             =   945
      Width           =   10820
      Begin VB.VScrollBar VScroll1 
         Height          =   6825
         Left            =   10500
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   7230
         Left            =   0
         ScaleHeight     =   7230
         ScaleWidth      =   10455
         TabIndex        =   13
         Top             =   0
         Width           =   10450
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   6060
            TabIndex        =   18
            Top             =   180
            Width           =   645
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   7425
            TabIndex        =   17
            Top             =   180
            Width           =   645
         End
         Begin VB.CommandButton Command1 
            Caption         =   "."
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
            Index           =   0
            Left            =   1995
            TabIndex        =   16
            Top             =   180
            Width           =   180
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
            Height          =   540
            Index           =   0
            Left            =   5010
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   525
            Width           =   5265
         End
         Begin VB.Label Label6 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   495
            Index           =   0
            Left            =   285
            TabIndex        =   20
            Top             =   525
            Width           =   4425
         End
         Begin VB.Label Label17 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label17"
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
            Index           =   0
            Left            =   2400
            TabIndex        =   39
            Top             =   180
            Width           =   2055
         End
         Begin VB.Label Label11 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   0
            Left            =   3780
            TabIndex        =   32
            Top             =   690
            Width           =   645
         End
         Begin VB.Label Label10 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   0
            Left            =   3045
            TabIndex        =   31
            Top             =   690
            Width           =   645
         End
         Begin VB.Label Label9 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   0
            Left            =   2280
            TabIndex        =   30
            Top             =   690
            Width           =   645
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Index           =   0
            Left            =   8895
            TabIndex        =   23
            Top             =   180
            Width           =   660
         End
         Begin VB.Label Label8 
            Caption         =   "Lab."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   9870
            TabIndex        =   27
            ToolTipText     =   "Doble Click para Conmutar Sede de Actividad"
            Top             =   210
            Width           =   540
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   0
            X2              =   10500
            Y1              =   1185
            Y2              =   1185
         End
         Begin VB.Label Label2 
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
            Index           =   0
            Left            =   5430
            TabIndex        =   26
            Top             =   195
            Width           =   645
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta:"
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
            Index           =   0
            Left            =   6825
            TabIndex        =   25
            Top             =   195
            Width           =   540
         End
         Begin VB.Label Label4 
            Caption         =   "Horas:"
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
            Index           =   0
            Left            =   8295
            TabIndex        =   24
            Top             =   195
            Width           =   540
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   0
            Left            =   705
            TabIndex        =   22
            ToolTipText     =   "Doble Click para Limpiar Todo"
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "OT:"
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
            Index           =   0
            Left            =   -135
            TabIndex        =   21
            Top             =   225
            Width           =   750
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Detalle:"
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
            Index           =   0
            Left            =   4515
            TabIndex        =   19
            Top             =   315
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fecha"
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
      Left            =   9080
      TabIndex        =   8
      Top             =   105
      Width           =   1800
      Begin VB.CommandButton Command12 
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
         Left            =   1470
         TabIndex        =   10
         Top             =   315
         Width           =   180
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   420
         TabIndex        =   9
         Top             =   315
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   6150
      Left            =   11025
      ScaleHeight     =   6090
      ScaleWidth      =   900
      TabIndex        =   7
      Top             =   -105
      Width           =   960
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   90
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   35
         Top             =   5415
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command4 
         Height          =   645
         Left            =   105
         Picture         =   "ITARFE09.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Cierra y Abandona la Aplicacion"
         Top             =   210
         Width           =   645
      End
      Begin VB.CommandButton Command3 
         Height          =   645
         Left            =   105
         Picture         =   "ITARFE09.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Imprime Reporte"
         Top             =   4665
         Width           =   645
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "ITARFE09.frx":0C08
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5700
         Width           =   1515
      End
   End
   Begin VB.Frame Frame26 
      Caption         =   "Responsable"
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
      Left            =   2940
      TabIndex        =   0
      Top             =   105
      Width           =   6000
      Begin VB.OptionButton Option3 
         Caption         =   "Ezeq."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   3990
         TabIndex        =   1
         Top             =   315
         Width           =   795
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Matias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2880
         TabIndex        =   4
         Top             =   315
         Width           =   1170
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Omar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   1890
         TabIndex        =   2
         Top             =   315
         Width           =   960
      End
      Begin VB.OptionButton Option3 
         Caption         =   "EPB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Flavio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   945
         TabIndex        =   5
         Top             =   315
         Width           =   1275
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Natalia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   4935
         TabIndex        =   3
         Top             =   315
         Width           =   900
      End
   End
   Begin VB.Label Label18 
      Height          =   255
      Left            =   9720
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Horas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   38
      Top             =   7960
      Width           =   1575
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11880
      Y1              =   7810
      Y2              =   7810
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   8520
      TabIndex        =   37
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "I-Tar-Fe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   645
      Left            =   105
      TabIndex        =   11
      Top             =   210
      Width           =   2640
   End
End
Attribute VB_Name = "ITARFE09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RESPO$

Private Sub Command1_Click(Index As Integer)
    If FECHANUM(Label14) <= 33 Then
       Command12.SetFocus
       Exit Sub
    End If
    '
    Call COPYSTRU("CORTRAB1", "OTRASEL")
    ' CONDI$ = "(TecRespo = '" & RESPO$ & "' OR TecRespo = '' OR TecRespo IS NULL) "
    CONDI$ = "(STATUS < 3 OR STATUS IS NULL) "
    CONDI$ = CONDI$ & " AND (Fech_Cierre < '01/01/90' OR Fech_Cierre IS NULL) "
    CONDI$ = CONDI$ + " ORDER BY Nume_OT ASC"
    Call CARGALISEL("!OTRASEL", "CORTRAB", "Nume_OT/A12;Referencia/A66", CONDI$)
    Call SELECHO("!OTRASEL")
    AA1 = VALACT1$("!otrasel")
    NUCOTE$ = AJUSTI$(VALACT1$("!OTRASEL"), 12)
    If TRIM$(NUCOTE$) <> "" Then
       Label15(Index).Caption = NUCOTE$
    End If

End Sub
'
Private Sub Command12_Click()
    If Label18 = "CONTROL" Then Exit Sub ' para no poder cambiar fecha y salir
    '
    Call SELECHO("FECHA")
    VDEVU$ = VALACT1$("FECHA")
    If VDEVU$ <> "" Then Label14 = VDEVU$
End Sub

Sub CARGABLOIMG() 'CARGA BLOQUE DE IMAGEN
  '
  'Agrega un Conjunto de Imagen.
  'IMAGEN
  INTERVER = Line1(0).Y1
  NUPIC% = Text1.Count
  Load Line1(NUPIC%)
  Line1(NUPIC%).Y1 = Line1(NUPIC% - 1).Y1 + INTERVER
  Line1(NUPIC%).Y2 = Line1(NUPIC% - 1).Y2 + INTERVER
  Line1(NUPIC%).Visible = True
  If Line1(NUPIC%).Y1 > Picture4.Height - 500 Then
    Picture4.Height = Line1(NUPIC%).Y1 + 500
  End If
  VScroll1.Max = (Picture4.Height - Picture3.Height) / 10 + 1
  '
  Load Command1(NUPIC%): Command1(NUPIC%).Visible = True: Command1(NUPIC%).Top = Command1(NUPIC% - 1).Top + INTERVER
  '
  Load Text1(NUPIC%): Text1(NUPIC%).Visible = True: Text1(NUPIC%).Top = Text1(NUPIC% - 1).Top + INTERVER
  Load Text2(NUPIC%): Text2(NUPIC%).Visible = True: Text2(NUPIC%).Top = Text2(NUPIC% - 1).Top + INTERVER
  Load Text3(NUPIC%): Text3(NUPIC%).Visible = True: Text3(NUPIC%).Top = Text3(NUPIC% - 1).Top + INTERVER
  '
  Load Label2(NUPIC%): Label2(NUPIC%).Visible = True: Label2(NUPIC%).Top = Label2(NUPIC% - 1).Top + INTERVER
  Load Label3(NUPIC%): Label3(NUPIC%).Visible = True: Label3(NUPIC%).Top = Label3(NUPIC% - 1).Top + INTERVER
  Load Label4(NUPIC%): Label4(NUPIC%).Visible = True: Label4(NUPIC%).Top = Label4(NUPIC% - 1).Top + INTERVER
  Load Label5(NUPIC%): Label5(NUPIC%).Visible = True: Label5(NUPIC%).Top = Label5(NUPIC% - 1).Top + INTERVER
  Load Label6(NUPIC%): Label6(NUPIC%).Visible = True: Label6(NUPIC%).Top = Label6(NUPIC% - 1).Top + INTERVER
  Load Label7(NUPIC%): Label7(NUPIC%).Visible = True: Label7(NUPIC%).Top = Label7(NUPIC% - 1).Top + INTERVER
  Load Label8(NUPIC%): Label8(NUPIC%).Visible = True: Label8(NUPIC%).Top = Label8(NUPIC% - 1).Top + INTERVER
  Load Label9(NUPIC%): Label9(NUPIC%).Visible = True: Label9(NUPIC%).Top = Label9(NUPIC% - 1).Top + INTERVER
  Load Label10(NUPIC%): Label10(NUPIC%).Visible = True: Label10(NUPIC%).Top = Label10(NUPIC% - 1).Top + INTERVER
  Load Label11(NUPIC%): Label11(NUPIC%).Visible = True: Label11(NUPIC%).Top = Label11(NUPIC% - 1).Top + INTERVER
  Load Label13(NUPIC%): Label13(NUPIC%).Visible = True: Label13(NUPIC%).Top = Label13(NUPIC% - 1).Top + INTERVER
  Load Label15(NUPIC%): Label15(NUPIC%).Visible = True: Label15(NUPIC%).Top = Label15(NUPIC% - 1).Top + INTERVER
  Load Label17(NUPIC%): Label17(NUPIC%).Visible = True: Label17(NUPIC%).Top = Label17(NUPIC% - 1).Top + INTERVER
  '
  JJ& = NUPIC% + 1
  Call LIMPIATEX(JJ&)
  Call FRESHALL
  '
End Sub

Private Sub Command2_Click()
    If FECHANUM(Label14) <= 33 Then
       Command12.SetFocus
       Exit Sub
    End If
    '
    FEJEC% = FECHANUM(Label14)
    If FEJEC% < 1 Then Exit Sub
    '
    Command2.Enabled = False
    CONDI$ = "FechEjec = '" & FETEXCOR1$(FEJEC%) & "' "
    CONDI$ = CONDI$ + " AND RespEjec = '" & RESPO$ & "' "
    Call ACTUREGISTRO("DETOTRA", "StatActi", CONDI$, 9, RAFE&, "NOMESS")
    '
    SQLX$ = "SELECT * FROM DETOTRA "
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    '
    Dim OTRASE As ADODB.Recordset
    Set OTRASE = New ADODB.Recordset
    OTRASE.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
    '
    SUMHOR = 0
    For JJ& = 1 To Label15.Count
      NUCOTE$ = TRIM$(Label15(JJ& - 1))
      If NUCOTE$ <> "" Then
        ' ESTO ES PARA ELIMINAR REGISTROS DUPLICADOS
        If TRIM$(Text3(JJ& - 1)) <> TRIM$(REFX$) Then GoTo 25
        If Int(HORANUM(Text1(JJ& - 1))) <> COMIX% Then GoTo 25
        If Int(HORANUM(Text2(JJ& - 1))) <> FINIX% Then GoTo 25
        If TRIM$(UCase$(Label8(JJ& - 1))) <> LUEJ$ Then GoTo 25
        If TRIM$(UCase$(Label9(JJ& - 1))) <> CONCE$ Then GoTo 25
        GoTo 39
        '
25      FEXI% = FEJEC%
        REFX$ = Text3(JJ& - 1)
        COMIX% = Int(HORANUM(Text1(JJ& - 1)))
        FINIX% = Int(HORANUM(Text2(JJ& - 1)))
        HORAX = DIFHORA(COMIX%, FINIX%)
          SUMHOR = SUMHOR + HORAX
        LUEJ$ = TRIM$(UCase$(Label8(JJ& - 1)))
        CONCE$ = TRIM$(UCase$(Label9(JJ& - 1)))
        TARIF = XVALO(Label10(JJ& - 1))
        IMPOR = XVALO(Label11(JJ& - 1))
        '
        With OTRASE
          .AddNew
          !nume_ot = NUCOTE$
          !FechEjec = CVDAT(FEXI%)
          !Descrip = SAFETEXT$(REFX$ + " ")
          !HORACOM = COMIX%
          !HORAFIN = FINIX%
          !TiemEjec = HORAX
          !RespEjec = RESPO$
          !LugEjec = LUEJ$
          !Concepto = CONCE$
          !TarifHor = TARIF
          !Importe = IMPOR
          .Update
        End With
      End If
39  Next JJ&
    OTRASE.Close
    Set OTRASE = Nothing
    '
    Label12 = FORMATNUM$(SUMHOR, "F10.2")
    '
    CONDI$ = CONDI$ & " AND StatActi = 9 "
    Call BORRAREGISTROS("DETOTRA", CONDI$, RAFE, "NOMESS")
    '
    Command2.Enabled = True
    '
End Sub

Private Sub Command3_Click()
   If FECHANUM(Label14) <= 33 Then
      Command12.SetFocus
      Exit Sub
   End If
   '
   Command3.Enabled = False
   If TRIM$(Label14) <> "" Then
     PARAM$ = "(" + Label14 + ";" + RESPO$ + ")"
     Call CONECRUN("?ACTIPORFETEC" + PARAM$, "Actividad por Fecha y Tecnico")
   End If
   '
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Unload Me
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Call CARGABLOIMG
   Command5.Enabled = True
End Sub

Private Sub Form_Load()
   RESPO$ = "EPB"
   HPII% = HOY(HOS$)
   Label14 = HOS$
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   '
   Cancel = 0
   SUMHOR = 0
   For JJ& = 1 To Label5.Count
      SUMHOR = SUMHOR + XVALO(Label5(JJ& - 1))
   Next JJ&
   If Abs(XVALO(Label12) - SUMHOR) > 0.05 Then
      Call MENSERR(24, "Debe Grabar Antes de Salir (Dif.= " & TRIM$(FORMATNUM$(Abs(XVALO(Label12) - SUMHOR), "F12.4")) & ")")
      Cancel = 1
      Visible = 1
      Exit Sub
   End If
   '
   If Label18 = "CONTROL" Then
     CANOR = XVALO(Label12)
     If CANOR < 7 Then
       Call MENSERR(24, "Parte de Horas del " + Label14 + " Incompleto !!!")
       Cancel = 1
       Visible = 1
     End If
   End If
   '
End Sub

Private Sub Label14_Change()
   Call CARLISTAPEN
End Sub

Sub CARLISTAPEN()
   '
   Label12 = "": SUMHOR = 0
   For JJ& = 1 To Line1.Count
      Call LIMPIATEX(JJ&)
   Next JJ&
   '
   FEJEC% = FECHANUM(Label14)
   If FEJEC% < 1 Then Exit Sub
   '
   SQLX$ = "SELECT * FROM DETOTRA INNER JOIN CORTRAB "
   SQLX$ = SQLX$ + " ON DETOTRA.Nume_OT = CORTRAB.Nume_OT "
   SQLX$ = SQLX$ + " WHERE DETOTRA.FechEjec = '" & FETEXCOR1$(FEJEC%) & "' "
   SQLX$ = SQLX$ + " AND DETOTRA.RespEjec = '" & RESPO$ & "' "
   SQLX$ = SQLX$ + " ORDER BY DETOTRA.HORACOM ASC "
   '
   Dim OTRASE As ADODB.Recordset
   Set OTRASE = New ADODB.Recordset
   Set OTRASE = FLEXCONN.Execute(SQLX$)
   '
   JJ& = 0
   With OTRASE
     Do While Not .EOF
       JJ& = JJ& + 1
       If JJ& > Line1.Count Then
         Call CARGABLOIMG
       End If
       Label15(JJ& - 1) = OTRASE!nume_ot
       Text1(JJ& - 1) = FORMATNUM$(XVALO(OTRASE!HORACOM), "H5")
       Text2(JJ& - 1) = FORMATNUM$(XVALO(OTRASE!HORAFIN), "H5")
       Label5(JJ& - 1) = FORMATNUM$(DIFHORA(OTRASE!HORACOM, OTRASE!HORAFIN), "F6.2")
          SUMHOR = SUMHOR + XVALO(Label5(JJ& - 1))
       Text3(JJ& - 1) = SAFETEXT$(OTRASE!Descrip)
       Label8(JJ& - 1).Caption = "Lab."
          If UCase$(Left$(SAFETEXT$(OTRASE!LugEjec), 3)) = "COM" Then Label8(JJ& - 1).Caption = "Comp"
       Label9(JJ& - 1).Caption = SAFETEXT$(OTRASE!Concepto)
       Label10(JJ& - 1).Caption = FORMATNUM$(XVALO(OTRASE!TarifHor), "F12.2")
       Label11(JJ& - 1).Caption = FORMATNUM$(XVALO(OTRASE!Importe), "F12.2")
     .MoveNext
     Loop
   End With
   OTRASE.Close
   Set OTRASE = Nothing
   '
   Label12 = FORMATNUM$(SUMHOR, "F10.2")
   '
End Sub

Private Sub Label15_Change(Index As Integer)
    Label6(Index) = ""
    Label6(Index) = TRIM$(VALOBADA("CORTRAB", "Descripcion", "Nume_Ot = '" & TRIM$(Label15(Index)) & "'", "NOMESS"))
      NCLIE% = XVALO(VALOBADA("CORTRAB", "NroClie", "Nume_Ot = '" & TRIM$(Label15(Index)) & "'", "NOMESS"))
    Label17(Index).Caption = FANTCLI$(NCLIE%)
End Sub

Private Sub Label15_DblClick(Index As Integer)
    If COMALTER%("Confirma Eliminar la Tarea Activa ?\\No, Conservarla\Si, Borrarla") = 2 Then
       JJ& = Index + 1
       Call LIMPIATEX(JJ&)
       Call Command2_Click
       Call CARLISTAPEN
    End If
End Sub

Private Sub Label6_DblClick(Index As Integer)
   If Label18 = "CONTROL" Then Exit Sub ' para no poder salir por la O/T
   '
   NOTRA$ = TRIM$(Label15(Index))
   If NOTRA$ <> "" Then
      FADFLESQL.Label15 = NOTRA$
      Hide
   End If
End Sub

Private Sub Label8_DblClick(Index As Integer)
   If Left$(UCase$(Label8(Index)), 3) = "LAB" Then
        Label8(Index) = "Comp"
      Else
        Label8(Index) = "Lab."
    End If
End Sub

Private Sub Option3_Click(Index As Integer)
   '
   If Label18 = "CONTROL" Then Exit Sub ' para no poder cambiar responsable y salir
   '
   RESPO$ = "EPB"
   If Option3(1) = True Then RESPO$ = "FG"
   If Option3(5) = True Then RESPO$ = "OD"
   If Option3(2) = True Then RESPO$ = "MC"
   If Option3(4) = True Then RESPO$ = "EG"
   If Option3(3) = True Then RESPO$ = "NC"
   '
   Call SOLOPROPIO
   '
   Call CARLISTAPEN
   '
End Sub

Sub SOLOPROPIO()
   ' LOS TECNICOS SOLO AJUSTAN SU PROPIO PARTE
   If USERID$ = "MC" Then Option3(2) = True
   If USERID$ = "EG" Then Option3(4) = True
   If USERID$ = "OD" Then Option3(5) = True
End Sub
'
Private Sub Text1_GotFocus(Index As Integer)
    If FECHANUM(Label14) <= 33 Then
       Command12.SetFocus
       Exit Sub
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text2(Index).SetFocus
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    COMIX% = HORANUM(Text1(Index).TEXT)
    Text1(Index) = FORMATNUM$(COMIX%, "H5")
    Call CALDIFHOR(Index)
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    If FECHANUM(Label14) <= 33 Then
       Command12.SetFocus
       Exit Sub
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text3(Index).SetFocus
    End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    FINIX% = HORANUM(Text2(Index).TEXT)
    Text2(Index) = FORMATNUM$(FINIX%, "H5")
    Call CALDIFHOR(Index)
End Sub

Sub CALDIFHOR(Index)
    COMIX% = HORANUM(Text1(Index).TEXT)
    FINIX% = HORANUM(Text2(Index).TEXT)
    Label5(Index) = "0.00"
    If FINIX% > COMIX% Then
       Label5(Index) = FORMATNUM(DIFHORA(COMIX%, FINIX%), "F6.2")
    End If
    '
End Sub
'
Private Sub Text3_GotFocus(Index As Integer)
    If FECHANUM(Label14) <= 33 Then
       Command12.SetFocus
       Exit Sub
    End If
End Sub

Private Sub VScroll1_Change()
   On Error Resume Next 'POR ERROR '6' DESBORDAMIENTO
     Picture4.Top = (-10) * VScroll1.Value
   On Error GoTo 0
End Sub

Private Sub VScroll1_Scroll()
  On Error Resume Next 'POR ERROR '6' DESBORDAMIENTO
   Picture4.Top = (-10) * VScroll1.Value
  On Error GoTo 0
End Sub

Sub LIMPIATEX(JJ&)
    Text1(JJ& - 1) = ""
    Text2(JJ& - 1) = ""
    Text3(JJ& - 1) = ""
    Label5(JJ& - 1) = ""
    Label6(JJ& - 1) = ""
    Label8(JJ& - 1) = "Lab."
    Label9(JJ& - 1).Caption = ""
    Label10(JJ& - 1).Caption = ""
    Label11(JJ& - 1).Caption = ""
    Label17(JJ& - 1).Caption = ""
    Label15(JJ& - 1) = ""
End Sub
