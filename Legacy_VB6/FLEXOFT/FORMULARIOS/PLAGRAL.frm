VERSION 5.00
Begin VB.Form PLAGRAL 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plan General de Cuentas"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8610
      Picture         =   "PLAGRAL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Imprime Listado de Cuentas Contables "
      Top             =   5040
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   8610
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   20
      Top             =   3885
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8610
      Picture         =   "PLAGRAL.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprime Listado del Plan de Cuentas"
      Top             =   2730
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Selección de Cuenta Activa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   210
      TabIndex        =   13
      Top             =   6090
      Width           =   9255
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
         Left            =   3780
         TabIndex        =   22
         Top             =   525
         Width           =   435
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   4515
         TabIndex        =   19
         Top             =   210
         Width           =   3375
         Begin VB.TextBox Text14 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   105
            MaxLength       =   10
            TabIndex        =   1
            Text            =   " "
            Top             =   315
            Width           =   3165
         End
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   7980
         Picture         =   "PLAGRAL.frx":0974
         TabIndex        =   5
         ToolTipText     =   "Eldición de la Cuenta Activa"
         Top             =   1115
         Width           =   1000
      End
      Begin VB.CommandButton Command31 
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
         Left            =   2835
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   525
         Width           =   175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   7980
         Picture         =   "PLAGRAL.frx":0C7E
         TabIndex        =   4
         ToolTipText     =   "Elimina la Cuenta Activa del Plan"
         Top             =   840
         Width           =   1000
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Insert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   7980
         Picture         =   "PLAGRAL.frx":0F88
         TabIndex        =   3
         ToolTipText     =   "Intercala Cuenta en el Plan "
         Top             =   575
         Width           =   1000
      End
      Begin VB.TextBox Text11 
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
         Left            =   1050
         TabIndex        =   0
         Top             =   525
         Width           =   1800
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Append"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   7980
         Picture         =   "PLAGRAL.frx":1292
         TabIndex        =   2
         ToolTipText     =   "Agrega Cuenta al Final del Plan"
         Top             =   300
         Width           =   1000
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel:"
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
         Left            =   2730
         TabIndex        =   23
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   18
         Top             =   1030
         Width           =   4005
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   0
         TabIndex        =   14
         Top             =   630
         Width           =   960
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5655
      Left            =   210
      TabIndex        =   11
      Top             =   210
      Width           =   8100
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   0
         TabIndex        =   12
         Top             =   315
         Width           =   8100
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   360
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   9000
         Begin VB.Label Label5 
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Nivel         Cuenta                    Referencia    "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   8625
         End
      End
      Begin VB.Line Line4 
         X1              =   7000
         X2              =   10
         Y1              =   0
         Y2              =   420
      End
   End
   Begin VB.CommandButton Command2 
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
      Height          =   800
      Left            =   8610
      Picture         =   "PLAGRAL.frx":159C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1890
      Width           =   855
   End
   Begin VB.CommandButton Command9 
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
      Height          =   800
      Left            =   8610
      Picture         =   "PLAGRAL.frx":18A6
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8610
      Picture         =   "PLAGRAL.frx":19F0
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Archivo Maestro de Cuentas Contables"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Height          =   800
      Left            =   8610
      Picture         =   "PLAGRAL.frx":1B3A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Aprueba y Graba el Plan de Cuentas"
      Top             =   1050
      Width           =   855
   End
   Begin VB.Line Line5 
      X1              =   8275
      X2              =   8280
      Y1              =   4305
      Y2              =   4725
   End
   Begin VB.Line Line3 
      X1              =   8275
      X2              =   8275
      Y1              =   1365
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   8265
      X2              =   8265
      Y1              =   1365
      Y2              =   1785
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   180
      X2              =   8275
      Y1              =   4725
      Y2              =   4725
   End
End
Attribute VB_Name = "PLAGRAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   If CONTROL$("PROPSCG", "PLAGRAL") = "GUEST" Then
      Call MENSERR(24, "Sistema Distribuido - Invitado en Plan de Cuentas\Imposible Grabar Plan de Cuentas.")
      Unload Me
      Exit Sub
   End If
   '
   Call GUARDAPLAN
   Call RECUECON
   '
End Sub
'
Private Sub Command11_Click()
   '
   NIVEI% = Val(Text1.TEXT)
   If NIVEI% < 0 Or NIVEI% > 4 Then Exit Sub
   '
   CUE% = CUEINT%(Text11.TEXT)
   If CUE% < 1 Or CUE% > NUCUEN% Then Exit Sub
   '
   ZZ$ = Space$(64)
   NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + ".....", 5)
   Call REPLA(ZZ$, NVX$, 1, 5)
   Call REPLA(ZZ$, CODCUE$(CUE%), 10, 12)
   Call REPLA(ZZ$, DENOCUE$(CUE%), 26, 32)
   List1.AddItem ZZ$
   '
   Text1.TEXT = 0
   Text11.TEXT = ""
   Label4.Caption = ""
   List1.ListIndex = List1.ListCount - 1
   '
End Sub

Private Sub Command12_Click()
   '
   NIVEI% = Val(Text1.TEXT)
   If NIVEI% < 0 Or NIVEI% > 4 Then Exit Sub
   '
   CUE% = CUEINT%(Text11.TEXT)
   If CUE% < 1 Or CUE% > NUCUEN% Then Exit Sub
   '
   U& = List1.ListIndex
   If U& < 0 Then Exit Sub
   If U& > List1.ListCount - 1 Then Exit Sub
   '
   ZZ$ = Space$(64)
   NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + ".....", 5)
   Call REPLA(ZZ$, NVX$, 1, 5)
   Call REPLA(ZZ$, CODCUE$(CUE%), 10, 12)
   Call REPLA(ZZ$, DENOCUE$(CUE%), 26, 32)
   List1.AddItem ZZ$, U&
   '
   Text1.TEXT = 0
   Text11.TEXT = ""
   Label4.Caption = ""
   List1.ListIndex = U&
   '
End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   If CONTROL$("PROPSCG", "PLAGRAL") = "GUEST" Then
      Call MENSERR(24, "Sistema Distribuido - Invitado en Plan de Cuentas\Imposible Editar B.Datos de Cuentas.")
      GoTo 99
   End If
   '
   Call GUARDAPLAN
   Call CARCUEINGAS
   Call RECARPLAN
99 Command3.Enabled = True
   '
End Sub

Private Sub Command4_Click()
   '
   U& = List1.ListIndex
   If U& < 0 Then Exit Sub
   If U& > List1.ListCount - 1 Then Exit Sub
   '
   List1.RemoveItem U&
   If U& < List1.ListCount Then
      List1.ListIndex = U&
   End If
   '
End Sub

Private Sub Command5_Click()
   '
   JJ& = 0
   FIN& = List1.ListCount
   For U& = 1 To List1.ListCount
      X$ = List1.List(U& - 1)
      JJ& = JJ& + 1
      Call GRAREG("PLANCUE", X$, JJ&)
   Next U&
   '
   Call SETULTREG("PLANCUE", JJ&)
   Call CIERRARCH("PLANCUE")
   Call FINAL("*PLANCUE")
   '
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call FINAL("*LICUECO")
   Command6.Enabled = True
End Sub

'
Private Sub Command8_Click()
   '
   U& = List1.ListIndex
   If U& < 0 Then Exit Sub
   If U& > List1.ListCount - 1 Then Exit Sub
   '
   X$ = List1.List(U&)
   NIVEX$ = Left$(X$, 5)
   While Left$(NIVEX$, 1) = "."
      NIVEX$ = Mid$(NIVEX$, 2)
   Wend
   NIVEI% = Val(NIVEX$)
   CUE% = CUEINT%(Mid$(X$, 10, 12))
   Text11.TEXT = CODCUE$(CUE%)
   Label4.Caption = DENOCUE$(CUE%)
   Text1.TEXT = TRIM$(Str$(NIVEI%))
   '
   List1.RemoveItem U&
   If U& < List1.ListCount Then
      List1.ListIndex = U&
   End If
   '
End Sub

'
Private Sub Form_Load()
    '
    Caption = Caption + " - " + TRIM$(EJACT$)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Refresh
    Call COPYSTRU("ECUECON", "ECUECOT")
    Call RECARPLAN
    '
End Sub
'

Private Sub Form_Unload(Cancel As Integer)
    TIASI% = 0
    Call CIERRACONTA
    Call LIBARCH("ASICONT")
End Sub

Private Sub List1_DblClick()
   '
   U& = List1.ListIndex
   If U& < 0 Then Exit Sub
   If U& > List1.ListCount - 1 Then Exit Sub
   '
   X$ = List1.List(U&)
   NIVEX$ = Left$(X$, 5)
   While Left$(NIVEX$, 1) = "."
      NIVEX$ = Mid$(NIVEX$, 2)
   Wend
   NIVEI% = Val(NIVEX$)
   CUE% = CUEINT%(Mid$(X$, 10, 12))
   Text11.TEXT = CODCUE$(CUE%)
   Label4.Caption = DENOCUE$(CUE%)
   Text1.TEXT = TRIM$(Str$(NIVEI%))
   '
   List1.RemoveItem U&
   If U& < List1.ListCount Then
      List1.ListIndex = U&
   End If
   '
End Sub

Private Sub Command2_Click()
    Call HELPONLINE("PLAGRAL")
End Sub
'

Private Sub Command9_Click()
    '
    FEX = FEPROAS
    List1.Clear
    Text11.TEXT = ""
    Label4.Caption = ""
    Text14.TEXT = ""
    Call CIERRARCH("*.*")
    Unload PLAGRAL
End Sub


Private Sub Text11_DBLCLICK()
   RG1$ = REGSEL$("ECUECOT")
   If Len(RG1$) > 4 Then
      Text11.TEXT = Left$(RG1$, 12)
      Label4.Caption = Mid$(RG1$, 13, 28)
   End If
End Sub
'
Private Sub COMMAND31_CLICK()
   RG1$ = REGSEL$("ECUECOT")
   If Len(RG1$) > 4 Then
      Text11.TEXT = Left$(RG1$, 12)
      Label4.Caption = Mid$(RG1$, 13, 28)
   End If
End Sub
'
Sub RECARPLAN()
    '
    Static YAVERI%
    '
    Screen.MousePointer = 11
    If UCase$(EXTE$) = "OLD" Then YAVERI% = 1
    If YAVERI% < 1 Then
       On Error Resume Next
       If FileLen(LUDAT$ + "PLAGRAL.DAT") >= FileLen(LUDAT$ + "PLAGRAL.OLD") Then YAVERI% = 1
       On Error GoTo 0
    End If
    '
    NXP% = FILEOPEN%(LUDAT$ + "PLAGRAL" + "." + EXTE$, 0, 4)
    Dim IVP As String * 4
    If YAVERI% < 1 Then
      Dim YAPLA%(32767)
      NXQ% = FILEOPEN%(LUDAT$ + "PLAGRAL" + ".OLD", 0, 4)
      Dim IVQ As String * 4
    End If
    '
    List1.Clear
    JJ& = 0
    '
    UCUEPLAN$ = "": UREPLA% = 0
    FIN& = LOF(NXP%) / 4
    For KK% = 1 To LOF(NXP%) / 4
       UU& = KK%
       Call TRACE(20, UU&, FIN&)
       Get #NXP%, KK%, IVP$
       CUE% = CVI(Mid$(IVP$, 1, 2))
       NIVEI% = CVI(Mid$(IVP$, 3, 2))
       If NIVEI% < 0 Then NIVEI% = 0
       If NIVEI% > 4 Then NIVEI% = 4
       '
       If CUE% > 0 Then
         ZZ$ = Space$(64)
         If NIVEI% >= 0 Then
           If NIVEI% <= 4 Then
             NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + ".....", 5)
             Call REPLA(ZZ$, NVX$, 1, 5)
             Call REPLA(ZZ$, CODCUE$(CUE%), 10, 12)
             Call REPLA(ZZ$, DENOCUE$(CUE%), 26, 32)
             List1.AddItem ZZ$
             UCUEPLAN$ = Mid$(ZZ$, 10, 12)
             If YAVERI% < 1 Then YAPLA%(CUE%) = 1
             UREPLA% = KK%
           End If
         End If
       End If
       '
    Next KK%
    '
    If YAVERI% < 1 Then
      FIN& = LOF(NXQ%) / 4
      For KK% = 1 To LOF(NXQ%) / 4
        UU& = KK%
        Call TRACE(20, UU&, FIN&)
        Get #NXQ%, KK%, IVQ$
        CUE% = CVI(Mid$(IVQ$, 1, 2))
        If CUE% > 0 Then
          If YAPLA%(CUE%) < 1 Then
            NIVEI% = CVI(Mid$(IVQ$, 3, 2))
            If NIVEI% < 0 Then NIVEI% = 0
            If NIVEI% > 4 Then NIVEI% = 4
            '
            COCUE$ = CODCUE$(CUE%)
            If COCUE$ > UCUEPLAN$ Then
              ZZ$ = Space$(64)
              If NIVEI% >= 0 Then
                If NIVEI% <= 4 Then
                  NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + ".....", 5)
                  Call REPLA(ZZ$, NVX$, 1, 5)
                  Call REPLA(ZZ$, CODCUE$(CUE%), 10, 12)
                  Call REPLA(ZZ$, DENOCUE$(CUE%), 26, 32)
                  List1.AddItem ZZ$
                  YAPLA%(CUE%) = 1
                End If
              End If
            End If
          End If
        End If
      Next KK%
      YAVERI% = 1
    End If
    '
    List1.Refresh
    Screen.MousePointer = 1
    '
End Sub
'
Sub GUARDAPLAN()
   '
   Screen.MousePointer = 11
   NXP% = FILEOPEN%(LUDAT$ + "PLAGRAL" + "." + EXTE$, 0, 4)
   Dim IVP As String * 4
   '
   JJ& = 0: FIN& = List1.ListCount
   For U& = 1 To List1.ListCount
      Call TRACE(20, U&, FIN&)
      X$ = List1.List(U& - 1)
      NIVEX$ = Left$(X$, 5)
      While Left$(NIVEX$, 1) = "."
         NIVEX$ = Mid$(NIVEX$, 2)
      Wend
      NIVEI% = Val(NIVEX$)
      CUE% = CUEINT%(Mid$(X$, 10, 12))
      IVP$ = String$(4, 0)
      Call REPLA(IVP$, MKI$(CUE%), 1, 2)
      Call REPLA(IVP$, MKI$(NIVEI%), 3, 2)
      JJ& = JJ& + 1
      Put #NXP%, JJ&, IVP$
   Next U&
   '
   For U& = JJ& + 1 To LOF(NXP%) / 4
      LSet IVP$ = String$(4, 0)
      Put #NXP%, U&, IVP$
   Next U&
   Close #NXP%
   Screen.MousePointer = 1
   '
End Sub

Sub RECUECON()
   '
   Screen.MousePointer = 11
   Dim FCUEN As String * 64
   Dim YAPLAN%(32767, 2)
   Dim IVP As String * 4
   NOHAYPLAN% = 0
   '
   EJERCI$ = EJACT$
   LU$ = LUDAT$
   ACUEN% = FILEOPEN%(LU$ + "CUENTAS." + EXTE$, 0, 64)
   For UI% = LOF(ACUEN%) / 64 To 1 Step -1
       Get #ACUEN%, UI%, FCUEN$
       If Asc(Left$(FCUEN$, 1)) > 0 Then
          NUCUEN% = UI%
          GoTo 5
       End If
   Next UI%
   NUCUEN% = 0
   '
5  Close #ACUEN%: ACUEN% = 0
   EXTEPLAN$ = "DAT": IPLAN% = 1: LPLAN$ = "Actual"
   '
6  NXP% = FILEOPEN%(LUDAT$ + "PLAGRAL." + EXTEPLAN$, 0, 4)
   If LOF(NXP%) < 4 Then NOHAYPLAN% = NOHAYPLAN% = NOHAYPLAN% + 1
   '
   FIN& = LOF(NXP%) / 4
   For U& = 1 To LOF(NXP%) / 4
       Call TRACE(20, U&, FIN&)
       Get #NXP%, U&, IVP$
       ICUE% = CVI(Left$(IVP$, 2))
       If ICUE% > 0 Then
          If ICUE% <= 32767 Then
             YAPLAN%(ICUE%, IPLAN%) = YAPLAN%(ICUE%, IPLAN%) + 1
          End If
       End If
   Next U&
   Close #NXP%: NXP% = 0
   '
   For KKI% = 1 To NUCUEN%
       If YAPLAN%(KKI%, IPLAN%) > 1 Then
          Call MENSERR(24, "Cuenta '" + TRIM$(CODCUE$(KKI%)) + "'\Duplicada en el Plan de Cuentas Ejercicio " + LPLAN$ + ".")
       End If
   Next KKI%
   '
   If EXTEPLAN$ = "DAT" Then
       EXTEPLAN$ = "OLD"
       IPLAN% = 2
       LPLAN$ = "Anterior"
       GoTo 6
   End If
   '
   K2& = 0
   FIN& = ULTREG&("ECUECOT")
   For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
      Z2$ = REGLEIDO$("ECUECOT", U&)
      ICUE% = CVI(Mid$(Z2$, 41, 2))
      If ICUE% > 0 Then
         If ICUE% <= NUCUEN% Then
            If YAPLAN%(ICUE%, 1) Or YAPLAN%(ICUE%, 2) > 0 Or NOHAYPLAN% = 2 Then
               K2& = K2& + 1
               Call GRAREG("ECUECON", Z2$, K2&)
            End If
         End If
      End If
   Next U&
   '
   Call SETULTREG("ECUECON", K2&)
   Call CIERRARCH("ECUECON")
   Call CIERRARCH("ECUECOT")
   Call FRESHECHO("ECUECON")
   Screen.MousePointer = 1
   '
End Sub
