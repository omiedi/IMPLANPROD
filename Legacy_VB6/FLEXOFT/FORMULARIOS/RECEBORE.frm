VERSION 5.00
Begin VB.Form RECEBORE 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción Bombas para Evaluar Reparación"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   10860
   Icon            =   "RECEBORE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDC0DD&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12000
      Begin VB.Frame Frame4 
         BackColor       =   &H00DDC0DD&
         Caption         =   "Orden de Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1020
         Left            =   7980
         TabIndex        =   30
         Top             =   3675
         Width           =   2745
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   210
            TabIndex        =   31
            Text            =   " "
            Top             =   495
            Width           =   2325
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDC0DD&
         Caption         =   "Remito Número"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1020
         Left            =   7980
         TabIndex        =   28
         Top             =   2625
         Width           =   2745
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
            Height          =   300
            Left            =   210
            TabIndex        =   29
            Text            =   " "
            Top             =   495
            Width           =   2325
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1050
         TabIndex        =   27
         Top             =   1095
         Width           =   195
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
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
         Left            =   8925
         Picture         =   "RECEBORE.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   4830
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
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
         Left            =   9870
         Picture         =   "RECEBORE.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   4830
         Width           =   855
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9255
         TabIndex        =   18
         Text            =   " "
         Top             =   1095
         Width           =   1485
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DDC0DD&
         Caption         =   "Serie Número"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1020
         Left            =   7980
         TabIndex        =   15
         Top             =   1575
         Width           =   2745
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
            Height          =   300
            Left            =   210
            TabIndex        =   0
            Text            =   " "
            Top             =   495
            Width           =   2325
         End
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
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
         Height          =   720
         Left            =   7980
         Picture         =   "RECEBORE.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba, Graba e Imprime Presupuesto"
         Top             =   4830
         Width           =   855
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00DDC0DD&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   3960
         Left            =   210
         TabIndex        =   8
         Top             =   1575
         Width           =   7680
         Begin VB.TextBox Text28 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3225
            Left            =   315
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Top             =   525
            Width           =   7155
         End
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9615
         TabIndex        =   2
         Text            =   " "
         Top             =   315
         Width           =   960
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10560
         TabIndex        =   3
         Top             =   315
         Width           =   180
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   5220
         TabIndex        =   23
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   7815
         TabIndex        =   22
         Top             =   1155
         Width           =   1380
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   5220
         TabIndex        =   21
         Top             =   735
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   5745
         TabIndex        =   20
         Top             =   315
         Width           =   2220
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   5745
         TabIndex        =   19
         Top             =   705
         Width           =   2220
      End
      Begin VB.Label REPRESUP 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   225
         Left            =   6510
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   4935
         TabIndex        =   16
         Top             =   1785
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9615
         TabIndex        =   14
         Top             =   705
         Width           =   1110
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   2205
         TabIndex        =   13
         Top             =   1095
         Width           =   2850
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   1365
         TabIndex        =   12
         Top             =   1095
         Width           =   750
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   1365
         TabIndex        =   11
         Top             =   705
         Width           =   3690
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   1365
         TabIndex        =   10
         Top             =   315
         Width           =   3690
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   315
         TabIndex        =   9
         Top             =   1095
         Width           =   750
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   420
         Picture         =   "RECEBORE.frx":0DC8
         ToolTipText     =   "Click Aquí para cambiar Modalidad de Facturación"
         Top             =   105
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Presup.Nro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   7770
         TabIndex        =   7
         Top             =   735
         Width           =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   8820
         TabIndex        =   6
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   105
         TabIndex        =   5
         Top             =   735
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
End
Attribute VB_Name = "RECEBORE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   Call SELECHO("DEUDOR1")
   NCLIE% = Val(VALACT1$("DEUDOR1"))
   If NCLIE% > 0 Then
      Label11 = TRIM$(Str$(NCLIE%))
   End If
   Command1.Enabled = False
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call GRAPRESU(OKX%)
   If OKX% > 0 Then
     CANCELPRINT% = 0
     Call PRIPRESU
     Call BLANFORMU
     Label18 = ""
     Command14.Enabled = True
     Label11 = ""
     Label18 = TRIM$(Str$(NUPRESUP&))
     Label10 = ""
     Hide
   End If
   Command14.Enabled = True
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Unload Me
    Hide
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PRESUP01")
    Command3.Enabled = True
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   Text9.TEXT = VALACT1$("SELFECHA")
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    '
    CUHOR = Val(CONTROL$("REPARA", "VALHORA"))
    COEFI = Val(CONTROL$("REPARA", "COEFPREA"))
    Label20 = CSTRING$(MKS$(CUHOR * COEFI), 4, 8, 2, 2)
    '
    Label18 = TRIM$(Str$(NUPRESUP&))
    '
End Sub
'

Private Sub Form_Resize()
   If WindowState = 1 Then
      Call CIERRARCH("*.*")
      Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Label11 = ""
   Label18 = "": Label10 = ""
   Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     openforms = DoEvents
   Wend
End Sub

Private Sub Label11_Change()
   '
   NC% = Val(Label11)
   Label12 = ""
   Label13 = ""
   Label14 = ""
   Label15 = ""
   Label16 = ""
   Label17 = ""
   '
   If NC% > 0 Then
     If RASOCLI$(NC%) <> "" Then
       Label12 = RASOCLI$(NC%): Label12.Refresh
       Label13 = DOMICLI$(NC%)
       Label14 = CPOSCLI$(NC%)
       Label15 = LOCACLI$(NC%)
       Label16 = TEL1CLI$(NC%)
       Label17 = FAXCLI$(NC%)
       DoEvents
     End If
   End If
   '
End Sub

Private Sub Label18_Change()
   '
   NPX& = Val(Label18)
   Call BLANFORMU
   For UI& = ULTREG&("PRESREPA") To 1 Step -1
     XX$ = REGLEIDO$("PRESREPA", UI&)
     If CVS(Left$(XX$, 4)) = NPX& Then
       REPRESUP = TRIM$(Str$(UI&))
       NCLIE% = CVI(Mid$(XX$, 67, 2))
       Label11 = TRIM$(Str$(NCLIE%))
       CII% = CVI(Mid$(XX$, 69, 2))
       Label10 = CODEXT$(CII%)
       TEPRUEBA = DESCRIT$(CII%)
       Text1 = Mid$(XX$, 71, 15)
       Text10 = Mid$(XX$, 86, 9)
       TALON = Mid$(XX$, 95, 6)
       Text28 = TRIM$(Mid$(XX$, 129, 368))
       Label19 = CSTRING$(Mid$(XX$, 101, 4), 4, 8, 2, 2)
       If CVS(Mid$(XX$, 105, 4)) > 0.5 Then
         Label20 = CSTRING$(Mid$(XX$, 105, 4), 4, 8, 2, 2)
       End If
       '
       Call CAREPUES
       Call PREREPUES
       '
       On Error Resume Next
       Text10.SetFocus
       On Error GoTo 0
       openforms = DoEvents
       Exit Sub
     End If
   Next UI&
   '
End Sub


Private Sub TEPRUEBA_Change()
   'DESCRI$ = TEPRUEBA.Caption
   'If TEPRUEBA.Width < Label8.Width Then
   '  DESCRI$ = Chr$(10) + Chr$(13) + DESCRI$
   'End If
   'Label8.Caption = DESCRI$
End Sub

Function VALICLIE%()
   VCC% = 0
   If Val(Label11) > 0 Then
     If Val(Label11) <= 32767 Then
       NCLIE% = Val(Label11)
       If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function
'
Sub CAREPUES()
   '
   Call SETULTREG("!MATEROF", 0)
   NPX& = Val(Label18)
   NPI% = NPX& Mod 10000
   RFF$ = RECFILT$("REPUREPA", 2, MKI$(NPI%))
   '
   URE& = ULTREG&("REPUREPA")
30 For KK& = 1 To Len(RFF$) Step 4
     REG& = CVS(Mid$(RFF$, KK&, 4))
     If REG& > 0 Then
       If REG& <= URE& Then
         XX$ = REGLEIDO$("REPUREPA", REG&)
         If CVS(Left$(XX$, 4)) = NPX& Then
           CII% = CVI(Mid$(XX$, 7, 2))
           CANTI = CVS(Mid$(XX$, 9, 4))
           '
           ZZ$ = REGBLAN$("MATEROF")
           Call REPLA(ZZ$, MKI$(CII%), 1, 2)
           Call REPLA(ZZ$, UNIMED$(CII%), 3, 2)
           Call REPLA(ZZ$, MKS$(CANTI), 5, 4)
           Call REPLA(ZZ$, MKD$(CDbl(CANTI)), 9, 8)
           Call REPLA(ZZ$, MKS$(0), 17, 4)
           Call REGAPP("!MATEROF", ZZ$)
           '
         End If
       End If
     End If
   Next KK&
   Call CIERRARCH("!MATEROF")
   '
End Sub
'
Sub PREREPUES()
   '
   List1.Clear
   SUMA# = 0
   SUMAREP = ""
   COEFI = Val(CONTROL$("REPARA", "COEFPREA"))
   For U& = 1 To ULTREG&("!MATEROF")
     ZZ$ = REGLEIDO$("!MATEROF", U&)
     TTXX$ = Space$(128)
     CII% = CVI(Left$(ZZ$, 2))
     Call REPLA(TTXX$, CODEXT$(CII%), 1, 16)
     Call REPLA(TTXX$, DESCRIT$(CII%), 17, 38)
     NEBRUX$ = Mid$(ZZ$, 5, 4)
     Call REPLA(TTXX$, CSTRING$(NEBRUX$, 4, 10, 1, 2), 54, 10)
     COSUX$ = MKS$(COEFI * COSUNI(CII%))
     Call REPLA(TTXX$, CSTRING$(COSUX$, 4, 10, 2, 2), 64, 10)
     ITOX$ = MKS$(CVS(NEBRUX$) * CVS(COSUX$))
     Call REPLA(TTXX$, CSTRING$(ITOX$, 4, 12, 2, 2), 74, 12)
     SUMA# = SUMA# + Val(Mid$(TTXX$, 74, 12))
     List1.AddItem TTXX$
   Next U&
   SUMAREP = CSTRING$(MKD$(SUMA#), 4, 11, 2, 2)
   '
End Sub
'
Sub BLANFORMU()
   '
   Label8 = ""
   Label10 = ""
   Label12 = ""
   Label13 = ""
   Label14 = ""
   Label15 = ""
   Label16 = ""
   Label17 = ""
   Text1 = ""
   Text10 = ""
   TALON = ""
   TEPRUEBA = ""
   Text28 = ""
   Label19 = ""
   Label20 = ""
   '
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   '
End Sub
'

Private Sub Text9_GotFocus()
   On Error Resume Next
   Command9.SetFocus
   On Error GoTo 0
End Sub


Sub PRIPRESU()
    '
    FORIPRE$ = TRIM$(CONTROL$("FOREPAR", "DOCUREP5"))
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Imposible Imprimir Recepción de Reparación.\Falta Definir Formulario Electrónico.")
       Exit Sub
    End If
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Recepción de Reparación.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
10  NC% = Val(Label11)
    TIPODOC$ = "Recepción de Reparación"
    DESTIDOC% = NC%
    NUMEDOC$ = TRIM$(Label18)
    FECHDOC$ = TRIM$(Text9.TEXT)
    '
    CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = DEPOSAL$
    CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = Text2
    CODPARAM$(3) = "N-AUTENT": DOCPARAM$(3) = Text3
    CAPARDOC% = 3
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "COLOR"
    CODTABU1$(4) = "UNIMED"
    CODTABU1$(5) = "NRO-LOTE"
    CODTABU1$(6) = "CANTIDAD"
    CODTABU1$(7) = "N-OCOMPR"
    CODTABU1$(8) = "CAKILOS"
    CODTABU1$(9) = "F-PLANO"
    CODTABU1$(10) = "FECH-REC"
    CACOLTAB1% = 10
    '
    CAITAB1% = 1
    TETABU1$(2, CAITAB1%) = "S/Nro." + TRIM$(Text1) + " - Ref." + TRIM$(Text10)
    TETABU1$(6, CAITAB1%) = "1"
    '
    Call FRATEXTO(Text28, 46)
    For KKI% = 1 To CARETEX%
      CAITAB1% = CAITAB1% + 1
      TETABU1$(2, CAITAB1%) = RETEX$(KKI%)
    Next KKI%                           '
    '
    '
    Call PRINTDOC("@" + FORIPRE$)
    '
End Sub

Sub GRAPRESU(OKX%)
   '
   OKX% = 0
   Call BLOQARCH("PRESREPA")
   '
   REG& = Val(REPRESUP)
   If REG& < 1 Then REG& = 1 + ULTREG&("PRESREPA")
   If REG& < 1 Or REG& > 1 + ULTREG&("PRESREPA") Then
      Call COMUNI("Imposible Registrar - Error de Indices")
      GoTo 98
   End If
   '

   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = Val(Label11)
   '
   If TRIM$(Text1) = "" Then
      MERRO$ = "Falta Número de Serie\o Identificación del Equipo."
      GoTo 98
   End If
   '
   If Len(TRIM$(Text1)) > 15 Then
        MERRO$ = "Exceso Longitud Número de Serie"
        Text1 = Left$(TRIM$(Text1), 15)
        GoTo 98
   End If
   '
   If Len(TRIM$(Text10)) > 9 Then
        MERRO$ = "Exceso Longitud Referencia"
        Text10 = Left$(TRIM$(Text10), 9)
        GoTo 98
   End If
   '
   If TRIM$(Text10) = "" Then
      MERRO$ = "Falta Referencia del Cliente\o Número de Pedido de Cotización."
      GoTo 98
   End If
   '
   HOII% = HOY(HOS$)
   FEMI% = FECHANUM(Text9.TEXT)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
   XX$ = REGBLAN$("PRESREPA")
   REFEPRE$ = "S/Nro." + Text1 + " - " + Left$(TRIM$(Label12), 30) + " - R: " + TRIM$(Text10) + " - " + TRIM$(Text9)
   NPLAL& = Val(Label18)
   Call REPLA(XX$, MKS$(NPLAL&), 1, 4)
   Call REPLA(XX$, REFEPRE$, 5, 60)
   Call REPLA(XX$, MKI$(NCLIE%), 67, 2)
   Call REPLA(XX$, TRIM$(Text1), 71, 15)
   Call REPLA(XX$, TRIM$(Text10), 86, 9)
   Call REPLA(XX$, TRIM$(TALONX$), 95, 6)
   Call REPLA(XX$, MKS$(0), 101, 4)
   Call REPLA(XX$, MKS$(0), 105, 4)
   Call REPLA(XX$, MKS$(0), 109, 4)
   Call REPLA(XX$, MKS$(0), 113, 4)
   Call REPLA(XX$, MKS$(0), 117, 4)
   Call REPLA(XX$, MKI$(0), 121, 2)   ' NO SE EMITIO PRESUPUESTO
   Call REPLA(XX$, Text28 + Space$(368), 129, 368)
   Call REPLA(XX$, "", 497, 16)
   Call GRAREG("PRESREPA", XX$, REG&)
   Call CIERRARCH("PRESREPA")
   '
   NPLAL& = Val(Label18)
   NPLAI% = NPLAL& Mod 10000
   '
   TTXYZ$ = Space$(92)
   NROREP& = Val(Label18)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   Call REPLA(TTXYZ$, "Recepción Material para Presupuestación de Reparación:", 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   '
   TTXYZ$ = Space$(92)
   Call REPLA(TTXYZ$, Space$(56), 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call REPLA(TTXYZ$, "Cliente: " + Label12, 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call REPLA(TTXYZ$, "    Tel: " + TRIM$(Label16), 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call REPLA(TTXYZ$, "    Fax: " + TRIM$(Label17), 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   '
   Call REPLA(TTXYZ$, Space$(56), 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call REPLA(TTXYZ$, "  S/Nro: " + Text1, 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call REPLA(TTXYZ$, "Referen: " + Text10, 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call REPLA(TTXYZ$, Space$(56), 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   '
   Call FRATEXTO(Text28, 46)
   For KKI% = 1 To CARETEX%
     TTXX$ = Space$(92)
     Call REPLA(TTXX$, RETEX$(KKI%), 10, 46)
     If KKI% = 1 Then Call REPLA(TTXX$, " Observ: ", 1, 9)
     Call REGAPP("ANOTAREP", Left$(MKS$(NROREP&) + Space$(10) + TTXX$, 96))
   Next KKI%
   '
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + String$(88, "="))
   Call CIERRARCH("ANOTAREP")
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   OKX% = 1
   Screen.MousePointer = 1
   Exit Sub
   '
98 Call CIERRARCH("PRESREPA")
   Call CIERRARCH("ANOTAREP")
   '
   Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   '
End Sub

Function NUPRESUP&()
   '
   NPP& = 1
   '
   FIN& = ULTREG&("PRESREPA")
   If FIN& > 0 Then
      For U& = FIN& To FIN& - 128 Step -1
         If U& < 1 Then GoTo 99
         XX$ = REGLEIDO$("PRESREPA", U&)
         NPPX& = CVS(Mid$(XX$, 1, 4))
         If NPPX& >= NPP& Then
            NPP& = 1 + NPPX&
         End If
      Next U&
   End If
   '
99 NUPRESUP& = NPP&
   '
End Function
'

