VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form RIVACOM07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen de IVA"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2160
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox IMPOGRA 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   3465
      TabIndex        =   23
      ToolTipText     =   "  "
      Top             =   315
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5160
      Picture         =   "RIVACOM07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   540
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00CDDADA&
      Caption         =   "TOTALES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   210
      TabIndex        =   6
      Top             =   210
      Width           =   4905
      Begin VB.TextBox IMPOGRA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   930
         TabIndex        =   4
         Text            =   " "
         ToolTipText     =   "  "
         Top             =   2010
         Width           =   1845
      End
      Begin VB.TextBox IMPOGRA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   930
         TabIndex        =   3
         Text            =   " "
         ToolTipText     =   "  "
         Top             =   1740
         Width           =   1845
      End
      Begin VB.TextBox IMPOGRA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   930
         TabIndex        =   2
         Text            =   " "
         ToolTipText     =   "  "
         Top             =   1470
         Width           =   1845
      End
      Begin VB.TextBox IMPOGRA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   930
         TabIndex        =   1
         Text            =   " "
         ToolTipText     =   "  "
         Top             =   1215
         Width           =   1845
      End
      Begin VB.TextBox IMPOGRA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   930
         TabIndex        =   0
         ToolTipText     =   "  "
         Top             =   930
         Width           =   1845
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00CDDADA&
         BorderStyle     =   0  'None
         Height          =   1650
         Left            =   40
         TabIndex        =   7
         Top             =   720
         Width           =   850
         Begin VB.Label TASIVA 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   80
            TabIndex        =   13
            ToolTipText     =   "Doble Click Para Configurar Tasa 3 I.V.A."
            Top             =   810
            Width           =   750
         End
         Begin VB.Label TASIVA 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "10.5%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   80
            TabIndex        =   12
            ToolTipText     =   "Doble Click Para Configurar Tasa 2 I.V.A."
            Top             =   570
            Width           =   750
         End
         Begin VB.Label TASIVA 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "27.0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   80
            TabIndex        =   11
            ToolTipText     =   "Doble Click Para Configurar Tasa 1 I.V.A."
            Top             =   270
            Width           =   750
         End
         Begin VB.Label TASIVA 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "21.0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   105
            TabIndex        =   10
            ToolTipText     =   "Doble Click Para Configurar Tasa Basica I.V.A."
            Top             =   0
            Width           =   750
         End
         Begin VB.Label TASIVA 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Exento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   0
            TabIndex        =   9
            Top             =   1080
            Width           =   800
         End
         Begin VB.Label TASIVA 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "N/Grav"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   0
            TabIndex        =   8
            Top             =   1350
            Width           =   795
         End
      End
      Begin VB.Label IMPOIVA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   5
         Left            =   3120
         TabIndex        =   21
         Top             =   2010
         Width           =   1635
      End
      Begin VB.Label IMPOIVA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   4
         Left            =   3120
         TabIndex        =   20
         Top             =   1740
         Width           =   1635
      End
      Begin VB.Label IMPOIVA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   3
         Left            =   3120
         TabIndex        =   19
         Top             =   1470
         Width           =   1635
      End
      Begin VB.Label IMPOIVA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   2
         Left            =   3120
         TabIndex        =   18
         Top             =   1215
         Width           =   1635
      End
      Begin VB.Label IMPOIVA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   1
         Left            =   3120
         TabIndex        =   17
         Top             =   930
         Width           =   1635
      End
      Begin VB.Label IMPOIVA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   3120
         TabIndex        =   16
         Top             =   660
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "                   Base Imponible        Importe IVA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   15
         Top             =   300
         Width           =   3690
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   3885
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label ITOGRA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   930
         TabIndex        =   14
         Top             =   660
         Width           =   1845
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "RIVACOM07.frx":030A
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "RIVACOM07.frx":053E
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1260
      TabIndex        =   25
      Top             =   3360
      Width           =   1800
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1260
      TabIndex        =   24
      Top             =   2940
      Width           =   960
   End
   Begin VB.Label TONECOM 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "RIVACOM07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   For KKII% = 0 To 5
      If XVALO(IMPOGRA(KKII%).TEXT) < (-0.005) Then GoTo 20
      If XVALO(IMPOIVA(KKII%).Caption) < (-0.005) Then GoTo 20
   Next KKII%
   Hide
   Exit Sub
   '
20 Call MENSERR(24, "No se Admiten Valores Negativos")
   '
End Sub

Private Sub Form_Load()
    For IXX% = 0 To 3
      TXXX$ = TRIM$(Str$(IXX%))
      Tasa = XVALO(Control$("", "TASIVA-" + TXXX$))
      TASIVA(IXX%).Caption = "0.00%"
      If Tasa > 0.005 Then
          TASIVA(IXX%).Caption = TRIM$(CSTRING$(MKS$(Tasa), 3, 6, 2, 2) + "%")
      End If
    Next IXX%
    Label2.Caption = TASIVA(0).Caption
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 1
End Sub

Private Sub IMPOGRA_Change(Index As Integer)
   '
   TASIX$ = TRIM$(REPLACAR$(TASIVA(Index).Caption, "%", ""))
   TASI = XVALO(TASIX$)
   IGRA# = XVALO(IMPOGRA(Index).TEXT)
   If Index < 4 Then
     IMPOIVA(Index) = CSTRING$(MKD$(TASI * IGRA# / 100), 4, 14, 2, 2)
     IMPOIVA(Index).Refresh
   End If
   '
   TOGRAVA# = 0
   For KKI% = 1 To 5
     IGRA# = XVALO(IMPOGRA(KKI%).TEXT)
     TOGRAVA# = TOGRAVA# + IGRA#
   Next KKI%
   '
   TOTIMPU# = XVALO(TONECOM.Caption)
   ITOGRA.Caption = CSTRING$(MKD$(TOTIMPU# - TOGRAVA#), 4, 13, 2, 2)
   ITOGRA.Refresh
   IMPOGRA(0).TEXT = ITOGRA.Caption
End Sub

Private Sub IMPOGRA_DblClick(Index As Integer)
   For KKII% = 0 To 5
      IMPOGRA(KKII%).TEXT = ""
      IMPOIVA(KKII%).Caption = ""
   Next KKII%
   IMPOGRA(Index).TEXT = TRIM$(FORMATNUM$(XVALO(TONECOM.Caption), "F15.2"))
End Sub

Private Sub IMPOGRA_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii% = 13 Then
      If Index < 5 Then
         IMPOGRA(Index + 1).SetFocus
       Else
         Command1.SetFocus
      End If
   End If
End Sub

Private Sub IMPOGRA_LostFocus(Index As Integer)
   IGRA# = XVALO(IMPOGRA(Index).TEXT)
   IMPOGRA(Index).TEXT = TRIM$(FORMATNUM$(IGRA#, "F13.2"))
End Sub

Private Sub IMPOIVA_Change(Index As Integer)
   SUMAIVA# = 0
   Label2 = ""
   For KKII% = 0 To 5
      SUMAIVA# = SUMAIVA# + ROUND(XVALO(IMPOIVA(KKII%)))
      If Abs(XVALO(IMPOGRA(KKII%))) > 0.005 Then
         If Label2 = "" Then
             Label2 = TASIVA(KKII%).Caption
           Else
             Label2 = "Varios"
         End If
      End If
   Next KKII%
   Label3 = FORMATNUM$(SUMAIVA#, "F15.2")
End Sub

Private Sub ITOGRA_CHANGE()
   '
   
   TASIX$ = TRIM$(REPLACAR$(TASIVA(0).Caption, "%", ""))
   TASI = XVALO(TASIX$)
   IGRA# = XVALO(ITOGRA.Caption)
   IMPOGRA(0) = TRIM$(FORMATNUM$(IGRA#, "F15.2"))
   IMPOIVA(0) = CSTRING$(MKD$(TASI * IGRA# / 100), 4, 14, 2, 2)
   IMPOIVA(0).Refresh
   '
End Sub

Private Sub ITOGRA_DblClick()
   For KKII% = 0 To 5
      IMPOGRA(KKII%).TEXT = ""
      IMPOIVA(KKII%).Caption = ""
   Next KKII%
   ITOGRA.Caption = TRIM$(FORMATNUM$(XVALO(TONECOM.Caption), "F15.2"))
End Sub



Private Sub TASIVA_DblClick(Index As Integer)
    '
    If Index > 3 Then Exit Sub
    If DERACCE%("STPTIVA", "Configuracion de Tasas de I.V.A.") >= 2 Then
        If Index = 0 Then
            TIP$ = "Basica"
        Else
            TIP$ = "Tasa - " & Index
        End If
        TASACT$ = FORMATNUM$(Control$("", "TASIVA-" & TRIM$(Str$(Index))), "F5.2")
10      TASAX$ = InputBox("Tasa I.V.A. " & TIP$, "Configuracion Tasa I.V.A.", TASACT$)
        If TRIM$(TASAX$) = "" Then TASAX$ = TASACT$          ' POR SI CANCELA
        If XVALO(TASAX$) > 0 And XVALO(TASAX$) < 100 Then
            Call GRACONTROL$("", "TASIVA-" & TRIM$(Str(Index)), TASAX$)
            TASIVA(Index) = FORMATNUM$(XVALO(TASAX$), "F5.2") & "%"
            'Actualiza importes de IVA segun modificacion de tasas
            TASIX$ = TRIM$(REPLACAR$(TASIVA(Index).Caption, "%", ""))
            TASI = XVALO(TASIX$)
            IGRA# = XVALO(ITOGRA.Caption)
            IMPOGRA(Index) = TRIM$(FORMATNUM$(IGRA#, "F15.2"))
            IMPOIVA(Index) = CSTRING$(MKD$(TASI * IGRA# / 100), 4, 14, 2, 2)
            IMPOIVA(Index).Refresh
        Else
            Call MENSERR(24, "Tasa Ingresada no Valida")
            GoTo 10
        End If
    End If
    '
End Sub

Private Sub TONECOM_Change()
   For KKII% = 0 To 5
      IMPOGRA(KKII%).TEXT = ""
      IMPOIVA(KKII%).Caption = ""
   Next KKII%
   ITOGRA = TRIM$(FORMATNUM$(XVALO(TONECOM), "F15.2"))
   ITOGRA.Refresh
End Sub

