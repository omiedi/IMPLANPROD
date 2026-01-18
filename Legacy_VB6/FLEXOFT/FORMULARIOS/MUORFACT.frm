VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form MUORFACT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orden de Fabricación Activa"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0D0D0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8700
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9465
      Begin VB.PictureBox Picture999 
         Height          =   435
         Index           =   1
         Left            =   1680
         ScaleHeight     =   375
         ScaleWidth      =   1215
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame7"
         Height          =   4670
         Index           =   5
         Left            =   7455
         TabIndex        =   26
         Top             =   1470
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame6"
         Height          =   4670
         Index           =   4
         Left            =   5985
         TabIndex        =   25
         Top             =   1470
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame5"
         Height          =   4670
         Index           =   3
         Left            =   5250
         TabIndex        =   24
         Top             =   1470
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame4"
         Height          =   4670
         Index           =   2
         Left            =   4410
         TabIndex        =   23
         Top             =   1470
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame3"
         Height          =   4670
         Index           =   1
         Left            =   3465
         TabIndex        =   22
         Top             =   1470
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   4670
         Index           =   0
         Left            =   630
         TabIndex        =   21
         Top             =   1470
         Width           =   15
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4530
         Left            =   105
         TabIndex        =   1
         Top             =   1605
         Width           =   9075
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   105
         ScaleHeight     =   300
         ScaleWidth      =   9045
         TabIndex        =   19
         Top             =   1365
         Width           =   9075
         Begin VB.Label Label16 
            BackColor       =   &H00E0E0E0&
            Caption         =   " N-O Operación                  Máquina    Cant.  Horas   Comienza       Finaliza"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   30
            TabIndex        =   20
            Top             =   0
            Width           =   9045
         End
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   4020
         OleObjectBlob   =   "MUORFACT.frx":0000
         Top             =   2100
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   435
         Index           =   0
         Left            =   0
         OleObjectBlob   =   "MUORFACT.frx":0234
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   5040
         TabIndex        =   30
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
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
         Left            =   3435
         TabIndex        =   29
         Top             =   285
         Width           =   1485
      End
      Begin VB.Label Label18 
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
         Height          =   255
         Left            =   8220
         TabIndex        =   18
         Top             =   525
         Width           =   960
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Calculada"
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
         Left            =   6645
         TabIndex        =   17
         Top             =   585
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. O/F"
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
         Left            =   105
         TabIndex        =   16
         Top             =   280
         Width           =   960
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
         Height          =   255
         Left            =   1155
         TabIndex        =   15
         Top             =   210
         Width           =   1380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   2415
         TabIndex        =   14
         Top             =   280
         Width           =   960
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   3465
         TabIndex        =   13
         Top             =   210
         Width           =   330
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Dias de Atraso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7380
         TabIndex        =   12
         Top             =   915
         Width           =   720
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   8190
         TabIndex        =   11
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Solicitada"
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
         Left            =   6645
         TabIndex        =   10
         Top             =   280
         Width           =   1485
      End
      Begin VB.Label Label8 
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
         Height          =   255
         Left            =   8220
         TabIndex        =   9
         Top             =   210
         Width           =   960
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Producto"
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
         Left            =   105
         TabIndex        =   8
         Top             =   580
         Width           =   960
      End
      Begin VB.Label Label10 
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
         Height          =   255
         Left            =   1155
         TabIndex        =   7
         Top             =   525
         Width           =   2010
      End
      Begin VB.Label Label11 
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
         Height          =   465
         Left            =   1155
         TabIndex        =   6
         Top             =   840
         Width           =   3000
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
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
         Left            =   3915
         TabIndex        =   5
         Top             =   795
         Width           =   1485
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   5520
         TabIndex        =   4
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
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
         Left            =   3915
         TabIndex        =   3
         Top             =   1110
         Width           =   1485
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   5520
         TabIndex        =   2
         Top             =   1035
         Width           =   1065
      End
   End
End
Attribute VB_Name = "MUORFACT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CONSIMULACION%
Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   XMUOFA = Left
   YMUOFA = Top
End Sub

Private Sub Frame2_Click()
   Hide
End Sub

Private Sub Label2_Change()
   '
   Screen.MousePointer = 11
   NORFA$ = Label2
   Label4.Caption = ""
   Label6.Caption = ""
   Label8.Caption = ""
   Label10.Caption = ""
   Label11.Caption = ""
   Label13.Caption = ""
   Label15.Caption = ""
   Label20.Caption = ""
   List2.Clear
   '
   'SE VALIDA EN EL CASO QUE SEA 'SF' ES SIMULACION DE FABRICACION LEE OTRO ARCHVO
   ABUSCA$ = "ORDEFABP"
   If Left$(TRIM$(NORFA$), 2) = "SF" Then
      ABUSCA$ = "SRDEFABP"
   End If

5  For U& = ULTREG&(ABUSCA$) To 1 Step -1
      XX$ = REGLEIDO$(ABUSCA$, U&)
      If TRIM$(Mid$(XX$, 69, 12)) = NORFA$ Then
         CIIX% = CVI(Mid$(XX$, 83, 2))
         Label10.Caption = CODEXT$(CIIX%)
         Label11.Caption = DESCRIT$(CIIX%)
         Label13.Caption = Format$(CVD(Mid$(XX$, 89, 8)), "######0.0")
         Label6.Caption = Mid$(XX$, 97, 8)
         FEX = CVI(Mid$(XX$, 105, 2))
         FESOLENT = FEX
         Label8.Caption = FECHATEX$(FEX)
         Label4.Caption = Asc(Mid$(XX$, 108, 1))
         Label15.Caption = Format$(CVD(Mid$(XX$, 89, 8)) - CVD(Mid$(XX$, 119, 8)), "######0.0")
            NORFA$ = TRIM$(Label2.Caption)
         Label20.Caption = SAFETEXT$(VALOBADA("ORDEFABR", "ID_ORFA", "IDSUBOR='" & NORFA$ & "'"))
         GoTo 10
      End If
   Next U&
   '
   If Left$(TRIM$(NORFA$), 2) = "SF" Then
        If ABUSCA$ = "SRDEFABP" Then
          ABUSCA$ = "SRDEFABR"
          GoTo 5
         End If
    Else
        If ABUSCA$ = "ORDEFABP" Then
           ABUSCA$ = "ORDEFABR"
           GoTo 5
        End If
   End If
   '
10 ARCHOPERFAP$ = "OPERFAP0"
   Call COPYSTRU("OPERFAP", ARCHOPERFAP$)
   FULTOP = 0
   If Left$(TRIM$(NORFA$), 2) = "SF" Then
      ARCHOPERFAP$ = "SOPERFAP"
   End If
   For U& = 1 To ULTREG&(ARCHOPERFAP$)
      XX$ = REGLEIDO$(ARCHOPERFAP$, U&)
      If TRIM$(Left$(XX$, 12)) = NORFA$ Then
         DESCRIP1$ = ""
         TTXX$ = Space$(128)
         Call REPLA(TTXX$, CSTRING$(Mid$(XX$, 15, 2), 1, 4, 0, 2), 1, 4)
         If XVALO(TTXX$) > 0 Then
           OPER% = XVALO(TTXX$)
           If Left$(NORFA$, 2) = "OF" Then DESCRIP1$ = DESCRIOP$(CODIFAB%(NORFA$), OPER%)
           If Left$(NORFA$, 2) = "SF" Then DESCRIP1$ = DESCRIOP$(CODIFAB%(NORFA$, "SIMU_ORDEFABR"), OPER%)
         End If
         'Call REPLA(TTXX$, Mid$(XX$, 17, 48), 6, 26)
         Call REPLA(TTXX$, TRIM$(DESCRIP1$), 6, 26)
         Call REPLA(TTXX$, Mid$(XX$, 73, 6), 33, 6)
         'Call REPLA(TTXX$, CSTRING$(Mid$(XX$, 79, 8), 4, 10, 1, 2), 44, 10)
         'Call REPLA(TTXX$, CSTRING$(Mid$(XX$, 87, 8), 4, 10, 1, 2), 54, 10)
         SALDO# = CVD(Mid$(XX$, 79, 8)) - CVD(Mid$(XX$, 87, 8))
         If SALDO# < 0 Then SALDO# = 0
         Call REPLA(TTXX$, CSTRING$(MKD$(SALDO#), 4, 10, 1, 2), 39, 10)
         Call REPLA(TTXX$, Left$(HORMINSE$(60 * CVS(Mid$(XX$, 131, 4))), 6), 50, 6)
         FEX1 = CVI(Mid$(XX$, 135, 2))
         HORA1% = CVI(Mid$(XX$, 137, 2))
         FEX2 = CVI(Mid$(XX$, 139, 2))
         If FEX2 > FULTOP Then FULTOP = FEX2
         HORA2% = CVI(Mid$(XX$, 141, 2))
         If FEX1 > FEX2 Then
            FEX1 = 0: FEX2 = 0
            HORA1% = 0: HORA2% = 0
            Call REPLA(XX$, String$(8, 0), 135, 8)
            Call GRAREG(ARCHOPERFAP$, XX$, U&)
         End If
         Call REPLA(TTXX$, Left$(FECHATEX$(FEX1), 5), 58, 5)
         Call REPLA(TTXX$, Left$(HORATEXNOR$(HORA1%), 5), 64, 5)
         Call REPLA(TTXX$, Left$(FECHATEX$(FEX2), 5), 72, 5)
         Call REPLA(TTXX$, Left$(HORATEXNOR$(HORA2%), 5), 78, 5)
         List2.AddItem TTXX$
      End If
   Next U&
   
   Label18 = FECHATEX(FULTOP)
   DIATRA = 0: If FULTOP > FESOLENT Then DIATRA = DIENTRE(FESOLENT, FULTOP)
   If DIATRA > 999 Then DIATRA = 999
   Label6 = "": If DIATRA > 0.1 Then Label6 = TRIM$(Str$(DIATRA))
   '
   'List2.Height = 80 + 225 * List2.ListCount
   Height = List2.Top + List2.Height + 440
   Top = 8500 - Height
   Left = 400 + (12000 - Width) / 2
   '
   For KKI% = 0 To 5
     Frame1(KKI%).Top = Picture1.Top
     Frame1(KKI%).Height = Picture1.Height + List2.Height - 100
   Next KKI%
   Refresh
   List2.Refresh
   Screen.MousePointer = 1
   
'   Screen.MousePointer = 11
'   NORFA$ = Label2
'   Label4.Caption = ""
'   Label6.Caption = ""
'   Label8.Caption = ""
'   Label10.Caption = ""
'   Label11.Caption = ""
'   Label13.Caption = ""
'   Label15.Caption = ""
'   List2.Clear
'   '
'   ARCHORDEFABP$ = "ORDEFABP"
'   ARCHORDEFABR$ = "ORDEFABR"
'   If CONSIMULACION% > 0 Then
'     Call COPYSTRU("ORDEFABP", "SRDEFABP")
'     Call COPYSTRU("ORDEFABR", "SRDEFABR")
'     ARCHORDEFABP$ = "SRDEFABP"
'     ARCHORDEFABR$ = "SRDEFABR"
'   End If
'
''   ABUSCA$ = "ORDEFABP"
'   ABUSCA$ = ARCHORDEFABP$
'5  For U& = ULTREG&(ABUSCA$) To 1 Step -1
'      XX$ = REGLEIDO$(ABUSCA$, U&)
'      If TRIM$(Mid$(XX$, 69, 12)) = NORFA$ Then
'         CIIX% = CVI(Mid$(XX$, 83, 2))
'         Label10.Caption = CODEXT$(CIIX%)
'         Label11.Caption = DESCRIT$(CIIX%)
'         Label13.Caption = Format$(CVD(Mid$(XX$, 89, 8)), "######0.0")
'         Label6.Caption = Mid$(XX$, 97, 8)
'         FEX = CVI(Mid$(XX$, 105, 2))
'         FESOLENT = FEX
'         Label8.Caption = FECHATEX$(FEX)
'         Label4.Caption = Asc(Mid$(XX$, 108, 1))
'         Label15.Caption = Format$(CVD(Mid$(XX$, 89, 8)) - CVD(Mid$(XX$, 119, 8)), "######0.0")
'         GoTo 10
'      End If
'   Next U&
'   '
'   If ABUSCA$ = "ORDEFABP" Then
'      ABUSCA$ = "ORDEFABR"
'      GoTo 5
'   End If
'   '
'   If CONSIMULACION% > 0 Then
'    If ABUSCA$ = "SRDEFABP" Then
'       ABUSCA$ = "SRDEFABR"
'       GoTo 5
'    End If
'   End If
'
'10 FULTOP = 0
'   For U& = 1 To ULTREG&("OPERFAP")
'      XX$ = REGLEIDO$("OPERFAP", U&)
'      If TRIM$(Left$(XX$, 12)) = NORFA$ Then
'         TTXX$ = Space$(128)
'         Call REPLA(TTXX$, CSTRING$(Mid$(XX$, 15, 2), 1, 4, 0, 2), 1, 4)
'         Call REPLA(TTXX$, Mid$(XX$, 17, 48), 6, 26)
'         Call REPLA(TTXX$, Mid$(XX$, 73, 6), 33, 6)
'         'Call REPLA(TTXX$, CSTRING$(Mid$(XX$, 79, 8), 4, 10, 1, 2), 44, 10)
'         'Call REPLA(TTXX$, CSTRING$(Mid$(XX$, 87, 8), 4, 10, 1, 2), 54, 10)
'         SALDO# = CVD(Mid$(XX$, 79, 8)) - CVD(Mid$(XX$, 87, 8))
'         If SALDO# < 0 Then SALDO# = 0
'         Call REPLA(TTXX$, CSTRING$(MKD$(SALDO#), 4, 10, 1, 2), 39, 10)
'         Call REPLA(TTXX$, Left$(HORMINSE$(60 * CVS(Mid$(XX$, 131, 4))), 6), 50, 6)
'         FEX1 = CVI(Mid$(XX$, 135, 2))
'         HORA1% = CVI(Mid$(XX$, 137, 2))
'         FEX2 = CVI(Mid$(XX$, 139, 2))
'         If FEX2 > FULTOP Then FULTOP = FEX2
'         HORA2% = CVI(Mid$(XX$, 141, 2))
'         If FEX1 > FEX2 Then
'            FEX1 = 0: FEX2 = 0
'            HORA1% = 0: HORA2% = 0
'            Call REPLA(XX$, String$(8, 0), 135, 8)
'            Call GRAREG("OPERFAP", XX$, U&)
'         End If
'         Call REPLA(TTXX$, Left$(FECHATEX$(FEX1), 5), 58, 5)
'         Call REPLA(TTXX$, Left$(HORATEXNOR$(HORA1%), 5), 64, 5)
'         Call REPLA(TTXX$, Left$(FECHATEX$(FEX2), 5), 72, 5)
'         Call REPLA(TTXX$, Left$(HORATEXNOR$(HORA2%), 5), 78, 5)
'         List2.AddItem TTXX$
'      End If
'   Next U&
'
'   Label18 = FECHATEX(FULTOP)
'   DIATRA = 0: If FULTOP > FESOLENT Then DIATRA = DIENTRE(FESOLENT, FULTOP)
'   If DIATRA > 999 Then DIATRA = 999
'   Label6 = "": If DIATRA > 0.1 Then Label6 = TRIM$(Str$(DIATRA))
'   '
'   List2.Height = 80 + 225 * List2.ListCount
'   Height = List2.Top + List2.Height + 440
'   Top = 8500 - Height
'   Left = 400 + (12000 - Width) / 2
'   '
'   For KKI% = 0 To 5
'     Frame1(KKI%).Top = Picture1.Top
'     Frame1(KKI%).Height = Picture1.Height + List2.Height - 100
'   Next KKI%
'   Refresh
'   List2.Refresh
'   Screen.MousePointer = 1
End Sub

