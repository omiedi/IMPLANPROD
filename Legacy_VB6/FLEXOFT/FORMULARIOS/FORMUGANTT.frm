VERSION 5.00
Begin VB.Form FORMUGANTT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Secuencia de Carga por Equipo Productivo"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   9870
      Sorted          =   -1  'True
      TabIndex        =   30
      Top             =   3150
      Width           =   1905
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9870
      TabIndex        =   26
      Top             =   2050
      Width           =   1905
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Paisaje"
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
         Left            =   105
         TabIndex        =   29
         Top             =   550
         Width           =   960
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Retrato"
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
         Left            =   105
         TabIndex        =   28
         Top             =   300
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.CommandButton Command5 
         Height          =   540
         Left            =   1155
         Picture         =   "FORMUGANTT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   210
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Orden de Fabricación Activa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   210
      TabIndex        =   8
      Top             =   5565
      Width           =   9465
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   210
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   1575
         Width           =   9045
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8085
         TabIndex        =   23
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Qt.Saldo:"
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
         Left            =   6510
         TabIndex        =   22
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
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
         Left            =   5250
         TabIndex        =   21
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Qt.Proyectada:"
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
         Left            =   3675
         TabIndex        =   20
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1155
         TabIndex        =   19
         Top             =   1110
         Width           =   8100
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1155
         TabIndex        =   18
         Top             =   735
         Width           =   2010
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Producto:"
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
         Left            =   105
         TabIndex        =   17
         Top             =   780
         Width           =   960
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8085
         TabIndex        =   16
         Top             =   375
         Width           =   1170
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Solic:"
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
         Left            =   6510
         TabIndex        =   15
         Top             =   420
         Width           =   1485
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
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
         Left            =   5250
         TabIndex        =   14
         Top             =   375
         Width           =   1170
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Destino:"
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
         Left            =   4200
         TabIndex        =   13
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3465
         TabIndex        =   12
         Top             =   380
         Width           =   330
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   2415
         TabIndex        =   11
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
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
         Height          =   260
         Left            =   1155
         TabIndex        =   10
         Top             =   380
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. O/F:"
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
         Left            =   105
         TabIndex        =   9
         Top             =   420
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5220
      Left            =   210
      ScaleHeight     =   5160
      ScaleWidth      =   9405
      TabIndex        =   6
      Top             =   210
      Width           =   9465
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         Left            =   840
         TabIndex        =   31
         Top             =   4940
         Width           =   8320
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   4380
         Left            =   9185
         TabIndex        =   25
         Top             =   560
         Width           =   225
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   -105
         X2              =   9660
         Y1              =   4935
         Y2              =   4935
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   9765
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   840
         X2              =   840
         Y1              =   0
         Y2              =   6930
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Rango"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   9870
      TabIndex        =   1
      Top             =   105
      Width           =   1900
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mensual"
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
         Index           =   3
         Left            =   315
         TabIndex        =   5
         Top             =   1365
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Quincenal"
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
         Index           =   2
         Left            =   315
         TabIndex        =   4
         Top             =   1050
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Semanal"
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
         Index           =   1
         Left            =   315
         TabIndex        =   3
         Top             =   735
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Diario"
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
         Left            =   315
         TabIndex        =   2
         Top             =   420
         Width           =   1275
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   9870
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   5670
      Width           =   1905
   End
   Begin VB.Label TEPRUEBA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10395
      TabIndex        =   7
      Top             =   5775
      Visible         =   0   'False
      Width           =   1020
   End
End
Attribute VB_Name = "FORMUGANTT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click()
   PORIAN = Printer.Orientation
   If Option2.Value = True Then
      Printer.Orientation = 1
    Else
      Printer.Orientation = 2
   End If
   PRINTFORM
   Printer.Orientation = PORIAN
End Sub

Private Sub Form_Load()
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
   End If
   '
   FEIPER = HOY(HO$)
   While DIASEM%(FEIPER) <> 1
      FEIPER = DIANTE(FEIPER)
   Wend
   '
   HScroll1.Min = FEIPER - 32 * 13
   HScroll1.Max = FEIPER + 32 * 13
   HScroll1.Value = FEIPER
   '
   Call DIBULINVER
   Call CALCUFECHA
   Call DIBUFECHA
   Call CARLIORF
   Call DIBUEQUIP
   Call PREGANTT
   If List1.ListIndex >= 0 Then
      NORFA$ = TRIM$(Left$(List1.TEXT, 24))
      Call MUEDATORF(NORFA$)
   End If
   '
   YACARGANTT% = 1
   '
End Sub

Private Sub HScroll1_Change()
   If YACARGANTT% = 1 Then
      FEIPERANT = FEIPER
      FEIPER = HScroll1.Value
      SENTI% = 1
      If FEIPER < FEIPERANT Then SENTI% = -1
10    FEX = FEIPER
      If Val(Left$(FECHATEX$(FEX), 2)) < 1 Then
         FEIPER = FEIPER + SENTI%
         GoTo 10
      End If
      '
      If Val(Mid$(FECHATEX$(FEX), 4, 2)) < 1 Then
         FEIPER = FEIPER + SENTI%
         GoTo 10
      End If
      FF1 = DIPOST(FEX)
      FF2 = DIANTE(FF1)
      If FF2 <> FEIPER Then
         FEIPER = FEIPER + SENTI%
         GoTo 10
      End If
      '
   End If
   '
   HScroll1.Value = FEIPER
   FORMUGANTT.Picture1.Cls
   Call DIBULINVER
   Call CALCUFECHA
   Call DIBUFECHA
   Call DIBUEQUIP
   Call PREGANTT
End Sub

Private Sub List1_Click()
   If CONTROCLICK% = 0 Then
      Call PREGANTT
      NORFA$ = TRIM$(Left$(List1.TEXT, 24))
      Call MUEDATORF(NORFA$)
   End If
End Sub

Private Sub List3_Click()
   CIIX% = CODINT%(List3.TEXT)
   CONTROCLICK% = (-1)
   OFACT& = 0
   For U& = 1 To List1.ListCount
      FORMUGANTT.List1.Selected(U& - 1) = False
      V& = Val(Mid$(FORMUGANTT.List1.List(U& - 1), 25))
      If CIOFA%(V&) = CIIX% Then
         FORMUGANTT.List1.Selected(U& - 1) = True
         OFACT& = U&
      End If
   Next U&
   CONTROCLICK% = 0
   List1.Refresh
   Call PREGANTT
   If OFACT& > 0 Then
      List1.ListIndex = OFACT& - 1
      NORFA$ = TRIM$(Left$(List1.TEXT, 24))
      Call MUEDATORF(NORFA$)
   End If
End Sub

Private Sub Option1_Click(Index As Integer)
   FORMUGANTT.Picture1.Cls
   Call DIBULINVER
   Call CALCUFECHA
   Call DIBUFECHA
   Call DIBUEQUIP
   Call PREGANTT
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   For UI% = 1 To cabarr%
      If X >= XINI1(UI%) Then
        If X <= XFIN1(UI%) Then
          If Y >= YINI1(UI%) Then
            If Y <= YINI1(UI%) + 145 Then
              NORFAX$ = NUOFA$(UI%)
              For U& = 1 To List1.ListCount
                 V& = Val(Mid$(List1.List(U& - 1), 25))
                 If NORFB$(V&) = NORFAX$ Then
                   CONTROCLICK% = -1
                   FORMUGANTT.List1.ListIndex = U& - 1
                   CONTROCLICK% = 0
                   OFACT& = U&
                   Call PREGANTT
                   Call MUEDATORF(NORFAX$)
                   Exit Sub
                 End If
              Next U&
            End If
          End If
        End If
      End If
   Next UI%
End Sub

Private Sub VScroll1_Change()
   PRIEQVIS% = VScroll1.Value + 1
   FORMUGANTT.Picture1.Cls
   Call DIBULINVER
   Call CALCUFECHA
   Call DIBUFECHA
   Call DIBUEQUIP
   Call PREGANTT
End Sub
