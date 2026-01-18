VERSION 5.00
Begin VB.Form RESUVE00 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Gestión - Resúmenes Mensuales de Ventas"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5165
      ScaleHeight     =   75
      ScaleWidth      =   1005
      TabIndex        =   38
      Top             =   315
      Width           =   1065
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H80000016&
      Height          =   855
      Left            =   315
      ScaleHeight     =   795
      ScaleWidth      =   4575
      TabIndex        =   4
      Top             =   210
      Width           =   4635
      Begin VB.TextBox MESAN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         TabIndex        =   5
         Text            =   " "
         Top             =   210
         Width           =   2220
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   225
         Left            =   2835
         TabIndex        =   7
         Top             =   840
         Width           =   15
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Período:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   10
         Left            =   -735
         TabIndex        =   6
         Top             =   315
         Width           =   2220
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Gestión de Cuentas Corrientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Left            =   315
      TabIndex        =   3
      Top             =   1260
      Width           =   8835
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Dispositivo de Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4515
         TabIndex        =   35
         Top             =   3360
         Width           =   4005
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Impresión"
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
            Left            =   2310
            TabIndex        =   37
            Top             =   420
            Width           =   1170
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Pantalla"
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
            Left            =   630
            TabIndex        =   36
            Top             =   420
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Resumen Contable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   4515
         TabIndex        =   29
         Top             =   1680
         Width           =   4005
         Begin VB.PictureBox Picture15 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3135
            TabIndex        =   30
            Top             =   420
            Width           =   3200
            Begin VB.CommandButton Command22 
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
               Left            =   2580
               Picture         =   "RESUVE00.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   34
               ToolTipText     =   "Re-Imputación de Comprobantes de Ventas"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command19 
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
               Picture         =   "RESUVE00.frx":014A
               Style           =   1  'Graphical
               TabIndex        =   32
               ToolTipText     =   "Consulta / Listado por Cuenta Contable"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command20 
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
               Left            =   2020
               Picture         =   "RESUVE00.frx":0454
               Style           =   1  'Graphical
               TabIndex        =   31
               ToolTipText     =   "Consulta de Asiento Contable por Pantalla"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Contabilización de Ventas"
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
               TabIndex        =   33
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Resumen Operativo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   4515
         TabIndex        =   25
         Top             =   315
         Width           =   4005
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2055
            TabIndex        =   26
            Top             =   420
            Width           =   2115
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
               Picture         =   "RESUVE00.frx":075E
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Resumen Men- sual por Cuenta"
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
               TabIndex        =   28
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cobranza y Ajustes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   315
         TabIndex        =   18
         Top             =   2520
         Width           =   4005
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2475
            TabIndex        =   19
            Top             =   840
            Width           =   2535
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
               Left            =   -20
               Picture         =   "RESUVE00.frx":0BA0
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Detalle de Cobranza en Cuenta Corriente"
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
               Top             =   105
               Width           =   1905
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2475
            TabIndex        =   22
            Top             =   420
            Width           =   2535
            Begin VB.CommandButton Command7 
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
               Picture         =   "RESUVE00.frx":0EAA
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Ajustes y Otros Movimientos"
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
               Left            =   735
               TabIndex        =   24
               Top             =   0
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ventas"
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
         Left            =   315
         TabIndex        =   8
         Top             =   315
         Width           =   4005
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1155
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   15
            Top             =   1260
            Width           =   2325
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
               Picture         =   "RESUVE00.frx":11B4
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Facturación en Cuenta Corriente y de Contado"
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
               Left            =   735
               TabIndex        =   17
               Top             =   0
               Width           =   1515
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   12
            Top             =   840
            Width           =   2325
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
               Left            =   -20
               Picture         =   "RESUVE00.frx":14BE
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Notas de Débito"
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
               Left            =   735
               TabIndex        =   14
               Top             =   105
               Width           =   1410
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   9
            Top             =   420
            Width           =   2325
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
               Picture         =   "RESUVE00.frx":17C8
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Notas de Crédito"
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
               Left            =   735
               TabIndex        =   11
               Top             =   105
               Width           =   1695
            End
         End
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   855
      Left            =   8295
      Picture         =   "RESUVE00.frx":1AD2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   220
      Width           =   855
   End
   Begin VB.CommandButton Command29 
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
      Height          =   855
      Left            =   7350
      Picture         =   "RESUVE00.frx":1C1C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   220
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Direct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6405
      Picture         =   "RESUVE00.frx":1F26
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   220
      Width           =   855
   End
   Begin VB.Line Line2 
      X1              =   4945
      X2              =   4945
      Y1              =   210
      Y2              =   1050
   End
   Begin VB.Line Line1 
      X1              =   315
      X2              =   6405
      Y1              =   1060
      Y2              =   1060
   End
   Begin VB.Image Image4 
      Height          =   510
      Left            =   4725
      Picture         =   "RESUVE00.frx":2230
      Top             =   550
      Width           =   2010
   End
End
Attribute VB_Name = "RESUVE00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub


Sub Command19_Click()
    If Option1.Value = True Then
         Call CONECRUN("*COSUDIV", "Consulta por Cuenta Contable")
         Exit Sub
      ElseIf Option2.Value = True Then
         Call CONECRUN("*SOSUDIV", "Listado de Ventas por Cuenta Contable")
         Exit Sub
    End If
End Sub

Sub Command2_Click()
    MOX% = ALTERNA%("Secuencia por Recibo\Detalle por Cliente")
    On MOX% GoTo 203, 204
    Exit Sub
    '
203 Screen.MousePointer = 11
    LU$ = LUDAT$
    FileCopy LU$ + "DETCOBV.DAT", LU$ + "DETACCX.DAT"
    C1$ = TRIM$(REMEVE00.MESAN)
    C2$ = TRIM$(REMEVE00.Text4(0).TEXT)
    C3$ = TRIM$(REMEVE00.Text4(1).TEXT)
    Call HEADERS("DETACCX", "")
    Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Call HEADERS("DETACCX", "Rubro: Recibos de Cobranza")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODECCX")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDECCX")
    End If
    Exit Sub
    '
204 Call HEADERS("RECOBCLI", "")
    Call HEADERS("RECOBCLI", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Call HEADERS("RECOBCLI", "Rubro: Recibos de Cobranza")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*COCOBCL")
      ElseIf Option2.Value = True Then
         Call FINAL("*SOCOBCL")
    End If
    Exit Sub
    '
End Sub
'

Sub Command20_Click()
    Call CONASIDIA
End Sub

Sub Command22_Click()
   '
   Command22.Enabled = False
   '
5  CORRESP1$ = TRIM$(MESAN)
   CORRESQ1$ = MESANO$(CORRESP1$, DES%, HAS%)
   FEX = DES%: DESFE1$ = FECHATEX$(FEX)
   FEX = HAS%: HASFE1$ = FECHATEX$(FEX)
   '
   CORRESQ1$ = TRIM$(CORRESP1$)
   ABREMES1$ = Left$(CORRESQ1$, 3) + Mid$(DESFE1$, 7, 2)
   LUXI$ = LUDAT$
   ARCHI$ = "ICV" + ABREMES1$ + ".DAT"
   N2% = FILEOPEN(LU$ + ARCHI$, 0, 96)
   Dim BULEC As String * 96
   '
10 NUFACLI$ = REGSEL$("IFACVEN")
   If Len(NUFACLI$) <= 4 Then
      Close #N2%
      GoTo 99
   End If
   NF$ = Left$(NUFACLI$, 18)
   NC% = CVI(Mid$(NUFACLI$, 59, 2))
   IMPO# = CVD(Mid$(NUFACLI$, 51, 8))
   NRO& = CVS(Mid$(NUFACLI$, 61, 4))
   If Abs(IMPO#) < 0.005 Then
      Call MENSERR(24, "Documento Anulado")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   For U& = 1 To LOF(N2%) / 96
     Get #N2%, U&, BULEC$
     If Left$(BULEC$, 2) = Left$(NF$, 2) Then
       If CVS(Mid$(BULEC$, 3, 4)) = NRO& Then
         If CVI(Mid$(BULEC$, 7, 2)) = NC% Then
           REG& = U&
           Y6$ = Mid$(BULEC$, 17, 80)
           GoTo 20
         End If
       End If
     End If
   Next U&
   Close #N2%
   Call MENSERR(24, "Documento no Encontrado - Consulte.")
   GoTo 99
   '
20 ATRI$ = TRIM$(Left$(NUFACLI$, 50)) + " ($ " + TRIM$(CSTRING$(MKD$(IMPO#), 4, 12, 2, 2)) + ")"
   '
   Call BLANARCH("!CARICOM")
   For J% = 1 To 8
      IPA# = 0
      CUE% = CVI(Mid$(Y6$, 10 * J% - 9, 2))
      If CUE% > 0 Then
         IPA# = CVD(Mid$(Y6$, 10 * J% - 7, 8))
         Y$ = REGBLAN$("!CARICOM")
         Call REPLA(Y$, MKI$(CUE%), 1, 2)
         Call REPLA(Y$, MKD$(IPA#), 3, 8)
         Call REGAPP("!CARICOM", Y$)
      End If
   Next J%
   Call CIERRARCH("!CARICOM")
   '
30 VTB% = VENTABU%("CARIFAVE/" + ATRI$)
   If VTB% < 0 Then GoTo 10
   '
   TOTIMPU# = 0: Y6$ = ""
   For U& = 1 To ULTREG&("!CARICOM")
       Y$ = REGLEIDO$("!CARICOM", U&)
       CUEB$ = Left$(Y$, 2)
       If CVI(CUEB$) > 0 Then
          DECUE$ = ECOARCH$("CUECON", CUEB$)
          If DECUE$ <> "" Then
             IMBB$ = Mid$(Y$, 3, 8)
             If CVD(IMBB$) < 0 Then
                Call MENSERR(24, "No se Admiten Valores Negativos")
                GoTo 30
             End If
             Y6$ = Y6$ + CUEB$ + IMBB$
             TOTIMPU# = TOTIMPU# + CVD(IMBB$)
          End If
       End If
   Next U&
   '
   If Abs(Abs(IMPO#) - TOTIMPU#) >= 0.005 Then
      Call MENSERR(24, "Suma de Imputaciones Difiere de Total")
      GoTo 30
   End If
   '
   Y6$ = Y6$ + String$(80, 0)
   Get #N2%, REG&, BULEC$
   X$ = BULEC$
   Call REPLA(X$, Y6$, 17, 80)
   BULEC$ = X$
   Put #N2%, REG&, BULEC$
   Close #N2%
   GoTo 5
   '
99 Call FRESHECHO("IFACVEN")
   Call COMUNI("Para Efectivizar los Cambios\Debe Recalcular el Resumen Mensual Ahora.")
   Command22.Enabled = True
   Hide
   '
End Sub

Sub Command3_Click()
    Screen.MousePointer = 11
    LU$ = LUDAT$
    FileCopy LU$ + "DETFACV.DAT", LU$ + "DETACCX.DAT"
    C1$ = TRIM$(MESAN)
    C2$ = TRIM$(REMEVE00.Text4(0).TEXT)
    C3$ = TRIM$(REMEVE00.Text4(1).TEXT)
    Call HEADERS("DETACCX", "")
    Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Call HEADERS("DETACCX", "Rubro: Facturacion en Cuenta Corriente / Contado")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODECCX")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDECCX")
    End If
End Sub

Sub Command4_Click()
    Screen.MousePointer = 11
    LU$ = LUDAT$
    FileCopy LU$ + "DETDEBV.DAT", LU$ + "DETACCX.DAT"
    C1$ = TRIM$(REMEVE00.MESAN)
    C2$ = TRIM$(REMEVE00.Text4(0).TEXT)
    C3$ = TRIM$(REMEVE00.Text4(1).TEXT)
    Call HEADERS("DETACCX", "")
    Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Call HEADERS("DETACCX", "Rubro: Notas de Debito")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODECCX")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDECCX")
    End If
End Sub


Sub Command5_Click()
    Screen.MousePointer = 11
    LU$ = LUDAT$
    FileCopy LU$ + "DETCREV.DAT", LU$ + "DETACCX.DAT"
    C1$ = TRIM$(REMEVE00.MESAN)
    C2$ = TRIM$(REMEVE00.Text4(0).TEXT)
    C3$ = TRIM$(REMEVE00.Text4(1).TEXT)
    Call HEADERS("DETACCX", "")
    Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Call HEADERS("DETACCX", "Rubro: Notas de Credito")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODECCX")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDECCX")
    End If
End Sub

Sub Command6_Click()
    If Option1.Value = True Then
         Call FINAL("*RESUCCD")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIRECCD")
    End If
End Sub

Sub Command7_Click()
    Screen.MousePointer = 11
    LU$ = LUDAT$
    FileCopy LU$ + "DETOTRV.DAT", LU$ + "DETACCX.DAT"
    C1$ = TRIM$(REMEVE00.MESAN)
    C2$ = TRIM$(REMEVE00.Text4(0).TEXT)
    C3$ = TRIM$(REMEVE00.Text4(1).TEXT)
    Call HEADERS("DETACCX", "")
    Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Call HEADERS("DETACCX", "Rubro: Otros Movimientos de Cuenta Corriente")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODECCX")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDECCX")
    End If
End Sub

Sub Command8_Click()
    Call ACCDIR("COGEVEN")
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub MESAN_GOTFOCUS()
    Command1.SetFocus
End Sub

Sub CONASIDIA()
    '
    MESA$ = TRIM$(MESAN)
    MESB$ = MESANO$(MESA$, DES%, HAS%)
    If MESB$ <> MESA$ Then
      Exit Sub
    End If
    '
    Screen.MousePointer = 11
    J& = 0
    Call BLANARCH("!ASIDIAFI")
    TTXXY$ = Chr$(1) + Chr$(1) + MKI$(HAS%)
    For IL& = 1 To ULTREG&("ASIDIAR")
        X$ = REGLEIDO$("ASIDIAR", IL&)
        If Left$(X$, 4) = TTXXY$ Then
            CUENU& = CVI(Mid$(X$, 5, 2))
            IDEB# = CVD(Mid$(X$, 39, 8))
            If Abs(IDEB#) >= 0.005 Then
                IHAB# = 0
                If IDEB# < 0 Then
                    IHAB# = Abs(IDEB#)
                    IDEB# = 0
                End If
                CUES$ = Space$(12)
                If CUENU& > 0 Then
                    If CUENU& <= ULTREG&("CUENTAS") Then
                        CUES$ = Left$(REGLEIDO$("CUENTAS", CUENU&), 12)
                    End If
                End If
                DENOCUES$ = ECOARCH$("ECUECON", CUES$)
                '
                Z$ = REGBLAN$("ASIDIAFI")
                Call REPLA(Z$, CUES$, 1, 12)
                Call REPLA(Z$, DENOCUES$, 13, 30)
                Call REPLA(Z$, MKD$(IDEB#), 43, 8)
                Call REPLA(Z$, MKD$(IHAB#), 51, 8)
                '
                J& = J& + 1
                Call GRAREG("!ASIDIAFI", Z$, J&)
            End If
        End If
    Next IL&
    Call SETULTREG("!ASIDIAFI", J&)
    '
    For K1& = 1 To J&
        For K2& = K1& + 1 To J&
            X1$ = REGLEIDO$("!ASIDIAFI", K1&)
            X2$ = REGLEIDO$("!ASIDIAFI", K2&)
            If Left$(X1$, 12) > Left$(X2$, 12) Then
                X0$ = X1$
                X1$ = X2$
                X2$ = X0$
                Call GRAREG("!ASIDIAFI", X1$, K1&)
                Call GRAREG("!ASIDIAFI", X2$, K2&)
            End If
        Next K2&
    Next K1&
    Call CIERRARCH("!ASIDIAFI")
    '
    Screen.MousePointer = 1
    VTB% = VENTABU%("ASIDIAFI")
    '
End Sub
'
