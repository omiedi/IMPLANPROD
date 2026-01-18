VERSION 5.00
Begin VB.Form CONTAB00 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Contabilidad General"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10860
   ClipControls    =   0   'False
   Icon            =   "CONTAB00.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox SORTLIST 
      Height          =   255
      Left            =   6300
      Sorted          =   -1  'True
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Control Pos.Diferida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3990
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   9870
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   37
      Top             =   3360
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Plan"
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
      Left            =   9870
      Picture         =   "CONTAB00.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Plan General de Cuentas"
      Top             =   4620
      Width           =   855
   End
   Begin VB.CommandButton Command27 
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
      Left            =   9870
      Picture         =   "CONTAB00.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Acceso Directo a Aplicaciones Complementarias"
      Top             =   2310
      Width           =   855
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Setup"
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
      Left            =   9870
      Picture         =   "CONTAB00.frx":075E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Configuración de Funcionamiento"
      Top             =   3675
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
      Left            =   9870
      Picture         =   "CONTAB00.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1365
      Width           =   855
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
      Left            =   9870
      Picture         =   "CONTAB00.frx":0EAA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   420
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consultas e Informes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      Left            =   5040
      TabIndex        =   9
      Top             =   315
      Width           =   4635
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Otros Reportes de Control"
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
         Left            =   210
         TabIndex        =   30
         Top             =   4620
         Width           =   4215
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   31
            Top             =   420
            Width           =   3375
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
               Left            =   2760
               Picture         =   "CONTAB00.frx":0FF4
               Style           =   1  'Graphical
               TabIndex        =   33
               ToolTipText     =   "Ajustee por Inflación"
               Top             =   0
               Width           =   560
            End
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
               Picture         =   "CONTAB00.frx":12FE
               Style           =   1  'Graphical
               TabIndex        =   32
               ToolTipText     =   "Informes Standard Normalizados"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Informes Contables Complementarios"
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
               Left            =   840
               TabIndex        =   34
               Top             =   105
               Width           =   1845
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Opciones de Balance"
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
         Left            =   210
         TabIndex        =   26
         Top             =   3150
         Width           =   4215
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2685
            TabIndex        =   27
            Top             =   420
            Width           =   2745
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
               Picture         =   "CONTAB00.frx":1608
               Style           =   1  'Graphical
               TabIndex        =   28
               ToolTipText     =   "Balances de Sumas y Saldos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Balance Mensual y Ge- neral - Sumas y Saldos"
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
               Left            =   630
               TabIndex        =   29
               Top             =   120
               Width           =   2160
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Informes de Diario"
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
         Left            =   210
         TabIndex        =   21
         Top             =   420
         Width           =   4215
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   22
            Top             =   420
            Width           =   3375
            Begin VB.CommandButton Command13 
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
               Picture         =   "CONTAB00.frx":1912
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "Libro Diario General"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command12 
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
               Left            =   2760
               Picture         =   "CONTAB00.frx":1C1C
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "Mayor por Cuenta"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Reportes de la Contabilidad General"
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
               Left            =   735
               TabIndex        =   25
               Top             =   120
               Width           =   1845
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Asientos Contables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      Left            =   210
      TabIndex        =   7
      Top             =   315
      Width           =   4635
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Apertura y Cierre"
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
         Left            =   210
         TabIndex        =   16
         Top             =   4620
         Width           =   4215
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   17
            Top             =   420
            Width           =   3375
            Begin VB.CommandButton Command21 
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
               Left            =   560
               Picture         =   "CONTAB00.frx":1F26
               Style           =   1  'Graphical
               TabIndex        =   35
               ToolTipText     =   "Asiento de Apertura de Ejercicio"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command7 
               Enabled         =   0   'False
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
               Left            =   2760
               Picture         =   "CONTAB00.frx":2230
               Style           =   1  'Graphical
               TabIndex        =   19
               ToolTipText     =   "Asiento de Cierre de Ejercicio"
               Top             =   0
               Width           =   560
            End
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
               Picture         =   "CONTAB00.frx":253A
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Apertura de Nuevo Ejercicio Contable"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Apertura y Cierre de Ejercicio"
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
               Left            =   1260
               TabIndex        =   20
               Top             =   120
               Width           =   1530
            End
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mayorizar"
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
         Left            =   210
         TabIndex        =   12
         Top             =   3150
         Width           =   2010
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   1530
            TabIndex        =   13
            Top             =   420
            Width           =   1590
            Begin VB.CommandButton Command11 
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
               Picture         =   "CONTAB00.frx":2844
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Generar Asiento Resumen de Caja y Bancos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Resumen de C.y B."
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
               Left            =   630
               TabIndex        =   15
               Top             =   105
               Width           =   1005
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Asientos de Diario"
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
         Left            =   210
         TabIndex        =   8
         Top             =   420
         Width           =   4215
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   10
            Top             =   420
            Width           =   3375
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
               Picture         =   "CONTAB00.frx":298E
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Asiento Manual de Diario"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command14 
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
               Left            =   2760
               Picture         =   "CONTAB00.frx":2C98
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Asiento Resumen de Compras"
               Top             =   0
               Width           =   560
            End
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
               Left            =   2200
               Picture         =   "CONTAB00.frx":2FA2
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Asiento Resumen de Ventas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Registrar / Con- sultar As. Diario"
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
               Left            =   720
               TabIndex        =   11
               Top             =   120
               Width           =   1740
            End
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   9300
      Picture         =   "CONTAB00.frx":33E4
      Top             =   5900
      Width           =   2010
   End
End
Attribute VB_Name = "CONTAB00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
   Call CIERRACONTA
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    'Call BLOQARCH("ASICONT")
    TIASI% = 1
    ASICAJA.Show 1
    Unload ASICAJA
    Command10.Enabled = True
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    'Call BLOQARCH("ASICONT")
    MODOX% = ALTERNA%("Generar Asiento Resumen de Caja-Bancos\Borrar Asiento Resumen de C-B")
    '
    If MODOX% = 1 Then
        TIASI% = 3
        ASIDIAR.Show 1
        Unload ASIDIAR
      ElseIf MODOX% = 2 Then
        Call BORRARESCAJ
    End If
    '
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    Command12.Enabled = False
    RESCUEDIAR.Show 1
    Command12.Enabled = True
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    Call DIARIOGEN   ' Impresion Libro Diario General
    Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    'Call BLOQARCH("ASICONT")
    TIASI% = 2
    ASIDIAR.Show 1
    Unload ASIDIAR
    Command14.Enabled = True
End Sub

Private Sub Command15_Click()
    Command15.Enabled = False
    Call SUDIARCAJ   ' Impresion Subdiario de Caja y Bancos
    Command15.Enabled = True
End Sub

Private Sub Command17_Click()
    Command17.Enabled = False
    RESCUECAJA.Show 1
    Command17.Enabled = True
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
    Call SUCAJINEG    ' Subdiario de Caja (Ingresos - Egresos
    Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    OPBALANCE.Show 1
    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    'Call BLOQARCH("ASICONT")
    TIASI% = 0
    ASIDIAR.Show 1
    Unload ASIDIAR
    Command2.Enabled = True
End Sub

Private Sub Command20_Click()
    Command20.Enabled = False
    Call CONECRUN("AJUINFLA", "Ajuste por Inflación")
    Command20.Enabled = True
End Sub

Private Sub COMMAND21_CLICK()
    Command21.Enabled = False
    'Call BLOQARCH("ASICONT")
    TIASI% = 4
    ASIDIAR.Show 1
    Unload ASIDIAR
    Command21.Enabled = True
End Sub

Private Sub Command22_Click()
    Command22.Enabled = False
    'Call CONECRUN("#SCG21S", "Actualizacion del Plan General de Cuentas")
    PLAGRAL.Show 1
    Command22.Enabled = True
End Sub


Private Sub Command23_Click()
    Command23.Enabled = False
    GESVAL00.Show 1
    Command23.Enabled = True
End Sub


Private Sub Command24_Click()
   '
   ' control posicion diferida
   '
   Dim FEREGIS%(8192), IMPORE#(8192), REFERE$(8192)
   Dim CAMP%, CUEMP%(2048), MEPA%(2048)
   '
1 Call SELECHO("ECUECON")
   CUEX$ = VALACT1$("ECUECON")
   If CUEX$ = "" Then Exit Sub
   '
   CUE% = CUEINT%(CUEX$)
   If CUE% < 1 Then
      Call MENSERR(24, "Cuenta Inexistente")
      GoTo 1
   End If
   '
   Call FRESHALL
   CUEBIN$ = MKI$(CUE%)
   CODENOCUE$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
   JJ& = 0: TRANSPORTE# = 0
   SUMATO# = 0
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   Call CIERRACONTA
   SE$ = "0": EXTE$ = "OLD"
   JJ& = 0
   '
8 DES% = 0: HAS% = 32767
   '
   PMC$ = PUNMODIAR$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIA% = CVI(Mid$(PMC$, KK%, 2))
      Call LEEDIARIO(REGDIA%)
 If EXTE$ = "DAT" Then
   If CVS(E12$) = 1 Then
     GoTo 9    ' SALTEA APERTURA EJERCICIO 2002
   End If
 End If
      If Asc(E18$) <> 1 Then  ' SALTEA ASIENTOS RESUMEN DE CAJA
         Call LEEPASDIA(REGDIA%, 1)
         If TRIM$(DETAPAS$) = "" Then
            Call LEEPASDIA(REGDIA%, 2)
            DETAPAS$ = REFASIEN$
         End If
         MARCADIA$ = "*"
         If FECHASIEN% < DES% Then
             TRANSPORTE# = TRANSPORTE# + IMPOPAS#
             GoTo 9
           ElseIf FECHASIEN% <= HAS% Then
             GoSub 190
         End If
      End If
9 Next KK%
   '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   PMC$ = PUNMOCAJA$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIA% = CVI(Mid$(PMC$, KK%, 2))
      Call LEEPASCAJ(REGDIA%, 1)
      If TRIM$(DETAPAS$) = "" Then
         Call LEEPASCAJ(REGDIA%, 2)
         DETAPAS$ = REFASIEN$
      End If
      MARCADIA$ = ""
      If FECHASIEN% < DES% Then
          TRANSPORTE# = TRANSPORTE# + IMPOPAS#
          GoTo 11
        ElseIf FECHASIEN% <= HAS% Then
          GoSub 190
      End If
11 Next KK%
   '
   Call CIERRACONTA
   If Abs(TRANSPORTE#) >= 0.005 Then
      Call MENSERR(24, "Hay error en Transporte")
      Stop
   End If
   '
   If EXTE$ = "OLD" Then
     EXTE$ = "DAT"
     GoTo 8
   End If
   '
   AAA = SUMATO#
   Stop
   GoTo 100
   '
   '
190 JJ& = JJ& + 1
    FEREGIS%(JJ&) = FECHASIEN%
    IMPORE#(JJ&) = IMPOPAS#
    REFERE$(JJ&) = MARCADIA$ + DETAPAS$
    SUMATO# = SUMATO# + IMPOPAS#
    Return
    '
   '
   '
   '
   '
   'REVISA POSICION DIFERIDA
   '
100 Dim REFE$(2048)
    Dim FECHAMO%(2048)
    Dim IPODIF#(2048)
    '
    JJ1& = 0
    EXTE$ = "DAT"
   '
   If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUEX$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUEX$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUEX$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
   End If
   '
   HOII% = HOY(HOS$)
   DES% = FECHANUM("01/10/01")
   HAS% = FECHANUM("31/12/02")
   '
10 Screen.MousePointer = 11
   KLM& = 0
   '
   CUEBIN$ = MKI$(CUE%)
   CODENOCUE$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
   JJ1& = 0
   SALINI# = 0
   '
   If DES% >= FECHANUM("01/10/01") Then
      CCCC$ = TRIM$(CODCUE$(CUE%))
      If CCCC$ = "111401" Then SALINI# = 11409.33
      If CCCC$ = "111502" Then SALINI# = 11791.1
      If CCCC$ = "111403" Then SALINI# = 2414.3
      If CCCC$ = "111406" Then SALINI# = 29769.7
      If CCCC$ = "111506" Then SALINI# = 381.86
   End If
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   EXTEA$ = EXTE$
   If EXTEA$ = "DAT" Then
     Close #NX9%: NX9% = 0
     Close #NX10%: NX10% = 0
     Close #NX11%: NX11% = 0
     EXTE$ = "OLD"
   End If
   '
12 PMC$ = PUNMODIAR$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIA% = CVI(Mid$(PMC$, KK%, 2))
      Call LEEDIARIO(REGDIA%)
      If CVS(E12$) < 1.05 Then GoTo 19 ' SALTEA ASIENTO APERTURA
      If Asc(E18$) <> 1 Then  ' SALTEA ASIENTOS RESUMEN DE CAJA
         Call LEEPASDIA(REGDIA%, 1)
         If TRIM$(DETAPAS$) = "" Then
            Call LEEPASDIA(REGDIA%, 2)
            DETAPAS$ = REFASIEN$
         End If
         MARCADIA$ = "*"
         FEMOREA% = CVI(E15$)
         IMPO# = CVD(E16$)
         If FEMOREA% < DES% Then
             If FEMOREA% > FECHANUM("31/09/01") Then GoSub 90 'TOTALIZAR INICIAL
          ElseIf FEMOREA% <= HAS% Then
            FEMORE% = FEMOREA%
            FEREGI% = FEMOREA%
            FEMIS% = FEMOREA%
            GoSub 95 'OK
         End If
      End If
19 Next KK%
   '
   If EXTE$ = "OLD" Then
     If EXTEA$ = "DAT" Then
       Close #NX9%: NX9% = 0
       Close #NX10%: NX10% = 0
       Close #NX11%: NX11% = 0
       EXTE$ = "DAT"
       EXTEA$ = ""
       GoTo 12
     End If
   End If
   '
   '
20 ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   EXTEA$ = EXTE$
   If EXTEA$ = "DAT" Then
     Close #NX5%: NX5% = 0
     Close #NX6%: NX6% = 0
     Close #NX7%: NX7% = 0
     EXTE$ = "OLD"
   End If
   '
22 PMC$ = PUNMOCAJA$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIA% = CVI(Mid$(PMC$, KK%, 2))
      Call LEEPASCAJ(REGDIA%, 1)
      FEMOVI% = CVI(D15$)
      IMPO# = CVD(D16$)
      '
      If TRIM$(DETAPAS$) = "" Then
         Call LEEPASCAJ(REGDIA%, 2)
         DETAPAS$ = REFASIEN$
      End If
      MARCADIA$ = ""
      If InStr(DETAPAS$, "- OP") > 0 Then GoTo 29
      If Left$(DETAPAS$, 8) = "DEPOSITO" Then
         FEMORE% = FECHACRE%(FEMOVI%, Val(Mid$(DETAPAS$, 10, 3)))
         If FEMORE% < DES% Then
            If FEMORE% > FECHANUM("31/09/01") Then GoSub 90
          ElseIf FEMORE% <= HAS% Then
 'FEMORE% = FEMOVI%    ' para poder comparar
            FEMIS% = FEMOVI%
            GoSub 95    ' ok
         End If
         GoTo 29
      End If
      If FEMOVI% < DES% Then
        If FEMOVI% > FECHANUM("31/09/01") Then GoSub 90
        GoTo 29
      End If
      If FEMOVI% >= DES% Then
        If FEMOVI% <= HAS% Then
          FEMORE% = FEMOVI%
          FEMIS% = FEMOVI%
          GoSub 95  'ok
        End If
      End If
      '
29 Next KK%
   '
   If EXTE$ = "OLD" Then
     If EXTEA$ = "DAT" Then
       Close #NX5%: NX5% = 0
       Close #NX6%: NX6% = 0
       Close #NX7%: NX7% = 0
       EXTE$ = "DAT"
       EXTEA$ = ""
       GoTo 22
     End If
   End If
   '
   For U& = 1 To ULTREG&("CHEQUEMI")
     X$ = REGLEIDO$("CHEQUEMI", U&)
      PROV% = CVI(Left$(X$, 2))
      NUMERO$ = Mid$(X$, 3, 12)
      BANCO$ = Mid$(X$, 15, 16)
      EMIS% = CVI(Mid$(X$, 47, 2)) ' toma fecha de O/Pago
      IMPO# = (-1) * CVD(Mid$(X$, 33, 8))
      NOPAGO& = CVS(Mid$(X$, 41, 4))
      VENCI% = CVI(Mid$(X$, 49, 2))
      QUE% = CVI(Mid$(X$, 45, 2))
            For UX% = 1 To CAMP%
               If MEPA%(UX%) = QUE% Then GoTo 38
            Next UX%
            Call MENSERR(24, "Medio de Pago" + Str$(QUE%) + " no Definido")
            Call FINAL("")
            '
38    CUEH% = CUEMP%(UX%)
      If CUEH% = CUE% Then
        FEMORE% = VENCI%      ' TOMA FECHA DE EMISION - CAMBIE ACA
        FEMIS% = EMIS%
        If FEMORE% < DES% Then
          If FEMORE% > FECHANUM("31/09/01") Then GoSub 90
          'GoSub 90
        End If
        If FEMORE% >= DES% Then
          If FEMORE% <= HAS% Then
           FEX = EMIS%
            DETAPAS$ = "Ch." + TRIM$(NUMERO$)
            DETAPAS$ = DETAPAS$ + " - OP." + TRIM$(Str$(NOPAGO&))
            DETAPAS$ = DETAPAS$ + " - " + TRIM$(RASOCLI$((-1) * PROV%))
            VINODEOPA% = 1
            GoSub 95 ' OK
            VINODEOPA% = 0
          End If
        End If
      End If
   Next U&
   '
   FEX = DES%
   FEMORE% = DIANTE(FEX)
   DETAPAS$ = "Saldo al " + FECHATEX$(FEX)
   IMPO# = SALINI#
   FEMIS% = FEMORE%
   GoSub 95
   '
AAA = SUPODIF#
   '
SUDIFER# = 0
111 For KK11& = 1 To JJ&
If Abs(IMPORE#(KK11&)) >= 0.005 Then
  SUDIFER# = SUDIFER# + IMPORE#(KK11&)
  MsgBox REFERE$(KK11&) & IMPORE#(KK11&)
  TTT1& = TTT1& + 1
End If
Next KK11&
Stop
   '
   '
   '
   '
   '
90 SALINI# = SALINI# + IMPO#
   Return
   '
95 If Abs(IMPO#) >= 0.005 Then
      If VINODEOPA% = 1 Then
        If FEMIS% < DES% Then
          GoSub 90     ' DESCONTAR DE SALDO INICIAL
          GoTo 951
        End If
      End If
      '
      SUPODIF# = SUPODIF# + IMPO#
      For KK11& = 1 To JJ&
        If FEREGIS%(KK11&) = FEMIS% Then
          If Abs(IMPORE#(KK11&) - IMPO#) < 0.005 Then
            IMPORE#(KK11&) = IMPORE#(KK11&) - IMPO#
            GoTo 951
          End If
        End If
      Next KK11&
      '
      JJ& = JJ& + 1
      FEREGIS%(JJ&) = FEMORE%
      IMPORE#(JJ&) = (-1) * IMPO#
      REFERE$(JJ&) = MARCADIA$ + DETAPAS$
951 End If
   Return
   '

End Sub


Private Sub Command26_Click()
   OPCHEQ00.Show 1
End Sub

Private Sub Command27_Click()
    Call ACCDIR("CONTAGEN")
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONTAGEN")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    EPAC$ = EMPREAC$
    If InStr(EPAC$, "CIME-SOI") < 1 Then
         Call OPNOIN("Informes Normalizados")
       Else
         Call CONECRUN("#SCG47B", "Informes Standard Normalizados")
    End If
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    'Call BLOQARCH("ASICONT")
    TIASI% = 0
    ASICAJA.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    'Call BLOQARCH("ASICONT")
    TIASI% = 3
    ASICAJA.Show 1
    Unload ASICAJA
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    If DERACCE%("ABREJER", "Apertura de Ejercicio Contable") = 2 Then
       Call ABREJER    ' Apertura de Nuevo Ejercicio Contable
    End If
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    Call OPNOIN("Asiento de Cierre de Ejercicio")
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    CUECHECA$ = CONTROL$("", "CUECHECA")
    CUEICHEQ% = CUEINT%(CUECHECA$)
    If TRIM$(CUECHECA$) = "" Or CUEICHEQ% < 1 Or CUEICHEQ% > NUCUEN% Then
       Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores")
       OPCHEQ00.Show 1
       GoTo 99
    End If
    '
    'Call BLOQARCH("ASICONT")
    TIASI% = 2
    ASICAJA.Show 1
    Unload ASICAJA
99  Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    'Call BLOQARCH("ASICONT")
    TIASI% = 1
    ASIDIAR.Show 1
    Unload ASIDIAR
    Command9.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    Caption = Caption + " - " + TRIM$(EJACT$)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
End Sub

Sub DIARIOGEN()
   '
   Call FRESHALL
   Call CARSELMESA
   MODOEMI% = COMALTER%("Puede Enitir el Libro Diario General\en Forma Mensual o General del Ejercicio\\Mensual\General")
   If MODOEMI% < 1 Or MODOEMI% > 2 Then Exit Sub
   '
10 DES% = COMEJ%: HAS% = FINEJ%
   CORRESP$ = EJACT$
   If MODOEMI% = 1 Then
      Call SELECHO("SELMESA")
      MESASEL$ = VALACT1$("SELMESA")
      If TRIM$(MESASEL$) = "" Then
         Exit Sub
      End If
      '
      MMXX0$ = TRIM$(MESASEL$)
      MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
      If TRIM$(MMXX1$) = "" Then
         Exit Sub
      End If
      '
20    MESASEL$ = MMXX1$
      If DES% < COMEJ% Or HAS% > FINEJ% Then
         Call MENSERR(24, "Período Fuera de Límites\del Ejercicio Activo.")
         GoTo 10
      End If
      '
      CORRESP$ = MESASEL$
   End If
   '
   MODOXX2% = COMALTER%("Puede Incluir los Asientos de Caja y Bancos\o Dejar Solo los Asientos Resumen\\Asientos de C & B\Asientos Resumen")
   If MODOXX2% < 1 Or MODOXX2% > 2 Then GoTo 99
   Call GRACONTROL("CONTAGEN", "INCLUCAJ", Str$(MODOXX2% Mod 2))
   MODOXX3% = COMALTER%("Puede Re-Numerar Correlativamente los Asientos\o Expresar los Números Provisorios de Asientos\\Re-Numerar Asientos\Números Provisorios")
   If MODOXX3% < 1 Or MODOXX3% > 2 Then GoTo 99
   Call GRACONTROL("CONTAGEN", "RENUMERA", Str$(MODOXX3% Mod 2))
   '
   Screen.MousePointer = 11
   SUDEBE# = 0
   SUHABE# = 0
   SORTLIST.Clear
   '
   For KK% = 1 To MAXDIARIO%
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      Call LEEPASDIA(KK%, 1)
      FEAS% = CVI(E15$)
      If MODOXX2% = 1 Then
        If Asc(E18$) = 1 Then GoTo 29 ' SALTEA ASIENTOS RESUMEN DE CAJA
      End If
      RESOR$ = Left$(TRIM$(Str$(10000 + FEAS%)) + Space$(10), 10)
      RESOR$ = RESOR$ + "DI" + TRIM$(Str$(10000 + KK%))
      SORTLIST.AddItem RESOR$
29 Next KK%
   '
   If MODOXX2% = 1 Then
     For KK% = 1 To MAXDIA%
       UU& = KK%
       Call TRACE(20, UU&, FIN&)
       Call LEEPASCAJ(KK%, 1)
       FEAS% = CVI(D15$)
       RESOR$ = Left$(TRIM$(Str$(10000 + FEAS%)) + Space$(10), 10)
       RESOR$ = RESOR$ + "SU" + TRIM$(Str$(10000 + KK%))
       SORTLIST.AddItem RESOR$
     Next KK%
   End If
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   FIN& = SORTLIST.ListCount    ' MAXDIARIO%
   YATRA% = 0: JJ& = 0: ASIANT = 0: ASIANTX = 0
   '
   NUASDEF& = 0
   If MODOXX3% = 1 Then
      Screen.MousePointer = 1
      NUASDEF& = Val(InputBox("Re-Numera a Partir de", "Numeración Asientos Libro Diario", "1.0")) - 1
      If NUASDEF& < 0 Then Exit Sub
      Screen.MousePointer = 11
   End If
   '
   For KK% = 1 To FIN&     'MAXDIARIO%
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      RESOR$ = SORTLIST.List(KK% - 1)
      REGII% = Val(Mid$(RESOR$, 13)) - 10000
      GRUXI$ = TRIM$(Mid$(RESOR$, 11, 2))
      SGX1% = 1
      FEASI% = Val(TRIM$(Left$(RESOR$, 10))) - 10000
      MODOLE% = 1
      If FEASI% >= DES% And FEASI% <= HAS% Then MODOLE% = 2
      If GRUXI$ = "DI" Then
         Call LEEPASDIA(REGII%, MODOLE%)
         FEAS% = CVI(E15$)
         IPAS# = CVD(E16$)
         CUE% = CVI(E11$)
         NUAS = CVS(E12$)
         REFEL$ = "- " + TRIM$(DETAPAS$)
         REFE$ = REPLACAR$(REFEL$, "@", "A")
         If TRIM$(REFE$) = "-" Then
           REFE$ = "- " + TRIM$(DENOCUE$(CUE%))
         End If
         DENOMI$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
        Else
         SGX1% = (-1)
         Call LEEPASCAJ(REGII%, MODOLE%)
         FEAS% = CVI(D15$)
         IPAS# = CVD(D16$)
         CUE% = CVI(D11$)
         NUAS = CVS(D12$)
         REFEL$ = "- " + TRIM$(DETAPAS$)
         REFE$ = REPLACAR$(REFEL$, "@", "A")
         If TRIM$(REFE$) = "-" Then
           REFE$ = "- " + TRIM$(DENOCUE$(CUE%))
         End If
         DENOMI$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
      End If
      '
      If NUAS * SGX1% <> ASIANT Then
        NUASDEF& = NUASDEF& + 1
        ASIANT = NUAS * SGX1%
      End If
      '
      NUADE% = NUASDEF& Mod 32700
      If SGX1% > 0 Then
          Get #NX10%, REGII%, E10$
          EXXX$ = E10$
          Call REPLA(EXXX$, MKI$(NUADE%), 7, 2)
          E10$ = EXXX$
          Put #NX10%, REGII%, E10$
        Else
          Get #NX6%, REGII%, D10$
          EXXX$ = D10$
          Call REPLA(EXXX$, MKI$(NUADE%), 7, 2)
          D10$ = EXXX$
          Put #NX6%, REGII%, D10$
      End If
      '
      If FEAS% < DES% Then
         If IPAS# > 0 Then
              SUDEBE# = SUDEBE# + IPAS#
            Else
              SUHABE# = SUHABE# - IPAS#
         End If
         GoTo 90
      End If
      '
      If YATRA% = 0 Then
         Z$ = REGBLAN$("DIARGEN")
         FEX = DES%
         DIA0% = DIANTE(FEX)
         Call REPLA(Z$, MKI$(DIA0%), 1, 2)
         Call REPLA(Z$, MKS$(99999999999#), 3, 4)
         Call REPLA(Z$, "Transporte", 7, 53)
         Call REPLA(Z$, MKD$(SUDEBE#), 105, 8)
         Call REPLA(Z$, MKD$(SUHABE#), 113, 8)
         JJ& = JJ& + 1
         Call GRAREG("DIARGEN", Z$, JJ&)
         Z$ = REGBLAN$("DIARGEN")
         JJ& = JJ& + 1
         Call GRAREG("DIARGEN", Z$, JJ&)
         YATRA% = 1
      End If
      '
      If FEAS% <= HAS% Then
        If NUAS * SGX1% <> ASIANTX Then
           ASIANTX = NUAS * SGX1%
           Z$ = REGBLAN$("DIARGEN")
           Call REPLA(Z$, MKI$(FEAS%), 1, 2)
           Call REPLA(Z$, MKS$(NUAS), 3, 4)
           If MODOXX3% = 1 Then
             Call REPLA(Z$, MKS$(NUASDEF&), 3, 4)
           End If
           Call REPLA(Z$, REFASIEN$, 7, 53)
           JJ& = JJ& + 1
           Call GRAREG("DIARGEN", Z$, JJ&)
         End If
         '
         Z$ = REGBLAN$("DIARGEN")
         Call REPLA(Z$, MKI$(FEAS%), 1, 2)
         Call REPLA(Z$, MKS$(NUAS), 3, 4)
         If MODOXX3% = 1 Then
            Call REPLA(Z$, MKS$(NUASDEF&), 3, 4)
         End If
         Call REPLA(Z$, REFE$, 7, 53)
         Call REPLA(Z$, DENOMI$, 60, 45)
         If IPAS# > 0 Then
              Call REPLA(Z$, MKD$(IPAS#), 105, 8)
            Else
              Call REPLA(Z$, MKD$((-1) * IPAS#), 113, 8)
         End If
         JJ& = JJ& + 1
         Call GRAREG("DIARGEN", Z$, JJ&)
      End If
90 Next KK%
   '
   Call SETULTREG("DIARGEN", JJ&)
   Call CIERRARCH("DIARGEN")
   '
   Call HEADERS("DIARGEN", "")
   Call HEADERS("DIARGEN", "Corresponde a: " + CORRESP$)
   '
   Screen.MousePointer = 1
   Call CONECRUN("*SUBDIAR", "Libro Diario General")
   '
99 End Sub

Sub SUDIARCAJ()
   '
   Call FRESHALL
   Call CARSELMESA
   MODOEMI% = COMALTER%("Puede Enitir el Subdiario de Caja y Bancos\en Forma Mensual o General del Ejercicio\\Mensual\General")
   If MODOEMI% < 1 Or MODOEMI% > 2 Then Exit Sub
   '
10 DES% = COMEJ%: HAS% = FINEJ%
   CORRESP$ = EJACT$
   If MODOEMI% = 1 Then
      Call SELECHO("SELMESA")
      MESASEL$ = VALACT1$("SELMESA")
      If TRIM$(MESASEL$) = "" Then
         Exit Sub
      End If
      '
      MMXX0$ = TRIM$(MESASEL$)
      MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
      If TRIM$(MMXX1$) = "" Then
         Exit Sub
      End If
      '
20    MESASEL$ = MMXX1$
      If DES% < COMEJ% Or HAS% > FINEJ% Then
         Call MENSERR(24, "Período Fuera de Límites\del Ejercicio Activo.")
         GoTo 10
      End If
      '
      CORRESP$ = MESASEL$
   End If
   '
   Screen.MousePointer = 11
   SUDEBE# = 0
   SUHABE# = 0
   '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   FIN& = MAXDIA%
   YATRA% = 0: JJ& = 0: ASIANT = 0
   For KK% = 1 To MAXDIA%
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      Call LEEPASCAJ(KK%, 1)
      FEAS% = CVI(D15$)
      IPAS# = CVD(D16$)
      '
      If FEAS% < DES% Then
         If IPAS# > 0 Then
              SUDEBE# = SUDEBE# + IPAS#
            Else
              SUHABE# = SUHABE# - IPAS#
         End If
         GoTo 90
      End If
      '
      If YATRA% = 0 Then
         Z$ = REGBLAN$("SUDIACAJ")
         FEX = DES%
         DIA0% = DIANTE(FEX)
         Call REPLA(Z$, MKI$(DIA0%), 1, 2)
         Call REPLA(Z$, MKS$(99999999999#), 3, 4)
         Call REPLA(Z$, "Transporte", 7, 53)
         Call REPLA(Z$, MKD$(SUDEBE#), 105, 8)
         Call REPLA(Z$, MKD$(SUHABE#), 113, 8)
         JJ& = JJ& + 1
         Call GRAREG("SUDIACAJ", Z$, JJ&)
         Z$ = REGBLAN$("SUDIACAJ")
         JJ& = JJ& + 1
         Call GRAREG("SUDIACAJ", Z$, JJ&)
         YATRA% = 1
      End If
      '
      If FEAS% <= HAS% Then
         Call LEEPASCAJ(KK%, 2)
         FEAS% = CVI(D15$)
         IPAS# = CVD(D16$)
         CUE% = CVI(D11$)
         NUAS = CVS(D12$)
         '
         If NUAS <> ASIANT Then
           Z$ = REGBLAN$("SUDIACAJ")
           Call REPLA(Z$, MKI$(FEAS%), 1, 2)
           Call REPLA(Z$, MKS$(NUAS), 3, 4)
           Call REPLA(Z$, REFASIEN$, 7, 53)
           JJ& = JJ& + 1
           Call GRAREG("SUDIACAJ", Z$, JJ&)
           ASIANT = NUAS
         End If
         '
         REFEL$ = "- " + TRIM$(DETAPAS$)
         REFE$ = REPLACAR$(REFEL$, "@", "A")
         If TRIM$(REFE$) = "-" Then
           REFE$ = "- " + TRIM$(DENOCUE$(CUE%))
         End If
         DENOMI$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
         '
         Z$ = REGBLAN$("SUDIACAJ")
         Call REPLA(Z$, MKI$(FEAS%), 1, 2)
         Call REPLA(Z$, MKS$(NUAS), 3, 4)
         Call REPLA(Z$, REFE$, 7, 53)
         Call REPLA(Z$, DENOMI$, 60, 45)
         If IPAS# > 0 Then
              Call REPLA(Z$, MKD$(IPAS#), 105, 8)
            Else
              Call REPLA(Z$, MKD$((-1) * IPAS#), 113, 8)
         End If
         JJ& = JJ& + 1
         Call GRAREG("SUDIACAJ", Z$, JJ&)
      End If
90 Next KK%
   '
   Call SETULTREG("SUDIACAJ", JJ&)
   Call CIERRARCH("SUDIACAJ")
   '
   Call HEADERS("SUDIACAJ", "")
   Call HEADERS("SUDIACAJ", "Corresponde a: " + CORRESP$)
   '
   Screen.MousePointer = 1
   Call CONECRUN("*SUBCAJA", "Libro Subdiario de Caja y Bancos")
   '
End Sub
'
Sub SUCAJINEG()           ' Subdiario de Caja - Ingresos y Egresos
   '
   Dim NASISEL(32767)
   Call FRESHALL
   Screen.MousePointer = 11
   CCCX% = CUEINT%("")          ' para determinar NUCUEN%
   HAS% = HOY(HO$): If HAS% > FINEJ% Then HAS% = FINEJ%
   FEX = HAS%: HASX$ = FECHATEX$(FEX)
   DESX$ = "01" + Mid$(HASX$, 3)
   DES% = FECHANUM(DESX$)
   For I1% = 1 To NUCUEN%
      If InStr(UCase$(DENOCUE$(I1%)), "CAJA") > 0 Then
         CUEI% = I1%
         COCUE$ = CODCUE$(CUEI%)
         GoTo 5
      End If
   Next I1%
   '
5  VDEF$ = MKI$(DES%) + MKI$(HAS%) + COCUE$
   Screen.MousePointer = 1
   '
6  OBX$ = OBJPLA$("DESHASFECHCUEN", VDEF$)
   If OBX$ = "" Then Exit Sub
   VDEF$ = OBX$
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   COCUE$ = TRIM$(Mid$(OBX$, 5, 12))
   If DES% > HAS% Then GoTo 6
   If DES% < COMEJ% Then GoTo 6
   If HAS% > FINEJ% Then GoTo 6
   If COCUE$ = "" Then GoTo 6
   CUEI% = CUEINT%(COCUE$)
   If CUEI% < 1 Or CUEI% > NUCUEN% Then GoTo 6
   '
   FEX = DES%: DESX$ = FECHATEX$(FEX)
   FEX = HAS%: HASX$ = FECHATEX$(FEX)
   PERIO$ = "del " + DESX$ + " al " + HASX$
   '
8  MODOEMI% = COMALTER%("Puede Enitir el Subdiario\por Ingresos o Egresos\\Ingresos\Egresos\Cancelar")
   If MODOEMI% < 1 Or MODOEMI% > 2 Then GoTo 6
   '
   SIGNO% = 1: INEG$ = "Ingresos"
   If MODOEMI% = 2 Then SIGNO% = (-1): INEG$ = "Egresos"
   '
   CORRESP$ = TRIM$(COCUE$) + " " + TRIM$(DENOCUE$(CUEI%)) + " (" + INEG$ + ") " + PERIO$
   '
   'Identificar asientos con movimiento
   '
   Screen.MousePointer = 11
   FIN& = MAXDIA%: PRIMO% = 0: ULTMO% = 0
   YATRA% = 0: NUASANT = 0: CASISEL% = 0
   For KK% = 1 To MAXDIA%
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      Call LEEPASCAJ(KK%, 1)
      CUE% = CVI(D11$)
      NUAS = CVS(D12$)
      FEAS% = CVI(D15$)
      IPAS# = CVD(D16$)
      '
      If FEAS% >= DES% Then
        If FEAS% <= HAS% Then
          If PRIMO% = 0 Then PRIMO% = KK%
          ULTMO% = KK%
          If CUE% = CUEI% Then
            If Sgn(IPAS#) = SIGNO% Then
              If NUAS <> NUASANT Then
                CASISEL% = CASISEL% + 1
                NASISEL(CASISEL%) = NUAS
                NUASANT = NUAS
              End If
            End If
          End If
        End If
      End If
   Next KK%
   '
   If CASISEL% < 1 Then
      Call COMUNI("Sin Movimientos en el Período Elegido")
      GoTo 8
   End If
   '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   FIN& = ULTMO% - PRIMO% + 1
   NUASANT = 0: ACUM# = 0
   YATRA% = 0: JJ& = 0
   For KK% = PRIMO% To ULTMO%
      UU& = KK% - PRIMO% + 1
      Call TRACE(20, UU&, FIN&)
      Call LEEPASCAJ(KK%, 1)
      CUE% = CVI(D11$)
      NUAS = CVS(D12$)
      '
      If NUAS <> NUASANT Then
        SELE% = 0
        For UI% = 1 To CASISEL%
          If NASISEL(UI%) = NUAS Then
            SELE% = 1
            GoTo 20
          End If
        Next UI%
      End If
      '
20    If SELE% = 1 Then
        Z$ = REGBLAN$("SUDINEG")
        If NUAS <> NUASANT Then
           Call LEEPASCAJ(KK%, 2)
           FEX = CVI(D15$)
           Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
           Call REPLA(Z$, MKS$(NUAS), 9, 4)
           REFEL$ = REFASIEN$
           REFE$ = REPLACAR$(REFEL$, "@", "A")
           Call REPLA(Z$, REFE$, 13, 48)
           JJ& = JJ& + 1
           Call GRAREG("SUDINEG", Z$, JJ&)
           Z$ = REGBLAN$("SUDINEG")
        End If
        '
        IPAS# = CVD(D16$)
        REFEL$ = "- " + TRIM$(DETAPAS$)
        REFE$ = REPLACAR$(REFEL$, "@", "A")
        Call REPLA(Z$, REFE$, 13, 48)
        Call REPLA(Z$, TRIM$(CODCUE$(CUE%)), 61, 12)
        If IPAS# > 0 Then
             Call REPLA(Z$, MKD$(IPAS#), 73, 8)
           Else
             Call REPLA(Z$, MKD$(Abs(IPAS#)), 81, 8)
        End If
        If CUE% = CUEI% Then
           ACUM# = ACUM# + Abs(IPAS#)
           Call REPLA(Z$, MKD$(ACUM#), 89, 8)
        End If
        JJ& = JJ& + 1
        Call GRAREG("SUDINEG", Z$, JJ&)
        '
        NUASANT = NUAS
        '
      End If
      '
   Next KK%
   '
   Call SETULTREG("SUDINEG", JJ&)
   Call CIERRARCH("SUDINEG")
   '
   Call HEADERS("SUDINEG", "")
   Call HEADERS("SUDINEG", "Cuenta: " + CORRESP$)
   '
   Screen.MousePointer = 1
   Call CONECRUN("*SUDINEG", "Subdiario de Caja (Ingresos y Egresos)")
   '
   GoTo 6
   '
End Sub
'
Sub ABREJER()
   '
   Dim ARQ$(16)
   '
   If EJER% <> 1 Then
      Call MENSERR(24, "Apertura Imposible\Ingrese al Ejercicio Actual")
      Exit Sub
   End If
   '
   TTXXX$ = "Al Abrir un Nuevo Ejercicio Contable\el Ejercicio ACTUAL pasa a ser ANTERIOR\y los Datos del ANTERIOR se Pierden !!!\   \Desea Continuar o Cancelar Apertura ?"
   If COMALTER%(TTXXX$ + "\\Cancelar Apertura\Continuar") <> 2 Then
      Exit Sub
   End If
   '
   TTXXX$ = "Para su Propia Seguridad de sus Datos\es IMPRESCINDIBLE Haber Confeccionado un BACKUP\Reciente de los Archivos '*.OLD'\de su Base de Datos !!!\   \Sólo si Tiene ese Backup en Lugar Seguro\Pinche sobre 'Si, Continuar', Caso Contrario sobre 'No, Cancelar'"
   If COMALTER%(TTXXX$ + "\\No, Cancelar\Si, Continuar") <> 2 Then
      Exit Sub
   End If
   '
   VDEF$ = String$(80, 0)
   Call REPLA(VDEF$, MKI$(COMEJ%), 1, 2)
   Call REPLA(VDEF$, EJACT$, 3, 32)
   FEX = FINEJ%
   CMJ1% = DIPOST(FEX)
   Call REPLA(VDEF$, MKI$(CMJ1%), 41, 2)
   Call REPLA(VDEF$, "", 43, 32)
   '
10 OBX$ = OBJPLA$("ABREJER", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   VDEF$ = OBX$
   DENUE$ = TRIM$(Mid$(OBX$, 43, 32))
   If DENUE$ = "" Then
      Call MENSERR(24, "Falta Denominación del Nuevo Ejercicio")
      GoTo 10
   End If
   '
   CMJACT% = COMEJ%
   APNUE% = CVI(Mid$(OBX$, 41, 2))
   If APNUE% < CMJACT% Then
      Call MENSERR(24, "Nueva Apertura es Anterior\al Cierre del Ejercicio Actual")
      GoTo 10
   End If
   '
   FEX = APNUE%
   APNU$ = FECHATEX$(FEX)
   ANX% = Val(Mid$(APNU$, 7, 2)) + 1: If ANX% > 99 Then ANX% = ANX% - 100
   AN$ = TRIM$(Str$(ANX%)): While Len(AN$) < 2: AN$ = "0" + AN$: Wend
   FE$ = Left$(APNU$, 6) + AN$
   FEX = FECHANUM(FE$)
   CIENUE% = DIANTE(FEX)
   '
   ' MENSAJE DE ADVERTENCIA y confirmacion
   '
   Screen.MousePointer = 11
   LU$ = LUDAT$
   NSE% = 1           ' numero de subempresas
   On Error Resume Next
   Kill LU$ + "*.NEW"
   On Error GoTo 0
   '
   N1% = FILEOPEN%(LU$ + "VARIAB.DAT", 0, 32)
   Dim X As String * 32
   If APNUE% = CMJACT% Then
      Get #N1%, 2, X$
      LSet X$ = Left$(DENUE$ + Space$(28), 28) + MKI$(APNUE%) + MKI$(CIENUE%)
      Put #N1%, 2, X$
      Close #N1%
      GoTo 99
   End If
   '
   N2% = FILEOPEN%(LU$ + "VARIAB.NEW", 0, 32)
   Dim Y As String * 32
   '
   For I% = 1 To LOF(N1%) / 32
      Get #N1%, I%, X$
      Z$ = X$: If I% > 10 Then Z$ = String$(32, 0)
      LSet Y$ = Z$
      Put #N2%, I%, Y$
   Next I%
   '
   Get #N2%, 2, Y$: Put #N2%, 3, Y$
   LSet Y$ = Left$(DENUE$ + Space$(28), 28) + MKI$(APNUE%) + MKI$(CIENUE%)
   Put #N2%, 2, Y$
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "SUBEMPRE.DAT", 0, 64)
   Dim X64 As String * 64
   N2% = FILEOPEN%(LU$ + "SUBEMPRE.NEW", 0, 64)
   Dim Y64 As String * 64
   '
   For I% = 1 To LOF(N1%) / 64
      Get #N1%, I%, X64$
      Z$ = X64$
      LSet Y64$ = Z$
      Put #N2%, I%, Y64$
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "CUENTAS.DAT", 0, 64)
   N2% = FILEOPEN%(LU$ + "CUENTAS.NEW", 0, 64)
   '
   For I% = 1 To LOF(N1%) / 64
      Get #N1%, I%, X64$
      Z$ = X64$
      LSet Y64$ = Z$
      Put #N2%, I%, Y64$
      NUCUENUE% = I%
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "DENOCUE.DAT", 0, 128)
   Dim X128 As String * 128
   N2% = FILEOPEN%(LU$ + "DENOCUE.NEW", 0, 128)
   Dim Y128 As String * 128
   '
   For I% = 1 To LOF(N1%) / 128
      Get #N1%, I%, X128
      Z$ = X128$
      LSet Y128$ = Z$
      Put #N2%, I%, Y128
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "INVECUE.DAT", 0, 16)
   Dim X16 As String * 16
   N2% = FILEOPEN%(LU$ + "INVECUE.NEW", 0, 16)
   Dim Y16 As String * 16
   '
   For I% = 1 To LOF(N1%) / 16
      Get #N1%, I%, X16
      Z$ = X16$
      LSet Y16$ = Z$
      Put #N2%, I%, Y16$
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "PLAGRAL.DAT", 0, 16)
   N2% = FILEOPEN%(LU$ + "PLAGRAL.NEW", 0, 16)
   '
   For I% = 1 To LOF(N1%) / 16
      Get #N1%, I%, X16$
      Z$ = X16$
      LSet Y16$ = Z$
      Put #N2%, I%, Y16$
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "PLANALT.DAT", 0, 16)
   N2% = FILEOPEN%(LU$ + "PLANALT.NEW", 0, 16)
   '
   For I% = 1 To LOF(N1%) / 16
      Get #N1%, I%, X16$
      Z$ = X16$
      LSet Y16$ = Z$
      Put #N2%, I%, Y16$
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "TRADENO.DAT", 0, 128)
   N2% = FILEOPEN%(LU$ + "TRADENO.NEW", 0, 128)
   '
   For I% = 1 To LOF(N1%) / 128
      Get #N1%, I%, X128$
      Z$ = X128$
      LSet Y128$ = Z$
      Put #N2%, I%, Y128$
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "DENENCA.DAT", 0, 128)
   N2% = FILEOPEN%(LU$ + "DENENCA.NEW", 0, 128)
   '
   For I% = 1 To LOF(N1%) / 128
      Get #N1%, I%, X128$
      Z$ = X128$
      LSet Y128$ = Z$
      Put #N2%, I%, Y128$
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "INDAJU.DAT", 0, 132)
   Dim X132 As String * 132
   N2% = FILEOPEN%(LU$ + "INDAJU.NEW", 0, 132)
   Dim Y132 As String * 132
   '
   For I% = 1 To LOF(N1%) / 132
      Get #N1%, I%, X132$
      Z$ = X132$
      LSet Y132$ = Z$
      Put #N2%, I%, Y132$
   Next I%
   Close #N2%: Close #N1%
   '
   N2% = FILEOPEN%(LU$ + "SALCAJA.NEW", 0, 80)
   Dim Y80 As String * 80
   '
   For I% = 1 To NUCUENUE%
      LSet Y80$ = String$(80, 45)
      Put #N2%, I%, Y80$
   Next I%
   Close #N2%: Close #N1%
   '
   Dim X112 As String * 112
   Dim Y112 As String * 112
   '
   For SX% = 1 To NSE%
      '
      SXX$ = Mid$(Str$(SX% - 1), 2)
      N1% = FILEOPEN%(LU$ + "CIERRE" + SXX$ + ".NEW", 0, 112)
      '
      For I% = 1 To NUCUENUE%
         LSet Y112$ = String$(112, 0)
         Put #N1%, I%, Y112$
      Next I%
      Close #N1%
      '
      N1% = FILEOPEN%(LU$ + "ENCDIAR" + SXX$ + ".NEW", 0, 64)
      LSet Y64$ = String$(64, 0)
      Put #N1%, 1, Y64$
      Close #N1%
      '
      N1% = FILEOPEN%(LU$ + "DIARIO" + SXX$ + ".NEW", 0, 32)
      LSet Y32$ = String$(32, 0)
      Put #N1%, 1, Y32$
      Close #N1%
      '
      N1% = FILEOPEN%(LU$ + "IDIARIO" + SXX$ + ".NEW", 0, 128)
      LSet Y128$ = String$(128, 0)
      Put #N1%, 1, Y128$
      Close #N1%
      '
      N1% = FILEOPEN%(LU$ + "ENCAJA" + SXX$ + ".NEW", 0, 64)
      LSet Y64$ = String$(64, 0)
      Put #N1%, 1, Y64$
      Close #N1%
      '
      N1% = FILEOPEN%(LU$ + "SUCAJA" + SXX$ + ".NEW", 0, 64)
      LSet Y64$ = String$(64, 0)
      Put #N1%, 1, Y64$
      Close #N1%
      '
      N1% = FILEOPEN%(LU$ + "ISUCAJA" + SXX$ + ".NEW", 0, 128)
      LSet Y128$ = String$(128, 0)
      Put #N1%, 1, Y128$
      Close #N1%
      '
   Next SX%
   '
   N1% = FILEOPEN%(LU$ + "INFORMES.DAT", 0, 6)
   Dim X6 As String * 6
   N2% = FILEOPEN%(LU$ + "INFORMES.NEW", 0, 6)
   Dim Y6 As String * 6
   '
   For I% = 1 To LOF(N1%) / 6
      Get #N1%, I%, X6$
      Z$ = X6$
      LSet Y6$ = Z$
      Put #N2%, I%, Y6$
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "ENCAINF.DAT", 0, 64)
   N2% = FILEOPEN%(LU$ + "ENCAINF.NEW", 0, 64)
   '
   For I% = 1 To LOF(N1%) / 64
      Get #N1%, I%, X64$
      Z$ = X64$
      LSet Y64$ = Z$
      Put #N2%, I%, Y64$
      NUCUENUE% = I%
   Next I%
   Close #N2%: Close #N1%
   '
   ARQ$(1) = "VARIAB"
   ARQ$(2) = "SUBEMPRE"
   ARQ$(3) = "CUENTAS"
   ARQ$(4) = "DENOCUE"
   ARQ$(5) = "INVECUE"
   ARQ$(6) = "PLAGRAL"
   ARQ$(7) = "PLANALT"
   ARQ$(8) = "TRADENO"
   ARQ$(9) = "DENENCA"
   ARQ$(10) = "INDAJU"
   ARQ$(11) = "SALCAJA"
   ARQ$(12) = "INFORMES"
   ARQ$(13) = "ENCAINF"
   '
   For I% = 1 To 13
      ARXQ$ = ARQ$(I%)
      On Error Resume Next
      Kill LU$ + ARXQ$ + ".OLD"
      '
      FileCopy LU$ + ARXQ$ + ".DAT ", LU$ + ARXQ$ + ".OLD"
      FileCopy LU$ + ARXQ$ + ".NEW ", LU$ + ARXQ$ + ".DAT"
      '
      Kill LU$ + ARXQ$ + ".NEW"
      On Error GoTo 0
   Next I%
   '
   ARQ$(1) = "CIERRE"
   ARQ$(2) = "ENCDIAR"
   ARQ$(3) = "DIARIO"
   ARQ$(4) = "IDIARIO"
   ARQ$(5) = "ENCAJA"
   ARQ$(6) = "SUCAJA"
   ARQ$(7) = "ISUCAJA"
   '
   For I% = 1 To 7
      For SX% = 1 To NSE%
         SXX$ = TRIM$(Str$(SX% - 1))
         ARXQ$ = ARQ$(I%) + SXX$
         On Error Resume Next
         Kill LU$ + ARXQ$ + ".OLD"
         '
         FileCopy LU$ + ARXQ$ + ".DAT ", LU$ + ARXQ$ + ".OLD"
         FileCopy LU$ + ARXQ$ + ".NEW ", LU$ + ARXQ$ + ".DAT"
         '
         Kill LU$ + ARXQ$ + ".NEW"
         On Error GoTo 0
      Next SX%
      '
   Next I%
   '
99 Screen.MousePointer = 1
   Call COMUNI("Apertura Completa.\  \Re-Inicie Ahora el Sistema\de Contabilidad General.")
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
End Sub

Sub BORRARESCAJ()
   '
   REFE$ = TRIM$("RESUMEN DE CAJA-BANCOS AL")
   LS% = MAXENCDIARIO%
   JJ& = 0
   For L1% = 1 To LS%
      Call LEENCDIARIO(L1%)
      If Left$(E7$, 25) = REFE$ Then             ' ENCONTRO EL ASIENTO
         XX$ = REGBLAN$("SELASIEN")
         Call REPLA(XX$, E1$, 1, 4)
         Call REPLA(XX$, E7$, 5, 40)
         JJ& = JJ& + 1
         Call GRAREG("!SELASIEN", XX$, JJ&)
      End If
   Next L1%
   '
   Call SETULTREG("!SELASIEN", JJ&)
   Call CIERRARCH("!SELASIEN")
   Call FRESHECHO("!SELASIEN")
   '
   If ULTREG&("!SELASIEN") < 1 Then
      Call MENSERR(24, "No hay Asientos Resumen\que Puedan Anularse.")
      Exit Sub
   End If
   '
   Call SELECHO("!SELASIEN")
   NROAS = Val(VALACT1$("!SELASIEN"))
   If NROAS < 1 Then
      Exit Sub
   End If
   '
   Call BLOQARCH("ACONT" + EXTE$)
   'Call BLOQARCH("ASICONT")
   Screen.MousePointer = 11
   FIN& = MAXDIA%
   For I% = 1 To MAXDIA%
      U& = I%
      Call TRACE(20, U&, FIN&)
      Call LEESUCAJA(I%)
      NRXX = CVS(D19$)
      If Abs(NRXX - NROAS) < 0.05 Then
         DXX$ = D10$
         Call REPLA(DXX$, MKS$(0), 55, 4)
         LSet D10$ = DXX$
         Put #NX6%, I%, D10$
      End If
   Next I%
   '
   Close #NX6%: NX6% = 0
   '
   FIN& = MAXDIARIO%
   MMXX% = MAXDIARIO%
   I1% = 0
   For I% = 1 To MAXDIARIO%
      Call LEEDIARIO(I%)
      If Abs(CVS(E12$) - NROAS) > 0.01 Then
         I1% = I1% + 1
         Put #NX10%, I1%, E10$
      End If
   Next I%
   '
   For I% = I1% + 1 To MMXX%
      LSet E10$ = String$(32, 0)
      Put #NX10%, I%, E10$
   Next I%
   Close #NX10%: NX10% = 0
   '
   FIN& = MAXENCDIARIO%
   MMXX% = MAXENCDIARIO%
   I1% = 0
   For I% = 1 To MAXENCDIARIO%
      Call LEENCDIARIO(I%)
      If Abs(CVS(E1$) - NROAS) > 0.01 Then
         I1% = I1% + 1
         Put #NX9%, I1%, E0$
      End If
   Next I%
   '
   For I% = I1% + 1 To MMXX%
      LSet E0$ = String$(64, 0)
      Put #NX9%, I%, E0$
   Next I%
   Close #NX9%: NX9% = 0
   '
   Screen.MousePointer = 1
   '
   Call CIERRARCH("*.*")
   Close
   End
   '
End Sub

Private Sub Form_Resize()
   '
   If WindowState = 1 Then
     Call CIERRACONTA
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRACONTA
   Hide
End Sub
