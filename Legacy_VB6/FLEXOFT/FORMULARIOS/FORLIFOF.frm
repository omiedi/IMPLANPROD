VERSION 5.00
Begin VB.Form FORLIFOF 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidación de Fondo Fijo (Ordenes de Pago)"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8200
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12000
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2745
         Left            =   10680
         TabIndex        =   29
         Top             =   1155
         Width           =   1170
         Begin VB.CommandButton Command11 
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
            Height          =   780
            Left            =   210
            Picture         =   "FORLIFOF.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Apertura de Nueva Cuenta de Proveedor"
            Top             =   1785
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Anular"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   210
            Picture         =   "FORLIFOF.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Anulacion de Liquidacion de Fondo Fijo"
            Top             =   945
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            Caption         =   "C.Cte."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   210
            Picture         =   "FORLIFOF.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   105
            Width           =   855
         End
      End
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
         Height          =   330
         Left            =   735
         MaxLength       =   72
         TabIndex        =   28
         Top             =   6930
         Width           =   9570
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   4905
         Left            =   210
         ScaleHeight     =   4845
         ScaleWidth      =   11505
         TabIndex        =   7
         Top             =   1260
         Width           =   11565
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   1.11565e5
            Begin VB.Line Line10 
               X1              =   6930
               X2              =   6930
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line13 
               X1              =   10430
               X2              =   10430
               Y1              =   0
               Y2              =   315
            End
            Begin VB.Line Line1 
               X1              =   10080
               X2              =   10080
               Y1              =   0
               Y2              =   1260
            End
            Begin VB.Line Line2 
               X1              =   8505
               X2              =   8505
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "    Cta.    Proveedor / Denominación                Documento                                  Saldo             Importe "
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
               Left            =   0
               TabIndex        =   14
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line8 
               X1              =   735
               X2              =   735
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line9 
               X1              =   4515
               X2              =   4515
               Y1              =   0
               Y2              =   1050
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5630
            Left            =   0
            TabIndex        =   8
            Top             =   315
            Width           =   1.11565e5
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
               Height          =   400
               Left            =   10710
               Picture         =   "FORLIFOF.frx":091E
               TabIndex        =   33
               ToolTipText     =   "Eldición del Item Activo"
               Top             =   3990
               Width           =   580
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Del"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   10710
               Picture         =   "FORLIFOF.frx":0C28
               TabIndex        =   32
               ToolTipText     =   "Elimina el Item Activo "
               Top             =   3465
               Width           =   580
            End
            Begin VB.CommandButton Command12 
               Caption         =   "Ins."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   10710
               Picture         =   "FORLIFOF.frx":0F32
               TabIndex        =   31
               ToolTipText     =   "Intercala un Item Antes del Activo"
               Top             =   2940
               Width           =   580
            End
            Begin VB.CommandButton Command6 
               Caption         =   "Add."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   10710
               Picture         =   "FORLIFOF.frx":123C
               TabIndex        =   30
               ToolTipText     =   "Agrega un Item al Final"
               Top             =   2415
               Width           =   580
            End
            Begin VB.ListBox List2 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4110
               Left            =   -30
               TabIndex        =   9
               Top             =   0
               Width           =   10450
            End
            Begin VB.Line Line6 
               X1              =   10395
               X2              =   11445
               Y1              =   2310
               Y2              =   2310
            End
            Begin VB.Line Line5 
               X1              =   11480
               X2              =   11480
               Y1              =   2310
               Y2              =   4515
            End
            Begin VB.Line Line14 
               BorderWidth     =   2
               X1              =   0
               X2              =   10395
               Y1              =   2085
               Y2              =   2085
            End
            Begin VB.Line Line24 
               X1              =   10425
               X2              =   10425
               Y1              =   0
               Y2              =   4515
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   8295
               TabIndex        =   16
               Top             =   4200
               Width           =   1695
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Width           =   1695
            End
            Begin VB.Line Line4 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   2415
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   6615
               TabIndex        =   12
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   3570
               TabIndex        =   11
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Liquidado Fondo Fijo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   105
               TabIndex        =   10
               Top             =   4200
               Width           =   4635
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   11550
               Y1              =   4515
               Y2              =   4515
            End
         End
      End
      Begin VB.CommandButton Command2 
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
         Left            =   10895
         Picture         =   "FORLIFOF.frx":1546
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   210
         Width           =   855
      End
      Begin VB.CommandButton BOTNEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   10670
         Picture         =   "FORLIFOF.frx":1690
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba - Registra - Imprime Liquidación"
         Top             =   6300
         Width           =   1065
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
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
         Left            =   10500
         TabIndex        =   0
         Top             =   210
         Width           =   180
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Obs:"
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
         TabIndex        =   27
         Top             =   7035
         Width           =   750
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000E&
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
         Height          =   330
         Left            =   1575
         TabIndex        =   26
         Top             =   6405
         Width           =   4110
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "M-P:"
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
         TabIndex        =   25
         Top             =   6510
         Width           =   750
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000E&
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
         Height          =   330
         Left            =   735
         TabIndex        =   24
         Top             =   6405
         Width           =   645
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000E&
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
         Left            =   9030
         TabIndex        =   23
         Top             =   735
         Width           =   1650
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000E&
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
         Left            =   9450
         TabIndex        =   22
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Documento no Válido como Factura"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5565
         TabIndex        =   21
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   840
         Width           =   2115
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   6405
         TabIndex        =   20
         Top             =   210
         Width           =   435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Liquidación de Fondo Fijo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   210
         TabIndex        =   19
         Top             =   420
         Width           =   6000
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   8400
         TabIndex        =   18
         Top             =   6405
         Width           =   1905
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Liquidado:"
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
         Left            =   5880
         TabIndex        =   17
         Top             =   6510
         Width           =   2430
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Numero:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7560
         TabIndex        =   5
         Top             =   735
         Width           =   1380
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Left            =   8610
         TabIndex        =   4
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pos.IVA:"
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
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "FORLIFOF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$


Private Sub BOTNEXT_Click()
   If List2.ListCount < 1 Or Val(Label20.Caption) < 0.005 Then
       Call MENSERR(24, "Falta Ingresar Detalle de Liquidacion")
       Exit Sub
   End If
   '
   If Val(Label15) < 1 Then
       Call MENSERR(24, "Imposible Registrar.\Medio de Pago no Válido.")
       Exit Sub
   End If
   '
   Call GRALIFOF(OKEY%)
   If OKEY% < 0 Then
      Exit Sub
   End If
   '
   If OKEY% > 0 Then
      Call PRILIFOF
   End If
   '
   Call CLEARFORLIFOF
   Call LIFOFIJ1
   '
End Sub


Private Sub Command1_Click()
   Command1.Enabled = False
   Call ANULIFF
   Command1.Enabled = True
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("AMAPRO/NEW", "Alta de Nueva Cuenta de Proveedor")
    Call FRESHECHO("PROVED1")
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Call MENSERR(24, "Opción Inserción no Disponible.\Debe Agregarse al Final de la Lista.")
End Sub

Private Sub Command2_Click()
    If List2.ListCount > 0 Then
          List2.Clear
        Else
          Call CIERRARCH("*.*")
          FORPASTA.Option1.Value = True
          FORPASTA.Show
          Unload FORLIFOF
    End If
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    PROLLAMA$ = Name
    Hide
    RESCUPRO.Show
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   '
   If List2.ListIndex < 0 Then Exit Sub
   If List2.ListIndex > List2.ListCount - 1 Then Exit Sub
   '
   ILISTA& = List2.ListIndex
   List2.RemoveItem ILISTA&
   List2.ListIndex = (-1)
   If List2.ListCount > ILISTA& Then
       List2.ListIndex = ILISTA&
   End If
   '
   Call LIFOFIJ4
   '
End Sub

Private Sub Command6_Click()
    '
    Call LIFOFIJ3(OKEY%)
    If OKEY% <> 1 Then
       Exit Sub
    End If
    '
End Sub
'


Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
10  FORIPRE$ = TRIM$(CONTROL$("", "FORLIFOF"))
    'If FORIPRE$ = "" Then
    '    Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Pedido de Cliente.")
    '    OPORDESPA.Show 1
    '    GoTo 10
    'End If
    '
    'If FORIPRE$ <> "" Then
    '    LU$ = LUDAT$
    '    If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
    '        RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
    '        Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Pedido de Cliente:\  \'" + RMCC$ + "'.")
    '        OPORDESPA.Show 1
    '        GoTo 10
    '    End If
    'End If
    '
    FF% = HOY(HO$)
    Label13.Caption = HO$
    Refresh
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    FORPASTA.Option1.Value = True
    FORPASTA.Show
End Sub
