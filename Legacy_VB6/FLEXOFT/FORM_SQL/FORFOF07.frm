VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FORLIFOF 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidación de Fondo Fijo (Ordenes de Pago)"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11730
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   1440
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   34
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C7D9CC&
      BorderStyle     =   0  'None
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
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   10890
         Picture         =   "FORFOF07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   240
         Width           =   650
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   1215
         MaxLength       =   72
         TabIndex        =   24
         Top             =   6930
         Width           =   9090
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   4905
         Left            =   240
         ScaleHeight     =   4875
         ScaleWidth      =   11385
         TabIndex        =   5
         Top             =   1200
         Width           =   11415
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   80
            OleObjectBlob   =   "FORFOF07.frx":014A
            TabIndex        =   35
            Top             =   30
            Width           =   735
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H0050765A&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   5630
            Left            =   -120
            TabIndex        =   6
            Top             =   360
            Width           =   1.11565e5
            Begin VB.CommandButton Command4 
               Caption         =   "C.Cte."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   10720
               Picture         =   "FORFOF07.frx":01AC
               Style           =   1  'Graphical
               TabIndex        =   31
               ToolTipText     =   "Resumen de Cuenta Corriente"
               Top             =   120
               Width           =   650
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Anular"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   10720
               Picture         =   "FORFOF07.frx":04B6
               Style           =   1  'Graphical
               TabIndex        =   30
               ToolTipText     =   "Anulacion de Liquidacion de Fondo Fijo"
               Top             =   840
               Width           =   650
            End
            Begin VB.CommandButton Command11 
               Caption         =   "New"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   10720
               Picture         =   "FORFOF07.frx":07C0
               Style           =   1  'Graphical
               TabIndex        =   29
               ToolTipText     =   "Apertura de Nueva Cuenta de Proveedor"
               Top             =   1560
               Width           =   650
            End
            Begin VB.CommandButton Command8 
               Caption         =   "Edit"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   10720
               Picture         =   "FORFOF07.frx":0ACA
               TabIndex        =   28
               ToolTipText     =   "Eldición del Item Activo"
               Top             =   3990
               Width           =   580
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Del"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   10720
               Picture         =   "FORFOF07.frx":0DD4
               TabIndex        =   27
               ToolTipText     =   "Elimina el Item Activo "
               Top             =   3465
               Width           =   580
            End
            Begin VB.CommandButton Command12 
               Caption         =   "Ins."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   10720
               Picture         =   "FORFOF07.frx":10DE
               TabIndex        =   26
               ToolTipText     =   "Intercala un Item Antes del Activo"
               Top             =   2940
               Width           =   580
            End
            Begin VB.CommandButton Command6 
               Caption         =   "Add."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   10720
               Picture         =   "FORFOF07.frx":13E8
               TabIndex        =   25
               ToolTipText     =   "Agrega un Item al Final"
               Top             =   2415
               Width           =   580
            End
            Begin VB.ListBox List2 
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
               Height          =   4080
               Left            =   120
               TabIndex        =   7
               Top             =   0
               Width           =   10455
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
               ForeColor       =   &H8000000B&
               Height          =   225
               Left            =   8295
               TabIndex        =   12
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
               Height          =   315
               Left            =   3600
               TabIndex        =   11
               Top             =   4200
               Width           =   1455
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
               TabIndex        =   10
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
               TabIndex        =   9
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Liquidado Fondo Fijo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   225
               Left            =   240
               TabIndex        =   8
               Top             =   4200
               Width           =   4635
            End
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   900
            OleObjectBlob   =   "FORFOF07.frx":16F2
            TabIndex        =   36
            Top             =   30
            Width           =   2415
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   4620
            OleObjectBlob   =   "FORFOF07.frx":1778
            TabIndex        =   37
            Top             =   30
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   7470
            OleObjectBlob   =   "FORFOF07.frx":17E0
            TabIndex        =   38
            Top             =   30
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   9060
            OleObjectBlob   =   "FORFOF07.frx":1840
            TabIndex        =   39
            Top             =   30
            Width           =   975
         End
      End
      Begin VB.CommandButton BOTNEXT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   10670
         Picture         =   "FORFOF07.frx":18A4
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba - Registra - Imprime Liquidación"
         Top             =   6300
         Width           =   825
      End
      Begin VB.CommandButton Command9 
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
         Height          =   300
         Left            =   10500
         TabIndex        =   0
         Top             =   210
         Width           =   180
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Observación"
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
         Left            =   120
         TabIndex        =   23
         Top             =   7035
         Width           =   975
      End
      Begin VB.Label Label17 
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
         Height          =   330
         Left            =   1575
         TabIndex        =   22
         Top             =   6405
         Width           =   4110
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "M-Pago"
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
         Left            =   140
         TabIndex        =   21
         Top             =   6500
         Width           =   615
      End
      Begin VB.Label Label15 
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
         Height          =   330
         Left            =   735
         TabIndex        =   20
         Top             =   6405
         Width           =   645
      End
      Begin VB.Label Label14 
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
         Left            =   9030
         TabIndex        =   19
         Top             =   735
         Width           =   1650
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   9450
         TabIndex        =   18
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Documento no Válido como Factura"
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
         Left            =   5400
         TabIndex        =   17
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   720
         Width           =   2115
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   6240
         TabIndex        =   16
         Top             =   120
         Width           =   435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Liquidación de Fondo Fijo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   210
         TabIndex        =   15
         Top             =   420
         Width           =   6000
      End
      Begin VB.Label Label20 
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
         Height          =   330
         Left            =   8400
         TabIndex        =   14
         Top             =   6405
         Width           =   1905
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Liquidado"
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
         Left            =   5880
         TabIndex        =   13
         Top             =   6510
         Width           =   2430
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Numero"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7560
         TabIndex        =   4
         Top             =   810
         Width           =   1380
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8880
         TabIndex        =   3
         Top             =   285
         Width           =   495
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   2640
      OleObjectBlob   =   "FORFOF07.frx":1BAE
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FORFOF07.frx":1C28
      Top             =   0
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu ver 
      Caption         =   "Ver"
      Begin VB.Menu verff 
         Caption         =   "Visualizar Fondo Fijo"
      End
   End
   Begin VB.Menu CL 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu FondoFijos 
         Caption         =   "Reportes de  Fondos Fijos"
      End
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
          ORPAGO07.Option1.Value = True
          'ORPAGO07.Show
          'FORLIFOF.Hide
          Unload Me
          ORPAGO07.Visible = True
    End If
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    PROLLAMA$ = Name
    RECUEP07.Show
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


Private Sub FondoFijos_Click()
    Call FINAL("?FONFIJO")
End Sub

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
    
    Call AplicaFuenteAnchoFijo(List2, , 10)
    Call AplicaFuenteAnchoFijo(Label10, , 10)
    Call AplicaFuenteAnchoFijo(Label20, , 10)
    
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture2)
    Refresh
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    ORPAGO07.Option1.Value = True
    'ORPAGO07.Show
End Sub

Sub ANULIFF()
    '
    Screen.MousePointer = 11
    '
    FECHALIM% = XVALO(CONTROL$("", "CIEMECOM"))
    Desde = FETEXCOR1$(FECHALIM%)
    '
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    SQLX$ = "SELECT DISTINCT CODICOM_LAR, SUM(VALABSCOMP) AS IMPORTE FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE (TipoMovim = 'OPAG' AND OpciMovim = 'FF') "
    SQLX$ = SQLX$ + " AND FechEmisi > '" & Desde & "' "
    SQLX$ = SQLX$ + " GROUP BY CODICOM_LAR Order by CODICOM_LAR ASC "
    Set CABATEMP = READSET(SQLX$)
    '
    JJ& = 0
    NIVCONXXX% = NIVCONFI%
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
            COMPROB$ = SAFETEXT$(!CODICOM_LAR)
            CONDI$ = "CODICOM_LAR='" & COMPROB$ & "'"
          NCLIE% = XVALO(VALOBADA("CAJABANC", "NUPROVE", CONDI$))
            IMPO# = XVALO(!IMPORTE)
          IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
          NPX& = XVALO(Mid$(COMPROB$, 4))
          y$ = MKS$(NPX&) + AJUSTI$(COMPROB$, 18) + IMPOS$
          JJ& = JJ& + 1
          Call GRAREG("!SELCUECO", y$, JJ&)
70        .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("!SELCUECO", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUECO") < 1 Then
       Call MENSERR(24, "No se han Encontrado Fondos Fijos que Puedan Anularse.")
       GoTo 99
    End If
    '
    Call CIERRARCH("!SELCUECO")
    Call FRESHECHO("!SELCUECO")
    '
12  Call SELECHO("!SELCUECO/LISTA DE FONDOS FIJOS")
    If TRIM$(VALACT2$("!SELCUECO")) = "" Then GoTo 99
    '
    DOCULA$ = TRIM$(Left$(VALACT2$("!SELCUECO"), 18))
    If COMALTER%("¿Desea Anular el Comprobante: '" + DOCULA$ + "'?\  Importe: " + IMPOS$ + "\\No, Abandonar\Si, Confirma Anulacion") <> 2 Then GoTo 99
    If VERICONCILIA%(DOCULA$) < 1 Then GoTo 99
    '
    CONDI$ = "SELECT DISTINCT NUPROVE FROM CAJABANC WITH(NOLOCK) WHERE (TipoMovim = 'OPAG' AND OpciMovim = 'FF') AND FechEmisi > '" & Desde & "' GROUP BY NUPROVE "
    Set Rst = READSET(CONDI$)
    XX$ = "": JJ& = 0
    With Rst
        Do While Not .EOF
            XX$ = XX$ & MKI$(XVALO(!nuprove))
            JJ& = JJ& + 2
            .MoveNext
        Loop
    End With
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    CONDI$ = "CodiCom_Lar = '" & DOCULA$ & "'"
    Call BORRAREGISTROS("CAJABANC", CONDI$, REGAF&)
    CONDI$ = "CoDocAp_Lar = '" & DOCULA$ & "'"
    Call BORRAREGISTROS("APLIPAGO", CONDI$, REGAF&)
    If JJ& > 0 Then
        For II& = 1 To JJ& Step 2
            NPROV% = CVI(Mid$(XX$, II&, 2))
            Call RAPLIPRO(NPROV%)
        Next II&
    End If
    
    Call COMUNI("Anulación de Fondo Fijo Realizada")

    Call LIFOFIJ1
    
99 Exit Sub
   '


'   '
'   Screen.MousePointer = 11
'   Dim NPROVE%(512)
'   '
'   FECHALIM% = XVALO(CONTROL$("", "CIEMECOM"))
'   FIN& = ULTREG&("CUECORR")
'   NULIFFA$ = ""
'   Call BLANARCH("!SELDOCAP")
'   Call FRESHECHO("!SELDOCAP")
'   '
'   For REG& = 1 To FIN&
'      Call TRACE(20, REG&, FIN&)
'      x$ = REGLEIDO$("CUECORR", REG&)
'      TIX$ = Mid$(x$, 7, 2)
'      If TIX$ = "FF" Then
'         If CVI(Left$(x$, 2)) > FECHALIM% Then
'            If Mid$(x$, 12, 4) <> "anul" Then
'               NULIFF$ = Mid$(x$, 17, 8)
'               If NULIFF$ <> NULIFFA$ Then
'                 NROX$ = Mid$(x$, 7, 18)
'                 FEMI = CVI(Left$(x$, 2))
'                 REGTEX$ = CSTRING$(MKS$(REG&), 3, 8, 0, 2)
'                 '
'                 IMPO# = 0: KK% = 0
'                 For U& = REG& To FIN&
'                    x$ = REGLEIDO$("CUECORR", U&)
'                    If Mid$(x$, 17, 8) <> NULIFF$ Then GoTo 20
'                    IMPO# = IMPO# + CVD(Mid$(x$, 115, 8))
'                    KK% = KK% + 1
'                 Next U&
'                 '
'20               z$ = Space$(128)
'                 IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 14, 2, 2)
'                 CAITEX$ = CSTRING$(MKI$(KK%), 1, 6, 0, 2)
'                 Call REPLA(z$, NROX$, 1, 18)
'                 Call REPLA(z$, FECHATEX$(FEMI), 24, 8)
'                 Call REPLA(z$, IMPTEX$, 36, 14)
'                 Call REPLA(z$, CAITEX$, 51, 6)
'                 Call REPLA(z$, MKS$(REG&), 61, 4)
'                    '
'                 Call REGAPP("!SELDOCAP", z$)
'                 NULIFFA$ = NULIFF$
'                 KQ& = KQ& + 1
'               End If
'            End If
'         End If
'      End If
'   Next REG&
'   Call CIERRARCH("!SELDOCAP")
'   '
'   Screen.MousePointer = 1
'   y$ = REGSEL$("!SELDOCAP")
'   If Len(y$) <= 4 Then
'      Exit Sub
'   End If
'   '
'   REG& = CVS(Mid$(y$, 61, 4))
'   IMPOS$ = "$ " + TRIM$(Mid$(y$, 36, 14))
'   DOCU$ = Left$(y$, 18)
'   NUDOCU$ = Mid$(DOCU$, 11, 8)
'   NORPA& = XVALO(NUDOCU$)
'   '
'   If COMALTER%("Documento: " + DOCU$ + "\  Importe: " + IMPOS$ + "\\Abandona\Confirma Anulacion") <> 2 Then Exit Sub
'   '
'   ' anulacion
'   ' *********
'   '
'   Screen.MousePointer = 11
'   Call BLOQARCH("CUECORR")
'   Call BLOQARCH("PROVED2")
'   CAPROVE% = 0
'   '
'   For I& = REG& To FIN&
'      '
'      XX0$ = REGLEIDO$("CUECORR", I&)
'      If Mid$(XX0$, 7, 2) <> "FF" Then GoTo 30
'      If Mid$(XX0$, 17, 8) <> NUDOCU$ Then GoTo 30
'      '
'      FF1% = CVI(Left$(XX0$, 2))
'      NC% = CVI(Mid$(XX0$, 5, 2))
'      NF$ = Mid$(XX0$, 7, 18)
'      '
'      Call REPLA(NF$, "anul", 6, 4)
'      XD# = CVD(Mid$(XX0$, 115, 8))
'      XH# = CVD(Mid$(XX0$, 123, 8))
'      IM# = CVD(Mid$(SS8X$, 3, 8)) + XD# - XH#
'      Call REPLA(SS8X$, MKI$(FF1%), 1, 2)
'      Call REPLA(SS8X$, MKD$(IM#), 3, 8)
'      Call REPLA(XX0$, NF$, 7, 18)
'      Call REPLA(XX0$, MKD$(0), 115, 8)
'      Call REPLA(XX0$, MKD$(0), 123, 8)
'      Call REPLA(XX0$, SS8X$, 139, 60)
'      Call GRAREG("CUECORR", XX0$, I&)
'      '
'      NC1% = (-1) * NC%
'      RC& = REGICLI&(NC1%)
'      DX$ = REGLEIDO$("PROVED2", RC&)
'      SALDO# = CVD(Mid$(DX$, 109, 8)) + XD# - XH#
'      Call REPLA(DX$, MKD$(SALDO#), 109, 8)
'      Call GRAREG("PROVED2", DX$, RC&)
'      '
'      For KK% = 1 To CAPROVE%
'         If NPROVE%(KK%) = NC% Then GoTo 29
'      Next KK%
'      CAPROVE% = CAPROVE% + 1: NPROVE%(CAPROVE%) = NC%
'      '
'29 Next I&
'   '
'30 Call CIERRARCH("CUECORR")
'   FINCH& = ULTREG&("CHEQUEMI")
'     '
'   For U& = 1 To FINCH&
'      Call TRACE(20, U&, FINCH&)
'      z$ = REGLEIDO$("CHEQUEMI", U&)
'      If CVS(Mid$(z$, 41, 4)) = NORPA& Then
'         If CVI(Mid$(z$, 31, 2)) = FF1% Then
'            Call REPLA(z$, MKD$(0), 33, 8)
'            Call GRAREG("CHEQUEMI", z$, U&)
'         End If
'      End If
'   Next U&
'   Call CIERRARCH("CHEQUEMI")
'   '
'   FINAP& = ULTREG&("APLIPAG")
'   '
'   For U& = 1 To FINAP&
'      Call TRACE(20, U&, FINAP&)
'      z$ = REGLEIDO$("APLIPAG", U&)
'      If XVALO(Mid$(z$, 17, 6)) = NORPA& Then
'         If CVI(Left$(z$, 2)) = FF1% Then
'            Call REPLA(z$, "anul", 10, 4)
'            Call REPLA(z$, MKD$(0), 23, 8)
'            Call GRAREG("APLIPAG", z$, U&)
'         End If
'      End If
'   Next U&
'   Call CIERRARCH("APLIPAG")
'   '
'   Screen.MousePointer = 1
'   If COMALTER%("Anulación Completa.\  \Puede Ahora Reconstruir la Composición de Deuda\de las Cuentas Correspondientes a Esta Liquidación.\\Reconstruir\Continuar") = 1 Then
'      For KK% = 1 To CAPROVE%
'         If TICUEPRO((-1) * NPROVE%(KK%)) = 0 Then
'            Call REAPLICREPRO(NPROVE%(KK%))
'         End If
'      Next KK%
'   End If
'   '
'   Call CIERRARCH("*.*")
'   Call LIBARCH("*.*")
'   '
'99 Exit Sub
'   '
End Sub
'

Sub GRALIFOF(OKEY%)
   '
   Screen.MousePointer = 11
   FENUE$ = CONTROL$("", "FENUORPA")
   FLIM% = FECHANUM(FENUE$)
   If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
   '
   UNRO& = 0: KK% = 0
   Screen.MousePointer = 11
   '
   ' OBTIENE LA CONFIGURACION DE LA CANTIDAD DE MOVIMIENTOS QUE GENERAR EN LA CUENTA CONTABLE
   ' DEL MEDIO DE PAGO. UNO POR EL TOTAL DE LO CANCELADO, O UNO POR CADA CBTE. CANCELADO
   VariosMovimientosParaPago% = 1
   If CONTROL$("ORPASTA", "FFCAMOMP") = "UNOSOLO" Then VariosMovimientosParaPago% = 0
   '
   ' OBTIENE PROXIMO NUMERO DE FONDO FIJO
   CONDI$ = "OPCIMOVIM='FF'"
   UNRO& = XVALO(VALOBADA("CAJABANC", "MAX(NUMECOMP)", CONDI$, "NOMESS"))
'   For I& = ULTREG&("CUECORR") To 1 Step -1
'      Screen.MousePointer = 11
'      x$ = REGLEIDO$("CUECORR", I&)
'      If UCase$(Mid$(x$, 7, 2)) = "FF" Then
'        If Mid$(x$, 10, 4) <> "anul" Then
'          FEMO% = CVI(Mid$(x$, 1, 2))
'          If FEMO% >= FLIM% Then
'            If XVALO(Mid$(x$, 19, 6)) > UNRO& Then
'               UNRO& = XVALO(Mid$(x$, 19, 6))
'            End If
'            KK% = KK% + 1
'            If KK% >= 128 Then GoTo 99
'          End If
'        End If
'      End If
'   Next I&
   '
99 NROPRO& = UNRO& + 1
   Screen.MousePointer = 1
   '
   ' VERIFICA QUE EL PROXIMO NO SEA MAYOR AL QUE APARECE EN PANTALLA
   FC% = FECHANUM(FORLIFOF.Label13.Caption)
   NRO& = XVALO(Mid$(FORLIFOF.Label14.Caption, 4))
   If NROPRO& > NRO& Then
       Call CIERRARCH("CUECORR")
       Call CIERRARCH("PROVED2")
       Call CIERRARCH("APLIPAG")
       Call CIERRARCH("CHEQUEMI")
       Call MENSERR(24, "Al Grabar esta Liquidación\el Número Propuesto Estaba Ocupado.\  \Verifique el Nuevo Número Propuesto\y Vuelva a Grabar e Imprimir.")
       FORLIFOF.Label14.Caption = "FF." + TRIM$(Str$(NROPRO&))
       FORLIFOF.Label14.Refresh
       OKEY% = (-1)
       Exit Sub
    End If
    '
   RefApli$ = "Liq.Fondo Fijo Nro." + Str$(NRO&)
   '
   ' RECORRE LISTA DE COMPROBANTES A PAGAR
   For U& = 1 To FORLIFOF.List2.ListCount
      Z$ = FORLIFOF.List2.LIST(U& - 1)
      NC% = XVALO(Mid$(Z$, 1, 5))
      DOCUCA$ = Mid$(Z$, 39, 18)
      IPEN# = XVALO(Mid$(Z$, 59, 12))
      CANCE# = XVALO(Mid$(Z$, 72, 12))
      '
      If CANCE# > 0 Then
        ORD% = XVALO(Mid$(Z$, 97, 4))
        If ORD% > 0 Then
           If ORD% <= 6 Then
             RECILAR$ = FORLIFOF.Label14.Caption
             NRECIB$ = FORLIFOF.Label14.Caption
             CoComCor$ = NRECIB$
             CoComLar$ = RECILAR$
             NumeProv% = NC%
             NumeClie = 0
             NC1% = (-1) * NC%
             Referen$ = rasocli$(NC1%)
             CuentaCo% = CUECOCLI%(NC1%)
             FechEmi = CVDAT(FC%)
             FeContab = CVDAT(FC%)
             FeAcredi = CVDAT(FC%)
             TiMovi$ = "OPAG"
             OpMovi$ = "FF"
             VaMovi$ = "CCORR"
             Descrip$ = RefApli$
             NumeTal = 0
             NumeCom = NRO&
             StatInit = 0
             ImpoDeb# = CANCE#
             BasImpo# = CANCE#    ' aqui va entonces el neto sin IVA
             Observa$ = ""
             If Observa$ = "" Then Observa$ = rasocli$(NC1%)
             Call GRACABAN
             
             Dim Sacrepen As ADODB.Recordset
             Set Sacrepen = New ADODB.Recordset
             strsql = "Select * from SacrePen "
             strsql = strsql & " where NuProve = " & NumeProv%
             strsql = strsql & " and CODICOM_LAR = '" & DOCUCA$ & "'"
             strsql = strsql & " and ORVENCIM = " & ORD%
             strsql = strsql & " and ( ABS(SALDACRE) > 0 OR ABS(SALDDEBE) > 0 ) "
25           On Error Resume Next
             Sacrepen.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
             If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                On Error GoTo 0
                Call CapturaErrorOpen
                GoTo 25
             End If
             On Error GoTo 0
             ICAN# = CANCE#
             SICAN% = Sgn(CANCE#)
             DocOrCor$ = ""
             TotOrig# = 0
             IvaOrig# = 0
             With Sacrepen
              If Not (.EOF And .BOF) Then
                .MoveFirst
                Do While Not (.EOF)
850               PARCAN# = ICAN#
                  If SICAN% * (XVALO(!SaldAcre) - XVALO(!SaldDebe)) < SICAN% * ICAN# Then PARCAN# = XVALO(!SaldAcre) - XVALO(!SaldDebe)
                  If PARCAN# > 0 Then
                      !SaldAcre = XVALO(!SaldAcre) - PARCAN#
                    Else
                      !SaldDebe = XVALO(!SaldDebe) + PARCAN#
                  End If
                  .Update
                  TOTAPLI# = TOTAPLI# + PARCAN#   ' SUMA APLICACIONES -  agregado por tema bimonetario 22/03
                  '
                  DocOrCor$ = SAFETEXT$(!CODICOM_COR)
                  TotOrig# = XVALO(!IBruComp)
                  IIVAORIG# = 0
                  On Error Resume Next
                  IvaOrig# = XVALO(!IIvaComp)
                  On Error GoTo 0
                  '
                  ICAN# = ICAN# - PARCAN#
                  If Abs(ICAN#) >= 0.005 Then
                     .MoveNext
                  Else
                     Exit Do
                  End If
                Loop
               End If
             End With
             Sacrepen.Close
             Set Sacrepen = Nothing
             '
             DocOrLar$ = TRIM$(Mid$(Z$, 39, 18))
             ImpApli# = CANCE#
             Call GRAPLIPAG
              
             Tipo% = XVALO(FORLIFOF.Label15.Caption)
             CUECONTABINT% = CUEMEPAG%(Tipo%)
             BANCO$ = ""
             NUCHE$ = ""
             IMCHE# = CANCE#
             ImporteTotalCancelado# = ImporteTotalCancelado# + CANCE#
             
             FECHE% = FC%: If FECHE% < 1 Then FECHE% = HOY(HOS$)
             FEVEN% = FC%: If FEVEN% < FECHE% Then FEVEN% = FECHE%
             IDEVA$ = TRIM$(FORLIFOF.Label17.Caption)
             '
             HOII% = HOY(HOS$)
             If IMCHE# > 0 And VariosMovimientosParaPago% = 1 Then
                '
                CoComCor$ = NRECIB$
                CoComLar$ = RECILAR$
                NumeClie = 0
                NumeProv% = NC%
                FechEmi = CVDAT(FECHE%)
                  FEX = FC%
                Referen$ = RefApli$
                FeContab = CVDAT(FC%)
                '
                FeAcredi = CVDAT(FEVEN%)
                CuentaCo% = CUECONTABINT%
                TiMovi$ = "CHENTR"
                OpMovi$ = "OP"
                VaMovi$ = ""
                StatInit = 0
                Descrip$ = IDEVA$
                NumeTal = COBRACT%
                NumeCom = NRO&
                ImpoDeb# = (-1) * Abs(IMCHE#)
                BasImpo# = 0               ' CERO PARA LA SALIDA DEL CHEQUE
                Observa$ = RefApli$
                '
                Call GRACABAN
               End If
           End If
         End If
      End If
   Next U&
   
   If ImporteTotalCancelado# > 0.005 And VariosMovimientosParaPago% = 0 Then
     '
     Tipo% = XVALO(FORLIFOF.Label15.Caption)
     CUECONTABINT% = CUEMEPAG%(Tipo%)
     BANCO$ = ""
     NUCHE$ = ""
     IMCHE# = ImporteTotalCancelado#
     FECHE% = FC%: If FECHE% < 1 Then FECHE% = HOY(HOS$)
     FEVEN% = FC%: If FEVEN% < FECHE% Then FEVEN% = FECHE%
     IDEVA$ = TRIM$(FORLIFOF.Label17.Caption)
     '
     HOII% = HOY(HOS$)
     '
     CoComCor$ = NRECIB$
     CoComLar$ = RECILAR$
     NumeClie = 0
     NumeProv% = NC%
     FechEmi = CVDAT(FECHE%)
       FEX = FC%
     Referen$ = RefApli$
     FeContab = CVDAT(FC%)
     '
     FeAcredi = CVDAT(FEVEN%)
     CuentaCo% = CUECONTABINT%
     TiMovi$ = "CHENTR"
     OpMovi$ = "OP"
     VaMovi$ = ""
     StatInit = 0
     Descrip$ = IDEVA$
     NumeTal = COBRACT%
     NumeCom = NRO&
     ImpoDeb# = (-1) * Abs(IMCHE#)
     BasImpo# = 0               ' CERO PARA LA SALIDA DEL CHEQUE
     Observa$ = RefApli$
     '
     Call GRACABAN
   End If
    

    '
'   '
'   For U& = 1 To FORLIFOF.List2.ListCount
'      z$ = FORLIFOF.List2.List(U& - 1)
'      NC% = XVALO(Mid$(z$, 1, 5))
'      TI% = 14: VA% = 1: MO$ = "$": SUC% = U&
'      DOCORI$ = Mid$(z$, 39, 18)
'      IPEN# = XVALO(Mid$(z$, 59, 12))
'      CANCE# = XVALO(Mid$(z$, 72, 12))
'      If CANCE# > 0 Then
'         NREG& = XVALO(Mid$(z$, 89, 8))
'         If NREG& > 0 Then
'            If NREG& <= ULTREG&("CUECORR") Then
'               ORD% = XVALO(Mid$(z$, 97, 4))
'               If ORD% > 0 Then
'                  If ORD% <= 6 Then
'                     SS$ = REGLEIDO$("CUECORR", NREG&)
'                     '
'                     SALDO$ = Mid$(SS$, 139, 60)
'                     SALDEU# = CVD(Mid$(SALDO$, 10 * ORD% - 7, 8))
'                     SALDEU# = SALDEU# - CANCE#
'                     Call REPLA(SALDO$, MKD$(SALDEU#), 10 * ORD% - 7, 8)
'                     '
'                     If Abs(SALDEU#) < 0.005 Then
'                        SALDO$ = Left$(SALDO$, 10 * (ORD% - 1)) + Mid$(SALDO$, 10 * ORD% + 1) + String$(10, 0)
'                     End If
'                     '
'                     Call REPLA(SS$, SALDO$, 139, 60)
'                     Call GRAREG("CUECORR", SS$, NREG&)
'                     '
'                     DOCORI$ = Mid$(SS$, 7, 18)
'                     SS77$ = Mid$(SS$, 51)
'                     ITOTORI# = Abs(CVD(Mid$(SS$, 115, 8)) - CVD(Mid$(SS$, 123, 8)))
'                     IVAORI# = Abs(CVD(Mid$(SS77$, 25, 8)) + CVD(Mid$(SS77$, 33, 8)) + CVD(Mid$(SS77$, 41, 8)))
'                     '
'                     Call APLIPAG(FC%, (-1) * NC%, 14, 1, SUC%, NRO&, CANCE#, RefApli$, DOCORI$, ITOTORI#, IVAORI#)
'                     '
'                  End If
'               End If
'            End If
'         End If
'         '
'         TOFAC#(0) = CANCE#
'         TOFAC#(1) = 0: TOFAC#(2) = TOFAC#(0)
'         FVII%(1) = FC%
'         IMV#(1) = 0
'         Call CUECORR(FC%, (-1) * NC%, TI%, VA%, SUC%, NRO&, RefApli$, "$", TOFAC#(), FVII%(), IMV#(), GRU%(), IPAR#(), 1, FC%)
'         '
'         NUCHE$ = Space$(12)
'         IMCHE# = CANCE#
'         FECHE% = FC%: If FECHE% < 1 Then FECHE% = HOY(HOS$)
'         FEVEN% = FC%: If FEVEN% < FECHE% Then FEVEN% = FECHE%
'         Tipo% = XVALO(FORLIFOF.Label15.Caption)
'         '
'         Z1$ = REGBLAN$("CHEQUEMI")
'         Call REPLA(Z1$, MKI$(NC%), 1, 2)
'         Call REPLA(Z1$, NUCHE$, 3, 12)
'         Call REPLA(Z1$, BANCO$, 15, 16)
'         Call REPLA(Z1$, MKI$(FECHE%), 31, 2)
'         Call REPLA(Z1$, MKD$(IMCHE#), 33, 8)
'         Call REPLA(Z1$, MKS$(NRO&), 41, 4)
'         Call REPLA(Z1$, MKI$(Tipo%), 45, 2)
'         Call REPLA(Z1$, MKI$(FC%), 47, 2)
'         Call REPLA(Z1$, MKI$(FEVEN%), 49, 2)
'         Call REGAPP("CHEQUEMI", Z1$)
'         '
'      End If
'      '
'   Next U&
'   '
'   Call CIERRARCH("CUECORR")
'   Call CIERRARCH("PROVED2")
'   Call CIERRARCH("APLIPAG")
'   Call CIERRARCH("CHEQUEMI")
'   '
   OKEY% = 1
'   '
   Screen.MousePointer = 1
   '
End Sub

Sub LIFOFIJ4()
   '
   TOLIFOF# = 0
   For U& = 1 To FORLIFOF.List2.ListCount
      Z$ = FORLIFOF.List2.LIST(U& - 1)
      TOLIFOF# = TOLIFOF# + XVALO(Mid$(Z$, 72, 12))
   Next U&
   '
   TOLITEX$ = CSTRING$(MKD$(TOLIFOF#), 3, 12, 2, 2)
   FORLIFOF.Label10.Caption = TOLITEX$
   FORLIFOF.Label10.Refresh
   FORLIFOF.Label20.Caption = TOLITEX$
   FORLIFOF.Label20.Refresh
   '
End Sub
'
'



Sub PRILIFOF()
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORLIFOF"))
    If FORIPRE$ = "" Then
       Exit Sub
    End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Recibo de Cobranza:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    CODFOR$ = FORIPRE$
    CANCELPRINT% = 0
    If CONTROL$("ORPASTA", "SUPRINT") = "SI" Then CANCELPRINT% = 1
    '
    TTXX$ = TRIM$(FORLIFOF.Text1.TEXT)
    Call GRAREG("!OBSERVOF", TTXX$, 1)
    Call SETULTREG("!OBSERVOF", 1)
    Call CIERRARCH("!OBSERVOF")
    '
    Screen.MousePointer = 11
    HOII% = HOY(HOS$)
    FFII% = FECHANUM(FORLIFOF.Label13.Caption)
    FECHDOC$ = FETEXTO$(FFII%)
    NROPE& = XVALO(Mid$(FORLIFOF.Label14.Caption, 4))
    NUMEDOC$ = TRIM$(Str$(NROPE&))
    TIPODOC$ = "Liquidacion de Fondo Fijo"
    '
    CODTABU1$(1) = "TIPO-VAL"
    CODTABU1$(2) = "BANC-CHE"
    CODTABU1$(3) = "NUME-CHE"
    CODTABU1$(4) = "IMPO-PEN"
    CODTABU1$(5) = "IMPO-CAN"
    CACOLTAB1% = 5
    '
    DESTIDOC% = 0
    CAITAB1% = 0
    CAITCRU% = 0
    '
    Z$ = Space$(100)
    Call REPLA(Z$, CSTRING$(MKI$(NC%), 1, 5, 0, 2), 1, 5)
    Call REPLA(Z$, rasocli$((-1) * NC%), 8, 30)
    Call REPLA(Z$, Mid$(OBX$, 3, 18), 39, 18)
    Call REPLA(Z$, CSTRING$(MKD$(IPEN#), 3, 12, 2, 2), 59, 12)
    Call REPLA(Z$, CSTRING$(MKD$(ICAN#), 3, 12, 2, 2), 72, 12)
    Call REPLA(Z$, CSTRING$(MKS$(REG&), 3, 8, 0, 2), 89, 8)
    Call REPLA(Z$, CSTRING$(MKI$(KKI%), 1, 4, 0, 2), 97, 4)
    '
    SUMATOTA# = 0: J& = 0
    For U& = 1 To FORLIFOF.List2.ListCount
       Z$ = FORLIFOF.List2.LIST(U& - 1)
       NCX% = XVALO(Left$(Z$, 5))
       CLIX$ = rasocli$((-1) * NCX%)
       DOCU$ = Mid$(Z$, 39, 18)
       IPE$ = TRIM$(Mid$(Z$, 59, 12))
       ICA$ = Mid$(Z$, 72, 12)
       '
       If XVALO(ICA$) >= 0.005 Then
          SUMATOTA# = SUMATOTA# + CDbl(XVALO(ICA$))
          '
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = TRIM$(Str$(NCX%))
          TETABU1$(2, CAITAB1%) = CLIX$
          TETABU1$(3, CAITAB1%) = DOCU$
          TETABU1$(4, CAITAB1%) = IPE$
          TETABU1$(5, CAITAB1%) = ICA$
       End If
       '
240 Next U&
     '
     CODPARAM$(1) = "CUEN-CLI"
     DOCPARAM$(1) = FORLIFOF.Label15.Caption
     CODPARAM$(2) = "RASO-CLI"
     DOCPARAM$(2) = FORLIFOF.Label17.Caption
     CODPARAM$(3) = "IMP-TOTA"
     DOCPARAM$(3) = TRIM$(CSTRING$(MKD$(SUMATOTA#), 3, 12, 2, 2))
     '
     ITEXTO$ = NUMTEX$(SUMATOTA#) + " ctvs"
     LOTEX% = XVALO(ATRIFORM$(CODFOR$, "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
     Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
     IMPLET1$ = "": IMPLET2$ = "": IMPLET3$ = ""
     If TRIM$(NRX1$) <> "" Then
       IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
     End If
     If TRIM$(NRX2$) <> "" Then
       IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
     End If
     If TRIM$(NRX3$) <> "" Then
       IMPLET3$ = TRIM$(NRX3$) + " " + "***************************************************************************************************************"
     End If
     '
     CODPARAM$(4) = "IMP-LET1"
     DOCPARAM$(4) = NRX1$
     CODPARAM$(5) = "IMP-LET2"
     DOCPARAM$(5) = NRX2$
     CODPARAM$(6) = "IMP-LET3"
     DOCPARAM$(6) = NRX3$
     CAPARDOC% = 5
     '
     Call PRINTDOC("FORLIFOF")
     Screen.MousePointer = 1
     '
End Sub
'

Sub CLEARFORLIFOF()
    '
    FORLIFOF.List2.Clear
    FORLIFOF.Label15.Caption = ""
    FORLIFOF.Label14.Caption = ""
    FORLIFOF.Label17.Caption = ""
    FORLIFOF.Label10.Caption = ""
    FORLIFOF.Label20.Caption = ""
    '
End Sub
'
'
Sub LIFOFIJ3(OKEY%)
    '
    OKEY% = 0
    If MODORPA% = 0 Then
        MODORPA% = 1       ' por emision
        If CONTROL$("ORPASTA", "MODORDE") = "VENCIM" Then
           MODORPA% = 2
        End If
    End If
    '
10  Call SELECHO("PROVED1")
    NC% = XVALO(VALACT1$("PROVED1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    NC1% = (-1) * NC%
    If TICUEPRO%(NC1%) <> 0 Then  ' cuenta de gastos
       '
       VDEF$ = String$(64, 0)
       Call REPLA(VDEF$, MKI$(NC%), 1, 2)
       Call REPLA(VDEF$, "Gasto Fondo Fijo", 3, 18)
       Call REPLA(VDEF$, MKD$(0), 21, 8)
       Call REPLA(VDEF$, MKD$(0), 29, 8)
       Call REPLA(VDEF$, String$(6, 0), 59, 6)
       '
       OBX$ = OBJPLA$("CANGASFIJ", VDEF$)
       GoTo 111
       '
    End If
    '
    Screen.MousePointer = 11
    Call BLANARCH("!SELDOCAP")
    '
100 SQLX$ = "SELECT * FROM SACREPEN WITH(NOLOCK) WHERE NUPROVE=" & NC% & " ORDER BY CODICOM_LAR, ORVENCIM"
    Dim Rst As ADODB.Recordset
    Set Rst = READSET(SQLX$)
    '
    KQ& = 0
    With Rst
        Do While Not .EOF
                '
                NROX$ = SAFETEXT$(!CODICOM_LAR)
                TIX$ = Mid$(NROX$, 1, 2)
                SG% = (-1)
                TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                                    If TIX$ = "TF" Then TIPOS$ = "Ticket-Factura"
                                    If TIX$ = "RH" Then TIPOS$ = "Rec.Honorarios"
                                    If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                                    If TIX$ = "CP" Then TIPOS$ = "Crédito Propio"
                                    If TIX$ = "OP" Then TIPOS$ = "Orden de Pago": SG% = 1
                                    If TIX$ = "FF" Then TIPOS$ = "O/P Fondo Fijo": SG% = 1
                                    If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
                                    If TIX$ = "DP" Then TIPOS$ = "Debito Propio": SG% = 1
                                    If TIX$ = "CC" Then TIPOS$ = "Comp.Cuenta": SG% = 1
                                    
                IMPO# = XVALO(!IBruComp) * SG%
                SALDO# = Abs(XVALO(!SaldDebe) - XVALO(!SaldAcre)) * Sgn(IMPO#)
                '
                IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 14, 2, 2)
                SALTEX$ = CSTRING$(MKD$(Abs(SALDO#)), 3, 14, 2, 2)
                '
                KKI% = XVALO(!OrVencim)
                FEMI = CVINT(!FECHEMISI)
                '
                If SALDO# <= (-0.005) Then
                    '
                    Z$ = Space$(60) + MKS$(0)
                    Call REPLA(Z$, NROX$, 1, 18)
                    Call REPLA(Z$, FECHATEX$(FEMI), 19, 8)
                    Call REPLA(Z$, IMPTEX$, 27, 14)
                    Call REPLA(Z$, SALTEX$, 41, 14)
                    'Call REPLA(z$, MKS$(REG&), 61, 4)
                    Call REPLA(Z$, MKI$(KKI%), 59, 2)
                    '
                    Call REGAPP("!SELDOCAP", Z$)
                    KQ& = KQ& + 1
                    '
                End If
        
            .MoveNext
        Loop
    End With

'100 MAXDEU& = ULTREG&("CUECORR")
'    RFD$ = RECFILT$("CUECORR", 4, MKI$(NC%))
'    '
'    RESCUPRO.LINDICE.Clear
'    For U& = 1 To Len(RFD$) Step 4
'        REG& = CVS(Mid$(RFD$, U&, 4))
'        x$ = REGLEIDO$("CUECORR", REG&)
'        If CVI(Mid$(x$, 5, 2)) <> NC% Then
'            x$ = REGLEIDO$("CUECORR", 1)
'            Call GRAREG("CUECORR", x$, 1)
'            Call CIERRARCH("CUECORR")
'            GoTo 100
'        End If
'        '
'        FEMI = CVI(Left$(x$, 2))
'        IMPO# = CVD(Mid$(x$, 115, 8)) - CVD(Mid$(x$, 123, 8))
'        SALDEU$ = Mid$(x$, 139, 60)
'        For KKI% = 1 To 6
'           FEVE = CVI(Mid$(SALDEU$, 10 * KKI% - 9, 2))
'           SALDO# = CVD(Mid$(SALDEU$, 10 * KKI% - 7, 8))
'           If Abs(SALDO#) >= 0.005 Then
'              If MODORPA% = 1 Then
'                  FF% = FEMI      ' orden emision
'                Else
'                  FF% = FEVE      ' orden vencimiento
'              End If
'              z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2) + CSTRING$(MKI$(KKI%), 1, 4, 0, 2)
'              RESCUPRO.LINDICE.AddItem z$
'           End If
'        Next KKI%
'    Next U&
'    '
'    SUMADEU# = 0: SUMAFAC# = 0: KQ& = 0
'    For U& = 1 To RESCUPRO.LINDICE.ListCount
'        RESCUPRO.LINDICE.ListIndex = U& - 1
'        REG& = XVALO(Mid$(RESCUPRO.LINDICE.TEXT, 7, 8))
'        KKI% = XVALO(Mid$(RESCUPRO.LINDICE.TEXT, 15, 4))
'        If REG& > 0 Then
'          If REG& <= MAXDEU& Then
'            If KKI% > 0 Then
'              If KKI% <= 6 Then
'                x$ = REGLEIDO$("CUECORR", REG&)
'                TIX$ = Mid$(x$, 7, 2)
'                SG% = (-1)
'                TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
'                                    If TIX$ = "TF" Then TIPOS$ = "Ticket-Factura"
'                                    If TIX$ = "RH" Then TIPOS$ = "Rec.Honorarios"
'                                    If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
'                                    If TIX$ = "CP" Then TIPOS$ = "Crédito Propio"
'                                    If TIX$ = "OP" Then TIPOS$ = "Orden de Pago": SG% = 1
'                                    If TIX$ = "FF" Then TIPOS$ = "O/P Fondo Fijo": SG% = 1
'                                    If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
'                                    If TIX$ = "DP" Then TIPOS$ = "Debito Propio": SG% = 1
'                                    If TIX$ = "CC" Then TIPOS$ = "Comp.Cuenta": SG% = 1
'                                    '
'                IMPO# = CVD(Mid$(x$, 115, 8)) - CVD(Mid$(x$, 123, 8))
'                SALDO# = (Abs(CVD(Mid$(x$, 139 + 10 * KKI% - 8, 8)))) * Sgn(IMPO#)
'                '
'                IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 14, 2, 2)
'                SALTEX$ = CSTRING$(MKD$(Abs(SALDO#)), 3, 14, 2, 2)
'                REGTEX$ = CSTRING$(MKS$(REG&), 3, 8, 0, 2)
'                IRETEX$ = CSTRING$(MKI$(KKI%), 1, 2, 0, 2)
'                '
'                NROX$ = Mid$(x$, 7, 18)
'                FEMI = CVI(Mid$(x$, 1, 2))
'                '
'                If SALDO# <= (-0.005) Then
'                    '
'                    z$ = Space$(60) + MKS$(0)
'                    Call REPLA(z$, NROX$, 1, 18)
'                    Call REPLA(z$, FECHATEX$(FEMI), 19, 8)
'                    Call REPLA(z$, IMPTEX$, 27, 14)
'                    Call REPLA(z$, SALTEX$, 41, 14)
'                    Call REPLA(z$, MKS$(REG&), 61, 4)
'                    Call REPLA(z$, MKI$(KKI%), 59, 2)
'                    '
'                    Call REGAPP("!SELDOCAP", z$)
'                    KQ& = KQ& + 1
'                    '
'                End If
'              End If
'            End If
'          End If
'        End If
'    Next U&
'    '
'    Call CIERRARCH("!SELDOCAP")
    '
    Screen.MousePointer = 1
    If KQ& < 1 Then
        Call COMUNI("No hay Obligaciones Pendientes de Pago\para la Cuenta Elegida.")
        GoTo 10
    End If
    '
105 Call FRESHECHO("!SELDOCAP")
    RGSEL$ = REGSEL$("!SELDOCAP")
    If Len(RGSEL$) <= 4 Then
        GoTo 10
    End If
    REGISEL& = CVS(Mid$(RGSEL$, Len(RGSEL$) - 3, 4))
    '
    VDEF$ = String$(64, 0)
    Call REPLA(VDEF$, MKI$(NC%), 1, 2)
    Call REPLA(VDEF$, Mid$(RGSEL$, 1, 18), 3, 18)
    Call REPLA(VDEF$, MKD$(XVALO(Mid$(RGSEL$, 41, 14))), 21, 8)
    Call REPLA(VDEF$, MKD$(XVALO(Mid$(RGSEL$, 41, 14))), 29, 8)
    Call REPLA(VDEF$, Mid$(RGSEL$, 59, 6), 59, 6)
    '
110 OBX$ = OBJPLA$("CANFACFIJ", VDEF$)
    '
111 If OBX$ = "" Then GoTo 10
    '
    IPEN# = CVD(Mid$(OBX$, 21, 8))
    ICAN# = CVD(Mid$(OBX$, 29, 8))
    REG& = CVS(Mid$(OBX$, 61, 4))
    KKI% = CVI(Mid$(OBX$, 59, 2))
    '
    If TICUEPRO%(NC1%) <> 0 Then
       IPEN# = ICAN#
       REG& = 0
       KKI% = 0
    End If
    '
    If ICAN# < 0.005 Then GoTo 10
    If ICAN# > IPEN# Then
       Call MENSERR(24, "Importe a Cancelar Excede Pendiente.")
       GoTo 110
    End If
    '
    Z$ = Space$(100)
    Call REPLA(Z$, CSTRING$(MKI$(NC%), 1, 5, 0, 2), 1, 5)
    Call REPLA(Z$, rasocli$(NC1%), 8, 30)
    Call REPLA(Z$, Mid$(OBX$, 3, 18), 39, 18)
    Call REPLA(Z$, CSTRING$(MKD$(IPEN#), 3, 12, 2, 2), 59, 12)
    Call REPLA(Z$, CSTRING$(MKD$(ICAN#), 3, 12, 2, 2), 72, 12)
    Call REPLA(Z$, CSTRING$(MKS$(REG&), 3, 8, 0, 2), 89, 8)
    Call REPLA(Z$, CSTRING$(MKI$(KKI%), 1, 4, 0, 2), 97, 4)
    '
    For U& = 1 To FORLIFOF.List2.ListCount
       If Left$(Z$, 70) = Left$(FORLIFOF.List2.LIST(U& - 1), 70) Then
          Call MENSERR(24, "Aplicación Repetida")
          GoTo 105
       End If
    Next U&
    '
    FORLIFOF.List2.AddItem Z$
    FORLIFOF.List2.Refresh
    On Error Resume Next
    List2.TopIndex = List2.ListCount - 16
    On Error GoTo 0

    Call LIFOFIJ4
    '
    If TICUEPRO%(NC1%) = 0 Then
       For U& = REGISEL& + 1 To KQ&
          x$ = REGLEIDO$("!SELDOCAP", U&)
          Call GRAREG("!SELDOCAP", x$, U& - 1)
       Next U&
       KQ& = KQ& - 1
       Call SETULTREG("!SELDOCAP", KQ&)
       If KQ& > 0 Then GoTo 105
    End If
    GoTo 10
    '
999 End Sub
'
Sub LIFOFIJ2()
     '
     FECHALIM% = XVALO(CONTROL$("", "CIEMECOM"))
     '
110  NROPRO& = XVALO(Mid$(FORLIFOF.Label14.Caption, 4))
     HOI% = HOY(HO$)
     '
115  VDEF$ = MKS$(NROPRO&) + MKI$(HOI%) + MKI$(MP%)
     x$ = OBJPLA$("LIFOFIJ-NUE", VDEF$)
     If x$ = "" Then
        Call CIERRARCH("*.*")
        Unload FORLIFOF
        Exit Sub
     End If
     '
     NRO& = CVS(Left$(x$, 4))
     CONDI$ = "OPCIMOVIM='FF' AND NUMECOMP=" & NRO&
     If CantRegistros("CAJABANC", CONDI$, "NOMESS") > 0 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Número Pre-Existente")
        GoTo 110
     End If
     
'     NRO& = CVS(Left$(x$, 4))
'     If NRO& < NROPRO& Then
'         Screen.MousePointer = 11
'         For I& = ULTREG&("CUECORR") To 1 Step -1
'            XXX$ = REGLEIDO$("CUECORR", I&)
'            If UCase$(Mid$(XXX$, 7, 2)) = "FF" Then
'                If Mid$(XXX$, 12, 4) <> "anul" Then
'                   If XVALO(Mid$(XXX$, 19, 6)) = NRO& Then
'                      Screen.MousePointer = 1
'                      Call MENSERR(24, "Número Pre-Existente")
'                      GoTo 110
'                   End If
'                End If
'            End If
'         Next I&
'         Screen.MousePointer = 1
'     End If
     '
     FORLIFOF.Label14.Caption = "FF." + TRIM$(Str$(NRO&))
     FORLIFOF.Label14.Refresh
     '
     FC% = CVI(Mid$(x$, 5, 2))
     If FC% < 7000 Or FC% > HOY(HO$) + 1 Then
         Call MENSERR(24, "Fecha no Válida")
         GoTo 110
     End If
     '
     If FC% <= FECHALIM% Then
         Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
         GoTo 110
     End If
     '
     FEX = FC%
     FORLIFOF.Label13.Caption = FECHATEX$(FEX)
     FORLIFOF.Label13.Refresh
     '
     MP% = CVI(Mid$(x$, 7, 2))
     If MP% < 1 Then
        Call MENSERR(24, "Medio de Pago no Válido")
        GoTo 110
     End If
     Tipo% = MP%
     BANCO$ = ECOARCH$("MEPAGO", MKI$(MP%))
     If BANCO$ = "" Then
        Call MENSERR(24, "Medio de Pago no Válido")
        MP% = 0
        GoTo 110
     End If
     '
     FORLIFOF.Label15.Caption = Tipo%
     FORLIFOF.Label15.Refresh
     FORLIFOF.Label17.Caption = BANCO$
     FORLIFOF.Label17.Refresh
     '
End Sub

Sub LIFOFIJ1()
   '
   FENUE$ = CONTROL$("", "FENUORPA")
   FLIM% = FECHANUM(FENUE$)
   If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
   '
   UNRO& = 0: KK% = 0
   Screen.MousePointer = 11
   CONDI$ = "OPCIMOVIM='FF'"
   UNRO& = XVALO(VALOBADA("CAJABANC", "MAX(NUMECOMP)", CONDI$, "NOMESS"))
'   For I& = ULTREG&("CUECORR") To 1 Step -1
'      Screen.MousePointer = 11
'      x$ = REGLEIDO$("CUECORR", I&)
'      If UCase$(Mid$(x$, 7, 2)) = "FF" Then
'        If Mid$(x$, 10, 4) <> "anul" Then
'          FEMO% = CVI(Mid$(x$, 1, 2))
'          If FEMO% >= FLIM% Then
'            If XVALO(Mid$(x$, 19, 6)) > UNRO& Then
'               UNRO& = XVALO(Mid$(x$, 19, 6))
'            End If
'            KK% = KK% + 1
'            If KK% >= 128 Then GoTo 99
'          End If
'        End If
'      End If
'   Next I&
   '
99 NROPRO& = UNRO& + 1
   Screen.MousePointer = 1
   '
   FORLIFOF.Label14.Caption = "FF." + TRIM$(Str$(NROPRO&))
   FORLIFOF.Label14.Refresh
   '
   Call LIFOFIJ2
   '
End Sub
'


Private Sub verff_Click()
        CONDI$ = "TipoMovim = 'OPAG' AND CODICOM_LAR LIKE 'FF%' Order by FECHEMISI ASC, CODICOM_LAR ASC "
        Call CARGALISEL("!INDIORPA", "CAJABANC", "NUMECOMP/I8;FECHCONTA/D8;FECHEMISI/D0;CODICOM_LAR/A12 ", CONDI$, "DISTINCT/NOMESS")

        ' LA RUTINA DE CARGA DE LA LISTA DE ORDENES DE PAGO SE LLEVO AL FRM ORPAGO07 RUTINA ACTULISTAOP
        Call TRALOCAL("INDIORPA", "")
        '
3       Call SELECHO("!INDIORPA/Indice de Ordenes de Pago Emitidas")
        NPX& = XVALO(VALACT1$("!INDIORPA"))
        If NPX& < 1 Then Exit Sub
        '
        CONDI$ = "TipoMovim = 'OPAG' AND CODICOM_LAR LIKE 'FF%' AND NUMECOMP= " & NPX&
        Stat% = VALOBADA("CAJABANC", "STATUS", CONDI$, "NOMESS")
        If Stat% = 9 Then Call COMUNI("La Orden de Pago de Fondo Fijo" & NPX& & " se Encuentra Anulada")
        '
        TIDOCUM$ = "Liquidacion de Fondo Fijo"
        NUDOCU$ = " " + TRIM$(Str$(NPX&)) + " -"
        DOCUNU& = 0
5       Call MUESTRADOC(DOCUNU&, , TIDOCUM$, NUDOCU$, NUDOCU$)
10      GoTo 3
End Sub


