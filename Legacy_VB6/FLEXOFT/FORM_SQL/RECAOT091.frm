VERSION 5.00
Begin VB.Form RECAOT091 
   BackColor       =   &H00FFFFD0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción y Etiquetado"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command23 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6000
      Picture         =   "RECAOT091.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Recepcion de Material en Transito"
      Top             =   2580
      Width           =   650
   End
   Begin VB.CommandButton Command30 
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
      Left            =   3420
      TabIndex        =   33
      Top             =   3120
      Width           =   175
   End
   Begin VB.TextBox Text6 
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
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   32
      Top             =   3120
      Width           =   420
   End
   Begin VB.CommandButton Command18 
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
      Left            =   2280
      TabIndex        =   5
      ToolTipText     =   "Seleccione para qué Solicitud de Fabricación es lo que Recibe"
      Top             =   3135
      Width           =   175
   End
   Begin VB.CommandButton Command4 
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
      Left            =   1770
      TabIndex        =   28
      ToolTipText     =   "Seleccione aquí si recibe de un LABORATORIO"
      Top             =   1215
      Width           =   175
   End
   Begin VB.TextBox Text5 
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
      Left            =   1110
      MaxLength       =   5
      TabIndex        =   27
      ToolTipText     =   "Seleccione aquí si recibe de un LABORATORIO"
      Top             =   1230
      Width           =   675
   End
   Begin VB.CommandButton Command3 
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
      Left            =   2160
      TabIndex        =   25
      Top             =   2295
      Width           =   175
   End
   Begin VB.TextBox Text2 
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
      Left            =   1095
      MaxLength       =   12
      TabIndex        =   24
      Top             =   2295
      Width           =   1155
   End
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
      Height          =   690
      Left            =   6000
      Picture         =   "RECAOT091.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Continuar a la Siguiente Pantalla"
      Top             =   3105
      Width           =   650
   End
   Begin VB.TextBox Text9 
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
      Left            =   870
      TabIndex        =   4
      ToolTipText     =   "Seleccione para qué Solicitud de Fabricación es lo que Recibe"
      Top             =   3135
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
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
      Left            =   5535
      TabIndex        =   16
      Top             =   570
      Width           =   175
   End
   Begin VB.TextBox Text4 
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
      Left            =   3450
      MaxLength       =   18
      TabIndex        =   2
      ToolTipText     =   "Este Campo es OBLIGATORIO"
      Top             =   2310
      Width           =   2280
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
      Height          =   285
      Left            =   3450
      MaxLength       =   24
      TabIndex        =   3
      ToolTipText     =   "Este Campo es OPCIONAL"
      Top             =   2625
      Width           =   2280
   End
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
      Left            =   1110
      MaxLength       =   5
      TabIndex        =   0
      ToolTipText     =   "Seleccione aquí si recibe de una DROGUERIA"
      Top             =   1680
      Width           =   675
   End
   Begin VB.CommandButton Command5 
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
      Left            =   1770
      TabIndex        =   1
      ToolTipText     =   "Seleccione aquí si recibe de una DROGUERIA"
      Top             =   1665
      Width           =   175
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   2620
      Left            =   5895
      ScaleHeight     =   2565
      ScaleWidth      =   915
      TabIndex        =   7
      Top             =   -120
      Width           =   975
      Begin VB.CommandButton AGenRep 
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
         Left            =   110
         Picture         =   "RECAOT091.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   840
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Height          =   650
         Left            =   110
         Picture         =   "RECAOT091.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1440
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Left            =   110
         Picture         =   "RECAOT091.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "RECAOT091.frx":0EAA
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   2160
         Width           =   1515
      End
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   2670
      TabIndex        =   39
      Top             =   1995
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Ltota 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "C. Imp:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   7
      Left            =   1935
      TabIndex        =   38
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFD0&
      Caption         =   "(Recepcion y Etiquetado)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   600
      TabIndex        =   37
      Top             =   525
      Width           =   2940
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dep:"
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
      Left            =   2160
      TabIndex        =   35
      Top             =   3200
      Width           =   810
   End
   Begin VB.Label Label38 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3660
      TabIndex        =   34
      Top             =   3135
      Width           =   2085
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Laboratorio:"
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
      Left            =   60
      TabIndex        =   30
      Top             =   1305
      Width           =   1035
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
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
      Left            =   2070
      TabIndex        =   29
      Top             =   1230
      Width           =   3645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido de Suministro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   60
      TabIndex        =   26
      Top             =   2160
      Width           =   945
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFD0&
      Caption         =   "Depósito Caótico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   240
      TabIndex        =   22
      Top             =   120
      Width           =   2940
   End
   Begin VB.Label Label29 
      BackColor       =   &H00E0E0E0&
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
      Left            =   840
      TabIndex        =   21
      ToolTipText     =   "Cantidad nominal"
      Top             =   3495
      Width           =   4890
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
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
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   3200
      Width           =   705
   End
   Begin VB.Label Label34 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "I-R Numero:"
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
      Left            =   3465
      TabIndex        =   18
      Top             =   330
      Width           =   1320
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Left            =   4830
      TabIndex        =   17
      Top             =   255
      Width           =   885
   End
   Begin VB.Label Label24 
      BackColor       =   &H00E0E0E0&
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
      Left            =   2070
      TabIndex        =   15
      Top             =   1680
      Width           =   3645
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remito:"
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
      Left            =   2565
      TabIndex        =   14
      Top             =   2400
      Width           =   810
   End
   Begin VB.Label Label28 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4500
      TabIndex        =   13
      Top             =   585
      Width           =   1035
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   3765
      TabIndex        =   12
      Top             =   660
      Width           =   615
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Factura: "
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
      Left            =   2370
      TabIndex        =   11
      Top             =   2715
      Width           =   1065
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor:"
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
      Left            =   120
      TabIndex        =   10
      Top             =   1755
      Width           =   945
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha:"
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
      Left            =   1470
      TabIndex        =   9
      Top             =   4830
      Width           =   1335
   End
   Begin VB.Label Label18 
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
      Height          =   300
      Left            =   1470
      TabIndex        =   8
      Top             =   4830
      Width           =   1335
   End
End
Attribute VB_Name = "RECAOT091"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VIENEDETEXT2%

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub
Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

'
Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   '
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Label28 = VDEV$
   End If
   '
   Command10.Enabled = True
End Sub

Private Sub Command18_Click()
   
   Command18.Enabled = False
   '
   Call SELECHO("ECORDEF")
   VDEV$ = VALACT1$("ECORDEF")
   '
   If VDEV$ <> "" Then
      Text9 = VDEV$
   End If
   '
   Command18.Enabled = True
   
End Sub

Private Sub Command2_Click()
   '
   Command2.Enabled = False
   If XVALO(Text1) > 0 And XVALO(Text5) > 0 Then
      Call MENSERR(24, "Cliente y Proveedor no Compatible")
      GoTo 99
   End If
   '
   If RASOCLI(XVALO(Text5)) = "" And RASOCLI((-1) * XVALO(Text1)) = "" Then
      If TRIM$(Label24) = "" Then
         COMUNI ("Proveedor Inexistente")
         On Error Resume Next
         Text1.SetFocus
         On Error GoTo 0
         GoTo 99
      End If
      If TRIM(Label3) = "" Then
         COMUNI ("Cliente Inexistente")
         Text5.SetFocus
         GoTo 99
      End If
   End If
   '
   If TRIM$(Text4) = "" Then
      COMUNI ("Falta Ingresar Remito")
      Text4.SetFocus
      GoTo 99
   End If
   ' Validacion de ingreso de O/C en Carpetas de Importacion.
   If CODCARPE$ <> "" Then
      If TRIM$(Text2) = "" Then
          COMUNI ("Falta Ingresar Pedido de Suministro")
          Text3.SetFocus
          GoTo 99
      End If
   End If
   '
    DEPENTRA% = XVALO(Text6)
    If DEPENTRA% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPENTRA%)) = "" Then
       Call MENSERR(24, "Deposito no Válido")
       GoTo 99
    End If
   '
   DEPOTRANS% = XVALO(Control$("BOLRECEP", "DEPTRANS")) ' DEPOSITO MATERIAL EN TRANSITO
   If DEPENTRA% = DEPOTRANS% Then
        COMUNI ("Depósito de Entrada No Válido\Configurado Como Material en Transito")
       GoTo 99
   End If

   IDCOMP$ = TRIM(Text2)
   If IDCOMP$ <> "" Then
      ESTAORCOM% = XVALO(VALOBADA("OCOMENCA", "STAT_Ocom", "ID_COMPROB='" & IDCOMP$ & "'"))
      If ESTAORCOM% >= 9 Then
        Call MENSERR(24, "Imposible Procesar.\Documento '" & IDCOMP$ & "' Figura Anulado")
        GoTo 99
      End If
   End If
   '
   If TRIM$(Text9) <> "" Then
      If ECOARCH$("ECORDEF", TRIM$(Text9)) = "" Then
         COMUNI ("Destino Inexistente")
         Text9.SetFocus
         GoTo 99
       End If
   End If
   '
   RECAOT092.Show 1
   If CODCARPE$ <> "" Then
      Unload Me
   End If
   '
99 Command2.Enabled = True

End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   '
   If TRIM$(Text4) = "" Then
      COMUNI ("Falta Ingresar Nº de Remito")
      Text4.SetFocus
      GoTo 99
   End If
   '
   DEPENTRA% = XVALO(Text6)
   If DEPENTRA% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPENTRA%)) = "" Then
       Call MENSERR(24, "Deposito no Válido")
       GoTo 99
   End If
   '
   DEPOTRANS% = XVALO(Control$("BOLRECEP", "DEPTRANS")) ' DEPOSITO MATERIAL EN TRANSITO
   If DEPENTRA% = DEPOTRANS% Then
        COMUNI ("Depósito Seleccionado No Válido\Configurado como Material en Transito")
       GoTo 99
   End If
   '
   On Error Resume Next
   RECMATRAN09.Label13 = TRIM$(Label33)    'numero de recepcion
   RECMATRAN09.Label10 = TRIM$(Text4)      'numero de remito
   RECMATRAN09.Label16 = DEPENTRA%         'deposito
   RECMATRAN09.Label20 = ECOARCH$("TADEPOSI", Chr$(DEPENTRA%))
   '
   RECMATRAN09.Show 1
   '
   On Error GoTo 0
99 Command23.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   '
   ' DEPURA ORDENS DE COMPRA
   Dim OCOENCA As ADODB.Recordset
   Set OCOENCA = New ADODB.Recordset
   CONDI0$ = "ID_COMPROB <> '' AND Stat_Ocom < 3 "
   NCLI% = XVALO(Text5): NPRO% = XVALO(Text1)
   If NCLI% > 0 Then
       CONDI0$ = CONDI0$ + " AND NUME_CLI = " & NCLI% & " "
     ElseIf NPRO% > 0 Then
       CONDI0$ = CONDI0$ + " AND NPRO_OCOM = " & NPRO% & " "
   End If
   CONDI$ = CONDI0$ & " ORDER BY FGEN_OCOM ASC"
   sqlx$ = "SELECT * FROM OCOMENCA WITH(NOLOCK) WHERE " & CONDI$
   Set OCOENCA = FLEXCONN.Execute(FILTSQL$(sqlx$))
   '
   With OCOENCA
     Do While Not .EOF
       IDOC$ = SAFETEXT$(!ID_COMPROB)
       NUOC = XVALO(!Nume_Ocom)
       NPED = XVALO(!Nume_Pedi)
       If NUOC > 0 Then
           QTR% = CantRegistros("OCOMDETA", "Nume_Ocom = " & NUOC & " AND (CANT_OCOM - CANT_REC) > 0.00005")
         ElseIf NPED > 0 Then
           QTR% = CantRegistros("OCOMDETA", "Nume_Pedi = " & NPED & " AND (CANT_OCOM - CANT_REC) > 0.00005")
       End If
       If QTR% < 1 Then
          Call ACTUREGISTRO("OCOMENCA", "Stat_Ocom", CONDI0$ & " AND ID_COMPROB = '" & IDOC$ & "'", 3, RAFE&)
       End If
     .MoveNext
     Loop
   End With
   OCOENCA.Close
   Set OCOENCA = Nothing
   '
   If CODCARPE$ = "" Then
        Call CARGALISEL("!PEDSUMI", "OCOMENCA", "ID_COMPROB/A16;Refe_OCom/A48", CONDI$)
        Call SELECHO("!PEDSUMI")
        VIENEDETEXT2% = 1
        Text2 = TRIM(VALACT1("!PEDSUMI"))
        VIENEDETEXT2% = 0
   Else
        CONDI$ = CONDI0$ + " AND CODCARPIMP = '" & CODCARPE$ & "'"
        CONDI$ = CONDI$ + " ORDER BY FGEN_OCOM ASC "
        Call CARGALISEL("!PEDSUMI", "OCOMENCA", "ID_COMPROB/A16;Refe_OCom/A48", CONDI$)
        Call SELECHO("!PEDSUMI")
        VIENEDETEXT2% = 1
        Text2 = TRIM(VALACT1("!PEDSUMI"))
        VIENEDETEXT2% = 0
   End If
   '
   Command3.Enabled = True
End Sub

Private Sub Command30_Click()
   Call SELECHO("TADEPOSI")
   NUMDEP% = XVALO(VALACT1$("TADEPOSI"))
   Text6 = VALACT1$("TADEPOSI")
End Sub

Private Sub Command4_Click()
  Command4.Enabled = False
  '
  Call SELECHO("DEUDOR1")
  NPROV$ = VALACT1$("DEUDOR1")
    
  If NPROV$ <> "" Then
     Text5.TEXT = NPROV$
  End If
  '
  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
  Command5.Enabled = False
  '
  Call SELECHO("PROVED1")
  NPROV$ = VALACT1$("PROVED1")
    
  If NPROV$ <> "" Then
     Text1.TEXT = NPROV$
     Call Text1_KeyPress(13)
  End If
  '
  Command5.Enabled = True
End Sub

Private Sub Form_Activate()
   COMP$ = Control("", "CODMBREC")
   If TRIM$(COMP$) = "" Then COMP$ = "BR"
   Label33 = TRIM$(Str$(ProNroComprobante(COMP$))) 'Proximo Nro Comprobante
End Sub

Private Sub Form_Load()
   HOII% = HOY(HOS$)
   COMP$ = Control("", "CODMBREC")
   If TRIM$(COMP$) = "" Then COMP$ = "BR"
   Label28 = HOS$
   Label33 = TRIM$(Str$(ProNroComprobante(COMP$))) 'Proximo Nro Comprobante
   '
   Dim RECUTIMA As ADODB.Recordset
   Set RECUTIMA = New ADODB.Recordset
   Set RECUTIMA = FLEXCONN.Execute(FILTSQL$("SELECT IDENLOTE FROM LOTINGRE WHERE TIMATE IS NULL"))
   With RECUTIMA
      Do While Not .EOF
        IDLO$ = !idenLote
        CIXI% = XVALO(VALOBADA("LOTINGRE", "Cod_Inte", "IdenLote = '" & IDLO$ & "'"))
        TIMA% = TIMATE%(CIXI%)
        Call ACTUREGISTRO("LOTINGRE", "Timate", "IdenLote = '" & IDLO$ & "'", TIMA%, RAFE&)
      .MoveNext
      Loop
    End With
    '
    Call GENLISORF    'Genera lista de seleccion de ordenes de produccion
    '
    End Sub

Private Sub Image2_DblClick()
   '
   CONDIX$ = " between '01/mar/2011' and '31/mar/2011'"
   '
   CONDI$ = "fechgen" + CONDIX$
   QT1 = CantRegistros&("ORDEFABR", CONDI$, "NOMESS")

   CONDI$ = "codimovi='ir' and fechmov " + CONDIX$
   QT2 = CantRegistros&("MOVISTO", CONDI$, "NOMESS")


   CONDI$ = "canconsu < canalta"
   QT3 = CantRegistros("LOTINGRE", CONDI$, "NOMESS")

   CONDI$ = "fechmov" + CONDIX$
   QT4 = CantRegistros("MOVISTO", CONDI$, "NOMESS")
   '
   MsgBox QT1 & "   " & QT2 & "   " & QT3 & "   " & QT4
   '
End Sub

Private Sub Text1_Change()
 
   If XVALO(Text1) < 1 Then
      Label24 = ""
      Exit Sub
   End If
   '
   If XVALO(Text1) > 32767 Then
      Text1.SetFocus
      Exit Sub
   End If
   '
   If XVALO(Text1) > 0 Then
      CPROV% = XVALO(Text1)
      Label24 = RASOCLI(-CPROV%)
   End If
   
End Sub

Private Sub Text1_DblClick()
    Call Command5_Click
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
       Call Command5_Click
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text4.SetFocus
   End If
End Sub

Private Sub Text1_LostFocus()
   '
   Call Text4_LostFocus
   '
End Sub

Private Sub Text2_Change()
   IDCOMP$ = TRIM(Text2)
   '
   If TRIM$(Label7) = "" Then
     Text1.Enabled = True
     Text5.Enabled = True
     Command4.Enabled = True
     Command5.Enabled = True
   End If
   '
   If IDCOMP$ <> "" Then
      Text5 = XVALO(VALOBADA("OCOMENCA", "NUME_CLI", "ID_COMPROB='" & IDCOMP$ & "'"))
      Text1 = XVALO(VALOBADA("OCOMENCA", "Npro_Ocom", "ID_COMPROB='" & IDCOMP$ & "'"))
      Text9 = SAFETEXT(VALOBADA("OCOMENCA", "ANEX_OCOM", "ID_COMPROB='" & IDCOMP$ & "'"))
      '
      Text1.Enabled = False
      Text5.Enabled = False
      Command4.Enabled = False
      Command5.Enabled = False
   End If
End Sub


Private Sub Text3_GotFocus()
   Text3.SelStart = 0
   Text3.SelLength = Text3.MaxLength
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text9.SetFocus
   End If
End Sub

Private Sub Text4_GotFocus()
   '
   Text4.SelStart = 0
   Text4.SelLength = Text4.MaxLength
   '
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text3.SetFocus
   End If
End Sub

Private Sub Text4_LostFocus()
  'Valida Remito
  If Text4 <> "" And Text1 <> "" Then
     nroprov% = XVALO(TRIM$(Text1))
     CONDI$ = "DOSECLAR = '" & TRIM$(Text4) & "' AND NUME_PROV=" & nroprov%
      
     'CONDI$ = "DOCUINGRE = '" & TRIM$(Text4) & "' AND ORIGEN='" & RASOCLI$(NROPROV%) & "'"
     CantRem% = CantRegistros("MOVISTO", CONDI$)
     '
     If CantRem% > 0 Then
        'Traer Remito registrado a Pantalla
        COMUNI ("Remito ya Ingresado")
        NROLOTE$ = VALOBADA("MOVISTO", "IDENLOTE", CONDI$)
        Label33 = XVALO(TRIM$(Mid$(NROLOTE$, 4, 6)))
        NPROV% = VALOBADA("MOVISTO", "Nume_Prov", "IDENLOTE ='" & NROLOTE$ & "'")
        Text1 = NPROV%
        CONDI$ = "IDENLOTE='" & NROLOTE$ & "'"
        FF$ = VALOBADA("LOTINGRE", "FECHALTA", CONDI$)
        Num% = CVINT(FF$)
        FechFinal$ = FECHATEX(Num%)
        Label28 = FechFinal$
        Text9 = VALOBADA("LOTINGRE", "Cod_Desti", CONDI$)  'Trae el Destino
       Else
        HOII% = HOY(HOS$)
        Label28 = HOS$
        COMP$ = Control("", "CODMBREC")
        If TRIM$(COMP$) = "" Then COMP$ = "BR"
        Label33 = TRIM$(Str$(ProNroComprobante(COMP$))) 'Proximo Nro Comprobante
     End If
   
  End If
   
End Sub

Sub LIMPTEXT()
    Text1 = ""
    COMP$ = Control("", "CODMBREC")
    If TRIM$(COMP$) = "" Then COMP$ = "BR"
    Label33 = TRIM$(Str$(ProNroComprobante(COMP$))) 'Proximo Nro Comprobante
    HOII% = HOY(HOS$)
    Label28 = HOS$
    Text4 = ""
    Text3 = ""
    Text9 = ""
    Text2 = ""
    Label29 = ""
    Text5 = ""
    Text6 = ""
    Text1.Enabled = True
    Command5.Enabled = True
    Text5.Enabled = True
    Command4.Enabled = True
End Sub


Private Sub Text5_Change()
   If XVALO(Text5) < 1 Then
      Label3 = ""
      Exit Sub
   End If
   '
   If XVALO(Text5) > 32767 Then
      Text5.SetFocus
      Exit Sub
   End If
   '
   If XVALO(Text5) > 0 Then
      NC% = XVALO(Text5)
      Label3 = RASOCLI(NC%)
   End If
End Sub

Private Sub Text6_Change()
   Label38 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text6)))
End Sub

Private Sub Text6_DblClick()
   Call Command30_Click
End Sub

Private Sub Text9_Change()
    '
    SOFA$ = TRIM$(Text9)
    DES$ = ECOARCH$("ECORDEF", SOFA$)
    Label29 = DES$
    If Text9 <> "" Then
      If VIENEDETEXT2% = 0 Then
        Text2 = SAFETEXT(VALOBADA("OCOMENCA", "ID_COMPROB", "ANEX_OCOM='" & SOFA$ & "' AND STAT_OCOM < 8"))
      End If
    End If
    '
End Sub

Private Sub Text9_DblClick()
    Call Command18_Click
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
       Call Command18_Click
    End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Command2.SetFocus
   End If
End Sub
Private Sub Command14_Click()
  
  If DERACCE%("CFGRECE", "Configuración de Recepción") > 1 Then
     Command14.Enabled = False
     CFGRECEP.Show 1
     Command14.Enabled = True
  End If
   
End Sub

