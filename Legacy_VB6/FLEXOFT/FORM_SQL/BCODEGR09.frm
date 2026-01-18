VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form BCODEGR09 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lectura de Código de Barra"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   10065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   9795
      Begin VB.TextBox Text2 
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
         Index           =   1
         Left            =   1240
         MaxLength       =   6
         TabIndex        =   32
         Text            =   " "
         Top             =   765
         Width           =   780
      End
      Begin VB.CommandButton Command1 
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
         Height          =   330
         Index           =   1
         Left            =   2040
         Picture         =   "BCODEGR09.frx":0000
         TabIndex        =   31
         ToolTipText     =   "Desplegar Lista de Depósitos"
         Top             =   765
         Width           =   175
      End
      Begin VB.TextBox Text3 
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
         Index           =   1
         Left            =   1240
         MaxLength       =   6
         TabIndex        =   30
         Text            =   " "
         Top             =   285
         Width           =   780
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
         Height          =   330
         Index           =   1
         Left            =   2040
         Picture         =   "BCODEGR09.frx":030A
         TabIndex        =   29
         ToolTipText     =   "Desplegar Lista de Sectores"
         Top             =   285
         Width           =   175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito Asociado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   8230
         TabIndex        =   39
         Top             =   220
         Width           =   1140
      End
      Begin VB.Label Label5 
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
         Index           =   1
         Left            =   9000
         TabIndex        =   37
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "A Depósito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   36
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   2235
         TabIndex        =   35
         Top             =   765
         Width           =   7380
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde Sector"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   680
         TabIndex        =   34
         Top             =   220
         Width           =   855
      End
      Begin VB.Label Label8 
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
         Index           =   1
         Left            =   2235
         TabIndex        =   33
         Top             =   285
         Width           =   5220
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   9795
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
         Height          =   330
         Index           =   0
         Left            =   2040
         Picture         =   "BCODEGR09.frx":0614
         TabIndex        =   21
         ToolTipText     =   "Desplegar Lista de Sectores"
         Top             =   765
         Width           =   175
      End
      Begin VB.TextBox Text3 
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
         Index           =   0
         Left            =   1240
         MaxLength       =   6
         TabIndex        =   20
         Text            =   " "
         Top             =   765
         Width           =   780
      End
      Begin VB.CommandButton Command1 
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
         Height          =   330
         Index           =   0
         Left            =   2040
         Picture         =   "BCODEGR09.frx":091E
         TabIndex        =   19
         ToolTipText     =   "Desplegar Lista de Depósitos"
         Top             =   285
         Width           =   175
      End
      Begin VB.TextBox Text2 
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
         Index           =   0
         Left            =   1240
         MaxLength       =   6
         TabIndex        =   18
         Text            =   " "
         Top             =   285
         Width           =   780
      End
      Begin VB.Label Label8 
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
         Index           =   0
         Left            =   2235
         TabIndex        =   27
         Top             =   765
         Width           =   5220
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "A Sector"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   510
         TabIndex        =   26
         Top             =   890
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   2235
         TabIndex        =   25
         Top             =   285
         Width           =   7380
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde Depósito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   520
         TabIndex        =   24
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito Asociado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   8220
         TabIndex        =   23
         Top             =   710
         Width           =   1140
      End
      Begin VB.Label Label5 
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
         Index           =   0
         Left            =   9000
         TabIndex        =   22
         Top             =   765
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "INGRESO MANUAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1060
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   9135
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
         Height          =   330
         Left            =   6120
         Picture         =   "BCODEGR09.frx":0C28
         TabIndex        =   16
         ToolTipText     =   "Desplegar Lista de Lotes"
         Top             =   240
         Width           =   175
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
         Height          =   330
         Left            =   3360
         Picture         =   "BCODEGR09.frx":0F32
         TabIndex        =   15
         ToolTipText     =   "Desplegar Lista Código"
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
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
         Left            =   7440
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text5 
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
         Left            =   4200
         TabIndex        =   2
         Top             =   240
         Width           =   2000
      End
      Begin VB.TextBox Text4 
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
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   2600
      End
      Begin VB.Label Label18 
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
         Left            =   840
         TabIndex        =   38
         Top             =   600
         Width           =   5355
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3810
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   6720
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Habilitar Ingreso Manual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5805
      Picture         =   "BCODEGR09.frx":123C
      TabIndex        =   10
      Top             =   5160
      Width           =   3420
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   6120
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton BOTNEXT 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   9360
      Picture         =   "BCODEGR09.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Continúa Carga de Remito."
      Top             =   5160
      Width           =   650
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   5535
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "BCODEGR09.frx":1850
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   960
      OleObjectBlob   =   "BCODEGR09.frx":1A84
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   3765
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   6641
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorSel    =   14737632
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   3120
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label VALI_BCODE09 
      Caption         =   "VALI_BCODE09"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu mnumenu 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuborrarregselec 
         Caption         =   "Eliminar Item de la Lista a Transferir"
      End
      Begin VB.Menu separa1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransferir 
         Caption         =   "Aprobar Transferencia"
      End
      Begin VB.Menu separa2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "BCODEGR09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOTITFO&
Public DEVOLUCION%

Function SaldoLoteMovisto(CIXI%, LOTE$, DEPO%)
    SALDO = 0
    SQLX$ = "Select SUM(CANTINGRE) AS ACUIN, SUM(CANTSALID) AS ACUSA,COD_INTE,IDENLOTE "
    SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXI%
    SQLX$ = SQLX$ + " AND DepoMovi = " & DEPO%
    SQLX$ = SQLX$ + " AND IDENLOTE = '" & LOTE$ & "'"
    SQLX$ = SQLX$ + " GROUP BY COD_INTE,IDENLOTE"
    
    Set rs = READSET(SQLX$)
    With rs
     SALDO = XVALO(!ACUIN) - XVALO(!ACUSA)
    End With
    SaldoLoteMovisto = SALDO
End Function

Private Sub BOTNEXT_Click()
        '
        ecosal$ = TRIM$(ECOARCH$("TADEPOSI", MKI$(XVALO(Text2(DEVOLUCION%)))))
        ecoing$ = TRIM$(ECOARCH$("TADEPOSI", MKI$(XVALO(Label5(DEVOLUCION%)))))
        If ecosal$ = "" Then
           Call COMUNI("Depósito de No Válido o Inexistente")
           GoTo 99
        End If
        If ecoing$ = "" Then
           Call COMUNI("Sector No Válido o Depósito Asociado Inexistente")
           GoTo 99
        End If
        
        BOTNEXT.Enabled = False
        For U& = 1 To GRID.Rows - 1
            CI% = GRID.TextMatrix(U&, 1)
            CANTSA = XVALO(GRID.TextMatrix(U&, 5))
            If CANTSA < 1 Then
               Call COMUNI("Falta Cantidad: " & CODEXT$(CI%))
               GoTo 99
            End If
            LOTE$ = TRIM$(GRID.TextMatrix(U&, 4))
            If DEVOLUCION% < 1 Then
              DEPOSAL% = XVALO(Text2(DEVOLUCION%))
              DEPOING% = XVALO(Label5(DEVOLUCION%))
            Else
              'INVIERTO
              DEPOSAL% = XVALO(Label5(DEVOLUCION%))
              DEPOING% = XVALO(Text2(DEVOLUCION%))
            End If
            '
            If LOTE$ = "" Then
              CANTILOTE = STODEPOS(CI%, DEPOSAL%)
            Else
              CANTILOTE = SaldoLoteMovisto(CI%, LOTE$, DEPOSAL%)
            End If
            YAENGRILLA = SUMA_CANT_SEGUNDATO(Me.GRID, SAFETEXT$(CI%), 1, 5)    'SUMA SI YA HAY EN GRILLA DEL LOTE Y LE RESTA EL VALOR DE LA FILA ACTUAL
            
            If YAENGRILLA > CANTILOTE Then
               DEPOS$ = TRIM$(ECOARCH$("ECODEP", MKI$(DEPOSAL%)))
               Call COMUNI("STOCKS INSUFICIENTE: " & CODEXT$(CI%) & "\En Depósito: " & DEPOS$ & " - Saldo: " & CANTILOTE)
               GoTo 99
            End If
            
            '
        Next U&
        '
        Label18 = ""
        VALI_BCODE09 = "APROB"
        Hide
99      BOTNEXT.Enabled = True
End Sub

Private Sub Command1_Click(Index As Integer)
    Command1(Index).Enabled = False
    Call SELECHO("TADEPOSI")
    VDEV$ = VALACT1$("TADEPOSI")
    If VDEV$ <> "" Then
       Text2(Index).TEXT = VDEV$
    End If
    Command1(Index).Enabled = True
End Sub

Private Sub Command13_Click()
    '
    Command13.Enabled = False
    If DERACCE%("STPCOBAR", "Configuración de Funcionamiento de Codigos de Barra") = 2 Then
        STPKBN04.Show 1
    End If
    Command13.Enabled = True
    '
End Sub

Private Sub Command2_Click()

  If Command2.Caption = "Habilitar Ingreso Manual" Then
     Command2.Caption = "Habilitar Ingreso por Lecotora"
     Text4 = "": Text5 = "": Text6 = ""
     Text4.Enabled = True: Text5.Enabled = True: Text6.Enabled = True
     Text1 = "": Text1.Enabled = False
     Command3.Enabled = True: Command4.Enabled = True
     Call PINTATEXT(Text4)
     
  Else
     Text4 = "": Text5 = "": Text6 = ""
     Text4.Enabled = False: Text5.Enabled = False: Text6.Enabled = False
     Text1 = "": Text1.Enabled = True
     Command2.Caption = "Habilitar Ingreso Manual"
     Command3.Enabled = False: Command4.Enabled = False
     Call PINTATEXT(Text1)
     
  End If
  
  '
  '
End Sub


Private Sub Command3_Click()
   Command3.Enabled = False
   Call ZELECHOAHP(1, 1)
   COD$ = RESP_ZELE_VECT(2)
   If COD$ <> "" Then
     Text4 = ""
     Text4 = COD$
   End If
   
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    FRMZELECHO.TXTBUSCAR = ""
    CIXI% = CODINT%(Me.Text4)
    If CIXI% < 1 Then Call COMUNI("Falta Ingresar Código"): GoTo 99
    If DEVOLUCION% < 1 Then
      DEPO% = XVALO(Text2(DEVOLUCION%))
    Else
      DEPO% = Me.Label5(DEVOLUCION%)
    End If
    If TRIM$(Label1(DEVOLUCION%)) = "" Then Call COMUNI("Depósito Inexistente o No Válido"): GoTo 99
    '
    Campos$ = "Código/A20;Descripción/A32;Lote/A12;Saldo/F12.1"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , ANTOT)
    FRMZELECHO.Width = 10000
    i& = 0
    SQLX$ = "Select SUM(CANTINGRE) AS ACUIN, SUM(CANTSALID) AS ACUSA,COD_EXTE,IDENLOTE "
    SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXI%
    SQLX$ = SQLX$ + " AND DepoMovi = " & DEPO%
    SQLX$ = SQLX$ + " GROUP BY COD_EXTE,IDENLOTE"

'    SQLX$ = "SELECT SUM(CANTINGRE) AS INGRE,SUM(CANTSALID) AS SALI, IDENLOTE,COD_EXTE"
'    SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK) WHERE COD_INTE = " & CIXI%
'    SQLX$ = SQLX$ + "  AND DEPOMOVI = " & DEPO%
''    SQLX$ = SQLX$ + "  HAVING SUM(CANTINGRE) > SUM(CANTSALID)"
'    SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,IDENLOTE"
'    SQLX$ = SQLX$ + " ORDER BY IDENLOTE"
    Set rs = READSET(SQLX$)
    FRMZELECHO.msf2.Redraw = False
    With rs
      Do While Not .EOF
        SALDO = XVALO(!ACUIN) - XVALO(!ACUSA)
        If SALDO < 0.05 Then GoTo 30
        i& = i& + 1
        
        FRMZELECHO.msf2.Rows = i& + 1
        FRMZELECHO.msf2.TextMatrix(i&, 1) = SAFETEXT$(!Cod_Exte)
        FRMZELECHO.msf2.TextMatrix(i&, 2) = DESCRIT0$(CODINT%(SAFETEXT$(!Cod_Exte)))
        FRMZELECHO.msf2.TextMatrix(i&, 3) = SAFETEXT$(!idenlote)
        FRMZELECHO.msf2.TextMatrix(i&, 4) = FORMATNUM$(SALDO, "F10.1")
30      .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    FRMZELECHO.msf2.Redraw = True
    '
50  FRMZELECHO.Show 1
    
    CI1 = XVALO(RESP_ZELE_VECT(1))
    CODI$ = RESP_ZELE_VECT(2)
    LOTE$ = RESP_ZELE_VECT(3)
    
    Me.Text5 = LOTE$
    Unload FRMZELECHO
99  Command4.Enabled = True
End Sub

Private Sub Command5_Click(Index As Integer)

End Sub

Private Sub Command9_Click(Index As Integer)
    Command9(Index).Enabled = False
    Call SELECHO("SECPROD")
    If VALACT1$("SECPROD") <> "" Then Text3(Index) = VALACT1$("SECPROD")
    Command9(Index).Enabled = True
End Sub



Private Sub Form_Load()
    '
    XXTITU$ = Space$(0)
    '
    VALI_BCODE09 = ""
    '
    If DEVOLUCION% > 0 Then
      Frame2(0).Visible = False
      Frame2(1).Visible = True
      Frame2(1).Top = 0
      Text2(DEVOLUCION%) = Control$("", "DEPOCEN")
    Else
      Frame2(0).Visible = True
      Frame2(1).Visible = False
      Text2(DEVOLUCION%) = Control$("", "DEPOCEN")
    End If
    Campos$ = "CI/F0;Código/A20;Descripción/A32;Lote/A12;Cant./F10.1"
    Call CARGA_ENCABEZADO(GRID, Campos$, BCODEGR09)
    Me.Text4.Enabled = False: Me.Text5.Enabled = False: Me.Text6.Enabled = False
    Me.Command3.Enabled = False: Me.Command4.Enabled = False
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
End Sub

Private Sub GRID_Click()
    If GRID.Rows < 2 Then Exit Sub
    REG& = GRID.MouseRow
    On Error Resume Next
    Call COLOREAR_GRILLA_SOLAFILA(Me.GRID, &HE0E0E0, REG&)
    On Error GoTo 0

End Sub

Private Sub mnuborrarregselec_Click()
    While i& < GRID.Rows - 1
       i& = i& + 1
       If GRID.TextMatrix(i&, 0) = "*" Then
          Call cmdEliminaItem(GRID, i&)
          i& = i& - 1
       End If
    Wend
End Sub

Private Sub mnuconfiglectura_Click()
    If DERACCE%("STPCOBAR", "Configuración de Funcionamiento de Codigos de Barra") = 2 Then
        STPKBN04.Show 1
    End If

End Sub

Private Sub mnuSalir_Click()
  GRID.Rows = 1
  VALI_BCODE09 = ""
  Call CIERRARCH("*.*")
  Screen.MousePointer = 1
  Hide 'Unload Me

End Sub

Private Sub mnuTransferir_Click()
   Call BOTNEXT_Click
End Sub

Private Sub Text1_Change()

'    If Len(TRIM$(Text1)) = 25 Then
'       Call Text1_KeyPress(13)
'    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    '
        'VALIDACIONES
        '
        If KeyAscii <> 13 Then Exit Sub
        If TRIM$(Text1) = "" Then Exit Sub

'        'LECTURA DE POSICIONES CONFIGURADAS
'        If TRIM$(CONTROL("LECCODBA", "OPLECTUR")) = "LONGFIJA" Then
'            PCOD_IN% = XVALO(CONTROL("LECCODBA", "CODIINIC"))
'            PCOD_FIN% = XVALO(CONTROL("LECCODBA", "CODIGFIN"))
'            PCANT_IN% = XVALO(CONTROL("LECCODBA", "CANTINIC"))
'            PCANT_FIN% = XVALO(CONTROL("LECCODBA", "CANTIFIN"))
'            If PCOD_IN% < 1 Or PCOD_FIN% < 1 Or PCANT_IN% < 1 Or PCANT_FIN% < 1 Then
'               Call COMUNI("Falta Información de Posiciones de Lectura")
'               CNFGBARRA.Show 1
'               GoTo 60
'            End If
'        End If
        '
'        If TRIM$(CONTROL("LECCODBA", "OPLECTUR")) = "CARACTER" Then
'           CARCTR$ = TRIM$(CONTROL("LECCODBA", "CARASEPA"))
'           If TRIM$(CARCTR$) = "" Then Call COMUNI("Falta Ingresar Caracter Separador"): CNFGBARRA.Show 1: GoTo 60
'        End If
        '
        'OBTENGO CODIGO INTERNO DE EXTERNO O UPC
'        CODI$ = LEECODIGOBARRA$(Text1, PCOD_IN%, PCOD_FIN%)
'        CANTI = LEECANTIBARRA(Text1, PCANT_IN%, PCANT_FIN%)
        '
        sLecturaCodigoBarras$ = Text1
        sLecturaCodigoBarras$ = REPLACAR(sLecturaCodigoBarras$, "'", "-")
        sLecturaCodigoBarras$ = REPLACAR(sLecturaCodigoBarras$, "ç", "|")
        
        Call TEXTOLOTES(sLecturaCodigoBarras$, "|")
        On Error Resume Next
        For i& = 1 To UBound(CADENATEX$) 'RECORRO EL VECTOR QUE TIENE CUALES SON LAS CELDAS EDITABLES
            Select Case i&
            Case 1
             CODI$ = CADENATEX$(i&)
            Case 2
             LOTE$ = TRIM$(CADENATEX$(i&))
            Case 3
             CANTI = XVALO(CADENATEX$(i&))
            End Select
        Next i&
        On Error GoTo 0

        CI% = CODINT%(CODI$) ' CODIGO INTERNO DE CODIGO EXTERNO
        If CI% < 1 Then
          CI% = CODINTBARRA%(CODI$) ' CODIGO INTERNO DE UPC
        End If
        '
        'VALIDO QUE EL CODIGO SEA VALIDO
        If CI% < 1 Then
           Call MENSERR(24, "Código Inexistente o No Válido")
           Text1.SelStart = 0
           Text1.SelLength = Len(Text1)
           GoTo 60
        End If
        '

        REG& = GRID.Rows
        GRID.Rows = GRID.Rows + 1
        GRID.Row = REG&
        GRID.TextMatrix(REG&, 1) = CI%
        GRID.TextMatrix(REG&, 2) = CODEXT$(CI%)
        GRID.TextMatrix(REG&, 3) = DESCRIT$(CI%)
        GRID.TextMatrix(REG&, 4) = LOTE$
        GRID.TextMatrix(REG&, 5) = FORMATNUM$(CANTI, "F10.1")
        '
50      Text1 = ""
60      Call PINTATEXT(Text1)
    '    '
99 End Sub



Private Sub Text2_Change(Index As Integer)
   Label1(Index).Caption = ""
   If XVALO(Text2(Index)) < 32767 Then Label1(Index).Caption = TRIM$(ECOARCH$("TADEPOSI", MKI$(XVALO(Text2(Index)))))
   
End Sub

Private Sub Text3_Change(Index As Integer)
  Label5(Index).ToolTipText = ""
  CS1 = XVALO(Text3(Index))
  If CS1 < 32767 Then
    Label8(Index).Caption = TRIM$(ECOARCH$("SECPROD", MKS$(CS1)))
    REGI_SECTOR$ = FILTERGETRECORD$("SECPROD", 1, MKS$(CS1))
    DEPODEST% = Asc(Mid$(REGI_SECTOR$, 81, 1))
    Label5(Index) = DEPODEST%
    Label5(Index).ToolTipText = ECOARCH$("TADEPOSI", MKI$(DEPODEST%))
  End If
End Sub

Private Sub Text4_Change()
    Label18 = ""
    CIXI% = CODINT(TRIM$(Text4))
    If CIXI% > 0 Then
      Label18 = DESCRIT0$(CIXI%)
    End If
    
End Sub

Private Sub Text4_DblClick()
   Call Command3_Click
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call PINTATEXT(Text5)
End Sub

Private Sub Text5_DblClick()
   Call Command4_Click
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call PINTATEXT(Text6)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If TRIM$(Text4) = "" Then Exit Sub
    
    Text1 = TRIM$(Text4) & Chr(124)
    Text1 = Text1 & TRIM$(Text5) & Chr(124)
    Text1 = Text1 & TRIM$(Text6)
    Call Text1_KeyPress(13)
    Text4 = ""
    Text5 = ""
    Text6 = ""
  End If
  
End Sub

Private Sub Text6_LostFocus()
   Text6 = FORMATNUM$(XVALO(Text6), "F10.1")
End Sub
