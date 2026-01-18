VERSION 5.00
Begin VB.Form MARCH_GES04 
   BackColor       =   &H00D0F0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Archivos Maestros - Actualización y Consulta"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Height          =   4335
      Left            =   4800
      ScaleHeight     =   4275
      ScaleWidth      =   915
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      Begin VB.Image Image3 
         Height          =   4500
         Left            =   0
         Picture         =   "MARCH_GES04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   975
      End
   End
   Begin VB.ListBox SORTLIST 
      Height          =   255
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   32
      Top             =   3885
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   540
      Left            =   2205
      TabIndex        =   30
      Top             =   105
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   330
      Left            =   105
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   225
      Left            =   1155
      TabIndex        =   28
      Top             =   105
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   330
      Left            =   3465
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404080&
      Height          =   4335
      Left            =   4800
      ScaleHeight     =   4275
      ScaleWidth      =   915
      TabIndex        =   19
      Top             =   0
      Width           =   975
      Begin VB.CommandButton Command2 
         Caption         =   "QC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   105
         Picture         =   "MARCH_GES04.frx":2398
         TabIndex        =   25
         ToolTipText     =   "Sub-Sistema de Gestión de Calidad"
         Top             =   1680
         Width           =   650
      End
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
         Height          =   630
         Left            =   105
         Picture         =   "MARCH_GES04.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Importación / Conversión de Datos"
         Top             =   2415
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   100
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   22
         Top             =   3230
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   23
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   100
         Picture         =   "MARCH_GES04.frx":262C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   960
         Width           =   650
      End
      Begin VB.CommandButton Command24 
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
         Height          =   640
         Left            =   100
         Picture         =   "MARCH_GES04.frx":2936
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   315
         Width           =   650
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   105
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "MARCH_GES04.frx":2A80
         Top             =   3520
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0FF&
      Caption         =   "Actualización Bases de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4425
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   225
         Left            =   2730
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Materiales, Fórmulas y Rutas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   210
         TabIndex        =   9
         Top             =   420
         Width           =   4005
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   10
            Top             =   420
            Width           =   2325
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
               Left            =   0
               Picture         =   "MARCH_GES04.frx":49A2
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Maestro de Artí- culos de Stock"
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
               Left            =   840
               TabIndex        =   12
               Top             =   105
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   2
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command4 
            Caption         =   "Inventarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   8
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Negativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   7
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Consignación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   6
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Críticos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   5
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Estadísticas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   4
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Análisis A-B-C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   3
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Clientes / Proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   210
         TabIndex        =   1
         Top             =   1890
         Width           =   4005
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   16
            Top             =   840
            Width           =   2325
            Begin VB.CommandButton Command18 
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
               Picture         =   "MARCH_GES04.frx":4CAC
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Clientes en Cuenta Corriente"
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
               TabIndex        =   18
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   13
            Top             =   420
            Width           =   2325
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
               Left            =   -20
               Picture         =   "MARCH_GES04.frx":50EE
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Proveedores y Subcontratistas"
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
               TabIndex        =   15
               Top             =   0
               Width           =   1590
            End
         End
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Archi 
      Caption         =   "Archivos Maestros"
      Begin VB.Menu Mastok 
         Caption         =   "Materiales, Fórmulas, Rutas"
      End
      Begin VB.Menu PLH1 
         Caption         =   "-"
      End
      Begin VB.Menu Amacli 
         Caption         =   "Clientes - Cuentas Corrientes"
      End
      Begin VB.Menu Proveed 
         Caption         =   "Proveedores / Contratistas"
      End
      Begin VB.Menu PLH2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu cal 
      Caption         =   "Cálidad"
      Begin VB.Menu MNUSSCAL 
         Caption         =   "SUB-SISTEMA DE GESTION DE CALIDAD"
      End
   End
   Begin VB.Menu mnuimp 
      Caption         =   "Importador"
      Begin VB.Menu mnuimportar 
         Caption         =   "Importación / Conversión de Datos"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "MARCH_GES04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CUEBALRECI%, CUEBALPAGO%, CUEBALAJUS%, CUEBALCART%
'
Private Sub Amacli_Click()
   Call Command18_Click
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   If EXISTE%("AMASTO04.EXE") > 0 Then
       Call CONECRUN("AMASTO04", "Maestro de Items de Stock")
     Else
       AMAS_GES04.Show 1
   End If
   Command1.Enabled = True
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
      ACLI_GES04.Show 1
    Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    ACRE_GES04.Show 1
    'Call CONECRUN("AMAPRO", "Actualización de Base de Datos de Proveedores")
    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
 Command2.Enabled = False
    If EXISTE%("ISO9K04.EXE") > 0 Then
       Call CONECRUN("ISO9K04", "SUB-SISTEMA DE GESTION DE CALIDAD")
    End If
    Command2.Enabled = True
End Sub

Private Sub Command21_Click()
    Command21.Enabled = False
    RTA% = ALTERNA%("Importación desde Planillas Excel.\Conversión de 'DAT' a 'MDB'.")
    If RTA% = 1 Then
        IMPOR04.Show 1
      ElseIf RTA% = 2 Then
        Call CONVBADA 'CONVERSION DE .DAT A .MDB
    End If
    Command21.Enabled = True
End Sub

Private Sub Command24_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FINAL("")
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("ARCHIMA")
End Sub

Private Sub Command3_Click()
   Call APERBASDAT("CABAN")
   With CajaBanc
      .MoveFirst
      Do While Not .EOF
         If Mid$(!CodiCom_Lar, 6, 4) = "anul" Then
            If Abs(!ValAbsComp) > 0 Then
               AA2 = !NuProve
               AAA = !ValAbsComp
               .Edit
               !ValAbsComp = 0
               !Idebecomp = 0
               !IHABECOMP = 0
               !BASEIMPO = 0
               .Update
            End If
         End If
      .MoveNext
      Loop
   End With
   
End Sub

Private Sub Command6_Click()
   '
   ' ESTO ES PARA RECUPERAR MOVIMIENTOS 'AS' EN FERVI
   '
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE NuProve > 0 "
   SQLX$ = SQLX$ + "AND CodiCom_Lar like 'AS-*' "
   Set Cabasel1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With Cabasel1
     On Error Resume Next
130    .MoveFirst
       If Err = 0 Then
         .Delete
         GoTo 130
       End If
     On Error GoTo 0
   End With
   '
   SQLX$ = "SELECT * FROM SACREPEN "
   SQLX$ = SQLX$ + "WHERE NuProve > 0 "
   SQLX$ = SQLX$ + "AND CodiCom_Lar like 'AS-*' OR CodiCom_Lar like 'FF-*' "
   Set Cabasel1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With Cabasel1
     On Error Resume Next
140    .MoveFirst
       If Err = 0 Then
         .Delete
         GoTo 140
       End If
     On Error GoTo 0
   End With
   '
   Call GRAFACRE("AS")
   '
End Sub

Sub MIGRACHECAR()
   '
' VERSION DEL 30 MAYO 2007
   Dim IDECH$(), IMPOX#(), MARCA1%(), MARCA2%(), MARCA3%()
   HOII% = HOY(HOS$)
   CUECARTE% = CUECOINT%(Control$("", "CUECHECA"))
   '
   Call COPYSTRU("BADACHE", "BADACHE1")
   ABC1% = COPYFILE%(LUDAT$ + "BADACHE.DAT", LUDAT$ + "BADACHE1.DAT")
   '
   FIN& = ULTREG&("BADACHE")
   ReDim IDECH$(FIN&), IMPOX#(FIN&), MARCA1%(FIN&), MARCA2%(FIN&), MARCA3%(FIN&)
   For U& = 1 To FIN&
      Label1 = U&
      Label1.Visible = True
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("BADACHE", U&)
      IDEVA$ = TRIM$(Left$(XX$, 40))
         IDECH$(U&) = IDEVA$
      Import# = CVD(Mid$(XX$, 49, 8))
         IMPOX#(U&) = Import#
'      FEMICHE% = CVI(Mid$(XX$, 57, 2))
'      FEVECHE% = CVI(Mid$(XX$, 59, 2))
'      FRECEP% = CVI(Mid$(XX$, 61, 2))
'      DORECEP$ = TRIM$(Mid$(XX$, 63, 10))
'      ORICHE$ = TRIM$(Mid$(XX$, 73, 30))
'      DESTI$ = TRIM$(Mid$(XX$, 103, 30))
'      FENTREGA% = CVI(Mid$(XX$, 133, 2))
'      OB1$ = TRIM$(Mid$(XX$, 135, 40))
'      OB2$ = TRIM$(Mid$(XX$, 175, 40))
'      OB3$ = TRIM$(Mid$(XX$, 215, 40))
         MARCA1%(U&) = 0
         MARCA2%(U&) = 0
         MARCA3%(U&) = 0
   Next U&
   '
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
   Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   SQLY$ = "SELECT * FROM CAJABANC "
   SQLY$ = SQLY$ + "WHERE TipoMovim = 'CHENTR' ORDER BY DescriMov ASC"
   Set chetemp2 = CueCorri.OpenRecordset(SQLY$, dbOpenDynaset)
   '
NOFIGU% = 0
   With CHETEMP1
     .MoveFirst
     Do While Not .EOF
       IDEVA$ = !Descrimov
       If !CueContaI <> CUECARTE% Then GoTo 19
       For JJ& = 1 To FIN&
         If IDECH$(JJ&) = IDEVA$ Then
           If MARCA1%(JJ&) = 0 Then
              MARCA1%(JJ&) = 1
              GoTo 19
           End If
         End If
       Next JJ&
       '
       If CVINT%(!FeDeposEnt) < 1 Then
         AA3$ = ""
         On Error Resume Next
         AA3$ = !Document
         On Error GoTo 0
         If AA3$ = "" Then
           AA4$ = ""
           On Error Resume Next
           AA4$ = !DestiCheq
           On Error GoTo 0
           If AA4$ = "" Then
             Import# = !Idebecomp
             ABCX1$ = IDEVA$
             GoSub BUSCALASALIDA
             If ENCONTROLASALIDA% > 0 Then
                GoTo 19
             End If
             '
             FENTREGA% = CVINT%(!FechConta) ' Si no en BADACHE, sale misma fecha que entró
             FEMICHE% = CVINT%(!FechEmisi)
             DOCUEN$ = "Mig.Datos"
             DESTI$ = "Cuenta Balance Cartera"
             OB1$ = "Conversión Datos DAT-MDB "
             .Edit
             !Status = 2
             !FeDeposEnt = CVDAT(FENTREGA%)
             !FECHENTRE = CVDAT(FENTREGA%)
             !Document = DOCUEN$
             !DestiCheq = DESTI$
             !REFESALCHE = OB1$
             !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
             !UsuaSalChe = OB3$ ' AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
             .Update
             '
             GoSub AGREGASALIDA
             GoTo 19
           End If
         End If
       End If
       '
19   .MoveNext
     Loop
   End With
   Exit Sub
   '
   ' **********************************************************************
BUSCALASALIDA:
   ENCONTROSALIDA% = 0
   With chetemp2
     .FindFirst "DescriMov = '" & IDEVA$ & "' " ' and VariMovim = '1' "
37   If .NoMatch = False Then
        If Abs(Import# - IHABECOMP) < 0.005 Then
          ENCONTROLASALIDA% = 1
          GoTo 85
        End If
        .FindNext "DescriMov = '" & IDEVA$ & "' "
        GoTo 37
     End If
  End With
'              .Edit
'              !VariMovim = "2"
'              .Update
'              GoTo 29    ' LO ENCONTRO
85 Return
   '
   ' **********************************************************************
AGREGASALIDA:
    '
    With chetemp2
      .AddNew
      !CodiCom_Cor = "MIGRA"
      !CodiCom_Lar = "MIGRA"
      !REFECOM = DESTI$
      !FechConta = CVDAT(FENTREGA%)
      !FechAcred = CVDAT(FENTREGA%)
      !CueContaI = CUECARTE%
      !CENCOSTO = ""
      !ItemConc = ""
      !NUCLIEN = 0
      !NuProve = 0
      !TipoMovim = "CHENTR"
      !OPCIMOVIM = "CH"
      !VARIMOVIM = "1"
      !Descrimov = IDEVA$
      !FechEmisi = CVDAT(FEMICHE%)
      !CodiEmpr = 0
      !CodiSucu = 0
      !NUMETALO = 0
      !NumeComp = 0
      !Status = 1
        On Error Resume Next
      !Mone_Emis = 1
      !Tipo_Cam = 1
        On Error GoTo 0
      !ValAbsComp = Abs(Import#)
      !Idebecomp = 0
      !IHABECOMP = Import#
      !Impo_Dola = (-1) * Import#
      !BASEIMPO = Abs(Import#)
      !NROCAI = ""
        On Error Resume Next
      !IncSubIva = ""
        On Error GoTo 0
      !FEVTOCAI = CVDAT(0)
      !OBSERGEN = ""
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now
      !NuAsiDiar = ""
     .Update
     '
     .AddNew
      !CodiCom_Cor = "MIGRA"
      !CodiCom_Lar = "MIGRA"
      !REFECOM = DESTI$
      !FechConta = CVDAT(FENTREGA%)
      !FechAcred = CVDAT(FENTREGA%)
      !CueContaI = CUEBALCART%
      !CENCOSTO = ""
      !ItemConc = ""
      !NUCLIEN = 0
      !NuProve = 0
      !TipoMovim = "MIGRA"
      !OPCIMOVIM = ""
      !VARIMOVIM = ""
      !Descrimov = IDEVA$
      !FechEmisi = CVDAT(FEMICHE%)
      !CodiEmpr = 0
      !CodiSucu = 0
      !NUMETALO = 0
      !NumeComp = 0
      !Status = 0
        On Error Resume Next
      !Mone_Emis = 1
      !Tipo_Cam = 1
        On Error GoTo 0
      !ValAbsComp = Import#
      !Idebecomp = Import#
      !IHABECOMP = 0
      !Impo_Dola = Import#
      !BASEIMPO = Abs(Import#)
      !NROCAI = ""
        On Error Resume Next
      !IncSubIva = ""
        On Error GoTo 0
      !FEVTOCAI = CVDAT(0)
      !OBSERGEN = ""
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now
      !NuAsiDiar = ""
      .Update
   End With
Return
'
'
End Sub
'
Sub MIGRACARCHE()
   
   Dim IDECH$()
   HOII% = HOY(HOS$)
   CUECARTE% = CUECOINT%(Control$("", "CUECHECA"))
   '
   FIN& = ULTREG&("BADACHE")
   ReDim IDECH$(FIN&)
   UBACHE& = FIN&
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("BADACHE", U&)
      IDEVA$ = TRIM$(Left$(XX$, 40))
      IDECH$(U&) = IDEVA$
   Next U&
   
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI' ORDER BY DESCRIMOV ASC"
   Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With CHETEMP1
      FIN& = .RecordCount
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           U& = .AbsolutePosition
           Call TRACE(20, U&, FIN&)
           .Edit
           !VARIMOVIM = "1"
           .Update
         .MoveNext
         Loop
      End If
   End With
   On Error GoTo 0
   '
   FIN& = ULTREG&("BADACHE")
   ReDim IDECH$(FIN&)
   UBACHE& = FIN&
   Label1.ToolTipText = "Hasta " & FIN&
   For U& = 1 To FIN&
      DIFECART# = 0
      Label1 = U&
      Label1.Visible = True
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("BADACHE", U&)
      IDEVA$ = TRIM$(Left$(XX$, 40))
         IDECH$(U&) = IDEVA$
      Import# = CVD(Mid$(XX$, 49, 8))
      FEMICHE% = CVI(Mid$(XX$, 57, 2))
      FEVECHE% = CVI(Mid$(XX$, 59, 2))
      FRECEP% = CVI(Mid$(XX$, 61, 2))
      DORECEP$ = TRIM$(Mid$(XX$, 63, 10))
      ORICHE$ = TRIM$(Mid$(XX$, 73, 30))
      DESTI$ = TRIM$(Mid$(XX$, 103, 30))
      FENTREGA% = CVI(Mid$(XX$, 133, 2))
      OB1$ = TRIM$(Mid$(XX$, 135, 40))
      OB2$ = TRIM$(Mid$(XX$, 175, 40))
      OB3$ = TRIM$(Mid$(XX$, 215, 40))
      '
      ' BUSCA POR IDEVA$
      If IDEVA$ = "" Then GoTo 19
15    SQLZ$ = "SELECT * FROM CAJABANC "
      SQLZ$ = SQLZ$ + "WHERE TipoMovim = 'CHRECI' "
      SQLZ$ = SQLZ$ + "AND DescriMov = '" & IDEVA$ & "' and VariMovim = '1' "
      Set CHETEMP3 = CueCorri.OpenRecordset(SQLZ$, dbOpenDynaset)
      '
      With CHETEMP3
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
              On Error GoTo 0
              Call REPLA(XX$, "**", 255, 2)
              Call GRAREG("BADACHE", XX$, U&)
              .Edit
              !VARIMOVIM = "2"
              .Update
              GoTo 19    ' LO ENCONTRO
            Else
              On Error GoTo 0
              CAGRE% = CAGRE% + 1
              GoSub AGREGAENTRACHE
              GoTo 15
         End If
      End With
19 Next U&
   '
   SQLY$ = "SELECT * FROM CAJABANC "
   SQLY$ = SQLY$ + "WHERE TipoMovim = 'CHENTR' ORDER BY DESCRIMOV ASC"
   Set chetemp2 = CueCorri.OpenRecordset(SQLY$, dbOpenDynaset)
   '
   With chetemp2
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
         On Error GoTo 0
         FINIX& = .RecordCount
         JKU& = 0
         .MoveFirst
         Do While Not .EOF
           JKU& = JKU& + 1
           Call TRACE(20, JKU&, FINIX&)
           .Edit
           !VARIMOVIM = "1"
           .Update
         .MoveNext
         Loop
      End If
   End With
   On Error GoTo 0
   '
   Label1.ToolTipText = "Hasta " & FIN&
   For U& = 1 To FIN&
      DIFECART# = 0
      Label1 = U&
      Label1.Visible = True
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("BADACHE", U&)
      IDEVA$ = TRIM$(Left$(XX$, 40))
         IDECH$(U&) = IDEVA$
      Import# = CVD(Mid$(XX$, 49, 8))
      FEMICHE% = CVI(Mid$(XX$, 57, 2))
      FEVECHE% = CVI(Mid$(XX$, 59, 2))
      FRECEP% = CVI(Mid$(XX$, 61, 2))
      DORECEP$ = TRIM$(Mid$(XX$, 63, 10))
      ORICHE$ = TRIM$(Mid$(XX$, 73, 30))
      DESTI$ = TRIM$(Mid$(XX$, 103, 30))
      FENTREGA% = CVI(Mid$(XX$, 133, 2))
      OB1$ = TRIM$(Mid$(XX$, 135, 40))
      OB2$ = TRIM$(Mid$(XX$, 175, 40))
      OB3$ = TRIM$(Mid$(XX$, 215, 40))
      '
      If FENTREGA% > 0 Then
        ' BUSCA POR IDEVA$
        If IDEVA$ = "" Then GoTo 39
25      SQLZ$ = "SELECT * FROM CAJABANC "
        SQLZ$ = SQLZ$ + "WHERE TipoMovim = 'CHENTR' "
        SQLZ$ = SQLZ$ + "AND DescriMov = '" & IDEVA$ & "' and VariMovim = '1' "
        Set CHETEMP3 = CueCorri.OpenRecordset(SQLZ$, dbOpenDynaset)
        '
        With CHETEMP3
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
              On Error GoTo 0
              .Edit
              !VARIMOVIM = "2"
              .Update
              GoTo 29    ' LO ENCONTRO
            Else
              On Error GoTo 0
              GoSub AGREGASALIDA
              GoTo 25
          End If
        End With
        '
29      SQLZ$ = "SELECT * FROM CAJABANC "
        SQLZ$ = SQLZ$ + "WHERE TipoMovim = 'CHRECI' "
        SQLZ$ = SQLZ$ + "AND DescriMov = '" & IDEVA$ & "' and VariMovim = '2' "
        Set CHETEMP3 = CueCorri.OpenRecordset(SQLZ$, dbOpenDynaset)
        '
        With CHETEMP3
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
              On Error GoTo 0
              .Edit
              !VARIMOVIM = "3"
              !Status = 2
              !FeDeposEnt = CVDAT(FENTREGA%)
              !FECHENTRE = CVDAT(FENTREGA%)
              If TRIM$(" " & !Document) = "" Then !Document = "Mig.Datos"
              If TRIM$(" " & !DestiCheq) = "" Then !DestiCheq = DESTI$
              If TRIM$(" " & !REFESALCHE) = "" Then !REFESALCHE = OB1$
              If TRIM$(" " & !IDENSALCHE) = "" Then !IDENSALCHE = OB2$ ' AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
              If TRIM$(" " & !UsuaSalChe) = "" Then !UsuaSalChe = OB3$ ' AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
              .Update
          End If
        End With
      End If
39 Next U&
   '
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "AND Status = 1 "
   Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With CHETEMP1
      On Error Resume Next
      .MoveLast
      If Err Then
         On Error GoTo 0
         GoTo 59
      End If
      On Error GoTo 0
      '
      FIN& = .RecordCount
      .MoveFirst
      Do While Not .EOF
        U& = .AbsolutePosition
        Call TRACE(20, U&, FIN&)
        IDEVA$ = !Descrimov
        For JJ& = 1 To UBACHE&
          If IDECH$(JJ&) = IDEVA$ Then
            XX$ = REGLEIDO$("BADACHE", JJ&)
            Import# = CVD(Mid$(XX$, 49, 8))
            FEMICHE% = CVI(Mid$(XX$, 57, 2))
            FEVECHE% = CVI(Mid$(XX$, 59, 2))
            FRECEP% = CVI(Mid$(XX$, 61, 2))
            DORECEP$ = TRIM$(Mid$(XX$, 63, 10))
            ORICHE$ = TRIM$(Mid$(XX$, 73, 30))
            DESTI$ = TRIM$(Mid$(XX$, 103, 30))
            FENTREGA% = CVI(Mid$(XX$, 133, 2))
            OB1$ = TRIM$(Mid$(XX$, 135, 40))
            OB2$ = TRIM$(Mid$(XX$, 175, 40))
            OB3$ = TRIM$(Mid$(XX$, 215, 40))
            '
            If FENTREGA% > 1 Then
                GoTo 45
              Else
                GoTo 49
            End If
          End If
        Next JJ&
        '
        ' no lo encontró
        FENTREGA% = CVINT(!FechConta)     ' SALE EN LA MISMA FECHA DE ENTRADA
        DESTI$ = "Baja por Consistencia"
        OB1$ = "No ingresado en BADACHE.DAT"
        OB2$ = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
        OB3$ = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
        ' fin de no lo encontró
        '
45      .Edit
        !Status = 2
        !VARIMOVIM = "3"
        !FeDeposEnt = CVDAT(FENTREGA%)
        !FECHENTRE = CVDAT(FENTREGA%)
        If TRIM$(" " & !Document) = "" Then !Document = "Mig.Datos"
        If TRIM$(" " & !DestiCheq) = "" Then !DestiCheq = DESTI$
        If TRIM$(" " & !REFESALCHE) = "" Then !REFESALCHE = OB1$
        If TRIM$(" " & !IDENSALCHE) = "" Then !IDENSALCHE = OB2$ ' AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
        If TRIM$(" " & !UsuaSalChe) = "" Then !UsuaSalChe = OB3$ ' AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
        .Update
49    .MoveNext
      Loop
59 End With
Exit Sub
'
AGREGAENTRACHE:
    With CHETEMP1
      .AddNew
      !CodiCom_Cor = DORECEP$
      !CodiCom_Lar = DORECEP$
      !REFECOM = ORICHE$
      !FechConta = CVDAT(FRECEP%)
      !FechAcred = CVDAT(FEVECHE%)
      !CueContaI = CUECARTE%
      !CENCOSTO = ""
      !ItemConc = ""
      !NUCLIEN = 0
      !NuProve = 0
      !TipoMovim = "CHRECI"
      !OPCIMOVIM = "CH"
      !VARIMOVIM = "1"
      !Descrimov = IDEVA$
      !FechEmisi = CVDAT(FEMICHE%)
      !CodiEmpr = 0
      !CodiSucu = 0
      !NUMETALO = 0
      !NumeComp = 0
      !Status = 1
        On Error Resume Next
      !Mone_Emis = 1
      !Tipo_Cam = 1
        On Error GoTo 0
      !ValAbsComp = Abs(Import#)
      !Idebecomp = Import#
      !IHABECOMP = 0
      !Impo_Dola = Import#
      !BASEIMPO = Abs(Import#)
      !NROCAI = ""
        On Error Resume Next
      !IncSubIva = ""
        On Error GoTo 0
      !FEVTOCAI = CVDAT(0)
      !OBSERGEN = ""
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now
      !NuAsiDiar = ""
     .Update
     '
     .AddNew
      !CodiCom_Cor = DORECEP$
      !CodiCom_Lar = DORECEP$
      !REFECOM = ORICHE$
      !FechConta = CVDAT(FRECEP%)
      !FechAcred = CVDAT(FEVECHE%)
      !CueContaI = CUEBALCART%
      !CENCOSTO = ""
      !ItemConc = ""
      !NUCLIEN = 0
      !NuProve = 0
      !TipoMovim = "MIGRA"
      !OPCIMOVIM = ""
      !VARIMOVIM = ""
      !Descrimov = IDEVA$
      !FechEmisi = CVDAT(FEMICHE%)
      !CodiEmpr = 0
      !CodiSucu = 0
      !NUMETALO = 0
      !NumeComp = 0
      !Status = 0
        On Error Resume Next
      !Mone_Emis = 1
      !Tipo_Cam = 1
        On Error GoTo 0
      !ValAbsComp = Import#
      !Idebecomp = 0
      !IHABECOMP = Import#
      !Impo_Dola = (-1) * Import#
      !BASEIMPO = Abs(Import#)
      !NROCAI = ""
        On Error Resume Next
      !IncSubIva = ""
        On Error GoTo 0
      !FEVTOCAI = CVDAT(0)
      !OBSERGEN = ""
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now
      !NuAsiDiar = ""
      .Update
   End With
Return
'
'
AGREGASALIDA:
    With chetemp2
      .AddNew
      !CodiCom_Cor = "MIGRA"
      !CodiCom_Lar = "MIGRA"
      !REFECOM = DESTI$
      !FechConta = CVDAT(FENTREGA%)
      !FechAcred = CVDAT(FENTREGA%)
      !CueContaI = CUECARTE%
      !CENCOSTO = ""
      !ItemConc = ""
      !NUCLIEN = 0
      !NuProve = 0
      !TipoMovim = "CHENTR"
      !OPCIMOVIM = "CH"
      !VARIMOVIM = "1"
      !Descrimov = IDEVA$
      !FechEmisi = CVDAT(FEMICHE%)
      !CodiEmpr = 0
      !CodiSucu = 0
      !NUMETALO = 0
      !NumeComp = 0
      !Status = 0
        On Error Resume Next
      !Mone_Emis = 1
      !Tipo_Cam = 1
        On Error GoTo 0
      !ValAbsComp = Abs(Import#)
      !Idebecomp = 0
      !IHABECOMP = Import#
      !Impo_Dola = (-1) * Import#
      !BASEIMPO = Abs(Import#)
      !NROCAI = ""
        On Error Resume Next
      !IncSubIva = ""
        On Error GoTo 0
      !FEVTOCAI = CVDAT(0)
      !OBSERGEN = ""
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now
      !NuAsiDiar = ""
     .Update
     '
     .AddNew
      !CodiCom_Cor = "MIGRA"
      !CodiCom_Lar = "MIGRA"
      !REFECOM = DESTI$
      !FechConta = CVDAT(FENTREGA%)
      !FechAcred = CVDAT(FENTREGA%)
      !CueContaI = CUEBALCART%
      !CENCOSTO = ""
      !ItemConc = ""
      !NUCLIEN = 0
      !NuProve = 0
      !TipoMovim = "MIGRA"
      !OPCIMOVIM = ""
      !VARIMOVIM = ""
      !Descrimov = IDEVA$
      !FechEmisi = CVDAT(FEMICHE%)
      !CodiEmpr = 0
      !CodiSucu = 0
      !NUMETALO = 0
      !NumeComp = 0
      !Status = 0
        On Error Resume Next
      !Mone_Emis = 1
      !Tipo_Cam = 1
        On Error GoTo 0
      !ValAbsComp = Import#
      !Idebecomp = Import#
      !IHABECOMP = 0
      !Impo_Dola = Import#
      !BASEIMPO = Abs(Import#)
      !NROCAI = ""
        On Error Resume Next
      !IncSubIva = ""
        On Error GoTo 0
      !FEVTOCAI = CVDAT(0)
      !OBSERGEN = ""
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now
      !NuAsiDiar = ""
      .Update
   End With
Return
'
'
End Sub

Private Sub Command7_Click()
   Call MIGRACARCHE
   Screen.MousePointer = 1
   Call COMUNI("Fin de Control y Recuperación")
End Sub

Private Sub Command8_Click()
    Call ABRECABAN
    With CajaBanc
       .MoveFirst
       Do While Not .EOF
       Sudebe# = Sudebe# + !Idebecomp - !IHABECOMP
       .MoveNext
       Loop
    End With
End Sub

Private Sub Command9_Click()
   For I1& = 1 To ULTREG&("CUECORR")
      XX$ = REGLEIDO$("CUECORR", I1&)
      ImpoDeb# = Abs(CVD(Mid$(XX$, 107, 8)))
      If ImpoDeb# > 0 Then
        TTXX$ = TTXX$ + Mid$(XX$, 193, 1)
      End If
   Next I1&
   MsgBox TTXX$

End Sub

Private Sub Exit_Click()
   Call Command24_Click
End Sub

Private Sub Form_Load()
   '


   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture3.Visible = True
      Me.BackColor = &HFCD7B6
      Frame1.BackColor = &HFCD7B6
      Frame4.BackColor = &HFCD7B6
      Frame8.BackColor = &HFCD7B6
    End If

   EJER% = 1
   Call TRANSLABELS(Name)
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   If Left$(APLINVO$, 6) = "AMACLI" Then
       Call Command18_Click  ' Actualización de Base de Datos de Clientes
       Call FINAL("")
     ElseIf Left$(APLINVO$, 6) = "AMASTO" Then
       Call Command1_Click   ' Actualización Maestro de Items de Stock
       Call FINAL("")
     ElseIf Left$(APLINVO$, 6) = "AMAPRO" Then
       Call Command19_Click  ' Actualización de Base de Datos de Proveedores
       Call FINAL("")
     ElseIf Left$(APLINVO$, 8) = "OPMASPRO" Then
       Call ACRE_GES04.Command14_Click  ' Valores por Omisión Base de Datos de Proveedores
       Call FINAL("")
      ElseIf Left$(APLINVO$, 8) = "LIMASTER" Then
       Call OPLIMAS04.LIMASTER(0)
       Call FINAL("")
      ElseIf Left$(APLINVO$, 8) = "TABLES" Then
       OPMAST04.Show 1
       Call FINAL("")
   End If
   '
   If EXISTE%("ISO9K04.EXE") = 0 Then 'Verifica si existe el modulo de calidad
     Command2.Visible = False
   End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub

Private Sub Help_Click()
   Call Command29_Click
End Sub



Private Sub Label11_DblClick()
   If COMALTER%("Realmente Desea Convertir\los Números de Proveedores ?\\No, Continuar\Si, Convertir") = 2 Then
     OPX% = CLACONTRA%
     If OPX% = 1 Then
       Screen.MousePointer = 11
10     Call APERBASDAT("MASTER")            ' AQUI ABRIR BASE
       With Master                ' AQUI PONER TABLA
         On Error Resume Next
         .MoveFirst
         If Err Then
           On Error GoTo 0
           GoTo 20
         End If
         On Error GoTo 0
         Do While Not .EOF
           PRORI% = !PROVHABI     ' AQUI AJUSTAR VARIABLE
           PRNUE% = REPNU%(PRORI%)
           If PRNUE% <> PRORI% Then
             .Edit
             !PROVHABI = PRNUE%
             .Update
           End If
         .MoveNext
         Loop
       End With
       '
20     Call APERBASDAT("COMPRAS")            ' AQUI ABRIR BASE
       With OcomDeta               ' AQUI PONER TABLA
         On Error Resume Next
         .MoveFirst
         If Err Then
           On Error GoTo 0
           GoTo 30
         End If
         On Error GoTo 0
         Do While Not .EOF
           PRORI% = !Npro_Ocom     ' AQUI AJUSTAR VARIABLE
           PRNUE% = REPNU%(PRORI%)
           If PRNUE% <> PRORI% Then
             .Edit
             !Npro_Ocom = PRNUE%   ' AQUI AJUSTAR VARIABLE
             .Update
           End If
         .MoveNext
         Loop
       End With
       '
30     Call APERBASDAT("COMPRAS")            ' AQUI ABRIR BASE
       With OcomEnca               ' AQUI PONER TABLA
         On Error Resume Next
         .MoveFirst
         If Err Then
           On Error GoTo 0
           GoTo 40                 ' aqui al proximo bloque
         End If
         On Error GoTo 0
         Do While Not .EOF
           PRORI% = !Npro_Ocom     ' AQUI AJUSTAR VARIABLE
           PRNUE% = REPNU%(PRORI%)
           If PRNUE% <> PRORI% Then
             .Edit
             !Npro_Ocom = PRNUE%   ' AQUI AJUSTAR VARIABLE
             .Update
           End If
         .MoveNext
         Loop
       End With
       '
40     Call APERBASDAT("STOCKS")            ' AQUI ABRIR BASE
       With Bolrecep              ' AQUI PONER TABLA
         On Error Resume Next
         .MoveFirst
         If Err Then
           On Error GoTo 0
           GoTo 50                 ' aqui al proximo bloque
         End If
         On Error GoTo 0
         Do While Not .EOF
           PRORI% = !NumProv       ' AQUI AJUSTAR VARIABLE
           PRNUE% = REPNU%(PRORI%)
           If PRNUE% <> PRORI% Then
             .Edit
             !NumProv = PRNUE%     ' AQUI AJUSTAR VARIABLE
             .Update
           End If
         .MoveNext
         Loop
       End With
       '
50     Call APERBASDAT("STOCKS")            ' AQUI ABRIR BASE
       With Movisto               ' AQUI PONER TABLA
         On Error Resume Next
         .MoveFirst
         If Err Then
           On Error GoTo 0
           GoTo 60                 ' aqui al proximo bloque
         End If
         On Error GoTo 0
         Do While Not .EOF
           PRORI% = !Nume_Prov     ' AQUI AJUSTAR VARIABLE
           PRNUE% = REPNU%(PRORI%)
           If PRNUE% <> PRORI% Then
             .Edit
             !Nume_Prov = PRNUE%   ' AQUI AJUSTAR VARIABLE
             .Update
           End If
         .MoveNext
         Loop
       End With
       '
       Screen.MousePointer = 1
       '
     End If
   End If
   '
60 Call COMUNI("Conversion Completa")
   '
End Sub

Function REPNU%(NORIG%)
   '
   RPN% = NORIG%
   '
   If NORIG% = 1 Then RPN% = 10: GoTo 90
   If NORIG% = 2 Then RPN% = 34: GoTo 90
   If NORIG% = 3 Then RPN% = 61: GoTo 90
   If NORIG% = 4 Then RPN% = 64: GoTo 90
   If NORIG% = 5 Then RPN% = 66: GoTo 90
   If NORIG% = 6 Then RPN% = 68: GoTo 90
   If NORIG% = 7 Then RPN% = 96: GoTo 90
   If NORIG% = 8 Then RPN% = 98: GoTo 90
   If NORIG% = 9 Then RPN% = 295: GoTo 90
   If NORIG% = 10 Then RPN% = 304: GoTo 90
   If NORIG% = 11 Then RPN% = 308: GoTo 90
   If NORIG% = 12 Then RPN% = 333: GoTo 90
   If NORIG% = 13 Then RPN% = 425: GoTo 90
   If NORIG% = 14 Then RPN% = 665: GoTo 90
   If NORIG% = 15 Then RPN% = 858: GoTo 90
   If NORIG% = 16 Then RPN% = 1014: GoTo 90
   If NORIG% = 17 Then RPN% = 985: GoTo 90
   If NORIG% = 18 Then RPN% = 138: GoTo 90
   If NORIG% = 19 Then RPN% = 314: GoTo 90
   If NORIG% = 20 Then RPN% = 107: GoTo 90
   If NORIG% = 21 Then RPN% = 128: GoTo 90
   If NORIG% = 22 Then RPN% = 393: GoTo 90
   If NORIG% = 23 Then RPN% = 980: GoTo 90
   If NORIG% = 24 Then RPN% = 848: GoTo 90
   If NORIG% = 25 Then RPN% = 170: GoTo 90
   If NORIG% = 26 Then RPN% = 310: GoTo 90
   If NORIG% = 27 Then RPN% = 766: GoTo 90
   If NORIG% = 28 Then RPN% = 48: GoTo 90
   If NORIG% = 29 Then RPN% = 842: GoTo 90
   If NORIG% = 30 Then RPN% = 82: GoTo 90
   If NORIG% = 31 Then RPN% = 434: GoTo 90
   If NORIG% = 32 Then RPN% = 300: GoTo 90
   If NORIG% = 33 Then RPN% = 860: GoTo 90
   If NORIG% = 34 Then RPN% = 519: GoTo 90
   '
90 REPNU% = RPN%
   '
End Function
'
Private Sub Mastok_Click()
   Call Command1_Click
End Sub

Private Sub mnuimportar_Click()
   Call Command21_Click
End Sub

Private Sub MNUSSCAL_Click()
    Call Command2_Click
End Sub

Private Sub Proveed_Click()
   Call Command19_Click
End Sub

'\\\OT-05-0656-WAR-03/10/05///
Sub CONVBADA() 'CONVERSIÓN DE .DAT A .MDB
   '
   If COMALTER%("La Conversión de Datos\Debe Correrse en el Servidor.\\Continuar\Cancelar") <> 1 Then
      Exit Sub
   End If
   '
   If COMALTER%("Antes de la Conversión de Datos\Conviene en el Sistema DAT\Pedir un Listado de Cuentas\por Cobrar y por Pagar\para Verificar la Composición de Deuda.\\Continuar\Cancelar") <> 1 Then
      Exit Sub
   End If
   '
   Label1.Visible = True
   If COMALTER%("Plan de Cuentas Contables\\Migrar\Saltear") = 1 Then Call MIGRAPLANCUEN
   '
   For IXX% = 0 To 6
      TXXX$ = TRIM$(Str$(IXX%))
      CUEM$ = Control$("", "CUEIVA-" + TXXX$)
      If CUEM$ = "" Or CUEINT%(CUEM$) < 1 Then
         Call MENSERR(24, "Falta Definir Cuentas Contables\Correspondientes a IVA Crédito Fiscal Tasa " + TTXX$ + ".")
         Call FINAL("")
      End If
   Next IXX%
   '
   FINI& = ULTREG&("PROVINB")
   If FINI& < 1 Then
      Call MENSERR(24, "Falta Tabla de Provincias")
      Call FINAL("")
   End If
   For I& = 1 To FINI&
      X$ = REGLEIDO$("PROVINB", I&)
      CUEII% = CVI(Mid$(X$, 27, 2))
      If CUEII% < 1 Or ECOARCH$("CUECON", MKI$(CUEII%)) = "" Then
         Call MENSERR(24, "Faltan Cuentas Contables de Percepcion Ingresos Brutos")
         Call FINAL("")
      End If
   Next I&
   If CUEIPROVIN%("B") < 1 Then
      Call MENSERR(24, "Falta Cuenta Contable I.Brutos Provincia 'B'")
      Call FINAL("")
   End If
   '
   CUEBAL04.Show 1
   Call FRESHALL
   CUEBALRECI% = CUEINT%(CUEBAL04.Label1)
   CUEBALPAGO% = CUEINT%(CUEBAL04.label2)
   CUEBALAJUS% = CUEINT%(CUEBAL04.Label3)
   CUEBALCART% = CUEINT%(CUEBAL04.Label4)
   If CUEBALRECI% < 1 Or CUEBALPAGO% < 1 Or CUEBALAJUS% < 1 Or CUEBALCART% < 1 Then
      Call FINAL("")
   End If
   '
   If COMALTER%("Clientes\\Migrar\Saltear") = 1 Then Call MIGRACLIEN
   If COMALTER%("Proveedores\\Migrar\Saltear") = 1 Then Call MIGRAPROVE
   If COMALTER%("Maestro de Items de Stock\\Migrar\Saltear") = 1 Then Call MIGRAMASTER
   If COMALTER%("Formulas de Producto\\Migrar\Saltear") = 1 Then Call MIGRAFORMULA
   If COMALTER%("Secuencia de Operaciones\\Migrar\Saltear") = 1 Then Call MIGRASECOPER
   If COMALTER%("Pedidos de Clientes\\Migrar\Saltear") = 1 Then Call MIGRAPEDIDOS
   If COMALTER%("Ordenes de Compras\\Migrar\Saltear") = 1 Then Call MIGRACOMPRAS
   If COMALTER%("Movimiendos de Stocks\\Migrar\Saltear") = 1 Then Call MIGRASTOCKS
   If COMALTER%("Boletas de Recepción\\Migrar\Saltear") = 1 Then Call MIGRABOLREC
   If COMALTER%("Material en Consignacion\\Migrar\Saltear") = 1 Then Call MIGRACONSIGNA
   If COMALTER%("Ordenes de Fabricacion\\Migrar\Saltear") = 1 Then Call MIGRAORFAB
   If COMALTER%("Reservas para O/F\\Migrar\Saltear") = 1 Then Call MIGRARESERVA
   '
   If COMALTER%("Presupuestos Emitidos\\Migrar\Saltear") = 1 Then Call MIGRAPRESU
   '
   If COMALTER%("Lista de Precios de Venta\\Migrar\Saltear") = 1 Then Call MIGRAPRECIOS
   '
   If COMALTER%("Cuecorri + Cajabanc\\Migrar\Saltear") <> 1 Then GoTo 50
   '
   Call BORRACUECORR
   '
   Call CONVERFAC
   Call CONVERCOB
   Call CONVERDEU
   Call CONVAPLICOB
   '
   Call GRAFACRE
   Call MIGRADECHE
   '
   PORCABA% = 0
   If ULTREG&("CAJABANC") > 0 Then
      PORCABA% = 1
   End If
   '
   If PORCABA = 0 Then
         Call MIGRACARCHE
      Else
         Call MARCACARCHE
   End If
   '
50 If COMALTER%("Contabilidad General\\Migrar\Saltear") = 1 Then
      Call MIGRACONTA
      Call MIGRACONTB
   End If
   '
'90 MsgBox "Falta Migrar Contabilidad General"
   '
95 Screen.MousePointer = 1
  Label1.Visible = False
  Call COMUNI("Ha Finalizado con Éxito la Conversión.")
  '
End Sub

Sub MARCACARCHE()
   '
   RFF$ = RECFILT$("CAJABANC", 16, Chr$(1)) + RECFILT$("CAJABANC", 16, Chr$(11))
   JJ& = 0
   FIN& = Len(RFF$)
   '
   Call APERBASDAT("CABAN")
   
   For I& = 1 To FIN& Step 4
     Call TRACE(20, I&, FIN&)
     RECHE& = CVS(Mid$(RFF$, I&, 4))
     X$ = REGLEIDO$("CAJABANC", RECHE&)
     '
     If UCase$(Mid$(X$, 69, 2)) <> "CH" Then  ' no es un cheque
       GoTo 19
     End If
     '
     If CVI(Mid$(X$, 269, 2)) > 0 Then ' esta anulado
       Call REPLA(X$, Chr$(9), 121, 1)
       Call GRAREG("CAJABANC", X$, RECHE&)
       GoTo 19
     End If
     '
     If CVI(Mid$(X$, 275, 2)) > 0 Then ' salio de cartera
       Call REPLA(X$, Chr$(2), 121, 1)
       Call GRAREG("CAJABANC", X$, RECHE&)
       GoTo 19
     End If
     '
     IMCHE# = CVD(Mid$(X$, 153, 8)) - CVD(Mid$(Z$, 161, 8))
     IDEVA$ = Mid$(X$, 71, 40)
     IMCHEX$ = TRIM$(FORMATNUM$(Abs(IMCHE#), "F14.2"))
       If IMCHEX$ = "" Then IMCHEX$ = "0"
     
     SQLX$ = "SELECT * FROM CAJABANC "
     SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
     SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
     SQLX$ = SQLX$ + "AND DescriMov = '" & IDEVA$ & "' "
     SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IMCHEX$ & ") < 0.005 "
     Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
     '
     On Error Resume Next
     CHETEMP1.MoveFirst
     If Err Then
          On Error GoTo 0
          MsgBox "No encontrado"
        Else
          On Error GoTo 0
          CHETEMP1.Edit
          CHETEMP1!Status = 11
          CHETEMP1.Update
     End If
19 Next I&
   '
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
   SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
   SQLX$ = SQLX$ + "AND STATUS = 1"
   Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   On Error Resume Next
   With CHETEMP1
   .MoveFirst
   If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
          IDEVA$ = !Descrimov
          Import# = !Idebecomp
          ABCX1$ = IDEVA$
          GoSub BUSCALASALIDA
          If ENCONTROLASALIDA% > 0 Then
             GoTo 29
          End If
          '
          FENTREGA% = CVINT%(!FechConta) ' Si no en BADACHE, sale misma fecha que entró
          FEMICHE% = CVINT%(!FechEmisi)
          DOCUEN$ = "Mig.Datos"
          DESTI$ = "Cuenta Balance Cartera"
          OB1$ = "Conversión Datos DAT-MDB "
          .Edit
          !Status = 2
          !FeDeposEnt = CVDAT(FENTREGA%)
          !FECHENTRE = CVDAT(FENTREGA%)
          !Document = DOCUEN$
          !DestiCheq = DESTI$
          !REFESALCHE = OB1$
          !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
          !UsuaSalChe = OB3$ ' AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
          .Update
          '
          GoSub AGREGASALIDA
          GoTo 29
          '
29     .MoveNext
       Loop
   End If
   End With
   '
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
   SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
   SQLX$ = SQLX$ + "AND STATUS = 11"
   Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   On Error Resume Next
   With CHETEMP1
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
         .Edit
         !Status = 1
         .Update
         .MoveNext
         Loop
      End If
   End With
   On Error GoTo 0
Exit Sub
   '
BUSCALASALIDA:
    ENCONTROSALIDA% = 0
    With CajaBanc
     .FindFirst "DescriMov = '" & IDEVA$ & "' " ' and VariMovim = '1' "
37   If .NoMatch = False Then
        If Abs(Import# - IHABECOMP) < 0.005 Then
          .Edit
          !Status = 2
          .Update
          ENCONTROLASALIDA% = 1
          GoTo 85
        End If
        .FindNext "DescriMov = '" & IDEVA$ & "' "
        GoTo 37
     End If
   End With
85 Return
   '
   ' **********************************************************************
AGREGASALIDA:
    '
    With CajaBanc
      .AddNew
      !CodiCom_Cor = "MIGRA"
      !CodiCom_Lar = "MIGRA"
      !REFECOM = DESTI$
      !FechConta = CVDAT(FENTREGA%)
      !FechAcred = CVDAT(FENTREGA%)
      !CueContaI = CUECARTE%
      !CENCOSTO = ""
      !ItemConc = ""
      !NUCLIEN = 0
      !NuProve = 0
      !TipoMovim = "CHENTR"
      !OPCIMOVIM = "CH"
      !VARIMOVIM = "1"
      !Descrimov = IDEVA$
      !FechEmisi = CVDAT(FEMICHE%)
      !CodiEmpr = 0
      !CodiSucu = 0
      !NUMETALO = 0
      !NumeComp = 0
      !Status = 1
        On Error Resume Next
      !Mone_Emis = 1
      !Tipo_Cam = 1
        On Error GoTo 0
      !ValAbsComp = Abs(Import#)
      !Idebecomp = 0
      !IHABECOMP = Import#
      !Impo_Dola = (-1) * Import#
      !BASEIMPO = Abs(Import#)
      !NROCAI = ""
        On Error Resume Next
      !IncSubIva = ""
        On Error GoTo 0
      !FEVTOCAI = CVDAT(0)
      !OBSERGEN = ""
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now
      !NuAsiDiar = ""
     .Update
     '
     .AddNew
      !CodiCom_Cor = "MIGRA"
      !CodiCom_Lar = "MIGRA"
      !REFECOM = DESTI$
      !FechConta = CVDAT(FENTREGA%)
      !FechAcred = CVDAT(FENTREGA%)
      !CueContaI = CUEBALCART%
      !CENCOSTO = ""
      !ItemConc = ""
      !NUCLIEN = 0
      !NuProve = 0
      !TipoMovim = "MIGRA"
      !OPCIMOVIM = ""
      !VARIMOVIM = ""
      !Descrimov = IDEVA$
      !FechEmisi = CVDAT(FEMICHE%)
      !CodiEmpr = 0
      !CodiSucu = 0
      !NUMETALO = 0
      !NumeComp = 0
      !Status = 0
        On Error Resume Next
      !Mone_Emis = 1
      !Tipo_Cam = 1
        On Error GoTo 0
      !ValAbsComp = Import#
      !Idebecomp = Import#
      !IHABECOMP = 0
      !Impo_Dola = Import#
      !BASEIMPO = Abs(Import#)
      !NROCAI = ""
        On Error Resume Next
      !IncSubIva = ""
        On Error GoTo 0
      !FEVTOCAI = CVDAT(0)
      !OBSERGEN = ""
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now
      !NuAsiDiar = ""
      .Update
   End With
Return

'
'
End Sub

Sub CONVERDEU()
    '
   Picture1.ToolTipText = "Conversión de Saldos Deudores"
   '
65 Call APERBASDAT("CABAN")
   With SaDeuPen
      On Error Resume Next
140    .MoveFirst
       If Err = 0 Then
          .Delete
          GoTo 140
       End If
      On Error GoTo 0
   End With
   '
   FIN& = ULTREG&("SVD202")
   Label1.ToolTipText = "Hasta " & FIN&
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Label1 = U&
      Label1.Refresh
      X$ = REGLEIDO$("SVD202", U&)
      Call DATDEUPEN(X$, TIDOCU$, NFCX$, NFCY$, OPMOVIX$, VAX$, NC%, FFI%, FVENCI%, SIGX%, IORIG#, IPENDE#)

      ' AQUI GRABAR CAJABANC PARA FACTURA
      NOCLOSE% = 1
30    CoComCor$ = NFCX$
      CoComLar$ = NFCY$
      CodiEmp% = 0
      CodiSuc% = 0
      NumeClie% = GRUDIVCLI%(NC%)
      NumeProv% = 0
      FechEmi = CVDAT(FFI%)
      MONEMI% = 1
      IBruTot# = IORIG#
      IIvaTot# = 0
      FeVenc = CVDAT(FVENCI%)
      ImVenc = IPENDE#
      SIGNO% = 1
      '
      NuorVen% = 1
      Call GRADEUPEN
      '
   Next U&
   '
End Sub

Sub CONVAPLICOB()
   '
   Call APERBASDAT("CABAN")
   With ApliCobra
   On Error Resume Next
13   .MoveFirst
     If Err = 0 Then
       .Delete
       GoTo 13
     End If
     On Error GoTo 0
   End With
   '
   FIN& = ULTREG&("APLICOB")
   Label1.ToolTipText = "Hasta " & FIN&
   Label1.Visible = True
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     Label1 = U&
     Label1.Refresh
     XX$ = REGLEIDO$("APLICOB", U&)
     NumeClie% = CVI(Mid$(XX$, 3, 2))
     '
     SQLX$ = "SELECT CodiCom_Lar FROM CAJABANC "
     SQLX$ = SQLX$ + "where NuClien = " & NumeClie%
     Set MOFASEL = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
'mofasel.MoveFirst
'Do While Not mofasel.EOF
'If InStr(mofasel!cODIcOM_Lar, "21143") > 0 Then
'AAA = bbb
'End If
'mofasel.MoveNext
'Loop
'MOFASEL.MoveLast
'AA1 = MOFASEL.RecordCount
'5 MOFASEL.MoveNext
'AA2 = MOFASEL!CodiCom_Lar + "   " + CoComLar$
'GoTo 5
     CoComLar$ = Mid$(XX$, 5, 18)
        NUMEFA$ = Left$(CoComLar$, 3) + "???????" + Mid$(CoComLar$, 11)
        NUMEFB$ = Left$(CoComLar$, 3) + "?????" + Mid$(CoComLar$, 9, 8) + "??"
        MOFASEL.FindFirst "CodiCom_Lar LIKE '" & NUMEFA$ & "' OR CodiCom_Lar LIKE '" & NUMEFB$ & "' "
        If MOFASEL.NoMatch = False Then
            CoComLar = MOFASEL.CodiCom_Lar
            GoTo 15
        End If
'MsgBox "No Encontré " + CoComLar$
15      NROX$ = CoComLar$
     CoComCor$ = Left$(NROX$, 3) + Mid$(NROX$, 9, 1) + "-" + Mid$(NROX$, 14)
     ImpApli# = CVD(Mid$(XX$, 23, 8))
     RefApli$ = Mid$(XX$, 31, 24)
     DocOrLar$ = Mid$(XX$, 55, 18)
        NUMEFA$ = Left$(DocOrLar$, 3) + "???????" + Mid$(DocOrLar$, 11)
        NUMEFB$ = Left$(DocOrLar$, 3) + "?????" + Mid$(DocOrLar$, 9, 8) + "??"
        MOFASEL.FindFirst "CodiCom_Lar LIKE '" & NUMEFA$ & "' OR CodiCom_Lar LIKE '" & NUMEFB$ & "' "
        If MOFASEL.NoMatch = False Then
            DocOrLar = MOFASEL.CodiCom_Lar
            GoTo 25
        End If
'MsgBox "No Encontré " + DocOrLar$
25      NROX$ = DocOrLar$
     DocOrCor$ = Left$(NROX$, 3) + Mid$(NROX$, 9, 1) + "-" + Mid$(NROX$, 14)
     TotOrig# = CVD(Mid$(XX$, 73, 8))
     IvaOrig# = CVD(Mid$(XX$, 81, 8))
     Call GRAPLICOB
   Next U&
   '
End Sub
'
Sub CONVERCOB()
   '
   Picture1.ToolTipText = "Conversion de Cobranzas"
   Call APERBASDAT("CABAN")
   NOCLOSE% = 1
   '
   CUECARTE% = CUECOINT%(Control$("", "CUECHECA"))
   FIN& = ULTREG&("COBRAN")
   DIFECOB# = 0
   '
   PORCABA% = 0
   If ULTREG&("CAJABANC") > 0 Then
      PORCABA% = 1
   End If
   '
   Label1.ToolTipText = "Hasta " & FIN&
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Label1 = U&
      Label1.Refresh
      X$ = REGLEIDO$("COBRAN", U&)
      NC% = CVI(Left$(X$, 2))
      NROO& = CVS(Mid$(X$, 3, 4))
      FC% = CVI(Mid$(X$, 7, 2))
      TOTAVAL# = CVD(Mid$(X$, 15, 8))
      If TOTAVAL# < 0.005 Then GoTo 89 ' ESTÁ ANULADO
      COBRACT% = CVI(Mid$(X$, 23, 2))
      '
      RECICOR$ = TRIM$(Str$(COBRACT%))
      While Len(RECICOR$) < 2: RECICOR$ = "0" + RECICOR$: Wend
      RECICOR$ = TRIM$(Str$(NROO&)) + "-" + RECICOR$
      While Len(RECICOR$) < 9: RECICOR$ = "0" + RECICOR$: Wend
      RECICOR$ = "RB-" + RECICOR$
      NRECIB$ = RECICOR$
      '
      RECILAR$ = TRIM$(Str$(NROO&))
      While Len(RECILAR$) < 8: RECILAR$ = "0" + RECILAR$: Wend
      RECILAR$ = TRIM$(Str$(COBRACT%)) + "-" + RECILAR$
      While Len(RECILAR$) < 13: RECILAR = "0" + RECILAR$: Wend
      RECILAR$ = "RB-X-" + RECILAR$
      '
      CoComCor$ = NRECIB$
      CoComLar$ = RECILAR$
      Referen$ = RASOCLI$(NC%)
      SUVALOR# = 0
      FeContab = CVDAT(FC%)
      FeAcredi = CVDAT(FC%)
      '
      '
      If PORCABA% = 1 Then
         RFF$ = RECFILT$("CAJABANC", 1, NRECIB$)
         For KKU& = 1 To Len(RFF$) Step 4
            REG& = CVS(Mid$(RFF$, KKU&, 4))
            Z$ = REGLEIDO$("CAJABANC", REG&)
            '
            IMCHE# = CVD(Mid$(Z$, 153, 8))
            If IMCHE# > 0 Then
              FC1% = CVI(Mid$(Z$, 61, 2))
              If FC1% = FC% Then
                 SUVALOR# = SUVALOR# + IMCHE#
                 CodiEmp% = Asc(Mid$(Z$, 113, 1))
                 CodiSuc% = Asc(Mid$(Z$, 114, 1))
                 NumeClie% = NC%
                 NumeProv% = 0
                 '
                 FECHE% = CVI(Mid$(Z$, 111, 2))
                 FEVEN% = CVI(Mid$(Z$, 63, 2))
                 CUITLIB$ = ""                 ' NO HAY DATOS PARA MIGRAR
                 COBRACT% = CVI(Mid$(Z$, 115, 2))
                 SUCUI% = 0                    ' NO HAY DATOS PARA MIGRAR
                 '
                   FEX = FEVEN%: FEVECHE$ = FECHATEX$(FEX)
                 FeContab = CVDAT(FC%)
                   CUECO% = CVI(Mid$(Z$, 65, 2))
                 CuentaCo% = CUECO%
                 TiMovi$ = "CHRECI"
                 OpMovi$ = "CH"
                 VaMovi$ = ""
                   IDEVA$ = Mid$(Z$, 71, 40)
                 Descrip$ = IDEVA$
                   If FECHE% < 1 Then FECHE% = FC%
                 FechEmi = CVDAT(FECHE%)
                 FeAcredi = CVDAT(FEVEN%)
                 NumeTal = COBRACT%
                 NumeCom = CVD(Mid$(Z$, 117, 4))
                    STATUI% = Asc(Mid$(Z$, 121, 1))
                    If CUECO% = CUECARTE% Then STATUI% = 1
                 StatInit = STATUI%
                 ImpoDeb# = IMCHE#
                 Observa$ = ""
                   If CUITLIB$ <> "" Then
                     Observa$ = "CUIT Libr: " + CUITLIB$
                   End If
                 Call GRACABAN
                 '
              End If
            End If
         Next KKU&
         GoTo 850
      End If
      '
      '

840   ' ***** aqui grabar cheques en CAJABANC.DAT
      '
      RFF$ = RECFILT$("CHECAR", 11, MKS$(NROO&))
      SUVALOR# = 0
      For KKU& = 1 To Len(RFF$) Step 4
        REG& = CVS(Mid$(RFF$, KKU&, 4))
        Z$ = REGLEIDO$("CHECAR", REG&)
        '
        FC1% = CVI(Mid$(Z$, 45, 2))
        '
        If FC1% = FC% Then
           TIPO% = CVI(Left$(Z$, 2))
           BANCO$ = TRIM$(Mid$(Z$, 17, 16))
           If BANCO$ = "" Then
             BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(TIPO%)))
           End If
           If Left$(UCase$(BANCO$), 3) = "CH." Then
             If Len(BANCO$) > 3 Then
               BANCO$ = Mid$(BANCO$, 4)
             End If
           End If
           NUCHE$ = TRIM$(Mid$(Z$, 5, 12))
           IMCHE# = CVD(Mid$(Z$, 33, 8))
             SUVALOR# = SUVALOR# + IMCHE#
           FECHE% = CVI(Mid$(Z$, 43, 2))
           FEVEN% = CVI(Mid$(Z$, 41, 2))
           CUITLIB$ = ""                 ' NO HAY DATOS PARA MIGRAR
           COBRACT% = 0
           SUCUI% = 0                    ' NO HAY DATOS PARA MIGRAR
           '
           If Abs(IMCHE#) > 0 Then
             '
               FEX = FEVEN%: FEVECHE$ = FECHATEX$(FEX)
             FeContab = CVDAT(FC%)
               CUECO% = CUEMECOB%(TIPO%)
             CuentaCo% = CUECO%
             NumeClie% = NC%
             NumeProv% = 0
             TiMovi$ = "CHRECI"
             OpMovi$ = "CH"
             VaMovi$ = ""
               IDEVA$ = AJUSTI$(BANCO$, 18)
               IDEVA$ = IDEVA$ + AJUSTI$(NUCHE$, 14)
               IDEVA$ = IDEVA$ + FEVECHE$
             Descrip$ = IDEVA$
               If FECHE% < 1 Then FECHE% = FC%
             FechEmi = CVDAT(FECHE%)
               If FEVEN% < FECHE% Then FEVEN% = FECHE%
               If FEVEN% < FC% Then FEVEN% = FC%
             FeAcredi = CVDAT(FEVEN%)
             NumeTal = COBRACT%
             CodiEmp% = 0
             CodiSuc% = SUCUI%
             NumeCom = NROO&
                STATUI% = 0
                If CUECO% = CUECARTE% Then STATUI% = 1
             StatInit = STATUI%
             ImpoDeb# = IMCHE#
             Observa$ = ""
               If CUITLIB$ <> "" Then
                 Observa$ = "CUIT Libr: " + CUITLIB$
               End If
             Call GRACABAN
             '
           End If
        End If
      Next KKU&
      '
850   If Abs(TOTAVAL#) >= 0.005 Then
         FeContab = CVDAT(FC%)
           CUECO% = CUECOCLI%(NC%)
         CuentaCo% = CUECO%
         NumeClie% = NC%
         NumeProv% = 0
         TiMovi$ = "RCOB"
         OpMovi$ = "CCORR"
         VaMovi$ = ""
         Descrip$ = "Cobranza Cta.Corriente " + NRECIB$
         FechEmi = CVDAT(FC%)
         FeAcredi = CVDAT(FC%)
         NumeTal = COBRACT%
         CodiEmp% = 0
         CodiSuc% = SUCUI%
         NumeCom = NROO&
         StatInit = 0
         ImpoDeb# = (-1) * TOTAVAL#   ' SE FUERZA QUE CUADRE
         Call GRACABAN
      End If
      '
      If Abs(SUVALOR# - TOTAVAL#) > 0.005 Then
        DIFECOB# = ROUND(TOTAVAL# - SUVALOR#)
        Referen$ = "Dif.Importación " + NRECIB$
        CuentaCo% = CUEBALRECI%
        TiMovi$ = "AJUSTE"
        OpMovi$ = ""
        VaMovi$ = ""
        Descrip$ = "Dif.Importación " + NRECIB$
        ImpoDeb# = DIFECOB#
        Observa$ = ""
        Call GRACABAN
      End If
      '
89 Next U&
   '
End Sub

'\\\OT-05-0656-WAR-FIN///
Sub CONVERFAC()
   '
   Picture1.ToolTipText = "Conversión de Facturas"
   FIN& = ULTREG&("FACTUR")
   Label1.ToolTipText = "Hasta " & FIN&
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Label1 = U&
      Label1.Refresh
      X$ = REGLEIDO$("FACTUR", U&)
      Call DATFACTUR(X$, TIDOCU$, NFCX$, NFCY$, OPMOVIX$, VAX$, NC%, FFI%, SIGX%, TOTAGEN#, NEFA7#, IVAD1#, IVAD2#, RNI#, IPER)
      '
      ' AQUI GRABAR CAJABANC PARA FACTURA
      Call APERBASDAT("CABAN")
      NOCLOSE% = 1
30    CoComCor$ = NFCX$
      CoComLar$ = NFCY$
      CodiEmp% = 0
      CodiSuc% = 0
      NumeClie% = GRUDIVCLI%(NC%)
      NumeProv% = 0
      FechEmi = CVDAT(FFI%)
      Referen$ = RASOCLI$(NC%)
      FeContab = CVDAT(FFI%)
      FeAcredi = CVDAT(FFI%)
      CuentaCo% = CUECOCLI%(NC%)
      TiMovi$ = "FVEN"
      OpMovi$ = OPMOVIX$
      VaMovi$ = VAX$
      Descrip$ = TIDOCU$ + " " + VAX$ + " " + Mid$(NFCY$, 4, 13)
      NumeTal = XVALO(Mid$(NFCY$, 4, 4))
      NumeCom = XVALO(Mid$(NFCY$, 9, 8))
      StatInit = 0
      ImpoDeb = TOTAGEN# * SIGX%
      Observa$ = RASOCLI$(NC%)
      Call GRACABAN
      '
      'graba venta neta - Verificar que QQQ sea la IC de combinacion Cliente-Tipo de Comprobante
      TOTANETX# = NEFA7#
      If Abs(TOTANETX#) >= 0.005 Then
        PPVI% = PIVACLI%(NC%)
        If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
        T16$ = Control$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
        QQ1% = CUECOINT%(T16$)
        'If QQ1% < 1 Then QQ1% = QQQ%
        CuentaCo% = QQ1%   ' TOMA LA QUE APARECE EN PANTALLA
        TiMovi$ = "VENTA"
        OpMovi$ = OPMOVIX$
        VaMovi$ = VAX
        Descrip$ = "Mercaderia Vendida " + NFCX$
        StatInit = 0
        ImpoDeb = TOTANETX# * SIGX% * (-1)
        Call GRACABAN
      End If
      '
      'graba iva inscripto
      TOTANETX# = IVAD1#
      If Abs(TOTANETX#) >= 0.005 Then
          If CUIVA0% < 1 Then
             CUEIVA$ = Control$("SVD451S", "CUEIVA")
             CUIVA0% = CUECOINT%(CUEIVA$)
          End If
          CuentaCo% = CUIVA0%
          TiMovi$ = "IVA"
          OpMovi$ = "DeFi"
          VaMovi$ = "21.0 %"
          Descrip$ = "Iva Debito Fiscal " + NFCX$
          StatInit = 0
          ImpoDeb = TOTANETX# * SIGX% * (-1)
          BasImpo# = ROUND#(TOTANETX# / 0.21, 2)
          Call GRACABAN
      End If
      BasImpo# = 0
      '
      'graba iva inscripto
      TOTANETX# = IVAD2#
      If Abs(TOTANETX#) >= 0.005 Then
          If CUIVA1% < 1 Then
             CUEIVA$ = Control$("SVD451S", "CUEIVA")
             CUIVA1% = CUECOINT%(CUEIVA$)
          End If
          CuentaCo% = CUIVA1%
          TiMovi$ = "IVA"
          OpMovi$ = "DeFi"
          VaMovi$ = "10.5 %"
          Descrip$ = "Iva Debito Fiscal " + NFCX$
          StatInit = 0
          ImpoDeb = TOTANETX# * SIGX% * (-1)
          BasImpo# = ROUND#(TOTANETX# / 0.105, 2)
          Call GRACABAN
      End If
      BasImpo# = 0
      '
      'graba iva NO inscripto
      TOTANETX# = RNI#
      If Abs(TOTANETX#) >= 0.005 Then
        If CUADU% < 1 Then
           CUEIVA$ = Control$("SVD451S", "CUEADU")
           CUADU% = CUECOINT%(CUEIVA$)
        End If
        CuentaCo% = CUADU%
        TiMovi$ = "IVA"
        OpMovi$ = "DeFi"
        VaMovi$ = "RNI"
        Descrip$ = "IVA Resp.No Inscripto " + NFCX$
        StatInit = 0
        ImpoDeb = TOTANETX# * SIGX% * (-1)
        BasImpo# = ROUND#(NEFA7#, 2)
        Call GRACABAN
      End If
      BasImpo# = 0
      '
      'graba percepcion I.Brutos
      TOTANETX# = IPER
      If Abs(TOTANETX#) >= 0.005 Then
        If CUIVA% < 1 Then
           CUEIVA$ = Control$("PERIBRU", "CUECONTA")
           CUIVA% = CUECOINT%(CUEIVA$)
        End If
        CuentaCo% = CUIVA%
        TiMovi$ = "IBRU"
        OpMovi$ = "IB-B"
        VaMovi$ = "PERC"
        Descrip$ = "Perc.IB.Bs.As. " + NFCX$
        StatInit = 0
        ImpoDeb = TOTANETX# * SIGX% * (-1)
        BasImpo# = ROUND#(NEFA7#, 2)
        Call GRACABAN
      End If
      BasImpo# = 0
      '
   Next U&
   '
End Sub

Sub MIGRADECHE()
   '
   Picture1.ToolTipText = "Conversion de obranzas"
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
   SQLX$ = SQLX$ + "AND Status = 1 "
   SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
   Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
              '
   NOCLOSE% = 0
   Call APERBASDAT("CABAN")
   NOCLOSE% = 1
   '
   FIN& = ULTREG&("CAJABANC")
   For REG& = 1 To FIN&
      Call TRACE(20, REG&, FIN&)
      Z$ = REGLEIDO$("CAJABANC", REG&)
      '
      CoComCor$ = Left$(Z$, 12): NRECIB$ = Left$(Z$, 12)
      If Left$(CoComCor$, 2) <> "DB" Then GoTo 29
      '
      CoComLar$ = Left$(Z$, 12)
      CodiEmp% = Asc(Mid$(Z$, 113, 1))
      CodiSuc% = Asc(Mid$(Z$, 114, 1))
      NumeClie% = CVI(Mid$(Z$, 67, 2))
      NumeProv% = 0
      FechEmi = CVDAT(CVI(Mid$(Z$, 111, 2)))
      Referen$ = Mid$(Z$, 13, 40): DESTIN$ = TRIM$(Mid$(Z$, 13, 48))

      IMCHE# = CVD(Mid$(Z$, 153, 8)) - CVD(Mid$(Z$, 161, 8))
      FC% = CVI(Mid$(Z$, 61, 2))
      FECHE% = CVI(Mid$(Z$, 111, 2))
      FEVEN% = CVI(Mid$(Z$, 63, 2))
      CUITLIB$ = ""                 ' NO HAY DATOS PARA MIGRAR
      COBRACT% = CVI(Mid$(Z$, 115, 2))
      SUCUI% = Asc(Mid$(Z$, 114, 1))
      '
      If Abs(IMCHE#) > 0 Then
      '
             FEX = FEVEN%: FEVECHE$ = FECHATEX$(FEX)
          FeContab = CVDAT(FC%)
             CUECO% = CVI(Mid$(Z$, 65, 2))
          CuentaCo% = CUECO%
             NumeClie% = CVI(Mid$(Z$, 67, 2))
             NumeProv% = 0
             TiMovi$ = "DEPOBAN"
             If IMCHE# < 0 Then TiMovi$ = "CHENTR"
             OpMovi$ = Mid$(Z$, 69, 2)
             VaMovi$ = ""
                IDEVA$ = Mid$(Z$, 71, 40)
             Descrip$ = IDEVA$
             If FECHE% < 1 Then FECHE% = FC%
             FechEmi = CVDAT(FECHE%)
             FeAcredi = CVDAT(FEVEN%)
             NumeTal = COBRACT%
             CodiEmp% = 0
             CodiSuc% = SUCUI%
             NumeCom = CVD(Mid$(Z$, 117, 4))
                 STATUI% = Asc(Mid$(Z$, 121, 1))
             StatInit = STATUI%
             ImpoDeb# = IMCHE#
             Observa$ = ""
               If CUITLIB$ <> "" Then
                 Observa$ = "CUIT Libr: " + CUITLIB$
               End If
             Call GRACABAN
             '
             If IMCHE# < 0 Then
                 'BUSCAR SI ESTÁ INGRESADO
                 ' CASO AFIRMATIVO:
                 ' 1) MARCAR LA SALIDA DE CARTERA
                 ' 2) PONER CUENTA = CUENTA DE CARTERA DE VALORES
                 VUELTA% = 0: ENCUEN% = 0
                 IMCHEX$ = TRIM$(FORMATNUM$(Abs(IMCHE#), "F14.2"))
                    If IMCHEX$ = "" Then IMCHEX$ = "0"
                 SQLX$ = "DescriMov = '" & IDEVA$ & "' "
                 SQLX1$ = SQLX$
                 SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IMCHEX$ & ") < 0.005 "
                 'SQLX$ = SQLX$ + "AND Status = 1"
                 '
                 With CHETEMP1
                    .FindFirst SQLX$
                    If .NoMatch = False Then
                       On Error GoTo 0
                       .Edit
                       !FeDeposEnt = CVDAT(FC%)
                       !Document = NRECIB$
                       !DestiCheq = Left$(DESTIN$, 48)
                         CUECONTABINT% = !CueContaI  ' PARA QUE LA SALIDA SEA DE LA MISMA CUENTA QUE ENTRO
                       !REFESALCHE = "Deposito Bancario " + NRECIB$
                       !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
                       !UsuaSalChe = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
                       !Status = 2                     ' Marca Salida de Cartera
                       !IdenBaja = IDACTU$             ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
                       .Update
                       ENCUEN% = 1
                    End If
                 End With
                 '
             End If
      End If
      '
29  Next REG&
'      '
End Sub


Sub DATFACTUR(X$, TIDOCU$, NFCX$, NFCY$, OPMOVIX$, VAMOVIX$, NC%, FFI%, SIGX%, TOTAGEN#, NETO#, IVAD1#, IVAD2#, RNI#, IPER)
'**********
    '
    Static CTXX%, MODOSABO%, CUEPER$, CUEPIBRU%
    '
    If CTXX% = 0 Then
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       MODOSABO% = 0: CUEPIBRU% = 0
       If InStr(EPAC$, "SABO") > 0 Or InStr(EPAC$, "DOSIVA") > 0 Then
         MODOSABO% = 1
         CUEPER$ = Control$("PERIBRU", "CUECONTA")
         CUEPIBRU% = CUECOINT%(CUEPER$)
       End If
    End If
    '
    NC% = CVI(Left$(X$, 2))
    FFI% = CVI(Mid$(X$, 7, 2))
       If FFI% = 0 Then FFI% = FFANT%
    FFANT% = FFI%
    '
    NRO& = CVS(Mid$(X$, 3, 4))
    NUSUCUVE% = CVI(Mid$(X$, 47, 2))
    NSUCX$ = TRIM$(Str$(NUSUCUVE%)): While Len(NSUCX$) < 4: NSUCX$ = "0" + NSUCX$: Wend
    '
    TIPO% = Asc(Mid$(X$, 13, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
    TICOMX$ = "FC": OPMOVIX$ = "FAC": TIDOCU$ = "Factura"
    If TIPO% = 3 Then TICOMX$ = "ND": OPMOVIX$ = "DEB": TIDOCU$ = "N.Debito"
    If TIPO% = 6 Then TICOMX$ = "NC": OPMOVIX$ = "CRE": TIDOCU$ = "N.Credito"
    If TIPO% = 0 Or TIPO% = 7 Then TICOMX$ = "AJ": OPMIVIX$ = "AJUD": TIDOCU$ = "Ajuste"
    '
    VA% = Asc(Mid$(X$, 14, 1)): If VA% < 1 Or VA% > 3 Then VA% = 1 ' default factura
    VAX$ = "A"
    If VA% = 2 Then VAX$ = "B"
    If VA% = 3 Then VAX$ = "E"
    '
    NFCZ$ = TRIM$(Str$(NRO&)): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
    NFCX$ = TICOMX$ + "-" + Mid$(NFCZ$, 3) + "-" + VAX$
    NFCY$ = TICOMX$ + "-" + NSUCX$ + "-" + NFCZ$ + "-" + VAX$
    '
    OPMOVIX$ = "FC"
    If TIPO% = 3 Then OPMOVIX$ = "ND"
    If TIPO% = 6 Then OPMOVIX$ = "NC"
    VAMOVIX$ = VAX$
    '
    NETO1# = (Int(100 * (CVD(Mid$(X$, 15, 8))) + 0.5)) / 100
    IVAD1# = (Int(100 * CVD(Mid$(X$, 23, 8)) + 0.5)) / 100
    RNI# = (Int(100 * CVD(Mid$(X$, 31, 8)) + 0.5)) / 100
    NETO2# = (Int(100 * CVD(Mid$(X$, 49, 8)) + 0.5)) / 100
    IVAD2# = (Int(100 * CVD(Mid$(X$, 57, 8)) + 0.5)) / 100
    IPER = 0
    '
    If MODOSABO% = 1 Then
       IPER = IVAD2#
       IVAD2# = 0
    End If
    '
    If Abs(NETO2#) > 0.1 Then
      If Abs(100 * IVAD2# / NETO2# - TAIVA(0)) < 0.2 Then
        NETO1# = NETO1# + NETO2#
        Call REPLA(X$, MKD$(NETO1#), 15, 8)
        NETO2# = 0
        Call REPLA(X$, MKD$(NETO2#), 49, 8)
        IVAD1# = IVAD1# + IVAD2#
        Call REPLA(X$, MKD$(IVAD1#), 23, 8)
        IVAD2# = 0
        Call REPLA(X$, MKD$(IVAD2#), 57, 8)
      End If
    End If
    '
    NETO# = NETO1# + NETO2#
    IVAD# = IVAD1# + IVAD2#
    TOTAGEN# = NETO# + IVAD# + RNI# + CDbl(IPER) + IRET#
    '
    SIGX% = 1
    If TIPO% > 3 Or TOTAGEN# < 0 Then
        NETO1# = Abs(NETO1#)
        NETO2# = Abs(NETO2#)
        NETO# = Abs(NETO#)
        IVAD1# = Abs(IVAD1#)
        IVAD2# = Abs(IVAD2#)
        IVAD# = Abs(IVAD#)
        RNI# = Abs(RNI#)
        IPER = Abs(IPER)
        TOTAGEN# = Abs(TOTAGEN#)
        SIGX% = (-1)
    End If
    '
End Sub

Sub DATDEUPEN(X$, TIDOCU$, NFCX$, NFCY$, OPMOVIX$, VAMOVIX$, NC%, FFI%, FVENCI%, SIGX%, IORIG#, IPENDE#)

'**********
    '
    NC% = CVI(Left$(X$, 2))
    FFI% = CVI(Mid$(X$, 19, 2))
       If FFI% = 0 Then FFI% = FFANT%
    FFANT% = FFI%
    '
    FVENCI% = CVI(Mid$(X$, 9, 2))
    If FVENCI% < FFI% Then FVENCI% = FFI%
    '
    NRO& = CVS(Mid$(X$, 5, 4))
    NUSUCUVE% = CVI(Mid$(X$, 29, 2))
    NSUCX$ = TRIM$(Str$(NUSUCUVE%)): While Len(NSUCX$) < 4: NSUCX$ = "0" + NSUCX$: Wend
    '
    TIPO% = Asc(Mid$(X$, 3, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
    TICOMX$ = "FC": OPMOVIX$ = "FAC": TIDOCU$ = "Factura"
    If TIPO% = 3 Then TICOMX$ = "ND": OPMOVIX$ = "DEB": TIDOCU$ = "N.Debito"
    If TIPO% = 4 Then TICOMX$ = "RB": OPMOVIX$ = "RCOB": TIDOCU$ = "R.Cobranza"
    If TIPO% = 6 Then TICOMX$ = "NC": OPMOVIX$ = "CRE": TIDOCU$ = "N.Credito"
    If TIPO% = 0 Or TIPO% = 7 Then TICOMX$ = "AJ": OPMIVIX$ = "AJUD": TIDOCU$ = "Ajuste"
    '
    VA% = Asc(Mid$(X$, 4, 1)): If VA% < 1 Or VA% > 3 Then VA% = 1 ' default factura
    VAX$ = "A"
    If VA% = 2 Then VAX$ = "B"
    If VA% = 3 Then VAX$ = "E"
    If TIPO% = 4 Then
      VA% = 0
      VAX$ = "X"
    End If
    '
    NFCZ$ = TRIM$(Str$(NRO&)): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
    NFCX$ = TICOMX$ + "-" + Mid$(NFCZ$, 3) + "-" + VAX$
    NFCY$ = TICOMX$ + "-" + NSUCX$ + "-" + NFCZ$ + "-" + VAX$
    '
    OPMOVIX$ = "FC"
    If TIPO% = 3 Then OPMOVIX$ = "ND"
    If TIPO% = 4 Then OPMOVIX$ = "RB"
    If TIPO% = 6 Then OPMOVIX$ = "NC"
    VAMOVIX$ = VAX$
    '
    IPENDE# = (Int(100 * (CVD(Mid$(X$, 11, 8))) + 0.5)) / 100
    IORIG# = (Int(100 * CVD(Mid$(X$, 21, 8)) + 0.5)) / 100
    '
    SIGX% = 1
    If TIPO% > 3 Or IORIG# < 0 Then
       SIGX% = (-1)
    End If
    '
    IPENDE# = SIGX% * Abs(IPENDE#)
    IORIG# = SIGX% * Abs(IORIG#)
    '
End Sub

Sub GRAFACRE(Optional TIMOFIL$)
   '
   Static YAINDEXCHE%
   '
   Dim FEVEN%(10)
   Dim VENCIPESOS#(10)
   '
   Dim CAMP%, CUEMP%(2048), MEPA%(2048)
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
   '
   PORCABA% = 0
   If ULTREG&("CAJABANC") > 0 Then
      PORCABA% = 1
   End If
   '
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
   SQLX$ = SQLX$ + "AND Status = 1 "
   SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
   Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
              '
   NOCLOSE% = 0
   Call APERBASDAT("CABAN")
   NOCLOSE% = 1
   FIN& = ULTREG&("CUECORR")
   Label1.ToolTipText = "Hasta " & FIN&
   For U& = 1 To FIN&
      Label1 = U&
      Call TRACE(20, U&, FIN&)
      Label1 = U&
      Label1.Refresh
      XX$ = REGLEIDO$("CUECORR", U&)
      FF% = CVI(Left$(XX$, 2))
      NC% = CVI(Mid$(XX$, 5, 2))
      NROX$ = Mid$(XX$, 7, 18)
      NRECIB$ = Left$(NROX$, 3) + Mid$(NROX$, 13, 6) + "-" + Mid$(NROX$, 8, 2)
      TIDOCOM$ = Left$(NROX$, 2)
      '
      ' ESTO ES PARA RECUPERAR AJUSTES DE SALDOS
      If TIMOFIL$ <> "" Then
         If TIMOFIL$ = "AS" And TIDOCOM$ = "FF" Then GoTo 12
         If TIDOCOM$ <> TIMOFIL$ Then
            GoTo 15
         End If
      End If
      '
12    SUC% = 0
      NRO& = XVALO(Mid$(NROX$, 11, 8))
      Observa$ = Mid$(XX$, 25, 24)
      ICONTAX$ = Mid$(XX$, 199, 80)
      '
      If TIDOCOM$ = "OP" Or TIDOCOM$ = "FF" Then
          GoSub 30
        Else
          If TIDOCOM$ = "AS" Or TIDOCOM$ = "AJ" Then
             AAA = BBB
          End If
          GoSub 20
      End If
      '
15 Next U&
   GoSub GRAPLIP
Exit Sub
   '
   '
20 NC1% = (-1) * NC%
   '
   CoComCor$ = NRECIB$
   CoComLar$ = NROX$
   CodiEmp% = 0
   CodiSuc% = SUC%
   NumeClie% = 0
   NumeProv% = NC%
   FechEmi = CVDAT(FF%)
   '
   Referen$ = RASOCLI$(NC1%)
   FeContab = CVDAT(FF%)
   FeAcredi = CVDAT(FF%)
   NumeTal = 0
   NumeCom = NRO&
      If Observa$ = "" Then Observa$ = RASOCLI$(NC1%)
   NumCai$ = ""
   FechCai = CVDAT(FECHANUM("01/01/80"))
   InclIva$ = ""
   TOFA0# = 0
   '
   ' graba detalle por imputación de compras
   For KKI% = 1 To Len(ICONTAX$) Step 10
     CuentaCo% = CVI(Mid$(ICONTAX$, KKI%, 2))
     CentroCos$ = ""
     ItConcep$ = "" 'P/ conceptos
     ImpoDeb# = Abs(CVD(Mid$(ICONTAX$, KKI% + 2, 8)))
     If TIDOCOM$ = "AS" Then
        If KKI% = 1 Then
            ImpoDeb# = CVD(Mid$(XX$, 123, 8)) - CVD(Mid$(XX$, 115, 8))
          Else
            ImpoDeb# = 0
        End If
        If CuentaCo% < 1 Then CuentaCo% = CUEBALAJUS%
     End If
     If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
     '
     If CuentaCo% > 0 Then
       If Abs(ImpoDeb#) >= 0.005 Then
         DEMOVX$ = DENOCUE$(CuentaCo%)
         TiMovi$ = "COMPRA"
         OpMovi$ = TIDOCOM$
         VaMovi$ = Mid$(NROX$, 4, 1)
         Descrip$ = TRIM$(DEMOVX$) + " - " + RASOCLI$(NC1%)
         TOFA0# = TOFA0# + ImpoDeb#
         Call GRACABAN
       End If
     End If
     CentroCos$ = ""
   Next KKI%
   '
   ' grabar importes de IVA
   SUMIVA# = 0
   IEXEN# = ROUND(CVD(Mid$(XX$, 59, 8)))
   If XVALO(IEXEN#) >= 0.005 Then   ' GRABAR IMPORTE EXENTO
     TiMovi$ = "IVA"
     OpMovi$ = "CrFi"
     VaMovi$ = "Exento"
     ImpoDeb# = 0
     BasImpo# = ROUND(IEXEN#)
     Descrip$ = "IVA B.Imp. Exento " + NRECIB$
     CuentaCo% = CUECOINT%(Control$("", "CUEIVA-0"))
     TOFA0# = TOFA0# + ImpoDeb#
     Call GRACABAN
   End If
   '
   INOGRA# = ROUND(CVD(Mid$(XX$, 67, 8)))
   If XVALO(INOGRA#) >= 0.005 Then   ' GRABAR IMPORTE NO GRAVADO
     TiMovi$ = "IVA"
     OpMovi$ = "CrFi"
     VaMovi$ = "No Grav"
     ImpoDeb# = 0
     BasImpo# = INOGRA#
     Descrip$ = "IVA B.Imp. No Grav." + NRECIB$
     CuentaCo% = CUECOINT%(Control$("", "CUEIVA-0"))
     TOFA0# = TOFA0# + ImpoDeb#
     Call GRACABAN
   End If
   '
'  HASTA AQUI LLEGUE *************************************
   For KKI% = 75 To 114 Step 8
     TiMovi$ = "IVA"
     OpMovi$ = "CrFi"
     VaMovi$ = ""
     Descrip$ = "IVA Credito Fiscal " + NRECIB$
     ImpoDeb# = Abs(CVD(Mid$(XX$, KKI%, 8)))
     If TIDOCOM$ = "AS" Then
        ImpoDeb# = ImpoDeb# * Sgn(CVD(Mid$(XX$, 123, 8)) - CVD(Mid$(XX$, 115, 8)))
     End If
     If Abs(ImpoDeb#) < 0.005 Then GoTo 29
     '
     BasImpo# = 0
     If KKI% = 75 Then
         VaMovi$ = "21.00 %"
         CuentaCo% = CUECOINT%(Control$("", "CUEIVA-0"))
       ElseIf KKI% = 83 Then
         VaMovi$ = "Aduana"
         CuentaCo% = CUECOINT%(Control$("", "CUEIVA-4"))
       ElseIf KKI% = 91 Then
         VaMovi$ = "Per-IVA"
         Descrip$ = "Percepcion I.V.A. " + NRECIB$
         CuentaCo% = CUECOINT%(Control$("", "CUEIVA-5"))
       ElseIf KKI% = 99 Then
         TiMovi$ = "GAN"
         VaMovi$ = "Per-Gan"
         Descrip$ = "Percepcion Ganancias " + NRECIB$
         CuentaCo% = CUECOINT%(Control$("", "CUEIVA-6"))
       ElseIf KKI% = 107 Then
         TiMovi$ = "I-BRU"
         VaMovi$ = "PERC"
         Descrip$ = "Percepcion I.Brutos " + NRECIB$
' ojo que esto está mal
         CuentaCo% = CUEIPROVIN%("B")
     End If
     '
     If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
     If Abs(ImpoDeb#) >= 0.005 Then
       TOFA0# = TOFA0# + ImpoDeb#
       Call GRACABAN
     End If
     BasImpo# = 0
29 Next KKI%
   '
   ' graba factura en c.corriente proveedor
   CuentaCo% = CUECOINT%(CUEMADRE$(-1 * NC%))
   ImpoDeb# = (-1) * TOFA0#
     'If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
   TiMovi = "FCOM"
   OpMovi = Left$(Text11, 2)
   VaMovi = Label20
   Descrip$ = CoComLar$ + " - " + RASOCLI$(NC1%)
   FeAcredi = CVDAT(FF%)
   Call GRACABAN
   '
   ' Grabar Vencimientos de Facturas ******************************************
   SIGNO% = 1
   If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then SIGNO% = (-1)
     'IBruTot# = Abs(TOFAC#(0))
   IBruTot# = TOFA0#
   IIvaTot# = SUMIVA#
   '
   PRICONVEN% = -1
   SALBRUTO# = TOFA0# * SIGNO%
   TXVENCI$ = Mid$(XX$, 139, 60)
   For KKI% = 1 To 6
      FEVEN%(KKI%) = CVI(Mid$(TXVENCI$, 10 * KKI% - 9, 2))
      VENCIPESOS#(KKI%) = CVD(Mid$(TXVENCI$, 10 * KKI% - 7, 8))
   Next KKI%
   '
   For KKI% = 1 To 6
      If FEVEN(KKI%) <> 0 Then
        If Abs(VENCIPESOS#(KKI%)) >= 0.005 Then
           NuorVen% = KKI%
           FeVenc = CVDAT(FEVEN(KKI%))
           ImVenc# = VENCIPESOS#(KKI%)
           Call GRACREPEN
         End If
      End If
   Next KKI%
   '
Return
'
30 ' GRABA ORDENES DE PAGO
    IDEBEX# = CVD(Mid$(XX$, 115, 8))
    IHABEX# = CVD(Mid$(XX$, 123, 8))
    DIFEPAG# = 0
    'If Abs(IDEBEX#) + Abs(IHABEX#) < 0.005 Then
    '   GoTo 399    ' ESTÁ ANULADA - PERO SE GRABA IGUAL
    'End If
    '
    CoComCor$ = NRECIB$
    CoComLar$ = NROX$
    IMPOP# = ROUND(IDEBEX# - IHABEX#)
    CodiEmp% = 0
    CodiSuc% = SUC%
    NumeClie% = 0
    NumeProv% = NC%
    FechEmi = CVDAT(FF%)
    Referen$ = RASOCLI$((-1) * NC%)
    FeContab = CVDAT(FF%)
    FeAcredi = CVDAT(FF%)
    CuentaCo% = CUECOCLI%(NC1%)
    TiMovi$ = "OPAG"
    OpMovi$ = "OP"
    VaMovi$ = "CCORR"
    Descrip$ = "Orden de Pago " + NRECIB$
    '
    If Left$(NRECIB$, 1) = "A" Then ' PARA AJUSTES DE SALDO
       TiMovi = "AJUP"              ' AJUSTE DE SALDO PROVEEDORES
       OpMovi = "AS"                ' AJUSTE DE SALDO
       VaMovi = "CCORR"             ' CUENTA CORRIENTE
       Descrip = "Aj.Saldo - " + RASOCLI$((-1) * NC%)
    End If
    '
    NumeTal = COBRACT%
    NumeCom = NRO&
    StatInit = 0
    ImpoDeb# = IMPOP#
    BasImpo# = IMPOP#    ' aqui va entonces el neto sin IVA
    Observa$ = ""
    If Observa$ = "" Then Observa$ = RASOCLI$(NC1%)
    '
    'DIFEPAG# = DIFEPAG# + ImpoDeb#
    If TICUEPRO%((-1) * NC%) = 0 Then
         DIFEPAG# = DIFEPAG# + ImpoDeb#
         Call GRACABAN
       Else
         RFG$ = RECFILT$("DETPAGAS", 3, MKS$(NRO&))
         If RFG$ = "" Then
            DIFEPAG# = DIFEPAG# + ImpoDeb#
            Call GRACABAN
            GoTo 32
         End If
         For K2& = 1 To Len(RFG$) Step 4
           K1& = CVS(Mid$(RFG$, K2&, 4))
           Z$ = REGLEIDO$("DETPAGAS", K1&)
           If CVI(Mid$(Z$, 7, 2)) = FF% Then
             CuentaCo% = CVI(Mid$(Z$, 17, 2))
             ImpoDeb# = CVD(Mid$(Z$, 9, 8))
             VaMovi$ = "CGAST"
             Referen$ = DENOCUE$(CuentaCo%)
             CentroCos$ = ""
             ItConcep$ = "" 'P/ conceptos
             DIFEPAG# = DIFEPAG# + ImpoDeb#
             Call GRACABAN
           End If
         Next K2&
    End If
    '
    'grabar cheques entregados
32  If YAINDEXCHE% = 0 Then
        XX1$ = REGLEIDO$("CHEQUEMI", 1)
        Call GRAREG("CHEQUEMI", XX1$, 1)
        Call CIERRARCH("CHEQUEMI")
        YAINDEXCHE% = 1
    End If
    '
    RFGH$ = RECFILT$("CHEQUEMI", 6, MKS$(NRO&))
    If PORCABA% = 1 Then
       RFGH$ = RECFILT$("CAJABANC", 1, NRECIB$)
    End If
    For KKUU& = 1 To Len(RFGH$) Step 4
       RECHE& = CVS(Mid$(RFGH$, KKUU&, 4))
       If PORCABA% = 0 Then
            Z$ = REGLEIDO$("CHEQUEMI", RECHE&)
         Else
            Z$ = REGLEIDO$("CAJABANC", RECHE&)
            IMCHE# = CVD(Mid$(Z$, 161, 8))
            FFX1% = CVI(Mid$(Z$, 61, 2))
       End If
       If (PORCABA% = 0 And CVI(Mid$(Z$, 1, 2)) = NC%) Or (PORCABA% = 1 And IMCHE# > 0 And FFX1% = FF%) Then ' COINCIDE PROVEEDOR
          BANCO$ = AJUSTI$(Mid$(Z$, 15, 16), 18)
          NUCHE$ = AJUSTI$(Mid$(Z$, 3, 12), 14)
            FEX = CVI(Mid$(Z$, 49, 2))
            FEVEN1% = FEX
          FEVECH$ = FECHATEX$(FEX)
          IDEVA$ = BANCO$ + NUCHE$ + FEVECH$
          IMCHE# = CVD(Mid$(Z$, 33, 8))
          FECHE% = CVI(Mid$(Z$, 31, 2))
          If PORCABA% = 1 Then
             FECHE% = CVI(Mid$(Z$, 111, 2))
             FEVEN1% = CVI(Mid$(Z$, 63, 2))
             IDEVA$ = Mid$(Z$, 71, 40)
             IMCHE# = CVD(Mid$(Z$, 161, 8))
          End If
          '
          'BUSCAR SI ESTÁ INGRESADO
          ' CASO AFIRMATIVO:
          ' 1) MARCAR LA SALIDA DE CARTERA
          ' 2) PONER CUENTA = CUENTA DE CARTERA DE VALORES
          VUELTA% = 0: ENCUEN% = 0
          IMCHEX$ = TRIM$(FORMATNUM$(IMCHE#, "F14.2"))
             If IMCHEX$ = "" Then IMCHEX$ = "0"
          SQLX$ = "DescriMov = '" & IDEVA$ & "' "
          SQLX1$ = SQLX$
          SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IMCHEX$ & ") < 0.005 "
          'SQLX$ = SQLX$ + "AND Status = 1"
          '
          SQLZ$ = "SELECT * FROM CAJABANC "
          SQLZ$ = SQLZ$ + "WHERE TipoMovim = 'CHRECI'"
          SQLZ$ = SQLZ$ + "AND Status = 1 "
          SQLZ$ = SQLZ$ + "AND OPCIMOVIM = 'CH' "
          SQLZ$ = SQLZ$ + "AND DescriMov = '" & IDEVA$ & "' "
          SQLZ$ = SQLZ$ + "AND ABS(ValAbsComp - " & IMCHEX$ & ") < 0.005 "
          Set CHETEMP11 = CueCorri.OpenRecordset(SQLZ$, dbOpenDynaset)
          
          With CHETEMP11
             On Error Resume Next
             .MoveFirst
             If Err < 1 Then
                On Error GoTo 0
                .Edit
                !FeDeposEnt = CVDAT(FF%)
                !Document = NRECIB$
                !DestiCheq = RASOCLI$((-1) * NC%)
                  CUECONTABINT% = !CueContaI  ' PARA QUE LA SALIDA SEA DE LA MISMA CUENTA QUE ENTRO
                !REFESALCHE = "Orden de Pago Nro." + Str$(NRO&)
                !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
                !UsuaSalChe = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
                !Status = 2                     ' Marca Salida de Cartera
                !IdenBaja = IDACTU$             ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
                .Update
                ENCUEN% = 1
             End If
             On Error GoTo 0
          End With
          '
          ' ESTO SI NO ESTÁ INGRESADO
          If ENCUEN% < 1 Then
             QUE% = CVI(Mid$(Z$, 45, 2))
             For UX% = 1 To CAMP%
                 If MEPA%(UX%) = QUE% Then GoTo 38
             Next UX%
             ' PONER AQUI UNA OPCIÓN POR DEFAULT
             UX% = 3   ' ESTO ES EFECTIVO EN FERVI
             '
38           CUECONTABINT% = CUEMP%(UX%)
          End If
          '
          NC1% = (-1) * NC%
          IMPOP# = IMPOP# - IMCHE#
          CoComCor$ = NRECIB$
          CoComLar$ = NROX$
          CodiEmp% = 0
          CodiSuc% = SUC%
          NumeClie% = 0
          NumeProv% = NC%
          FechEmi = CVDAT(FECHE%)
            FEX = FF%
          Referen$ = RASOCLI$(NC1%) + " - " + FECHATEX$(FEX)
          FeContab = CVDAT(FF%)
          '
          FeAcredi = CVDAT(FEVEN1%)
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
          Observa$ = "O.Pago " + TRIM$(Str$(NRO&)) + " - " + RASOCLI$(NC1%)
          '
          DIFEPAG# = DIFEPAG# + ImpoDeb#
          Call GRACABAN
           '
       End If
    Next KKUU&
    '
    If Abs(DIFEPAG#) > 0.005 Then
       Referen$ = "Dif.Importación " + NRECIB$
       CuentaCo% = CUEBALPAGO%
       TiMovi$ = "AJUSTE"
       OpMovi$ = ""
       VaMovi$ = ""
       Descrip$ = "Dif.Importación " + NRECIB$
       ImpoDeb# = (-1) * DIFEPAG#
       Observa$ = ""
       Call GRACABAN
    End If
    '
   ' Grabar Vencimientos de Facturas ******************************************
39 SIGNO% = (-1)
   IBruTot# = Abs(ROUND(CVD(Mid$(XX$, 59, 8))))
   IIvaTot# = 0
   '
   PRICONVEN% = -1
   SALBRUTO# = Abs(IBruTot#) * SIGNO%
   TXVENCI$ = Mid$(XX$, 139, 60)
   For KKI% = 1 To 6
      FEVEN%(KKI%) = CVI(Mid$(TXVENCI$, 10 * KKI% - 9, 2))
      VENCIPESOS#(KKI%) = CVD(Mid$(TXVENCI$, 10 * KKI% - 7, 8))
   Next KKI%
   '
   For KKI% = 1 To 6
      If FEVEN(KKI%) <> 0 Then
        If Abs(VENCIPESOS#(KKI%)) >= 0.005 Then
           NuorVen% = KKI%
           FeVenc = CVDAT(FEVEN(KKI%))
           ImVenc# = VENCIPESOS#(KKI%)
           Call GRACREPEN
         End If
      End If
   Next KKI%
   '
399 Return

GRAPLIP:
  With ApliPago
    On Error Resume Next
1301   .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 1301
    End If
    On Error GoTo 0
  End With
  '
  FIN& = ULTREG&("APLIPAG")
  For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("APLIPAG", U&)
     NumeProv% = CVI(Mid$(XX$, 3, 2))
     CoComLar$ = Mid$(XX$, 5, 18)
        NROX$ = CoComLar$
     CoComCor$ = Left$(NROX$, 3) + Mid$(NROX$, 9, 1) + "-" + Mid$(NROX$, 14)
     ImpApli# = CVD(Mid$(XX$, 23, 8))
     RefApli$ = Mid$(XX$, 31, 24)
     DocOrLar$ = Mid$(XX$, 55, 18)
        NROX$ = DocOrLar$
     DocOrCor$ = Left$(NROX$, 3) + Mid$(NROX$, 9, 1) + "-" + Mid$(NROX$, 14)
     TotOrig# = CVD(Mid$(XX$, 73, 8))
     IvaOrig# = CVD(Mid$(XX$, 81, 8))
     Call GRAPLIPAG
   Next U&
Return

End Sub

Sub MIGRACLIEN()
  Picture1.ToolTipText = "Depuración - Clientes"
  '
  Call APERBASDAT("CLIEN")
  With Deudor12
    On Error Resume Next
10  .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 10
    End If
    On Error GoTo 0
  End With
  '
  Picture1.ToolTipText = "Conversión de Clientes"
  '
  'LECTURA
  UR& = ULTREG&("DEUDOR1")
  For U& = 1 To ULTREG&("DEUDOR1")
    '
    X1$ = REGLEIDO$("DEUDOR1", U&)
      NCI% = CVI(Mid$(X1$, 1, 2))
    X2$ = REGLEIDO$("DEUDOR2", U&)
    X3$ = FILTERGETRECORD$("BADACLIE", 1, MKI$(NCI%))
    'LECTURA
    RASO$ = Mid$(X3$, 3, 62)
    FANTA$ = Mid$(X3$, 65, 32)
    DOMI$ = Mid$(X3$, 97, 64)
    CPOSTX$ = Mid$(X3$, 161, 16)
    CPOST% = 0: If XVALO(CPOSTX$) < 32767 Then CPOST% = XVALO(CPOSTX$)
    LOCA$ = TRIM$(Mid$(X3$, 177, 48))
    PAISOK$ = Mid$(X3$, 225, 32)
    TELE$ = Mid$(X3$, 257, 24)
    FAXX$ = TRIM$(Mid$(X3$, 281, 24))
    CONTACTX$ = TRIM$(Mid$(X3, 385, 48))
    MAILPRO$ = TRIM$(Mid$(X3$, 305, 48))
    HORARIX$ = TRIM$(Mid$(X3$, 449, 32))
    ZONCLI$ = Mid$(X3$, 481, 12)
    CUIT$ = TRIM$(Mid$(X2$, 75, 15))
      If CUIT$ <> "" Then VCC% = VALICUIT%(CUIT$)
    GruCli% = CVI(Mid$(X2$, 125, 2))
    '
    NCITXT$ = CStr(NCI%)
    Do Until Len(NCITXT$) = 5
      NCITXT$ = "0" + NCITXT$
    Loop
    CODICUEN$ = "CL-" & NCITXT$
    '
    TELEF2$ = Mid$(X2$, 65, 10)
    PIVA% = Asc(Mid$(X2$, 104, 1))
    CONDI% = Asc(Mid$(X2$, 118, 1))
    PROVI$ = TRIM$(Mid$(X2$, 117, 1))
    LIPRE% = Asc(Mid$(X2$, 119, 1))
    COVEN% = Asc(Mid$(X2$, 120, 1))
    LICRE = CVS(Mid$(X2$, 121, 4))
    INGBRU$ = Mid$(X2$, 90, 13)
    CAIBRU$ = Mid$(X2$, 103, 1)
    CUEMA% = CVI(Mid$(X2$, 127, 2))
    GRUDIV% = 0
    '
    DOMIENT1$ = ""
    LOCAENT1$ = ""
    TRANSPOR$ = ""
    NTRANX% = 0
    HORARIENT$ = ""
    If NCI% > 0 Then
      XX$ = FILTERGETRECORD$("LENTRECL", 1, MKI$(NCI%))
      DOMIENT1$ = TRIM$(Mid$(XX$, 3, 32))
      LOCAENT1$ = TRIM$(Mid$(XX$, 35, 32))
      TRANSPOR$ = TRIM$(Mid$(XX$, 67, 32))
      NTRANX% = CVI(Mid$(XX$, 99, 2))
         If NTRANX% = 8224 Then NTRANX% = 0
      HORARIENT$ = TRIM$(Mid$(XX$, 101, 24))
    End If
    '
    DESSTAN$ = ""
    XXX$ = FILTERGETRECORD$("DESCUCLI", 1, MKI$(NCI%))
    DESSTAN$ = Mid$(XXX$, 3, 14)
    '
    ' anotaciones
    RNC$ = RECFILT$("NOTACLI", 1, MKI$(NCI%))
    URN& = ULTREG&("NOTACLI")
    '
    NOTAC$ = ""
    For UKL& = 1 To Len(RNC$) Step 4
       RENOTA1& = CVS(Mid$(RNC$, UKL&, 4))
       If RENOTA1& > 0 Then
          TTYY$ = TRIM$(Mid$(REGLEIDO$("NOTACLI", RENOTA1&), 3))
          If NOTAC$ <> "" Then
             NOTAC$ = NOTAC$ + Chr$(13) + Chr$(10)
          End If
          NOTAC$ = NOTAC$ + RTrim$(TTYY$)
       End If
    Next UKL&
    '
    'GRABACION
    On Error GoTo 0
    With Deudor12
      .AddNew
      !Stat_Cli = 1                        'STATUS ACTIVO
      !Nume_Cli = NCI%                     'CODIGO ENTERO
      !Codi_Cli = CODICUEN$                'CODIGO TEXT
      !Tcpr_Cli = 0                        'TIPO CTA
      !GrCo_Cli = GruCli%                  'TIPO CLI  VA A GRUPO CLI
      !Raso_Cli = RASO$                    'RAZON SOCIAL
      !Fant_Cli = Left$(FANTA$, 10)        'FANTASIA
      !Domi_Cli = DOMI$                    'DIRECCION
      !Loca_Cli = LOCA$                    'LOCALIDAD
      !Prov_Cli = PROVI$                   'COD PROVINCIA
      !Pais_Cli = PAISOK$                  'PAIS
      !Cpos_Cli = CPOSTX$                  'CODIGO POSTAL
      !Tele_Cli = TELE$                    'TELEFONO
'      !Tele2_Cli = TELEF2$                '2º TELEFONO
      !Faxi_Cli = FAXX$                    'FAX
      !Cuit_Cli = CUIT$                    'CUIT SIN VALIDAR
      !Cpag_Cli = CONDI%                   'CONDICION DE PAGO
      !Nibr_Cli = INGBRU$                  'IMP BRUTOS
      !Caib_Cli = CAIBRU$
      !Mail_Cli = MAILPRO$                 'MAIL
      !Piva_Cli = PIVA%                    'POS IVA
      '!Posivat_Cli = PIVATXT$             'POSICION DEL IVA EN TEXTO
      '!C_Ret = CRET$                      'RETENCION
      !Hora_Cli = HORARIX$
      !Lcre_Cli = LICRE                    'LIMITE CREDITO
      !Vend_Cli = COVEN%                   'CODIGO VENDEDOR
      !Ctac_Cli = CONTACTX$                'CONTACTO
      !Dsta_Cli = DESSTAN$                 'DESCUENTO
      !Lentre0_Cli = DOMIENT1$             'LUGAR DE ENTREGA
      !Lentre1_Cli = LOCAENT1$             'LOCALIDAD
      !Lentre2_Cli = TRANSPOR$             'TRANSPORTE
      !NTrans_Cli = NTRANX%
      !Hora_Ent = HORARIENT$
      !Zona_Cli = ZONCLI$
      !Obser_Entre = ""
      !Recar_Embal = 0
      !Gru_Div = GRUDIV%                   'GRUPO DIVISIONAL
      !CuMa_Cli = CUEMA%                   'CTA CONTABLE MADRE
      !Lpre_Cli = LIPRE%                   'LISTA DE PRECIOS
      !Nota_Cli = NOTAC$                   'ANOTACION
      .Update
      Call TRACE(20, U&, UR&)
    End With
  Next U&
  '
End Sub

Sub MIGRAPROVE()
   '
19 Picture1.ToolTipText = "Del - Proveedores"
  '
  Call APERBASDAT("PROVE")
  With Proved12
    On Error Resume Next
20    .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 20
    End If
    On Error GoTo 0
  End With
  '
  Picture1.ToolTipText = "Conversión de Proveedores"
  '
  'LECTURA
   UR& = ULTREG&("PROVED1")
  For U& = 1 To ULTREG&("PROVED1")
    '
    X1$ = REGLEIDO$("PROVED1", U&)
      NCI% = CVI(Mid$(X1$, 1, 2))
    X2$ = REGLEIDO$("PROVED2", U&)
    X3$ = FILTERGETRECORD$("BADAPROV", 1, MKI$(NCI%))
   'LECTURA
    NCITXT$ = CStr(NCI%)
    Do Until Len(NCITXT$) = 5
      NCITXT$ = "0" + NCITXT$
    Loop
    CODICUEN$ = "PR-" & NCITXT$
    '
    RASO$ = TRIM$(Mid$(X3$, 3, 62))
      If RASO$ = "" Then RASO$ = TRIM$(Mid$(X1$, 3, 30))
    FANTASI$ = TRIM$(Left(Mid$(X3$, 65, 32), 10))
      If FANTASI$ = "" Then FANTASI$ = TRIM$(Mid$(X2$, 65, 10))
    DOMI$ = TRIM$(Mid$(X3$, 97, 64))
      If DOMI$ = "" Then DOMI$ = TRIM$(Left$(X2$, 25))
    CPOSTX$ = TRIM$(Mid$(X3$, 161, 16))
      If CPOSTX$ = "" Then CPOSTX$ = Str$(CVI(Mid$(X2$, 28, 2)))
    CPOST% = 0: If XVALO(CPOSTX$) < 32767 Then CPOST% = XVALO(CPOSTX$)
    LOCA$ = TRIM$(Mid$(X3$, 177, 48))
      If LOCA$ = "" Then LOCA$ = TRIM$(Mid$(X2$, 30, 25))
    PAISOK$ = Mid$(X3$, 225, 32)
    TELE$ = TRIM$(Mid$(X3$, 257, 24))
      If TELE$ = "" Then TELE$ = TRIM$(Mid$(X2$, 55, 10))
    FAXX$ = Mid$(X3$, 281, 24)
    MAILPRO$ = Mid$(X3$, 305, 48)
    CONTACTX$ = Mid$(X3$, 385, 32)
    TIPOCTA% = Asc(Mid$(X2$, 119, 1))
    CUIT$ = TRIM$(Mid$(X2$, 75, 15))
      If CUIT$ <> "" Then VCC% = VALICUIT%(CUIT$)
    NUIBRU$ = TRIM$(Mid$(X2$, 90, 13))
    CODIPROVI1$ = Mid$(X2$, 117, 1)
    CODIVA1% = Asc(Mid$(X2$, 120, 1))
    IBRU$ = ""
    CONDI% = Asc(Mid$(X2$, 118, 1))
    LIPRE% = 0
    COVEN% = 0
    LICRE = 0
    CUEMA% = CVI(Mid$(X2$, 127, 2))
    '
    TRT$ = FILTERGETRECORD$("TARETIB", 1, MKI$(NCI%))
    CATIBX$ = Mid$(TRT$, 55, 1)

    ' anotaciones
    RNC$ = RECFILT$("NOTAPRO", 1, MKI$(NCI%))
    URN& = ULTREG&("NOTACLI")
    '
    NOTAC$ = ""
    For UKL& = 1 To Len(RNC$) Step 4
       RENOTA1& = CVS(Mid$(RNC$, UKL&, 4))
       If RENOTA1& > 0 Then
          TTYY$ = TRIM$(Mid$(REGLEIDO$("NOTAPRO", RENOTA1&), 3))
          If NOTAC$ <> "" Then
             NOTAC$ = NOTAC$ + Chr$(13) + Chr$(10)
          End If
          NOTAC$ = NOTAC$ + RTrim$(TTYY$)
       End If
    Next UKL&
    '
   'GRABACION
   With Proved12
    .AddNew
    !Nume_Cli = NCI%                     'CODIGO ENTERO
    !Stat_Cli = 1                        'STATUS ACTIVO
    !Codi_Cli = CODICUEN$                'CODIGO TEXT
    !Tcpr_Cli = TIPOCTA%                 'TIPO CTA
    !Raso_Cli = RASO$                    'RAZON SOCIAL
    !Fant_Cli = FANTASI$                 'FANTASIA
    !Domi_Cli = DOMI$                    'DIRECCION
    !Loca_Cli = LOCA$                    'LOCALIDAD
    !Prov_Cli = CODIPROVI1$              'COD PROVINCIA
    !Pais_Cli = PAISOK$                  'PAIS
    !Cpos_Cli = CPOSTX$                  'CODIGO POSTAL
    !Tele_Cli = TELE$                    'TELEFONO
    '!Tele2_Cli = TELEF2$                '2º TELEFONO
    !Faxi_Cli = FAXX$                    'FAX
    !Cuit_Cli = CUIT$                    'CUIT SIN VALIDAR
    !Cpag_Cli = CONDI%                   'CONDICION DE PAGO
    !Nibr_Cli = NUIBRU$                  'IMG BRUTOS
    !Caib_Cli = CATIBX$
    !Mail_Cli = MAILPRO$                 'MAIL
    !Piva_Cli = CODIVA1%                 'POS IVA
    '!C_Ret = CRET$                      'RETENCION
    !Ctac_Cli = CONTACTX$
    !CuMa_Cli = CUEMA%                   'CUENTA MADRE
    !Nota_Cli = NOTAC$                   'ANOTACION
    .Update
    Call TRACE(20, U&, UR&)
   End With
  Next U&
  '
End Sub

Sub MIGRAMASTER()
   '
29 Picture1.ToolTipText = "Del - Articulos"
  '
  Call APERBASDAT("MASTER")
  With Master
    On Error Resume Next
30    .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 30
    End If
    On Error GoTo 0
  End With
  '
  Picture1.ToolTipText = "Conversión de Articulos"
  '
  'LECTURA
  FIN& = ULTREG&("MASTER")
  MRECU% = 0
  For U& = 2 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("MASTER", U&)
     CI% = CVI(Mid$(X$, 1, 2))
     GIVA% = Asc(Mid$(X$, 3, 1))
     CODIT$ = Mid$(X$, 77, 16)
     CODALT$ = ""
     DESCOMP$ = ""
     DESIT$ = Mid$(X$, 5, 64)
     TIMAT% = Asc(Mid$(X$, 93, 1))
     CRITREP1% = Asc(Mid$(X$, 94, 1))
     UNIME$ = TRIM$(Mid$(X$, 109, 2))
     STOMI = CVI(Mid$(X$, 105, 2))
       If STOMI > 30000 Then STOMI = (-1) * STOMI / 100
       If STOMI < (-30000) Then STOMI = (-30000)
     LECON = CVI(Mid$(X$, 107, 2))
       If LECON > 30000 Then LECON = (-1) * LECON / 100
       If LECON < (-30000) Then LECON = (-30000)
     LOMAX = CVI(Mid$(X$, 103, 2))
     'LOMAX = 1
       If LOMAX < Abs(LECON) Then LOMAX = LECON
     KILUNI = CVS(Mid$(X$, 99, 4))
     KILMET = ""
     COSUN = CVS(Mid$(X$, 111, 4))
     MONEC% = Asc(Mid$(X$, 95, 1))
     FECOU% = CVI(Mid$(X$, 115, 2))
       PRH = CVI(Mid$(X$, 123, 2))
       While PRH > 30000: PRH = PRH - 10000: Wend
     PRHAB% = PRH
     '
     GRUCO% = CVI(Mid$(X$, 125, 2))
     '
     FCP = CVI(Mid$(X$, 127, 2)) / 100   ' % DE NACIONALIZACION
     '
     REV1$ = TRIM$(Mid$(X$, 117, 4))
     FERE1% = CVI(Mid$(X$, 121, 2))
     '
     URN& = ULTREG&("ESPECIFI")
     ESPEC$ = ESPECIFI$(CI%)
     '
     ATRIA$ = ""
     REG& = Asc(Mid$(X$, 97, 1))
     If REG& > 0 Then ATRIA$ = TRIM$(REGLEIDO$("TATALLES", REG&))
     '
     ATRIB$ = ""
     REG& = Asc(Mid$(X$, 96, 1))
     If REG& > 0 Then ATRIB$ = TRIM$(REGLEIDO$("TACOLOR", REG&))
     '
     On Error GoTo 0
32   If CODINT%(CODIT$) > 0 Then
       MRECU% = MRECU% + 1
       CODIT1$ = "zrecu-" + TRIM$(Str$(MRECU%))
       Call COMUNI("Codigo '" + CODIT$ + "' Pre-existente en BD.\Reemplazado por '" + CODIT1$ + "'.")
       CODIT$ = CODIT1$
       GoTo 32 ' XQ SINO  ENTRA EN UN LOOP SIN FIN
     End If
     '
     If CODINT%(CODIT$) = 0 Then
       With Master
         .AddNew
         !CODINT = CI%
         !NPLANO = CODALT$
         !DESCRIDOS = DESCOMP$
         !TALLE = ATRIA$
         !COLOR = ATRIB$
         !CODIGO = CODIT$
         !DESCRIPCION = DESIT$
         !Status = 1
         !UMEDIDA = AJUSTI$(UNIME$, 4)
         !TIMATE = TIMAT%
         !CRITREP = CRITREP1%
         !Monecos = MONEC%
         !FECOSTO = CVDAT(FECOU%)
         !COSUNI = COSUN
         !STOMIN = STOMI
         !LOREPOS = LECON
         !STOMAX = LOMAX
         !PROVHABI = PRHAB%
         !Alias = " "
         !DESCOMPO = 0
         !GRUPIVA = GIVA%
         !Grucoven = GRUCO%
         !CLinea = " "
         !Cfamilia = " "
         !Revision = REV1$
         !FEREVIS = CVDAT(FERE1%)
         !Observa = " "
         !ESPECIFICACION = ESPEC$
         !PESUNI = KILUNI
         !PESMETRO = 0 'KILMET
         !PORNACIO = FCP
        .Update
       End With
       '
     End If
     '
     Call TRACE(20, U&, FIN&)
   Next U&
  '
End Sub

Sub MIGRAFORMULA()
  '
  Picture1.ToolTipText = "Del - Formulas"
  '
  Call APERBASDAT("MASTER")
  With Formulas
    On Error Resume Next
35    .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 35
    End If
    On Error GoTo 0
  End With
  '
  Picture1.ToolTipText = "Conversión de Formulas"
  '
  'LECTURA
  FIN& = ULTREG&("MASTER")
  For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     CIXI% = U&
     CAIT% = 0
     RFF$ = RECFILT$("ESTRUNUE", 1, MKI$(CIXI%))
     If RFF$ <> "" Then
       REG& = CVS(Left$(RFF$, 4))
       If REG& > 0 Then
         CODICON$ = CODEXT$(CIXI%)
         X$ = REGLEIDO$("ESTRUNUE", REG&)
         DPS$ = Mid$(X$, 3, 120)
         If ULTREG&("ESTRUDOS") >= REG& Then
           DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", REG&)
         End If
         '
         RFF$ = RECFILT$("ESTRULAR", 1, MKI$(CIXI%))
         If Len(RFF$) >= 4 Then
           RELA& = CVS(RFF$)
           If RELA& > 0 Then
             If RELA& <= ULTREG&("ESTRULAR") Then
               DPS$ = DPS$ + Mid$(REGLEIDO$("ESTRULAR", RELA&), 3)
             End If
           End If
         End If
         '
         While Len(DPS$) Mod 8 <> 0
           DPS$ = DPS$ + Chr$(0)
         Wend
         '
         For K% = 1 To Len(DPS$) Step 8
           DPSX$ = Mid$(DPS$, K%, 8)
           CIXXII% = CVI(Left$(DPSX$, 2))
           If CIXXII% > 0 Then
             If Abs(CVS(Mid$(DPSX$, 3))) >= 0.000005 Then
               CODXX$ = CODEXT$(CIXXII%)
               If TRIM$(CODXX$) <> "" Then
                 CAIT% = CAIT% + 1
                 CIJ%(CAIT%) = CIXXII%
                 USOI(CAIT%) = CVS(Mid$(DPSX$, 3, 4))
                 PDI%(CAIT%) = Asc(Mid$(DPSX$, 7, 1))
                 '
                 DEXX$ = DESCRIT$(CIJ%(CAIT%))
                 UNXX$ = UNIMED$(CIJ%(CAIT%))
                 USON = USOI(CAIT%)
                 USOX$ = CSTRING$(MKS$(USON), 4, 12, 4, 2)
                 PDX$ = CSTRING$(Mid$(DPSX$, 7, 1), 1, 4, 0, 2)
                 XXX$ = Space$(64)
                 Call REPLA(XXX$, CODXX$, 1, 16)
                 Call REPLA(XXX$, DEXX$, 17, 28)
                 Call REPLA(XXX$, UNXX$, 47, 2)
                 Call REPLA(XXX$, USOX$, 49, 12)
                 Call REPLA(XXX$, PDX$, 61, 4)
                 EXPLOLIN(CAIT%) = XXX$
               End If
             End If
           End If
         Next K%
         '
         Call EXPLOGR(CODICON$, 1)
         '
       End If
     End If
  Next U&
  '
End Sub

Sub MIGRASECOPER()
  '
  Picture1.ToolTipText = "Del - Sec_Oper"
  '
36 Call ABRESECOP
   Dim RESECO$(256), DESOPE$(256)
  With Secoper
    On Error Resume Next
37    .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 37
    End If
    On Error GoTo 0
  End With
  '
  Picture1.ToolTipText = "Conversión de Secuencia de Operaciones"
  '
  'LECTURA
  FIN& = ULTREG&("MASTER")
  For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     CI% = U&
     RFFX$ = RECFILT$("SECOPER", 1, MKI$(CI%))
     CASEC% = 0
     For I1& = 1 To Len(RFFX$) Step 4
        I& = CVS(Mid$(RFFX$, I1&, 4))
        X$ = REGLEIDO$("SECOPER", I&)
        '
        NOPY% = CVI(Mid$(X$, 3, 2))
        If NOPY% > ULTOP% Then ULTOP% = NOPY%
        NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
        DOPE$ = Mid$(X$, 5, 48)
        REGO& = I&
        REGOX$ = AJUSTD$(Str$(REGO&), 6)
        SECO$ = NOPEX$ + "     " + DOPE$ + " " + REGOX$
        CASEC% = CASEC% + 1
        RESECO$(CASEC%) = Left$(X$ + String$(128, 0), 128) + Left$(NUHOPROF$(CI%, NOPY%) + Space$(16), 16)
        DESOPE$(CASEC%) = ""
     Next I1&
     '
     RNC$ = RECFILT$("OBSEOPER", 1, MKI$(CI%))
     URN& = ULTREG&("OBSEOPER")
     '
     For U3& = 1 To Len(RNC$) Step 4
        RENOTA& = CVS(Mid$(RNC$, U3&, 4))
        If RENOTA& > 0 Then
          If RENOTA& <= URN& Then
            TTZZ$ = REGLEIDO$("OBSEOPER", RENOTA&)
            NOPY% = CVI(Mid$(TTZZ$, 3, 2))
            For KKI% = 1 To CASEC%
              If CVI(Mid$(RESECO$(KKI%), 3, 2)) = NOPY% Then
                DESOPE$(KKI%) = DESOPE$(KKI%) + Mid$(TTZZ$, 5)
              End If
            Next KKI%
          End If
        End If
     Next U3&
     '
     For LL1% = 1 To CASEC%
        LL% = LL1% + 1
        With Secoper
        .AddNew
        !CODICONJ = CI%
        !CODXCONJ = CODEXT$(CI%)
        !NumeOper = CVI(Mid$(RESECO$(LL% - 1), 3, 2))
        !AltOper = Asc(Mid$(RESECO$(LL% - 1), 88, 1))
        !DescOper = TRIM$(Mid$(RESECO$(LL% - 1), 5, 48))
        !statoper = 1
        !CodMaq0 = TRIM$(Mid$(RESECO$(LL% - 1), 53, 6))
        !CodMaq1 = TRIM$(Mid$(RESECO$(LL% - 1), 103, 6))
        !CodMaq2 = TRIM$(Mid$(RESECO$(LL% - 1), 109, 6))
        !CodMaq3 = TRIM$(Mid$(RESECO$(LL% - 1), 115, 6))
        !HorsPrep = CVS(Mid$(RESECO$(LL% - 1), 59, 4))
        !HorsFijo = CVS(Mid$(RESECO$(LL% - 1), 63, 4))
        !MinuStan = CVS(Mid$(RESECO$(LL% - 1), 67, 4))
        !HorsLimp = CVS(Mid$(RESECO$(LL% - 1), 79, 4))
        !HorsMHer = CVS(Mid$(RESECO$(LL% - 1), 95, 4))
        !PieCiclo = CVS(Mid$(RESECO$(LL% - 1), 99, 4))
        !PorDemor = CVS(Mid$(RESECO$(LL% - 1), 75, 4))
        If !MinuStan > 0.005 Then
          !Cadencia = 60 / !MinuStan
        End If
        !CantiOps = CVS(Mid$(RESECO$(LL% - 1), 83, 4))
        !CategOps = Asc(Mid$(RESECO$(LL% - 1), 87, 1))
        !CoefSolp = CVI(Mid$(RESECO$(LL% - 1), 89, 2))
        !LOTESTAN = CVS(Mid$(RESECO$(LL% - 1), 91, 4))
        !ProgCNum = TRIM$(Mid$(RESECO$(LL% - 1), 121, 8))
        !HoProFab = TRIM$(Mid$(RESECO$(LL% - 1), 129, 16))
        !ProTerOp = CVI(Mid$(RESECO$(LL% - 1), 149, 2))
        !PreUnid = CVD(Mid$(RESECO$(LL% - 1), 151, 8))
        !MonePrec = CVI(Mid$(RESECO$(LL% - 1), 159, 2))
        !FechPrec = CVDAT(CVI(Mid$(RESECO$(LL% - 1), 161, 2)))
        '
        !OBSEOPER = DESOPE(LL% - 1)
        !NUOPER_PREP = CANPRE
        !CATOPER_PREP = CATPRE%
        .Update
        End With
     Next LL1%
     '
  Next U&
  '
End Sub

Sub MIGRAPEDIDOS()
  Picture1.ToolTipText = "Del - Pedidos"
  '
  'PEDENCA
  Call APERBASDAT("PEDCLI")
  With PedEnca
    On Error Resume Next
40    .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 40
    End If
    On Error GoTo 0
  End With
  '
  Picture1.ToolTipText = "Conversión de Pedidos"
  '
  FIN& = ULTREG&("PEDENCA")
  For U& = 1 To FIN&
    X$ = REGLEIDO$("PEDENCA", U&)
    'LECTURA
    NUPE& = CVS(Mid$(X$, 1, 4)) 'NUMERO DE PEDIDO
    REFE$ = Mid$(X$, 5, 48) 'REFERENCIA
    FEPE% = CVI(Mid$(X$, 53, 2)) 'FECHA DE PEDIDO
    NUCLI% = CVI(Mid$(X$, 55, 2)) 'NUMERO DE CLIENTE
    NUOC1$ = Mid$(X$, 57, 16) 'NUMERO DE ORDEN DE COMPRA
    DOMIENT1$ = Mid$(X$, 73, 32) 'DOMICILIO DE ENTREGA
    CPLOCA$ = Mid$(X$, 105, 32) 'CP-LOCALIDAD
    TRANSP1$ = Mid$(X$, 137, 32) 'TRANPORTE
    OBS1$ = Mid$(X$, 169, 32) 'OBSERVACION UNO
    OBS2$ = Mid$(X$, 201, 32) 'OBSERVACION DOS
    OBS3$ = Mid$(X$, 233, 32) 'OBSERVACION TRES
    FESOL1% = CVI(Mid$(X$, 265, 2)) 'FECHA DE ENTREGA SOLICITADA
    STAT1% = Asc(Mid$(X$, 267, 1)) 'STATUS (1-ABIER,3-CUMP,9-ANUL)
    LIPRE1% = Asc(Mid$(X$, 268, 1)) 'LISTA DE PRECIOS
    COPA1% = Asc(Mid$(X$, 269, 1)) 'CONDICION DE PAGO
    VEND1% = Asc(Mid$(X$, 269, 1)) 'VENDEDOR
    DTO1$ = Mid$(X$, 271, 14) 'DESCUENTO APLICADO
    CAMPVER$ = Mid$(X$, 289, 16)
    'GRABACION
    With PedEnca
      .AddNew
        !NroPed = NUPE&
        !Referencia = REFE$
        !TipoPed = "" 'POR CONVERSIÓN
        !FECHEMIS = CVDAT(FEPE%)
        !NroClie = NUCLI%
        !RasoClie = RASOCLI$(NUCLI%)
        !NroOcom = NUOC1$
        !DOMIENT = DOMIENT1$
        !LOCAENT = CPLOCA$
        !TranspEnt = TRANSP1$
        !Observa = OBS1$ + " " + OBS2$ + " " + OBS3$
        !FeSolEnt = CVDAT(FESOL1%)
        !Status = STAT1%
        !FinVigen = CVDAT(0)
        !ListPre = LIPRE1%
        !CondPag = COPA1%
        !Vended = VEND1%
        !TDescApli = DTO1$
        !AnotaPed = ""
        !FApro_Nope = CVDAT(0)
        '!Usua_Apro = 0'POR CONVERSIÓN
      .Update
    End With
    Call TRACE(20, U&, FIN&)
  Next U&
  'PEDDETA
  Picture1.ToolTipText = "Del - Conversión de Pedidos"
  Call APERBASDAT("PEDCLI")
  With PedDeta
    On Error Resume Next
50    .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 50
    End If
    On Error GoTo 0
  End With
  '
  Picture1.ToolTipText = "Conversión de Pedidos"
  'LECTURA
  FIN& = ULTREG&("PEDDETA")
  For U& = 1 To FIN&
    X$ = REGLEIDO$("PEDDETA", U&)
    'LECTURA
    NUPE& = CVS(Mid$(X$, 1, 4)) 'NUMERO DE PEDIDO
    FEPE% = CVI(Mid$(X$, 5, 2)) 'FECHA DEL PEDIDO
    NUCLI% = CVI(Mid$(X$, 7, 2)) 'NUMERO DE CLIENTE
    NUOC1$ = Mid$(X$, 9, 16) 'NUMERO DE ORDEN DE COMPRA
    FESOL1% = CVI(Mid$(X$, 25, 2)) 'FECHA ENTREGA SOLICITADA
    STAT1% = Asc(Mid$(X$, 27, 1)) 'STATUS
    LisPre% = Asc(Mid$(X$, 28, 1)) 'LISTA DE PRECIO
    COPA1% = Asc(Mid$(X$, 29, 1)) 'CONDICION DE PAGO
    VEND1% = Asc(Mid$(X$, 30, 1)) 'VENDEDOR
    PORCOM = CVI(Mid$(X$, 31, 2)) 'PORCENTAJE DE COMISION
    CI% = CVI(Mid$(X$, 33, 2)) 'CODIGO DEL MATERIAL
    COLO$ = Mid$(X$, 35, 16) 'COLOR
    TALL$ = ""
    MONEMI1% = Asc(Mid$(X$, 52, 1))
    TIPOCA1 = CVS(Mid$(X$, 53, 4))
    LAR1 = CVS(Mid$(X$, 59, 4)) 'LARGO
    ANC1 = CVS(Mid$(X$, 63, 4)) 'ANCHO
    ESP1 = CVS(Mid$(X$, 67, 4)) 'ESPESOR
    PRELI1 = CVS(Mid$(X$, 71, 4)) 'PRECIO DE LISTA
    CASI1 = CVS(Mid$(X$, 75, 4)) 'CANTIDAD ASIGNADA
    CAPE1 = CVS(Mid$(X$, 79, 4)) 'CANTIDAD PEDIDA
    PREUNI1 = CVS(Mid$(X$, 83, 4)) 'PRECIO UNITARIO
    IMPTOT1 = CVD(Mid$(X$, 87, 8)) 'IMPORTE TOTAL
    CAEN1 = CVS(Mid$(X$, 95, 4)) 'CANTIDAD ENTREGADA
    CASA1 = CVS(Mid$(X$, 99, 4)) 'CANTIDAD SALDO
    FEUE% = CVI(Mid$(X$, 103, 2)) 'FECHA DE ULTIMA ENTREGA
    NURE1& = CVS(Mid$(X$, 105, 4)) 'NUMERO DE REMITO
    'RT-X-0002-00000030
    NURE3$ = ""
    If NURE1& > 0 Then
      NURE2$ = CStr(NURE1&)
      Do Until Len(NURE2$) = 8
       NURE2$ = "0" + NURE2$
      Loop
      NURE3$ = "RT-X-0000-" + NURE2$
    End If
    'GRABACION
    With PedDeta
      .AddNew
        !NroRemi = NURE3$ 'NUMERO PEDIDO FORMATEADO
        !TPedItem = ""
        !NroPed = NUPE&
        !TipoPed = "" 'POR CONVERSIÓN
        !FECHEMIS = CVDAT(FEPE%)
        !NroClie = NUCLI%
        !RasoClie = RASOCLI$(NUCLI%)
        !NroOcom = NUOC1$
        !FeSolEnt = CVDAT(FESOL1%)
        !FinVigen = CVDAT(0)
        !Status = STAT1%
        !ListPre = LisPre%
        !CondPag = COPA1%
        !Vended = VEND1%
        !PorComVend = PORCOM
        !NUORITEM = 0 'POR CONVERSIÓN
        !CodiInt = CI%
        !CodiExt = CODEXT$(CI%)
        !Descrip = DESCRIT0$(CI%)
        !SubCod = ""
        !TColor = COLO$
        !TTalle = TALL$
        !TLargo = LAR1
        !TAncho = ANC1
        !TEspesor = ESP1
        !VenRep = ""
        !MoneEmis = MONEMI1%
        !TIPOCAM = TIPOCA1
        !PreList = PRELI1
        !CanAsig = CASI1
        !CanPed = CAPE1
        !PreUnid = PREUNI1
        !ImpoTot = IMPTOT1
        !CantEnt = CAEN1
        !CanSaldo = CASA1
        !FeUltEnt = CVDAT(FEUE%)
        !FApro_Nope = CVDAT(FEPE%)
      .Update
    End With
    Call TRACE(20, U&, FIN&)
  Next U&
  '
End Sub

Sub MIGRACOMPRAS()
   '
59 Picture1.ToolTipText = "Del - Ordenes de Compra"
  'OCOMENCA
  Call APERBASDAT("COMPRAS")
  With OcomEnca
    On Error Resume Next
60    .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 60
    End If
    On Error GoTo 0
  End With
  Picture1.ToolTipText = "Conversión de Ordenes de Compra"
  '
  FIN& = ULTREG&("OCOMENCA")
  For U& = 1 To FIN&
    X$ = REGLEIDO$("OCOMENCA", U&)
    'LECTURA
    NUCO& = CVS(Mid$(X$, 1, 4)) 'NUMERO DE O/C
    REFE$ = Mid$(X$, 5, 44) 'REFERENCIA
    FEMICO% = CVI(Mid$(X$, 49, 2)) 'FECHA DE EMISION
    PROVE% = CVI(Mid$(X$, 51, 2)) 'PROVEEDOR
    TIPSUM% = Asc(Mid$(X$, 53, 1)) 'TIPO DE SUMINISTRO
    ESTOC% = Asc(Mid$(X$, 54, 1)) 'ESTADO DE LA O/C
    FEGEN% = CVI(Mid$(X$, 55, 2)) 'FECHA DE GENERACIÓN
    HORGEN% = CVI(Mid$(X$, 57, 2)) 'HORA DE GENERACIÓN
    USUGEN% = CVI(Mid$(X$, 59, 2)) 'USUARIO GENERACIÓN
    FEMOD% = CVI(Mid$(X$, 61, 2)) 'FECHA DE MODIFICACIÓN
    HOMOD% = CVI(Mid$(X$, 63, 2)) 'HORA REAL DE MODIFICACIÓN
    USUMOD% = CVI(Mid$(X$, 65, 2)) 'USUARIO MODIFICACION
    FEANU% = CVI(Mid$(X$, 67, 2)) 'FECHA REAL ANULACIÓN
    HOANU% = CVI(Mid$(X$, 69, 2)) 'HORA REAL ANULACIÓN
    USUANU% = CVI(Mid$(X$, 71, 2)) 'USUARIO ANULACIÓN
    MOTANUL$ = Mid$(X$, 73, 10) 'MOTIVO DE LA ANULACIÓN
    MOEVA% = Asc(Mid$(X$, 83, 1)) 'MONEDA DE VALORIZACIÓN
    MOEMI% = Asc(Mid$(X$, 84, 1)) 'MONEDA DE EMISIÓN
    TICA = CVS(Mid$(X$, 85, 4)) 'TIPO DE CAMBIO
    IBRU = CVD(Mid$(X$, 89, 8)) 'IMPORTE BRUTO SIN IVA
    PODE$ = Mid$(X$, 97, 16) 'PORCENTAJE DE DESCUENTO
    IMDE = CVD(Mid$(X$, 113, 8)) 'IMPORTE DEL DESCUENTO
    IMIV = CVD(Mid$(X$, 121, 8)) 'IMPORTE DEL IVA
    IMOT = CVD(Mid$(X$, 129, 8)) 'IMPORTE OTROS
    IMOT2 = CVD(Mid$(X$, 137, 8)) 'IMPORTE OTROS DOS
    TOTOC = CVD(Mid$(X$, 145, 8)) 'TOTAL ORDEN DE COMPRA
    COPA% = CVI(Mid$(X$, 153, 2)) 'CONDICIÓN DE PAGO
    ATESR$ = Mid$(X$, 163, 30) 'ATE. SEÑOR:
    OBS$ = Mid$(X$, 193, 64) 'OBSERVACIONES CERO
    OBS1$ = Mid$(X$, 257, 64) 'OBSERVACIONES UNO
    OBS2$ = Mid$(X$, 321, 64) 'OBSERVACIONES DOS
    OBS3$ = Mid$(X$, 385, 64) 'OBSERVACIONES TRES
    OBS4$ = Mid$(X$, 449, 64) 'OBSERVACIONES CUATRO
    'GRABACIÓN
    With OcomEnca
      .AddNew
        !Nume_Ocom = NUCO&
        !Refe_Ocom = REFE$
        !Femi_Ocom = CVDAT(FEMICO%)
        !Npro_Ocom = PROVE%
        !RaSo_Prov = Left$(RASOCLI$(PROVE% * (-1)), 48)
        !TSum_Ocom = TIPSUM%
        !Stat_Ocom = ESTOC%
        !FGen_Ocom = CVDAT(FEGEN%)
        !Usua_Gen = USUGEN%
        !FApro_Ocom = CVDAT(FEMICO%)
        !Usua_Apro = 0 'POR CONVERSIÓN
        !FMod_Ocom = CVDAT(FEMOD%)
        !Usua_Mod = USUMOD%
        !FAnu_Ocom = FEANU%
        !Usua_Anul = USUANU%
        !Moti_Anul = MOTANUL$
        !Mone_Cos = MOEVA%
        !Mone_Emis = MOEMI%
        !Tcam_Ocom = TICA
        !ToBru_Ocom = IBRU
        !Pdes_Ocom = PODE$
        !Ides_Ocom = IMDE
        !Iiva_Ocom = IMIV
        !Iotr1_Ocom = IMOT
        !Iotr2_Ocom = IMOT2
        !Total_Ocom = TOTOC
        !Cpag_Ocom = COPA%
        !Embala_Ocom = 0 'Por Conversión
        !AtSr_Ocom = ATESR$
        !Obse_Ocom = OBS$ + " " + OBS1$ + " " + OBS2$ + " " + OBS3$ + " " + OBS4$
        !Lentrega = "" 'Por Conversión
      .Update
    End With
    Call TRACE(20, U&, FIN&)
  Next U&
  'OCOMDETA
  Picture1.ToolTipText = "Del - Ordenes de Compra"
    Call APERBASDAT("COMPRAS")
    With OcomDeta
      On Error Resume Next
70      .MoveFirst
      If Err = 0 Then
        .Delete
        GoTo 70
      End If
      On Error GoTo 0
  End With
  '
  Picture1.ToolTipText = "Conversión de Ordenes de Compra"
  '
  FIN& = ULTREG&("OCOMDETA")
  NUANT& = 0
  For U& = 1 To FIN&
    X$ = REGLEIDO$("OCOMDETA", U&)
    'LECTURA
    NUCO& = CVS(Mid$(X$, 1, 4)) 'NUMERO DE O/C
    FEMICO% = CVI(Mid$(X$, 5, 2)) 'FECHA DE EMISION
    PROVE% = CVI(Mid$(X$, 7, 2)) 'PROVEEDOR
    CI% = CVI(Mid$(X$, 9, 2)) 'CODIGO DEL MATERIAL
    UNME$ = Mid$(X$, 11, 4) 'UNIDAD DE MEDIDA SEGUN MAESTRO
    COLO$ = Mid$(X$, 15, 16) 'COLOR
    MEDI$ = Mid$(X$, 31, 24) 'MEDIDAS
    CAPE = CVS(Mid$(X$, 55, 4)) 'CANTIDAD PEDIDA
    UNOC$ = Mid$(X$, 59, 4) 'UNIDAD DE COMPRA
    COEUN = CVS(Mid$(X$, 63, 4)) 'COEFICIENTE DE UNIDAD
    PUNI = CVD(Mid$(X$, 67, 8)) 'PRECIO UNITARIO POR UNIDAD
    DEIT = CVS(Mid$(X$, 75, 4)) 'PORCENTAJE DE DESCUENTO DEL ITEM
    PUNINET = CVD(Mid$(X$, 79, 8)) 'PRECIO UNITARIO NETO
    IMPONET = CVD(Mid$(X$, 87, 8)) 'IMPORTE NETO DEL ITEM
    FENSOL% = CVI(Mid$(X$, 95, 2)) 'FECHA DE ENTREGA SOLICITADA
    COMP$ = Mid$(X$, 97, 2) 'CONSIGNA MATERIA PRIMA
    CARE = CVS(Mid$(X$, 99, 4)) 'CANTIDAD RECIBIDAD
    COPE% = CVI(Mid$(X$, 103, 2)) 'CODIGO DE OPERACIÓN
    DOCUORI$ = Mid$(X$, 105, 12) 'DOCUMENTO DE ORIGEN
    NUORSE& = CVS(Mid$(X$, 117, 4)) 'NUMERO DE ORDEN DE SERVICIO
    MONCO% = Asc(Mid$(X$, 121, 1)) 'MONEDA DE COSTEO
    MONE% = Asc(Mid$(X$, 122, 1)) 'MONEDA DE EMISION
    TIPSUM% = Asc(Mid$(X$, 123, 1)) 'TIPO DE SUMINISTRO
    STAOC% = Asc(Mid$(X$, 124, 1)) 'STATUS DE LA O/C
    'GRABACIÓN
    With OcomDeta
      .AddNew
        !Nume_Ocom = NUCO&
          If NUCO& <> NUANT& Then
             NORI% = 0
             NUANT& = NUCO&
          End If
          NORI% = NORI% + 1
        !NuOrd_Item = NORI%
        !Femi_Ocom = CVDAT(FEMICO%)
        !Npro_Ocom = PROVE%
        !RaSo_Prov = RASOCLI$(PROVE% * (-1))
        !Cod_Inte = CI%
        !COD_EXTE = CODEXT$(CI%)
        !Descritem = DESCRIT0$(CI%)
        !Delibre = "" 'POR CONVERSIÓN
        !IContab = 0 'POR CONVERSIÓN
        !Sect_Ocom = "" 'POR CONVERSIÓN
        !Uni_Mas = UNME$
        !Colo_Ocom = COLO$
        !Medi_Ocom = Left$(MEDI$, 20)
        !Valo_Unid = "" 'POR CONVERSIÓN
        !Unids_Ocom = 0 'POR CONVERSIÓN
        !Cant_Ocom = CAPE
        !Uni_Com = UNOC$
        !Coef_Unids = COEUN
        !Pre_Unit = PUNI
        !Desc_Item = DEIT
        !PrUn_Neto = PUNINET
        !INet_Item = IMPONET
        !Fentre_Sol = CVDAT(FENSOL%)
        !Consimat = COMP$
        !Cant_Rec = CARE
        !Cod_Oper = COPE%
        !Doc_Orig = DOCUORI$
        !N_OrServ = NUORSE&
        !Mone_Cos = MONCO%
        !Mone_Emis = MONE%
        !TSum_Ocom = TIPSUM%
        !Stat_Ocom = STAOC%
        !Nro_Otrab = "" 'POR CONVERSIÓN
        !Nro_Grup = "" 'POR CONVERSIÓN
      .Update
    End With
    Call TRACE(20, U&, FIN&)
  Next U&
  '
End Sub

Sub MIGRASTOCKS()
   '
79 Picture1.ToolTipText = "Del - Movimientos de Stock"
  Call APERBASDAT("STOCKS")
  With Movisto
    On Error Resume Next
80    .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 80
    End If
    On Error GoTo 0
  End With
  '
  Picture1.ToolTipText = "Conversión de Movimientos de Stocks"
  '
  FIN& = ULTREG&("MOVISTO")
  DOCANT$ = ""
  PREFIREM$ = TRIM$(Control$("", "PREFIREM")): If PREFIREM$ = "" Then PREFIREM$ = "0001"
  For U& = 1 To FIN&
    X$ = REGLEIDO$("MOVISTO", U&)
    'LECTURA
      FEC% = CVI(Mid$(X$, 1, 2)) 'FECHA DE MOVIMIENTO
      CI% = CVI(Mid$(X$, 3, 2)) 'CODIGO DE PIEZA
      LOT$ = Mid$(X$, 5, 16) 'LOTE / PARTIDA
      COMO$ = Mid$(X$, 21, 2) 'CODIGO DE MOVIMIENTO
      DOC1$ = Mid$(X$, 23, 10) 'DOCUMENTO PRIMARIO ASOCIADO
      DOC1LAR$ = DOC1$
      If COMO$ = "RT" Then
         NUMEDO$ = TRIM$(Mid$(DOC1$, 4))
         While Len(NUMEDO$) < 8: NUMEDO$ = "0" + NUMEDO$: Wend
         DOC1LAR$ = Left$(DOC1$, 2) + "-X-" + PREFIREM$ + "-" + NUMEDO$
      End If
      DOC2$ = Mid$(X$, 33, 10) 'DOCUMENTO SECUNDARIO
      DOC2LAR$ = DOC2$
      If Left$(DOC2$, 2) = "FC" Then
         NUMEDO$ = TRIM$(Mid$(DOC2$, 4))
         While Len(NUMEDO$) < 8: NUMEDO$ = "0" + NUMEDO$: Wend
         DOC2LAR$ = Left$(DOC2$, 2) + "-0001-" + NUMEDO$ + "-A"
      End If
      REFE$ = Mid$(X$, 43, 30) 'REFERENCIA
      VAUN# = CVD(Mid$(X$, 73, 8)) 'VALOR UNITARIO
      MONE1$ = Mid$(X$, 81, 2) 'MONEDA
      DEP% = Asc(Mid$(X$, 83, 1)) 'DEPOSITO
      NUCLI% = CVI(Mid$(X$, 84, 2)) 'NUEMRO DE CLIENTE
      NUPRO% = CVI(Mid$(X$, 86, 2)) 'NUMERO DE PROVEEDOR
      CAEN = CVD(Mid$(X$, 88, 8)) 'CANTIDAD ENTRADA
      CASA = CVD(Mid$(X$, 96, 8)) 'CANTIDAD SALIDA
      CASALD = CVD(Mid$(X$, 104, 8)) 'CANTIDAD SALDO
      Fill$ = Mid$(X$, 112, 10) 'FILLER NO IDENTIFICADO
      FEMO% = CVI(Mid$(X$, 122, 2)) 'FECHA DE MOVIMIENTO
      USU% = Asc(Mid$(X$, 126, 1)) 'USUARIO
    'GRABACIÓN
    With Movisto
      .AddNew
        !FECHMOV = CVDAT(FEC%)
        !Cod_Inte = CI%
        !COD_EXTE = CODEXT$(CI%)
        !Descrip = DESCRIT0(CI%)
        !IdenLote = LOT$
        !CodiMovi = COMO$
        !DoPriCor = DOC1$
        !DOPRILAR = DOC1LAR$
           If DOC1$ <> DOCANT$ Then
             NORI% = 0
             DOCANT$ = DOC1$
           End If
        !DoSecCor = DOC2$
        !DoSecLar = DOC2LAR$
        !NumRem_Cli = "" 'POR CONVERSIÓN
           NORI% = NORI% + 1
        !NUORITEM = NORI%
        !NuPedCli = 0 'POR CONVERSIÓN
        !ReferCor = REFE$
        !VALOUNIT = VAUN
        !MoneVal = XVALO(MONE1$)
        !DEPOMOVI = DEP%
        !Nume_Clie = NUCLI%
        !Nume_Prov = NUPRO%
        !CANTINGRE = CAEN
        !CANTSALID = CASA
        !CantSaldo = CASALD
        !FeReMovi = CVDAT(FEMO%)
      .Update
    End With
      '
    Call TRACE(20, U&, FIN&)
  Next U&
  '
End Sub

Sub MIGRABOLREC()
  Picture1.ToolTipText = "Del - Boletas de Rececpción"
  Call APERBASDAT("STOCKS")
  With Bolrecep
    On Error Resume Next
90    .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 90
    End If
    On Error GoTo 0
  End With
  '
  Picture1.ToolTipText = "Conversión de Boletas de Recepción"
  '
  FIN& = ULTREG&("BOLRECEP")
  For U& = 1 To FIN&
    X$ = REGLEIDO$("BOLRECEP", U&)
    'LECTURA
    NUBORE& = CVS(Mid$(X$, 1, 4)) 'NUMERO DE BOLETA DE RECEPCIÓN
    DES$ = Mid$(X$, 5, 64) 'DESCRIPCION
    LOT$ = Mid$(X$, 69, 12) 'IDENTIFICADOR DE LOTE
    FERE% = CVI(Mid$(X$, 81, 2)) 'FECHA DE RECEPCION
    CI% = CVI(Mid$(X$, 83, 2)) 'CODIGO DEL MATERIAL
    UM$ = Mid$(X$, 85, 4) 'UNIDAD DE MEDIDA
    CANOM = CVD(Mid$(X$, 89, 8)) 'CANITDAD NOMINAL
    DEPEN% = Asc(Mid$(X$, 97, 1)) 'DEPOSITO DE ENTRADA
    PROVE% = CVI(Mid$(X$, 105, 2)) 'NUMERO DE PROVEEDOR
    LOTPROV$ = Mid$(X$, 107, 12) 'LOTE DEL PROVEEDOR
    REPROV$ = Mid$(X$, 119, 16) 'NUMERO DE REMITO DEL PROVEEDOR
    FAPROV$ = Mid$(X$, 135, 16) 'NUMERO DE FACTURA DEL PROVEEDOR
    NUOC& = CVS(Mid$(X$, 151, 4)) 'NUMERO DE ORDEN DE COMPRA
    MOVAUN$ = Mid$(X$, 155, 4) 'MONEDA VALOR UNITARIO
    VAUN = CVS(Mid$(X$, 159, 4)) 'VALOR UNITARIO
    CAPRO1 = CVD(Mid$(X$, 163, 8)) 'CANTIDAD APROBADA
    CARE = CVD(Mid$(X$, 171, 8)) 'CANTIDAD RECHAZADA
    CAFA = CVD(Mid$(X$, 179, 8)) 'CANTIDAD FALTANTE
    FELI% = CVI(Mid$(X$, 187, 2)) 'FECHA DE LIBERACIÓN
    NUORSE& = CVS(Mid$(X$, 189, 4)) 'NUMERO DE ORDEN DE SERVICIO
    'GRABACIÓN
    Call APERBASDAT("STOCKS")
    With Bolrecep
      .AddNew
        !NumeBoRe = NUBORE&
        !Descrip = DES$
        !IdenLote = LOT$
        !FechRecep = CVDAT(FERE%)
        !Cod_Inte = CI%
        !COD_EXTE = CODEXT$(CI%)
        !DescripItem = DESCRIT0(CI%)
        !Unidad = UM$
        !CantNom = CANOM
        !DepoEnt = DEPEN%
        !NumProv = PROVE%
        !RaSo_Prov = RASOCLI$(PROVE% * (-1))
        !LoteCol_Prov = LOTPROV$
        !NumRem_Prov = REPROV$
        !NumFac_Prov = FAPROV$
        !NumOC = NUOC&
        !IContab = 0 'POR CONVERSIÓN
        !MoneVaUni = MOVAUN$
        !Valo_Unit = VAUN
        !CanToApro = CAPRO1
        !CanToRech = CARE
        !CanToFalta = CAFA
        !FechLibe = CVDAT(FELI%)
        !NumeOrdSer = NUORSE&
        !Observa = ""
        !Status = 1 'POR CONVERSIÓN
        !UsuApro = "" 'POR CONVERSIÓN
        !FeHoApro = CVDAT(FERE%) 'POR CONVERSIÓN
      .Update
    End With
    Call TRACE(20, U&, FIN&)
  Next U&
  '
End Sub

Sub MIGRACONSIGNA()
  Picture1.ToolTipText = "Del - Material en Consignación"
  Call APERBASDAT("STOCKS")
  With MaConsig
    On Error Resume Next
100   .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 100
    End If
    On Error GoTo 0
  End With
  '
  Picture1.ToolTipText = "Conversión Material en Consignación"
  '
  FIN& = ULTREG&("MACONSIG")
  For U& = 1 To FIN&
    X$ = REGLEIDO$("MACONSIG", U&)
    'LECTURA
    FERE% = CVI(Mid$(X$, 1, 2)) 'FECHA DE RESERVA
    PROVE% = CVI(Mid$(X$, 3, 2)) 'NUMERO DE PROVEEDOR
    SECC% = CVS(Mid$(X$, 5, 4)) 'SECCION
    NUOF$ = Mid$(X$, 9, 12) 'NUMERO DE ORDEN DE FABRICACIÓN
    NUOC1$ = Mid$(X$, 21, 12) 'NUMERO DE ORDEN DE COMPRA
    CI% = CVI(Mid$(X$, 33, 2)) 'CODIGO INTERNO DEL MATERIAL
    LOT$ = Mid$(X$, 35, 16) 'LOTE - PARTIDA
    UM$ = Mid$(X$, 51, 4) 'UNIDAD DE MEDIDA
    CASI = CVD(Mid$(X$, 55, 8)) 'CANTIDAD ASIGNADA
    CAEN = CVD(Mid$(X$, 63, 8)) 'CANTIDAD ENTREGADA
    FENT% = CVI(Mid$(X$, 71, 2)) 'FECHA DE ENTREGA
    NURE$ = Mid$(X$, 73, 16) 'NUMERO DE REMITO
    VAUN = CVD(Mid$(X$, 93, 8)) 'VALOR UNITARIO
    MON% = CVI(Mid$(X$, 101, 2)) 'MONEDA
    VATO = CVD(Mid$(X$, 103, 8)) 'VALOR TOTAL
    CADE = CVD(Mid$(X$, 111, 8)) 'CANTIDAD RENDIDA (DEVUELTA)
    SAPE = CVD(Mid$(X$, 119, 8)) 'SALDO PENDIENTE
    OBS$ = Mid$(X$, 129, 112) 'OBSERVACIONES
    'GRABACIÓN
    Call APERBASDAT("STOCKS")
    With MaConsig
      .AddNew
        !Femi_Ocom = CVDAT(FERE%)
        !Npro_Ocom = PROVE%
        !RaSo_Prov = RASOCLI$(PROVE% * (-1))
        !Secc_Ejec = SECC%
        !Deno_Secc = "" 'POR CONVERSIÓN
        !Nro_Orfa = NUOF$
        !Nro_Ocom = NUOC1$
        !Cod_Inte = CI%
        !COD_EXTE = CODEXT$(CI%)
        !Descritem = DESCRIT0(CI%)
        !Lote_Ocom = LOT$
        !Unidad = UM$
        !Cant_Asig = CASI
        !Cant_Entreg = CAEN
        !Fecha_Entreg = CVDAT(FENT%)
        !NRemi_Consig = NURE$
        !Valo_Unit = VAUN
        !MONEDA = MON%
        !Valo_Tot = VATO
        !Cant_Rendid = CADE
        !Saldo_Pend = SAPE
        !Observa = OBS$
      .Update
    End With
    Call TRACE(20, U&, FIN&)
  Next U&
  '
End Sub

Sub MIGRAORFAB()
  Picture1.ToolTipText = "Del - Ordenes de Fabricación"
  Call APERBASDAT("ORFAB")
  With OrdeFabr
    On Error Resume Next
110   .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 110
    End If
    On Error GoTo 0
  End With
  '
  Picture1.ToolTipText = "Conversión Ordenes de Fabricación"
  '
  FIN& = ULTREG&("ORDEFABR")
  For U& = 1 To FIN&
    X$ = REGLEIDO$("ORDEFABR", U&)
    'LECTURA
    NUOF1& = CVS(Mid$(X$, 1, 4)) 'NUMERO DE ORDEN DE FABRICACION
    REFEOF$ = Mid$(X$, 5, 64) 'REFERENCIA
    LOTOF$ = Mid$(X$, 69, 12) 'IDENTIFICADOR DE LOTE
    FEGEOF% = CVI(Mid$(X$, 81, 2)) 'FECHA DE GENERACIÓN
    COPIOF% = CVI(Mid$(X$, 83, 2)) 'CODIGO DE PIEZA
    UMOF$ = Mid$(X$, 85, 4) 'UNIDAD DE MEDIDA
    CAPROYE = CVD(Mid$(X$, 89, 8)) 'CANTIDAD PROYECTADA
    DEOF$ = Mid$(X$, 97, 8) 'CLIENTE O DEPOSITO DESTINO
    FESOF% = CVI(Mid$(X$, 105, 2)) 'FECHA SOLICITADA DE ENTREGA
    PRIOF% = Asc(Mid$(X$, 107, 1)) 'PRIORIDAD DE LA OF
    STATOF% = Asc(Mid$(X$, 108, 1)) 'STATUS DE LA OF
    FELAOF% = CVI(Mid$(X$, 109, 2)) 'FECHA DE LANZAMIENTO DE LA OF
    COSEOF% = CVS(Mid$(X$, 113, 4)) 'CODIGO DEL SECTOR PRODUCTIVO
    FECIOF% = CVI(Mid$(X$, 117, 2)) 'FECHA DE CIERRE DE LA OF
    CAPROF = CVD(Mid$(X$, 119, 8)) 'CANTIDAD APROBADA OF
    CAREOF = CVD(Mid$(X$, 127, 8)) 'CANTIDAD RECHAZADA OF
    IDOROF$ = Mid$(X$, 135, 8) 'IDENTIFICADOR ORDEN RETRABAJO
    FENOF% = CVI(Mid$(X$, 147, 2)) 'FECHA PROYECTADA DE ENTREGA
    FETEOF% = CVI(Mid$(X$, 149, 2)) 'FECHA TEMPRANA DE ENTREGA
    FERECOF% = CVI(Mid$(X$, 151, 2)) 'FECHA DE RECALCULO
    OBSOF$ = Mid$(X$, 155, 4) 'OBSERVACIONES
    'GRABACIÓN
    Call APERBASDAT("ORFAB")
    With OrdeFabr
      .AddNew
        !NumeOrFa = NUOF1&
        !Referen = REFEOF$
        !IdSubOr = LOTOF$
        !FechGen = CVDAT(FEGEOF%)
        !Cod_Inte = COPIOF%
        !COD_EXTE = CODEXT$(COPIOF%)
        !Descrip = DESCRIT0(COPIOF%)
        !Unidad = UMOF$
        !CanProy = CAPROYE
        !Destino = DEOF$
        !FechEnSol = CVDAT(FESOF%)
        !Priorid = PRIOF%
        !StatuOrf = STATOF%
        !FechLan = CVDAT(FELAOF%)
        !SecciPro = COSEOF%
        !FechCierr = CVDAT(FECIOF%)
        !CanApro = CAPROF
        !CanRech = CAREOF
        !IdenRetra = IDOROF$
        !FechProyEntre = CVDAT(FENOF%)
        !FechTempEntre = CVDAT(FETEOF%)
        !FechRecal = CVDAT(FERECOF%)
        !Observa = OBSOF$
      .Update
    End With
    Call TRACE(20, U&, FIN&)
  Next U&
  '
End Sub

Sub MIGRARESERVA()
  Picture1.ToolTipText = "Del - Reservas"
  Call APERBASDAT("RESERVA")
  With Reserva
    On Error Resume Next
120   .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 120
    End If
    On Error GoTo 0
  End With
  '
  Picture1.ToolTipText = "Conversión Reservas de Materiales"
  '
  FIN& = ULTREG&("RESERVA")
  For U& = 1 To FIN&
    X$ = REGLEIDO$("RESERVA", U&)
    'LECTURA
    CI% = CVI(Mid$(X$, 1, 2)) 'CODIGO INTERNO DEL MATERIAL
    NOF$ = Mid$(X$, 3, 12) 'NUMERO DE ORDEN DE FABRICACION
    FERE% = CVI(Mid$(X$, 17, 2)) 'FECHA DE RESERVA
    CARE = CVD(Mid$(X$, 19, 8)) 'CANTIDAD RESERVADA
    CACO = CVD(Mid$(X$, 27, 8)) 'CANTIDAD CONSUMIDA
    SALEN = CVD(Mid$(X$, 35, 8)) 'SALDO A ENTREGAR
    LOT$ = Mid$(X$, 43, 16) 'LOTE O PARTIDA
    DEP% = Asc(Mid$(X$, 59, 1)) 'DEPOSITO
    USOUNI = CVS(Mid$(X$, 75, 4)) 'USO UNITARIO
    'GRABACIÓN
    With Reserva
      .AddNew
        !Cod_Inte = CI%
        !COD_EXTE = CODEXT$(CI%)
        !Descrip = DESCRIT0(CI%)
        !IdSubOr = NOF$
        !FechRes = CVDAT(FERE%)
        !CantRes = CARE
        !CantCons = CACO
        !Saldo_Entre = SALEN
        !IdenLote = LOT$
        !Deposito = DEP%
        !UsoUnit = USOUNI
      .Update
    End With
  Next U&
  '
End Sub

Sub MIGRAPLANCUEN()

5  Call ABREPLANCUEN
  Picture1.ToolTipText = "Del - Plan de Cuentas"
  Plancuen.Close
  Set Plancuen = Contable.OpenRecordset("PLANCUEN", dbOpenDynaset)
  With Plancuen
    On Error Resume Next
1201   .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 1201
    End If
    On Error GoTo 0
  End With
  '
  Picture1.ToolTipText = "Plan de Cuentas Contables"
  '
     FIN& = ULTREG&("ECUECOT")
     If FIN& > 0 Then
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         IVP$ = REGLEIDO$("ECUECOT", U&)
         '
         CUE% = CVI(Mid$(IVP$, 41, 2))
         NIVEI% = 2
         '
         If CUE% > 0 Then
           ZZ$ = Space$(256)
           COCUE$ = TRIM$(Left$(IVP$, 12))
           DECUE$ = TRIM$(Mid$(IVP$, 13, 28))
           '
           If COCUE$ <> "" Then
             Call REPLA(ZZ$, NVX$, 1, 6)
             Call REPLA(ZZ$, COCUE$, 9, 12)
             Call REPLA(ZZ$, DECUE$, 23, 44)
             Call REPLA(ZZ$, DECUE$, 129, 128)
             '
             RECUEN& = CUE%
             XX$ = REGLEIDO$("CUENTAS", RECUEN&)
             TICUEX$ = UCase$(Mid$(XX$, 13, 1))
             If TICUEX$ = "H" Then
                 TICUEX$ = "Haber"
               ElseIf TICUEX$ = "I" Then
                 TICUEX$ = "Ingresos"
               ElseIf TICUEX$ = "G" Then
                 TICUEX$ = "Gastos"
               Else
                 TICUEX$ = "Debe"
             End If
             YY$ = REGLEIDO$("DENOCUE", RECUEN&)
             '
             IMPUBLE$ = ""
             STATUX$ = ""
             '
             With Plancuen
               .AddNew
               !CodExCue = COCUE$
               !DENOCUE = DECUE$
               !CODINCUE = CUE%
               !STATCUE = 1
               !NIVAGRU = NIVEI%
               !TipoCue = TICUEX$
               !NuOrdIt = 0
               !POSIMPU = 1
               !Observa = ""
               !Salini = 0
               !Sudebe = 0
               !Suhabe = 0
               !Result = 0
               !Salfin = 0
               !IndiAjus = 0
               !MoneExtr = 0
               .Update
             End With
             '
           End If
         End If
       Next U&
     End If
     '
     ' PONE EN ORDEN SEGÚN PLAN GENERAL DE CUENTAS
     FIN& = ULTREG&("PLAGRAL")
     For U& = 1 To FIN&
       XX$ = REGLEIDO$("PLAGRAL", U&)
       CII% = CVI(XX$)
       NIVEI% = CVI(Mid$(XX$, 3, 2))
       '
       With Plancuen
          .FindFirst "CODINCUE = " & CII%
          If .NoMatch = False Then
             .Edit
             !NIVAGRU = NIVEI%
             !NuOrdIt = U&
             .Update
          End If
       End With
       '
     Next U&
     '
     ' LOS QUE NO ESTÁN EN EL PLAN DE CUENTAS VAN AL FINAL Y SUSPENDIDAS
     U& = FIN&
     With Plancuen
        .MoveFirst
        Do While Not .EOF
           If !NuOrdIt = 0 Then
             U& = U& + 1
             .Edit
             !NuOrdIt = U&
             !STATCUE = 1
             .Update
           End If
        .MoveNext
        Loop
     End With
     '
End Sub

Sub BORRACUECORR()
  '
6 Picture1.ToolTipText = "Del - Cajabanc"
  Call APERBASDAT("CABAN")
  With CajaBanc
    On Error Resume Next
130   .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 130
    End If
    On Error GoTo 0
  End With
End Sub

Sub MIGRAPRESU()
  '\\\OT-07-0414-FL-05/11/07///
  Call ABREPRESVEN
  With PresuVen
    On Error Resume Next
10  .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 10
    End If
    On Error GoTo 0
  End With
  
  'LECTURA
  ULTPRESU& = 0
  FIN& = ULTREG&("PRESUPU")
  For U& = 1 To FIN&
    Call TRACE(20, U&, FIN&)
    X3$ = REGLEIDO$("PRESUPU", U&)
    '
    'LECTURA
    NUMPRE& = CVS(Mid$(X3$, 1, 4))           'NUMERO DE PRESUPUESTO
    FECHE% = CVI(Mid$(X3$, 5, 2))            'FECHA DE EMISION
    NUMECLI1% = CVI(Mid$(X3$, 7, 2))         'NUMERO DE CLIENTE
    REFER$ = Mid$(X3$, 9, 16)                'REFERENCIA
    ATENSR$ = Mid$(X3$, 25, 40)              'ATENCION SR.
    LISTPREC% = Asc(Mid$(X3$, 65, 1))        'LISTA DE PRECIOS
    CONPAGO% = Asc(Mid$(X3$, 66, 1))         'CONDICION DE PAGO
    VENDEDO% = Asc(Mid$(X3$, 67, 1))         'VENDEDOR
    CODART% = CVI(Mid$(X3$, 69, 2))          'CODIGO DEL MATERIAL
    PRECLIST# = CVS(Mid$(X3$, 71, 4))        'PRECIO DE LISTA
    DESCUEN$ = Mid$(X3$, 75, 14)             'DESCUENTO
    DESCUE# = CVS(Mid$(X3$, 89, 4))          '% DESCUENTO
    IMPDESC# = CVS(Mid$(X3$, 93, 4))         'IMPORTE DESCUENTO
    PRECUNIT# = CVS(Mid$(X3$, 97, 4))        'PRECIO UNITARIO
    CANTIDAD# = CVS(Mid$(X3$, 101, 4))       'CANTIDAD
    IMPTOTNET# = CVD(Mid$(X3$, 105, 8))      'IMPORTE TOTAL NETO
    FECVENT% = CVI(Mid$(X3$, 113, 2))        'FECHA DE VENTA
    NUORDVEN& = CVS(Mid$(X3$, 115, 4))       'NUMERO DE RODEN DE VENTA
    NUMFACT& = CVS(Mid$(X3$, 119, 4))        'NUMERO DE FACTURA
    '
    'GRABACION
    On Error GoTo 0
    With PresuVen
      .AddNew
        !NroPres = NUMPRE&
        !Cod_Rubro = " "
        !NumeCRubro = 0
        !Cod_Anio = " "
        FECHE1 = CVDAT(FECHE%)
        !FECHEMIS = FECHE1
        !NroClie = NUMECLI1%
        RASOCLIEN$ = ECOARCH$("DEUDOR1", MKI(NUMECLI1%))
        !RasoClie = RASOCLIEN$
        !Referen = REFER$
        !AteSen = ATENSR$
        !LisPre = LISTPREC%
        !ConPag = CONPAGO%
        !Vended = VENDEDO%
        !MonedEmi = 1
        If ULTPRESU& <> NUMPRE& Then
          NUORIT% = 1: ULTPRESU& = NUMPRE&
        Else
          NUORIT% = NUORIT% + 1
        End If
        !NuOrdIt = NUORIT%
        !CodiInt = CODART%
        !CodiExt = CODEXT$(CODART%)
           Descrip$ = ECOARCH$("MASTER", MKI$(CODART%))
        !Descrip = Left$(Descrip$, 64)
        !ESPECIF = Descrip$
        !PreList = PRECLIST#
        !TDESCUEN = DESCUEN$
        !PorcDesc = DESCUE#
        !ImpoDesc = IMPDESC#
        !PreUnit = PRECUNIT#
        !Cantit = CANTIDAD#
        !ImToNe = IMPTOTNET#
        FECVENT1 = CVDAT(FECVENT%)
        !FechVent = FECVENT1
        !NroPedi = NUORDVEN&
        NroFac_Cor = NUMFACT&
      .Update
    End With
  Next U&
  '
  '\\\OT-07-0414-FIN///
End Sub

Sub MIGRAPRECIOS()
    '
    '\\\OT-07-0437-14-11-07///
   Call APERBASDAT("PEDCLI")
   With ListaPrec
     On Error Resume Next
10   .MoveFirst
     If Err < 1 Then
        .Delete
        GoTo 10
     End If
     On Error GoTo 0
   End With
   '
   'Lee y graba cada registro del archivo LISTAS.DAT a la tabla LISTAPRE
   FIN& = ULTREG&("LISTAS")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X3$ = REGLEIDO$("LISTAS", U&)
      '
      'LECTURA
      CODLIST% = Asc(Mid$(X3$, 1, 1))          'CODIGO DE LISTA
      DENOMIN$ = Mid$(X3$, 2, 46)              'DENOMINACION DE LA LISTA
      FORMATO% = Asc(Mid$(X3$, 48, 1))         'FORMATO
      DECIMA% = Asc(Mid$(X3$, 49, 1))          'DECIMALES
      MONEDA% = Asc(Mid$(X3$, 50, 1))          'MONEDA
      ABREVI$ = Mid$(X3$, 51, 12)              'ABREVIATURA
      VIGENC% = CVI(Mid$(X3$, 63, 2))          'VIGENCIA
      '
      'GRABACION
      With ListaPrec
        .AddNew
        !Cod_Lista = CODLIST%
        !DENOM_LISTA = DENOMIN$
        !VIGENCIA = VIGENC%
        !FORMATO = FORMATO%
        !DECIMALES = DECIMA%
        !MONEDA = MONEDA%
        !DESC_ABREV = ABREVI$
        .Update
      End With
      '
      ARCHIPRE$ = TRIM$(Str$(CODLIST%))
      While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
      ARCHIPRE$ = "LIPRE" + ARCHIPRE$
      FINA& = ULTREG&(ARCHIPRE$)
      ' Lee y graba cada dato del archivo LIPREnnn en la tabla PREVENTA
      For J& = 1 To FINA&
         'LECTURA
         X3$ = REGLEIDO$(ARCHIPRE$, U&)
         CODPIE% = CVI(Mid$(X3$, 1, 2))         'CODIGO DE PIEZA
         DESCRI$ = Mid$(X3$, 3, 42)             'DESCRIPCION
         CARACONT$ = Mid$(X3$, 45, 1)           'CARACTER DE CONTROL
         NUMORIGI$ = Mid$(X3$, 46, 16)          'NUMERO ORIGINAL
         PRECUNIT# = CVD(Mid$(X3$, 62, 8))      'PRECIO UNITARIO
         CANTLIMI% = CVI(Mid$(X3$, 70, 2))      'CANTIDAD LIMITE
         PORDESCU% = Asc(Mid$(X3$, 72, 1))      'PORCENTAJE DE DESCUENTO
         '
         'GRABACION
         With Preventa
            .AddNew
            !Cod_Lista = CODLIST%
            !COD_PIEZA = CODPIE%
            !Num_Orden = J&
            !DESC_PIEZA = DESCRI$
            !Denom_Pieza = DESCRIT$(CODPIE%)
            !CARAC_CONTROL = CARACONT$
            !NUM_ORIGINAL = NUMORIGI$
            !PRECIO_UNITARIO = PRECUNIT#
            !CANT_LIMITE = CANTLIMI%
            !PORC_LIMITE = PORDESCU%
            .Update
         End With
      Next J&
      '
   Next U&
   '\\\OT-07-0437-FIN///
End Sub

Sub MIGRACONTA()
   '
   AVARIA$ = LOADFILE$(LUDAT$ + "VARIAB.DAT") ', 0, 32)
   V$ = Mid$(AVARIA$, 33, 32)
   COMEJ1% = CVI(Mid$(V$, 29, 2))
   FINEJ1% = CVI(Mid$(V$, 31, 2))
   DES% = COMEJ1%: HAS% = FINEJ1%
   '
   MODOXX2% = 1 ' INCLUYE LOS ASIENTOS DE CAJA Y BANCOS
   MODOXX3% = 1 ' RENUMERA LOS ASIENTOS
   '
   Screen.MousePointer = 11
   Sudebe# = 0: Suhabe# = 0
   '
   SORTLIST.Clear
   FIN& = MAXDIARIO%
   For KK% = 1 To MAXDIARIO%
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      Call LEEPASDIA(KK%, 2)
      '
      If Asc(Left$(E8$, 1)) = 1 Then GoTo 29 ' SALTEA ASIENTOS RESUMEN DE VENTAS
      If Asc(Left$(E8$, 1)) = 2 Then GoTo 29 ' SALTEA ASIENTOS RESUMEN DE COMPRAS
      If Asc(Left$(E8$, 1)) = 3 Then GoTo 29 ' SALTEA ASIENTOS RESUMEN DE CAJA Y BANCOS
      '
      FEAS% = CVI(E15$)
      RESOR$ = Left$(TRIM$(Str$(10000 + FEAS%)) + Space$(10), 10)
      RESOR$ = RESOR$ + "DI" + TRIM$(Str$(100000 + KK%))
      SORTLIST.AddItem RESOR$
29 Next KK%
   '
' ACA VER COMO SE INCLUYEN LOS DE CAJA Y BANCOS
   If MODOXX2% = 1 Then
     FIN& = MAXDIA&
     For KKL& = 1 To MAXDIA&
       UU& = KKL&
       Call TRACE(20, UU&, FIN&)
       Call LEEPASCAJ(KKL&, 2)
       If Asc(Left$(D8$, 1)) = 1 Then GoTo 39 ' SALTEA ASIENTOS RESUMEN DE VENTAS
       If Asc(Left$(D8$, 1)) = 2 Then GoTo 39 ' SALTEA ASIENTOS RESUMEN DE COMPRAS
       FEAS% = CVI(D15$)
       RESOR$ = Left$(TRIM$(Str$(10000 + FEAS%)) + Space$(10), 10)
       RESOR$ = RESOR$ + "CA" + TRIM$(Str$(100000 + KKL&))
       SORTLIST.AddItem RESOR$
39   Next KKL&
   End If
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   FIN& = SORTLIST.ListCount    ' MAXDIARIO%
   YATRA% = 0: JJ& = 0: ASIANT = 0: ASIANTX = 0
   '
   NUASDEF& = 0
   N1% = FILEOPEN%(LUDAT$ + "ASIENTOS.TXT", 2, 256)

   For KK% = 1 To FIN&     'MAXDIARIO%
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      RESOR$ = SORTLIST.List(KK% - 1)
      REGII& = Val(Mid$(RESOR$, 13)) - 100000
      GRUXI$ = TRIM$(Mid$(RESOR$, 11, 2))
      SGX1% = 1
      FEASI% = Val(TRIM$(Left$(RESOR$, 10))) - 10000
      MODOLE% = 1
      If FEASI% >= DES% And FEASI% <= HAS% Then MODOLE% = 2
      If GRUXI$ = "DI" Then
         REGIJ% = REGII&
         Call LEEPASDIA(REGIJ%, 2)
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
         Call LEEPASCAJ(REGII&, 2)
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
      If FEAS% < DES% Then
         If IPAS# > 0 Then
              Sudebe# = Sudebe# + IPAS#
            Else
              Suhabe# = Suhabe# - IPAS#
         End If
         GoTo 90
      End If
      '
      FECHASTA$ = FECHATEX$(FEAS%)
      If FEAS% <= HAS% Then
        If NUAS * SGX1% <> ASIANTX Then
           ASIANTX = NUAS * SGX1%
           Z$ = Space$(256)
           Call REPLA(Z$, FECHASTA$, 16, 8)
           Call REPLA(Z$, FORMATNUM$(NUASDEF&, "F6.0"), 1, 6)
           JJ& = JJ& + 1
         End If
         '
         Call REPLA(Z$, FECHASTA$, 16, 8)
         Call REPLA(Z$, FORMATNUM$(NUASDEF&, "F6.0"), 1, 6)
         Call REPLA(Z$, REFASIEN$, 32, 48)
         Call REPLA(Z$, TRIM$(CODCUE$(CUE%)), 83, 16)
         Call REPLA(Z$, REFE$, 100, 46)
         Call REPLA(Z$, Space$(32), 146, 32)
         If IPAS# > 0 Then
              Call REPLA(Z$, FORMATNUM$(IPAS#, "F13.2"), 146, 13)
            Else
              Call REPLA(Z$, FORMATNUM$((-1) * IPAS#, "F13.2"), 162, 13)
         End If
         Print #N1%, Z$
      End If
90 Next KK%
   '
   Close #N1%
   Screen.MousePointer = 1
   '
99 End Sub

Sub MIGRACONTB()
   '
   Dim LISTA$(256)
   Call ABREPLANCUEN
   '
20 On Error Resume Next
   Set ENCASIGRA = Contable.OpenRecordset("ENCABEASIEN", dbOpenDynaset)
   If Err Then
     If COMALTER%("No se Puede Abrir 'ENCABEASIEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 20
     GoTo 90
   End If
   '
21 Set DETASIGRA = Contable.OpenRecordset("DETAASIEN", dbOpenDynaset)
   If Err Then
     If COMALTER%("No se Puede Abrir 'DETAASIEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 21
     GoTo 90
   End If
   '
   N1% = FILEOPEN%(LUDAT$ + "ASIENTOS.TXT", 1, 256)
   JJ& = 0
   FIN& = LOF(N1%) / 256
   Do While Not EOF(N1%)
     Call TRACE(20, U&, FIN&)
     Line Input #N1%, TTXX$
     NUMERO = XVALO(TRIM$(Mid$(TTXX$, 1, 12)))
     If NUMERO <> NUMEROANT Then
        If JJ& > 0 Then GoSub GRABATODO
        JJ& = 0
        NUMEROANT = NUMERO
        FECHA$ = TRIM$(Mid$(TTXX$, 16, 8))
        NUASI$ = NUPROA$(FECHA$)
        REFEX8$ = TRIM$(Mid$(TTXX$, 32, 48))
        SUMADEBE# = 0
     End If
     JJ& = JJ& + 1
     LISTA(JJ&) = TTXX$
     IDEBE# = XVALO(TRIM$(Mid$(LISTA$(JJ&), 146, 13)))
     SUMADEBE# = SUMADEBE# + ROUND(IDEBE#)
   Loop
   GoSub GRABATODO
   '
   ENCASIGRA.Close
   DETASIGRA.Close
   Call COMUNI("Fin de la Importación Definitiva")
   Exit Sub
   '
GRABATODO:
   With ENCASIGRA
      .AddNew
      !NUIDASIEN = NUASI$
      !Refeasien = REFEX8$
          FECHAU% = FECHANUM(FECHA$)
      !FECHASIEN = CVDAT(FECHAU%)
          PPXX% = InStr(NUASI$, "-")
      !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
      !SucuAsien = 0
      !NuDfAsien = 0
      !TIPOASIEN = 0
         !StatAsien = 1  ' activo
      !EditAsien = 1
      !QPasAsien = JJ&
      !DebeAsien = SUMADEBE#
      !Observa = " "
      !Nume_Usuar = USNBR%
      .Update
   End With
   '
   ORPASE% = 0
   For K2% = 1 To JJ&
      COCUE$ = TRIM$(Mid$(LISTA$(K2%), 83, 16))
      DEPASE$ = TRIM$(Mid$(LISTA$(K2%), 100, 46))
      If Left$(DEPASE$, 1) = "-" Then DEPASE$ = TRIM$(Mid$(DEPASE$, 2))
      IDEBE# = XVALO(TRIM$(Mid$(LISTA$(K2%), 146, 13)))
         SUMADEBE# = SUMADEBE# + ROUND(IDEBE#)
      IHABE# = XVALO(TRIM$(Mid$(LISTA$(K2%), 162, 13)))
      MACON$ = " "
      CENCOS1$ = " "
      '
      With DETASIGRA
         .AddNew
         !NUIDASIEN = NUASI$
           ORPASE% = ORPASE% + 1
         !OrNuPase = ORPASE%
         !REFEPASE = DEPASE$
           FECHAU% = FECHANUM(FECHA$)
         !FECHASIEN = CVDAT(FECHAU%)
           PPXX% = InStr(NUASI$, "-")
         !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
         !SucuAsien = 0
         !NuDfAsien = 0
         !StatPase = 1
         !CUECOEXT = COCUE$
         !CUECOINT = CUEINT%(COCUE$)
         !IDebePase = IDEBE#
         !IHabePase = IHABE#
         !MARCONSI = MACON$
         !CenCos = CENCOS1$ '\\\OT-06-0244-WAR-16/06/06///
         .Update
      End With
   Next K2%
   '
Return

90 End Sub

Function NUPROA$(FECHA$)
   '
   NUPROA$ = ""
   '
   FEASX = FECHANUM(FECHA$)
   fechas$ = FECHATEX$(FEASX)
   FECHAU% = FEASX
   FEASYYYY = CVDAT(FECHAU%)
   NUIDAS$ = Mid$(fechas$, 7, 2) + Mid$(fechas$, 4, 2) + Mid$(fechas$, 1, 2)
   If Left$(NUIDAS$, 1) = "0" Then NUIDAS$ = Mid$(NUIDAS$, 2)
   ORDAS% = 0
   '
   SUCUI% = 0
   Set SELECASIEN = Contable.OpenRecordset("SELECT * FROM EncaBEAsien WHERE (STR$(FECHASIEN) = '" & FEASYYYY & "' and SucuAsien = " & SUCUI% & ") ORDER BY  ORFEASIEN DESC;", dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   SELECASIEN.MoveFirst
   If Err Then GoTo 20
   ORDAS% = SELECASIEN!OrFeAsien
   '
20 On Error GoTo 0
   ORDAS% = ORDAS% + 1
   ORDA$ = TRIM$(Str$(ORDAS%))
   While Len(ORDA$) < 3: ORDA$ = "0" + ORDA$: Wend
   NUPROA$ = NUIDAS$ + "-" + ORDA$
   '
End Function


