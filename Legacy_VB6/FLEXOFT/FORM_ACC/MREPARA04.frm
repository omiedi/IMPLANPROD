VERSION 5.00
Begin VB.Form MREPARA04 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reparaciones - Actualización y Consultas"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   7440
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MREPARA04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   1065
      Left            =   3700
      ScaleHeight     =   1005
      ScaleWidth      =   3735
      TabIndex        =   26
      Top             =   0
      Width           =   3795
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   945
         ScaleHeight     =   75
         ScaleWidth      =   1020
         TabIndex        =   30
         Top             =   700
         Width           =   1080
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   12000
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
         Height          =   700
         Left            =   2900
         Picture         =   "MREPARA04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   180
         Width           =   650
      End
      Begin VB.CommandButton Command6 
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
         Height          =   700
         Left            =   2205
         Picture         =   "MREPARA04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   180
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   105
         Picture         =   "MREPARA04.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Anotador de Novedades"
         Top             =   180
         Width           =   650
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consultas - Re-impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   105
      TabIndex        =   15
      Top             =   2940
      Width           =   7200
      Begin VB.OptionButton Option9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Despacho de Reparación"
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
         Left            =   2100
         TabIndex        =   32
         Top             =   420
         Width           =   1515
      End
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
         Height          =   960
         Left            =   6510
         Picture         =   "MREPARA04.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Genera y Emite la Consulta Marcada"
         Top             =   400
         Width           =   560
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Control Estadístico"
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
         Left            =   4410
         TabIndex        =   25
         Top             =   1150
         Width           =   2115
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "O/R's en Curso"
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
         Left            =   4410
         TabIndex        =   24
         Top             =   850
         Width           =   1920
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Tareas por O/Rep."
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
         Left            =   4410
         TabIndex        =   23
         Top             =   550
         Width           =   2130
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Material por O/Rep."
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
         Left            =   4410
         TabIndex        =   22
         Top             =   250
         Width           =   2115
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Orden de Reparación"
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
         Left            =   2835
         TabIndex        =   20
         Top             =   960
         Width           =   1300
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Presupuesto al Cliente"
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
         Left            =   1470
         TabIndex        =   19
         Top             =   945
         Width           =   1380
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Presupuesto Interno "
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
         Left            =   105
         TabIndex        =   18
         Top             =   960
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Planilla de Reparación"
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
         Left            =   700
         TabIndex        =   17
         Top             =   420
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   4200
         X2              =   4200
         Y1              =   120
         Y2              =   1800
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Preparación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   105
      TabIndex        =   11
      Top             =   105
      Width           =   3480
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   12
         Top             =   525
         Width           =   2430
         Begin VB.CommandButton Command20 
            BackColor       =   &H8000000A&
            Height          =   600
            Left            =   1810
            Picture         =   "MREPARA04.frx":0EAA
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Búsqueda Cliente por CUIT, Domicilio, Nombre de Fantasía, etc"
            Top             =   0
            Width           =   560
         End
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
            Picture         =   "MREPARA04.frx":2210
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Planilla de Presupuesto"
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
            TabIndex        =   14
            Top             =   105
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Presupuestos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   105
      TabIndex        =   7
      Top             =   1600
      Width           =   3480
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   8
         Top             =   420
         Width           =   2430
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
            Picture         =   "MREPARA04.frx":251A
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Generar Presupuesto por una Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar y Emitir un Presupuesto"
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
            TabIndex        =   10
            Top             =   105
            Width           =   1380
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reparaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   3700
      TabIndex        =   0
      Top             =   1110
      Width           =   3585
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   1
         Top             =   945
         Width           =   2325
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
            Picture         =   "MREPARA04.frx":2824
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cierre de Orden de Reparación"
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
            TabIndex        =   3
            Top             =   60
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   4
         Top             =   420
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
            Left            =   0
            Picture         =   "MREPARA04.frx":2B2E
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Generar una Orden de Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Orden de Reparación"
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
            TabIndex        =   6
            Top             =   70
            Width           =   1485
         End
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7680
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Repara 
      Caption         =   "Reparaciones"
      Begin VB.Menu Planipre 
         Caption         =   "Planilla de Presupuestación"
         Begin VB.Menu Pladir 
            Caption         =   "Directa"
         End
         Begin VB.Menu Plabuscli 
            Caption         =   "Búsqueda de Cliente"
         End
      End
      Begin VB.Menu plh3 
         Caption         =   "-"
      End
      Begin VB.Menu Genpresu 
         Caption         =   "Generar Presupuesto"
      End
      Begin VB.Menu Ordrep 
         Caption         =   "Orden de Reparación"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu Cierrorep 
         Caption         =   "Cierre de O/Reparación"
      End
      Begin VB.Menu plh2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu Consul 
      Caption         =   "Consultas y Re-Impresión"
      Begin VB.Menu Anota 
         Caption         =   "Anotador de Seguimiento"
      End
      Begin VB.Menu plh5 
         Caption         =   "-"
      End
      Begin VB.Menu Reimpre 
         Caption         =   "Re-Impresiones"
         Begin VB.Menu Plarepa 
            Caption         =   "Planilla de Reparación"
         End
         Begin VB.Menu Presuint 
            Caption         =   "Presupuesto Interno"
         End
         Begin VB.Menu Presucli 
            Caption         =   "Presupuesto al Cliente"
         End
         Begin VB.Menu Orderep 
            Caption         =   "Orden de Reparación"
         End
      End
      Begin VB.Menu plh21 
         Caption         =   "-"
      End
      Begin VB.Menu Mater_Orep 
         Caption         =   "Materiales Consumidos"
      End
      Begin VB.Menu Oper_Orep 
         Caption         =   "Tareas Realizadas"
      End
   End
   Begin VB.Menu Tablaux 
      Caption         =   "Tablas Auxiliares"
      Begin VB.Menu Master 
         Caption         =   "Maestro de Artículos"
      End
      Begin VB.Menu Clientes 
         Caption         =   "Padrón de Clientes"
      End
      Begin VB.Menu plh4 
         Caption         =   "-"
      End
      Begin VB.Menu Lispre 
         Caption         =   "Lista de Precios"
      End
   End
   Begin VB.Menu Setup 
      Caption         =   "Configuración"
   End
   Begin VB.Menu Help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "MREPARA04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Anota_Click()
   Call Command8_Click
End Sub

Private Sub Cierrorep_Click()
   Call Command2_Click
End Sub

Private Sub Clientes_Click()
   Call PLAPRERE04.Command11_Click
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    Call CIERROREPA
    Command2.Enabled = True
    '
End Sub

Private Sub Command20_Click()
    Command20.Enabled = False
    Call PLAPREPB
    Command20.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call PRESUREPA
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call OREPARA
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call PLAPREPA
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   CONFREPA04.Show 1
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    '
    Command7.Enabled = False
    If Option5.Value = True Then
        Call MATEROREP
      ElseIf Option6.Value = True Then
        Call OPEROREP
      ElseIf Option7.Value = True Then
        Call OREPABI
      ElseIf Option8.Value = True Then
        Call ESTAREP
      Else
        Call PRESEDOC
    End If
    Command7.Enabled = True
    '
End Sub

Sub MATEROREP()
    '
    Screen.MousePointer = 11
    Call COPYSTRU("PRESREPA", "OREPARPE")
    JJ& = 0
    For U& = 1 To ULTREG&("PRESREPA")
      XX$ = REGLEIDO$("PRESREPA", U&)
      If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
        JJ& = JJ& + 1
        Call GRAREG("OREPARPE", XX$, JJ&)
      End If
    Next U&
    Call SETULTREG("OREPARPE", JJ&)
    Call CIERRARCH("OREPARPE")
    Screen.MousePointer = 1
    '
    Call FINAL("*COREPUES")

99 End Sub
'
Sub OPEROREP()
   '
   Call COPYSTRU("INDIOTRA", "INDIOREP") ' INDICE DE TODAS LAS o/t'S
   '
   Screen.MousePointer = 11
   JJ& = 0
   REBLA$ = REGBLAN$("INDIOREP")
   FIN& = ULTREG&("PRESREPA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
       NOREPARA& = CVS(Left$(XX$, 4))
       NORFA$ = "OR-" + TRIM$(Str$(NOREPARA&))
       DORFA$ = TRIM$(Mid$(XX$, 5, 60))
       FEMI% = CVI(Mid$(XX$, 121, 2))
       '
       ZZ$ = REBLA$
       Call REPLA(ZZ$, NORFA$, 1, 12)
       Call REPLA(ZZ$, DORFA$, 13, 48)
       Call REPLA(ZZ$, MKI$(FEMI%), 61, 2)
       JJ& = JJ& + 1
       Call GRAREG("INDIOREP", ZZ$, JJ&)
     End If
   Next U&
   Call SETULTREG("INDIOREP", JJ&)
   Call CIERRARCH("INDIOREP")
   Screen.MousePointer = 1
   '
   Call FINAL("*COMOBRAD")
   '
End Sub
'
Sub PRESEDOC()
    '
3   If Option1.Value = True Then
          ASEL$ = "PRESREPA"
          TSEL$ = "Indice General Planillas de Reparación"
          TIDOCX$ = TRIM$("Planilla de Reparación")
        ElseIf Option2.Value = True Then
          '
          If DERACCE%("VEPREPAR", "Visualizar Presupuestos de Reparación") < 2 Then
            Exit Sub
          End If
          '
          Screen.MousePointer = 11
          Call COPYSTRU("PRESREPA", "PREPAEMI")
          JJ& = 0
          For U& = 1 To ULTREG&("PRESREPA")
            XX$ = REGLEIDO$("PRESREPA", U&)
            If CVI(Mid$(XX$, 121, 2)) > 0 Then
              JJ& = JJ& + 1
              Call GRAREG("PREPAEMI", XX$, JJ&)
            End If
          Next U&
          Call SETULTREG("PREPAEMI", JJ&)
          Call CIERRARCH("PREPAEMI")
          Screen.MousePointer = 1
          ASEL$ = "PREPAEMI"
          TSEL$ = "Indice General Presupuestos de Reparación"
          TIDOCX$ = TRIM$("Presupuesto de Reparación")
        ElseIf Option3.Value = True Then
          '
          If DERACCE%("VEPREPAR", "Visualizar Presupuestos de Reparación") < 2 Then
            Exit Sub
          End If
          '
          Screen.MousePointer = 11
          Call COPYSTRU("PRESREPA", "PREPAEMI")
          JJ& = 0
          For U& = 1 To ULTREG&("PRESREPA")
            XX$ = REGLEIDO$("PRESREPA", U&)
            If CVI(Mid$(XX$, 121, 2)) > 0 Then
              JJ& = JJ& + 1
              Call GRAREG("PREPAEMI", XX$, JJ&)
            End If
          Next U&
          Call SETULTREG("PREPAEMI", JJ&)
          Call CIERRARCH("PREPAEMI")
          Screen.MousePointer = 1
          ASEL$ = "PREPAEMI"
          TSEL$ = "Indice General Presupuestos de Reparación"
          TIDOCX$ = TRIM$("Presup.Reparación (Externo)")
        ElseIf Option4.Value = True Then
          Screen.MousePointer = 11
          Call COPYSTRU("PRESREPA", "OREPARPE")
          JJ& = 0
          For U& = 1 To ULTREG&("PRESREPA")
            XX$ = REGLEIDO$("PRESREPA", U&)
            If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
              JJ& = JJ& + 1
              Call GRAREG("OREPARPE", XX$, JJ&)
            End If
          Next U&
          Call SETULTREG("OREPARPE", JJ&)
          Call CIERRARCH("OREPARPE")
          Screen.MousePointer = 1
          ASEL$ = "OREPARPE"
          TSEL$ = "Indice General Ordenes de Reparación"
          TIDOCX$ = TRIM$("Orden de Reparación")
    End If
    '
4   Call FRESHECHO(ASEL$)
    Call SELECHO(ASEL$ + "/" + TSEL$)
    NPX& = Val(VALACT1$(ASEL$))
    If NPX& < 1 Then GoTo 99
    '
    Screen.MousePointer = 11
    NUDOCU$ = TRIM$(Str$(NPX&))
    TIDOCU$ = TIDOCX$ + " " + NUDOCU$ + " "
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
      End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a esta Reparación.")
    On Error Resume Next
    GoTo 99
    '
    '27/03/07 (OT-07-0113)
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
    '27/03/07 (FIN)
'5   For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    GoTo 99
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    FIN& = Len(TPSP$)
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7170
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7170
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'10  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 10
'      '
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'       EDITFORM.Picture1(KKI%).Visible = False
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
'    EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
    '
    GoTo 3
    '
99  End Sub

Sub PLAPREPA()
   '
10 Call SELECHO("DEUDOR1")
   NC% = Val(VALACT1$("DEUDOR1"))
   If NC% > 0 Then
     If RASOCLI$(NC%) <> "" Then
       PLAPRERE04.Label11 = TRIM$(Str$(NC%))
       PLAPRERE04.Label18 = TRIM$(Str$(NUPRESUP&))
       PLAPRERE04.Show 1
       GoTo 10
     End If
   End If
   '
End Sub

Sub PLAPREPB()
   Static ARGU$
   Command20.Enabled = False
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGU$)))
   If ARGU$ <> "" Then
     Screen.MousePointer = 11
     JJ& = 0
     Call COPYSTRU("DEUDOR1", "SELECLI")
     '
     Call APERBASDAT("CLIEN")
     On Error Resume Next
     With Deudor12
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           For KKI% = 1 To .Fields.Count
             On Error Resume Next
             VATEX$ = ""
             VATEX$ = .Fields(KKI% - 1).Value
             On Error GoTo 0
             If InStr(VATEX$, ARGU$) > 0 Then
               NCLIE% = !Nume_Cli
               RACLI$ = !Raso_Cli
               XX$ = MKI$(NCLIE%) + RACLI$
               JJ& = JJ& + 1
               Call GRAREG("!SELECLI", XX$, JJ&)
               Exit For
             End If
           Next KKI%
         .MoveNext
         Loop
       End If
     End With
     Call SETULTREG("!SELECLI", JJ&)
     '
     Screen.MousePointer = 1
     If JJ& < 1 Then
         Call MENSERR(24, "No se Encontraron Coincidencias")
         GoTo 10
       ElseIf JJ& = 1 Then
         PLAPRERE04.Label11 = TRIM$(Str$(CVI(Left$(REGLEIDO$("!SELECLI", 1), 2))))
         PLAPRERE04.Label18 = TRIM$(Str$(NUPRESUP&))
         PLAPRERE04.Show 1
         GoTo 10
       Else
         Call FRESHECHO("!SELECLI")
         Call SELECHO("!SELECLI")
         NCLIX$ = VALACT1$("!SELECLI")
         If NCLIX$ <> "" Then
           PLAPRERE04.Label11 = TRIM$(NCLIX$)
           PLAPRERE04.Label18 = TRIM$(Str$(NUPRESUP&))
           PLAPRERE04.Show 1
           GoTo 10
         End If
     End If
   End If
   '
End Sub

Sub PRESUREPA()
   '
   If DERACCE%("PRESREPA", "Generar Presupuestos Reparaciones") < 2 Then
     Exit Sub
   End If
   '
3  Call FRESHECHO("PRESREPA")
   PRESURE04.Label18 = ""
   Call SELECHO("PRESREPA/Indice General Planillas de Reparación")
   NPX& = Val(VALACT1$("PRESREPA"))
   If NPX& < 1 Then Exit Sub
   '
   For UI& = ULTREG&("PRESREPA") To 1 Step -1
     XX$ = REGLEIDO$("PRESREPA", UI&)
     If CVS(Left$(XX$, 4)) = NPX& Then GoTo 10
   Next UI&
   Call COMUNI("Error de Indices - Consulte.")
   Exit Sub
   '
10 REPRESUP = TRIM$(Str$(UI&))
   XX$ = REGLEIDO$("PRESREPA", Val(REPRESUP))
   '
   If CVS(Mid$(XX$, 125, 4)) > 0.5 Then
     Call COMUNI("Imposible Modificar este Presupuesto.\Corresponde O.Venta " + TRIM$(Str$(CVS(Mid$(XX$, 125, 4)))))
     Exit Sub
   End If
   '
   PRESURE04.Label18 = TRIM$(Str$(NPX&))
   PRESURE04.Show 1
   GoTo 3
   '
End Sub

Sub OREPARA()
   '
   If DERACCE%("OREPARA", "Generar / Cerrar Orden de Reparación") < 2 Then
     Exit Sub
   End If
   '
3  Screen.MousePointer = 11
   Call COPYSTRU("PRESREPA", "PREPAEMI")
   JJ& = 0
   For U& = 1 To ULTREG&("PRESREPA")
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVI(Mid$(XX$, 121, 2)) > 0 Then
       If CVS(Mid$(XX$, 125, 4)) < 0.5 Then
         JJ& = JJ& + 1
         Call GRAREG("PREPAEMI", XX$, JJ&)
       End If
     End If
   Next U&
   Call SETULTREG("PREPAEMI", JJ&)
   Call CIERRARCH("PREPAEMI")
   Screen.MousePointer = 1
   '
   If JJ& < 1 Then
     Call COMUNI("No Hay Presupuestos Pendientes\para Emitir Orden de Reparación.")
     Exit Sub
   End If
   '
   Call FRESHECHO("PREPAEMI")
   Call SELECHO("PREPAEMI/Indice Presupuestos de Reparación Pendientes")
   NPX& = Val(VALACT1$("PREPAEMI"))
   If NPX& < 1 Then Exit Sub
   '
   For UI& = ULTREG&("PRESREPA") To 1 Step -1
     XX$ = REGLEIDO$("PRESREPA", UI&)
     If CVS(Left$(XX$, 4)) = NPX& Then GoTo 10
   Next UI&
   Call COMUNI("Error de Indices - Consulte.")
   Exit Sub
   '
10 REPRESUP = TRIM$(Str$(UI&))
   XX$ = REGLEIDO$("PRESREPA", Val(REPRESUP))
   If CVI(Mid$(XX$, 121, 2)) < 1 Then
      Call COMUNI("Falta Generar y Emitir Presupuesto.")
      PRESURE04.Label18 = TRIM$(Str$(NPX&))
      PRESURE04.Show 1
      GoTo 3
   End If
   '
   If CVS(Mid$(XX$, 125, 4)) > 0.5 Then
      Call COMUNI("Orden de Reparación ya Emitida.")
      GoTo 3
   End If
   '
   ORDEREPA04.Label18 = ""
   ORDEREPA04.Label18 = TRIM$(Str$(NPX&))
   ORDEREPA04.Show 1
   GoTo 3
   '
End Sub

Sub CIERROREPA()
   '
   If DERACCE%("OREPARA", "Generar / Cerrar Orden de Reparación") < 2 Then
     Exit Sub
   End If
   '
3  Screen.MousePointer = 11
   Call COPYSTRU("PRESREPA", "OREPARPE")
   JJ& = 0
   For U& = 1 To ULTREG&("PRESREPA")
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
       If CVI(Mid$(XX$, 123, 2)) < 1 Then
         JJ& = JJ& + 1
         Call GRAREG("OREPARPE", XX$, JJ&)
       End If
     End If
   Next U&
   Call SETULTREG("OREPARPE", JJ&)
   Call CIERRARCH("OREPARPE")
   Screen.MousePointer = 1
   '
   If JJ& < 1 Then
     Call COMUNI("No hay Ordenes de Reparación Pendientes\que Puedan Cerrarse.")
     Exit Sub
   End If
   '
   Call SELECHO("OREPARPE/Indice de Ordenes de Reparación en Curso")
   NPX& = Val(VALACT1$("OREPARPE"))
   If NPX& < 1 Then Exit Sub
   '
   For UI& = ULTREG&("PRESREPA") To 1 Step -1
     XX$ = REGLEIDO$("PRESREPA", UI&)
     If CVS(Left$(XX$, 4)) = NPX& Then GoTo 10
   Next UI&
   Call COMUNI("Error de Indices - Consulte.")
   Exit Sub
   '
10 REPRE& = Str$(UI&)
   XX$ = REGLEIDO$("PRESREPA", REPRE&)
   If CVI(Mid$(XX$, 123, 2)) > 0 Then
      Call COMUNI("Orden de Reparación ya Cerrada")
      GoTo 3
   End If
   '
   If CVS(Mid$(XX$, 125, 4)) < 0.5 Then
      Call COMUNI("Falta Generar Orden de Reparación")
      '\\\OT-08-0868-OD-22/09/08///
      ORDEREPA04.Label18 = TRIM$(Str$(NPX&))
      ORDEREPA04.Show 1
      '\\\OT-08-0868-OD-FIN///
      GoTo 3
   End If
   '
   NCLIE% = CVI(Mid$(XX$, 67, 2))
   TALOX$ = TRIM$(Mid$(XX$, 95, 6))
   NROPE& = CVS(Mid$(XX$, 125, 4))
   VDEF$ = XX$ + String$(128, 0)
   HOII% = HOY(HOS$)
   Call REPLA(VDEF$, MKI$(HOII%), 123, 2)
   DEPORIG% = Val(CONTROL$("", "DEPOCEN"))
   Call REPLA(VDEF$, Chr$(DEPORIG%), 129, 1)
20 OBX$ = OBJPLA$("CIERROREPA", VDEF$)
   If OBX$ = "" Then
     Exit Sub
   End If
   '
   DP2% = Asc(Mid$(OBX$, 129, 1))
   If DP2% < 1 Or TRIM$(ECOARCH$("ECODEP", Chr$(DP2%))) = "" Then
      Call MENSERR(24, "Falta Depósito de Consumo")
      GoTo 20
   End If
   FECIE% = CVI(Mid$(OBX$, 123, 2))
   '
   Screen.MousePointer = 11
   Call SETULTREG("!MATEROF", 0)
   NPI% = NPX& Mod 10000
   RFF$ = RECFILT$("REPUREPA", 2, MKI$(NPI%))
   '
   URE& = ULTREG&("REPUREPA")
30 For KK& = 1 To Len(RFF$) Step 4
     REG& = CVS(Mid$(RFF$, KK&, 4))
     If REG& > 0 Then
       If REG& <= URE& Then
         XX$ = REGLEIDO$("REPUREPA", REG&)
         If CVS(Left$(XX$, 4)) = NPX& Then
           CII% = CVI(Mid$(XX$, 7, 2))
           CANTI = CVS(Mid$(XX$, 9, 4))
           '
           ZZ$ = REGBLAN$("MATEROF")
           Call REPLA(ZZ$, MKI$(CII%), 1, 2)
           Call REPLA(ZZ$, UNIMED$(CII%), 3, 2)
           Call REPLA(ZZ$, MKS$(CANTI), 5, 4)
           Call REPLA(ZZ$, MKD$(CDbl(CANTI)), 9, 8)
           Call REPLA(ZZ$, MKS$(REG&), 17, 4)
           STDIS = STOCKDIS(CII%)
           Call REPLA(ZZ$, MKS$(STDIS), 21, 4)
           FALTA = CANTI - STDIS: If FALTA < 0 Then FALTA = 0
           Call REPLA(ZZ$, MKS$(FALTA), 25, 4)
           Call REGAPP("!MATEROF", ZZ$)
           '
         End If
       End If
     End If
   Next KK&
   Call CIERRARCH("!MATEROF")
   Screen.MousePointer = 1
   '
   TITUPA$ = "Consumo de Repuestos O.Rep." + TRIM$(Str$(NPX&))
   CIXI% = CVI(Mid$(OBX$, 69, 2))
   CODPRO$ = CODEXT$(CIXI%)
   TITUPA$ = TITUPA$ + " - " + TRIM$(CODPRO$) + " - " + TRIM$(DESCRIT$(CIXI%)) + " - Cant.=" + Str$(Int(CACONJU# + 0.5))
   If NCLIE% > 0 Then TITUPA$ = TITUPA$ + " - " + RASOCLI$(NCLIE%)
   '
   VTB% = VENTABU%("MATEROF/TITUPAN=" + TITUPA$)
   If VTB% < 0 Then
      Exit Sub
   End If
   '
   Call CIERRARCH("MOVISTO")
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("REPUREPA")
   Call BLOQARCH("PRESREPA")
   '
   XX$ = REGLEIDO$("PRESREPA", REPRE&)
   Call REPLA(XX$, MKI$(FECIE%), 123, 2)
   Call GRAREG("PRESREPA", XX$, REPRE&)
   Call CIERRARCH("PRESREPA")
   '
   DOPRI$ = "OR-" + TRIM$(Str$(NPX&))
   DOSEC$ = "TL-" + TRIM$(TALOX$)
   REFE$ = "Reparación " + CODPRO$
   '
   For U& = 1 To ULTREG&("!MATEROF")
     ZZ$ = REGLEIDO$("!MATEROF", U&)
     '
     CI2% = CVI(Mid$(ZZ$, 1, 2))
     CAN = (-1) * CVS(Mid$(ZZ$, 5, 4))
     '
     USOUNI# = CAN
     RERES& = CVS(Mid$(ZZ$, 17, 4))
     If CI2% > 0 Then
       If CI2% <= NUMAS% Then
         If Abs(CAN) >= 0.05 Then
           LOTE$ = ""
           CMOV$ = "RP"
           '
           If Abs(CAN) >= 0.05 Then
              DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
              'Call MOVISTOK(FECIE%, CI2%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPX%, CAN)
              Call MOVISTOK(FECIE%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 1, NC%, DEPX%, CAN)
           End If
           REGMOVI& = ULTREG&("MOVISTO")
           '
           If RERES& > 0 Then
             If RERES& <= ULTREG&("REPUREPA") Then
               YY$ = REGLEIDO$("REPUREPA", RERES&)
               CARESE = CVS(Mid$(YY$, 21, 4))
               Call REPLA(YY$, MKS$((-1) * CAN), 25, 4)
               Call GRAREG("REPUREPA", YY$, RERES&)
               GoTo 99
             End If
           End If
           '
           YY$ = REGBLAN$("REPUREPA")
           Call REPLA(YY$, MKS$(NPX&), 1, 4)
             NPI% = NPX& Mod 10000
           Call REPLA(YY$, MKI$(NPI%), 5, 2)
           Call REPLA(YY$, MKI$(CI2%), 7, 2)
           Call REPLA(YY$, MKS$((-1) * CAN), 25, 4)
           Call REPLA(YY$, "OR-" + TRIM$(Str$(NPX&)), 33, 12)
           Call REGAPP("REPUREPA", YY$)
         End If
       End If
     End If
99 Next U&
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("REPUREPA")
   '
   ' CANCELAR TODAS LAS RESERVAS
   Call ABRESERVA
   SQLX$ = "SELECT * FROM RESERVA"
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & DOPRI$ & "' "
   Set RESELE = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   With RESELE
101   On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error Resume Next
         .Delete
         If Err Then GoTo 102
         GoTo 101
      End If
      On Error GoTo 0
   End With
102 RESELE.Close
   '
   HOII% = HOY(HOS$)
   TTXX$ = Space$(96)
   Call REPLA(TTXX$, HOS$ + "  " + "Cierre Orden de Reparación " + TRIM$(Str$(NPX&)), 1, 56)
   Call REPLA(TTXX$, USERTER$, 69, 24)
   Call REGAPP("ANOTAREP", Left$(MKS$(NPX&) + TTXX$, 96))
   Call REGAPP("ANOTAREP", MKS$(NPX&) + String$(88, "="))
   Call CIERRARCH("ANOTAREP")
   '
   TTXX$ = Space$(96)
   Call REPLA(TTXX$, HOS$ + "  " + "Cierre O/Rep " + TRIM$(Str$(NPX&)) + " - Talón " + TRIM$(TALOX$), 1, 56)
   Call REPLA(TTXX$, USERTER$, 69, 24)
   Call REGAPP("ANOTAPED", Left$(MKS$(NROPE&) + TTXX$, 96))
   Call REGAPP("ANOTAPED", MKS$(NROPE&) + String$(88, "="))
   Call CIERRARCH("ANOTAPED")
   '
   RFF$ = RECFILT$("PEDDETA", 1, MKS$(NROPE&))
   For U& = 1 To Len(RFF$) Step 4
     REPEDD& = CVS(Mid$(RFF$, U&, 4))
     PDC$ = REGLEIDO$("PEDDETA", REPEDD&)
     If CVI(Mid$(PDC$, 33, 2)) = CIXI% Then
       If CVI(Mid$(PDC$, 7, 2)) = NCLIE% Then
         CANTID = CVS(Mid$(PDC$, 79, 4))
         Call REPLA(PDC$, Chr$(1), 27, 1)
         Call REPLA(PDC$, String$(38, 0), 71, 38)
         Call REPLA(PDC$, MKS$(CANTID), 75, 4)
         Call REPLA(PDC$, MKS$(CANTID), 99, 4)
         Call REGAPP("DESPAPED", PDC$)
         GoTo 90
       End If
     End If
   Next U&
   '
   Call MENSERR(24, "No se Encontró la Nota de Venta\Correspondiente a esta Reparación.\  \Consulte a su Soporte de Sistemas.")
   '
90 Call CIERRARCH("DESPAPED")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   GoTo 3
   '
End Sub
'
Function PRONUOCO&()
   '
   PRONU& = 1
   REGXX& = ULTREG&("OCOMENCA")
   If REGXX& > 0 Then
     PRONU& = 1 + CVS(Left$(REGLEIDO$("OCOMENCA", REGXX&), 4))
   End If
   If PRONU& <= REGXX& Then
     Screen.MousePointer = 11
     For U& = 1 To REGXX&
       X0$ = REGLEIDO$("OCOMENCA", U&)
       If CVS(Left$(X0$, 4)) >= PRONU& Then
         PRONU& = 1 + CVS(Left$(X0$, 4))
       End If
     Next U&
     Screen.MousePointer = 1
   End If
   '
   PRONUOCO& = PRONU&
   '
End Function

Function PRONUOTER&()    ' numero orden de servicio terceros
   '
   PRONU& = 1
   REGXX& = ULTREG&("OPERTER")
   If REGXX& > 0 Then
     PRONU& = 1 + CVS(Mid$(REGLEIDO$("OPERTER", REGXX&), 75, 4))
   End If
   If PRONU& <= REGXX& Then
     Screen.MousePointer = 11
     For U& = 1 To REGXX&
       X0$ = REGLEIDO$("OPERTER", U&)
       If CVS(Mid$(X0$, 75, 4)) >= PRONU& Then
         PRONU& = 1 + CVS(Mid$(X0$, 75, 4))
       End If
     Next U&
     Screen.MousePointer = 1
   End If
   '
   PRONUOTER& = PRONU&
   '
End Function



Private Sub Command8_Click()
   Command8.Enabled = False
   ANOTAREP04.Show 1
   Command8.Enabled = True
End Sub


Private Sub Exit_Click()
   Call Command1_Click
End Sub

Private Sub Form_Load()
    '
    Show
    Refresh
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Close: End
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("PRESREPA")
    For UI& = 1 To FIN&
      Call TRACE(20, UI&, FIN&)
      XX$ = REGLEIDO$("PRESREPA", UI&)
      If CVS(Mid$(XX$, 125, 4)) < 0 Then
        If DERACCE%("OREPARA", "Generar / Cerrar Orden de Reparación") < 2 Then
          Call CIERRARCH("*.*")
          Call BAJALDISCO
          Call FINAL("")
        End If
        '
        NPX& = CVS(Left$(XX$, 4))
        '\\\OT-08-0868-OD-22/09/08///
        ORDEREPA04.Label18 = TRIM$(Str$(NPX&))
        ORDEREPA04.Text2 = TRIM$(Str$(Abs(CVS(Mid$(XX$, 125, 4)))))
        ORDEREPA04.Text2.Enabled = False
        Screen.MousePointer = 1
        ORDEREPA04.Show 1
        ORDEREPA04.Text2.Enabled = True
        '\\\OT-08-0868-OD-FIN///
        Screen.MousePointer = 11
      End If
   Next UI&
   Screen.MousePointer = 1
   '
End Sub

Function NUPRESUP&()
   '
   NPP& = 1
   '
   FIN& = ULTREG&("PRESREPA")
   If FIN& > 0 Then
      For U& = FIN& To FIN& - 128 Step -1
         If U& < 1 Then GoTo 99
         XX$ = REGLEIDO$("PRESREPA", U&)
         NPPX& = CVS(Mid$(XX$, 1, 4))
         If NPPX& >= NPP& Then
            NPP& = 1 + NPPX&
         End If
      Next U&
   End If
   '
99 NUPRESUP& = NPP&
   '
End Function
'

Private Sub Genpresu_Click()
   Call Command3_Click
End Sub

Private Sub Lispre_Click()
   Call CONFREPA04.Command5_Click
End Sub

Private Sub Master_Click()
   Call PLAPRERE04.Command8_Click
End Sub

Private Sub Mater_Orep_Click()
   Option5.Value = True
   Call Command7_Click
End Sub

Private Sub Oper_Orep_Click()
   Option6.Value = True
   Call Command7_Click
End Sub

Private Sub Option1_DblClick()
   Call PRESEDOC
End Sub

Private Sub Option2_DBLClick()
   Call PRESEDOC
End Sub

Private Sub Option3_DblClick()
   Call PRESEDOC
End Sub

Private Sub Option4_DblClick()
   Call PRESEDOC
End Sub





Private Sub Option9_DblClick()
'
3  Screen.MousePointer = 11
   Call COPYSTRU("PRESREPA", "OREPARPE")
   JJ& = 0
   For U& = 1 To ULTREG&("PRESREPA")
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
       If CVI(Mid$(XX$, 123, 2)) > 1 Then
         JJ& = JJ& + 1
         Call GRAREG("OREPARPE", XX$, JJ&)
       End If
     End If
   Next U&
   Call SETULTREG("OREPARPE", JJ&)
   Call CIERRARCH("OREPARPE")
   Screen.MousePointer = 1
   '
   If JJ& < 1 Then
     Call COMUNI("No hay Ordenes de Reparación Pendientes\Para Despachar.")
     Exit Sub
   End If
  '
  Call SELECHO("OREPARPE/Ordenes de Reparación Pendientes de Despacho")
  NPX& = Val(VALACT1$("OREPARPE"))
  If NPX& < 1 Then Exit Sub
   '
   XX$ = FILTERGETRECORD$("PRESREPA", 1, MKS(NPX&))
   '
   NUCLIE% = CVI(Mid$(XX$, 67, 2))
   codrep% = CVI(Mid$(XX$, 69, 2))
   NUSERI$ = Mid$(XX$, 71, 18)
   '
   '
   '
PRIDOCREMI:
    '
    PRINTERPORT$ = CONTROL$(IDENTER$, "PORTREM")
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    If TRIM$(FORIPRE$) <> "" Then
      '
      PREFIREM$ = CONTROL$("", "PREFIREM")
      If PREFIREM$ = "" Then PREFIREM$ = "0001"
      Call NUREMIFA(NUPRORE&, PREFIREM$)
      NROREMIT$ = NUPRORE&
      While Len(NROREMIT$) < 8: NROREMIT$ = "0" + NROREMIT$: Wend
      NROREMIT$ = PREFIREM$ + "-" + NROREMIT$
      '
      FECHDOC$ = FECHATEX(HOY(HO$))
      NUMEDOC$ = NROREMIT$
      TIPODOC$ = "Remito"
      '
      DESTIDOC% = NUCLIE%
      DESCMAT$ = "Reemitimos el Equipo " & CODEXT$(codrep%) & " - " & DESCRIT0$(codrep%)
      '
      CODTABU1$(1) = "DES-MAT"
      CODTABU1$(2) = "CANTIDAD"

      CACOLTAB1% = 2
      '
      CAITAB1% = 1
      TETABU1$(1, CAITAB1%) = DESCMAT$
      TETABU1$(2, CAITAB1%) = "1"
      CAITAB1% = CAITAB1% + 1
      TETABU1$(1, CAITAB1%) = "Con el Nº de Serie: " & NUSERI$
      '
      FF% = HOY(HO$)
      NDOCOR$ = "RT." & TRIM$(XVALO(PREFIREM$)) & "-" & NUPRORE&
      NDOLAR$ = "RT-X-" & NUMEDOC$
      REFEMO$ = "Reparacion " & CODEXT(codrep%)
      Call MOVISTOK(FF%, 0, "", "RT", NDOLAR$, NDOCOR$, "", "", REFEMO$, 1, 0, 1, NUCLIE%, 1, 1)
      Call PRINTDOC("FOREMIFA")   ' Remito de Facturacion
      '
    End If
    '

End Sub

Private Sub Orderep_Click()
   Option4.Value = True
   Call Command7_Click
End Sub

Private Sub Ordrep_Click()
   Call Command4_Click
End Sub

Private Sub Plabuscli_Click()
   Call Command20_Click
End Sub

Private Sub Pladir_Click()
   Call Command5_Click
End Sub

Private Sub Plarepa_Click()
   Option1.Value = True
   Call Command7_Click
End Sub

Private Sub Presucli_Click()
   Option3.Value = True
   Call Command7_Click
End Sub

Private Sub Presuint_Click()
   Option2.Value = True
   Call Command7_Click
End Sub

Private Sub Setup_Click()
   Call Command6_Click
End Sub

Private Sub ESTAREP()
  '
  Call DESHASFECHA(DES%, HAS%, PERIO$)
  '
  If TRIM$(PERIO$) = "" Then Exit Sub 'SI CANCELO
  '
  Screen.MousePointer = 11
  FIN& = ULTREG&("PRESREPA")
  JJ& = 0
  For U& = 1 To FIN&
    Call TRACE(20, U&, FIN&)
    XX$ = REGLEIDO$("PRESREPA", U&)
    FEIN1% = CVI(Mid$(XX$, 65, 2))   ' FECHA DE RECEPCION
    '
    If FEIN1% >= DES% Then
      If FEIN1% <= HAS% Then
        JJ& = JJ& + 1
        NO1% = JJ&
        NUREP1& = CVS(Mid$(XX$, 1, 4))
        CLIE1% = CVI(Mid$(XX$, 67, 2))
        CI1% = CVI(Mid$(XX$, 69, 2))
        FEPRE1% = CVI(Mid$(XX$, 121, 2))   ' FECHA PRESUPUESTO
        FECIE1% = CVI(Mid$(XX$, 123, 2))   ' FECHA DE CIERRE
        NUOVEN1 = CVS(Mid$(XX$, 125, 4))
        '
        'OBTENGO LA FECHA DE LA O/VEN
        XZ1$ = FILTERGETRECORD$("PEDENCA", 1, MKS$(NUOVEN1))
        FEOREP1% = CVI(Mid$(XZ1$, 53, 2))  ' FECHA DE APROBACION
        '
        YY$ = REGBLAN$("ESTAREP")
        '
        Call REPLA(YY$, MKI$(NO1%), 1, 2)
        Call REPLA(YY$, MKS$(NUREP1&), 3, 4)
        Call REPLA(YY$, MKI$(CLIE1%), 7, 2)
        Call REPLA(YY$, MKI$(CI1%), 9, 2)
        Call REPLA(YY$, MKI$(FEIN1%), 11, 2)
        Call REPLA(YY$, MKI$(FEPRE1%), 13, 2)
        Call REPLA(YY$, MKI$(FEOREP1%), 15, 2)
        Call REPLA(YY$, MKS$(NUOVEN1), 17, 4)
        Call REPLA(YY$, MKI$(FECIE1%), 21, 2)
        '
        Call GRAREG("ESTAREP", YY$, JJ&)
      End If
    End If
  Next U&
  '
  Call SETULTREG("ESTAREP", JJ&)
  Call CIERRARCH("ESTAREP")
  Screen.MousePointer = 1
  '
  If JJ& > 0 Then
    Call HEADERS("ESTAREP", "")
    Call HEADERS("ESTAREP", "Período: " & PERIO$)
    Call FINAL("*ESTAREP")
  End If
  '
End Sub
Private Sub OREPABI()
   '
   FIN& = ULTREG&("PRESREPA")
   For U& = 1 To FIN&
     Call TRACE(20, UI&, FIN&)
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
        If CVI(Mid$(XX$, 123, 2)) < 1 Then
           JJ& = JJ& + 1
           Call GRAREG("OREPARPE", XX$, JJ&)
        End If
     End If
   Next U&
   Call SETULTREG("OREPARPE", JJ&)
   Screen.MousePointer = 1
   '
   Call FINAL("*LIREPAPE")
   '
End Sub
