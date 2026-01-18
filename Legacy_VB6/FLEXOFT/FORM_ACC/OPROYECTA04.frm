VERSION 5.00
Begin VB.Form OPROYECTA04 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operaciones Proyectadas por Orden de Fabricación"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture16 
      Height          =   5865
      Left            =   9400
      ScaleHeight     =   5805
      ScaleWidth      =   1035
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   5100
         Left            =   0
         Picture         =   "OPROYECTA04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos de la O/F Activa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3585
      Left            =   105
      TabIndex        =   12
      Top             =   1260
      Width           =   9150
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
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
         Height          =   3300
         Left            =   50
         TabIndex        =   13
         Top             =   210
         Width           =   9045
         Begin VB.Frame Frame1 
            Caption         =   "Frame7"
            Height          =   1850
            Index           =   5
            Left            =   8520
            TabIndex        =   19
            Top             =   1400
            Width           =   15
         End
         Begin VB.Frame Frame1 
            Caption         =   "Frame6"
            Height          =   1850
            Index           =   4
            Left            =   7660
            TabIndex        =   18
            Top             =   1400
            Width           =   15
         End
         Begin VB.Frame Frame1 
            Caption         =   "Frame5"
            Height          =   1850
            Index           =   3
            Left            =   6550
            TabIndex        =   17
            Top             =   1400
            Width           =   15
         End
         Begin VB.Frame Frame1 
            Caption         =   "Frame4"
            Height          =   1850
            Index           =   2
            Left            =   5300
            TabIndex        =   16
            Top             =   1400
            Width           =   15
         End
         Begin VB.Frame Frame1 
            Caption         =   "Frame3"
            Height          =   1850
            Index           =   1
            Left            =   4500
            TabIndex        =   15
            Top             =   1400
            Width           =   15
         End
         Begin VB.Frame Frame1 
            Caption         =   "Frame1"
            Height          =   1850
            Index           =   0
            Left            =   630
            TabIndex        =   14
            Top             =   1400
            Width           =   15
         End
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
            Height          =   1635
            Left            =   105
            Sorted          =   -1  'True
            TabIndex        =   20
            Top             =   1605
            Width           =   8835
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Left            =   105
            ScaleHeight     =   270
            ScaleWidth      =   8775
            TabIndex        =   21
            Top             =   1365
            Width           =   8835
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   " N-O Operación                            Maq.         Real.     Saldo    Horas  "
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
               TabIndex        =   22
               Top             =   0
               Width           =   9045
            End
         End
         Begin VB.Label Label18 
            BackColor       =   &H00E0E0E0&
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
            Left            =   7980
            TabIndex        =   39
            Top             =   525
            Width           =   960
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Calculada:"
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
            Left            =   6405
            TabIndex        =   38
            Top             =   525
            Width           =   1485
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
            TabIndex        =   37
            Top             =   250
            Width           =   960
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   36
            Top             =   210
            Width           =   1170
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
            TabIndex        =   35
            Top             =   250
            Width           =   960
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   34
            Top             =   210
            Width           =   330
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Dias Atraso:"
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
            Left            =   7140
            TabIndex        =   33
            Top             =   885
            Width           =   960
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   465
            Left            =   8190
            TabIndex        =   32
            Top             =   840
            Width           =   750
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
            Left            =   6405
            TabIndex        =   31
            Top             =   250
            Width           =   1485
         End
         Begin VB.Label Label8 
            BackColor       =   &H00E0E0E0&
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
            Left            =   7980
            TabIndex        =   30
            Top             =   210
            Width           =   960
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
            TabIndex        =   29
            Top             =   525
            Width           =   960
         End
         Begin VB.Label Label10 
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   28
            Top             =   525
            Width           =   2010
         End
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
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
            Height          =   465
            Left            =   1155
            TabIndex        =   27
            Top             =   840
            Width           =   5160
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad:"
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
            TabIndex        =   26
            Top             =   255
            Width           =   1485
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   25
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo:"
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
            TabIndex        =   24
            Top             =   525
            Width           =   1485
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   23
            Top             =   525
            Width           =   1065
         End
      End
   End
   Begin VB.CommandButton Command8 
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
      Height          =   285
      Left            =   2940
      TabIndex        =   1
      Top             =   315
      Width           =   175
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   945
      TabIndex        =   0
      Top             =   315
      Width           =   2010
   End
   Begin VB.CommandButton Command3 
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
      Height          =   285
      Left            =   2940
      TabIndex        =   7
      Top             =   735
      Width           =   175
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00004000&
      Height          =   8460
      Left            =   9400
      ScaleHeight     =   8400
      ScaleWidth      =   1320
      TabIndex        =   4
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command2 
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
         Height          =   640
         Left            =   105
         Picture         =   "OPROYECTA04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Borrar Operación Activa"
         Top             =   4170
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   5
         Top             =   1785
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   12000
         End
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
         Height          =   640
         Left            =   105
         Picture         =   "OPROYECTA04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "OPROYECTA04.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   845
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "OPROYECTA04.frx":2AF6
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   2190
         Width           =   1515
      End
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
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
      Left            =   3360
      TabIndex        =   40
      Top             =   390
      Width           =   1335
   End
   Begin VB.Label Label117 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O/F Nro:"
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
      Left            =   -420
      TabIndex        =   11
      Top             =   810
      Width           =   1335
   End
   Begin VB.Label Label166 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   -420
      TabIndex        =   10
      Top             =   390
      Width           =   1335
   End
   Begin VB.Label Label144 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   735
      Width           =   1980
   End
   Begin VB.Label Label133 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   4725
      TabIndex        =   8
      Top             =   315
      Width           =   4500
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu borraoperacionactiva 
         Caption         =   "Borrar Operación Activa"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "OPROYECTA04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub borraoperacionactiva_Click()
    Call Command2_Click
End Sub

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub


Private Sub Command2_Click()
   Command2.Enabled = False
   For KKU% = 1 To List2.ListCount
     If List2.Selected(KKU% - 1) = True Then
       NOPESEL% = XVALO(Left$(List2.List(KKU% - 1), 4))
       GoTo 10
     End If
   Next KKU%
   Call MENSERR(24, "Ninguna Operación Seleccionada")
   Exit Sub
   '
10 If DERACCE%("BORROPER", "Eliminar Operación Proyectada") >= 2 Then
     TTXX$ = "¿ Realmente Desea Eliminar Definitivamente\la Operación de Fabricación Seleccionada ?\   \Esta Transacción no Podrá Revertirse !!!"
     If COMALTER%(TTXX$ + "\\No, Conservarla\Si, Eliminarla") = 2 Then
        '
        SEBORRO% = 0
        NORFA$ = Label2
        With OperFab
20        .FindFirst "IdSubOr = '" & NORFA$ & "' AND NumeOper = " & NOPESEL%
          If .NoMatch = False Then
            .Delete
            SEBORRO% = 1
            GoTo 20
          End If
        End With
        '
        If SEBORRO% < 1 Then
            Call COMUNI("La Operación no Pudo Completarse.")
          Else
            Call COMUNI("Operación de Borrado Completada.\   \Recuerde Recalcular la Secuencia\de Carga de Máquinas.")
        End If
     End If
   End If
   Command2.Enabled = True
End Sub

Private Sub Command29_Click()
  '
  Call OPNOIN("Ayuda FLEXOFT en Línea")
  '
End Sub

Private Sub Command3_Click()
   '
   NROOF$ = ""
   If ULTREG&("!ECORDEP") < 1 Then
     Call MENSERR(24, "No Hay Ordenes de Fabricación en Curso\para el Código Seleccionado")
     GoTo 90
   End If
   '
   If CODINT%(Text10) > 0 Then
     Call SELECHO("!ECORDEP/Ordenes de Trabajo en Curso")
     NROOF$ = AJUSTI$(VALACT1$("!ECORDEP"), 12)
     If TRIM$(NROOF$) = "" Then GoTo 90
   End If
   '
90 Label144 = NROOF$
   Screen.MousePointer = 1
   '
End Sub


Private Sub Command8_Click()
   Command8.Enabled = False
   Call SELECHO("MASTER")
   If VALACT1$("MASTER") <> "" Then
     Text10 = VALACT1$("MASTER")
   End If
   Command8.Enabled = True
End Sub

Sub Form_Load()
  '
  If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      
      MARCOBARRA.Top = 4700
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
  End If
  '
  Screen.MousePointer = 11
  Call CENTRAFORM(Me)
  HOII% = HOY(HOS$)
  Call BLANARCH("!OPERINF")
  Call SETULTREG("!OPERINF", 0)
  Call FRESHECHO("!OPERINF")
  'AGREGUE15/09/06MH
  '
  Screen.MousePointer = 1
  '
End Sub

Private Sub Label144_Change()
    '
    Label2 = Label144
    '
End Sub


Private Sub Label2_Change()
   Screen.MousePointer = 11
   NORFA$ = Label2
   Label4.Caption = ""
   Label6.Caption = ""
   Label8.Caption = ""
   Label10.Caption = ""
   Label11.Caption = ""
   Label13.Caption = ""
   Label15.Caption = ""
   List2.Clear
   Command2.Enabled = False
   '
   Call ABREORFAB
   '
   SQLX$ = " SELECT * FROM Ordefabr "
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
   Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With ORFATEMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Screen.MousePointer = 1
        Exit Sub
      End If
      On Error GoTo 0
      '
      CIIX% = !Cod_Inte
      Label10.Caption = CODEXT$(CIIX%)
      Label11.Caption = DESCRIT$(CIIX%)
      Label13.Caption = FORMATNUM$(!CanProy, "F9.1")
        FESOLENT = CVINT(!FechEnSol)
      Label8.Caption = FECHATEX$(FESOLENT)
        FECALC = CVINT(!FechProyEntre)
      Label18.Caption = FECHATEX$(FECALC)
        DIATRA& = DIENTRE(FESOLENT, FECALC): If DIATRA < 1 Then DIATRA = 0
      Label6 = TRIM$(Str$(DIATRA&))
      Label4.Caption = !Statuorf
        SALDORF = !CanProy - !CanApro: If SALDORF < 0 Then SALDORF = 0
      Label15.Caption = FORMATNUM$(SALDORF, "F9.1")
   End With
   '
   SQLX$ = " SELECT * FROM OPERFAB "
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
   SQLX$ = SQLX$ + "AND StatOper < 8 "                      ' no marca cumplidos"
   SQLX$ = SQLX$ + " ORDER BY NumeOper "
   '
   Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With FABRTMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Call MENSERR(24, "No Hay Operaciones Proyectadas\para la Orden de Fabricación Elegida")
         Screen.MousePointer = 1
         Exit Sub
       End If
       '
       On Error GoTo 0
       JJ& = 0
       Do While (.EOF = False)
         TTXX$ = Space$(128)
            NOPE% = !NumeOper      'numero de operación
         Call REPLA(TTXX$, FORMATNUM$(NOPE%, "I4"), 1, 4)
            DOPE$ = !DescOper      'Descripcion de la operacion
         Call REPLA(TTXX$, DOPE$, 6, 36)
            MAQUI$ = !Maquina      'Máquina
         Call REPLA(TTXX$, MAQUI$, 43, 6)
           CAPROD# = !CantProg      'Cantidad Programada
           CAREA# = !CantReal       'Cantidad realizada
         Call REPLA(TTXX$, FORMATNUM$(CAREA#, "F10.1"), 51, 10)
           SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
         Call REPLA(TTXX$, FORMATNUM$(SALDO#, "F10.1"), 62, 10)
           MODOX% = 0: If CAREA# > 0.05 Then MODO% = 1
           TIEMPEN = TIEMAQUI(CIIX%, NOPE%, SALDO#, MODOX%, 0)
         Call REPLA(TTXX$, Left$(HORMINSE$(60 * TIEMPEN), 6), 74, 6)
AAA = BBB
'         Call REPLA(TTXX$, CSTRING$(MKD$(SALDO#), 4, 10, 1, 2), 39, 10)
'         'Call REPLA(TTXX$, CSTRING$(Mid$(XX$, 87, 8), 4, 10, 1, 2), 54, 10)
'         SALDO# = CVD(Mid$(XX$, 79, 8)) - CVD(Mid$(XX$, 87, 8))
'         If SALDO# < 0 Then SALDO# = 0
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
         List2.AddItem TTXX$
       .MoveNext
       Loop
   End With
'
'********************************
'   For U& = 1 To ULTREG&("OPERFAP")
'      XX$ = REGLEIDO$("OPERFAP", U&)
'      If TRIM$(Left$(XX$, 12)) = NORFA$ Then
'      End If
'   Next U&
'
'   Label18 = FECHATEX(FULTOP)
'   DIATRA = 0: If FULTOP > FESOLENT Then DIATRA = DIENTRE(FESOLENT, FULTOP)
'   If DIATRA > 999 Then DIATRA = 999
'   Label6 = "": If DIATRA > 0.1 Then Label6 = TRIM$(Str$(DIATRA))
'   '
End Sub

Private Sub List2_Click()
   Command2.Enabled = True
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub Text10_Change()
   Label133 = DESCRIT0$(CODINT%(Text10))
   '
   CIXXI% = CODINT%(Text10)
   Label144 = "": CANOR% = 0
   If CIXXI% > 0 Then
      Call ABREORFAB
      SQLX$ = " SELECT * FROM ORDEFABR "
      SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXXI% & " "
      SQLX$ = SQLX$ + "AND StatUORF < 8 "                      ' no marca cumplidos"
      SQLX$ = SQLX$ + "AND CanApro < CanProy "                      ' no marca cumplidos"
      SQLX$ = SQLX$ + " ORDER BY IdSubOr ASC "
      '
      Set ORFASEL = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      JJ& = 0
      On Error Resume Next
      ORFASEL.MoveLast
      If Err Then
         GoTo 90
      End If
      Screen.MousePointer = 11
      CANOR% = ORFASEL.RecordCount
      With ORFASEL
        .MoveFirst
        Label144 = ORFASEL!IdSubOr
        Do While (.EOF = False)
          Y$ = REGBLAN("ECORDEP")
          Call REPLA(Y$, !IdSubOr, 1, 12)
          Call REPLA(Y$, ECOARCH$("ECORDEF", !IdSubOr), 13, 52)
          '
          JJ& = JJ& + 1
          Call GRAREG("!ECORDEP", Y$, JJ&)
          '
        .MoveNext
        Loop
      End With
90    Call SETULTREG("!ECORDEP", JJ&)
      Call CIERRARCH("!ECORDEP")
      Call FRESHECHO("!ECORDEP")
      Screen.MousePointer = 1
   End If
   '
End Sub

Private Sub Text10_DblClick()
   Call Command8_Click
End Sub
'
