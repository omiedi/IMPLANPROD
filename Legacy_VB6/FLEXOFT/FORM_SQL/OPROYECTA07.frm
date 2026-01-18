VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPROYECTA07 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operaciones Proyectadas por Orden de Fabricación"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos de la O/F Activa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
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
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1260
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
                  Italic          =   -1  'True
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
               Name            =   "Arial"
               Size            =   9
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
               Name            =   "Arial"
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
               Name            =   "Arial"
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
               Name            =   "Arial"
               Size            =   9
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2775
            TabIndex        =   35
            Top             =   255
            Width           =   600
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
               Name            =   "Arial"
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
               Name            =   "Arial"
               Size            =   9
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
               Name            =   "Arial"
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
               Name            =   "Arial"
               Size            =   9
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
               Name            =   "Arial"
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
               Name            =   "Arial"
               Size            =   9
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
               Name            =   "Arial"
               Size            =   9
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4275
            TabIndex        =   26
            Top             =   255
            Width           =   885
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
               Name            =   "Arial"
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
               Name            =   "Arial"
               Size            =   9
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
         Size            =   9.75
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
         Size            =   9.75
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
         Picture         =   "OPROYECTA07.frx":0000
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
         Picture         =   "OPROYECTA07.frx":030A
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
         Picture         =   "OPROYECTA07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   845
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "OPROYECTA07.frx":075E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   2190
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "OPROYECTA07.frx":2680
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   2400
      OleObjectBlob   =   "OPROYECTA07.frx":28B4
      TabIndex        =   43
      Top             =   120
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "Arial"
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
         Name            =   "Arial"
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
         Name            =   "Arial"
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
End
Attribute VB_Name = "OPROYECTA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
        SQLX$ = "DELETE FROM OPERFAB WHERE IdSubOr = '" & NORFA$ & "' AND NumeOper = " & NOPESEL%
        SEBORRO% = 1
        On Error Resume Next
        FLEXCONN.Execute SQLX$
        If Err Then
           On Error GoTo 0
           SEBORRO% = 0
        End If
        '
        If SEBORRO% < 1 Then
            Call COMUNI("La Operación no Pudo Completarse.")
          Else
            Call COMUNI("Operación de Borrado Completada.\   \Recuerde Recalcular la Secuencia\de Carga de Máquinas.")
        End If
        '
        Call Command1_Click
        Exit Sub
        '
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
  UTILIZA_SKIN% = 1
  Call APLICACIONSKINS(Me)

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
   SQLX$ = " SELECT COD_INTE, CanProy, FechEnSol, FechProyEntre, StatuOrf, CanApro  "
   SQLX$ = SQLX$ + " FROM Ordefabr WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
   'Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
'   Dim ORFATEMP As ADODB.Recordset
'   Set ORFATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   Set ORFATEMP = READSET(SQLX$)
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
      CIIX% = !cod_inte
      Label10.Caption = CODEXT$(CIIX%)
      Label11.Caption = DESCRIT$(CIIX%)
      Label13.Caption = FORMATNUM$(!CANPROY, "F9.1")
        FeSolEnt = CVINT(!FechEnSol)
      Label8.Caption = FECHATEX$(FeSolEnt)
        FECALC = CVINT(!FechProyEntre)
      Label18.Caption = FECHATEX$(FECALC)
        DIATRA& = DIENTRE(FeSolEnt, FECALC): If DIATRA < 1 Then DIATRA = 0
      Label6 = TRIM$(str$(DIATRA&))
      Label4.Caption = !StatuOrf
        SALDORF = !CANPROY - !CanApro: If SALDORF < 0 Then SALDORF = 0
      Label15.Caption = FORMATNUM$(SALDORF, "F9.1")
   End With
   Set ORFATEMP = Nothing
   '
   SQLX$ = " SELECT NUMEOPER, DESCOPER, MAQUINA, CANTPROG, CANTREAL FROM OPERFAB WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
   SQLX$ = SQLX$ + " AND StatOper < 8 "                      ' no marca cumplidos"
   SQLX$ = SQLX$ + " ORDER BY NumeOper "
   '
   'Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
'   Dim FABRTMP As ADODB.Recordset
'   Set FABRTMP = New ADODB.Recordset
'   FABRTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   Set FABRTMP = READSET(SQLX$)
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
            NOPE% = !NUMEOPER      'numero de operación
         Call REPLA(TTXX$, FORMATNUM$(NOPE%, "I4"), 1, 4)
            DOPE$ = !descoper      'Descripcion de la operacion
         Call REPLA(TTXX$, DOPE$, 6, 36)
            MAQUI$ = !MAQUINA      'Máquina
         Call REPLA(TTXX$, MAQUI$, 43, 6)
           CAPROD# = !CANTPROG      'Cantidad Programada
           CAREA# = !CANTREAL       'Cantidad realizada
         Call REPLA(TTXX$, FORMATNUM$(CAREA#, "F10.1"), 51, 10)
           SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
         Call REPLA(TTXX$, FORMATNUM$(SALDO#, "F10.1"), 62, 10)
           MODOX% = 0: If CAREA# > 0.05 Then MODO% = 1
           TIEMPEN = TIEMAQUI(CIIX%, NOPE%, SALDO#, MODOX%, 0)
         Call REPLA(TTXX$, Left$(HORMINSE$(60 * TIEMPEN), 6), 74, 6)
         List2.AddItem TTXX$
       .MoveNext
       Loop
   End With
   Set FABRTMP = Nothing
   '
End Sub

Private Sub List2_Click()
   Command2.Enabled = True
End Sub

Private Sub Text10_Change()
   Label133 = DESCRIT0$(CODINT%(Text10))
   '
   CIXXI% = CODINT%(Text10)
   Label144 = "": CANOR% = 0
   If CIXXI% > 0 Then
      Call ABREORFAB
      SQLX$ = " SELECT IDSUBOR FROM ORDEFABR WITH(NOLOCK) "
      SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXXI% & " "
      SQLX$ = SQLX$ + "AND StatUORF < 8 "                      ' no marca cumplidos"
      SQLX$ = SQLX$ + "AND CanApro < CanProy "                      ' no marca cumplidos"
      SQLX$ = SQLX$ + " ORDER BY IdSubOr ASC "
      '
      'Set ORFASEL = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
'      Dim ORFASEL As ADODB.Recordset
'      Set ORFASEL = New ADODB.Recordset
'      ORFASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
      '
      Set ORFASEL = READSET(SQLX$)
      JJ& = 0
      On Error Resume Next
      ORFASEL.MoveLast
      If Err Then
         On Error GoTo 0
         GoTo 90
      End If
      Screen.MousePointer = 11
      CANOR% = ORFASEL.RecordCount
      With ORFASEL
        .MoveFirst
        Label144 = ORFASEL!idsubor
        Do While (.EOF = False)
          Y$ = REGBLAN("ECORDEP")
          Call REPLA(Y$, !idsubor, 1, 12)
          Call REPLA(Y$, ECOARCH$("ECORDEF", !idsubor), 13, 52)
          '
          JJ& = JJ& + 1
          Call GRAREG("!ECORDEP", Y$, JJ&)
          '
        .MoveNext
        Loop
      End With
      Set ORFASEL = Nothing
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
