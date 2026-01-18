VERSION 5.00
Begin VB.Form APROGMRP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programacion de Necesidades de  Materiales"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "APROGMRP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11910
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1630
      Left            =   6560
      TabIndex        =   29
      Top             =   1540
      Width           =   4215
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
         Height          =   1070
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.CommandButton Command18 
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
      Left            =   10920
      Picture         =   "APROGMRP.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Configuración de Funcionamiento"
      Top             =   2115
      Width           =   855
   End
   Begin VB.CommandButton Command17 
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
      Left            =   10920
      Picture         =   "APROGMRP.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Ayuda en Línea (si Disponible)"
      Top             =   1215
      Width           =   855
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Calc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   10920
      TabIndex        =   18
      Top             =   3300
      Width           =   855
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   555
         TabIndex        =   20
         Top             =   840
         Width           =   615
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   12000
         End
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
         Height          =   735
         Left            =   120
         Picture         =   "APROGMRP.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Re-Cálculo de Necesidad de Componentes"
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Left            =   165
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
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
      Left            =   10920
      Picture         =   "APROGMRP.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Guarda Programa de Fabricación Activo"
      Top             =   6480
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Características"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6560
      TabIndex        =   11
      Top             =   210
      Width           =   4215
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
         Left            =   3840
         Picture         =   "APROGMRP.frx":106A
         TabIndex        =   31
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   30
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         TabIndex        =   33
         Top             =   400
         Width           =   1320
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   32
         Top             =   720
         Width           =   3825
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
      Left            =   10920
      Picture         =   "APROGMRP.frx":1374
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Seleccion de Programa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   105
      TabIndex        =   1
      Top             =   1365
      Width           =   6350
      Begin VB.CommandButton Command24 
         Caption         =   "New Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   5355
         Picture         =   "APROGMRP.frx":14BE
         TabIndex        =   10
         ToolTipText     =   "Agrega Articulo al Programa"
         Top             =   1020
         Width           =   870
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
         Height          =   600
         Left            =   5355
         Picture         =   "APROGMRP.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Nuevo Programa de Fabricacion"
         Top             =   420
         Width           =   435
      End
      Begin VB.CommandButton Command22 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   5775
         Picture         =   "APROGMRP.frx":1912
         TabIndex        =   8
         ToolTipText     =   "Importación de Programa de Fabricación"
         Top             =   420
         Width           =   450
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   105
         TabIndex        =   7
         Top             =   420
         Width           =   5160
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Programa de Fabricacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   105
      TabIndex        =   0
      Top             =   3265
      Width           =   10665
      Begin VB.PictureBox Picture2 
         Height          =   4005
         Left            =   4770
         ScaleHeight     =   3945
         ScaleWidth      =   5460
         TabIndex        =   35
         Top             =   420
         Width           =   5520
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
            Height          =   3660
            IntegralHeight  =   0   'False
            Left            =   -40
            TabIndex        =   36
            Top             =   320
            Width           =   8145
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Cliente"
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
            TabIndex        =   37
            Top             =   45
            Width           =   960
         End
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   3645
         Left            =   10290
         TabIndex        =   26
         Top             =   770
         Width           =   225
      End
      Begin VB.CommandButton Command2 
         Enabled         =   0   'False
         Height          =   250
         Left            =   120
         TabIndex        =   28
         Top             =   4410
         Width           =   4665
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   250
         Left            =   4770
         Max             =   10000
         TabIndex        =   27
         Top             =   4410
         Width           =   5505
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         Left            =   105
         TabIndex        =   2
         Top             =   760
         Width           =   8145
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   4275
         Left            =   105
         ScaleHeight     =   4215
         ScaleWidth      =   10365
         TabIndex        =   4
         Top             =   420
         Width           =   10425
         Begin VB.CommandButton Command15 
            Caption         =   "Codigo                   "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   1920
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Descripcion de Producto       "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   1920
            TabIndex        =   16
            Top             =   0
            Width           =   2740
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   225
            Left            =   5490
            TabIndex        =   6
            Top             =   550
            Width           =   4425
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo           Descripción              Cant."
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            TabIndex        =   5
            Top             =   360
            Width           =   11460
         End
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   7  'Diagonal Cross
      Height          =   960
      Left            =   240
      Shape           =   2  'Oval
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MRP-II"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   750
      Left            =   240
      TabIndex        =   25
      Top             =   705
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   870
      Picture         =   "APROGMRP.frx":1A5C
      Top             =   75
      Width           =   480
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "P.Terminados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   4095
      TabIndex        =   15
      Top             =   840
      Width           =   3585
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "de Fabricación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   3150
      TabIndex        =   14
      Top             =   525
      Width           =   3585
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Planificación "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   2415
      TabIndex        =   13
      Top             =   210
      Width           =   2955
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10370
      Picture         =   "APROGMRP.frx":1D66
      Top             =   7570
      Width           =   2010
   End
End
Attribute VB_Name = "APROGMRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IPROGRA&
Dim MESINPRO%

Private Sub command1_Click()
   '
   Call CIERRARCH("*.*")
   Hide
   '
End Sub

Private Sub Command10_Click()
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List5.ListCount - 1
     AAA$ = Mid$(List5.List(KIL& - 1), 1, 4)
     BBB$ = Mid$(List5.List(KIL&), 1, 4)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List5.List(KIL& - 1)
          List5.RemoveItem (KIL& - 1)
          List5.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List5.List(KIL& - 1)
          List5.RemoveItem (KIL& - 1)
          List5.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
End Sub

Private Sub Command11_Click()
  Command11.Enabled = False
  Call CIERRARCH("*.*")
  If Option1.Value = True Then
      Call MENSERR(24, "No Disponible por Pantalla.\  \Imprima o Exporte a Archivo Texto.")
    Else
      Call CONECRUN("*LICAPAC", "Listado Analisis de Capacidad")
  End If
  Command11.Enabled = True
End Sub

Private Sub Command12_Click()
  Command12.Enabled = False
  Call CIERRARCH("*.*")
  If Option1.Value = True Then
      Call CONECRUN("*COCAPAC", "Consulta Resumen de Capacidad")
    Else
      Call CONECRUN("*RECAPAC", "Listado Resumen de Capacidad")
  End If
  Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List1.ListCount - 1
     AAA$ = Mid$(List1.List(KIL& - 1), 39, 10)
     BBB$ = Mid$(List1.List(KIL&), 39, 10)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
End Sub

Private Sub Command14_Click()
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
   On Error Resume Next
   List1.TopIndex = 0
   On Error GoTo 0
   '
   Screen.MousePointer = 11
10 For KIL& = 1 To List1.ListCount - 1
     AAA$ = Mid$(List1.List(KIL& - 1), 17, 22)
     BBB$ = Mid$(List1.List(KIL&), 17, 22)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
   Call CARLIS2
   Screen.MousePointer = 1
   '
End Sub

Private Sub COMMAND15_CLICK()
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
   On Error Resume Next
   List1.TopIndex = 0
   On Error GoTo 0
   '
   Screen.MousePointer = 11
10 For KIL& = 1 To List1.ListCount - 1
     AAA$ = Mid$(List1.List(KIL& - 1), 1, 15)
     BBB$ = Mid$(List1.List(KIL&), 1, 15)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
   Call CARLIS2
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command16_Click()
   Command6.Enabled = False
   '
   If List1.ListCount > 0 Then
     Screen.MousePointer = 11
     Call COPYSTRU("DEPROFAB", "LPROFAB")
     RFF$ = RECFILT$("DEPROFAB", 1, MKS$(IPROGRA&))
     JJ& = 0
     For U& = 1 To Len(RFF$) Step 4
       REFI& = CVS(Mid$(RFF$, U&, 4))
       XX$ = REGLEIDO$("DEPROFAB", REFI&)
       JJ& = JJ& + 1
       Call GRAREG("LPROFAB", XX$, JJ&)
     Next U&
     Call SETULTREG("LPROFAB", JJ&)
     Call CIERRARCH("LPROFAB")
     '
     Call HEADERS("LPROFAB", "")
     Call HEADERS("LPROFAB", "Programa: " + List3.TEXT)
     Call FINAL("*LPROFAB")
     Screen.MousePointer = 1
   End If
   '
   Command6.Enabled = True
End Sub

Private Sub Command17_Click()
   Command17.Enabled = False
   Call COMUNI("No Disponible")
   Command17.Enabled = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   Call CONECRUN("FLSETUP", "Configuración de Funcionamiento")
   Call FINAL("")
   Command18.Enabled = True
End Sub


Private Sub COMMAND21_CLICK()
   Command21.Enabled = False
   URE& = ULTREG&("PROFABRI")
   VDEF$ = MKS$(URE& + 1) + Space$(48)
10 OBX$ = OBJPLA$("NUEPROFABCLI", VDEF$)
   If OBX$ <> "" Then
     DENOPRO$ = TRIM$(Mid$(OBX$, 5, 48))
     If DENOPRO$ <> "" Then
       NCLIE% = CVI(Mid$(OBX$, 53, 2))
       If NCLIE% > 0 Then
         If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) = "" Then
           Call MENSERR(24, "Numero de Cliente no Válido")
           VDEF$ = OBX$
           GoTo 10
         End If
       End If
       '
       Call BLOQARCH("PROFABRI")
       REBLA$ = REGBLAN$("PROFABRI")
       URE& = ULTREG&("PROFABRI")
       Call REPLA(REBLA$, MKS$(URE& + 1), 1, 4)
       Call REPLA(REBLA$, DENOPRO$, 5, 48)
       Call REPLA(REBLA$, MKI$(NCLIE%), 53, 2)
       Call REPLA(REBLA$, String$(10, 0), 55, 10)
       Call REPLA(REBLA$, Space$(192), 65, 192)
       Call REGAPP("PROFABRI", REBLA$)
       Call CIERRARCH("PROFABRI")
       '
       ILISTA& = List3.ListCount
       TTXX$ = Space$(128 + 192)
       Call REPLA(TTXX$, DENOPRO$, 1, 48)
       Call REPLA(TTXX$, TRIM$(Str$(NCLIE%)), 113, 8)
       Call REPLA(TTXX$, TRIM$(Str$(URE& + 1)), 121, 8)
       List3.AddItem TTXX$
       List3.ListIndex = ILISTA&
     End If
   End If
   Command21.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   If List3.ListIndex < 0 Then
     Call MENSERR(24, "Falta Seleccionar Programa de Fabricación")
     GoTo 99
   End If
   '
   REPRO& = 1 + List3.ListIndex
   REBLA$ = REGLEIDO$("PROFABRI", REPRO&)
   Text1 = TRIM$(Str$(CVI(Mid$(REBLA$, 53, 2))))
   Text2 = LTrim$(RTrim$(Mid$(REBLA$, 65, 192)))
   '
   Call CARPROGENT
   If List1.ListCount > 0 Then
     If COMALTER%("Al Importar un Programa de Fabricación\se Sobre-Escriben los Datos Previos\de Dicho Programa, con Pérdida Irreversible.\  \Está Seguro de Continuar ?\\No, Cancelar\Si, Continuar e Importar") <> 2 Then
       GoTo 99
     End If
   End If
   IMPORMRP.Show 1
   '
99 Command22.Enabled = True
End Sub

Private Sub Command24_Click()
   Command24.Enabled = False
   If List3.ListIndex < 0 Then
     Call MENSERR(24, "Falta Seleccionar Programa de Fabricación")
     GoTo 99
   End If
   '
10 VDEF$ = ""
   OBX$ = OBJPLA$("ENTRACODIGO", VDEF$)
   If OBX$ <> "" Then
     CIXI% = CVI(OBX$)
     If CIXI% > 0 Then
       CODEX$ = TRIM$(CODEXT$(CIXI%))
       For KU% = 1 To List1.ListCount
         If TRIM$(Left$(List1.List(KU% - 1), 16)) = CODEX$ Then
           Call MENSERR(24, "Codigo ya Programado")
           List1.ListIndex = KU% - 1
           List2.ListIndex = KU% - 1
           GoTo 10
         End If
       Next KU%
       '
       TTXX$ = Space$(64) + Space$(225)
       Call REPLA(TTXX$, CODEX$, 1, 16)
       Call REPLA(TTXX$, DESCRIT0$(CIXI), 17, 48)
       List1.AddItem TTXX$
       List1.ListIndex = List1.ListCount - 1
       List2.AddItem Mid$(TTXX$, 65)
       List2.ListIndex = List2.ListCount - 1
       GoTo 10
     End If
   End If
   '
99 Command24.Enabled = True
End Sub

Private Sub Command3_Click()
   NPROG% = 1 + List3.ListIndex
   If NPROG% < 1 Then Exit Sub
   '
   Command3.Enabled = False
   Dim FEPLA%(24)
   For KU% = 0 To 24
     TTYY$ = TRIM$(Mid$(Label5, 9 * KU% + 1, 9))
     If KU% > 0 Then
         TTYY$ = MESANO$(TTYY$, DES%, HAS%)
         FEPLA%(KU%) = HAS%
       Else
         FEPLA%(KU%) = DIANTE(FECHANUM("01/01/" + Mid$(HOS$, 7, 2)))
     End If
   Next KU%
   '
   NCLIE% = Val(Text1)
   If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) = "" Then NCLIE% = 0
   '
   REPROG& = NPROG%
   REBLA$ = REGLEIDO$("PROFABRI", REPROG&)
     DENOPRO$ = Left$(List3.List(NPROG% - 1), 48)
   Call REPLA(REBLA$, DENOPRO$, 5, 48)
   Call REPLA(REBLA$, MKI$(NCLIE%), 53, 2)
   Call REPLA(REBLA$, String$(10, 0), 55, 10)
   Call REPLA(REBLA$, Text2, 65, 192)
   Call GRAREG("PROFABRI", REBLA$, REPROG&)
   '
   REBLA$ = REGBLAN$("PROGENT")
   Call REPLA(REBLA$, MKI$(NCLIE%), 5, 2)
   JJ& = 0
   For U1& = 1 To List1.ListCount
     TTXX$ = List1.List(U1& - 1)
     CIXI% = CODINT%(Left$(TTXX$, 16))
     If CIXI% > 0 Then
       TTZZ$ = Mid$(TTXX$, 65, 225)
       For KU% = 0 To 24
         CAPROL = Val(Mid$(TTZZ$, 9 * KU% + 1, 9))
         If Abs(CAPROL) > 0.05 Then
           REPRO$ = REBLA$
           Call REPLA(REPRO$, MKI$(CIXI%), 1, 2)
           Call REPLA(REPRO$, MKI$(FEPLA%(KU%)), 3, 2)
           Call REPLA(REPRO$, MKS$(CAPROL), 7, 4)
           JJ& = JJ& + 1
           Call GRAREG("!PROGENT", REPRO, JJ&)
         End If
       Next KU%
     End If
   Next U1&
   Call SETULTREG("!PROGENT", JJ&)
   '
   Call FILTERUPDATE("PROGENT", "", 10, MKI$(NPROG%))
   Command3.Enabled = True
   '
End Sub


Private Sub Command4_Click()
TTXX$ = ""
For KU& = 1 To ULTREG&("PROGNETA")
XX$ = REGLEIDO$("PROGNETA", KU&)
TTXX$ = TTXX$ + Str$(CVI(Mid$(XX$, 3, 2)))
Next KU&
MsgBox TTXX$
Stop
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List4.ListCount - 1
     AAA$ = Mid$(List4.List(KIL& - 1), 1, 6)
     BBB$ = Mid$(List4.List(KIL&), 1, 6)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
End Sub

Private Sub Command5_Click()
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List4.ListCount - 1
     AAA$ = Mid$(List4.List(KIL& - 1), 9, 15)
     BBB$ = Mid$(List4.List(KIL&), 9, 15)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
End Sub

Private Sub Command6_Click()
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List4.ListCount - 1
     AAA$ = Mid$(List4.List(KIL& - 1), 24, 12)
     BBB$ = Mid$(List4.List(KIL&), 24, 12)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   AVAMRPSB.Show 1
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List5.ListCount - 1
     AAA$ = Mid$(List5.List(KIL& - 1), 24, 12)
     BBB$ = Mid$(List5.List(KIL&), 24, 12)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List5.List(KIL& - 1)
          List5.RemoveItem (KIL& - 1)
          List5.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List5.List(KIL& - 1)
          List5.RemoveItem (KIL& - 1)
          List5.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
End Sub

Private Sub Command9_Click()
   '
   Call SELECHO("DEUDOR1")
   VDEVU = Val(VALACT1$("DEUDOR1"))
   If VDEVU > 0 Then
     Text1 = TRIM$(Str$(VDEVU))
   End If
   '
End Sub

Private Sub Form_Load()
   '
   Call CENTRAFORM(Me)
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
10   Caption = Caption + "  -  " + EPAC$
     Call GRATITFOR(Caption)
   End If
   OPENFORMS = DoEvents
   '
   Call PREMESES
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("PROFABRI")
     XX$ = REGLEIDO$("PROFABRI", U&)
     DENOPRO$ = TRIM$(Mid$(XX$, 5, 48))
     If DENOPRO$ <> "" Then
       TTXX$ = Space$(128)
       Call REPLA(TTXX$, DENOPRO$, 1, 48)
       Call REPLA(TTXX$, TRIM$(Str$(U&)), 121, 8)
       List3.AddItem TTXX$
     End If
   Next U&
   If List3.ListCount > 0 Then
     List3.ListIndex = List3.ListCount - 1
   End If
   Screen.MousePointer = 1
   '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
      Call CIERRARCH("*.*")
      Call BAJALDISCO
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
End Sub

Private Sub HScroll1_Change()
  List2.Left = (-1) * 10 * HScroll1.Value - 40
  Label5.Left = List2.Left
End Sub

Private Sub HScroll1_Scroll()
  List2.Left = (-1) * 10 * HScroll1.Value - 40
  Label5.Left = List2.Left
End Sub

Private Sub Label1_DblClick()
   Call CARPROFAB
End Sub

Sub CARPROFAB()
   DEPROGRA$ = List3.TEXT
   IPROGRA& = Val(Mid$(DEPROGRA$, 121, 8))
   If IPROGRA& < 1 Then
        Call MENSERR(24, "Imposible Abrir Programa '0'")
        Exit Sub
      Else
        Call TRALOFIL("DEPROFAB", 1, MKS$(IPROGRA&))
        VTB% = VENTABU%("CARPROFA")
        If VTB% >= 0 Then
          Call TRACENFIL("DEPROFAB", 1, MKS$(IPROGRA&))
          Call LISPROFAB
        End If
   End If
End Sub
'
Sub LISPROFAB()
   DEPROGRA$ = List3.TEXT
   IPROGRA& = Val(Mid$(DEPROGRA$, 121, 8))
   If IPROGRA& < 1 Then
        Call MENSERR(24, "Imposible Abrir Programa '0'")
        Exit Sub
      Else
        SCMA = Screen.MousePointer
        Screen.MousePointer = 11
        RFF$ = RECFILT$("DEPROFAB", 1, MKS$(IPROGRA&))
        List1.Clear
        For U& = 1 To Len(RFF$) Step 4
          REFI& = CVS(Mid$(RFF$, U&, 4))
          XX$ = REGLEIDO$("DEPROFAB", REFI&)
          CIXI% = CVI(Mid$(XX$, 5, 2))
          CAXX = CVS(Mid$(XX$, 7, 4))
          TTXX$ = Space$(64)
          Call REPLA(TTXX$, CODEXT$(CIXI%), 1, 16)
          Call REPLA(TTXX$, DESCRIT0$(CIXI%), 17, 22)
          Call REPLA(TTXX$, FORMATNUM$(CAXX, "F10.1"), 39, 10)
          List1.AddItem TTXX$
        Next U&
        '
        Call TRALOFIL("RECALMAQ", 1, MKS$(IPROGRA&))
        Call TRALOFIL("RECALMOD", 1, MKS$(IPROGRA&))
        '
         Label5 = "": TOTIX = 0
        For U& = 1 To ULTREG&("!RECALMAQ")
          TTXX$ = Mid$(REGLEIDO$("!RECALMAQ", U&), 5)
          List4.AddItem TTXX$
          TOTIX = TOTIX + Val(Mid$(TTXX$, 34, 2)) + 60 * Val(Mid$(TTXX$, 31, 2)) + 3600 * Val(Mid$(TTXX$, 26, 4))
        Next U&
        'List4.Refresh
        If TOTIX >= 1 Then
          Label5 = HORMINSE$(TOTIX)
        End If
        '
        'List5.Clear:
        Label10 = "": TOTIX = 0
        For U& = 1 To ULTREG&("!RECALMOD")
          TTXX$ = Mid$(REGLEIDO$("!RECALMOD", U&), 5)
          List5.AddItem TTXX$
          TOTIX = TOTIX + Val(Mid$(TTXX$, 34, 2)) + 60 * Val(Mid$(TTXX$, 31, 2)) + 3600 * Val(Mid$(TTXX$, 26, 4))
        Next U&
        'List5.Refresh
        If TOTIX >= 1 Then
          Label10 = HORMINSE$(TOTIX)
        End If
        '
        Screen.MousePointer = SCMA
   End If
End Sub
'

Private Sub List1_Click()
   On Error Resume Next
   List2.ListIndex = List1.ListIndex
   On Error GoTo 0
End Sub

Private Sub List1_DblClick()
   CODEX$ = TRIM$(Left$(List1.TEXT, 16))
   VECAN$ = Mid$(List1.TEXT, 65, 225)
   ILISTA& = List1.ListIndex
   REBLA$ = REGBLAN$("PROGMEN")
   HOII% = HOY(HOS$)
   For KU% = 0 To 24
     TTYY$ = TRIM$(Mid$(Label5, 9 * KU% + 1, 9))
     CAMEN# = Val(TRIM$(Mid$(VECAN$, 9 * KU% + 1, 9)))
     DES% = 0: HAS% = DIANTE(FECHANUM("01/01/" + Mid$(HOS$, 7, 2)))
     If KU% > 0 Then
       TTYY$ = MESANO$(TTYY$, DES%, HAS%)
     End If
     REPRO$ = REBLA$
     Call REPLA(REPRO$, TTYY$, 1, 16)
     Call REPLA(REPRO$, MKI$(DES%), 17, 2)
     Call REPLA(REPRO$, MKI$(HAS%), 19, 2)
     Call REPLA(REPRO$, MKD$(CAMEN#), 21, 8)
     JJ& = KU% + 1
     Call GRAREG("!PROGMEN", REPRO$, JJ&)
   Next KU%
   Call SETULTREG("!PROGMEN", JJ&)
   '
   titupa$ = TRIM$(CODEX$) + " - " + TRIM$(DESCRIT0$(CODINT%(CODEX$)))
   VTB% = VENTABU%("PROGMEN/TITUPAN=" + titupa$)
   If VTB% >= 0 Then
     TTYY$ = ""
     For JJ& = 0 To 24
       XX$ = REGLEIDO$("!PROGMEN", JJ& + 1)
       TTZZ$ = AJUSTD$(TEXTBIN$(Mid$(XX$, 21, 8), "F9.1)"), 9)
       TTYY$ = TTYY$ + TTZZ$
     Next JJ&
     '
     TTXYZ$ = List1.List(ILISTA&)
     List1.RemoveItem ILISTA&
     List2.RemoveItem ILISTA&
     Call REPLA(TTXYZ$, TTYY$, 65, 225)
     List1.AddItem TTXYZ$, ILISTA&
     List2.AddItem TTYY$, ILISTA&
     On Error Resume Next
     List1.ListIndex = ILISTA&
     On Error GoTo 0
     '
   End If
End Sub

Sub CARLIS2()
   List2.Clear
   For U& = 1 To List1.ListCount
     TTXYZ$ = Mid$(List1.List(U& - 1), 65)
     List2.AddItem TTXYZ$
   Next U&
   '
   List2.TopIndex = List1.TopIndex
   List2.ListIndex = List1.ListIndex
   '
End Sub

Private Sub List1_Scroll()
   On Error Resume Next
   VScroll2.Value = List1.TopIndex
   On Error GoTo 0
End Sub

Private Sub List2_Click()
   List1.ListIndex = List2.ListIndex
End Sub

Private Sub List2_DblClick()
   Call List1_DblClick
End Sub

'
Private Sub List3_Click()
   List1.Clear
   List2.Clear
   Text1 = ""
   Text2 = ""
End Sub

Private Sub LIST3_DBLCLICK()
   REPRO& = 1 + List3.ListIndex
   REBLA$ = REGLEIDO$("PROFABRI", REPRO&)
   Text1 = TRIM$(Str$(CVI(Mid$(REBLA$, 53, 2))))
   Text2 = LTrim$(RTrim$(Mid$(REBLA$, 65, 192)))
   '
   Call CARPROGENT
End Sub


Sub PREMESES()
   '
   CAMESA% = 24
   TTXX$ = "   ATRASO"
   HOII = HOY(HOS$)
   MESINPRO% = 12 * (100 + Val(Mid$(HOS$, 7, 2)))
   '
   Label5 = "": Label5.Left = 0
   'List4.Left = -30
   
   For KKI% = 1 To 24
     TTXX$ = TTXX$ + AJUSTD$(ABREMESAN$(KKI% + MESINPRO%, 0), 9)
   Next KKI%
   Label5 = TTXX$
   List2.Width = Label5.Width + 500
   '
   HScroll1.Value = 0
   HScroll1.Max = (Label5.Width - HScroll1.Width) / 10
   On Error Resume Next
     HScroll1.SmallChange = HScroll1.Max / 100
     HScroll1.LargeChange = HScroll1.SmallChange * 5
   On Error GoTo 0
   '
End Sub

Sub CARPROGENT()
   '
   List1.Clear
   List2.Clear
   NPROG% = 1 + List3.ListIndex
   If NPROG% < 1 Then Exit Sub
   '
   Dim CODIPRO%(16384), PROGENT$(16384)
   Dim FEPLA%(24)
   '
   For KU% = 0 To 24
     TTYY$ = TRIM$(Mid$(Label5, 9 * KU% + 1, 9))
     If KU% > 0 Then
         TTYY$ = MESANO$(TTYY$, DES%, HAS%)
         FEPLA%(KU%) = HAS%
       Else
         FEPLA%(KU%) = DIANTE(FECHANUM("01/01/" + Mid$(HOS$, 7, 2)))
     End If
   Next KU%
   '
   For KKI% = 0 To NUMAS%
     PROGENT$(KKI%) = String$(100, 0)
   Next KKI%
   '
   RFF$ = RECFILT$("PROGENT", 10, MKI$(NPROG%))
   '
   FIN& = Len(RFF$)
   For U& = 1 To FIN& Step 4
     Call TRACE(20, U&, FIN&)
     I& = CVS(Mid$(RFF$, U&, 4))
     P00$ = REGLEIDO$("PROGENT", I&)
     CI% = CVI(Left$(P00$, 2))
     '
     If CI% > 0 Then
       If CI% <= NUMAS% Then
         If STATUIT%(CI%) >= 0 Then
           FEXI% = CVI(Mid$(P00$, 3, 2))
           If FEXI% <= FEPLA%(0) Then
              COLUX% = 0
              GoTo 20
            Else
              For KKJ% = 24 To 1 Step -1
                If FEXI% >= FEPLA%(KKJ%) Then
                  COLUX% = KKJ%
                  GoTo 20
                End If
              Next KKJ%
           End If
           GoTo 39
           '
20         CANPRO = CVS(Mid$(P00$, 7, 4))
           '
 'TTYYU$ = TTYYU$ + CODEXT$(CI%) + "     " + Str$(CANPRO) + Chr$(13) + Chr$(10)
           For KKI% = 1 To CACODIP%
             If CODIPRO%(KKI%) = CI% Then GoTo 30
           Next KKI%
           CACODIP% = CACODIP% + 1
           CODIPRO%(CACODIP%) = CI%
           KKI% = CACODIP%
           '
30         If COLUX% >= 0 Then
             If COLUX% <= 24 Then
               CANPRO = CANPRO + CVS(Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4))
               Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4) = MKS$(CANPRO)
             End If
           End If
         End If
       End If
     End If
39 Next U&
   '
'MsgBox TTYYU$
   For KKI% = 1 To CACODIP%
     TTXYZ$ = Space$(64) + Space$(225)
     CIKKI% = CODIPRO%(KKI%)
     Call REPLA(TTXYZ$, CODEXT$(CIKKI%), 1, 16)
     Call REPLA(TTXYZ$, DESCRIT0$(CIKKI%), 17, 48)
     '
     TTYY$ = ""
     For JJ& = 0 To 24
       XX$ = Mid$(PROGENT$(KKI%), 4 * JJ& + 1, 4)
       TTZZ$ = AJUSTD$(TEXTBIN$(XX$, "F9.1)"), 9)
       TTYY$ = TTYY$ + TTZZ$
     Next JJ&
     '
     Call REPLA(TTXYZ$, TTYY$, 65, 225)
     List1.AddItem TTXYZ$
     List2.AddItem TTYY$
     Screen.MousePointer = 1
   Next KKI%
   '
   MXX% = List1.ListCount - 16
   If MXX% < 0 Then MXX% = 0
   VScroll2.Value = 0
   VScroll2.Max = MXX%
   '
End Sub


Private Sub Text1_Change()
   If Val(Text1) < 1 Or Val(Text1) > 32767 Then
      Text1 = ""
      NCLIE% = 0
      Label8 = "Clientes Varios no Identificados"
      Exit Sub
    Else
      NCLIE% = Val(Text1)
      Label8 = TRIM$(ECOARCH("DEUDOR1", MKI$(NCLIE%)))
      If Label8 = "" Then
        Label8 = "Clientes Varios no Identificados"
      End If
   End If
End Sub

Private Sub VScroll2_Change()
   On Error Resume Next
   List1.TopIndex = VScroll2.Value
   List2.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub

Private Sub VScroll2_Scroll()
   On Error Resume Next
   List1.TopIndex = VScroll2.Value
   List2.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub
