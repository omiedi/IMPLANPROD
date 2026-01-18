VERSION 5.00
Begin VB.Form LISTAPRE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0FFEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listas de Precios de Venta"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11850
   Icon            =   "LISTAPRE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11850
   Begin VB.ListBox SORTLIST 
      Height          =   255
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Utilities"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   8715
      TabIndex        =   32
      Top             =   210
      Width           =   1740
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   140
         ScaleHeight     =   585
         ScaleWidth      =   1395
         TabIndex        =   39
         Top             =   1890
         Width           =   1450
         Begin VB.CommandButton Command13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "LISTAPRE.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Control y Depuración de Lista Activa"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Control"
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
            TabIndex        =   41
            Top             =   210
            Width           =   750
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   140
         ScaleHeight     =   585
         ScaleWidth      =   1395
         TabIndex        =   36
         Top             =   1155
         Width           =   1450
         Begin VB.CommandButton Command12 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "LISTAPRE.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Ordena por Código de Artículo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Sort por Código"
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
            TabIndex        =   38
            Top             =   105
            Width           =   750
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   140
         ScaleHeight     =   585
         ScaleWidth      =   1395
         TabIndex        =   33
         Top             =   420
         Width           =   1450
         Begin VB.CommandButton Command11 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Actualización por Porcentaje"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Ajuste +/- (%)"
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
            TabIndex        =   35
            Top             =   105
            Width           =   750
         End
      End
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
      Height          =   4785
      Left            =   210
      TabIndex        =   0
      Top             =   3465
      Width           =   10410
   End
   Begin VB.PictureBox Picture1 
      Height          =   435
      Left            =   210
      ScaleHeight     =   375
      ScaleWidth      =   10365
      TabIndex        =   30
      Top             =   3045
      Width           =   10430
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   $"LISTAPRE.frx":091E
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
         Left            =   0
         TabIndex        =   31
         Top             =   60
         Width           =   10410
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10760
      ScaleHeight     =   75
      ScaleWidth      =   945
      TabIndex        =   12
      Top             =   7350
      Width           =   1000
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Lista de Precio Activa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   4620
      TabIndex        =   11
      Top             =   210
      Width           =   3945
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   29
         Top             =   2250
         Width           =   435
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
         Left            =   945
         TabIndex        =   25
         Top             =   2220
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
         Left            =   945
         TabIndex        =   21
         Top             =   1590
         Width           =   175
      End
      Begin VB.CommandButton Command2 
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
         Left            =   3570
         TabIndex        =   19
         Top             =   990
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
         Left            =   630
         TabIndex        =   16
         Top             =   1000
         Width           =   1590
      End
      Begin VB.TextBox Text1 
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
         Left            =   210
         TabIndex        =   14
         Top             =   420
         Width           =   3585
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Decimal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2730
         TabIndex        =   28
         Top             =   1995
         Width           =   855
      End
      Begin VB.Label Label9 
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
         Left            =   1260
         TabIndex        =   27
         Top             =   2235
         Width           =   1800
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   26
         Top             =   1995
         Width           =   855
      End
      Begin VB.Label Label7 
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
         Left            =   630
         TabIndex        =   24
         Top             =   2235
         Width           =   330
      End
      Begin VB.Label Label6 
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
         Left            =   1260
         TabIndex        =   23
         Top             =   1600
         Width           =   2535
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formato:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   22
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label Label4 
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
         Left            =   630
         TabIndex        =   20
         Top             =   1600
         Width           =   330
      End
      Begin VB.Label Label3 
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
         Left            =   2730
         TabIndex        =   18
         Top             =   1000
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vigencia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2100
         TabIndex        =   17
         Top             =   750
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Abrev:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -105
         TabIndex        =   15
         Top             =   750
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Indice de Listas de Precios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   210
      TabIndex        =   9
      Top             =   210
      Width           =   4260
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   210
         TabIndex        =   10
         Top             =   420
         Width           =   3795
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   10760
      TabIndex        =   6
      Top             =   105
      Width           =   1000
      Begin VB.CommandButton Command1 
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
         Height          =   740
         Left            =   105
         Picture         =   "LISTAPRE.frx":09A8
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   265
         Width           =   750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "LISTAPRE.frx":0AF2
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1000
         Width           =   750
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00D0FFEE&
      Caption         =   "A-B-M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   10760
      TabIndex        =   3
      Top             =   1995
      Width           =   1000
      Begin VB.CommandButton Command6 
         Caption         =   "Opcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "LISTAPRE.frx":0DFC
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Opciones Varias"
         Top             =   3270
         Width           =   750
      End
      Begin VB.CommandButton Command21 
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
         Height          =   740
         Left            =   105
         Picture         =   "LISTAPRE.frx":1106
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Alta de Nuevo Item de Stock"
         Top             =   1015
         Width           =   750
      End
      Begin VB.CommandButton Command9 
         Caption         =   "More"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "LISTAPRE.frx":1250
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Opciones Complementarias"
         Top             =   2520
         Width           =   750
      End
      Begin VB.CommandButton COMMAND3 
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
         Height          =   740
         Left            =   105
         Picture         =   "LISTAPRE.frx":155A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Guarda Configuración de la Lista"
         Top             =   265
         Width           =   750
      End
      Begin VB.CommandButton Command7 
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
         Height          =   740
         Left            =   105
         Picture         =   "LISTAPRE.frx":1864
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Nueva Lista de Precios"
         Top             =   1765
         Width           =   750
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   10760
      TabIndex        =   1
      Top             =   6195
      Width           =   1000
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   105
         Picture         =   "LISTAPRE.frx":1B6E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Emite Lista de Precios Impresa"
         Top             =   265
         Width           =   750
      End
   End
   Begin VB.Line Line3 
      X1              =   210
      X2              =   11865
      Y1              =   8250
      Y2              =   8250
   End
   Begin VB.Line Line1 
      X1              =   10635
      X2              =   10635
      Y1              =   -105
      Y2              =   8295
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10290
      Picture         =   "LISTAPRE.frx":21D8
      Top             =   7740
      Width           =   2010
   End
End
Attribute VB_Name = "LISTAPRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LISTACTIVA&
'
Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command10_Click()
   Call SELECHO("VALIMONE")
   If VALACT1$("VALIMONE") <> "" Then
     Label7 = VALACT1$("VALIMONE")
   End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   If DERACCE%("APREPOR", "Actualización de Precios por Porcentaje") = 2 Then
      Call APREPOR
      Screen.MousePointer = 1
   End If
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call SORLIPRE
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
      If VERILLAVE% <> 1 Then GoTo 99
   End If
   '
   If DERACCE%("DEPUPREC", "Depurar Listas de Precios de Venta") < 2 Then
     GoTo 99
   End If
   '
   LIMPIAPRE.Show 1
   DoEvents
   '
   If LIMPIAPRE.Check1(2).Value = Checked Then
     Call VEPRECOS
     GoTo 99
   End If
   '
   LICOBAJ% = 0: If LIMPIAPRE.Check1(0).Value = Checked Then LICOBAJ% = 1
   LISINST% = 0: If LIMPIAPRE.Check1(1).Value = Checked Then LISINST% = 1
   LISDUPL% = 0: If LIMPIAPRE.Check1(3).Value = Checked Then LISDUPL% = 1
   LIMPIAPRE.Check1(0).Value = Unchecked
   LIMPIAPRE.Check1(1).Value = Unchecked
   LIMPIAPRE.Check1(2).Value = Unchecked
   LIMPIAPRE.Check1(3).Value = Unchecked
   DoEvents
   '
   If LICOBAJ% > 0 Or LISINST% > 0 Or LISDUPL% > 0 Then
     If LISTACTIVA& < 1 Then GoTo 99
     ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
     While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
     ARCHIPRE$ = "LIPRE" + ARCHIPRE$
     '
     Call BLOQARCH(ARCHIPRE$) '\\\OT-06-009-WAR-21/03/06///
     '
     Screen.MousePointer = 11
     JJ& = 0
     NMS% = NUMAS%
     VECODI$ = String$(NMS%, Chr$(0))
     FIN& = ULTREG&(ARCHIPRE$)
     '
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$(ARCHIPRE$, U&)
       CIXI% = CVI(Left$(X$, 2))
       '
       If LICOBAJ% = 1 Then
         If Mid$(CODEXT$(CIXI%), 16, 1) = Chr$(96) Then
           GoTo 19
         End If
       End If
       '
       If LISINST% = 1 Then
         If STOCKACT(CIXI%) < 0.05 Then
           GoTo 19
         End If
       End If
       '
       If LISDUPL% = 1 Then
         If CIXI% < 1 Or CIXI% > NMS% Then GoTo 19
         If Asc(Mid$(VECODI$, CIXI%, 1)) > 0 Then
           TTXX$ = "Código " + TRIM$(CODEXT$(CIXI%)) + " Duplicado en Lista Activa"
           If COMALTER%(TTXX$ + "\\Continuar Control\Eliminar") = 2 Then
             GoTo 19
           End If
         End If
         Mid$(VECODI$, CIXI%, 1) = Chr$(1)
       End If
       '
       JJ& = JJ& + 1
       Call GRAREG("!LIPREVEN", X$, JJ&)
       '
19   Next U&
     '
     Call SETULTREG("!LIPREVEN", JJ&)
     Call CIERRARCH("!LIPREVEN")
     Screen.MousePointer = 1
     '
     Call EDITLISTA(ARCHIPRE$)
     '
     Call LIBARCH(ARCHIPRE$) '\\\OT-06-009-WAR-21/03/06///
     '
   End If
   '
99 Command13.Enabled = True
   '
End Sub

Sub VEPRECOS()
   '
   If LISTACTIVA& < 1 Then Exit Sub
   '
   ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   '
10 VDEF$ = MKS$(1.5) + MKS$(4)
   OBX$ = OBJPLA$("COEFICOSPRE", VDEF$)
   If OBX$ <> "" Then
     Screen.MousePointer = 11
     COEFIMI = CVS(Left$(OBX$, 4))
     COEFIMA = CVS(Mid$(OBX$, 5, 4))
     If COEFIMI > COEFIMA Then GoTo 10
     '
     X$ = REGLEIDO$("LISTAS", LISTACTIVA&)
     TITU$ = TRIM$(Mid$(X$, 2, 46))
     Call HEADERS("LIPREVEN", "")
     Call HEADERS("LIPREVEN", "Lista: " + TITU$ + " - Precios a Revisar por Costos")
     '
     JJ& = 0
     FIN& = ULTREG&(ARCHIPRE$)
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$(ARCHIPRE$, U&)
       CIXI% = CVI(Left$(XX$, 2))
       PRE# = CVD(Mid$(XX$, 62, 8))
       COCAL = COSCALCU(CIXI%, MODOCO%)
       If PRE# < COEFIMI * COCAL Or PRE# > COEFIMA * COCAL Then
         JJ& = JJ& + 1
         Call GRAREG("LIPREVEN", XX$, JJ&)
       End If
     Next U&
     Call SETULTREG("LIPREVEN", JJ&)
     '
     If JJ& < 1 Then
          Call MENSERR(24, "Sin Valores Fuera de Rango.")
        Else
          Screen.MousePointer = 1
          Call FINAL("*LILIPRE") '
     End If
   End If
   '
End Sub
'
Private Sub Command2_Click()
   Call SELECHO("SELFECHA")
   If VALACT1$("SELFECHA") <> "" Then
     Label3 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub COMMAND21_CLICK()
   Command1.Enabled = False
   '
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
   Call REPLA(VDEF$, "U.", 109, 2)
   Call REPLA(VDEF$, Chr$(1), 93, 1)
   '
   OKX% = 0
   Call ALTACODIGO(VDEF$, OKX%)
   '
   Command1.Enabled = True
End Sub

Private Sub Command3_Click()
   LISTACTIVA& = List1.ListIndex + 1
   XX$ = REGLEIDO$("LISTAS", LISTACTIVA&)
   Mid$(XX$, 2, 46) = AJUSTI$(Text1, 46)
   Mid$(XX$, 51, 12) = AJUSTI$(Text2, 12)
   FEXI% = FECHANUM(Label3)
   Mid$(XX$, 63, 2) = MKI$(FEXI%)
   Mid$(XX$, 48, 1) = Chr$(Val(Label4))
   Mid$(XX$, 50, 1) = Chr$(Val(Label7))
   If Val(Text3) < 0 Or Val(Text3) > 5 Then Text3 = "2"
   Mid$(XX$, 49, 1) = Chr$(Val(Text3))
   Call GRAREG("LISTAS", XX$, LISTACTIVA&)
   Call CIERRARCH("LISTAS")
   Call CARILISTAS
End Sub

Private Sub Command4_Click()
   Call SELECHO("VALIFORM")
   If VALACT1$("VALIFORM") <> "" Then
     Label4 = VALACT1$("VALIFORM")
   End If
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call OPNOIN("")
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
        '
        Command7.Enabled = False

3000    ' ************************* nueva lista de precios
        '
        If ULTREG&("LISTAS") >= 99 Then
            Call COMUNI("Hay 99 Listas de Precios\ya Definidas\-\Capacidad del Sistema Saturada.")
            Exit Sub
        End If
        '
        NULIS% = 1 + ULTREG&("LISTAS")
        X$ = Chr$(NULIS%) + Mid$(REGBLAN$("LISTAS"), 2)
3010    X$ = OBJPLA$("LISPRE-NUEVA", X$)
        If X$ = "" Then
            GoTo 99
        End If
        '
        If TRIM$(Mid$(X$, 2, 46)) <> "" Then
          If Asc(Mid$(X$, 48, 1)) > 0 Then
            If Asc(Mid$(X$, 48, 1)) <= 2 Then
              If Asc(Mid$(X$, 49, 1)) <= 4 Then
                If Asc(Mid$(X$, 50, 1)) > 0 Then
                  If Asc(Mid$(X$, 50, 1)) <= 2 Then
                    If TRIM$(Mid$(X$, 51, 12)) <> "" Then
                      If CVI(Mid$(X$, 63, 2)) > 0 Then
                        Call REGAPP("LISTAS", X$)
                        Call CIERRARCH("LISTAS")
                        Call CARILISTAS
                        GoTo 99
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
        '
        Call COMUNI("Datos Incompletos\o No Válidos")
        GoTo 3010
        '
99      Command7.Enabled = True
        '
End Sub

Private Sub Command8_Click()
  Command8.Enabled = False
  Call LISTAPRE
  Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   OPLISPRE.Show 1
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
   '
   VERANTER$ = "ADMIPRE"
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
       Caption = Caption + " - " + EPAC$
   End If
   '
   Show
   Refresh
   '
   Screen.MousePointer = 11
   If ULTREG&("VALIFORM") < 3 Then
      Call BLANARCH("VALIFORM")
      Call REGAPP("VALIFORM", Chr$(1) + "Nro. Original")
      Call REGAPP("VALIFORM", Chr$(2) + "Descuento por Cantidad")
      Call REGAPP("VALIFORM", Chr$(3) + "Fecha Ultima Actualizacion")
      Call CIERRARCH("VALIFORM")
   End If
   '
   If ULTREG&("VALIMONE") < 2 Then
      Call BLANARCH("VALIMONE")
      Call REGAPP("VALIMONE", Chr$(1) + "Pesos")
      Call REGAPP("VALIMONE", Chr$(2) + "Dolares U$S")
      Call CIERRARCH("VALIMONE")
   End If
   '
   CONTROLI$ = TRIM$(CONTROL$("LIPREVEN", "CONTRODE"))
   If CONTROLI$ = "" Then
     Screen.MousePointer = 11
     For U& = 1 To ULTREG&("LIPRE001")
       X$ = REGLEIDO$("LIPRE001", U&)
       CII% = CVI(Left$(X$, 2))
       DEAR$ = TRIM$(UCase$(DESCRIT0$(CII%)))
       If InStr(DEAR$, "ARTICULO") > 0 Then
         REMA& = CII% + 1
         DEAR$ = TRIM$(Mid$(X$, 3, 42))
         Y$ = REGLEIDO$("MASTER", REMA&)
         Call REPLA(Y$, DEAR$, 5, 64)
         Call GRAREG("MASTER", Y$, REMA&)
       End If
     Next U&
     Call CIERRARCH("MASTER")
     'HOII% = HOY(HOS$)
     'Call GRACONTROL("LIPREVEN", "CONTRODE", HOS$)
     Screen.MousePointer = 1
   End If
   '
   Call CARILISTAS
   ILIOO& = 0
   If Val(APLINVO$) > 0 Then
     ILI00& = Val(APLINVO$) - 1
   End If
   LISTACTIVA& = ILI00&
   On Error Resume Next
   List1.ListIndex = ILI00&
   On Error GoTo 0
   Screen.MousePointer = 1
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label15_DblClick()
   Call List2_DblClick
End Sub

Private Sub Label4_Change()
   Label6 = ECOARCH$("VALIFORM", Chr$(Val(Label4)))
   If Val(Label4) = 3 Then
       Label15 = " Código                     Descripción                                             M.          Fecha Act.                      Precio"
     Else
       Label15 = " Código                     Descripción                                             M.          Nro.Original                    Precio"
   End If
   Label15.Refresh
End Sub

Private Sub Label4_DblClick()
   Call SELECHO("VALIFORM")
   If VALACT1$("VALIFORM") <> "" Then
     Label4 = VALACT1$("VALIFORM")
   End If
End Sub

Private Sub Label7_Change()
   Label9 = ECOARCH$("VALIMONE", Chr$(Val(Label7)))
End Sub

Private Sub Label7_DblClick()
   Call SELECHO("VALIMONE")
   If VALACT1$("VALIMONE") <> "" Then
     Label7 = VALACT1$("VALIMONE")
   End If
End Sub

Private Sub List1_Click()
   If List1.ListIndex <> LISTACTIVA& - 1 Then
     Call CARLISTA
   End If
End Sub

Sub CARLISTA()
   '
   ' 21/03/2006 - bloqueo de c odigos repetidos en la lista
   Dim YALIS%(32767)
   '
   Screen.MousePointer = 11
   LISTACTIVA& = List1.ListIndex + 1
   XX$ = REGLEIDO$("LISTAS", LISTACTIVA&)
   Text1 = Mid$(XX$, 2, 46)
   Text2 = Mid$(XX$, 51, 12)
   FEX = CVI(Mid$(XX$, 63, 2)): Label3 = FECHATEX$(FEX)
   Label4 = Asc(Mid$(XX$, 48, 1))
   Label7 = Asc(Mid$(XX$, 50, 1))
   Text3 = Asc(Mid$(XX$, 49, 1))
   '
   Call FRESHALL
   List2.Clear
   Call FRESHALL
   '
   ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   '
   Call COPYSTRU("LIPREVEN", ARCHIPRE$)
   '
   NMS% = NUMAS%
   FIN& = ULTREG&(ARCHIPRE$)
   JJ& = 0
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$(ARCHIPRE$, U&)
     CII% = CVI(Left$(XX$, 2))
     If CII% > 0 Then
       If CII% <= NMS% Then
         If YALIS%(CII%) < 1 Then
           DEX$ = Mid$(XX$, 3, 31)
           If TRIM$(DEX$) = "" Then
             DEX$ = DESCRIT0$(CII%)
             Call REPLA(XX$, DEX$, 3, 31)
           End If
           JJ& = JJ& + 1
           Call GRAREG(ARCHIPRE$, XX$, JJ&)
           YALIS%(CII%) = 1
           '
           CCC$ = TRIM$(Mid$(XX$, 45, 1))
           ORI$ = TRIM$(Mid$(XX$, 46, 16))
           PRE# = CVD(Mid$(XX$, 62, 8))
           '
           TTXX$ = Space$(128)
           Call REPLA(TTXX$, CODEXT$(CII%), 1, 16)
           Call REPLA(TTXX$, DEX$, 18, 31)
           Call REPLA(TTXX$, CCC$, 51, 2)
           If Val(Label4) <> 3 Then
               Call REPLA(TTXX$, ORI$, 55, 16)
             Else
               Call REPLA(TTXX$, ORI$, 58, 13)
           End If
           Call REPLA(TTXX$, CSTRING$(MKD$(PRE#), 3, 12, 4, 2), 72, 12)
           List2.AddItem TTXX$
         End If
       End If
     End If
     '
   Next U&
   Call SETULTREG(ARCHIPRE$, JJ&)
   Call CIERRARCH(ARCHIPRE$)
   '
   Screen.MousePointer = 1
   '
End Sub

Private Sub List2_DblClick()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
      If VERILLAVE% <> 1 Then Exit Sub
   End If
   '
   ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   '
   ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=" + TRIM$(List1.TEXT)
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
   '
   Call BLOQARCH(ARCHIPRE$) '\\\OT-06-009-WAR-21/03/06///
   Call TRALOCAL(ARCHIPRE$, "LIPREVEN")
   XII% = VENTABU%("LIPREVEN" + ATRI$)
   If XII% < 0 Then
     OPPX% = COMALTER%("Realmente CANCELA la Modificación\de Datos de esta Lista ?\\No, Conservarlos\Si, Cancelar")
     If OPPX% = 2 Then GoTo 90
   End If
   '
   If XII% >= 0 Then
     CADECI% = Val(Text3)
     If CADECI% < 0 Or CADECI% > 5 Then CADECI% = 2
     List2.Clear
     '
     FIN& = ULTREG&("!LIPREVEN")
     Screen.MousePointer = 11
     '
     HOII% = HOY(HOS$)
     If Val(Label4) = 3 Then
       For I& = 1 To FIN&
         NXX$ = REGLEIDO$("!LIPREVEN", I&)
         CI% = CVI(Left$(NXX$, 2))
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             PRELI# = CVD(Mid$(NXX$, 62, 8))
             RFF$ = RECFILT$(ARCHIPRE$, 1, MKI$(CI%))
             If Len(RFF$) >= 4 Then
               REANT& = CVS(RFF$)
               If REANT& > 0 Then
                 NYY$ = REGLEIDO$(ARCHIPRE$, REANT&)
                 If Abs(CVD(Mid$(NYY$, 62, 8)) - PRELI#) < 0.00005 Then
                   If TRIM$(Mid$(NYY$, 46, 16)) <> "" Then
                     GoTo 89
                   End If
                 End If
               End If
             End If
           End If
         End If
         Call REPLA(NXX$, HOS$, 46, 16)
         Call GRAREG("!LIPREVEN", NXX$, I&)
89     Next I&
     End If
     '
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        NXX$ = REGLEIDO$("!LIPREVEN", I&)
        CI% = CVI(Left$(NXX$, 2))
        MRX$ = Mid$(NXX$, 45, 1)
        If CI% > 0 Then
          If CI% <= NUMAS% Then
            PRELI# = CVD(Mid$(NXX$, 62, 8))
            PREX$ = "  " + TRIM$(CSTRING$(MKD$(PRELI#), 3, 12, CADECI%, 2))
            DE$ = DESCRIT0$(CI%)
            Call REPLA$(NXX$, DE$, 3, 42)
            Call REPLA(NXX$, PREX$ + "  ", 43 - Len(PREX$), 2 + Len(PREX$))
            Call REPLA(NXX$, MRX$, 44, 1)
            
            Call GRAREG(ARCHIPRE$, NXX$, I&)
          End If
        End If
     Next I&
     Call SETULTREG(ARCHIPRE$, FIN&)
     Call CIERRARCH(ARCHIPRE$)
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     Screen.MousePointer = 1
     '
     Call CARLISTA
   End If
   '
90 Call LIBARCH(ARCHIPRE$) '\\\OT-06-009-WAR-21/03/06///
   '
End Sub

Sub APREPOR()
    '
    Dim VALVIE(32767), VALNUE(32767)
    HOII% = HOY(HOS$)
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") > 0 Then
      If VERILLAVE% <> 1 Then Exit Sub
    End If
    '
    OPX% = COMALTER%("Puede Actualizar a Partir\de la Lista de Precios Actualmente Vigente\o de los Costos de Reposición para un Proveedor\\Lista Vigente\Costos de Reposición\Cancelar")
    Select Case OPX%
      Case 2
        Call APREPROV
        Exit Sub
      Case 3
        Exit Sub
    End Select
    '
    LISTACTIVA& = List1.ListIndex + 1
    X$ = REGLEIDO$("LISTAS", LISTACTIVA&)
    '
    RANGOAC% = 0
    If InStr(EPAC$, "DOSIVA") > 0 Then
      TTXX$ = "Opciones de Actualización por Porcentaje:\\Sólo Bombas\Sólo Repuestos\Toda la Lista"
      RANGOAC% = COMALTER%(TTXX$)
      If RANGOAC% < 1 Or RANGOAC% > 3 Then Exit Sub
      If RANGOAC% = 3 Then RANGOAC% = 0
    End If
    '
    Select Case RANGOAC%
      Case 0
        RACTU$ = "Lista Completa"
      Case 1
        RACTU$ = "Bombas Nacionales (Cod. 0...)"
      Case 2
        RACTU$ = "Repuestos (Cod. >= 1...)"
    End Select
    '
    Call DESHASCOD(cod1$, cod2$, CIKX1%, CIKX2%)
    If CIKX1% < 1 Or CIKX2% < 1 Then Exit Sub
    '
    VDEF$ = Left$(X$, 64) + MKS$(0)
    Call REPLA(VDEF$, MKI$(HOII%), 63, 2)
    CADECI% = Asc(Mid$(VDEF$, 49, 1))
    If CADECI% < 2 Or CADECI% > 4 Then CADECI% = 2
    Call REPLA(VDEF$, Chr$(CADECI%), 49, 1)
    '
5   OBX$ = OBJPLA$("APREVENPOR/Actualizacion por Porcentaje - " + RACTU$, VDEF$)
    If OBX$ = "" Then
      Exit Sub
    End If
    '
    PORAJU = CVS(Mid$(OBX$, 65, 4))
    If Abs(PORAJU) < 0.001 Then Exit Sub
    '
    If PORAJU < (-90) Or PORAJU > 99.99 Then
       Call MENSERR(24, "Porcentaje no Válido")
       GoTo 5
    End If
    '
    CADECI% = Asc(Mid$(OBX$, 49, 1))
    If CADECI% < 2 Or CADECI% > 4 Then
       Call MENSERR(24, "Decimales entre '2' y '4'")
       GoTo 5
    End If
    '
    LISNO% = Asc(Left$(OBX$, 1))
    ALISTA$ = TRIM$(Str$(LISNO%)): While Len(ALISTA$) < 3
                                     ALISTA$ = "0" + ALISTA$
                                   Wend
                                   '
    APRECIO$ = "LIPRE" + ALISTA$
    '
    Call BLOQARCH(APRECIO$) '\\\OT-06-009-WAR-21/03/06///
    '
    Call COPYSTRU("LIPREVEN", APRECIO$)
    REBLA$ = REGBLAN$("ACOSPOR")
    FIN& = ULTREG&(APRECIO$)
    JJ& = 0
    '
    Screen.MousePointer = 11
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      I0$ = REGLEIDO$(APRECIO$, U&)
      CIXI% = CVI(Left$(I0$, 2))
      CDXXX$ = CODEXT$(CIXI%)
      '
      PORAJX = PORAJU
      If RANGOAC% > 0 Then
        CDDX$ = Left$(CDXXX$, 1)
        If RANGOAC% = 1 Then
            If CDDX$ <> "0" Then PORAJX = 0
          ElseIf RANGOAC% = 2 Then
            If CDDX$ = "0" Then PORAJX = 0
        End If
      End If
      '
      If CDXXX$ < cod1$ Or CDXXX$ > cod2$ Then
        PORAJX = 0
      End If
      '
      COSVIE = CVD(Mid$(I0$, 62, 8))
      COSNUEX$ = CSTRING$(MKS$(COSVIE * (100 + PORAJX) / 100), 3, 12, CADECI%, 2)
      COSNUE = Val(COSNUEX$)
      ZZ$ = REBLA$
      Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
      Call REPLA(ZZ$, MKS$(COSVIE), 3, 4)
      Call REPLA(ZZ$, MKS$(COSNUE), 7, 4)
      JJ& = JJ& + 1
      Call GRAREG("!ACOSPOR", ZZ$, JJ&)
    Next U&
    '
    Call SETULTREG("!ACOSPOR", JJ&)
    Call CIERRARCH("!ACOSPOR")
    Screen.MousePointer = 1
    '
    ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Actualizacion de Precios por Porcentaje - " + RACTU$
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
    '
    VTB% = VENTABU%("ACOSPOR" + ATRI$)
    If VTB% >= 0 Then
      If COMALTER%("Desea Guardar los Nuevos Precios de Venta ?\\Cancelar\Guardar") = 2 Then
         '
         Screen.MousePointer = 11
         HOII% = HOY(HOS$)
         URE& = ULTREG&("!ACOSPOR")
         For U& = 1 To URE&
           Call TRACE(20, U&, URE&)
           X$ = REGLEIDO$("!ACOSPOR", U&)
           Y$ = REGLEIDO$(APRECIO$, U&)
           MRX$ = Mid$(Y$, 45, 1)
           COSNUEX$ = CSTRING$(Mid$(X$, 7, 4), 3, 12, CADECI%, 2)
           PRENUEVO# = Val(COSNUEX$)
           PREX$ = COSNUEX$
           '
           If Val(Label4) = 3 Then
             PREVIE# = CVD(Mid$(Y$, 62, 8))
             If Abs(PREVIE# - PRENUEVO#) < 0.00005 Or TRIM$(Mid$(Y$, 46, 16)) = "" Then
               Call REPLA(Y$, HOS$, 46, 16)
             End If
           End If
           '
           Call REPLA(Y$, MKD$(PRENUEVO#), 62, 8)
           Call REPLA(Y$, PREX$ + "  ", 43 - Len(PREX$), 2 + Len(PREX$))
           Call REPLA(Y$, MRX$, 44, 1)
           Call GRAREG(APRECIO$, Y$, U)
         Next U&
         '
         Call CIERRARCH(APRECIO$)
         Call BAJALDISCO
         Screen.MousePointer = 1
         '
         Call CARLISTA
         '
      End If
    End If
    '
    Call CIERRARCH(APRECIO$)
    Call LIBARCH(APRECIO$) '\\\OT-06-009-WAR-21/03/06///
    '
End Sub

Sub APREPROV()
    '
    LISTACTIVA& = List1.ListIndex + 1
    X$ = REGLEIDO$("LISTAS", LISTACTIVA&)
    '
    Call SELECHO("ACREDOR")
    VDEVU$ = VALACT1$("ACREDOR")
    NPRO% = Val(VDEVU$)
    If NPRO% < 1 Then Exit Sub
    '
    RACTU$ = RASOCLI$((-1) * NPRO%)
    VDEF$ = Left$(X$, 64) + MKS$(0)
    Call REPLA(VDEF$, MKI$(HOII%), 63, 2)
    CADECI% = Asc(Mid$(VDEF$, 49, 1))
    If CADECI% < 2 Or CADECI% > 4 Then CADECI% = 2
    Call REPLA(VDEF$, Chr$(CADECI%), 49, 1)
    '
5   OBX$ = OBJPLA$("APREVENPOR/Actualizacion por Porcentaje - " + RACTU$, VDEF$)
    If OBX$ = "" Then
      Exit Sub
    End If
    '
    PORAJU = CVS(Mid$(OBX$, 65, 4))
    'If Abs(PORAJU) < 0.001 Then Exit Sub
    '
    If PORAJU < (-90) Or PORAJU > 299.99 Then
       Call MENSERR(24, "Porcentaje no Válido")
       GoTo 5
    End If
    '
    CADECI% = Asc(Mid$(OBX$, 49, 1))
    If CADECI% < 2 Or CADECI% > 4 Then
       Call MENSERR(24, "Decimales entre '2' y '4'")
       GoTo 5
    End If
    '
    LISNO% = Asc(Left$(OBX$, 1))
    ALISTA$ = TRIM$(Str$(LISNO%)): While Len(ALISTA$) < 3
                                     ALISTA$ = "0" + ALISTA$
                                   Wend
                                   '
    APRECIO$ = "LIPRE" + ALISTA$
    '
    Call BLOQARCH(APRECIO$)
    '
    Call COPYSTRU("LIPREVEN", APRECIO$)
    REBLA$ = REGBLAN$("ACOSPOR")
    FIN& = ULTREG&(APRECIO$)
    JJ& = 0
    '
    Screen.MousePointer = 11
    PUCOD$ = ""
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      I0$ = REGLEIDO$(APRECIO$, U&)
      CIXI% = CVI(Left$(I0$, 2))
      CDXXX$ = CODEXT$(CIXI%)
      '
      COSVIE = CVD(Mid$(I0$, 62, 8))
      COSNUE = COSVIE
      ZZ$ = REBLA$
      Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
      Call REPLA(ZZ$, MKS$(COSVIE), 3, 4)
      Call REPLA(ZZ$, MKS$(COSNUE), 7, 4)
      JJ& = JJ& + 1
      Call GRAREG("!ACOSPOR", ZZ$, JJ&)
      PUCOD$ = PUCOD$ + MKI$(CIXI%)
    Next U&
    '
    AGREGAIT% = 0
    For CIXI% = 1 To NUMAS%
      If PROVHABI%(CIXI%) = NPRO% Then
        COSREP = COSPESOS(CIXI%)
        COSNUEX$ = CSTRING$(MKS$(COSREP * (100 + PORAJU) / 100), 3, 12, CADECI%, 2)
        COSNUE = Val(COSNUEX$)
        '
        For U& = 1 To JJ&
          If CVI(Mid$(PUCOD$, 2 * U& - 1, 2)) = CIXI% Then
            ZZ$ = REGLEIDO$("!ACOSPOR", U&)
            Call REPLA(ZZ$, MKS$(COSNUE), 7, 4)
            Call GRAREG("!ACOSPOR", ZZ$, U&)
            GoTo 29
          End If
        Next U&
        '
        ZZ$ = REBLA$
        Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
        Call REPLA(ZZ$, MKS$(0), 3, 4)
        Call REPLA(ZZ$, MKS$(COSNUE), 7, 4)
        JJ& = JJ& + 1
        Call GRAREG("!ACOSPOR", ZZ$, JJ&)
        PUCOD$ = PUCOD$ + MKI$(CIXI%)
        AGREGAIT% = 1
      End If
29  Next CIXI%
    '
    Call SETULTREG("!ACOSPOR", JJ&)
    Call CIERRARCH("!ACOSPOR")
    If AGREGAIT% > 0 Then
      Call COMUNI("Los Items Agregados Aparecen\al FINAL de la Lista.\ \Luego de Guardar los Nuevos Precios\Puede Realizar Ordenamiento Automatico.")
    End If
    Screen.MousePointer = 1
    '
    ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Actualizacion de Precios por Porcentaje - " + RACTU$
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
    '
    VTB% = VENTABU%("ACOSPOR" + ATRI$)
    If VTB% >= 0 Then
      If COMALTER%("Desea Guardar los Nuevos Precios de Venta ?\\Cancelar\Guardar") = 2 Then
         '
         Screen.MousePointer = 11
         HOII% = HOY(HOS$)
         URE& = ULTREG&("!ACOSPOR")
         UREA& = ULTREG&(APRECIO$)
         For U& = 1 To URE&
           Call TRACE(20, U&, URE&)
           X$ = REGLEIDO$("!ACOSPOR", U&)
           If U& <= UREA& Then
               Y$ = REGLEIDO$(APRECIO$, U&)
             Else
               CIXI% = CVI(Left$(X$, 2))
               Y$ = REGBLAN$(APRECIO$)
               Call REPLA(Y$, MKI$(CIXI%), 1, 2)
               Call REPLA(Y$, DESCRIT0$(CIXI%), 3, 42)
           End If
           MRX$ = Mid$(Y$, 45, 1)
           COSNUEX$ = CSTRING$(Mid$(X$, 7, 4), 3, 12, CADECI%, 2)
           PRENUEVO# = Val(COSNUEX$)
           PREX$ = COSNUEX$
           '
           If Val(Label4) = 3 Then
             PREVIE# = CVD(Mid$(Y$, 62, 8))
             If Abs(PREVIE# - PRENUEVO#) < 0.00005 Or TRIM$(Mid$(Y$, 46, 16)) = "" Then
               Call REPLA(Y$, HOS$, 46, 16)
             End If
           End If
           '
           Call REPLA(Y$, MKD$(PRENUEVO#), 62, 8)
           Call REPLA(Y$, PREX$ + "  ", 43 - Len(PREX$), 2 + Len(PREX$))
           Call REPLA(Y$, MRX$, 44, 1)
           Call GRAREG(APRECIO$, Y$, U&)
         Next U&
         '
         Call SETULTREG(APRECIO$, U&)
         Call CIERRARCH(APRECIO$)
         Call BAJALDISCO
         Screen.MousePointer = 1
         '
         Call CARLISTA
         If AGREGAIT% > 0 Then
           If COMALTER%("Desea Ordenar por Codigo Creciente ?\\No, Dejar\Si, Ordenar") = 2 Then
             Call Command12_Click
           End If
         End If
         '
      End If
    End If
    '
    Call BLOQARCH(APRECIO$)
    '
End Sub

Sub ALTACODIGO(VDEF$, OKX%)
   '
   Dim BUFERCOD As String * 2048
   '
10 OBX$ = OBJPLA$("ALTACODIGO", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   CODINUE$ = TRIM$(Mid$(OBX$, 77, 16))
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 10
   End If
   '
   If Len(CODINUE$) > 15 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que 15 Caracteres.")
      GoTo 10
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 10
   End If
   '
   If TRIM$(Mid$(OBX$, 5, 64)) = "" Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Descripcion")
      GoTo 10
   End If
   '
   TIMA% = Asc(Mid$(OBX$, 93, 1))
   If ECOARCH$("TATIMAT", Chr$(TIMA%)) = "" Then
      Call MENSERR(24, "Tipo de Material no Válido")
      GoTo 10
   End If
   '
   UNID$ = Mid$(OBX$, 109, 2)
   If ECOARCH$("UNIMED", UNID$) = "" Then
      Call MENSERR(24, "Unidad de Medida no Válida")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   '
   NMSI% = NUMAS% + 1
   REMAS& = NUMAS% + 2
   Call GRAREG("MASTER", OBX$, REMAS&)
   '
   X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(NMSI%), 1, 2)
   Call GRAREG("MASTER", X$, 1)
   '
   Call CIERRARCH("MASTER")
   Call BAJALDISCO
   Call REPUNCODIG
   '
   EL1$ = REGBLAN$("ECOMAST")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, Mid$(OBX$, 5, 64), 17, 46)
   Call REPLA(EL1$, MKI$(NMSI%), 63, 2)
   Call REGAPP("ECOMAST", EL1$)
   Call CIERRARCH("ECOMAST")
   Call FILESORT("ECOMAST", "", 1, 1, "ASC")
   Call BAJALDISCO
   '
   NVII1% = 1 + NUINV%
   EL1$ = REGLEIDO$("INVERT", 1)
   Call REPLA(EL1$, String$(16, 0), 1, 16)
   Call REPLA(EL1$, MKI$(NVII1%), 17, 2)
   Call GRAREG("INVERT", EL1$, 1)
   EL1$ = REGBLAN$("INVERT")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, MKI$(NMSI%), 17, 2)
   Call REGAPP("INVERT", EL1$)
   Call CIERRARCH("INVERT")
   Call FILESORT("INVERT", "", 1, 1, "ASC")
   Call BAJALDISCO
   '
   Call FRESHECHO("ECOMAST")
   Call FRESHECHO("MASTER")
   Call BAJALDISCO
   OKX% = 1
   '
99 Screen.MousePointer = 1

End Sub
'
Sub LISTAPRE()

2000    ' ************** impresion de listas de precios
        '
        LISTACTIVA& = List1.ListIndex + 1
        X$ = REGLEIDO$("LISTAS", LISTACTIVA&)
        TITU$ = TRIM$(Mid$(X$, 2, 46))
        '
        Call DESHASCOD(cod1$, cod2$, CIK1%, CIK2%)
        If CIK1% < 1 Or CIK2% < 1 Then Exit Sub
        '
2100    Call BLANARCH("LIPREVEN")
        Screen.MousePointer = 11
        '
        Call HEADERS("LIPREVEN", "")
        Call HEADERS("LIPREVEN", "Lista: " + TITU$)
        '
        ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
        While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
        LIPREAC$ = "LIPRE" + ARCHIPRE$
        FIN& = ULTREG&(LIPREAC$)
        '
        For I& = 1 To FIN&
          Call TRACE(20, I&, FIN&)
          NXX$ = REGLEIDO$(LIPREAC$, I&)
          CI% = CVI(Left$(NXX$, 2))
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              CDXX$ = CODEXT$(CI%)
              If CDXX$ >= cod1$ Then
                If CDXX$ <= cod2$ Then
                  RELIB% = Val(Mid$(NXX$, 45, 1))
                  '
                  For K% = 1 To RELIB%
                    Call REGAPP("LIPREVEN", MKI$(32767) + Mid$(REGBLAN$("LIPREVEN"), 3))
                  Next K%
                  Call REGAPP("LIPREVEN", NXX$)
                  '
                End If
              End If
            End If
          End If
        Next I&
        '
        Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        Call FINAL("*LILIPRE") '
        '
End Sub

Sub SORLIPRE()

    LISTACTIVA& = List1.ListIndex + 1
    X$ = REGLEIDO$("LISTAS", LISTACTIVA&)
    ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
    While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
    LIPREAC$ = "LIPRE" + ARCHIPRE$
    '
    '    Call DATLIPRE(X$)                                    ' datos de lista elegida
        '
    FIN& = ULTREG&(LIPREAC$)
    If FIN& < 1 Then
       Exit Sub
    End If
    Call CIERRARCH(LIPREAC$)
    LDT$ = LUDAT$
    '
    Call REGAPP(LIPREAC$, String$(128, 0))
    Call CIERRARCH("*.*")
    '
    Call BLOQARCH(LIPREAC$) '\\\OT-06-009-WAR-21/03/06///
    '
    Call TRALOCAL(LIPREAC$, "LIPREVEN")
    '
    Screen.MousePointer = 11
    SORTLIST.Clear
    '
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       NXX$ = REGLEIDO$(LIPREAC$, I&)
       CI% = CVI(Left$(NXX$, 2))
       CODE$ = CODEXT$(CI%)
       If CODE$ <> "" Then
          Z$ = Space$(24)
          Call REPLA(Z$, "A" + CODE$, 1, 18)
          Call REPLA(Z$, TRIM$(Str$(I&)), 19, 6)
          SORTLIST.AddItem Z$
       End If
    Next I&
    '
    FIN& = SORTLIST.ListCount
    '
    J& = 0
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       TESOR$ = SORTLIST.List(I& - 1)
       REG& = Val(Mid$(TESOR$, 19, 6))
       If REG& > 0 Then
          NXX$ = REGLEIDO$("!LIPREVEN", REG&)
          J& = J& + 1
          Call GRAREG(LIPREAC$, NXX$, J&)
       End If
    Next I&
    '
    Call SETULTREG&(LIPREAC$, J&)
    Call CIERRARCH("*.*")
    Call CARLISTA
    Screen.MousePointer = 1
    '
    Call LIBARCH(LIPREAC$) '\\\OT-06-009-WAR-21/03/06///
    '
    Call COMUNI("Ordenamiento Completo.")
    '
End Sub

Sub CARILISTAS()
   '
   List1.Clear
   For U& = 1 To ULTREG&("LISTAS")
     XX$ = REGLEIDO$("LISTAS", U&)
     List1.AddItem AJUSTD$(TRIM$(Str$(U&)), 2) + "-." + TRIM$(Mid$(XX$, 2, 46))
   Next U&
   List1.Refresh
   Call CIERRARCH("LISTAS")
   '
End Sub

Function VERILLAVE%()
   '
   VERILLAVE% = 0
10 OPX% = COMALTER%("Coloque Diskette Llave\\Continuar\Cancelar")
   If OPX% = 1 Then
      On Error GoTo 50
      SVL$ = SERIVOL$("A:")
      TECONTROL$ = ENCRIPTA$(TRIM$(UCase$(App.EXEName)) + SVL$)
      NX1% = FreeFile
      Dim TTX As String * 128
      Open "A:\" + TRIM$(UCase$(App.EXEName)) + ".KEY" For Random Access Read Write As #NX1% Len = 128
      Get #NX1%, 2, TTX$
      LTC% = Len(TECONTROL$)
VERILLAVE% = 1: GoTo 99
      If CVI(Mid$(TTX$, 18, 2)) = LTC% Then
        If Mid$(TTX$, 30, LTC%) = TECONTROL$ Then
          VERILLAVE% = 1
          GoTo 99
        End If
      End If
      Call COMUNI("Llave no Válida")
   End If
   GoTo 99
   '
50 Resume 51
51 GoTo 10
   '
99 End Function

Sub EDITLISTA(ARCHIPRE$)
   '
   ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=" + TRIM$(List1.TEXT)
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
   '
   XII% = VENTABU%("LIPREVEN" + ATRI$)
   If XII% >= 0 Then
     CADECI% = Val(Text3)
     If CADECI% < 0 Or CADECI% > 5 Then CADECI% = 2
     List2.Clear
     '
     FIN& = ULTREG&("!LIPREVEN")
     Screen.MousePointer = 11
     '
     HOII% = HOY(HOS$)
     If Val(Label4) = 3 Then
       For I& = 1 To FIN&
         NXX$ = REGLEIDO$("!LIPREVEN", I&)
         CI% = CVI(Left$(NXX$, 2))
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             PRELI# = CVD(Mid$(NXX$, 62, 8))
             RFF$ = RECFILT$(ARCHIPRE$, 1, MKI$(CI%))
             If Len(RFF$) >= 4 Then
               REANT& = CVS(RFF$)
               If REANT& > 0 Then
                 NYY$ = REGLEIDO$(ARCHIPRE$, REANT&)
                 If Abs(CVD(Mid$(NYY$, 62, 8)) - PRELI#) < 0.00005 Then
                   If TRIM$(Mid$(NYY$, 46, 16)) <> "" Then
                     GoTo 89
                   End If
                 End If
               End If
             End If
           End If
         End If
         Call REPLA(NXX$, HOS$, 46, 16)
         Call GRAREG("!LIPREVEN", NXX$, I&)
89     Next I&
     End If
     '
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        NXX$ = REGLEIDO$("!LIPREVEN", I&)
        CI% = CVI(Left$(NXX$, 2))
        MRX$ = Mid$(NXX$, 45, 1)
        If CI% > 0 Then
          If CI% <= NUMAS% Then
            PRELI# = CVD(Mid$(NXX$, 62, 8))
            PREX$ = "  " + TRIM$(CSTRING$(MKD$(PRELI#), 3, 12, CADECI%, 2))
            DE$ = DESCRIT0$(CI%)
            Call REPLA$(NXX$, DE$, 3, 42)
            Call REPLA(NXX$, PREX$ + "  ", 43 - Len(PREX$), 2 + Len(PREX$))
            Call REPLA(NXX$, MRX$, 44, 1)
            
            Call GRAREG(ARCHIPRE$, NXX$, I&)
          End If
        End If
     Next I&
     Call SETULTREG(ARCHIPRE$, FIN&)
     Call CIERRARCH(ARCHIPRE$)
     Screen.MousePointer = 1
     '
     Call CARLISTA
   End If

End Sub
