VERSION 5.00
Begin VB.Form WINDEMO1 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema Preconfigurado FLEXOFT - Demostración Interactiva"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox LISECU 
      Height          =   1260
      Left            =   6930
      TabIndex        =   25
      Top             =   1575
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Utilities"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   6930
      TabIndex        =   20
      Top             =   6405
      Visible         =   0   'False
      Width           =   4635
      Begin VB.CommandButton Command21 
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
         Left            =   315
         Picture         =   "WINDEMO1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Ayuda en Línea"
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Qwery"
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
         Left            =   2415
         Picture         =   "WINDEMO1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Consultas y Listados"
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command22 
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
         Left            =   1365
         Picture         =   "WINDEMO1.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Mail"
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
         Left            =   3465
         Picture         =   "WINDEMO1.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Correo Interno FLEXOFT"
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "SIM - Sistema de Manufactura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   6930
      TabIndex        =   11
      Top             =   3465
      Visible         =   0   'False
      Width           =   4635
      Begin VB.CommandButton Command11 
         Caption         =   "Files"
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
         Left            =   315
         Picture         =   "WINDEMO1.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Actualización y Consulta de Archivos Maestros"
         Top             =   525
         Width           =   855
      End
      Begin VB.CommandButton Command13 
         Caption         =   "MRP-I"
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
         Left            =   2415
         Picture         =   "WINDEMO1.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Planificación de Necesidades de Materiales"
         Top             =   525
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Stock"
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
         Left            =   1365
         Picture         =   "WINDEMO1.frx":11B4
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Control y Movimientos de Stocks"
         Top             =   525
         Width           =   855
      End
      Begin VB.CommandButton Command14 
         Caption         =   "MRP-II"
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
         Left            =   3465
         Picture         =   "WINDEMO1.frx":14BE
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Programación de Carga de Máquinas y Equipos"
         Top             =   525
         Width           =   855
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Factory"
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
         Left            =   315
         Picture         =   "WINDEMO1.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Ordenes de Trabajo y Fabricación"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Finan"
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
         Left            =   2415
         Picture         =   "WINDEMO1.frx":1AD2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Planeamiento y Control Financiero"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command16 
         Caption         =   "News"
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
         Left            =   1365
         Picture         =   "WINDEMO1.frx":1DDC
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Informes de Producción"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Presup"
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
         Left            =   3465
         Picture         =   "WINDEMO1.frx":221E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Costos y Presupuestación de Producto"
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Terminar"
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
      Left            =   10710
      Picture         =   "WINDEMO1.frx":2528
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   315
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5250
      Top             =   1050
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   790
      Left            =   1050
      TabIndex        =   2
      Top             =   450
      Visible         =   0   'False
      Width           =   2640
      Begin VB.Image Image9 
         BorderStyle     =   1  'Fixed Single
         Height          =   915
         Left            =   -105
         Picture         =   "WINDEMO1.frx":2672
         Stretch         =   -1  'True
         Top             =   -50
         Width           =   10815
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   3810
      Top             =   5955
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   945
      Top             =   5565
   End
   Begin VB.Label MODULO 
      BackStyle       =   0  'Transparent
      Caption         =   "y mucho más ..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Index           =   6
      Left            =   9135
      TabIndex        =   9
      Top             =   7560
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.Label MODULO 
      BackStyle       =   0  'Transparent
      Caption         =   "- Planeamiento Financiero"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Index           =   5
      Left            =   5565
      TabIndex        =   8
      Top             =   6720
      Visible         =   0   'False
      Width           =   4950
   End
   Begin VB.Label MODULO 
      BackStyle       =   0  'Transparent
      Caption         =   "- Presupuestación de Producto"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Index           =   4
      Left            =   5565
      TabIndex        =   7
      Top             =   6090
      Visible         =   0   'False
      Width           =   4950
   End
   Begin VB.Label MODULO 
      BackStyle       =   0  'Transparent
      Caption         =   "- Rastreabilidad según ISO 9000"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Index           =   3
      Left            =   5565
      TabIndex        =   6
      Top             =   5460
      Visible         =   0   'False
      Width           =   6420
   End
   Begin VB.Label MODULO 
      BackStyle       =   0  'Transparent
      Caption         =   "- Carga de Máquinas (MRP-II)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Index           =   2
      Left            =   5565
      TabIndex        =   5
      Top             =   4830
      Visible         =   0   'False
      Width           =   6420
   End
   Begin VB.Label MODULO 
      BackStyle       =   0  'Transparent
      Caption         =   "- Programación de Materiales (MRP-I) "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Index           =   1
      Left            =   5565
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   6945
   End
   Begin VB.Label MODULO 
      BackStyle       =   0  'Transparent
      Caption         =   "- Control de Stocks"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Index           =   0
      Left            =   5565
      TabIndex        =   3
      Top             =   3570
      Visible         =   0   'False
      Width           =   4950
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "al Servicio de la Producción "
      BeginProperty Font 
         Name            =   "Lucida Casual"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1275
      Left            =   525
      TabIndex        =   1
      Top             =   3045
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Un Universo de Herramientas Informáticas "
      BeginProperty Font 
         Name            =   "Lucida Casual"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1590
      Left            =   630
      TabIndex        =   0
      Top             =   1470
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   630
      Picture         =   "WINDEMO1.frx":157D8
      Top             =   5355
      Width           =   765
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   1410
      Picture         =   "WINDEMO1.frx":1772E
      Top             =   5355
      Width           =   765
   End
   Begin VB.Image Image3 
      Height          =   765
      Left            =   2190
      Picture         =   "WINDEMO1.frx":19684
      Top             =   5355
      Width           =   765
   End
   Begin VB.Image Image4 
      Height          =   765
      Left            =   2970
      Picture         =   "WINDEMO1.frx":1B5DA
      Top             =   5355
      Width           =   765
   End
   Begin VB.Image Image5 
      Height          =   765
      Left            =   630
      Picture         =   "WINDEMO1.frx":1D530
      Top             =   6195
      Width           =   765
   End
   Begin VB.Image Image6 
      Height          =   765
      Left            =   1410
      Picture         =   "WINDEMO1.frx":1F486
      Top             =   6195
      Width           =   765
   End
   Begin VB.Image Image7 
      Height          =   765
      Left            =   2190
      Picture         =   "WINDEMO1.frx":213DC
      Top             =   6195
      Width           =   765
   End
   Begin VB.Image Image8 
      Height          =   765
      Left            =   2970
      Picture         =   "WINDEMO1.frx":23332
      Top             =   6195
      Width           =   765
   End
End
Attribute VB_Name = "WINDEMO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Shell "DEMO027", 1
    Hide
    Close
    End
End Sub

Private Sub Command11_Click()
    Call SECUDEMO("ARCHIMA")
End Sub

Private Sub Command12_Click()
    Call SECUDEMO("STOCK")
End Sub

Private Sub Command13_Click()
'    Call SECUDEMO("PROGMAT")
End Sub

Private Sub Timer1_Timer()
    '
    Static VECES%
    '
    If YAEXTRA% < 1 Then
        '
        VECES% = VECES% + 1
        If VECES > 6 Then
            Frame5.Visible = True
        End If
        '
        If VECES% > 12 Then
            Label1.Visible = True
        End If
        '
        If VECES% > 18 Then
            Label2.Visible = True
        End If
        '
        If VECES% > 23 Then
            If VERIPAN% < 1 Then
              If Screen.Width < 12000 Then
                ANCHOPAN.Show
              End If
              VERIPAN% = 1
            End If
        End If
        '
        idx = Int((VECES% - 24) / 5)
        If idx >= 0 Then
           If idx <= 6 Then
              MODULO(idx).Visible = True
           End If
        End If
        
        If VECES% < 70 Then Exit Sub
        '
        Call EXTRACT
        EXTRA% = 1
        Call SECUDEMO("INTRODU")
        EXTRA% = 0
        Exit Sub
        '
    End If
    '
    NX% = FreeFile
    Open "COMMDEMO" For Input As #NX%
    Line Input #NX%, A$
    Close #NX%
    COMMA$ = LTrim$(RTrim$(UCase$(A$)))
    '
    If COMMA$ = "NEXT" Then
        '
        Timer2.Enabled = False
        Image1.Visible = False
        Image2.Visible = False
        Image3.Visible = False
        Image4.Visible = False
        Image5.Visible = False
        Image6.Visible = False
        Image7.Visible = False
        Image8.Visible = False
        '
        For K% = 0 To 6
           MODULO(K%).Visible = False
        Next K%
        '
        Command1.Visible = True
        Frame2.Visible = True
        Frame3.Visible = True
        Timer3.Enabled = True
        '
    End If
    '
End Sub
'
Private Sub Form_Load()
    '
    Image1.Visible = False
    Image1.Stretch = True
    Image2.Visible = False
    Image2.Stretch = True
    Image3.Visible = False
    Image3.Stretch = True
    Image4.Visible = False
    Image4.Stretch = True
    Image5.Visible = False
    Image5.Stretch = True
    Image6.Visible = False
    Image6.Stretch = True
    Image7.Visible = False
    Image7.Stretch = True
    Image8.Visible = False
    Image8.Stretch = True
    '
    Timer3.Enabled = False
    YAEXTRA% = 0
    VERIPAN% = 0
    Randomize
    '
End Sub

Private Sub Timer2_Timer()
    '
    Static VXX, px, py, tamano, INCREX, INCREY
    If VXX = 0 Then
        tamano = 240
        px = (ScaleWidth - tamano) / 2
        py = (ScaleHeight - tamano) / 2
        INCREX = 0.5 - Rnd
        INCREY = 0.5 - Rnd
        While (INCREX ^ 2 + INCREY ^ 2) ^ 0.5 < 0.36
        
            INCREX = INCREX * 1.2
            INCREY = INCREY * 1.2
        Wend
        signo = 1
    End If
    '
    VXX = VXX + 1
    If VXX = 9 Then
        VXX = 1
    End If
    '
    px = px + 100 * INCREX * tamano / 300
    py = py + 100 * INCREY * tamano / 300
    
    tamano = tamano + 40
    If px + tamano < 0 Then
        VXX = 0: Exit Sub
    End If
    '
    If px > ScaleWidth Then
        VXX = 0: Exit Sub
    End If
    '
    If py + tamano < 0 Then
        VXX = 0: Exit Sub
    End If
    '
    If py > ScaleHeight Then
        VXX = 0: Exit Sub
    End If
    '
    If tamano >= ScaleWidth Then VXX = 0
    '
    Select Case VXX
        '
        Case 1
            Image8.Visible = False
            Image1.Height = tamano
            Image1.Width = tamano
            Image1.Move px, py
            Image1.Visible = True
        Case 2
            Image1.Visible = False
            Image2.Height = tamano
            Image2.Width = tamano
            Image2.Move px, py
            Image2.Visible = True
        Case 3
            Image2.Visible = False
            Image3.Height = tamano
            Image3.Width = tamano
            Image3.Move px, py
            Image3.Visible = True
        Case 4
            Image3.Visible = False
            Image4.Height = tamano
            Image4.Width = tamano
            Image4.Move px, py
            Image4.Visible = True
        Case 5
            Image4.Visible = False
            Image5.Height = tamano
            Image5.Width = tamano
            Image5.Move px, py
            Image5.Visible = True
        Case 6
            Image5.Visible = False
            Image6.Height = tamano
            Image6.Width = tamano
            Image6.Move px, py
            Image6.Visible = True
        Case 7
            Image6.Visible = False
            Image7.Height = tamano
            Image7.Width = tamano
            Image7.Move px, py
            Image7.Visible = True
        Case 8
            Image7.Visible = False
            Image8.Height = tamano
            Image8.Width = tamano
            Image8.Move px, py
            Image8.Visible = True
    End Select
    '
End Sub

Private Sub Timer3_Timer()
    '
    Screen.MousePointer = 1
    Timer1.Enabled = False
    'Hide
    '
    COMMA$ = COMMDEM$
    If COMMA$ = "NEXT" Then
         CTX% = CTX% + 1
         Call PRESDEMO(CTX%)
         Exit Sub
       ElseIf COMMA$ = "PREV" Then
         CTX% = CTX% - 1: If CTX% < 1 Then CTX% = 1
         Call PRESDEMO(CTX%)
       ElseIf COMMA$ = "HOME" Then
         CTX% = 1
         Call PRESDEMO(CTX%)
       ElseIf COMMA$ = "END" Then
         CTX% = CAPAN% + 1
         Call PRESDEMO(CTX%)
    End If
    '
End Sub
