VERSION 5.00
Begin VB.Form STPIMP05 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0EEFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Impresora"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5745
   FillStyle       =   0  'Solid
   Icon            =   "STPIMP05.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
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
      Height          =   730
      Left            =   4830
      Picture         =   "STPIMP05.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Aprueba y Emite Listado"
      Top             =   1330
      Width           =   705
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Hoja Nro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   3360
      TabIndex        =   5
      Top             =   1260
      Width           =   1380
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   525
         TabIndex        =   6
         Text            =   " 1"
         Top             =   315
         Width           =   645
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Orientación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   210
      TabIndex        =   2
      Top             =   1260
      Width           =   3060
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Retrato"
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
         Index           =   0
         Left            =   315
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Paisaje"
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
         Index           =   1
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Impresora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   5340
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   315
         TabIndex        =   1
         Text            =   " "
         Top             =   420
         Width           =   4530
      End
   End
   Begin VB.Label Label1 
      Height          =   540
      Left            =   210
      TabIndex        =   8
      Top             =   2310
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "STPIMP05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IMPREDETER As String
Dim PORIAN%
Dim MODOSAL%

Sub Command1_Click()
    '
    If TRIM$(IMPREDETER$) <> TRIM$(Combo1.TEXT) Then
        Printer.TrackDefault = False
        NPPXX$ = TRIM$(Combo1.TEXT)
        For Each PRX9 In Printers
          If PRX9.DeviceName = NPPXX$ Then
            On Error Resume Next
            Set Printer = PRX9
            On Error GoTo 0
            GoTo 10
          End If
        Next PRX9
        '
10      If Printer.DeviceName <> Combo1.TEXT Then
          MENSA$ = "Imposible Aplicar Impresora Seleccionada:\"
          MENSA$ = MENSA$ + "'" + TRIM$(Combo1.TEXT) + "'\   \"
          MENSA$ = MENSA$ + "FLEXOFT ha Restaurado la Impresora Predeterminada:\"
          MENSA$ = MENSA$ + "'" + TRIM$(IMPREDETER$) + "'\  \"
          MENSA$ = MENSA$ + "Para Utilizar la Elegida, Cambie la Impresora\"
          MENSA$ = MENSA$ + "Predeterminada para Windows."
          Combo1.TEXT = IMPREDETER$
          Call COMUNI(MENSA$)
          Combo1.SetFocus
          Exit Sub
        End If
    End If
    '
20  On Error Resume Next
    If Option2(1).Value = True Then
          Printer.Orientation = 2
        Else
          Printer.Orientation = 1
    End If
    '
    Label1 = "OK"
    Hide
    Exit Sub
    '
199 ERNU$ = TRIM$(Str$(Err.Number))
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Impresión de\"
    MENSA$ = MENSA$ + "este Listado.\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Impresora Apagada o fuera de Línea.\"
    MENSA$ = MENSA$ + "   2.- Falta Papel o Papel Atorado.\"
    MENSA$ = MENSA$ + "   3.- Problemas en Cable de Señal de Impresora.\"
    MENSA$ = MENSA$ + "   4.- Problemas de Red - No hay Acceso a Impre-\"
    MENSA$ = MENSA$ + "        sora Remota.\"
    MENSA$ = MENSA$ + "   5.- Impresora de Matriz de Puntos Configurada\"
    MENSA$ = MENSA$ + "        para 'Calidad de Imprenta'. Corrija la Con-\"
    MENSA$ = MENSA$ + "        figuración para 'Borrador'.\"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    C$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(C$, 53, "Error de Impresión de Listado")
    If RESPU% = 4 Then
        Resume
    End If
    Call CIERRARCH("*.*")
    Call FINAL("")
    Close: End
    '
End Sub
'
Private Sub Form_Load()
    Refresh
    '
    PORIAN = Printer.Orientation
    Printer.TrackDefault = False
    '
    Combo1.Clear
    Dim PRX As Printer
    For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
    Next
    Combo1.TEXT = Printer.DeviceName
    IMPREDETER$ = Printer.DeviceName
    '
End Sub
'
