VERSION 5.00
Begin VB.Form OPORPASTA 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formularios de Comprobantes de Pago Standard"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Formularios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   210
      TabIndex        =   25
      Top             =   210
      Width           =   5160
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   42
         Text            =   " "
         Top             =   3150
         Width           =   1380
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
         Height          =   360
         Left            =   3780
         TabIndex        =   41
         Top             =   3150
         Width           =   175
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3675
         Width           =   3750
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   4095
         Picture         =   "OPORPASTA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3150
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   " Suprime  la  Impresión  al  Generar  el  Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   38
         Top             =   2415
         Width           =   4740
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   4095
         Picture         =   "OPORPASTA.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1365
         Width           =   855
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1890
         Width           =   3750
      End
      Begin VB.CommandButton Command6 
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
         Height          =   360
         Left            =   3780
         TabIndex        =   32
         Top             =   1365
         Width           =   175
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   31
         Text            =   " "
         Top             =   1365
         Width           =   1380
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   4095
         Picture         =   "OPORPASTA.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   3750
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
         Height          =   360
         Left            =   3780
         TabIndex        =   27
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   26
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   5145
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Retención Ing.Brutos:"
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
         Left            =   0
         TabIndex        =   43
         Top             =   3255
         Width           =   2325
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Liquidación Fondo Fijo:"
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
         Left            =   0
         TabIndex        =   35
         Top             =   1470
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de Pago:"
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
         Left            =   0
         TabIndex        =   30
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Configuración de Ordenes de Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4230
      Left            =   5460
      TabIndex        =   3
      Top             =   210
      Width           =   5160
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cartera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   210
         TabIndex        =   36
         Top             =   420
         Width           =   1650
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Utiliza Carte- ra de Valores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   105
            TabIndex        =   37
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ordenamiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   22
         Top             =   2835
         Width           =   1650
         Begin VB.OptionButton Option8 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Vencim."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   24
            Top             =   735
            Width           =   1275
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Emision"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   23
            Top             =   420
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Prefijo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   3675
         TabIndex        =   14
         Top             =   1575
         Width           =   1275
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   210
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "0001"
            Top             =   525
            Width           =   870
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Papel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   13
         Top             =   1575
         Width           =   3330
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   380
            Left            =   315
            TabIndex        =   19
            Top             =   735
            Width           =   2850
            Begin VB.OptionButton Option6 
               BackColor       =   &H00FFFFC0&
               Caption         =   "NewPage"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1575
               TabIndex        =   21
               Top             =   0
               Width           =   1275
            End
            Begin VB.OptionButton Option5 
               BackColor       =   &H00FFFFC0&
               Caption         =   "EndDoc"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Value           =   -1  'True
               Width           =   1275
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   435
            Left            =   315
            TabIndex        =   16
            Top             =   315
            Width           =   2850
            Begin VB.OptionButton Option3 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Formulario Continuo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   0
               TabIndex        =   18
               Top             =   0
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton Option4 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Hoja Suelta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1575
               TabIndex        =   17
               Top             =   0
               Width           =   1065
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Talonario en Uso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   1995
         TabIndex        =   9
         Top             =   2835
         Width           =   2955
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
            Height          =   360
            Left            =   2580
            TabIndex        =   11
            Top             =   525
            Width           =   175
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1425
            TabIndex        =   10
            Text            =   " "
            Top             =   525
            Width           =   1185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nueva Numeración"
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
            Left            =   165
            TabIndex        =   12
            Top             =   480
            Width           =   1170
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   1995
         TabIndex        =   6
         Top             =   420
         Width           =   1965
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pre-Impresa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   8
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Auto-Numera"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   7
            Top             =   630
            Width           =   1485
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   4095
         TabIndex        =   4
         Top             =   420
         Width           =   855
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   190
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "OP"
            Top             =   525
            Width           =   450
         End
      End
   End
   Begin VB.CommandButton Command7 
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
      Left            =   10750
      Picture         =   "OPORPASTA.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2205
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
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
      Left            =   10750
      Picture         =   "OPORPASTA.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1260
      Width           =   855
   End
   Begin VB.CommandButton command1 
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
      Left            =   10750
      Picture         =   "OPORPASTA.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10220
      Picture         =   "OPORPASTA.frx":107C
      Top             =   3910
      Width           =   2010
   End
End
Attribute VB_Name = "OPORPASTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text9.TEXT = VDEVU$
        Text8.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command2_Click()
    '
    FENUE$ = TRIM$(Text3.TEXT)
    If FENUE$ = "" Then
        FENUE$ = "01/01/80"
    End If
    If FECHANUM(FENUE$) < 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text3.SetFocus
        Exit Sub
    End If
    '
    Call GRACONTROL("", "FORPASTA", Text1.TEXT)
    Call GRACONTROL("", "FORLIFOF", Text6.TEXT)
    Call GRACONTROL("", "FORRETIB", Text9.TEXT)
    Call GRACONTROL("", "CODMORPA", "OP")    ' Text5.TEXT)
    '
    MONUME$ = "AUTONUME": If Option1.Value = True Then MONUME$ = "PREIMPRE"
    Call GRACONTROL("", "NUORPAST", MONUME$)
    '
    TIPAPEL$ = "FORMUCON": If Option4.Value = True Then TIPAPEL$ = "HOJASUEL"
    Call GRACONTROL("", "TIPAORPA", TIPAPEL$)
    '
    AVAHOJA$ = "ENDDOC": If Option6.Value = True Then AVAHOJA$ = "NEWPAGE"
    Call GRACONTROL("ORPASTA", "AVAHOJA", AVAHOJA$)
    '
    UTICHECA$ = "NO": If Check1.Value = 1 Then UTICHECA$ = "SI"
    Call GRACONTROL("ORPASTA", "UTICHECA", UTICHECA$)
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check2.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("ORPASTA", "SUPRINT", SUPRINT$)
    '
    MODORDE$ = "EMISION": MODORECI% = 1
    If Option8.Value = True Then
        MODORDE$ = "VENCIM"
        MODORECI% = 2
    End If
    Call GRACONTROL("ORPASTA", "MODORDE", MODORDE$)
    '
    Call GRACONTROL("", "FENUORPA", FENUE$)
    '
    PREFIREM$ = TRIM$(Text4.TEXT): If PREFIREM$ = "" Then PREFIREM$ = "0001"
    If PREFIREM$ <> "" Then
       If Len(PREFIREM$) <> 4 Then
          Call COMUNI("Prefijo no Válido")
          Text4.SetFocus
          Exit Sub
       End If
       For U& = 1 To 4
          UI% = Asc(Mid$(PREFIREM$, U&))
          If UI% < 48 Or UI% > 57 Then
             Call COMUNI("Prefijo no Válido")
             Text4.SetFocus
             Exit Sub
          End If
       Next U&
    End If
    Call GRACONTROL("", "PREFORPA", PREFIREM$)
    '
    Call CIERRARCH("*.*")
    '
    FORPASTA.List1.Clear
    FORPASTA.List2.Clear
    FORPASTA.Label4.Caption = ""
    FORPASTA.Label8.Caption = ""
    FORPASTA.Label10.Caption = ""
    FORPASTA.Label14.Caption = ""
    FORPASTA.Label18.Caption = ""
    FORPASTA.Label20.Caption = ""
    'FORPASTA.BOTNEXT.Enabled = True
    FORPASTA.Command5.Enabled = False
    '
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click()
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Command5_Click()
    Call CONECRUN("PRIFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
    If EXISTE%("PRIFORMS.EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

Private Sub Command6_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command8_Click()
    Call CONECRUN("PRIFORMS/" + TRIM$(Text6.TEXT), "Formularios de Impresion")
    If EXISTE%("PRIFORMS.EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

Private Sub Command9_Click()
    Call CONECRUN("PRIFORMS/" + TRIM$(Text9.TEXT), "Formularios de Impresion")
    If EXISTE%("PRIFORMS.EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

Private Sub Form_Load()
    '
    Text1.TEXT = CONTROL$("", "FORPASTA")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text6.TEXT = CONTROL$("", "FORLIFOF")
    Text7.TEXT = ECOARCH$("INDIFRM", Text6.TEXT)
    Text9.TEXT = CONTROL$("", "FORRETIB")
    Text8.TEXT = ECOARCH$("INDIFRM", Text9.TEXT)
    Text3.TEXT = CONTROL$("", "FENUORPA")
    '
    If CONTROL$("", "NUORPAST") = "AUTONUME" Then Option2.Value = True
    If CONTROL$("", "TIPAORPA") = "HOJASUEL" Then Option4.Value = True
    If CONTROL$("ORPASTA", "AVAHOJA") = "NEWPAGE" Then Option6.Value = True
    If CONTROL$("ORPASTA", "UTICHECA") = "SI" Then Check1.Value = 1
    If CONTROL$("ORPASTA", "SUPRINT") = "SI" Then Check2.Value = 1
    MODORECI% = 1: If CONTROL$("ORPASTA", "MODORDE") = "VENCIM" Then Option8.Value = True: MODORECI% = 2
    '
    PREFIREM$ = TRIM$(CONTROL$("", "PREFORPA"))
    Text4.TEXT = PREFIREM$
    Text5.TEXT = "OP"
    '
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub


Private Sub Text3_DBLCLICK()
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Text5_Click()
    Command2.SetFocus
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text7_CLICK()
    Text6.SetFocus
End Sub

Private Sub Text8_Click()
    Text9.SetFocus
End Sub

Private Sub Text9_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text9.TEXT = VDEVU$
        Text8.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
