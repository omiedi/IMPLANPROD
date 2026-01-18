VERSION 5.00
Begin VB.Form OPOCOMPRA 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compras y Expedición - Configuración y Formularios"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10860
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
   ScaleHeight     =   5175
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      Caption         =   "Cotiz."
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
      Left            =   6090
      Picture         =   "OPOCOMPRA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Actualizar Tabla de Tipo de Cambio"
      Top             =   460
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Param."
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
      Left            =   7035
      Picture         =   "OPOCOMPRA.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   460
      Width           =   855
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Orden de Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   210
      TabIndex        =   14
      Top             =   315
      Width           =   5160
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   210
         TabIndex        =   42
         Top             =   2730
         Width           =   1800
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Reemplazo Original"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   315
            TabIndex        =   43
            Top             =   315
            Width           =   1380
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   40
         Top             =   1470
         Width           =   1800
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Incluir las Especifica- ciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   315
            TabIndex        =   41
            Top             =   420
            Width           =   1380
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Documentos Anexos (Requiere Sistema de Calidad)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   210
         TabIndex        =   32
         Top             =   3675
         Width           =   4740
         Begin VB.CheckBox Check11 
            BackColor       =   &H00C0FFFF&
            Caption         =   "H.R.I.'s"
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
            Left            =   3570
            TabIndex        =   35
            Top             =   420
            Width           =   1065
         End
         Begin VB.CheckBox Check10 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Especificaciones"
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
            Left            =   1470
            TabIndex        =   34
            Top             =   420
            Width           =   1905
         End
         Begin VB.CheckBox Check9 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Planos"
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
            TabIndex        =   33
            Top             =   420
            Width           =   1065
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Costeo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   2100
         TabIndex        =   25
         Top             =   2310
         Width           =   2850
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
            Height          =   330
            Left            =   2415
            TabIndex        =   37
            Top             =   420
            Width           =   175
         End
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
            Height          =   330
            Left            =   1995
            TabIndex        =   36
            Text            =   " "
            Top             =   420
            Width           =   435
         End
         Begin VB.Label Label5 
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
            Left            =   210
            TabIndex        =   39
            Top             =   840
            Width           =   2430
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda:"
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
            Left            =   -420
            TabIndex        =   38
            Top             =   525
            Width           =   2325
         End
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
         TabIndex        =   24
         Text            =   " "
         Top             =   420
         Width           =   1380
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
         TabIndex        =   23
         Top             =   420
         Width           =   175
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
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
         Width           =   3750
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
         Picture         =   "OPOCOMPRA.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   420
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   2100
         TabIndex        =   17
         Top             =   1470
         Width           =   1920
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Auto"
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
            Left            =   1050
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   750
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Man"
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
            TabIndex        =   19
            Top             =   360
            Width           =   1380
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
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
         Height          =   720
         Left            =   4095
         TabIndex        =   15
         Top             =   1500
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
            TabIndex        =   16
            Text            =   "OC"
            Top             =   275
            Width           =   450
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   20
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Remito en Consignación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Left            =   5565
      TabIndex        =   3
      Top             =   1785
      Width           =   5160
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   210
         TabIndex        =   27
         Top             =   2310
         Width           =   4740
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
            Left            =   2205
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   " "
            Top             =   315
            Width           =   2385
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
            Left            =   1155
            TabIndex        =   29
            Text            =   " "
            Top             =   315
            Width           =   750
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
            Height          =   360
            Left            =   1890
            TabIndex        =   28
            Top             =   315
            Width           =   175
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Depósito:"
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
            Left            =   -1260
            TabIndex        =   31
            Top             =   420
            Width           =   2325
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
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
         Height          =   750
         Left            =   4095
         TabIndex        =   12
         Top             =   1470
         Width           =   855
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
            Left            =   190
            TabIndex        =   13
            Text            =   "RC"
            Top             =   275
            Width           =   450
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   210
         TabIndex        =   9
         Top             =   1470
         Width           =   3750
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Automatica"
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
            Left            =   2205
            TabIndex        =   11
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Manual"
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
            Left            =   525
            TabIndex        =   10
            Top             =   315
            Width           =   1380
         End
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
         Height          =   360
         Left            =   3780
         TabIndex        =   7
         Top             =   420
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
         Left            =   2415
         TabIndex        =   6
         Text            =   " "
         Top             =   420
         Width           =   1380
      End
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
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
         Width           =   3750
      End
      Begin VB.CommandButton Command6 
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
         Picture         =   "OPOCOMPRA.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   8
         Top             =   525
         Width           =   2325
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
      Left            =   7980
      Picture         =   "OPOCOMPRA.frx":0A68
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   460
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
      Left            =   8925
      Picture         =   "OPOCOMPRA.frx":0D72
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   460
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
      Left            =   9870
      Picture         =   "OPOCOMPRA.frx":107C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   460
      Width           =   855
   End
End
Attribute VB_Name = "OPOCOMPRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
    Call SELECHO("TAMONED")
    VDEVU$ = VALACT1$("TAMONED")
    If VDEVU$ <> "" Then
        Text9.TEXT = VDEVU$
        Label5 = VALACT2$("TAMONED")
    End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CARTAMONED
   Command11.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORMOCOM", Text1.TEXT)
    Call GRACONTROL("", "FOREMICO", Text3.TEXT)
    Call GRACONTROL("", "CODMOCOM", Text5.TEXT)
    Call GRACONTROL("", "CODRECON", Text6.TEXT)
    Call GRACONTROL("", "DEPOCONS", Text7.TEXT)
    Call GRACONTROL("OCOMPRA", "MONECOS", Text9.TEXT)
    '
    MONUME$ = "AUTO": If Option1.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMEOCOM", MONUME$)
    MONUME$ = "AUTO": If Option3.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMERECO", MONUME$)
    '
    MONUME$ = "NO"
    If Check9.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "PRIPLANO", MONUME$)
    '
    MONUME$ = "NO"
    If Check10.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "PRIESPEC", MONUME$)
    '
    MONUME$ = "NO"
    If Check11.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "PRIHRINS", MONUME$)
    '
    MONUME$ = "NO"
    If Check1.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "INCLUSPE", MONUME$)
    '
    MONUME$ = "NO"
    If Check2.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "REPLAORI", MONUME$)
    '
    Call CIERRARCH("*.*")
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
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command3.Enabled = False
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text3.TEXT), "Formularios de Impresion")
    Command3.Enabled = True
End Sub

Private Sub Command8_Click()
    Call CARDATCENCALC
End Sub

Private Sub Command9_Click()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Text8.TEXT = VALACT2$("TADEPOSI")
    End If
End Sub

Private Sub Form_Load()
    '
    J& = 0
    For U& = 1 To ULTREG&("SUCURSAL")
        X$ = REGLEIDO$("SUCURSAL", U&)
        CODEPO% = Asc(Mid$(X$, 64, 1))
        If CODEPO% > 0 Then
           If CODEPO% <> 99 Then
              If J& < 62 Then
                 J& = J& + 1
                 Y$ = REGBLAN$("TADEPOSI")
                 Call REPLA(Y$, Chr$(CODEPO%), 1, 1)
                 Call REPLA(Y$, Left$(X$, 63), 2, 63)
                 Call GRAREG("TADEPOSI", Y$, J&)
              End If
           End If
        End If
    Next U&
    Call SETULTREG("TADEPOSI", J&)
    Call CIERRARCH("SUCURSAL")
    Call CIERRARCH("TADEPOSI")
    Screen.MousePointer = 1
    '
    Text7.TEXT = CONTROL$("", "DEPOCONS")
    Text8.TEXT = ECOARCH$("TADEPOSI", Chr$(Val(Text7.TEXT)))
    '
    Text9.TEXT = CONTROL$("OCOMPRA", "MONECOS")
    Label5 = ECOARCH$("TAMONED", Chr$(Val(Text9.TEXT)))
    '
    Text1.TEXT = CONTROL$("", "FORMOCOM")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text3.TEXT = CONTROL$("", "FOREMICO")
    Text4.TEXT = ECOARCH$("INDIFRM", Text3.TEXT)
    '
    If CONTROL$("", "NUMEOCOM") = "MANUAL" Then Option1.Value = True
    If CONTROL$("", "NUMERECO") = "MANUAL" Then Option3.Value = True
    '
    CMO$ = TRIM$(CONTROL$("", "CODMOCOM")): If CMO$ <> "" Then Text5.TEXT = CMO$
    CMO$ = TRIM$(CONTROL$("", "CODRECON")): If CMO$ <> "" Then Text6.TEXT = CMO$
    '
    If CONTROL$("OCOMPRA", "PRIPLANO") = "SI" Then Check9.Value = 1
    If CONTROL$("OCOMPRA", "PRIESPEC") = "SI" Then Check10.Value = 1
    If CONTROL$("OCOMPRA", "PRIHRINS") = "SI" Then Check11.Value = 1
    If CONTROL$("OCOMPRA", "INCLUSPE") = "SI" Then Check1.Value = 1
    If CONTROL$("OCOMPRA", "REPLAORI") = "SI" Then Check2.Value = 1
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

Private Sub Text4_CLICK()
    Text2.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text7_DblClick()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Text8.TEXT = VALACT2$("TADEPOSI")
    End If
End Sub

Private Sub Text9_DblClick()
    Call SELECHO("TAMONED")
    VDEVU$ = VALACT1$("TAMONED")
    If VDEVU$ <> "" Then
        Text9.TEXT = VDEVU$
        Label5 = VALACT2$("TAMONED")
    End If
End Sub

Sub CARTAMONED()
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      JU& = 0
      For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         If TRIM$(Mid$(X$, 2, 27)) <> "" Then
           JU& = JU& + 1
           Call REPLA(X$, Chr$(JU&), 1, 1)
           Call GRAREG("TAMONED", X$, JU&)
         End If
      Next U&
      Call SETULTREG("TAMONED", JU&)
   End If
   Call CIERRARCH("TAMONED")
End Sub


