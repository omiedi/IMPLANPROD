VERSION 5.00
Begin VB.Form INSTRUING 
   BackColor       =   &H00E0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instrucciones de Ingenieria"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   520
      Left            =   5565
      TabIndex        =   29
      Top             =   1440
      Width           =   5055
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   1
         Left            =   1050
         TabIndex        =   36
         Top             =   105
         Width           =   2670
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   1
         Left            =   3795
         TabIndex        =   35
         Top             =   105
         Width           =   990
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
         Height          =   280
         Index           =   1
         Left            =   4770
         Picture         =   "INSTRUING.frx":0000
         TabIndex        =   34
         Top             =   105
         Width           =   175
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Aprobó:"
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
         Index           =   3
         Left            =   0
         TabIndex        =   37
         Top             =   105
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0C0&
      Caption         =   "Hoja de Actualizacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   105
      TabIndex        =   17
      Top             =   2100
      Width           =   10620
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   105
         TabIndex        =   44
         Top             =   315
         Width           =   10330
      End
      Begin VB.TextBox Text8 
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
         Left            =   1950
         TabIndex        =   40
         Top             =   1785
         Width           =   540
      End
      Begin VB.CommandButton Command10 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   10160
         Picture         =   "INSTRUING.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1575
         Width           =   300
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2625
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   1575
         Width           =   7540
      End
      Begin VB.TextBox Text7 
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
         Left            =   1290
         TabIndex        =   39
         Top             =   1785
         Width           =   540
      End
      Begin VB.TextBox Text6 
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
         Left            =   105
         TabIndex        =   38
         Top             =   1785
         Width           =   1065
      End
      Begin VB.Line Line6 
         X1              =   10185
         X2              =   10460
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line5 
         X1              =   2625
         X2              =   10460
         Y1              =   2110
         Y2              =   2110
      End
      Begin VB.Line Line4 
         X1              =   105
         X2              =   10450
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:         Hoja:    Mod:"
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
         TabIndex        =   43
         Top             =   1575
         Width           =   2430
      End
      Begin VB.Line Line2 
         X1              =   10440
         X2              =   10440
         Y1              =   315
         Y2              =   2100
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0C0&
      Caption         =   "Procedimiento - Encabezado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   105
      TabIndex        =   6
      Top             =   105
      Width           =   8310
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
         Height          =   280
         Index           =   0
         Left            =   5145
         Picture         =   "INSTRUING.frx":0614
         TabIndex        =   30
         Top             =   1440
         Width           =   175
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   0
         Left            =   1575
         TabIndex        =   32
         Top             =   1440
         Width           =   2565
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   0
         Left            =   4200
         TabIndex        =   31
         Top             =   1440
         Width           =   990
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Preparó:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   0
         Left            =   3360
         TabIndex        =   27
         Top             =   2310
         Width           =   4845
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1575
         TabIndex        =   26
         Top             =   840
         Width           =   6555
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4935
         TabIndex        =   25
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Preparó:"
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
         Index           =   2
         Left            =   525
         TabIndex        =   33
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
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
         Left            =   480
         TabIndex        =   24
         Top             =   500
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
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
         Left            =   105
         TabIndex        =   23
         Top             =   900
         Width           =   1380
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7380
         TabIndex        =   22
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "C.Interno:"
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
         Left            =   6300
         TabIndex        =   21
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label37 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1575
         TabIndex        =   20
         Top             =   420
         Width           =   2115
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Rev:"
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
         Left            =   3885
         TabIndex        =   19
         Top             =   525
         Width           =   1005
      End
   End
   Begin VB.Frame MARCOFICHA 
      BackColor       =   &H00E0E0C0&
      Caption         =   "Hoja de Procedimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Index           =   1
      Left            =   105
      TabIndex        =   8
      Top             =   4410
      Width           =   11670
      Begin VB.CommandButton Command7 
         Height          =   260
         Left            =   105
         TabIndex        =   11
         Top             =   3390
         Width           =   6600
      End
      Begin VB.TextBox Text28 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   735
         Width           =   11430
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Cortar"
         Height          =   260
         Left            =   6720
         TabIndex        =   15
         Top             =   3390
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Copiar"
         Height          =   260
         Left            =   7845
         TabIndex        =   14
         Top             =   3390
         Width           =   1170
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pegar"
         Height          =   260
         Left            =   9000
         TabIndex        =   13
         Top             =   3390
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Left            =   10260
         TabIndex        =   12
         Top             =   3390
         Width           =   1290
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   105
         ScaleHeight     =   375
         ScaleWidth      =   11370
         TabIndex        =   9
         Top             =   315
         Width           =   11430
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Instrucciones del Procedimiento"
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
            TabIndex        =   10
            Top             =   105
            Width           =   9885
         End
      End
      Begin VB.Line Line1 
         X1              =   105
         X2              =   11550
         Y1              =   3375
         Y2              =   3375
      End
      Begin VB.Line Line3 
         X1              =   11530
         X2              =   11530
         Y1              =   315
         Y2              =   3650
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0C0&
      Caption         =   "A-B-M-L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   10815
      TabIndex        =   3
      Top             =   1330
      Width           =   990
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
         Height          =   750
         Left            =   105
         Picture         =   "INSTRUING.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Alta de Nueva Hoja de Procedimientos"
         Top             =   300
         Width           =   750
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Print"
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
         Left            =   105
         Picture         =   "INSTRUING.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprime Hoja de Procedimientos"
         Top             =   1050
         Width           =   750
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0C0&
      Caption         =   "General"
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
      Left            =   8505
      TabIndex        =   0
      Top             =   105
      Width           =   3300
      Begin VB.CommandButton Command1 
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
         Height          =   750
         Left            =   105
         Picture         =   "INSTRUING.frx":10D2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton Command4 
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
         Height          =   750
         Left            =   1680
         Picture         =   "INSTRUING.frx":13DC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Ayuda en Línea"
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Restore"
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
         Left            =   840
         Picture         =   "INSTRUING.frx":16E6
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Restablece los Valores Anteriores"
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton Command3 
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
         Height          =   750
         Left            =   2415
         Picture         =   "INSTRUING.frx":19F0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cierra la Hoja de Procedimientos"
         Top             =   315
         Width           =   750
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0C0&
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
      Height          =   690
      Left            =   105
      TabIndex        =   28
      Top             =   1300
      Width           =   10620
   End
   Begin VB.Label DOCUORI 
      Height          =   225
      Left            =   10815
      TabIndex        =   46
      Top             =   3465
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label COMMLAB 
      Height          =   225
      Left            =   10815
      TabIndex        =   45
      Top             =   3255
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10340
      Picture         =   "INSTRUING.frx":1B3A
      Top             =   3800
      Width           =   2010
   End
End
Attribute VB_Name = "INSTRUING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   REGI& = Val(Label28)
   CIIINSTRU% = REGI&
   '
   Call BLOQARCH("INSTRUCI")
   XX$ = REGLEIDO$("INSTRUCI", REGI&)
   Call REPLA(XX$, AJUSTI$(Text3, 48), 17, 48)
   Call REPLA(XX$, AJUSTI$(Text1, 4), 65, 4)
   Call REPLA(XX$, AJUSTI$(Text4(0), 24), 71, 24)
   FEPREPA% = FECHANUM(Text5(0))
   Call REPLA(XX$, MKI$(FEPREPA%), 95, 2)
   Call REPLA(XX$, AJUSTI$(Text4(0), 24), 97, 24)
   FEAPROA% = FECHANUM(Text5(1))
   Call REPLA(XX$, MKI$(FEAPROA%), 121, 2)
   Call GRAREG("INSTRUCI", XX$, REGI&)
   Call CIERRARCH("INSTRUCI")
   '
   Call GRAINSTRUC
   Call GRAHOJAC
   '
End Sub

Private Sub Command10_Click()
   FEMODI% = FECHANUM(Text6.TEXT)
   If FEMODI% < 1 Then
      Call MENSERR(24, "Fecha no Válida")
      Text6.SetFocus
      Exit Sub
   End If
   '
   If Val(Text7) < 1 Or Val(Text7) > 9999 Then
      Call MENSERR(24, "Hoja no Válida")
      Text7.SetFocus
      Exit Sub
   End If
   hojanu% = Val(Text7)
   '
   If Val(Text8) < 1 Or Val(Text8) > 9999 Then
      Call MENSERR(24, "Modificacion Número no Válida")
      Text8.SetFocus
      Exit Sub
   End If
   MODINU% = Val(Text8)
   '
   DEMODI$ = TRIM$(Text9)
   If DEMODI$ = "" Then
      Call MENSERR(24, "Falta Descripción")
      Text9.SetFocus
      Exit Sub
   End If
   '
   NACTU% = 0
   If List1.ListCount > 0 Then
     NACTU% = 1 + Val(Mid$(List1.List(List1.ListCount - 1), 97, 6))
   End If
   '
   Call FRATEXTO(Text9, 62)
   ENCATEX$ = Space$(128)
   Call REPLA(ENCATEX$, TRIM$(Text6.TEXT), 1, 8)
   Call REPLA(ENCATEX$, AJUSTD$(Text7, 4), 10, 4)
   Call REPLA(ENCATEX$, AJUSTD$(Text8, 4), 15, 4)
   Call REPLA(ENCATEX$, RETEX$(1), 22, 62)
   Call REPLA(ENCATEX$, CSTRING$(MKI$(NACTU%), 1, 6, 0, 2), 97, 6)
   List1.AddItem ENCATEX$
   '
   For KKI% = 2 To CARETEX%
      ENCATEX$ = Space$(128)
      Call REPLA(ENCATEX$, RETEX$(KKI%), 22, 62)
      Call REPLA(ENCATEX$, CSTRING$(MKI$(NACTU%), 1, 6, 0, 2), 97, 6)
      List1.AddItem ENCATEX$
   Next KKI%
   '
   Text7 = ""
   Text8 = ""
   Text9 = ""
   '
End Sub

Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call CARINSTRUC
   Call CARHOJAC
   Command2.Enabled = True
End Sub

Private Sub COMMAND21_CLICK()
   Command21.Enabled = False
   Call DOCUMING.NUEINSTRU
   Command21.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   DOCUORI.Caption = ""
   COMMLAB.Caption = "PRINT"
   Command22.Enabled = True
End Sub

Private Sub Command3_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call MENSERR(24, "Ayuda en Línea no Disponible")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
End Sub

Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   ABC$ = Clipboard.GetText
10 For KKI% = 1 To Len(ABC)
     If Asc(Mid$(ABC$, KKI%, 1)) = 9 Then
       ABC$ = Left$(ABC$, KKI% - 1) + Space$(5) + Mid$(ABC$, KKI% + 1)
       GoTo 10
     End If
   Next KKI%
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + ABC$ + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub

Private Sub Command9_Click(Index As Integer)
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text5(Index) = VDEV$
   End If
End Sub

Private Sub COMMLAB_Change()
   If TRIM$(UCase$(COMMLAB.Caption)) = "PRINT" Then
      DOCUOR$ = TRIM$(DOCUORI.Caption)
      Call PRINPROCE(DOCUOR$)
      COMMLAB.Caption = ""
      DOCUORI.Caption = ""
   End If
End Sub

Private Sub Form_Load()
    '
    If Screen.Height > 9000 Then
       Top = (Screen.Height - Height) / 2
       Left = (Screen.Width - Width) / 2
    End If
    '
    Refresh
End Sub

Private Sub Form_Resize()
   DoEvents
   If WindowState = 1 Then
      Call BAJALDISCO
   End If
End Sub

Private Sub Label28_Change()
   '
   Call CARINSTRUC
   Call CARHOJAC
   '
End Sub

Private Sub Text1_Change()
   '
   'Call BLANFORMU
   'If CODINT%(Text1.TEXT) > 0 Then
   '    Call CARDATITEM
   'End If
   '
End Sub

Private Sub Text1_DblClick()
   '
   'Call GRADATITEM
   '
   'Call SELECHO("MASTER")
   'COD$ = TRIM$(VALACT1$("MASTER"))
   'If COD$ <> "" Then
   '  Text1.TEXT = COD$
   'End If
   '
End Sub

Private Sub Text1_GotFocus()
   '
   'Call GRADATITEM
   '
End Sub

Sub CARINSTRUC()
   '
   CIIINSTRU% = Val(Label28)
   If CIIINSTRU% < 1 Then Exit Sub
   '
   REGI& = CIIINSTRU%
   URN& = ULTREG&("INSTRUCD")
   XX$ = REGLEIDO$("INSTRUCI", REGI&)
   CODINSTRU$ = Left$(XX$, 16)
   DESINSTRU$ = Mid$(XX$, 17, 48)
   REVINSTRU$ = Mid$(XX$, 65, 4)
   REPREP$ = Mid$(XX$, 71, 24)
   FEPREP = CVI(Mid$(XX$, 95, 2))
   REAPRO$ = Mid$(XX$, 97, 24)
   FEAPRO = CVI(Mid$(XX$, 121, 2))
   '
   Label37 = CODINSTRU$
   Text3 = DESINSTRU$
   Text1 = REVINSTRU$
   Text4(0) = REPREP$
   Text5(0) = FECHATEX$(FEPREP)
   Text4(1) = REAPRO$
   Text5(1) = FECHATEX$(FEAPRO)
   OPENFORMS = DoEvents
   '
   Text28.TEXT = "": TTYZ$ = ""
   If CIIINSTRU% > 0 Then
     If CIIINSTRU% <= URN& Then
       Screen.MousePointer = 11
       RNC$ = RECFILT$("INSTRUCD", 1, MKI$(CIIINSTRU%))
       '
       For U& = 1 To Len(RNC$) Step 4
         RENOTA& = CVS(Mid$(RNC$, U&, 4))
         If RENOTA& > 0 Then
           If RENOTA& <= URN& Then
             TTYY$ = RTrim$(Mid$(REGLEIDO$("INSTRUCD", RENOTA&), 3))
             If TTYZ$ <> "" Then
               TTYZ$ = TTYZ$ + Chr$(13) + Chr$(10)
             End If
             TTYZ$ = TTYZ$ + RTrim$(TTYY$)
           End If
         End If
       Next U&
       Screen.MousePointer = 1
     End If
   End If
   '
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 30
   Next U&
   U& = 0
   '
30 Text28.TEXT = Left$(TTYZ$, U&)
   Text28.Refresh
   '
   HOII% = HOY(HOS$)
   Text6 = HOS$
   '
End Sub
'
Sub GRAINSTRUC()
   '
   URN& = ULTREG&("INSTRUCD")
   CIIINSTRU% = Val(Label28)
   If CIIINSTRU% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Call BLOQARCH("INSTRUCD")
   RNC$ = RECFILT$("INSTRUCD", 1, MKI$(CIIINSTRU%))
   KK& = (-3)
   '
   TTYZ$ = Text28.TEXT
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
   Next U&
   U& = 0
   '
10 TEXTO$ = Left$(TTYZ$, U&) + Chr$(13) + Chr$(10)
   Call FRATEXTO(TEXTO$, 92)
   For LL% = 1 To CARETEX%
      TTYY$ = MKI$(CIIINSTRU%) + RETEX$(LL%) + Space$(128)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("INSTRUCD", TTYY$, RENOTA&)
   Next LL%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("INSTRUCD", MKI$(0) + Space$(92), RENOTA&)
   Wend
   '
   Call CIERRARCH("INSTRUCD")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
End Sub
'
Sub CARHOJAC()
   '
   URN& = ULTREG&("INSTRUCI")
   CIIINSTRU% = Val(Label28)
   If CIIINSTRU% < 1 Then Exit Sub
   '
   List1.Clear
   If CIIINSTRU% > 0 Then
     If CIIINSTRU% <= URN& Then
       Screen.MousePointer = 11
       RNC$ = RECFILT$("HOJACTU", 1, MKI$(CIIINSTRU%))
       '
       For U& = 1 To Len(RNC$) Step 4
         RENOTA& = CVS(Mid$(RNC$, U&, 4))
         If RENOTA& > 0 Then
           If RENOTA& <= URN& Then
             TTYY$ = REGLEIDO$("HOJACTU", RENOTA&)
             TTZZ$ = Space$(128)
             Call REPLA(TTZZ$, Mid$(TTYY$, 5, 92), 1, 92)
             Call REPLA(TTZZ$, CSTRING$(Mid$(TTYY$, 3, 2), 1, 6, 0, 2), 97, 6)
             List1.AddItem TTZZ$
           End If
         End If
       Next U&
       Screen.MousePointer = 1
     End If
   End If
   '
End Sub
'
Sub GRAHOJAC()
   '
   URN& = ULTREG&("HOJACTU")
   CIIINSTRU% = Val(Label28)
   If CIIINSTRU% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Call BLOQARCH("HOJACTU")
   RNC$ = RECFILT$("HOJACTU", 1, MKI$(CIIINSTRU%))
   KK& = (-3)
   '
   For LL% = 1 To List1.ListCount
      TTZZ$ = List1.List(LL% - 1)
      NACTU% = Val(Mid$(TTZZ$, 97, 6))
      TTYY$ = MKI$(CIIINSTRU%) + MKI$(NACTU%) + Left$(TTZZ$, 92)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("HOJACTU", TTYY$, RENOTA&)
   Next LL%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("HOJACTU", MKI$(0) + MKI$(0) + Space$(92), RENOTA&)
   Wend
   '
   Call CIERRARCH("HOJACTU")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
End Sub

Sub BLANFORMU()
   '
   Label6.Caption = ""
   Label37.Caption = ""
   Label28.Caption = ""
   '
End Sub

Sub GRADATITEM()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
End Sub
'

Private Sub Text4_CLICK(Index As Integer)
   If TRIM$(Text4(Index)) = "" Then
      Text4(Index) = USERNAME$
   End If
End Sub

Private Sub Text5_DblClick(Index As Integer)
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text5(Index) = VDEV$
   End If
End Sub

Sub PRINPROCE(DOCUORIGEN$)
  '
  If List1.ListCount > 0 Then
     '
     FORIPRE$ = TRIM$(CONTROL$("FORGRAFI", "DOCUGRA4"))
     If FORIPRE$ <> "" Then
        Screen.MousePointer = 11
        FEX = HOY(HOS$)
        FECHDOC$ = HOS$
        NUMEDOC$ = ""
        TIPODOC$ = "Hoja de Actualización " + TRIM$(Label37)
        '
        CODPARAM$(1) = "COD-MAT"
        CODPARAM$(2) = "REF-MAT2"
        CODPARAM$(3) = "DES-MAT"
        CODPARAM$(4) = "VENDEDOR"
        CODPARAM$(5) = "COND-PAG"
        CODPARAM$(6) = "FECH-VEN"
        CAPARDOC% = 6
        '
        DOCPARAM$(1) = TRIM$(Label37)
        DOCPARAM$(2) = TRIM$(Text1)
        DOCPARAM$(3) = TRIM$(Text3)
        DOCPARAM$(4) = TRIM$(Text4(0))
        DOCPARAM$(5) = TRIM$(Text4(1))
        DOCPARAM$(6) = TRIM$(Text5(1))
        '
        CODTABU1$(1) = "NOP-OPER"
        CODTABU1$(2) = "DES-OPER"
        CODTABU1$(3) = "OBS-OPER"
        CACOLTAB1% = 3
        '
        CAITAB1% = 0
        ZANT$ = ""
        For U& = 1 To List1.ListCount
          Z$ = List1.List(U& - 1)
          If TRIM$(Left$(Z$, 18)) <> "" Then
              CAITAB1% = CAITAB1% + 1
              TETABU1$(1, CAITAB1%) = Mid$(Z$, 10, 4)
              TETABU1$(2, CAITAB1%) = Mid$(Z$, 15, 4)
              TETABU1$(3, CAITAB1%) = RTrim$(Mid$(Z$, 22, 62))
            Else
              TETABU1$(3, CAITAB1%) = TETABU1$(3, CAITAB1%) + " " + RTrim$(Mid$(Z$, 22, 62))
          End If
        Next U&
        '
        Call PRINTDOC("@" + FORIPRE$)
        Screen.MousePointer = 1
        '
     End If
  End If
  '
  FORIPRE$ = TRIM$(CONTROL$("FORGRAFI", "DOCUGRA2"))

  If FORIPRE$ = "" Then
       PRINTFORM
     Else
       Screen.MousePointer = 11
       FEX = HOY(HOS$)
       FECHDOC$ = HOS$
       NUMEDOC$ = ""
       TIPODOC$ = "Procedimiento " + TRIM$(Label37)
       '
       CODPARAM$(1) = "COD-MAT"
       CODPARAM$(2) = "REF-MAT2"
       CODPARAM$(3) = "DES-MAT"
       CODPARAM$(4) = "VENDEDOR"
       CODPARAM$(5) = "COND-PAG"
       CODPARAM$(6) = "FECH-VEN"
       CAPARDOC% = 6
       '
       DOCPARAM$(1) = TRIM$(Label37)
       DOCPARAM$(2) = TRIM$(Text1)
       DOCPARAM$(3) = TRIM$(Text3)
       DOCPARAM$(4) = TRIM$(Text4(0))
       DOCPARAM$(5) = TRIM$(Text4(1))
       DOCPARAM$(6) = TRIM$(Text5(1))
       '
       CODTABU1$(1) = "REF-MAT1"
       CACOLTAB1% = 1
       '
       CAITAB1% = 0
       Call FRATEXTO(Text28, 92)
       For U& = 1 To CARETEX%
         Z$ = "." + RETEX$(U&)
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = Z$
       Next U&
       '
       Call PRINTDOC("@" + FORIPRE$)
       Screen.MousePointer = 1
       '
  End If
  '
End Sub

