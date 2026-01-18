VERSION 5.00
Begin VB.Form FICALIB07 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha de Calibración"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      Height          =   8055
      Left            =   11025
      ScaleHeight     =   7995
      ScaleWidth      =   1275
      TabIndex        =   45
      Top             =   -240
      Width           =   1335
      Begin VB.CommandButton Command7 
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
         Height          =   690
         Left            =   120
         Picture         =   "FICALIB07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   5610
         Width           =   650
      End
      Begin VB.CommandButton Command5 
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
         Height          =   690
         Left            =   120
         Picture         =   "FICALIB07.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Alta de Nuevo Registro de Calibración"
         Top             =   4920
         Width           =   650
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
         Height          =   690
         Left            =   120
         Picture         =   "FICALIB07.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Imprime Ficha de Calibración"
         Top             =   3450
         Width           =   650
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
         Height          =   690
         Left            =   120
         Picture         =   "FICALIB07.frx":0BF6
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Alta de Nuevo Registro de Calibración"
         Top             =   2760
         Width           =   650
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
         Height          =   690
         Left            =   120
         Picture         =   "FICALIB07.frx":0D40
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Cierra la Ficha de Calibración"
         Top             =   480
         Width           =   650
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
         Height          =   690
         Left            =   120
         Picture         =   "FICALIB07.frx":0E8A
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Ayuda en Línea"
         Top             =   1170
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "FICALIB07.frx":1194
         Top             =   7470
         Width           =   1515
      End
      Begin VB.Label COMMLAB 
         Height          =   225
         Left            =   0
         TabIndex        =   53
         Top             =   3360
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label DOCUORI 
         Height          =   225
         Left            =   0
         TabIndex        =   52
         Top             =   6405
         Visible         =   0   'False
         Width           =   990
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2300
      Index           =   5
      Left            =   10290
      TabIndex        =   42
      Top             =   3780
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2300
      Index           =   4
      Left            =   8820
      TabIndex        =   41
      Top             =   3780
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2300
      Index           =   3
      Left            =   7350
      TabIndex        =   40
      Top             =   3780
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2300
      Index           =   2
      Left            =   5880
      TabIndex        =   39
      Top             =   3780
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2300
      Index           =   1
      Left            =   4450
      TabIndex        =   38
      Top             =   3780
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2300
      Index           =   0
      Left            =   3000
      TabIndex        =   31
      Top             =   3780
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2300
      Left            =   1575
      TabIndex        =   32
      Top             =   3780
      Width           =   15
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
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
      Height          =   1260
      Left            =   8625
      TabIndex        =   0
      Top             =   120
      Width           =   2220
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
         Height          =   690
         Left            =   240
         Picture         =   "FICALIB07.frx":30B6
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   400
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Undo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   1320
         Picture         =   "FICALIB07.frx":33C0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Restablece los Valores Anteriores"
         Top             =   400
         Width           =   650
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      Height          =   435
      Left            =   210
      ScaleHeight     =   375
      ScaleWidth      =   10545
      TabIndex        =   34
      Top             =   3780
      Width           =   10600
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "  Calibracion         IC-Nro           Utilizado       Calibración        Ajuste         Prox.Calibr.      Calibr.Nro."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   43
         Top             =   170
         Width           =   10515
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "    Fecha de                                   Patron           Requirio         Requirio        Fecha de       Registro de"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   10305
      End
   End
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
      Height          =   1860
      ItemData        =   "FICALIB07.frx":36CA
      Left            =   210
      List            =   "FICALIB07.frx":36CC
      TabIndex        =   33
      Top             =   4200
      Width           =   10605
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2160
      Left            =   315
      TabIndex        =   13
      Top             =   1470
      Width           =   10410
      Begin VB.ListBox List2 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1785
         TabIndex        =   37
         Top             =   1420
         Width           =   8520
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Componentes y/o Accesorios:"
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
         Left            =   105
         TabIndex        =   30
         Top             =   1365
         Width           =   1530
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Index           =   7
         Left            =   7245
         TabIndex        =   29
         Top             =   1050
         Width           =   3060
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Index           =   6
         Left            =   7245
         TabIndex        =   28
         Top             =   735
         Width           =   3060
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Index           =   5
         Left            =   7245
         TabIndex        =   27
         Top             =   420
         Width           =   3060
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Index           =   4
         Left            =   7245
         TabIndex        =   26
         Top             =   105
         Width           =   3060
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Index           =   3
         Left            =   1785
         TabIndex        =   25
         Top             =   1050
         Width           =   3690
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Index           =   2
         Left            =   1785
         TabIndex        =   24
         Top             =   735
         Width           =   3690
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Index           =   1
         Left            =   1785
         TabIndex        =   23
         Top             =   420
         Width           =   3690
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Index           =   0
         Left            =   1785
         TabIndex        =   22
         Top             =   105
         Width           =   3690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo:"
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
         Left            =   630
         TabIndex        =   21
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro de Serie:"
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
         Left            =   0
         TabIndex        =   20
         Top             =   735
         Width           =   1635
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
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
         Left            =   630
         TabIndex        =   19
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
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
         Index           =   0
         Left            =   630
         TabIndex        =   18
         Top             =   105
         Width           =   1005
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Intervalo:"
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
         Left            =   5775
         TabIndex        =   17
         Top             =   1050
         Width           =   1320
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Rango:"
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
         Left            =   6090
         TabIndex        =   16
         Top             =   105
         Width           =   1005
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Precisión:"
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
         Left            =   5880
         TabIndex        =   15
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación:"
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
         Left            =   5670
         TabIndex        =   14
         Top             =   735
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Instrumentos de Medición - Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   210
      TabIndex        =   2
      Top             =   105
      Width           =   8310
      Begin VB.Label Label2 
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
         Height          =   315
         Index           =   8
         Left            =   5775
         TabIndex        =   44
         Top             =   1155
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label14 
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
         Left            =   1890
         TabIndex        =   36
         Top             =   840
         Width           =   6210
      End
      Begin VB.Label Label16 
         Height          =   225
         Left            =   4830
         TabIndex        =   12
         Top             =   525
         Visible         =   0   'False
         Width           =   540
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
         Left            =   795
         TabIndex        =   10
         Top             =   495
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación:"
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
         Left            =   0
         TabIndex        =   9
         Top             =   900
         Width           =   1800
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         Left            =   1890
         TabIndex        =   6
         Top             =   420
         Width           =   1590
      End
   End
   Begin VB.Frame MARCOFICHA 
      BackColor       =   &H00C0FFC0&
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
      Height          =   1520
      Index           =   1
      Left            =   210
      TabIndex        =   3
      Top             =   6195
      Width           =   10620
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
         Height          =   1065
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   315
         Width           =   10275
      End
      Begin VB.Line Line1 
         X1              =   150
         X2              =   10430
         Y1              =   1375
         Y2              =   1375
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   10430
         X2              =   10430
         Y1              =   315
         Y2              =   1375
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Instrumentos de Medición - Calibración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   3465
      TabIndex        =   11
      Top             =   1320
      Width           =   7365
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   10710
      X2              =   10710
      Y1              =   3780
      Y2              =   6090
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   210
      X2              =   10815
      Y1              =   6075
      Y2              =   6075
   End
End
Attribute VB_Name = "FICALIB07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODIDOC%
Dim GRABATODO%
'
Private Sub Command1_Click()
   '
   GRABATODO% = 1
   Call GRADAFICAL
   GRABATODO% = 0
   Exit Sub
   
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call GRADAFICAL
   Call CARDAFICAL
   Call CARECALIB
   Command2.Enabled = True
End Sub

Private Sub COMMAND21_CLICK()
   Command21.Enabled = False
   Call GRADAFICAL
   '
   RECALIB07.Label37 = Label37
   RECALIB07.Label14 = Label14
   RECALIB07.Label28 = Label28
   RECALIB07.Label16 = Label16
   '
   For KKI% = 0 To 8
     RECALIB07.Label2(KKI%) = Label2(KKI%)
   Next KKI%
   '
   RECALIB07.List2.Clear
   For KKI% = 1 To List2.ListCount
     RECALIB07.List2.AddItem FICALIB07.List2.List(KKI% - 1)
   Next KKI%
   '
   RECALIB07.Label15.Caption = 1 + ULTREG&("RECALIB")
   RECALIB07.Show 1
   Call CARECALIB
   '
   Command21.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   Call GRADAFICAL
   Call PRIFICALIB
   Command22.Enabled = True
End Sub

Private Sub Command3_Click()
   Call GRADAFICAL
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call MENSERR(24, "Ayuda en Línea no Disponible")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Call GRADAFICAL
   '
   RECALIB07.Label37 = Label37
   RECALIB07.Label14 = Label14
   RECALIB07.Label28 = Label28
   RECALIB07.Label16 = Label16
   '
   For KKI% = 0 To 8
     RECALIB07.Label2(KKI%) = Label2(KKI%)
   Next KKI%
   '
   RECALIB07.List2.Clear
   For KKI% = 1 To List2.ListCount
     RECALIB07.List2.AddItem FICALIB07.List2.List(KKI% - 1)
   Next KKI%
   '
   RECALIB07.Label15.Caption = 1 + ULTREG&("RECALIB")
   RECALIB07.Show 1
   Call CARECALIB
   '
   Command5.Enabled = True
   '
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call GRADAFICAL
   CONFINSTR07.Show 1
   Command7.Enabled = True
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Screen.MousePointer = 1
    '
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Resize()
   DoEvents
   If WindowState = 1 Then
      Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label16_Change()
  '
  REINSTRUC& = Val(Label16)
  Label14 = ""
  Label15 = ""
  Label17 = ""
  Label18 = ""
  '
  If REINSTRUC& > 0 Then
    If REINSTRUC& <= ULTREG&("INSTRUCI") Then
      XX$ = REGLEIDO$("INSTRUCI", REINSTRUC&)
      CODINSTRU$ = Left$(XX$, 16)
      DESINSTRU$ = Mid$(XX$, 17, 48)
      REVINSTRU$ = Mid$(XX$, 65, 4)
      FEAPRO = CVI(Mid$(XX$, 121, 2))
      '
      Label14 = CODINSTRU$
      Label18 = DESINSTRU$
      Label15 = REVINSTRU$
      Label17 = FECHATEX$(FEAPRO)
    End If
  End If
  '
End Sub

Private Sub Label28_Change()
   Call CARECALIB
   Call CARDAFICAL
End Sub

Private Sub Label37_Click()
   Call GRADAFICAL
End Sub

Private Sub Label37_DblClick()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub List1_DblClick()
   '
   RECALIB07.Label37 = Label37
   RECALIB07.Label14 = Label14
   RECALIB07.Label28 = Label28
   RECALIB07.Label16 = Label16
   '
   For KKI% = 0 To 8
     RECALIB07.Label2(KKI%) = Label2(KKI%)
   Next KKI%
   '
   RECALIB07.List2.Clear
   For KKI% = 1 To List2.ListCount
     RECALIB07.List2.AddItem FICALIB07.List2.List(KKI% - 1)
   Next KKI%
   '
   RECALIB07.Label15.Caption = TRIM$(Mid$(List1.TEXT, 75, 8))
   RECALIB07.Show 1
   Call CARECALIB
   '
End Sub

Private Sub Text28_Change()
   MODIDOC% = 1
End Sub

Sub ALTAINSTRU()
   '
   VDEF$ = REGBLAN$("INSTRUME")
   INDINS% = 1 + ULTREG&("INSTRUME")
   Call REPLA(VDEF$, MKI$(INDINS%), 65, 2)
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("NUEINSTRUMEN", VDEF$)
   If OBX$ = "" Then
     Exit Sub
   End If
   '
   VDEF$ = OBX$
   CODINS$ = UCase$(TRIM$(Left$(OBX$, 12)))
   If CODINS$ = "" Then
      Call MENSERR(24, "Falta Código de Instrumento")
      GoTo 10
   End If
   Call REPLA(OBX$, CODINS$, 1, 12)
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("INSTRUME")
   For UI& = 1 To FIN&
      XX$ = REGLEIDO$("INSTRUME", UI&)
      If UCase$(TRIM$(Left$(XX$, 12))) = CODINS$ Then
         Call MENSERR(24, "Código Pre-Existente")
         GoTo 10
      End If
   Next UI&
   '
   DESINS$ = TRIM$(Mid$(OBX$, 13, 52))
   If DESINS$ = "" Then
      Call MENSERR(24, "Falta Denominación del Instrumento")
      GoTo 10
   End If
   '
   GRUIN% = CVI(Mid$(OBX$, 67, 2))
   If GRUIN% < 1 Or GRUIN% > ULTREG&("GRUINSTR") Then
      Call MENSERR(24, "Grupo de Instrumento Inexistente o No Válido")
      GoTo 10
   End If
   '
   Call CIERRARCH("INSTRUME")
   Call BLOQARCH("INSTRUME")
   INDINS% = 1 + ULTREG&("INSTRUME")
   Call REPLA(OBX$, MKI$(INDINS%), 65, 2)
   Call REGAPP("INSTRUME", OBX$)
   Call CIERRARCH("INSTRUME")
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
   On Error Resume Next
   INSTRUMEN.Label28 = TRIM$(Str$(INDINS%))
   INSTRUMENSA.Label28 = TRIM$(Str$(INDINS%))
   '
   On Error GoTo 0
   On Error Resume Next
   INSTRUMEN.Show 1
   If Err Then
     On Error GoTo 0
     INSTRUMENSA.Show 1
   End If
   On Error GoTo 0
   '
End Sub

Sub CARECALIB()
   '
   List1.Clear
   NINSTRU% = Val(Label28)
   If NINSTRU% < 1 Then Exit Sub
   '
   URN& = ULTREG&("RECALIB")
   RNC$ = RECFILT$("RECALIB", 2, MKI$(NINSTRU%))
   For KK& = 1 To Len(RNC$) Step 4
     REGI& = CVS(Mid$(RNC$, KK&, 4))
     If REGI& > 0 Then
       If REGI& <= URN& Then
         XX$ = REGLEIDO$("RECALIB", REGI&)
         ZZ$ = Space$(128)
         FEX = CVI(Mid$(XX$, 5, 2))
         Call REPLA(ZZ$, FECHATEX$(FEX), 3, 8)
         FEX = CVI(Mid$(XX$, 7, 2))
         Call REPLA(ZZ$, FECHATEX$(FEX), 63, 8)
         Call REPLA(ZZ$, Mid$(XX$, 17, 16), 13, 12)
         Call REPLA(ZZ$, Mid$(XX$, 33, 16), 25, 12)
         Call REPLA(ZZ$, Mid$(XX$, 49, 16), 37, 12)
         Call REPLA(ZZ$, Mid$(XX$, 65, 16), 49, 12)
         Call REPLA(ZZ$, CSTRING$(MKS$(REGI&), 3, 8, 0, 2), 75, 8)
         List1.AddItem ZZ$
       End If
     End If
   Next KK&

End Sub

Sub PRIFICALIB()
   '
   FORIPRE$ = TRIM$(CONTROL$("FORINMED", "DOCUMED2"))
   If FORIPRE$ = "" Then
     Call MENSERR(24, "Falta Definir Formulario de Impresión")
     CONFINSTR07.Show 1
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   Call FRATEXTO(Text28, 82)
   JJ& = 0
   For U& = 1 To CARETEX%
     Z$ = RETEX$(U&)
     JJ& = JJ& + 1
     Call GRAREG("!OBSERVOF", Z$, JJ&)
   Next U&
   CARE& = CARETEX%
   Call SETULTREG("!OBSERVOF", CARE&)
   Call CIERRARCH("!OBSERVOF")
   '
   FEX = HOY(HOS$)
   FECHDOC$ = HOS$
   NUMEDOC$ = ""
   TIPODOC$ = "Ficha de Calibración " + " - " + TRIM$(Label37)
   '
   CODPARAM$(1) = "FECH-LAN"
   CODPARAM$(2) = "UNIMED"
   CODPARAM$(3) = "REF-MAT1"
   CODPARAM$(4) = "REF-MAT2"
   CODPARAM$(5) = "REF-MAT3"
   CODPARAM$(6) = "NRO-LOTE"
   CODPARAM$(7) = "TALLE"
   CODPARAM$(8) = "COLOR"
   CODPARAM$(9) = "MEDIDA-1"
   CODPARAM$(10) = "EQUIPO1"
   CODPARAM$(11) = "EQUIPO2"
   CODPARAM$(12) = "COD-MAT"
   CODPARAM$(13) = "DES-MAT"
   CAPARDOC% = 13
   '
   DOCPARAM$(1) = FECHDOC$
   DOCPARAM$(2) = TRIM$(Label2(0))
   DOCPARAM$(3) = TRIM$(Label2(1))
   DOCPARAM$(4) = TRIM$(Label2(2))
   DOCPARAM$(5) = TRIM$(Label2(3))
   DOCPARAM$(6) = TRIM$(Label2(4))
   DOCPARAM$(7) = TRIM$(Label2(5))
   DOCPARAM$(8) = TRIM$(Label2(6))
   DOCPARAM$(9) = TRIM$(Label2(7))
   On Error Resume Next
     DOCPARAM$(10) = List2.List(0)
     DOCPARAM$(11) = List2.List(1)
   On Error GoTo 0
   DOCPARAM$(12) = TRIM$(Label37)
   DOCPARAM$(13) = TRIM$(Label14)
  
   '
   CODTABU1$(1) = "MEDIDA-3"
   CODTABU1$(2) = "N-PLANO"
   CODTABU1$(3) = "R-PLANO"
   CODTABU1$(4) = "F-PLANO"
   CODTABU1$(5) = "Z-PLANO"
   CODTABU1$(6) = "N-HPROF"
   CODTABU1$(7) = "N-HINSP"
   CACOLTAB1% = 7
   '
   CAITAB1% = 0
   For U& = 1 To List1.ListCount
     Z$ = List1.List(U& - 1)
     CAITAB1% = CAITAB1% + 1
     TETABU1$(1, CAITAB1%) = TRIM$(Mid$(Z$, 1, 12))
     TETABU1$(2, CAITAB1%) = TRIM$(Mid$(Z$, 13, 12))
     TETABU1$(3, CAITAB1%) = TRIM$(Mid$(Z$, 25, 12))
     TETABU1$(4, CAITAB1%) = TRIM$(Mid$(Z$, 37, 12))
     TETABU1$(5, CAITAB1%) = TRIM$(Mid$(Z$, 49, 12))
     TETABU1$(6, CAITAB1%) = TRIM$(Mid$(Z$, 61, 12))
     TETABU1$(7, CAITAB1%) = TRIM$(Mid$(Z$, 73, 12))
   Next U&
   '
   Call PRINTDOC("@" + FORIPRE$)
   Screen.MousePointer = 1
   '
End Sub

Sub CARDAFICAL()
   '
   CIIINSTRU% = Val(Label28)
   If CIIINSTRU% < 1 Then Exit Sub
   '
   REGI& = CIIINSTRU%
   URN& = ULTREG&("INSTRUME")
   URX& = ULTREG&("OBFICAL")
   '
   If REGI& <= URN& Then
     Text28.TEXT = "": TTYZ$ = ""
     If CIIINSTRU% > 0 Then
       If CIIINSTRU% <= URN& Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("OBFICAL", 1, MKI$(CIIINSTRU%))
         '
         For U& = 1 To Len(RNC$) Step 4
           RENOTA& = CVS(Mid$(RNC$, U&, 4))
           If RENOTA& > 0 Then
             If RENOTA& <= URX& Then
               TTYY$ = RTrim$(Mid$(REGLEIDO$("OBFICAL", RENOTA&), 3))
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
30   Text28.TEXT = Left$(TTYZ$, U&)
     Text28.Refresh
     '
     OPENFORMS = DoEvents
     MODIDOC% = 0
     '
   End If
   '
End Sub
'
Sub GRADAFICAL()
   '
   REGI& = Val(Label28)
   If REGI& < 1 Then Exit Sub
   If REGI& > ULTREG&("INSTRUME") Then Exit Sub
   If MODIDOC% <> 1 Then Exit Sub
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se han Modificado los Datos\de la Ficha de Calibración '" + TRIM$(Label37) + " - " + TRIM$(Label14) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       Call CARDAFICAL
       MODIDOC% = 0
       Exit Sub
     End If
   End If
   '
   CIIINSTRU% = REGI&
   URN& = ULTREG&("OBFICAL")
   RNC$ = RECFILT$("OBFICAL", 1, MKI$(CIIINSTRU%))
   KK& = (-3)
   '
   TTYZ$ = Text28.TEXT
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
   Next U&
   U& = 0
   '
10 TEXTO1$ = Left$(TTYZ$, U&) + Chr$(13) + Chr$(10)
   Call FRATEXTO(TEXTO1$, 82)
   For LL% = 1 To CARETEX%
      TTYY$ = MKI$(CIIINSTRU%) + RETEX$(LL%)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("OBFICAL", TTYY$, RENOTA&)
   Next LL%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("OBFICAL", MKI$(0) + Space$(82), RENOTA&)
   Wend
   '
   Call CIERRARCH("OBFICAL")
   MODIDOC% = 0
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
End Sub
