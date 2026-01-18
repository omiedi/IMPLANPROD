VERSION 5.00
Begin VB.Form RECALIB04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instrumentos de Medición y Ensayo"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2520
      Index           =   7
      Left            =   10340
      TabIndex        =   64
      Top             =   3780
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2520
      Index           =   6
      Left            =   9250
      TabIndex        =   63
      Top             =   3780
      Width           =   15
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame9"
      Height          =   120
      Left            =   10810
      TabIndex        =   61
      Top             =   6405
      Width           =   980
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Reg.Cal."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   7
      Left            =   10815
      TabIndex        =   60
      Top             =   5740
      Width           =   1015
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   105
         TabIndex        =   62
         Top             =   260
         Width           =   795
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2520
      Index           =   5
      Left            =   8160
      TabIndex        =   42
      Top             =   3780
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2520
      Index           =   4
      Left            =   7070
      TabIndex        =   41
      Top             =   3780
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2520
      Index           =   3
      Left            =   5980
      TabIndex        =   40
      Top             =   3780
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2520
      Index           =   2
      Left            =   4890
      TabIndex        =   39
      Top             =   3780
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2520
      Index           =   1
      Left            =   3800
      TabIndex        =   38
      Top             =   3780
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2520
      Index           =   0
      Left            =   2710
      TabIndex        =   31
      Top             =   3780
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2520
      Left            =   1575
      TabIndex        =   32
      Top             =   3780
      Width           =   15
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
      Height          =   2085
      ItemData        =   "RECALIB04.frx":0000
      Left            =   105
      List            =   "RECALIB04.frx":0002
      TabIndex        =   33
      Top             =   4200
      Width           =   10605
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
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
      Left            =   8520
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
         Picture         =   "RECALIB04.frx":0004
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
         Picture         =   "RECALIB04.frx":030E
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
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   10545
      TabIndex        =   34
      Top             =   3780
      Width           =   10600
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   $"RECALIB04.frx":0618
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
         Top             =   60
         Width           =   10305
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2160
      Left            =   210
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
      BackColor       =   &H00C0FFFF&
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
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   8310
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Procedimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   5775
         TabIndex        =   57
         Top             =   315
         Width           =   2325
         Begin VB.CommandButton Command6 
            Height          =   520
            Left            =   1785
            Picture         =   "RECALIB04.frx":06A0
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   315
            Width           =   435
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
            Height          =   310
            Index           =   8
            Left            =   210
            TabIndex        =   58
            Top             =   525
            Width           =   1500
         End
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
         Height          =   540
         Left            =   1890
         TabIndex        =   36
         Top             =   840
         Width           =   3710
      End
      Begin VB.Label Label16 
         Height          =   225
         Left            =   4410
         TabIndex        =   12
         Top             =   210
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
         Left            =   4860
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
         Left            =   3780
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
      BackColor       =   &H00C0FFFF&
      Caption         =   "Resultado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   1
      Left            =   105
      TabIndex        =   3
      Top             =   6350
      Width           =   11730
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "I-me"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   9
         Left            =   5880
         TabIndex        =   70
         Top             =   140
         Width           =   4530
         Begin VB.Label Label24 
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
            Height          =   255
            Left            =   1260
            TabIndex        =   74
            Top             =   135
            Width           =   1065
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Min:"
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
            Left            =   735
            TabIndex        =   73
            Top             =   150
            Width           =   435
         End
         Begin VB.Label Label21 
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
            Height          =   255
            Left            =   3150
            TabIndex        =   72
            Top             =   135
            Width           =   1065
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Max:"
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
            Left            =   2625
            TabIndex        =   71
            Top             =   150
            Width           =   435
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "E-prom"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   8
         Left            =   1260
         TabIndex        =   65
         Top             =   140
         Width           =   4530
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Max:"
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
            Left            =   2625
            TabIndex        =   69
            Top             =   150
            Width           =   435
         End
         Begin VB.Label Label18 
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
            Height          =   255
            Left            =   3150
            TabIndex        =   68
            Top             =   135
            Width           =   1065
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Min:"
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
            Left            =   735
            TabIndex        =   67
            Top             =   150
            Width           =   435
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
            Height          =   255
            Left            =   1260
            TabIndex        =   66
            Top             =   135
            Width           =   1065
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   6
         Left            =   10500
         TabIndex        =   55
         Top             =   525
         Width           =   1120
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
            Height          =   300
            Index           =   6
            Left            =   105
            MaxLength       =   12
            TabIndex        =   56
            Top             =   220
            Width           =   910
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Responsable"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   5
         Left            =   8610
         TabIndex        =   53
         Top             =   1215
         Width           =   3010
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
            Index           =   5
            Left            =   210
            MaxLength       =   32
            TabIndex        =   54
            Top             =   220
            Width           =   2690
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Prox.Calibr"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   4
         Left            =   7245
         TabIndex        =   51
         Top             =   1215
         Width           =   1275
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
            Index           =   4
            Left            =   210
            MaxLength       =   12
            TabIndex        =   52
            Top             =   210
            Width           =   960
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Requirió Ajuste"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   3
         Left            =   5460
         TabIndex        =   49
         Top             =   1215
         Width           =   1695
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
            Index           =   3
            Left            =   210
            MaxLength       =   12
            TabIndex        =   50
            Top             =   210
            Width           =   1380
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Req.Calibración"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   2
         Left            =   3675
         TabIndex        =   47
         Top             =   1215
         Width           =   1695
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
            Index           =   2
            Left            =   210
            MaxLength       =   12
            TabIndex        =   48
            Top             =   210
            Width           =   1380
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Patrón Utilizado"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   1890
         TabIndex        =   45
         Top             =   1215
         Width           =   1695
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
            Index           =   1
            Left            =   210
            MaxLength       =   12
            TabIndex        =   46
            Top             =   210
            Width           =   1380
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "IC-Nro"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   0
         Left            =   105
         TabIndex        =   43
         Top             =   1215
         Width           =   1695
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
            Index           =   0
            Left            =   210
            MaxLength       =   12
            TabIndex        =   44
            Top             =   220
            Width           =   1380
         End
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
         Height          =   540
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   630
         Width           =   10275
      End
      Begin VB.Line Line1 
         X1              =   105
         X2              =   10385
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   10395
         X2              =   10395
         Y1              =   630
         Y2              =   1155
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
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
      Left            =   3360
      TabIndex        =   11
      Top             =   1320
      Width           =   7365
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008080&
      Height          =   5775
      Left            =   10920
      ScaleHeight     =   5715
      ScaleWidth      =   1035
      TabIndex        =   75
      Top             =   -120
      Width           =   1095
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
         Left            =   150
         Picture         =   "RECALIB04.frx":09AA
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   4080
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
         Left            =   150
         Picture         =   "RECALIB04.frx":0DEC
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Cierra el Registro de Calibración"
         Top             =   315
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
         Left            =   150
         Picture         =   "RECALIB04.frx":0F36
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Ayuda en Línea"
         Top             =   1005
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
         Left            =   150
         Picture         =   "RECALIB04.frx":1240
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Imprime Registro de Calibración"
         Top             =   2985
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
         Left            =   150
         Picture         =   "RECALIB04.frx":18AA
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Alta de Nuevo Registro de Calibración"
         Top             =   2280
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -250
         Picture         =   "RECALIB04.frx":19F4
         Top             =   5280
         Width           =   1515
      End
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   10710
      X2              =   10710
      Y1              =   3780
      Y2              =   6510
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   105
      X2              =   10710
      Y1              =   6300
      Y2              =   6300
   End
End
Attribute VB_Name = "RECALIB04"
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
   Call GRADECALIB
   GRABATODO% = 0
   Exit Sub
   
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call GRADECALIB
   Call CARDECALIB
   Command2.Enabled = True
End Sub

Private Sub COMMAND21_CLICK()
   Command21.Enabled = False
   Call GRADECALIB
   Label15.Caption = 1 + ULTREG&("RECALIB")
   Call CADECALIB
   Command21.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   Call GRADECALIB
   Call PRIRECALIB
   Command22.Enabled = True
End Sub

Private Sub Command3_Click()
   Call GRADECALIB
   Call CIERRARCH("*.*")
   Hide
End Sub


Private Sub Command4_Click()
   Command4.Enabled = False
   Call MENSERR(24, "Ayuda en Línea no Disponible")
   Command4.Enabled = True
End Sub


Private Sub Command6_Click()
   Command6.Enabled = False
   Call GRADECALIB
   CIIINSTRU% = Val(Label16)
   If CIIINSTRU% < 1 Or CIIINSTRU% > ULTREG&("INSTRUCI") Then
      GoTo 99
      Exit Sub
   End If
   '
   On Error Resume Next
   INSTRUCCI.Label28 = TRIM$(Str$(CIIINSTRU%))
   INSTRUI04.Label28 = TRIM$(Str$(CIIINSTRU%))
   INSTRUCCI.Show 1
   If Err > 0 Then
     On Error GoTo 0
     INSTRUI04.Show 1
   End If
   On Error GoTo 0
   '
99 Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call GRADECALIB
   CONFINSTR04.Show 1
   Command7.Enabled = True
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call CENTRAFORM(Me)
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

Private Sub Label14_DblClick()
   XX$ = REGSEL$("INSTRUCI")
   If Len(XX$) > 4 Then
     CIIINSTRU% = REGNUM("INSTRUCI")
     Label16 = CIIINSTRU%
   End If
End Sub

Private Sub Label15_Change()
   If Val(Label15) <= ULTREG("RECALIB") Then
       Call CARDECALIB
     Else
       FEX = HOY(HOS$)
       Text1(6) = HOS$
       Text1(4) = ""
       Text1(0) = "IC." + TRIM$(Label15)
       Text1(1) = ""
       Text1(2) = ""
       Text1(3) = ""
       Text1(5) = USERNAME$
       Text28 = ""
       List1.Clear
   End If
   OPENFORMS = DoEvents
End Sub

Private Sub Label16_Change()
  '
  MODIDOC% = 1
  REINSTRUC& = Val(Label16)
  Label14 = ""
  Label15 = ""
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
      Label15 = REVINSTRU$
    End If
  End If
  '
End Sub

Private Sub Label22_Click(Index As Integer)
   Call CADECALIB
End Sub

Private Sub Label28_Change()
   '
End Sub

Private Sub Label37_Click()
   'Call GRADAINSTRUM
End Sub

Private Sub Label37_DblClick()
   XX$ = REGSEL$("INSTRUME")
   If Len(XX$) > 4 Then
     CIIINSTRU% = REGNUM&("INSTRUME")
     Label28 = CIIINSTRU%
   End If
End Sub


Private Sub List1_DblClick()
   Call CADECALIB
End Sub

Private Sub Text1_Change(Index As Integer)
   MODIDOC% = 1
End Sub

Private Sub Text1_DblClick(Index As Integer)
   If Index = 4 Or Index = 6 Then
     Call SELECHO("SELFECHA")
     VDEV$ = VALACT1$("SELFECHA")
     If VDEV$ <> "" Then
       Text1(Index) = VDEV$
     End If
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   Text1(Index).SelStart = 0
   Text1(Index).SelLength = Len(Text1(Index).TEXT)
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
   INSTRUMEN.Label28 = TRIM$(Str$(INDINS%))
   On Error Resume Next
   INSTRUMEN.Show 1
   On Error GoTo 0
   '
End Sub

Sub CADECALIB()
   JJ& = 0
   For KKL& = 1 To List1.ListCount
     X$ = List1.List(KKL& - 1)
     Z$ = REGBLAN$("DECALIB")
     Call REPLA(Z$, MKI$(Val(Label15)), 1, 2)
     Call REPLA(Z$, Mid$(X$, 1, 12), 3, 12)
     Call REPLA(Z$, MKD$(Val(Mid$(X$, 13, 9))), 15, 8)
     Call REPLA(Z$, MKD$(Val(Mid$(X$, 22, 9))), 23, 8)
     Call REPLA(Z$, MKD$(Val(Mid$(X$, 31, 9))), 31, 8)
     Call REPLA(Z$, MKD$(Val(Mid$(X$, 40, 9))), 39, 8)
     Call REPLA(Z$, MKD$(Val(Mid$(X$, 49, 9))), 47, 8)
     Call REPLA(Z$, MKD$(Val(Mid$(X$, 58, 9))), 55, 8)
     JJ& = JJ& + 1
     Call GRAREG("!DECALIB", Z$, JJ&)
   Next KKL&
   Call SETULTREG("!DECALIB", JJ&)
   Call CIERRARCH("!DECALIB")
   '
20 ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Carga Registro de Calibración"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(10 + Picture3.Left))
   '
   VTB% = VENTABU%("DECALIB" + ATRI$)
   If VTB% < 0 Then
      Exit Sub
   End If
   '
   List1.Clear
   DELTAMI# = 9999.999: DELTAMA# = 0
   DELCUMI# = 9999.999: DELCUMA# = 0
   For KKL& = 1 To ULTREG&("!DECALIB")
     Z$ = REGLEIDO$("!DECALIB", KKL&)
     X$ = Space$(80)
     Call REPLA(X$, Mid$(Z$, 3, 12), 1, 12)
     DELTA# = 0: DELTACUA# = 0
     For KKI% = 1 To 6
       TTXX$ = CSTRING$(Mid$(Z$, 7 + 8 * KKI%, 8), 3, 9, 3, 2)
       Call REPLA(X$, TTXX$, 4 + 9 * KKI%, 9)
       If KKI% = 1 Then
           XNOM# = Val(TTXX$)
         Else
           DELTA# = DELTA# + Abs(Val(TTXX$) - XNOM#)
           DELTACUA# = DELTACUA# + (Val(TTXX$) - XNOM#) ^ 2
       End If
     Next KKI%
     DELTA# = DELTA# / 5
     If DELTA# > 9999.999 Then DELTA# = 9999.999
     TTXX$ = CSTRING$(MKD$(DELTA#), 3, 9, 3, 2)
     Call REPLA(X$, TTXX$, 4 + 9 * 7, 9)
     If DELTA# < DELTAMI# Then DELTAMI# = DELTA#
     If DELTA# > DELTAMA# Then DELTAMA# = DELTA#
     DELTACUA# = 1.6 * Sqr(DELTACUA# / 4)
     If DELTACUA# > 9999.999 Then DELTACUA# = 9999.999
     TTXX$ = CSTRING$(MKD$(DELTACUA#), 3, 9, 3, 2)
     Call REPLA(X$, TTXX$, 4 + 9 * 8, 9)
     If DELTACUA# < DELCUMI# Then DELCUMI# = DELTACUA#
     If DELTACUA# > DELCUMA# Then DELCUMA# = DELTACUA#
     List1.AddItem X$
   Next KKL&
   Label7 = CSTRING$(MKD$(DELTAMI#), 3, 9, 3, 2)
   Label18 = CSTRING$(MKD$(DELTAMA#), 3, 9, 3, 2)
   Label24 = CSTRING$(MKD$(DELCUMI#), 3, 9, 3, 2)
   Label21 = CSTRING$(MKD$(DELCUMA#), 3, 9, 3, 2)
   List1.Refresh
   MODIDOC% = 1
   '
End Sub

Sub CARDECALIB()
   '
   NINFO% = Val(Label15)
   REGI& = NINFO%
   If REGI& < 1 Then Exit Sub
   If REGI& > ULTREG&("RECALIB") Then Exit Sub
   '
   XX$ = REGLEIDO$("RECALIB", REGI&)
   Label28 = CVI(Mid$(XX$, 3, 2))
   FEX = CVI(Mid$(XX$, 5, 2))
   Text1(6) = FECHATEX$(FEX)
   FEX = CVI(Mid$(XX$, 7, 2))
   Text1(4) = FECHATEX$(FEX)
   Text1(0) = Mid$(XX$, 17, 16)
   Text1(1) = Mid$(XX$, 33, 16)
   Text1(2) = Mid$(XX$, 49, 16)
   Text1(3) = Mid$(XX$, 65, 16)
   Text1(5) = Mid$(XX$, 81, 32)
   Text28 = Mid$(XX$, 129, 128)
   Call CIERRARCH("RECALIB")
   '
   URN& = ULTREG&("DECALIB")
   RNC$ = RECFILT$("DECALIB", 1, MKI$(NINFO%))
   KK& = (-3)
   '
   List1.Clear
   DELTAMI# = 9999.999: DELTAMA# = 0
   DELCUMI# = 9999.999: DELCUMA# = 0
   For KK& = 1 To Len(RNC$) Step 4
     KKL& = CVS(Mid$(RNC$, KK&, 4))
     Z$ = REGLEIDO$("DECALIB", KKL&)
     X$ = Space$(80)
     Call REPLA(X$, Mid$(Z$, 3, 12), 1, 12)
     DELTA# = 0: DELTACUA# = 0
     For KKI% = 1 To 6
       TTXX$ = CSTRING$(Mid$(Z$, 7 + 8 * KKI%, 8), 3, 9, 3, 2)
       Call REPLA(X$, TTXX$, 4 + 9 * KKI%, 9)
       If KKI% = 1 Then
           XNOM# = Val(TTXX$)
         Else
           DELTA# = DELTA# + Abs(Val(TTXX$) - XNOM#)
           DELTACUA# = DELTACUA# + (Val(TTXX$) - XNOM#) ^ 2
       End If
     Next KKI%
     DELTA# = DELTA# / 5
     If DELTA# > 9999.999 Then DELTA# = 9999.999
     TTXX$ = CSTRING$(MKD$(DELTA#), 3, 9, 3, 2)
     Call REPLA(X$, TTXX$, 4 + 9 * 7, 9)
     If DELTA# < DELTAMI# Then DELTAMI# = DELTA#
     If DELTA# > DELTAMA# Then DELTAMA# = DELTA#
     DELTACUA# = 1.6 * Sqr(DELTACUA# / 4)
     If DELTACUA# > 9999.999 Then DELTACUA# = 9999.999
     TTXX$ = CSTRING$(MKD$(DELTACUA#), 3, 9, 3, 2)
     Call REPLA(X$, TTXX$, 4 + 9 * 8, 9)
     If DELTACUA# < DELCUMI# Then DELCUMI# = DELTACUA#
     If DELTACUA# > DELCUMA# Then DELCUMA# = DELTACUA#
     List1.AddItem X$
   Next KK&
   Label7 = CSTRING$(MKD$(DELTAMI#), 3, 9, 3, 2)
   Label18 = CSTRING$(MKD$(DELTAMA#), 3, 9, 3, 2)
   Label24 = CSTRING$(MKD$(DELCUMI#), 3, 9, 3, 2)
   Label21 = CSTRING$(MKD$(DELCUMA#), 3, 9, 3, 2)
   List1.Refresh
   '
   Call CIERRARCH("DECALIB")
   MODIDOC% = 0
   '
End Sub

Sub GRADECALIB()
   '
   NINFO% = Val(Label15)
   REGI& = NINFO%
   If REGI& < 1 Then Exit Sub
   If REGI& > 1 + ULTREG&("RECALIB") Then Exit Sub
   If List1.ListCount < 1 Then Exit Sub
   If MODIDOC% <> 1 Then Exit Sub
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se han Modificado los Datos del\Registro de Calibración Numero " + TRIM$(Label15) + "\\Registrar Cambios\Cancelar") <> 1 Then
       MODIDOC% = 0
       Exit Sub
     End If
   End If
   '
   For KKI% = 0 To 6
      If TRIM$(Text1(KKI%)) = "" Then
         Call MENSERR(24, "Parámetro no válido")
         On Error Resume Next
         Text1(KKI%).SetFocus
         On Error GoTo 0
         Exit Sub
      End If
   Next KKI%
   '
   If FECHANUM(Text1(4)) < 1 Then
      Call MENSERR(24, "Fecha no Válida")
      On Error Resume Next
      Text1(4).SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   If FECHANUM(Text1(6)) < 1 Then
      Call MENSERR(24, "Fecha no Válida")
      On Error Resume Next
      Text1(6).SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   Call BLOQARCH("RECALIB")
   Call BLOQARCH("DECALIB")
   XX$ = REGLEIDO$("RECALIB", REGI&)
   Call REPLA(XX$, MKI$(NINFO%), 1, 2)
   Call REPLA(XX$, MKI$(Val(Label28)), 3, 2)
   FEXI% = FECHANUM(TRIM$(Text1(6)))
   Call REPLA(XX$, MKI$(FEXI%), 5, 2)
   FEXI% = FECHANUM(TRIM$(Text1(4)))
   Call REPLA(XX$, MKI$(FEXI%), 7, 2)
   Call REPLA(XX$, TRIM$(Text1(0)), 17, 16)
   Call REPLA(XX$, TRIM$(Text1(1)), 33, 16)
   Call REPLA(XX$, TRIM$(Text1(2)), 49, 16)
   Call REPLA(XX$, TRIM$(Text1(3)), 65, 16)
   Call REPLA(XX$, TRIM$(Text1(5)), 81, 32)
   Call REPLA(XX$, RTrim$(Text28), 129, 128)
   Call GRAREG("RECALIB", XX$, REGI&)
   Call CIERRARCH("RECALIB")
   '
   URN& = ULTREG&("DECALIB")
   RNC$ = RECFILT$("DECALIB", 1, MKI$(NINFO%))
   KK& = (-3)
   '
   For KKL& = 1 To List1.ListCount
     X$ = List1.List(KKL& - 1)
     Z$ = REGBLAN$("DECALIB")
     Call REPLA(Z$, MKI$(NINFO%), 1, 2)
     Call REPLA(Z$, Mid$(X$, 1, 12), 3, 12)
     Call REPLA(Z$, MKD$(Val(Mid$(X$, 13, 9))), 15, 8)
     Call REPLA(Z$, MKD$(Val(Mid$(X$, 22, 9))), 23, 8)
     Call REPLA(Z$, MKD$(Val(Mid$(X$, 31, 9))), 31, 8)
     Call REPLA(Z$, MKD$(Val(Mid$(X$, 40, 9))), 39, 8)
     Call REPLA(Z$, MKD$(Val(Mid$(X$, 49, 9))), 47, 8)
     Call REPLA(Z$, MKD$(Val(Mid$(X$, 58, 9))), 55, 8)
     Call REPLA(Z$, MKD$(Val(Mid$(X$, 67, 9))), 63, 8)
     Call REPLA(Z$, MKD$(Val(Mid$(X$, 76, 9))), 71, 8)
     '
     KK& = KK& + 4
     If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
     End If
     Call GRAREG("DECALIB", Z$, RENOTA&)
   Next KKL&
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("DECALIB", MKI$(0) + Space$(12) + String$(48, 0), RENOTA&)
   Wend
   '
   Call CIERRARCH("DECALIB")
   MODIDOC% = 0
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
End Sub

Sub PRIRECALIB()
   '
   FORIPRE$ = TRIM$(CONTROL$("FORINMED", "DOCUMED1"))
   If FORIPRE$ = "" Then
     Call MENSERR(24, "Falta Definir Formulario de Impresión")
     CONFINSTRUM.Show 1
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
   CARE& = CARETEX%: If CARE& > 3 Then CARE& = 3
   Call SETULTREG("!OBSERVOF", CARE&)
   Call CIERRARCH("!OBSERVOF")
   '
   FEX = HOY(HOS$)
   FECHDOC$ = HOS$
   NUMEDOC$ = TRIM$(Label15)
   TIPODOC$ = "Registro de Calibración " + TRIM$(Label15) + " - " + TRIM$(Label37)
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
   CODPARAM$(10) = "MEDIDA-2"
   CODPARAM$(11) = "MEDIDA-3"
   CODPARAM$(12) = "N-PLANO"
   CODPARAM$(13) = "R-PLANO"
   CODPARAM$(14) = "EQUIPO1"
   CODPARAM$(15) = "EQUIPO2"
   CODPARAM$(16) = "COD-MAT"
   CODPARAM$(17) = "DES-MAT"
   CODPARAM$(18) = "F-PLANO"
   CODPARAM$(19) = "Z-PLANO"
   CODPARAM$(20) = "POR-DESC"
   CODPARAM$(21) = "IMP-DESC"
   CAPARDOC% = 21
   '
   DOCPARAM$(1) = TRIM$(Text1(6))
   DOCPARAM$(2) = TRIM$(Label2(0))
   DOCPARAM$(3) = TRIM$(Label2(1))
   DOCPARAM$(4) = TRIM$(Label2(2))
   DOCPARAM$(5) = TRIM$(Label2(3))
   DOCPARAM$(6) = TRIM$(Text1(6))
   DOCPARAM$(7) = ""
   DOCPARAM$(8) = TRIM$(Label2(4))
   DOCPARAM$(9) = TRIM$(Label2(5))
   DOCPARAM$(10) = TRIM$(Label2(6))
   DOCPARAM$(11) = TRIM$(Label2(7))
   DOCPARAM$(12) = TRIM$(Label2(8))
   DOCPARAM$(13) = TRIM$(Text1(5))
   On Error Resume Next
     DOCPARAM$(14) = List2.List(0)
     DOCPARAM$(15) = List2.List(1)
   On Error GoTo 0
   DOCPARAM$(16) = TRIM$(Label37)
   DOCPARAM$(17) = TRIM$(Label14)
   DOCPARAM$(18) = TRIM$(Label7)
   DOCPARAM$(19) = TRIM$(Label18)
   DOCPARAM$(20) = TRIM$(Label24)
   DOCPARAM$(21) = TRIM$(Label21)
   '
   CODTABU1$(1) = "N-HPROF"
   CODTABU1$(2) = "N-HINSP"
   CODTABU1$(3) = "N-ORIGIN"
   CODTABU1$(4) = "N-OCOMPR"
   CODTABU1$(5) = "N-AUTENT"
   CODTABU1$(6) = "N-REMITO"
   CODTABU1$(7) = "N-BOLREC"
   CODTABU1$(8) = "POR-IVA"
   CODTABU1$(9) = "IMP-IVA"
   CACOLTAB1% = 9
   '
   CAITAB1% = 0
   For U& = 1 To List1.ListCount
     Z$ = List1.List(U& - 1)
     CAITAB1% = CAITAB1% + 1
     TETABU1$(1, CAITAB1%) = Mid$(Z$, 1, 12)
     TETABU1$(2, CAITAB1%) = Mid$(Z$, 13, 9)
     TETABU1$(3, CAITAB1%) = Mid$(Z$, 22, 9)
     TETABU1$(4, CAITAB1%) = Mid$(Z$, 31, 9)
     TETABU1$(5, CAITAB1%) = Mid$(Z$, 40, 9)
     TETABU1$(6, CAITAB1%) = Mid$(Z$, 49, 9)
     TETABU1$(7, CAITAB1%) = Mid$(Z$, 58, 9)
     TETABU1$(8, CAITAB1%) = Mid$(Z$, 67, 9)
     TETABU1$(9, CAITAB1%) = Mid$(Z$, 76, 9)
   Next U&
   '
   Call PRINTDOC("@" + FORIPRE$)
   Screen.MousePointer = 1
   '
End Sub
