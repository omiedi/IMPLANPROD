VERSION 5.00
Begin VB.Form FORDEMO 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8880
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   11970
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   7620
      ScaleHeight     =   405
      ScaleWidth      =   4305
      TabIndex        =   1
      Top             =   0
      Width           =   4370
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   750
         Left            =   1350
         TabIndex        =   43
         Top             =   0
         Width           =   3480
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   395
            Index           =   6
            Left            =   2555
            Picture         =   "FORDEMO.frx":0000
            Stretch         =   -1  'True
            ToolTipText     =   "Más Información (Si Disponible)"
            Top             =   0
            Width           =   395
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   395
            Index           =   5
            Left            =   2190
            Picture         =   "FORDEMO.frx":030A
            Stretch         =   -1  'True
            ToolTipText     =   "Imprime Pantalla"
            Top             =   0
            Width           =   395
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   395
            Index           =   4
            Left            =   1825
            Picture         =   "FORDEMO.frx":040C
            Stretch         =   -1  'True
            ToolTipText     =   "Finaliza Secuencia de Demostración"
            Top             =   0
            Width           =   395
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   395
            Index           =   3
            Left            =   1460
            Picture         =   "FORDEMO.frx":0716
            Stretch         =   -1  'True
            ToolTipText     =   "Pausa de Ejecución"
            Top             =   0
            Width           =   395
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Index           =   7
            Left            =   1095
            Picture         =   "FORDEMO.frx":0A20
            Stretch         =   -1  'True
            ToolTipText     =   "Acceso Directo a Secuencia"
            Top             =   0
            Width           =   390
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   395
            Index           =   2
            Left            =   730
            Picture         =   "FORDEMO.frx":0D2A
            Stretch         =   -1  'True
            ToolTipText     =   "Pantalla Anterior"
            Top             =   0
            Width           =   395
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   395
            Index           =   1
            Left            =   365
            Picture         =   "FORDEMO.frx":1034
            Stretch         =   -1  'True
            ToolTipText     =   "Pantalla Siguiente"
            Top             =   0
            Width           =   395
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   395
            Index           =   0
            Left            =   0
            Picture         =   "FORDEMO.frx":133E
            Stretch         =   -1  'True
            ToolTipText     =   "Comienza Secuencia de Demostración"
            Top             =   0
            Width           =   395
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   6000
            Y1              =   395
            Y2              =   395
         End
         Begin VB.Line Line2 
            X1              =   2940
            X2              =   2940
            Y1              =   0
            Y2              =   735
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   435
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   1485
         Begin VB.Frame Frame2 
            Caption         =   "Velocidad"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   1380
            Begin VB.HScrollBar HScroll1 
               Height          =   120
               LargeChange     =   10
               Left            =   210
               Max             =   500
               Min             =   30
               TabIndex        =   46
               Top             =   190
               Value           =   240
               Width           =   960
            End
         End
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      Height          =   1380
      Index           =   9
      Left            =   1890
      ScaleHeight     =   1320
      ScaleWidth      =   5520
      TabIndex        =   35
      Top             =   5040
      Visible         =   0   'False
      Width           =   5580
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000009&
         Height          =   330
         Index           =   9
         Left            =   -30
         ScaleHeight     =   270
         ScaleWidth      =   11940
         TabIndex        =   36
         Top             =   -25
         Width           =   12000
         Begin VB.Label Label2 
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   9
            Left            =   15
            TabIndex        =   37
            Top             =   10
            Width           =   11985
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   9
         Left            =   45
         TabIndex        =   38
         Top             =   325
         Width           =   3270
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line4 
         Index           =   9
         X1              =   5505
         X2              =   5500
         Y1              =   0
         Y2              =   5000
      End
      Begin VB.Line Line3 
         Index           =   9
         X1              =   0
         X2              =   12000
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   9
         X1              =   10
         X2              =   10
         Y1              =   210
         Y2              =   8000
      End
      Begin VB.Line Line6 
         Index           =   9
         X1              =   12000
         X2              =   0
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      Height          =   1380
      Index           =   8
      Left            =   1680
      ScaleHeight     =   1320
      ScaleWidth      =   5520
      TabIndex        =   31
      Top             =   4515
      Visible         =   0   'False
      Width           =   5580
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000009&
         Height          =   330
         Index           =   8
         Left            =   -30
         ScaleHeight     =   270
         ScaleWidth      =   11940
         TabIndex        =   32
         Top             =   -25
         Width           =   12000
         Begin VB.Label Label2 
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   8
            Left            =   15
            TabIndex        =   33
            Top             =   10
            Width           =   11985
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   8
         Left            =   45
         TabIndex        =   34
         Top             =   325
         Width           =   3270
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line4 
         Index           =   8
         X1              =   5505
         X2              =   5500
         Y1              =   0
         Y2              =   5000
      End
      Begin VB.Line Line3 
         Index           =   8
         X1              =   0
         X2              =   12000
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   8
         X1              =   10
         X2              =   10
         Y1              =   210
         Y2              =   8000
      End
      Begin VB.Line Line6 
         Index           =   8
         X1              =   12000
         X2              =   0
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      Height          =   1380
      Index           =   7
      Left            =   1470
      ScaleHeight     =   1320
      ScaleWidth      =   5520
      TabIndex        =   27
      Top             =   3990
      Visible         =   0   'False
      Width           =   5580
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000009&
         Height          =   330
         Index           =   7
         Left            =   -30
         ScaleHeight     =   270
         ScaleWidth      =   11940
         TabIndex        =   28
         Top             =   -25
         Width           =   12000
         Begin VB.Label Label2 
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   7
            Left            =   15
            TabIndex        =   29
            Top             =   10
            Width           =   11985
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   7
         Left            =   45
         TabIndex        =   30
         Top             =   325
         Width           =   3270
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line4 
         Index           =   7
         X1              =   5505
         X2              =   5500
         Y1              =   0
         Y2              =   5000
      End
      Begin VB.Line Line3 
         Index           =   7
         X1              =   0
         X2              =   12000
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   7
         X1              =   10
         X2              =   10
         Y1              =   210
         Y2              =   8000
      End
      Begin VB.Line Line6 
         Index           =   7
         X1              =   12000
         X2              =   0
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      Height          =   1380
      Index           =   6
      Left            =   1260
      ScaleHeight     =   1320
      ScaleWidth      =   5520
      TabIndex        =   23
      Top             =   3465
      Visible         =   0   'False
      Width           =   5580
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000009&
         Height          =   330
         Index           =   6
         Left            =   -30
         ScaleHeight     =   270
         ScaleWidth      =   11940
         TabIndex        =   24
         Top             =   -25
         Width           =   12000
         Begin VB.Label Label2 
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   6
            Left            =   15
            TabIndex        =   25
            Top             =   10
            Width           =   11985
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   6
         Left            =   45
         TabIndex        =   26
         Top             =   325
         Width           =   3270
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line4 
         Index           =   6
         X1              =   5505
         X2              =   5500
         Y1              =   0
         Y2              =   5000
      End
      Begin VB.Line Line3 
         Index           =   6
         X1              =   0
         X2              =   12000
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   6
         X1              =   10
         X2              =   10
         Y1              =   210
         Y2              =   8000
      End
      Begin VB.Line Line6 
         Index           =   6
         X1              =   12000
         X2              =   0
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      Height          =   1380
      Index           =   5
      Left            =   1050
      ScaleHeight     =   1320
      ScaleWidth      =   5520
      TabIndex        =   19
      Top             =   2940
      Visible         =   0   'False
      Width           =   5580
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000009&
         Height          =   330
         Index           =   5
         Left            =   -30
         ScaleHeight     =   270
         ScaleWidth      =   11940
         TabIndex        =   20
         Top             =   -25
         Width           =   12000
         Begin VB.Label Label2 
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   15
            TabIndex        =   21
            Top             =   10
            Width           =   11985
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   5
         Left            =   45
         TabIndex        =   22
         Top             =   325
         Width           =   3270
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line4 
         Index           =   5
         X1              =   5505
         X2              =   5500
         Y1              =   0
         Y2              =   5000
      End
      Begin VB.Line Line3 
         Index           =   5
         X1              =   0
         X2              =   12000
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   5
         X1              =   10
         X2              =   10
         Y1              =   210
         Y2              =   8000
      End
      Begin VB.Line Line6 
         Index           =   5
         X1              =   12000
         X2              =   0
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      Height          =   1380
      Index           =   4
      Left            =   840
      ScaleHeight     =   1320
      ScaleWidth      =   5520
      TabIndex        =   15
      Top             =   2415
      Visible         =   0   'False
      Width           =   5580
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000009&
         Height          =   330
         Index           =   4
         Left            =   -30
         ScaleHeight     =   270
         ScaleWidth      =   11940
         TabIndex        =   16
         Top             =   -25
         Width           =   12000
         Begin VB.Label Label2 
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   15
            TabIndex        =   17
            Top             =   10
            Width           =   11985
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   4
         Left            =   45
         TabIndex        =   18
         Top             =   325
         Width           =   3270
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line4 
         Index           =   4
         X1              =   5505
         X2              =   5500
         Y1              =   0
         Y2              =   5000
      End
      Begin VB.Line Line3 
         Index           =   4
         X1              =   0
         X2              =   12000
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   4
         X1              =   10
         X2              =   10
         Y1              =   210
         Y2              =   8000
      End
      Begin VB.Line Line6 
         Index           =   4
         X1              =   12000
         X2              =   0
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      Height          =   1380
      Index           =   3
      Left            =   630
      ScaleHeight     =   1320
      ScaleWidth      =   5520
      TabIndex        =   39
      Top             =   1890
      Visible         =   0   'False
      Width           =   5580
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000009&
         Height          =   330
         Index           =   3
         Left            =   -30
         ScaleHeight     =   270
         ScaleWidth      =   11940
         TabIndex        =   40
         Top             =   -25
         Width           =   12000
         Begin VB.Label Label2 
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   15
            TabIndex        =   41
            Top             =   10
            Width           =   11985
         End
      End
      Begin VB.Line Line6 
         Index           =   3
         X1              =   12000
         X2              =   0
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   3
         X1              =   10
         X2              =   10
         Y1              =   210
         Y2              =   8000
      End
      Begin VB.Line Line3 
         Index           =   3
         X1              =   0
         X2              =   12000
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line Line4 
         Index           =   3
         X1              =   5505
         X2              =   5500
         Y1              =   0
         Y2              =   5000
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   3
         Left            =   45
         TabIndex        =   42
         Top             =   325
         Width           =   3270
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      Height          =   1380
      Index           =   2
      Left            =   420
      ScaleHeight     =   1320
      ScaleWidth      =   5520
      TabIndex        =   11
      Top             =   1365
      Visible         =   0   'False
      Width           =   5580
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000009&
         Height          =   330
         Index           =   2
         Left            =   -30
         ScaleHeight     =   270
         ScaleWidth      =   11940
         TabIndex        =   12
         Top             =   -25
         Width           =   12000
         Begin VB.Label Label2 
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   15
            TabIndex        =   13
            Top             =   10
            Width           =   11985
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   2
         Left            =   45
         TabIndex        =   14
         Top             =   325
         Width           =   3270
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line4 
         Index           =   2
         X1              =   5505
         X2              =   5500
         Y1              =   0
         Y2              =   5000
      End
      Begin VB.Line Line3 
         Index           =   2
         X1              =   0
         X2              =   12000
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   2
         X1              =   10
         X2              =   10
         Y1              =   210
         Y2              =   8000
      End
      Begin VB.Line Line6 
         Index           =   2
         X1              =   12000
         X2              =   0
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      Height          =   1380
      Index           =   1
      Left            =   1365
      ScaleHeight     =   1320
      ScaleWidth      =   5520
      TabIndex        =   7
      Top             =   1785
      Visible         =   0   'False
      Width           =   5580
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000009&
         Height          =   330
         Index           =   1
         Left            =   -30
         ScaleHeight     =   270
         ScaleWidth      =   11940
         TabIndex        =   8
         Top             =   -25
         Width           =   12000
         Begin VB.Label Label2 
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   15
            TabIndex        =   9
            Top             =   10
            Width           =   11985
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   1
         Left            =   45
         TabIndex        =   10
         Top             =   315
         Width           =   3270
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   5505
         X2              =   5500
         Y1              =   0
         Y2              =   5000
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   0
         X2              =   12000
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   1
         X1              =   10
         X2              =   10
         Y1              =   210
         Y2              =   8000
      End
      Begin VB.Line Line6 
         Index           =   1
         X1              =   12000
         X2              =   0
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      Height          =   1380
      Index           =   0
      Left            =   2205
      ScaleHeight     =   1320
      ScaleWidth      =   5520
      TabIndex        =   2
      Top             =   2205
      Visible         =   0   'False
      Width           =   5580
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000009&
         Height          =   330
         Index           =   0
         Left            =   -30
         ScaleHeight     =   270
         ScaleWidth      =   11940
         TabIndex        =   5
         Top             =   -25
         Width           =   12000
         Begin VB.Label Label2 
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   15
            TabIndex        =   6
            Top             =   10
            Width           =   11985
         End
      End
      Begin VB.Line Line6 
         Index           =   0
         X1              =   12000
         X2              =   0
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   10
         X2              =   10
         Y1              =   210
         Y2              =   8000
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   0
         X2              =   12000
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   5505
         X2              =   5500
         Y1              =   0
         Y2              =   5000
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   0
         Left            =   45
         TabIndex        =   3
         Top             =   325
         Width           =   3270
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Line PUNTB 
      BorderWidth     =   3
      Index           =   9
      Visible         =   0   'False
      X1              =   0
      X2              =   315
      Y1              =   210
      Y2              =   420
   End
   Begin VB.Line PUNTA 
      BorderWidth     =   3
      Index           =   9
      Visible         =   0   'False
      X1              =   105
      X2              =   420
      Y1              =   0
      Y2              =   210
   End
   Begin VB.Line PUNTB 
      BorderWidth     =   3
      Index           =   8
      Visible         =   0   'False
      X1              =   0
      X2              =   315
      Y1              =   210
      Y2              =   420
   End
   Begin VB.Line PUNTA 
      BorderWidth     =   3
      Index           =   8
      Visible         =   0   'False
      X1              =   105
      X2              =   420
      Y1              =   0
      Y2              =   210
   End
   Begin VB.Line PUNTB 
      BorderWidth     =   3
      Index           =   7
      Visible         =   0   'False
      X1              =   0
      X2              =   315
      Y1              =   210
      Y2              =   420
   End
   Begin VB.Line PUNTA 
      BorderWidth     =   3
      Index           =   7
      Visible         =   0   'False
      X1              =   105
      X2              =   420
      Y1              =   0
      Y2              =   210
   End
   Begin VB.Line PUNTB 
      BorderWidth     =   3
      Index           =   6
      Visible         =   0   'False
      X1              =   0
      X2              =   315
      Y1              =   210
      Y2              =   420
   End
   Begin VB.Line PUNTA 
      BorderWidth     =   3
      Index           =   6
      Visible         =   0   'False
      X1              =   105
      X2              =   420
      Y1              =   0
      Y2              =   210
   End
   Begin VB.Line PUNTB 
      BorderWidth     =   3
      Index           =   5
      Visible         =   0   'False
      X1              =   0
      X2              =   315
      Y1              =   210
      Y2              =   420
   End
   Begin VB.Line PUNTA 
      BorderWidth     =   3
      Index           =   5
      Visible         =   0   'False
      X1              =   105
      X2              =   420
      Y1              =   0
      Y2              =   210
   End
   Begin VB.Line PUNTB 
      BorderWidth     =   3
      Index           =   4
      Visible         =   0   'False
      X1              =   0
      X2              =   315
      Y1              =   210
      Y2              =   420
   End
   Begin VB.Line PUNTA 
      BorderWidth     =   3
      Index           =   4
      Visible         =   0   'False
      X1              =   105
      X2              =   420
      Y1              =   0
      Y2              =   210
   End
   Begin VB.Line PUNTB 
      BorderWidth     =   3
      Index           =   3
      Visible         =   0   'False
      X1              =   0
      X2              =   315
      Y1              =   210
      Y2              =   420
   End
   Begin VB.Line PUNTA 
      BorderWidth     =   3
      Index           =   3
      Visible         =   0   'False
      X1              =   105
      X2              =   420
      Y1              =   0
      Y2              =   210
   End
   Begin VB.Line PUNTB 
      BorderWidth     =   3
      Index           =   2
      Visible         =   0   'False
      X1              =   0
      X2              =   315
      Y1              =   210
      Y2              =   420
   End
   Begin VB.Line PUNTA 
      BorderWidth     =   3
      Index           =   2
      Visible         =   0   'False
      X1              =   105
      X2              =   420
      Y1              =   0
      Y2              =   210
   End
   Begin VB.Line PUNTB 
      BorderWidth     =   3
      Index           =   1
      Visible         =   0   'False
      X1              =   0
      X2              =   315
      Y1              =   210
      Y2              =   420
   End
   Begin VB.Line PUNTA 
      BorderWidth     =   3
      Index           =   1
      Visible         =   0   'False
      X1              =   105
      X2              =   420
      Y1              =   0
      Y2              =   210
   End
   Begin VB.Line PUNTB 
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   4410
      X2              =   4725
      Y1              =   735
      Y2              =   945
   End
   Begin VB.Line PUNTA 
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   4515
      X2              =   4830
      Y1              =   525
      Y2              =   735
   End
   Begin VB.Line TRAMO1 
      BorderWidth     =   2
      Index           =   9
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line TRAMO2 
      BorderWidth     =   2
      Index           =   9
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line TRAMO3 
      BorderWidth     =   2
      Index           =   9
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line SOMBRA1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   9
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   525
      Y2              =   105
   End
   Begin VB.Line SOMBRA2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   9
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   840
      Y2              =   420
   End
   Begin VB.Line SOMBRA3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   9
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   1155
      Y2              =   735
   End
   Begin VB.Line TRAMO1 
      BorderWidth     =   2
      Index           =   8
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line TRAMO2 
      BorderWidth     =   2
      Index           =   8
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line TRAMO3 
      BorderWidth     =   2
      Index           =   8
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line SOMBRA1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   8
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   525
      Y2              =   105
   End
   Begin VB.Line SOMBRA2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   8
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   840
      Y2              =   420
   End
   Begin VB.Line SOMBRA3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   8
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   1155
      Y2              =   735
   End
   Begin VB.Line TRAMO1 
      BorderWidth     =   2
      Index           =   7
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line TRAMO2 
      BorderWidth     =   2
      Index           =   7
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line TRAMO3 
      BorderWidth     =   2
      Index           =   7
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line SOMBRA1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   7
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   525
      Y2              =   105
   End
   Begin VB.Line SOMBRA2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   7
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   840
      Y2              =   420
   End
   Begin VB.Line SOMBRA3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   7
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   1155
      Y2              =   735
   End
   Begin VB.Line TRAMO1 
      BorderWidth     =   2
      Index           =   6
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line TRAMO2 
      BorderWidth     =   2
      Index           =   6
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line TRAMO3 
      BorderWidth     =   2
      Index           =   6
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line SOMBRA1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   6
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   525
      Y2              =   105
   End
   Begin VB.Line SOMBRA2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   6
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   840
      Y2              =   420
   End
   Begin VB.Line SOMBRA3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   6
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   1155
      Y2              =   735
   End
   Begin VB.Line TRAMO1 
      BorderWidth     =   2
      Index           =   5
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line TRAMO2 
      BorderWidth     =   2
      Index           =   5
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line TRAMO3 
      BorderWidth     =   2
      Index           =   5
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line SOMBRA1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   5
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   525
      Y2              =   105
   End
   Begin VB.Line SOMBRA2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   5
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   840
      Y2              =   420
   End
   Begin VB.Line SOMBRA3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   5
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   1155
      Y2              =   735
   End
   Begin VB.Line TRAMO1 
      BorderWidth     =   2
      Index           =   4
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line TRAMO2 
      BorderWidth     =   2
      Index           =   4
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line TRAMO3 
      BorderWidth     =   2
      Index           =   4
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line SOMBRA1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   4
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   525
      Y2              =   105
   End
   Begin VB.Line SOMBRA2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   4
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   840
      Y2              =   420
   End
   Begin VB.Line SOMBRA3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   4
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   1155
      Y2              =   735
   End
   Begin VB.Line TRAMO1 
      BorderWidth     =   2
      Index           =   3
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line TRAMO2 
      BorderWidth     =   2
      Index           =   3
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line TRAMO3 
      BorderWidth     =   2
      Index           =   3
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line SOMBRA1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   3
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   525
      Y2              =   105
   End
   Begin VB.Line SOMBRA2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   3
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   840
      Y2              =   420
   End
   Begin VB.Line SOMBRA3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   3
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   1155
      Y2              =   735
   End
   Begin VB.Line TRAMO1 
      BorderWidth     =   2
      Index           =   2
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line TRAMO2 
      BorderWidth     =   2
      Index           =   2
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line TRAMO3 
      BorderWidth     =   2
      Index           =   2
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line SOMBRA1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   2
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   525
      Y2              =   105
   End
   Begin VB.Line SOMBRA2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   2
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   840
      Y2              =   420
   End
   Begin VB.Line SOMBRA3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   2
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   1155
      Y2              =   735
   End
   Begin VB.Line TRAMO1 
      BorderWidth     =   2
      Index           =   1
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line TRAMO2 
      BorderWidth     =   2
      Index           =   1
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line TRAMO3 
      BorderWidth     =   2
      Index           =   1
      Visible         =   0   'False
      X1              =   1155
      X2              =   2205
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line SOMBRA1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   1
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   525
      Y2              =   105
   End
   Begin VB.Line SOMBRA2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   1
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   840
      Y2              =   420
   End
   Begin VB.Line SOMBRA3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   1
      Visible         =   0   'False
      X1              =   1365
      X2              =   2100
      Y1              =   1155
      Y2              =   735
   End
   Begin VB.Line TRAMO3 
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   6615
      X2              =   7665
      Y1              =   1785
      Y2              =   1785
   End
   Begin VB.Line TRAMO2 
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   6615
      X2              =   7665
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Line TRAMO1 
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   6615
      X2              =   7665
      Y1              =   1155
      Y2              =   1155
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   1260
      TabIndex        =   4
      Top             =   1155
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11985
   End
   Begin VB.Line SOMBRA3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   6825
      X2              =   7560
      Y1              =   1995
      Y2              =   1575
   End
   Begin VB.Line SOMBRA2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   6825
      X2              =   7560
      Y1              =   1680
      Y2              =   1260
   End
   Begin VB.Line SOMBRA1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   6825
      X2              =   7560
      Y1              =   1365
      Y2              =   945
   End
   Begin VB.Image Image1 
      Height          =   8505
      Left            =   0
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "FORDEMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   Call CENTRAFORM(Me)
End Sub

Private Sub Image1_DblClick()
   If App.EXEName = "GENDEMO" Then
      Hide
   End If
End Sub

Private Sub Image2_Click(Index As Integer)
   '
   Image2(Index).Visible = False
   TINIX = Timer
   If Index% = 0 Then
      COMACON$ = "HOME"
    ElseIf Index% = 1 Then
      COMACON$ = "NEXT"
    ElseIf Index = 2 Then
      COMACON$ = "PREV"
    ElseIf Index = 3 Then
      COMACON$ = "WAIT"
    ElseIf Index = 4 Then
      COMACON$ = "END"
    ElseIf Index = 5 Then
      Picture1.Visible = False
      PrintForm
      Picture1.Visible = True
    ElseIf Index = 7 Then
      COMACON$ = "POINT"
   End If
   '
99 While Timer < TINIX + 0.2
     DoEvents
   Wend
   Image2(Index).Visible = True
   '
End Sub

