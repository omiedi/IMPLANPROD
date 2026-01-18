VERSION 5.00
Begin VB.Form CONPED04 
   BackColor       =   &H00E0FFD0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Pedidos por Número"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture6 
      Height          =   8295
      Left            =   10920
      ScaleHeight     =   8235
      ScaleWidth      =   1215
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
      Begin VB.Image Image7 
         Height          =   8730
         Left            =   0
         Picture         =   "CONPED04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command43 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1420
      Left            =   5400
      Picture         =   "CONPED04.frx":2398
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Actualizar Indice de Selección"
      Top             =   325
      Width           =   285
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   10920
      ScaleHeight     =   8355
      ScaleWidth      =   1110
      TabIndex        =   38
      Top             =   0
      Width           =   1170
      Begin VB.CommandButton Command6 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   150
         Picture         =   "CONPED04.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Consultas y Re-Impresión"
         Top             =   3360
         Width           =   650
      End
      Begin VB.CommandButton Command26 
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
         Height          =   650
         Left            =   150
         Picture         =   "CONPED04.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Configurar Comprobante de Despacho"
         Top             =   2520
         Width           =   650
      End
      Begin VB.CommandButton Command9 
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
         Height          =   650
         Left            =   150
         Picture         =   "CONPED04.frx":2C2E
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Cerrar y Salir"
         Top             =   300
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Height          =   650
         Left            =   150
         Picture         =   "CONPED04.frx":2D78
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Ayuda en Linea"
         Top             =   950
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   140
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   40
         Top             =   7440
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   0
            TabIndex        =   41
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   150
         Picture         =   "CONPED04.frx":3082
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Guarda la Configuracion"
         Top             =   1600
         Width           =   650
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -250
         Picture         =   "CONPED04.frx":338C
         Top             =   7800
         Width           =   1515
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Pedido / Orden de Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1550
      Left            =   105
      TabIndex        =   17
      Top             =   210
      Width           =   5310
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2205
         TabIndex        =   20
         Top             =   315
         Width           =   180
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4830
         TabIndex        =   33
         Top             =   315
         Width           =   330
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "St:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3465
         TabIndex        =   32
         Top             =   380
         Width           =   1275
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3360
         TabIndex        =   31
         Top             =   1125
         Width           =   1800
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O/C:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1995
         TabIndex        =   30
         Top             =   1170
         Width           =   1275
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1155
         TabIndex        =   26
         Top             =   1125
         Width           =   1605
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente::"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   25
         Top             =   765
         Width           =   750
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1155
         TabIndex        =   24
         Top             =   720
         Width           =   4020
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         TabIndex        =   23
         Top             =   380
         Width           =   1275
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1155
         TabIndex        =   22
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3255
         TabIndex        =   21
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Venta/Rep:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   -840
         TabIndex        =   19
         Top             =   1170
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   -840
         TabIndex        =   18
         Top             =   380
         Width           =   1905
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1550
      Left            =   5700
      TabIndex        =   14
      Top             =   210
      Width           =   5025
      Begin VB.Label LENTREGA 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   1155
         TabIndex        =   36
         Top             =   1125
         Width           =   3705
      End
      Begin VB.Label LENTREGA 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   1155
         TabIndex        =   35
         Top             =   720
         Width           =   3705
      End
      Begin VB.Label LENTREGA 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   1155
         TabIndex        =   34
         Top             =   315
         Width           =   3705
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   -840
         TabIndex        =   16
         Top             =   380
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transporte:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   -840
         TabIndex        =   15
         Top             =   1170
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   105
      TabIndex        =   3
      Top             =   6300
      Width           =   10620
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1175
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   735
         Width           =   7155
      End
      Begin VB.CommandButton Command3 
         Caption         =   "O.K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   745
         Left            =   8360
         Picture         =   "CONPED04.frx":52AE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1140
         Width           =   750
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         Height          =   425
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   10515
         TabIndex        =   5
         Top             =   315
         Width           =   10580
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            Height          =   15
            Left            =   0
            TabIndex        =   7
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   6
            Top             =   75
            Width           =   10200
         End
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   10665
         Y1              =   1900
         Y2              =   1900
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
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
         Left            =   8400
         TabIndex        =   29
         Top             =   840
         Width           =   4000
      End
      Begin VB.Line Line7 
         X1              =   10580
         X2              =   10580
         Y1              =   315
         Y2              =   1900
      End
      Begin VB.Line Line6 
         X1              =   8295
         X2              =   10605
         Y1              =   1130
         Y2              =   1130
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Registrar Anotación"
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
         Left            =   9345
         TabIndex        =   28
         Top             =   1310
         Width           =   960
      End
      Begin VB.Line Line5 
         X1              =   8295
         X2              =   8295
         Y1              =   735
         Y2              =   1995
      End
      Begin VB.Label Label9 
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
         Height          =   1165
         Left            =   0
         TabIndex        =   8
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label10 
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
         Height          =   1170
         Left            =   8310
         TabIndex        =   9
         Top             =   735
         Width           =   2260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Anotaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4450
      Left            =   105
      TabIndex        =   1
      Top             =   1800
      Width           =   10620
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   3795
         Left            =   10290
         TabIndex        =   37
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4005
         Left            =   8085
         TabIndex        =   13
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4110
         Left            =   1155
         TabIndex        =   12
         Top             =   630
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
         Height          =   3660
         Left            =   0
         TabIndex        =   11
         Top             =   735
         Width           =   10580
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10515
         TabIndex        =   2
         Top             =   315
         Width           =   10580
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   27
            Top             =   105
            Width           =   10200
         End
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10605
         Y1              =   4410
         Y2              =   4410
      End
      Begin VB.Line Line1 
         X1              =   10580
         X2              =   10580
         Y1              =   320
         Y2              =   6225
      End
   End
   Begin VB.Label Label1 
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
      Left            =   -105
      TabIndex        =   0
      Top             =   2625
      Width           =   1695
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu ver 
      Caption         =   "Ver"
      Begin VB.Menu visualizarpedidos 
         Caption         =   "Visualizar Pedidos"
      End
   End
   Begin VB.Menu config 
      Caption         =   "Configuración"
      Begin VB.Menu configuracion 
         Caption         =   "Configuración"
      End
   End
End
Attribute VB_Name = "CONPED04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command26_Click()
   Command26.Enabled = False
   CFGPED04.Show 1
   Command26.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   NPX& = XVALO(Label7)
   If NPX& > 0 Then
     If TRIM$(Text1) <> "" Then
       Call FRATEXTO(Text1, 56)
       
       For KKI% = 1 To CARETEX%
         TTX1$ = Space$(76)
         Call REPLA(TTX1$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTX1$, HOS$, 1, 8)
           Call REPLA(TTX1$, USERTER$, 69, 24)
         End If
         TTXX$ = TTXX$ + TTX1$ + Chr$(10) + Chr$(13)
       Next KKI%
       '
       '\\\OT-08-0098-12/03/08-///
       Call APERBASDAT("PEDCLI")
       SQLX$ = "SELECT ANOTAPED FROM PEDENCA"
       SQLX$ = SQLX$ + " WHERE NroPed = " & NPX& & ""
       Set PEDDDENCA = Ventas.OpenRecordset(SQLX$, dbOpenDynaset)
       ANOTA$ = PEDDDENCA!AnotaPed
       ANOTA$ = ANOTA$ + TTXX$
       ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
       
       SQLX$ = "UPDATE PEDENCA SET AnotaPed=" & "'" & ANOTA$ & "' WHERE NroPed=" & NPX&
       Ventas.Execute (SQLX$)
       Text1 = ""
       '
       List1.Clear
       Label7 = "": Label7 = NPX&
       '
     End If
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
3  Call SELECHO("INDIPEDI/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
   NPX& = XVALO(VALACT1$("INDIPEDI"))
   If NPX& < 1 Then GoTo 99
   '
   Label7 = TRIM$(Str$(NPX&))
   '
99 Command4.Enabled = True
End Sub

Private Sub Command43_Click()
    Command43.Enabled = False
    Screen.MousePointer = 11
    '
    SQLX$ = "SELECT * FROM PEDENCA ORDER BY NroPed ASC"
    Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With PEDENCTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
             NRODES& = !NroPed
             TTXX$ = MKS$(NRODES&)
               NOCO$ = !NroOcom
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIE% = !NroClie
               ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
             TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + ANULX$ 'Space$(52)
             JJ& = JJ& + 1
             Call GRAREG("INDIPEDI", TTXX$, JJ&)
             '\\\OT-05-0682-WAR-FIN///
        .MoveNext
        Loop
      End If
    End With
    '
    Call SETULTREG("INDIPEDI", JJ&)
    Call CIERRARCH("INDIPEDI")
    Call FRESHECHO("INDIPEDI")
    Screen.MousePointer = 1
    Command43.Enabled = True
    
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   NPX& = XVALO(Label7)
   '
   If NPX& > 0 Then
      TIDOCUM$ = "Pedido de Cliente"
      If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
        Call MENSERR(24, "Usuario No Autorizado para Visualizar Pedidos.\La Consulta Cambia a Orden de Despacho")
        TIDOCUM$ = "Orden de Despacho"
      End If
      NUDOCU$ = TRIM$(Str$(NPX&))
      While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
      NUDOCU$ = NUDOCU$ + " "
      Screen.MousePointer = 11
      '
      For UI& = ULTREG&("PDOCLST") To 1 Step -1
        XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
        If InStr(XX$, TIDOCUM$) > 4 Then
          PPXX% = InStr(XX$, "Nro.")
          If PPXX% > 0 Then
            XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
          End If
          If InStr(XX$, NUDOCU$) > 4 Then
            DOCUNU& = CVS(Left$(XX$, 4))
            GoTo 5
          End If
        End If
      Next UI&
      '
      Screen.MousePointer = 1
      Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
      GoTo 99
      '
5     Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
      
    End If
    '
99  Command6.Enabled = True
    '
End Sub

Private Sub Command9_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub configuracion_Click()
    Call Command26_Click
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      '
    End If

    Call CENTRAFORM(Me)
    '
    DoEvents
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Screen.MousePointer = 1
    HOII% = HOY(HOS$)
    Label9 = " " + HOS$
    Label12 = Left$(USERNAME, 17)
    '
End Sub

Private Sub Label7_Change()
   List1.Clear
   NPX& = XVALO(Label7)
   If NPX& < 1 Then Exit Sub
   '
20 Call APERBASDAT("PEDCLI")
   SQLX$ = "SELECT * FROM PEDENCA"
   SQLX$ = SQLX$ + " WHERE NroPed=" & NPX&
   Set CABATEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With CABATEMP
    On Error Resume Next
    .MoveFirst
    If Err < 1 Then
        On Error GoTo 0
        XX$ = SAFETEXT$(!AnotaPed): LONGI& = Len(XX$)
        Label11 = !RasoClie
        Label16 = !NroOcom
        FECEMI% = CVINT(!FECHEMIS)
        Label8 = FECHATEX$(FECEMI%)
        Label18 = !Status
        LENTREGA(0) = !DOMIENT
        LENTREGA(1) = !LOCAENT
        LENTREGA(2) = !TranspEnt
        REPARX% = 0: If UCase$(Left$(SAFETEXT$(!VenRep), 1)) = "R" Then REPARX% = 1
        '
        INICIO& = 1
        Do While INICIO& < LONGI&
            FIN& = InStr(INICIO&, XX$, Chr$(10))
            If FIN& > 0 Then
               TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
               List1.AddItem TTXX$
               INICIO& = FIN& + 2
            Else
               TTXX$ = Mid$(XX$, INICIO&)
               List1.AddItem TTXX$
               Exit Do
            End If
        Loop
    End If
    On Error GoTo 0
   End With
   List1.Refresh
   OPENFORMS = DoEvents
   CABATEMP.Close
   Set CABATEMP = Nothing
   '
   If REPARX% <> 1 Then ' NO ES UNA REPARACIóN
     Screen.MousePointer = 11
     HUBOREC% = 0
     If HUBORECU% > 0 Then HUBOREC% = 1
     Screen.MousePointer = 1
     If HUBOREC% > 0 Then
       Call COMUNI("Se Han Recuperado Acumulados\de Preparación y Despacho.")
     End If
   End If
   '
End Sub

Private Sub salir_Click()
   Call Command9_Click
End Sub

Private Sub visualizarpedidos_Click()
   Call Command6_Click
End Sub
