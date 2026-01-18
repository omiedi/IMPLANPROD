VERSION 5.00
Begin VB.Form FORFACDIP 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación en Cuenta Corriente "
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8200
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12000
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2745
         Left            =   10690
         TabIndex        =   44
         Top             =   1530
         Width           =   1170
         Begin VB.CommandButton Command4 
            Caption         =   "C.Cte."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   210
            Picture         =   "FORFACDIP.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   1785
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Notas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   210
            Picture         =   "FORFACDIP.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Notas de Clientes"
            Top             =   945
            Width           =   855
         End
         Begin VB.CommandButton Command11 
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
            Height          =   780
            Left            =   210
            Picture         =   "FORFACDIP.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
            Top             =   105
            Width           =   855
         End
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   10920
         ScaleHeight     =   75
         ScaleWidth      =   795
         TabIndex        =   42
         Top             =   1330
         Width           =   855
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.TextBox Text7 
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
         Left            =   5250
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   2640
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   5430
         Left            =   210
         ScaleHeight     =   5370
         ScaleWidth      =   11505
         TabIndex        =   17
         Top             =   1620
         Width           =   11565
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   1.11565e5
            Begin VB.Line Line10 
               X1              =   6700
               X2              =   6700
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line13 
               X1              =   10430
               X2              =   10430
               Y1              =   0
               Y2              =   315
            End
            Begin VB.Line Line1 
               X1              =   8925
               X2              =   8925
               Y1              =   0
               Y2              =   1155
            End
            Begin VB.Line Line2 
               X1              =   7840
               X2              =   7840
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   $"FORFACDIP.frx":0A56
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
               TabIndex        =   24
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line8 
               X1              =   630
               X2              =   630
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line9 
               X1              =   2520
               X2              =   2520
               Y1              =   0
               Y2              =   1050
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5630
            Left            =   0
            TabIndex        =   18
            Top             =   315
            Width           =   1.11565e5
            Begin VB.CommandButton Command6 
               Caption         =   "Add."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   10710
               Picture         =   "FORFACDIP.frx":0ADE
               TabIndex        =   38
               ToolTipText     =   "Agrega un Item al Final"
               Top             =   2415
               Width           =   580
            End
            Begin VB.CommandButton Command12 
               Caption         =   "Ins."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   10710
               Picture         =   "FORFACDIP.frx":0DE8
               TabIndex        =   37
               ToolTipText     =   "Intercala un Item Antes del Activo"
               Top             =   2970
               Width           =   580
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Del"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   10710
               Picture         =   "FORFACDIP.frx":10F2
               TabIndex        =   36
               ToolTipText     =   "Elimina el Item Activo "
               Top             =   3570
               Width           =   580
            End
            Begin VB.CommandButton Command8 
               Caption         =   "Edit"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   10710
               Picture         =   "FORFACDIP.frx":13FC
               TabIndex        =   35
               ToolTipText     =   "Eldición del Item Activo"
               Top             =   4095
               Width           =   580
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
               Height          =   4560
               Left            =   -30
               TabIndex        =   19
               Top             =   0
               Width           =   10450
            End
            Begin VB.Line Line6 
               X1              =   11490
               X2              =   11490
               Y1              =   2310
               Y2              =   5040
            End
            Begin VB.Line Line5 
               X1              =   10290
               X2              =   11550
               Y1              =   2310
               Y2              =   2310
            End
            Begin VB.Line Line14 
               BorderWidth     =   2
               X1              =   0
               X2              =   10395
               Y1              =   4560
               Y2              =   4560
            End
            Begin VB.Line Line24 
               X1              =   10425
               X2              =   10425
               Y1              =   0
               Y2              =   5040
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   8295
               TabIndex        =   26
               Top             =   4660
               Width           =   1695
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   0
               TabIndex        =   25
               Top             =   0
               Width           =   1695
            End
            Begin VB.Line Line4 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   5040
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   6615
               TabIndex        =   22
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   3570
               TabIndex        =   21
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Subtotal"
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
               Left            =   30
               TabIndex        =   20
               Top             =   4660
               Width           =   1695
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   11490
               Y1              =   5040
               Y2              =   5040
            End
         End
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   6405
         MaxLength       =   1
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   315
         Width           =   500
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
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
         Left            =   10920
         Picture         =   "FORFACDIP.frx":1706
         TabIndex        =   15
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   315
         Width           =   855
      End
      Begin VB.CommandButton BOTNEXT 
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
         Left            =   10670
         Picture         =   "FORFACDIP.frx":1A10
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Aprueba - Registra - Imprime Factura"
         Top             =   7140
         Width           =   1065
      End
      Begin VB.TextBox Text11 
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
         Left            =   9030
         TabIndex        =   4
         Text            =   " "
         Top             =   1155
         Width           =   1650
      End
      Begin VB.TextBox Text9 
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
         Left            =   9345
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   315
         Width           =   1170
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10500
         TabIndex        =   3
         Top             =   315
         Width           =   180
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
         Height          =   300
         Left            =   2310
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   2745
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
         Height          =   300
         Left            =   1470
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   750
      End
      Begin VB.TextBox Text3 
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
         Left            =   1470
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3585
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
         Height          =   300
         Left            =   1470
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   315
         Width           =   3585
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1050
         TabIndex        =   1
         Top             =   1155
         Width           =   180
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
         Height          =   300
         Left            =   210
         TabIndex        =   0
         Text            =   " "
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label Label2 
         Height          =   225
         Left            =   7245
         TabIndex        =   48
         Top             =   735
         Visible         =   0   'False
         Width           =   4530
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IVA-RNI:"
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
         TabIndex        =   40
         Top             =   7350
         Width           =   2220
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6195
         TabIndex        =   39
         Top             =   7245
         Width           =   1380
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8925
         TabIndex        =   34
         Top             =   7245
         Width           =   1380
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tot.Bruto:"
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
         Left            =   6615
         TabIndex        =   33
         Top             =   7350
         Width           =   2220
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3570
         TabIndex        =   32
         Top             =   7245
         Width           =   1380
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A.:"
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
         Left            =   2310
         TabIndex        =   31
         Top             =   7350
         Width           =   1170
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3780
         TabIndex        =   30
         Top             =   5775
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Valores:"
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
         Left            =   1470
         TabIndex        =   29
         Top             =   5880
         Width           =   2220
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         TabIndex        =   28
         Top             =   7245
         Width           =   1380
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tot,Neto:"
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
         Left            =   -1050
         TabIndex        =   27
         Top             =   7350
         Width           =   2220
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Numero:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7560
         TabIndex        =   14
         Top             =   1155
         Width           =   1380
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Left            =   8505
         TabIndex        =   13
         Top             =   315
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   315
         TabIndex        =   7
         Top             =   735
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pos.IVA:"
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
      TabIndex        =   12
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "FORFACDIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$


Private Sub BOTNEXT_Click()
   If ULTREG&("!CARCONPE") < 1 Then
       Call MENSERR(24, "Falta Ingresar Detalle de Factura")
       Exit Sub
   End If
   '
   OKEY% = 0
   Call GRAFACDIP(OKEY%)
   If OKEY% > 0 Then
      Call PRIFACDIP
      Call CLEARFORFACDIP
      Text1.SetFocus
   End If
   '
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        Text1.TEXT = TRIM$(Str$(NCXI%))
        NC% = NCXI%
        OKEY% = 0
        Call FACITEMI1(OKEY%)
        'Call CARDATPEDMAI(OKEY%)
        If OKEY% <> 1 Then
           Call CLEARFORFACDIP
           Exit Sub
        End If
    End If
    '
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("AMACLI/NEW", "Alta de Nueva Cuenta de Cliente")
    Call FRESHECHO("DEUDOR1")
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Call MENSERR(24, "Opción Inserción no Disponible.\Debe Agregarse al Final de la Lista.")
End Sub

Private Sub Command2_Click()
    If List2.ListCount > 0 Then
          Call CIERRARCH("!CARCONPE")
          Call CIERRARCH("!CARITPED")
          Call BLANARCH("!CARCONPE")
          Call BLANARCH("!CARITPED")
          '
          List2.Clear
          Text1.TEXT = ""
          Text2.TEXT = ""
          Text3.TEXT = ""
          Text4.TEXT = ""
          Text5.TEXT = ""
          HOII% = HOY(HOS$)
          Text9.TEXT = HOS$
          FEFAC% = HOII%
          '
          Label10.Caption = ""
          Label14.Caption = ""
          Label18.Caption = ""
          Label20.Caption = ""
          '
          Text1.SetFocus
        Else
          OPFACDIP.Show 1
    End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Text1.TEXT)
    Call CONECRUN("*NOTACLI/" + NCLI$, "Notas de Clientes")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RESCUDEU.Text1.TEXT = Text1.TEXT
    Hide
    PROLLAMA$ = "FORFACDIP"
    RESCUDEU.Show
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   '
   If List2.ListIndex < 0 Then Exit Sub
   If List2.ListIndex > List2.ListCount - 1 Then Exit Sub
   '
   ITID% = 1
   CACO = 1
   ILISTA& = List2.ListIndex
   List2.RemoveItem ILISTA&
   List2.ListIndex = (-1)
   If List2.ListCount > ILISTA& Then
       List2.ListIndex = ILISTA&
   End If
   '
   J& = 0
   For U& = 1 To ULTREG&("!CARITPED")
       Y$ = REGLEIDO$("!CARITPED", U&)
       If CVI(Left$(Y$, 2)) <> ITID% Then
          J& = J& + 1
          Call GRAREG("!CARITPED", Y$, J&)
       End If
   Next U&
   '
   ITOTCO# = 0
   For U& = 1 To List2.ListCount
      List2.ListIndex = U& - 1
      Z$ = List2.TEXT
      CAIT = Val(Mid$(Z$, 1, 4))
      CIIT% = CODINT%(Mid$(Z$, 7, 15))
      PRELI = Val(Mid$(Z$, 56, 9))
      PRECO = Val(Mid$(Z$, 65, 9))
      PREBA = PREBASE(CIIT%)
      IMPO# = CACO * CAIT * PRECO
      REIT& = Val(Mid$(Z$, 91, 6))
      '
      Y$ = REGBLAN$("CARITPED")
      Call REPLA(Y$, MKI$(ITID%), 1, 2)
      Call REPLA(Y$, MKI$(CIIT%), 3, 2)
      Call REPLA(Y$, MKS$(CAIT), 5, 4)
      Call REPLA(Y$, MKS$(PRELI), 9, 4)
      Call REPLA(Y$, MKS$(PREBA), 13, 4)
      Call REPLA(Y$, MKS$(PRECO), 17, 4)
      Call REPLA(Y$, MKD$(IMPO#), 21, 8)
      Call REPLA(Y$, MKS$(REIT&), 61, 4)
      J& = J& + 1
      Call GRAREG("!CARITPED", Y$, J&)
      ITOTCO# = ITOTCO# + IMPO#
   Next U&
   '
   Call SETULTREG("!CARITPED", J&)
   '
   PREUCO = ITOTCO# / CACO
   Label10.Caption = CSTRING$(MKD$(ITOTCO#), 3, 12, 2, 2)
   '
   ILISTA& = 0
   '
   X$ = REGLEIDO$("!CARCONPE", ILISTA& + 1)
   Call REPLA(X$, MKS$(PREUCO), 65, 4)
   Call REPLA(X$, MKD$(ITOTCO#), 69, 8)
   Call GRAREG("!CARCONPE", X$, ILISTA& + 1)
   '
   TOPEDIDO# = 0
   For U& = 1 To ULTREG&("!CARCONPE")
      X$ = REGLEIDO$("!CARCONPE", U&)
      TOPEDIDO# = TOPEDIDO# + CVD(Mid$(X$, 69, 8))
   Next U&
   '
   ALITOTA = 0
   If PIVAFAC% <> 5 Then ALITOTA = ALIVA
   FORFACDIP.Label14.Caption = CSTRING$(MKD$(TOPEDIDO#), 3, 11, 2, 2)
   If PIVAFAC% <> 5 Then FORFACDIP.Label18.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVA / 100), 3, 11, 2, 2)
   If RENOIN% = 1 Then
       FORFACDIP.Label21.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVARNI / 100), 3, 11, 2, 2)
       ALITOTA = ALIVA + ALIVARNI
   End If
   FORFACDIP.Label20.Caption = CSTRING$(MKD$(TOPEDIDO# * (1 + ALITOTA / 100)), 3, 11, 2, 2)
   '
End Sub

Private Sub Command6_Click()
    '
    VDEF$ = REGBLAN$("CARITPED")
10  OBX$ = OBJPLA$("CARITPED1" + "/ARCHECO(1)=" + ARCHIPRE$, VDEF$)
    If OBX$ = "" Then Exit Sub
    VDEF$ = OBX$
    CIIT% = CVI(Mid$(OBX$, 3, 2))
    CAIT = CVS(Mid$(OBX$, 5, 4))
    If CODEXT$(CIIT%) = "" Then
       Call MENSERR(24, "Codigo no Válido")
       GoTo 10
    End If
    '
    If CAIT < 0.1 Then
       Call MENSERR(24, "Falta Cantidad")
       GoTo 10
    End If
    '
    PREBA = PREBASE(CIIT%)
    PRELI = PRELISTA#(ARCHIPRE$, CIIT%)
    If PRELI < 0.005 Then
       Call MENSERR(24, "Artículo sin Precio")
    End If
    '
    PRECO = PRELI
    If PRECO < PREBA Then PRECO = PREBA
20  IMPO# = CACO * CAIT * PRECO
    '
    Call REPLA(OBX$, MKS$(PRELI), 9, 4)
    Call REPLA(OBX$, MKS$(PREBA), 13, 4)
    Call REPLA(OBX$, MKS$(PRECO), 17, 4)
    Call REPLA(OBX$, MKD$(IMPO#), 21, 8)
    '
    VDEF$ = OBX$ + "/ARCHECO(3)"
    CAANT = CAIT: PREANT = PRECO
25  Y$ = OBJPLA$("CARITFACDIP", VDEF$)
    If Y$ = "" Then Exit Sub
    If CVI(Mid$(Y$, 3, 2)) <> CIIT% Then
       Call MENSERR(24, "Cambio de Artículo no Válido\en esta Pantalla.")
       GoTo 10
    End If
    '
    '
    VDEF$ = Y$
    CIIT% = CVI(Mid$(Y$, 3, 2))
    CAIT = CVS(Mid$(Y$, 5, 4))
    PRELI = CVS(Mid$(Y$, 9, 4))
    PREBA = 0
    PRECO = CVS(Mid$(Y$, 17, 4))
    If PRECO < PREBA Then
       Call MENSERR(24, "Precio Cotizado Inferior a Precio de Base")
       GoTo 25
    End If
    ENVASE$ = TRIM$(Mid$(Y$, 33, 28))
    '
    ITID% = 1
    CACO = 1
    IMPO# = CACO * CAIT * PRECO
    '
    Call REPLA(Y$, MKI$(ITID%), 1, 2)
    Call REPLA(Y$, MKD$(IMPO#), 21, 8)
    Call REGAPP("!CARITPED", Y$)
    '
    Z$ = Space$(96)
    Call REPLA(Z$, CSTRING$(MKS$(CAIT), 3, 4, 0, 2), 1, 4)
    Call REPLA(Z$, CODEXT$(CIIT%), 7, 15)
    Call REPLA(Z$, TRIM$(ENVASE$ + " " + DESCRIT0$(CIIT%)), 23, 33)
    Call REPLA(Z$, CSTRING$(MKS$(PRELI), 3, 9, 2, 2), 56, 9)
    Call REPLA(Z$, CSTRING$(MKS$(PRECO), 3, 9, 2, 2), 65, 9)
    Call REPLA(Z$, CSTRING$(MKD$(IMPO#), 3, 10, 2, 2), 74, 10)
    FORFACDIP.List2.AddItem Z$
    List2.ListIndex = List2.ListCount - 1
    '
    ITOTCO# = 0
    For U& = 1 To ULTREG&("!CARITPED")
       Y$ = REGLEIDO$("!CARITPED", U&)
       If CVI(Left$(Y$, 2)) = ITID% Then
          CAIT = CVS(Mid$(Y$, 5, 4))
          PRECO = CVS(Mid$(Y$, 17, 4))
          IMPO# = CACO * CAIT * PRECO
          Call REPLA(Y$, MKD$(IMPO#), 21, 8)
          ITOTCO# = ITOTCO# + IMPO#
       End If
    Next U&
    PREUCO = ITOTCO# / CACO
    Label10.Caption = CSTRING$(MKD$(ITOTCO#), 3, 12, 2, 2)
    '
    ILISTA& = 0
    X$ = REGLEIDO$("!CARCONPE", ILISTA& + 1)
    Call REPLA(X$, MKS$(PREUCO), 65, 4)
    Call REPLA(X$, MKD$(ITOTCO#), 69, 8)
    Call GRAREG("!CARCONPE", X$, ILISTA& + 1)
    '
    TOPEDIDO# = 0
    For U& = 1 To ULTREG&("!CARCONPE")
       X$ = REGLEIDO$("!CARCONPE", U&)
       TOPEDIDO# = TOPEDIDO# + CVD(Mid$(X$, 69, 8))
    Next U&
    '
    ALITOTA = 0
    If PIVAFAC% <> 5 Then ALITOTA = ALIVA
    FORFACDIP.Label14.Caption = CSTRING$(MKD$(TOPEDIDO#), 3, 11, 2, 2)
    If PIVAFAC% <> 5 Then FORFACDIP.Label18.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVA / 100), 3, 11, 2, 2)
    If RENOIN% = 1 Then
       FORFACDIP.Label21.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVARNI / 100), 3, 11, 2, 2)
       ALITOTA = ALIVA + ALIVARNI
    End If
    FORFACDIP.Label20.Caption = CSTRING$(MKD$(TOPEDIDO# * (1 + ALITOTA / 100)), 3, 11, 2, 2)
    '
End Sub
'

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
        FEFAC% = FF%
    End If
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call BLOQARCH("FACTUDIP")
    '
    Show
    Refresh
    '
    OPFACDIP.Show 1
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    FEFAC% = FF%
    '
    Refresh
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub


Private Sub List2_Click()
    '
End Sub
'

Private Sub List2_DblClick()
    '
End Sub

Private Sub Text1_Change()
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label18.Caption = ""
    Label20.Caption = ""
    Label21.Caption = ""
    BOTNEXT.Enabled = True
End Sub

Private Sub Text1_DBLCLICK()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        Text1.TEXT = TRIM$(Str$(NCXI%))
        NC% = NCXI%
        OKEY% = 0
        Call FACITEMI1(OKEY%)
        'Call CARDATPEDMAI(OKEY%)
        If OKEY% <> 1 Then
           Call CLEARFORFACDIP
           Exit Sub
        End If
    End If
    '
End Sub

Private Sub Text1_GotFocus()
    Static CTXX%
    If CTXX% = 0 Then
        Refresh
        CTXX% = 1
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       OKEY% = 0
       Call FACITEMI1(OKEY%)
       'Call CARDATPEDMAI(OKEY%)
       If OKEY% <> 1 Then
           Call CLEARFORFACDIP
           Exit Sub
       End If
    End If
    '
End Sub

Private Sub Text11_CLICK()
   If MONUMERA$ = "" Then
      MONUMERA$ = "AUTO"
      If TRIM$(Control$("", "NUMFACTU")) = "MANUAL" Then
         MONUMERA$ = "MANUAL"
      End If
   End If
   If MONUMERA$ <> "MANUAL" Then
      Text1.SetFocus
   End If
End Sub

Private Sub Text11_LostFocus()
   NROPRO& = 0
   If Left$(Text11.TEXT, 3) = "FC." Then NROPRO& = Val(Mid$(Text11.TEXT, 4))
   Call NUFACTURA(NROPRO&)
   Text11.TEXT = "FC." + TRIM$(Str$(NROPRO&))
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text3_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text4_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text5_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text6_Click()
    Text1.SetFocus
End Sub

Private Sub Text6_DblClick()
    OPFACTUR.Show 1
End Sub

Private Sub Text7_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text9_DblClick()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FEFAC% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub Text9_LostFocus()
    VDEVV$ = TRIM$(Text9.TEXT)
    If FECHANUM(VDEVV$) > 0 Then
          FEFAC% = FECHANUM(VDEVV$)
        Else
          Call MENSERR(24, "Fecha no Válida")
          FEX = HOY(HO$)
          FEFAC% = FEX
          Text9.TEXT = HO$
          Text9.SetFocus
    End If
End Sub
