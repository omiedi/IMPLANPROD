VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PRESRE09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuestos por Reparaciones - Generación y Emisión"
   ClientHeight    =   8310
   ClientLeft      =   1800
   ClientTop       =   1485
   ClientWidth     =   11910
   Icon            =   "PRESRE09.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11910
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2325
      Left            =   0
      TabIndex        =   8
      Top             =   7170
      Width           =   10800
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
         Height          =   1020
         Left            =   225
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   0
         Width           =   7155
      End
      Begin VB.Label TOTAL 
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
         Height          =   300
         Index           =   2
         Left            =   9135
         TabIndex        =   28
         Top             =   740
         Width           =   1275
      End
      Begin VB.Label TOTAL 
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
         Height          =   300
         Index           =   1
         Left            =   9135
         TabIndex        =   27
         Top             =   370
         Width           =   1275
      End
      Begin VB.Label TOTAL 
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
         Height          =   300
         Index           =   0
         Left            =   9135
         TabIndex        =   26
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Presupuesto:"
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
         Index           =   3
         Left            =   7140
         TabIndex        =   11
         Top             =   815
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A:"
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
         Left            =   7140
         TabIndex        =   10
         Top             =   445
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Neto:"
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
         Left            =   7140
         TabIndex        =   9
         Top             =   75
         Width           =   1905
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox Picture14 
         BackColor       =   &H00400000&
         Height          =   8325
         Left            =   10935
         ScaleHeight     =   8265
         ScaleWidth      =   1155
         TabIndex        =   44
         Top             =   15
         Width           =   1215
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   120
            Top             =   4680
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton Command14 
            Caption         =   "O.K."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1030
            Left            =   60
            Picture         =   "PRESRE09.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Aprueba, Graba e Imprime Presupuesto"
            Top             =   7125
            Width           =   855
         End
         Begin VB.CommandButton Command3 
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
            Height          =   660
            Left            =   105
            Picture         =   "PRESRE09.frx":0974
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   1095
            Width           =   750
         End
         Begin VB.CommandButton Command13 
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
            Height          =   660
            Left            =   105
            Picture         =   "PRESRE09.frx":0C7E
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   3915
            Width           =   750
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Clientes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   105
            Picture         =   "PRESRE09.frx":10C0
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Acceso a Base de Datos de Clientes"
            Top             =   2505
            Width           =   750
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Articulos"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   105
            Picture         =   "PRESRE09.frx":1502
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   1800
            Width           =   750
         End
         Begin VB.CommandButton Command2 
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
            Height          =   660
            Left            =   105
            Picture         =   "PRESRE09.frx":180C
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   390
            Width           =   750
         End
         Begin VB.CommandButton Command15 
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
            Height          =   660
            Left            =   105
            Picture         =   "PRESRE09.frx":1956
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Ver y Re-Imprimir Presupuestos"
            Top             =   3210
            Width           =   750
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   150
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   45
            Top             =   6240
            Width           =   645
            Begin VB.Frame Frame7 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   46
               Top             =   105
               Width           =   12000
            End
         End
         Begin VB.Image Image2 
            Height          =   4095
            Left            =   120
            Top             =   5400
            Visible         =   0   'False
            Width           =   5460
         End
         Begin VB.Image Image3 
            Height          =   390
            Left            =   -285
            Picture         =   "PRESRE09.frx":1C60
            Top             =   6555
            Width           =   1515
         End
      End
      Begin VB.TextBox Text10 
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
         Left            =   8625
         TabIndex        =   36
         Text            =   " "
         Top             =   900
         Width           =   2130
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Repuestos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   4230
         Left            =   210
         TabIndex        =   21
         ToolTipText     =   "Doble-Click Aquí para Editar Repuestos"
         Top             =   2310
         Width           =   10515
         Begin VB.ListBox List1 
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
            Left            =   30
            TabIndex        =   23
            Top             =   555
            Width           =   10515
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo          Descripción                                Cant.  Pr.Unit    Importe"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   225
            Left            =   70
            TabIndex        =   22
            ToolTipText     =   "Doble-Click Aquí para Editar Repuestos"
            Top             =   315
            Width           =   10100
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Serie Número"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   915
         Left            =   7980
         TabIndex        =   20
         Top             =   1365
         Width           =   2745
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
            Height          =   300
            Left            =   210
            TabIndex        =   0
            Text            =   " "
            Top             =   495
            Width           =   2325
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   915
         Left            =   210
         TabIndex        =   12
         Top             =   1365
         Width           =   7680
         Begin VB.Label Label8 
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
            Height          =   510
            Left            =   210
            TabIndex        =   13
            Top             =   285
            Width           =   7305
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
            Height          =   300
            Left            =   210
            TabIndex        =   34
            Top             =   495
            Visible         =   0   'False
            Width           =   2325
         End
      End
      Begin VB.TextBox Text9 
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
         Height          =   300
         Left            =   9615
         TabIndex        =   2
         Text            =   " "
         Top             =   105
         Width           =   960
      End
      Begin VB.CommandButton Command9 
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
         Left            =   10560
         TabIndex        =   3
         Top             =   120
         Width           =   180
      End
      Begin VB.Label Label20 
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
         Height          =   300
         Left            =   5565
         TabIndex        =   43
         Top             =   6650
         Width           =   960
      End
      Begin VB.Label Label19 
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
         Height          =   300
         Left            =   3360
         TabIndex        =   42
         Top             =   6645
         Width           =   960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   5340
         TabIndex        =   41
         Top             =   150
         Width           =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   6825
         TabIndex        =   40
         Top             =   915
         Width           =   1800
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   5325
         TabIndex        =   39
         Top             =   525
         Width           =   1695
      End
      Begin VB.Label Label16 
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
         Height          =   300
         Left            =   5850
         TabIndex        =   38
         Top             =   105
         Width           =   2220
      End
      Begin VB.Label Label17 
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
         Height          =   300
         Left            =   5850
         TabIndex        =   37
         Top             =   495
         Width           =   2220
      End
      Begin VB.Line Line1 
         X1              =   9030
         X2              =   9030
         Y1              =   6510
         Y2              =   7060
      End
      Begin VB.Label REPRESUP 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   225
         Left            =   6510
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "M.de Obra:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   -735
         TabIndex        =   33
         Top             =   6690
         Width           =   2115
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Estimado de Horas:"
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
         Left            =   1365
         TabIndex        =   32
         Top             =   6720
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Hora:"
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
         Index           =   4
         Left            =   3570
         TabIndex        =   31
         Top             =   6720
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Importe:"
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
         Left            =   5670
         TabIndex        =   30
         Top             =   6720
         Width           =   1905
      End
      Begin VB.Label TOTAMOBRA 
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
         Height          =   300
         Left            =   7680
         TabIndex        =   29
         Top             =   6645
         Width           =   1275
      End
      Begin VB.Label SUMAREP 
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
         Height          =   300
         Left            =   9150
         TabIndex        =   25
         Top             =   6645
         Width           =   1275
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
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
         Height          =   300
         Left            =   4935
         TabIndex        =   24
         Top             =   1785
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   9630
         TabIndex        =   19
         Top             =   495
         Width           =   1110
      End
      Begin VB.Label Label15 
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
         Height          =   300
         Left            =   2310
         TabIndex        =   18
         Top             =   890
         Width           =   2850
      End
      Begin VB.Label Label14 
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
         Height          =   300
         Left            =   1470
         TabIndex        =   17
         Top             =   890
         Width           =   750
      End
      Begin VB.Label Label13 
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
         Height          =   300
         Left            =   1470
         TabIndex        =   16
         Top             =   490
         Width           =   3690
      End
      Begin VB.Label Label12 
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
         Height          =   300
         Left            =   1455
         TabIndex        =   15
         Top             =   105
         Width           =   3690
      End
      Begin VB.Label Label11 
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
         Height          =   300
         Left            =   420
         TabIndex        =   14
         Top             =   885
         Width           =   750
      End
      Begin VB.Line Line7 
         X1              =   210
         X2              =   10710
         Y1              =   7060
         Y2              =   7060
      End
      Begin VB.Line Line8 
         X1              =   10725
         X2              =   10725
         Y1              =   2445
         Y2              =   7060
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   420
         Picture         =   "PRESRE09.frx":3B82
         ToolTipText     =   "Click Aquí para cambiar Modalidad de Facturación"
         Top             =   105
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Presup.Nro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   7770
         TabIndex        =   7
         Top             =   525
         Width           =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   8820
         TabIndex        =   6
         Top             =   150
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   210
         TabIndex        =   5
         Top             =   630
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
End
Attribute VB_Name = "PRESRE09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG07"
   If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "AMACLI01"
   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
   Command11.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   CONFRE09.Show 1
   '
   CUHOR = Val(Control$("REPARA", "VALHORA"))
   PRESRE09.Label20 = CSTRING$(MKS$(CUHOR), 4, 8, 2, 2)
   COEFI = Val(Control$("REPARA", "COEFPREA"))
   If COEFI < 1 Then COEFI = 1
   '
   If ULTREG&("!MATEROF") > 0 Then
      Call PREREPUES
      GoTo 99
   End If
   '
   PRESRE09.List1.Clear
   SUREPUES = 0
   CODD$ = Label10
   Screen.MousePointer = 11
   RFF$ = RECFILT$("REPUESRE", 1, CODD$)
   For IK& = 1 To Len(RFF$) Step 4
     RERE& = CVS(Mid$(RFF$, IK&, 4))
     If RERE& > 0 Then
       If RERE& <= ULTREG&("REPUESRE") Then
         XX$ = REGLEIDO$("REPUESRE", RERE&)
         CIY% = CVI(Mid$(XX$, 17, 2))
         CODXX$ = CODEXT$(CIY%)
         DEXX$ = DESCRIT$(CIY%)
         COSUN = COEFI * COSUNI(CIY%)
         COSUX$ = CSTRING$(MKS$(COSUN), 4, 10, 2, 2)
         Uso = CVS(Mid$(XX$, 19, 4))
         USOX$ = CSTRING$(MKS$(Uso), 4, 10, 1, 2)
         COSTO = COSUN * Uso
         COSTX$ = CSTRING$(MKS$(COSTO), 4, 12, 2, 2)
         SUREPUES = SUREPUES + COSTO
         '
         XXX$ = Space$(64)
         Call REPLA(XXX$, CODXX$, 1, 16)
         Call REPLA(XXX$, DEXX$, 17, 38)
         Call REPLA(XXX$, USOX$, 54, 10)
         Call REPLA(XXX$, COSUX$, 64, 10)
         Call REPLA(XXX$, COSTX$, 74, 12)
         List1.AddItem XXX$
       End If
     End If
   Next IK&
   SUMAREP = CSTRING$(MKS$(SUREPUES), 4, 12, 2, 2)
   Screen.MousePointer = 1
   '
99 Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call GRAPRESU(OKX%)
   If OKX% > 0 Then
     CANCELPRINT% = 0
     Call PRIPRESU
     Call BLANFORMU
     Label18 = ""
     Command14.Enabled = True
     Label11 = ""
     Label18 = "": Label10 = ""
     Hide
   End If
   Command14.Enabled = True
End Sub

Sub Command15_Click()
    '
    Command15.Enabled = False
    '
'3   Call SELECHO("PRESREPA/Indice Presupuestos de Reparación")
'    NPX& = Val(VALACT1$("PRESREPA"))
    '
3   Call SELECHO("PREPAEMI/Indice Presupuestos de Reparación")
    NPX& = Val(VALACT1$("PREPAEMI"))

    If NPX& < 1 Then GoTo 99
'    '
'    NUDOCU$ = TRIM$(Str$(NPX&))
'    For UI& = ULTREG&("PRESREPA") To 1 Step -1
'      XX$ = REGLEIDO$("PRESREPA", UI&)
'      If CVS(Left$(XX$, 4)) = NPX& Then GoTo 4
'    Next UI&
'    Call COMUNI("Error de Indices - Consulte.")
'    GoTo 99
    '
4   NUDOCU$ = TRIM$(Str$(NPX&))
    Screen.MousePointer = 11
    TIDOCUM$ = "Presupuesto de Reparación " + NUDOCU$
    TIDOCU2$ = "Operaciones de Reparación PR-" + NUDOCU$
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
      End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
    On Error Resume Next
    GoTo 99
    '
5 Call MUESTRADOC(DOCUNU&)
'5   For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    Exit Sub
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 10
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'10  For UI& = ULTREG&("PDOCLST") To 1 Step -1
'      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'      If InStr(XX$, TIDOCU2$) > 4 Then
'        DOCUNU& = CVS(Left$(XX$, 4))
'        GoTo 15
'      End If
'    Next UI&
'    '
'15  REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'16  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 17
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 17
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 16
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 16
'    GoTo 18
'    '
'17  Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    Exit Sub
'    '
'18  While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 19
'      L& = L& - 1
'    Wend
'    '
'19  For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7170
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7170
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'50  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'       COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'       TPSP$ = Mid$(TPSP$, PPXX% + 1)
'       '
'       If Len(COPRIAT$) > 0 Then
'         COPRI% = Asc(Left$(COPRIAT$, 1))
'         ATRIP$ = Mid$(COPRIAT$, 2)
'         Call SHOWCOMMAND(COPRI%, ATRIP$)
'       End If
'       GoTo 50
'       '
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Visible = False
'      If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
''    EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command15.Enabled = True
    '
End Sub


Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Label11 = ""
    Label18 = "": Label10 = ""
    Hide
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PRESUP01")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    If DERACCE("RESCUDEU", "Resumen de Cuenta de Deudores") >= 2 Then
      RESCUDEU.Text1.TEXT = Label11
      PROLLAMA$ = "PRESURE"
      RESCUDEU.Show 1
    End If
    Command4.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "AMASTO07"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "AMASTO01"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   Text9.TEXT = VALACT1$("SELFECHA")
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    '
    CUHOR = Val(Control$("REPARA", "VALHORA"))
    COEFI = Val(Control$("REPARA", "COEFPREA"))
    Label20 = CSTRING$(MKS$(CUHOR * COEFI), 4, 8, 2, 2)
    '

End Sub
'

Private Sub Form_Resize()
   If WindowState = 1 Then
      Call CIERRARCH("*.*")
      Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Label11 = ""
   Label18 = "": Label10 = ""
   Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub Frame4_DblClick()
   Call EDITREPU
End Sub

Private Sub Label11_Change()
   '
   NC% = Val(Label11)
   Label12 = ""
   Label13 = ""
   Label14 = ""
   Label15 = ""
   Label16 = ""
   Label17 = ""
   '
   If NC% > 0 Then
     If rasocli$(NC%) <> "" Then
       Label12 = rasocli$(NC%): Label12.Refresh
       Label13 = DOMICLI$(NC%)
       Label14 = CPOSCLI$(NC%)
       Label15 = LOCACLI$(NC%)
       Label16 = TEL1CLI$(NC%)
       Label17 = FAXCLI$(NC%)

       DoEvents
     End If
   End If
   '
End Sub

Private Sub Label18_Change()
   '
   NPX& = Val(Label18)
   Call BLANFORMU

   If NPX& > 0 Then
      CONDI$ = "NumRepa = " & NPX& & ""
      Label11 = VALOBADA("REPARA", "Nume_Cli", CONDI$, "NOMESS")
      Text1 = VALOBADA("REPARA", "NumeroSerie", CONDI$, "NOMESS")
      Text10 = VALOBADA("REPARA", "Referen2", CONDI$, "NOMESS")
      NCLIE% = VALOBADA("REPARA", "Nume_Cli", CONDI$, "NOMESS")
      Text28 = VALOBADA("REPARA", "Observa", CONDI$, "NOMESS")
      REPRESUP = VALOBADA("REPARA", "NumRepa", CONDI$, "NOMESS")
      code% = VALOBADA("REPARA", "Cod_Inte", CONDI$, "NOMESS")
      Label10 = CODEXT$(code%)
      TEPRUEBA = DESCRIT$(code%)
   End If
   '
   Call CAREPUES
   Call PREREPUES
   '
   On Error Resume Next
   Text10.SetFocus
   On Error GoTo 0
   Openforms = DoEvents

'   For UI& = ULTREG&("PRESREPA") To 1 Step -1
'     XX$ = REGLEIDO$("PRESREPA", UI&)
'     If CVS(Left$(XX$, 4)) = NPX& Then
'       REPRESUP = TRIM$(Str$(UI&))
'       NCLIE% = CVI(Mid$(XX$, 67, 2))
'       Label11 = TRIM$(Str$(NCLIE%))
'       CII% = CVI(Mid$(XX$, 69, 2))
'       Label10 = CODEXT$(CII%)
'       TEPRUEBA = DESCRIT$(CII%)
'       Text1 = Mid$(XX$, 71, 15)
'       Text10 = Mid$(XX$, 86, 9)
'       TALON = Mid$(XX$, 95, 6)
'       Text28 = TRIM$(Mid$(XX$, 129, 368))
'       Label19 = CSTRING$(Mid$(XX$, 101, 4), 4, 8, 2, 2)
'       If CVS(Mid$(XX$, 105, 4)) > 0.5 Then
'        Label20 = CSTRING$(Mid$(XX$, 105, 4), 4, 8, 2, 2)
'       End If
'       '
'       Call CAREPUES
'       Call PREREPUES
'       '
'       On Error Resume Next
'       Text10.SetFocus
'       On Error GoTo 0
'       Openforms = DoEvents
'       Exit Sub
'     End If
'   Next UI&
   '
End Sub

Private Sub Label9_DblClick()
   Call EDITREPU
End Sub

Private Sub List1_DblClick()
   '
   Call EDITREPU
   '
End Sub

Private Sub SUMAREP_Change()
   TOTAMO = Val(SUMAREP) + Val(TOTAMOBRA)
   TOTAL(0) = CSTRING$(MKS$(TOTAMO), 3, 11, 2, 2)
End Sub

Private Sub TEPRUEBA_Change()
   DESCRI$ = TEPRUEBA.Caption
   If TEPRUEBA.Width < Label8.Width Then
     DESCRI$ = Chr$(10) + Chr$(13) + DESCRI$
   End If
   Label8.Caption = DESCRI$
End Sub


Private Sub LABEL20_Change()
   TOTAMO = Val(Label19) * Val(Label20)
   TOTAMOBRA = CSTRING$(MKS$(TOTAMO), 3, 11, 2, 2)
End Sub

Function VALICLIE%()
   VCC% = 0
   If Val(Label11) > 0 Then
     If Val(Label11) <= 32767 Then
       NCLIE% = Val(Label11)
       If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function
'
Sub CAREPUES()
   '
   Call SETULTREG("!MATEROF", 0)
   NPX& = Val(Label18)
'   NPI% = NPX& Mod 10000
'   RFF$ = RECFILT$("REPUREPA", 2, MKI$(NPI%))
'   '
'   URE& = ULTREG&("REPUREPA")
'30 For KK& = 1 To Len(RFF$) Step 4
'     REG& = CVS(Mid$(RFF$, KK&, 4))
'     If REG& > 0 Then
'       If REG& <= URE& Then
'         XX$ = REGLEIDO$("REPUREPA", REG&)
'         If CVS(Left$(XX$, 4)) = NPX& Then
'           CII% = CVI(Mid$(XX$, 7, 2))
'           CANTI = CVS(Mid$(XX$, 9, 4))
'           '
'           ZZ$ = REGBLAN$("MATEROF")
'           Call REPLA(ZZ$, MKI$(CII%), 1, 2)
'           Call REPLA(ZZ$, UNIMED$(CII%), 3, 2)
'           Call REPLA(ZZ$, MKS$(CANTI), 5, 4)
'           Call REPLA(ZZ$, MKD$(CDbl(CANTI)), 9, 8)
'           Call REPLA(ZZ$, MKS$(0), 17, 4)
'           Call REGAPP("!MATEROF", ZZ$)
'           '
'         End If
'       End If
'     End If
'   Next KK&
'   Call CIERRARCH("!MATEROF")
   '
   'Lectura desde la Tabla SQL
30 CONDI$ = "NumRepa = " & NPX& & " AND ORIGEN=1" ' ORIGEN=0 REPUESTOS DE TECNICA Y ORIGEN=1 DE COMERCIAL
   ORI$ = 0: If CantRegistros("REPUREPA", CONDI$) > 0 Then ORI$ = "1"
   SQLX$ = "SELECT * FROM REPUREPA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NumRepa = " & NPX& & " AND ORIGEN=" & ORI$ & " ORDER BY NuOrd_Item ASC"
   If NPX& > 0 Then
        Dim REPAR As ADODB.Recordset
        Set REPAR = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With REPAR
                Do While Not .EOF
                     CII% = !cod_inte
                     CANTI = !CANT_SOL
                     '
                     zz$ = REGBLAN$("MATEROF")
                     Call REPLA(zz$, MKI$(CII%), 1, 2)
                     Call REPLA(zz$, UNIMED$(CII%), 3, 2)
                     Call REPLA(zz$, MKS$(CANTI), 5, 4)
                     Call REPLA(zz$, MKD$(CDbl(CANTI)), 9, 8)
                     Call REPLA(zz$, MKS$(0), 17, 4)
                     Call REGAPP("!MATEROF", zz$)
                     '
                     .MoveNext
                 Loop
        End With
        REPAR.Close
        Set REPAR = Nothing
   End If
   '
   Call CIERRARCH("!MATEROF")
   '
End Sub
'
Sub PREREPUES()
   '
   List1.Clear
   SUMA# = 0
   SUMAREP = ""
   COEFI = Val(Control$("REPARA", "COEFPREA"))
   For U& = 1 To ULTREG&("!MATEROF")
     zz$ = REGLEIDO$("!MATEROF", U&)
     TTXX$ = Space$(128)
     CII% = CVI(Left$(zz$, 2))
     Call REPLA(TTXX$, CODEXT$(CII%), 1, 16)
     Call REPLA(TTXX$, DESCRIT$(CII%), 17, 38)
     NEBRUX$ = Mid$(zz$, 5, 4)
     Call REPLA(TTXX$, CSTRING$(NEBRUX$, 4, 10, 1, 2), 54, 10)
     COSUX$ = MKS$(COEFI * COSUNI(CII%))
     Call REPLA(TTXX$, CSTRING$(COSUX$, 4, 10, 2, 2), 64, 10)
     ITOX$ = MKS$(CVS(NEBRUX$) * CVS(COSUX$))
     Call REPLA(TTXX$, CSTRING$(ITOX$, 4, 12, 2, 2), 74, 12)
     SUMA# = SUMA# + Val(Mid$(TTXX$, 74, 12))
     List1.AddItem TTXX$
   Next U&
   SUMAREP = CSTRING$(MKD$(SUMA#), 4, 11, 2, 2)
   '
End Sub
'
Sub BLANFORMU()
   '
   Label8 = ""
   Label10 = ""
   Label12 = ""
   Label13 = ""
   Label14 = ""
   Label15 = ""
   Label16 = ""
   Label17 = ""
   Text1 = ""
   Text10 = ""
   TALON = ""
   TEPRUEBA = ""
   Text28 = ""
   Label19 = ""
   Label20 = ""
   List1.Clear
   '
   CUHOR = Val(Control$("REPARA", "VALHORA"))
   COEFI = Val(Control$("REPARA", "COEFPREA"))
   Label20 = CSTRING$(MKS$(CUHOR * COEFI), 4, 8, 2, 2)
   '
   For KK% = 0 To 2
      TOTAL(KK%) = ""
   Next KK%
   '
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   '
End Sub
'
Function NUPRESUP&()
   '
   NPP& = 1
   '
   FIN& = ULTREG&("PRESUPU")
   If FIN& > 0 Then
      For U& = FIN& To FIN& - 128 Step -1
         If U& < 1 Then GoTo 99
         XX$ = REGLEIDO$("PRESUPU", U&)
         NPPX& = CVS(Mid$(XX$, 1, 4))
         If NPPX& >= NPP& Then
            NPP& = 1 + NPPX&
         End If
      Next U&
   End If
   '
99 NUPRESUP& = NPP&
   '
End Function
'

Private Sub LABEL20_LostFocus()
   Label20 = CSTRING$(MKS$(Val(Label20)), 4, 8, 2, 2)
End Sub

Private Sub LABEL19_Change()
   TOTAMO = Val(Label19) * Val(Label20)
   TOTAMOBRA = CSTRING$(MKS$(TOTAMO), 3, 11, 2, 2)
End Sub

Private Sub LABEL19_LostFocus()
   Label19 = CSTRING$(MKS$(Val(Label19)), 4, 8, 2, 2)
End Sub

Private Sub Text9_GotFocus()
   Command9.SetFocus
End Sub

Private Sub Total_Change(Index As Integer)
   If Index = 0 Then
     NCLIE% = Val(Label11)
     If NCLIE% > 0 Then
       PIVA% = PIVACLI%(NCLIE%)
       FACIVA = ALIVA / 100
       If PIVA% = 2 Then
            FACIVA = 1.5 * ALICU / 100
          ElseIf PIVA% = 5 Then
            FACIVA = 0
            ALICU = 0
       End If
       TOTAL(1) = CSTRING$(MKS$(FACIVA * Val(TOTAL(0))), 4, 11, 2, 2)
       TOTAL(2) = CSTRING$(MKS$(Val(TOTAL(0)) + Val(TOTAL(1))), 4, 11, 2, 2)
     End If
   End If
End Sub
'
Sub PRIPRESU()
    '
    OPX% = COMALTER%("Opciones de Impresion:\\Emisión de Presupuestos Común\Generar Presupuesto en Word")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    If OPX% = 2 Then
        GENERAPRESUWORD
        Exit Sub
    End If
    
    FORIPRE$ = TRIM$(Control$("FOREPAR", "DOCUREP1"))
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Imposible Imprimir Presupuesto de Reparación.\Falta Definir Formulario Electrónico.")
       Exit Sub
    End If
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Presupuesto de Reparación.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    TTXYZ$ = Space$(92)
    NROREP& = Val(Label18)
    HOII% = HOY(HOS$)
    Call REPLA(TTXYZ$, HOS$, 1, 8)
    Call REPLA(TTXYZ$, USERTER$, 69, 24)
    Call REPLA(TTXYZ$, "Emisión Presupuesto de Reparación:", 11, 56)
    Call AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
'    Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
    TTXYZ$ = Space$(92)
    Call REPLA(TTXYZ$, Space$(56), 11, 56)
    Call AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
'    Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
    '
10  NC% = Val(Label11)
    TIPODOC$ = "Presupuesto de Reparación"
    DESTIDOC% = NC%
    NUMEDOC$ = TRIM$(Label18)
    FECHDOC$ = TRIM$(Text9.TEXT)
    '
    CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Label10)
    CODPARAM$(2) = "DES-CONJ": DOCPARAM$(2) = TRIM$(ECOARCH$("INDIREP", Label10))
    CODPARAM$(3) = "REF-MAT1": DOCPARAM$(3) = TRIM$(Text1)
    CODPARAM$(4) = "MEDIDA-1": DOCPARAM$(4) = TRIM$(Text10)
    CODPARAM$(5) = "REF-MAT2": DOCPARAM$(5) = TRIM$(Label19)
    CODPARAM$(6) = "REF-MAT3": DOCPARAM$(6) = TRIM$(Label20)
    CODPARAM$(7) = "PRE-NETO": DOCPARAM$(7) = TRIM$(SUMAREP)
    CODPARAM$(8) = "BRU-ITEM": DOCPARAM$(8) = TRIM$(TOTAMOBRA)
    CODPARAM$(9) = "IMP-NETO": DOCPARAM$(9) = TRIM$(TOTAL(0))
    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = TRIM$(TOTAL(1))
    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = TRIM$(TOTAL(2))
    PORI$ = ""
    On Error Resume Next
    PORI$ = TRIM$(Str$(100 * TOTAL(1) / TOTAL(0)))
    On Error GoTo 0
    CODPARAM$(12) = "POR-IVA": DOCPARAM$(12) = TRIM$(PORI$)
    CODPARAM$(13) = "MEDIDA-3": DOCPARAM$(13) = TRIM$(Text10)
    CODPARAM$(14) = "DES-MPRI": DOCPARAM$(14) = Label8
    CAPARDOC% = 14
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "PRE-LIST"
    CODTABU1$(5) = "NET-ITEM"
    CACOLTAB1% = 5
    '
    CAITAB1% = 0
    For U& = 1 To List1.ListCount
      zz$ = List1.List(U& - 1)
      CAITAB1% = CAITAB1% + 1
      TETABU1$(1, CAITAB1%) = Left$(zz$, 16)
      CII% = CODINT%(Left$(zz$, 16))
      TETABU1$(2, CAITAB1%) = DESCRIT$(CII%)
      TETABU1$(3, CAITAB1%) = TRIM$(Mid$(zz$, 56, 8))
      TETABU1$(4, CAITAB1%) = TRIM$(Mid$(zz$, 64, 10))
      TETABU1$(5, CAITAB1%) = TRIM$(Mid$(zz$, 74, 11))
      '
      TTXYZ$ = Space$(92)
      Call REPLA(TTXYZ$, Left$(zz$, 16), 22, 16)
      Call REPLA(TTXYZ$, Mid$(zz$, 56, 30), 37, 30)
      If U& = 1 Then Call REPLA(TTXYZ$, "Repuestos:", 11, 10)
      Call AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
'      Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
    Next U&
    TTXYZ$ = Space$(10) + String$(56, "-") + Space(26)
    Call AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
'    Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
    TTXYZ$ = Space$(92)
    Call REPLA(TTXYZ$, Space$(56), 11, 56)
    Call AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
'    Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
    '
    Call SETULTREG("!OBSERVOF", 0)
    Call FRATEXTO(Text28.TEXT, 56)
    For YX& = 1 To CARETEX%
       Call REGAPP("!OBSERVOF", RETEX$(YX&))
    Next YX&
    '
    Call PRINTDOC("@" + FORIPRE$)
    '
    FORIPRE$ = TRIM$(Control$("FOREPAR", "DOCUREP4"))
    NC% = Val(Label11)
    TIPODOC$ = "Operaciones de Reparación"
    DESTIDOC% = NC%
    NUMEDOC$ = "PR-" + TRIM$(Label18) '"7" n.presupuesto
    '
    CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Label10) 'codigo externo
    CODPARAM$(2) = "REF-MAT1": DOCPARAM$(2) = TRIM$(Text1) 'n.serie
    CODPARAM$(3) = "MEDIDA-3": DOCPARAM$(3) = TRIM$(Text10) 'referencia
    CODPARAM$(4) = "DES-MPRI": DOCPARAM$(4) = Label8 ' ver vacio
    CODPARAM$(5) = "TOT-TIEM": DOCPARAM$(5) = HORMINSE$(3600 * Val(Label19)) 'estimado de horas es el label19
    CAPARDOC% = 5
    '
    CODTABU1$(1) = "NOP-OPER"
    CODTABU1$(2) = "DES-OPER"
    CODTABU1$(3) = "TIE-PROD"
    CODTABU1$(4) = "NUOPERS"
    CODTABU1$(5) = "TOT-PROD"
    CACOLTAB1% = 5
    '
    CAITAB1% = 0
    CODD$ = Label10
    RFFX$ = RECFILT$("SECOPREP", 1, CODD$)
    For I1& = 1 To Len(RFFX$) Step 4
      i& = CVS(Mid$(RFFX$, I1&, 4))
      X$ = REGLEIDO$("SECOPREP", i&)
      NOPE% = CVI(Mid$(X$, 17, 2))
      DOPE$ = Mid$(X$, 19, 48)
      TIEX = CVS(Mid$(X$, 73, 4))
      CAOPS = CVS(Mid$(X$, 97, 4)): If CAOPS < 1 Then CAOPS = 1
      '
      CAITAB1% = CAITAB1% + 1
      TETABU1$(1, CAITAB1%) = Str$(NOPE%)
      TETABU1$(2, CAITAB1%) = DOPE$
      TETABU1$(3, CAITAB1%) = HORMINSE$(3600 * TIEX)
      TETABU1$(4, CAITAB1%) = TRIM$(Str$(CAOPS))
      TETABU1$(5, CAITAB1%) = HORMINSE$(3600 * TIEX * CAOPS)
      '
      TTXYZ$ = Space$(92)
      Call REPLA(TTXYZ$, DOPE$, 22, 22)
      Call REPLA(TTXYZ$, TETABU1$(3, CAITAB1%), 43, 10)
      Call REPLA(TTXYZ$, TETABU1$(4, CAITAB1%), 56, 4)
      Call REPLA(TTXYZ$, TETABU1$(5, CAITAB1%), 58, 10)
      If I1& = 1 Then Call REPLA(TTXYZ$, "   M.Obra:", 11, 10)
      '   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
      Call AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    Next I1&
    TTXYZ$ = Space$(10) + String$(56, "-") + Space(26)
    Call AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
'    Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
    TTXYZ$ = Space$(92)
    Call REPLA(TTXYZ$, Space$(56), 11, 56)
    Call REPLA(TTXYZ$, "Total Estimado Horas ...........................", 11, 56)
    Call REPLA(TTXYZ$, HORMINSE$(3600 * Val(Label19)), 58, 10)
    Call AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
'    Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
    TTXYZ$ = Space$(92)
    Call REPLA(TTXYZ$, Space$(56), 11, 56)
    Call AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
'    Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
    '
    Call FRATEXTO(Text28, 46)
    For KKI% = 1 To CARETEX%
      TTXX$ = Space$(92)
      Call REPLA(TTXX$, RETEX$(KKI%), 20, 46)
      If KKI% = 1 Then Call REPLA(TTXX$, " Observ: ", 11, 9)
      Call AGREGA_ANOTAREPARA(NROREP&, TTXX$)
'      Call REGAPP("ANOTAREP", Left$(MKS$(NROREP&) + Space$(10) + TTXX$, 96))
    Next KKI%
    '
    Call AGREGA_ANOTAREPARA(NROREP&, String$(88, "="))
'    Call REGAPP("ANOTAREP", MKS$(NROREP&) + String$(88, "="))
'    Call CIERRARCH("ANOTAREP")
    Call CIERRARCH("*.*")
    '
    Call PRINTDOC("@" + FORIPRE$)
    '
    '
    'FORIPRE$ = TRIM$(CONTROL$("FOREPAR", "DOCUREP0"))
    'NC% = Val(Label11)
    'TIPODOC$ = "Operaciones de Reparación"
    'DESTIDOC% = NC%
    'NUMEDOC$ = "PR-" + TRIM$(Label18)
    '
    'CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Label10)
    'CODPARAM$(2) = "REF-MAT1": DOCPARAM$(2) = TRIM$(Text1)
    'CODPARAM$(3) = "MEDIDA-3": DOCPARAM$(3) = TRIM$(Text10)
    'CODPARAM$(4) = "DES-MPRI": DOCPARAM$(4) = Label8
    'CAPARDOC% = 4
    '
    'CODTABU1$(1) = "NOP-OPER"
    'CODTABU1$(2) = "OBS-OPER"
    'CODTABU1$(3) = "TIE-PROD"
    'CACOLTAB1% = 3
    ''
    'CAITAB1% = 0
    'URN& = ULTREG&("OBSOPREP")
    'CODD$ = Label10
    'RFFX$ = RECFILT$("SECOPREP", 1, CODD$)
    'For I1& = 1 To Len(RFFX$) Step 4
    '  I& = CVS(Mid$(RFFX$, I1&, 4))
    '  X$ = REGLEIDO$("SECOPREP", I&)
    '  NOPE% = CVI(Mid$(X$, 17, 2))
    '  DOPE$ = Mid$(X$, 19, 48)
    '  TIEX = CVS(Mid$(X$, 73, 4))
    '  '
    '  RNC$ = RECFILT$("OBSOPREP", 1, CODD$)
    '  '
    '  For U& = 1 To Len(RNC$) Step 4
    '    RENOTA& = CVS(Mid$(RNC$, U&, 4))
    '    If RENOTA& > 0 Then
    '      If RENOTA& <= URN& Then
    '        TTZZ$ = REGLEIDO$("OBSOPREP", RENOTA&)
    '        NOPY% = CVI(Mid$(TTZZ$, 17, 2))
    '        If NOPY% = NOPE% Then
    '          DOPE$ = DOPE$ + Mid$(TTZZ$, 19)
    '        End If
    '      End If
    '    End If
    '  Next U&
    '  '
    '  CAITAB1% = CAITAB1% + 1
    '  TETABU1$(1, CAITAB1%) = Str$(NOPE%)
    '  TETABU1$(2, CAITAB1%) = DOPE$
    '  TETABU1$(3, CAITAB1%) = HORMINSE$(3600 * TIEX)
    '  '
    'Next I1&
    '
    'Call PRINTDOC("@" + FORIPRE$)
    '
    '
    FORIPRE$ = TRIM$(Control$("FOREPAR", "DOCUREP3"))
    If FORIPRE$ = "" Then
       Exit Sub
    End If
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Presupuesto de Reparación.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    PLAENREP$ = TRIM$(Control$("", "PLAENREP"))
    If PLAENREP$ = "" Then PLAENREP$ = "5 Días hábiles de recibida su órden de compra."
    '
    CONPAREP$ = TRIM$(Control$("", "CONPAREP"))
    If CONPAREP$ = "" Then CONPAREP$ = "Contado contra entrega"
    '
    VDEF$ = Space$(96) + Space$(48)
    Call REPLA(VDEF$, PLAENREP$, 1, 48)
    Call REPLA(VDEF$, CONPAREP$, 49, 48)
    '
15  OBX$ = OBJPLA$("PRESUREPA", VDEF$)
    If OBX$ = "" Then GoTo 15
    PLAENTRE$ = Left$(OBX$, 48)
    CONDIPAG$ = Mid$(OBX$, 49, 48)
    ATENTX$ = Mid$(OBX$, 97, 48)
    If TRIM$(PLAENTRE$) = "" Or TRIM$(CONDIPAG$) = "" Then
      Call MENSERR(24, "Datos Incompletos")
      VDEF$ = OBX$
      GoTo 15
    End If
    '
20  NC% = Val(Label11)
    TIPODOC$ = "Presup.Reparación (Externo)"
    DESTIDOC% = NC%
    NUMEDOC$ = TRIM$(Label18)
    FECHDOC$ = TRIM$(Text9.TEXT)
    '
    CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Label10)
    CODPARAM$(2) = "DES-CONJ": DOCPARAM$(2) = TRIM$(ECOARCH$("INDIREP", Label10))
    CODPARAM$(3) = "REF-MAT1": DOCPARAM$(3) = TRIM$(Text1)
    CODPARAM$(4) = "MEDIDA-1": DOCPARAM$(4) = TRIM$(Text10)
    CODPARAM$(5) = "REF-MAT2": DOCPARAM$(5) = ""
    CODPARAM$(6) = "REF-MAT3": DOCPARAM$(6) = ""
    CODPARAM$(7) = "PRE-NETO": DOCPARAM$(7) = ""
    CODPARAM$(8) = "BRU-ITEM": DOCPARAM$(8) = ""
    CODPARAM$(9) = "IMP-NETO": DOCPARAM$(9) = TRIM$(TOTAL(0))
    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = TRIM$(TOTAL(1))
    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = TRIM$(TOTAL(2))
    PORI$ = ""
    On Error Resume Next
    PORI$ = TRIM$(Str$(100 * TOTAL(1) / TOTAL(0)))
    On Error GoTo 0
    CODPARAM$(12) = "POR-IVA": DOCPARAM$(12) = TRIM$(PORI$)
    CODPARAM$(13) = "MEDIDA-3": DOCPARAM$(13) = TRIM$(Text10)
    CODPARAM$(14) = "REF-MAT2": DOCPARAM$(14) = TRIM$(PLAENTRE$)
    CODPARAM$(15) = "COND-PAG": DOCPARAM$(15) = TRIM$(CONDIPAG$)
    CODPARAM$(16) = "REF-MAT3": DOCPARAM$(16) = TRIM$(ATENTX$)
    CODPARAM$(17) = "DES-MPRI": DOCPARAM$(17) = Label8
    CAPARDOC% = 17
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CACOLTAB1% = 3
    '
    CAITAB1% = 0
    For U& = 1 To List1.ListCount
      zz$ = List1.List(U& - 1)
      CAITAB1% = CAITAB1% + 1
      TETABU1$(1, CAITAB1%) = Left$(zz$, 16)
      CII% = CODINT%(Left$(zz$, 16))
      TETABU1$(2, CAITAB1%) = DESCRIT$(CII%)
      TETABU1$(3, CAITAB1%) = TRIM$(Mid$(zz$, 56, 8))
    Next U&
    '
    Call PRINTDOC("@" + FORIPRE$)
    '
End Sub

Sub GRAPRESU(OKX%)
   '
   OKX% = 0
'   Call BLOQARCH("PRESREPA")
'   Call BLOQARCH("REPUREPA")
'   '
'   REG& = Val(REPRESUP)
'   If REG& < 1 Then REG& = 1 + ULTREG&("PRESREPA")
'   If REG& < 1 Or REG& > 1 + ULTREG&("PRESREPA") Then
'      Call COMUNI("Imposible Registrar - Error de Indices")
'      GoTo 98
'   End If
'   '
'   XX$ = REGLEIDO$("PRESREPA", REG&)
'   If CVS(Left$(XX$, 4)) <> Val(Label18) Then
'     Call REPLA(XX$, MKS$(Val(Label18)), 1, 4)
'   End If
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = Val(Label11)
   '
   If Len(TRIM$(Text1)) > 15 Then
        MERRO$ = "Exceso Longitud Número de Serie"
        Text1 = Left$(TRIM$(Text1), 15)
        GoTo 98
   End If
   '
   If Len(TRIM$(Text10)) > 9 Then
        MERRO$ = "Exceso Longitud Referencia"
        Text10 = Left$(TRIM$(Text10), 9)
        GoTo 98
   End If
   '
   HOII% = HOY(HOS$)
   FEMI% = FECHANUM(Text9.TEXT)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
   REFEPRE$ = TRIM$(Label10) + " - " + Left$(TRIM$(Label12), 30) + " - R: " + TRIM$(Text10) + " - " + TRIM$(Text9)
   ' Grabacion en Tabla SQL
   CONDI$ = "NumRepa = " & Label18 & ""
   Call ACTUREGISTRO("REPARA", "HoraMObra", CONDI$, TRIM$(Label19), REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "ValMObra", CONDI$, TRIM$(Label20), REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Imp_Mobra", CONDI$, TRIM$(TOTAMOBRA), REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Imp_Repu", CONDI$, TRIM$(SUMAREP), REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Imp_Tot", CONDI$, TRIM$(TOTAL(0)), REGS&, "NOMESS")
     FF% = FECHANUM(Text9): FECHA$ = FETEXCOR1$(FF%)
   Call ACTUREGISTRO("REPARA", "FechPres", CONDI$, FECHA$, REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Referen2", CONDI$, TRIM$(Text10), REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "NumeroSerie", CONDI$, TRIM$(Text1), REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Observa", CONDI$, Text28, REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "CodRepa", CONDI$, Label10, REGS&, "NOMESS")
   
'   Call REPLA(XX$, REFEPRE$, 5, 60)
'   Call REPLA(XX$, MKI$(NCLIE%), 67, 2)
'   Call REPLA(XX$, TRIM$(Text1), 71, 15)
'   Call REPLA(XX$, TRIM$(Text10), 86, 9)
'   Call REPLA(XX$, TRIM$(TALONX$), 95, 6)
'   Call REPLA(XX$, MKS$(Val(Label19)), 101, 4)
'   Call REPLA(XX$, MKS$(Val(Label20)), 105, 4)
'   Call REPLA(XX$, MKS$(Val(TOTAMOBRA)), 109, 4)
'   Call REPLA(XX$, MKS$(Val(SUMAREP)), 113, 4)
'   Call REPLA(XX$, MKS$(Val(TOTAL(0))), 117, 4)
'   Call REPLA(XX$, MKI$(FEMI%), 121, 2)
'   Call REPLA(XX$, Text28 + Space$(368), 129, 368)
'   Call REPLA(XX$, Label10, 497, 16)
'   Call GRAREG("PRESREPA", XX$, REG&)
'   Call CIERRARCH("PRESREPA")
'   '
'   NPLAL& = Val(Label18)
'   NPLAI% = NPLAL& Mod 10000

'   REBLA$ = REGBLAN$("REPUREPA")
'   Call REPLA(REBLA$, MKS$(NPLAL&), 1, 4)
'   Call REPLA(REBLA$, MKI$(NPLAI%), 5, 2)
'   '
'   RFF$ = RECFILT$("REPUREPA", 2, MKI$(NPLAI%))
'   RCC$ = ""
'   URE& = ULTREG&("REPUREPA")
'   For KK& = 1 To Len(RFF$) Step 4
'     RGG& = CVS(Mid$(RFF$, KK&, 4))
'     If RGG& > 0 Then
'       If RGG& <= URE& Then
'         XX$ = REGLEIDO$("REPUREPA", RGG&)
'         If CVS(Left$(XX$, 4)) = NPLAL& Then
'           RCC$ = RCC$ + MKS$(RGG&)
'         End If
'       End If
'     End If
'   Next KK&
'   '
'   KK& = -3
'   For UI% = 1 To List1.ListCount
'     ZZ$ = List1.List(UI% - 1)
'     CODD$ = TRIM$(Left$(ZZ$, 16))
'     CII% = CODINT%(CODD$)
'     If CII% > 0 Then
'       If CII% <= NUMAS% Then
'         CANTI = Val(Mid$(ZZ$, 54, 10))
'         XX$ = REBLA$
'         Call REPLA(XX$, MKI$(CII%), 7, 2)
'         Call REPLA(XX$, MKS$(CANTI), 9, 4)
'         KK& = KK& + 4
'         If KK& <= Len(RCC$) Then
'              RGG& = CVS(Mid$(RCC$, KK&, 4))
'              Call GRAREG("REPUREPA", XX$, RGG&)
'            Else
'              Call REGAPP("REPUREPA", XX$)
'         End If
'       End If
'     End If
'   Next UI%
   
   CONDI$ = "NumRepa = " & Label18 & " AND ORIGEN=1" ' ORIGEN=0 REPUESTOS DE TECNICA Y ORIGEN=1 DE COMERCIAL
   Call ACTUREGISTRO("REPUREPA", "MARBORRA", CONDI$, 1, REGS&, "NOMESS")
   KK& = -3
   For UI% = 1 To List1.ListCount
     zz$ = List1.List(UI% - 1)
     CODD$ = TRIM$(Left$(zz$, 16))
     CII% = CODINT%(CODD$)
     If CII% > 0 Then
       If CII% <= NUMAS% Then
         CANTI = Val(Mid$(zz$, 54, 10))
         KK& = KK& + 4
         '
         SQLX$ = "SELECT * FROM REPUREPA"
         SQLX$ = SQLX$ + " WHERE NumRepa = " & Label18 & ""
         Dim REPUES As ADODB.Recordset
         Set REPUES = New ADODB.Recordset
         REPUES.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
         With REPUES
           .AddNew
           !CodiEmpr = CodiEmp%
           !PRUN_NETO = 0
           !MARBORRA = 0
           !IMPO_NETO = 0
           !NUORD_ITEM = KK&
           !IDENLOTE = ""
           !cod_inte = CII%
           !CANT_SOL = CANTI
           !NUMREPA = Label18
           !ORIGEN = 1
           .Update
         End With
         REPUES.Close
         Set REPUES = Nothing
         '
       End If
     End If
   Next UI%
   CONDI2$ = "NumRepa = " & Label18 & " AND MARBORRA = 1 AND ORIGEN=1" ' ORIGEN=0 REPUESTOS DE TECNICA Y ORIGEN=1 DE COMERCIAL
   Call BORRAREGISTROS("REPUREPA", CONDI2$, REGS&, "NOMESS")
   
'   While KK& < Len(RCC$) - 3
'     KK& = KK& + 4
'     RGG& = CVS(Mid$(RCC$, KK&, 4))
'     Call GRAREG("REPUREPA", String$(128, 0), RGG&)
'   Wend
   '
   Call CIERRARCH("REPUREPA")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   OKX% = 1
   Screen.MousePointer = 1
   Exit Sub
   '
98 Call CIERRARCH("PRESREPA")
   Call CIERRARCH("REPUREPA")
   Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   '
End Sub

Private Sub TOTAMOBRA_Change()
   TOTAMO = Val(SUMAREP) + Val(TOTAMOBRA)
   TOTAL(0) = CSTRING$(MKS$(TOTAMO), 3, 11, 2, 2)
End Sub

Sub EDITREPU()
   '
   Screen.MousePointer = 11
   JJ& = 0
   For UI% = 1 To List1.ListCount
     zz$ = List1.List(UI% - 1)
     CODD$ = TRIM$(Left$(zz$, 16))
     CIXI% = CODINT%(CODD$)
     If CIXI% > 0 Then
       If CIXI% <= NUMAS% Then
         CANTU = Val(Mid$(zz$, 54, 10))
         XX$ = REGBLAN$("MATEROF")
         Call REPLA(XX$, MKI$(CIXI%), 1, 2)
         Call REPLA(XX$, UNIMED$(CIXI%), 3, 2)
         Call REPLA(XX$, MKS$(CANTU), 5, 4)
         Call REPLA(XX$, MKD$(CDbl(CANTU)), 9, 8)
         Call REPLA(XX$, MKS$(0), 17, 4)
         STDIS = STOCKDIS(CIXI%)
         Call REPLA(XX$, MKS$(STDIS), 21, 4)
         FALTA = CANTU - STDIS: If FALTA < 0 Then FALTA = 0
         Call REPLA(XX$, MKS$(FALTA), 25, 4)
         JJ& = JJ& + 1
         Call GRAREG("!MATEROF", XX$, JJ&)
       End If
     End If
   Next UI%
   '
   Call SETULTREG("!MATEROF", JJ&)
   '
20 ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Repuestos para Reparación"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Frame4.Top + Frame1.Top + 110)) 'Height - ScaleHeight))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(60 + Frame4.Left + Frame1.Left))
   '
   VTB% = VENTABU%("EDITREPU" + ATRI$)
   If VTB% < 0 Then
     Exit Sub
   End If
   Call PREREPUES
   '
End Sub

Sub AGREGA_ANOTAREPARA(NORXX&, TXTX$)
    '
    If NORXX& > 0 Then
       ANOTANTE = VALOBADA("REPARA", "ANOTAREP", "NumRepa = '" & NORXX& & "'")
       ANOTA$ = ANOTANTE + TXTX$ + CRLF$
       Call ACTUREGISTRO("REPARA", "ANOTAREP", "NumRepa = '" & NORXX& & "'", ANOTA$, RAFE&)
       '
'        SQLX$ = " SELECT * FROM REPARA "
'        SQLX$ = SQLX$ + " WHERE NumRepa = '" & NORXX& & "'"
'        Dim ABREREPARA As ADODB.Recordset
'        Set ABREREPARA = New ADODB.Recordset
'        ABREREPARA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'        On Error Resume Next
'        '
'        With ABREREPARA
'         If Not (.BOF And .EOF) Then
'            ANOTANTE$ = !AnotaRep
'            ANOTA$ = ANOTANTE$ + TXTX$ + Chr$(10) + Chr$(13)
'            !AnotaRep = ANOTA$
'            .Update
'         End If
'        End With
'        ABREREPARA.Close
'        Set ABREREPARA = Nothing
'        '
    End If
    '
End Sub

Sub GENERAPRESUWORD()

On Error GoTo ErrSub

    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Ubicacion del Presupuesto"
    'CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
    CommonDialog1.FilterIndex = 1
    On Error Resume Next
    CommonDialog1.ShowOpen
    On Error GoTo 0
    RUTAWORD$ = TRIM$(CommonDialog1.FileName)
    If RUTAWORD$ = "" Then
        Call MENSERR(24, "Ubicacion del Presupuesto Erronea")
        Exit Sub
    End If
    '
    'Variable de tipo Word Aplication
    Dim o_Word As Word.Application
    'Variable de objeto para hacer referencia al documento
    Dim documento As Word.Document
    'Variable de tipo tabla para poder hacer referencia a la tabla dentro del Word
    Dim Parrafo As Table
    'La F es para referenciar a las filas y la C para las columnas de la tabla de Word
    Dim f As Long
    Dim c As Integer
    Dim DATOSEMPRESA(4) As String
    
    'nuevo objeto de tipo Word Application
    Set o_Word = New Word.Application
    
    'hacemos visible el documento
    o_Word.Visible = True
    
    'Añadimos un Nuevo documento de word
    Set documento = o_Word.Documents.Add

    'ARMADO DEL ENCABEZADO
    '*********************
    Call AgregaLinea(documento, "", "", 10, True, wdAlignParagraphRight)
    TEXTO1$ = "Buenos Aires, " & FECHATEX$(HOY(HO$))
    Call AgregaLinea(documento, TEXTO1$, "", 12, True, wdAlignParagraphRight)
    Call AgregaLinea(documento, "", "", 10, True, wdAlignParagraphLeft)
    ' OBTIENE DATOS DE LA EMPRESA
    n1% = FILEOPEN(LUDAT$ + "SVD147.EXE", 0, 41)
    Dim X As String * 41
    TEXTO1$ = ""
    For KK% = 2 To 5
        If KK% <= LOF(n1%) / 41 Then
              Get #n1%, KK%, X
              TEXTO1$ = TEXTO1$ + DESENCRI$(X) + vbCrLf
        End If
    Next KK%
    Close #n1%
    '
    TEXTO1$ = Mid(TEXTO1$, 1, Len(TEXTO1$) - 2)
    Call AgregaLinea(documento, TEXTO1$, "", 12, True, wdAlignParagraphLeft)
    Call AgregaLinea(documento, " ", "", 12, True, wdAlignParagraphLeft)
    
    'CUERPO DEL PRESUPUESTO
    '**********************
    TEXTO1$ = "De nuestra mayor consideración:" & vbCrLf & "Adjunto enviamos presupuesto correspondiente" & _
                 " a la reparación nº " & Label18
    Call AgregaLinea(documento, TEXTO1$, "", 11, False, wdAlignParagraphLeft)
    Call AgregaLinea(documento, " ", "", 11, False, wdAlignParagraphLeft)
    CONDI$ = "NUMREPA = " & Label18
    TEXTO1$ = VALOBADA("REPARA", "REPARACION", CONDI$, "NOMESS")
    Call AgregaLinea(documento, TEXTO1$, "", 11, False, wdAlignParagraphLeft)
    Call AgregaLinea(documento, " ", "", 10, False, wdAlignParagraphLeft)
    
    'CARGA DE IMAGENES
    '*****************
    
    'DETERMINA SI EXISTEN IMAGENES HA INCLUIR AL PRESUPUESTO
    CONDI$ = "NUMREPA = " & Label18
    CANTIMAG% = CantRegistros("GALIMREP", CONDI$, "NOMESS")
    NN% = Int(CANTIMAG% / 3): If CANTIMAG% Mod 3 > 0 Then NN% = NN% + 1
    
    If CANTIMAG% > 0 Then
        'ENCABEZA LA PRESENTACION DE LA IMAGENES
        TEXTO1$ = "Adjuntamos imagenes de las fallas encontradas:"
        Call AgregaLinea(documento, TEXTO1$, "", 10, False, wdAlignParagraphLeft)
        Call AgregaLinea(documento, " ", "", 10, False, wdAlignParagraphLeft)
        
        'CREA LA TABLA
        Set Parrafo1 = documento.Paragraphs.Add
        Set Rango = Parrafo1.Range.Next
        Set Parrafo = documento.Tables.Add(Rango, NN%, 3)
        
        'OBTIENE LAS IMAGENES DESDE LA TABLA SQL
        SQLX$ = " SELECT * FROM GALIMREP "
        SQLX$ = SQLX$ + " WHERE NUMREPA = " & Label18 & " "
        SQLX$ = SQLX$ + " ORDER BY Norit_Imag ASC "
        Dim IMGTMP As ADODB.Recordset
        Set IMGTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        '
        With IMGTMP
            .MoveFirst
            For f = 1 To NN%
                    For c = 1 To 3
                        'LIMPIA EL PORTAPAPELES
                        Clipboard.Clear ' limpia el portapapeles
                        
                        'CARGA LA IMAGEN QUE ESTA EN LA TABLA SQL EN UN OBJETO IMAGE
                        FINAM$ = SAFETEXT$(!Nom_Imag)
                        XXXX$ = SAFETEXT$(!Imagen)
                        Call SaveFile(RUTATMP$ + "\" + FINAM$, XXXX$)
                        Image2.Picture = LoadPicture(RUTATMP$ & "\" & FINAM$)
                        Image2.Refresh
                        
                        'CARGA LA IMAGEN EN EL PORTAPAPELES
                        Clipboard.SetData Image2.Picture
                        'PEGA LA IMAGEN EN LA POSICION F,C DE LA TABLA
                        Parrafo.Cell(f, c).Range.Paste
                        
                        .MoveNext
                        If .EOF Then Exit For
                    Next c
            Next f
        End With
    End If
    
    'CARGA DE REPUESTOS
    '******************
       
    CONDI$ = "NUMREPA = " & Label18 & " AND ORIGEN=1" ' ORIGEN=0 REPUESTOS DE TECNICA Y ORIGEN=1 DE COMERCIAL
    CANREP% = CantRegistros("REPUREPA", CONDI$, "NOMESS")
    
    If CANREP% > 0 Then
        'ENCABEZA LA PRESENTACION DE LOS REPUESTOS
        Call AgregaLinea(documento, " ", "", 10, False, wdAlignParagraphLeft)
        TEXTO1$ = "Adjuntamos Informacion de los Repuestos:"
        Call AgregaLinea(documento, TEXTO1$, "", 10, False, wdAlignParagraphLeft)
        Call AgregaLinea(documento, " ", "", 10, False, wdAlignParagraphLeft)

        'CREA LA TABLA Y ESTABLECE EL ENCABEZADO DE LAS COLUMNAS
        Set Parrafo2 = documento.Paragraphs.Add
        Set Rango = Parrafo2.Range.Next
        Set Parrafo2 = documento.Tables.Add(Rango, CANREP% + 1, 6, wdWord9TableBehavior, wdAutoFitContent)
        
        Parrafo2.Cell(1, 1).Range.InsertAfter "Orden"
        Parrafo2.Cell(1, 2).Range.InsertAfter "Codigo"
        Parrafo2.Cell(1, 3).Range.InsertAfter "Descripcion"
        Parrafo2.Cell(1, 4).Range.InsertAfter "Cantidad"
        Parrafo2.Cell(1, 5).Range.InsertAfter "Pr.Unitario"
        Parrafo2.Cell(1, 6).Range.InsertAfter "Importe"
        
        'OBTIENE LISTA DE REPUESTOS DE LA TABLA SQL
        SQLX$ = "SELECT * FROM REPUREPA "
        SQLX$ = SQLX$ + " WHERE NUMREPA = " & Label18 & " AND ORIGEN=1" ' ORIGEN=0 REPUESTOS DE TECNICA Y ORIGEN=1 DE COMERCIAL
        Dim REPU As ADODB.Recordset
        Set REPU = FLEXCONN.Execute(FILTSQL$(SQLX$))
        '
        'COMPLETA LA TABLA CON LA INFORMACION
        FF = 1
        COEFI = Val(Control$("REPARA", "COEFPREA"))
        With REPU
          Do While .EOF = False
             CAN# = XVALO(!CANT_SOL)
             Ci% = XVALO(!cod_inte)
             Codigo$ = CODEXT$(Ci%)
             DESCRI$ = DESCRIT0$(Ci%)
             CANTI$ = FORMATNUM$(CAN#, "F12.2")
             COSUN = COEFI * COSUNI(Ci%)
             COSTO = COSUN * CAN#
             
             FF = FF + 1
             Parrafo2.Cell(FF, 1).Range.InsertAfter FF - 1
             Parrafo2.Cell(FF, 2).Range.InsertAfter Codigo$
             Parrafo2.Cell(FF, 3).Range.InsertAfter DESCRI$
             Parrafo2.Cell(FF, 4).Range.InsertAfter CANTI$
             Parrafo2.Cell(FF, 5).Range.InsertAfter FORMATNUM$(COSUN, "F12.2")
             Parrafo2.Cell(FF, 6).Range.InsertAfter FORMATNUM$(COSTO, "F12.2")
          .MoveNext
          Loop
        End With
    End If

    'DESCARGA EL OBJETO Y GUARDA EL ARCHIVO DOC
10  documento.SaveAs (RUTAWORD$ & ".DOC")
    Set o_Word = Nothing
    Set documento = Nothing
    Set Parrafo = Nothing

Exit Sub

ErrSub:

MsgBox Err.Description, vbCritical
    On Error Resume Next
    Set Parrafo = Nothing
    Set documento = Nothing
    Set o_Word = Nothing
End Sub

Sub AgregaLinea(documento, TEXTO$, FUENTE$, Tamanio, Negrita, alineacion)

    Set Parrafo = documento.Paragraphs.Add
    Set Rango = Parrafo.Range.Next
    Rango.ParagraphFormat.Alignment = alineacion
    Rango.Font.Size = Tamanio
    Rango.Font.Bold = Negrita
    Rango.InsertBefore TEXTO$

End Sub

