VERSION 5.00
Begin VB.Form AJINFLA 
   BackColor       =   &H00D0FFEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuste por Inflación"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   Begin VB.CommandButton Command6 
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
      Left            =   7035
      Picture         =   "AJINFLA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Inventarios Mercaderia de Importación"
      Top             =   315
      Width           =   750
   End
   Begin VB.CommandButton Command7 
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
      Left            =   7875
      Picture         =   "AJINFLA.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Re-Cálculo de Saldos Mensuales"
      Top             =   315
      Width           =   750
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Reportes y Listados"
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
      Left            =   7035
      TabIndex        =   21
      Top             =   1260
      Width           =   4740
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   3150
         ScaleHeight     =   585
         ScaleWidth      =   1395
         TabIndex        =   51
         Top             =   420
         Width           =   1450
         Begin VB.CommandButton Command4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "AJINFLA.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Listado Resumen R.E.I."
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado Resum."
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
            Left            =   630
            TabIndex        =   53
            Top             =   105
            Width           =   750
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1630
         ScaleHeight     =   585
         ScaleWidth      =   1395
         TabIndex        =   48
         Top             =   420
         Width           =   1450
         Begin VB.CommandButton Command2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "AJINFLA.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Listado Analítico R.E.I."
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado Analitico"
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
            Left            =   630
            TabIndex        =   50
            Top             =   105
            Width           =   750
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   105
         ScaleHeight     =   585
         ScaleWidth      =   1395
         TabIndex        =   45
         Top             =   420
         Width           =   1450
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
            Height          =   600
            Left            =   0
            Picture         =   "AJINFLA.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Asiento Contable R.E.I."
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Asiento Contab."
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
            Left            =   630
            TabIndex        =   47
            Top             =   105
            Width           =   750
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Configuración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   105
      TabIndex        =   7
      Top             =   210
      Width           =   6840
      Begin VB.Frame Frame1 
         BackColor       =   &H00D0FFEE&
         Caption         =   "Indices y Monedas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   210
         TabIndex        =   12
         Top             =   420
         Width           =   3165
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   13
            Top             =   840
            Width           =   2325
            Begin VB.CommandButton Command18 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "AJINFLA.frx":0F32
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Actualizacion Indices y Cotizaciones Mensuales"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Actualización Valores Mensuales"
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
               Left            =   630
               TabIndex        =   15
               Top             =   105
               Width           =   1650
            End
         End
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   16
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command19 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   -20
               Picture         =   "AJINFLA.frx":107C
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "Actualizar Tabla de Indices y Monedas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Tabla de Indi- ces y Monedas"
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
               Left            =   735
               TabIndex        =   18
               Top             =   0
               Width           =   1500
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00D0FFEE&
         Caption         =   "Base de Cuentas Contables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   3465
         TabIndex        =   8
         Top             =   420
         Width           =   3165
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   9
            Top             =   840
            Width           =   2325
            Begin VB.CommandButton Command8 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   -20
               Picture         =   "AJINFLA.frx":11C6
               Style           =   1  'Graphical
               TabIndex        =   10
               ToolTipText     =   "Tabla de Indices de Ajuste por Cuenta"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Tabla de Ajus- tes por Cuenta"
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
               Left            =   735
               TabIndex        =   11
               Top             =   105
               Width           =   1500
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   19
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command22 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "AJINFLA.frx":1310
               Style           =   1  'Graphical
               TabIndex        =   44
               ToolTipText     =   "Plan General de Cuentas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Plan de Cuentas Contables"
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
               Left            =   735
               TabIndex        =   20
               Top             =   0
               Width           =   1500
            End
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Ajuste por Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5610
      Left            =   105
      TabIndex        =   5
      Top             =   2625
      Width           =   11670
      Begin VB.Frame Frame6 
         BackColor       =   &H00D0FFEE&
         Caption         =   "Cuenta Activa"
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
         Index           =   2
         Left            =   105
         TabIndex        =   39
         Top             =   420
         Width           =   3690
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
            Left            =   1260
            TabIndex        =   41
            Top             =   360
            Width           =   2115
         End
         Begin VB.CommandButton Command5 
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
            Height          =   300
            Left            =   3360
            TabIndex        =   40
            Top             =   360
            Width           =   175
         End
         Begin VB.Label Label1 
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
            Left            =   210
            TabIndex        =   43
            Top             =   420
            Width           =   960
         End
         Begin VB.Label Label2 
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
            Height          =   300
            Left            =   210
            TabIndex        =   42
            Top             =   735
            Width           =   3375
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   11580
         TabIndex        =   37
         Top             =   5145
         Width           =   11640
         Begin VB.Label Label13 
            Caption         =   "Resultado Total"
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
            Left            =   40
            TabIndex        =   38
            Top             =   105
            Width           =   11565
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   11580
         TabIndex        =   35
         Top             =   1785
         Width           =   11640
         Begin VB.Label Label12 
            Caption         =   "Mes              S.Inicial     S.Final   Resultado    I.Act.       R.E.I    Cotiz.   Dif.Cambio"
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
            Left            =   40
            TabIndex        =   36
            Top             =   105
            Width           =   11565
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00D0FFEE&
         Caption         =   "Diferencia de Cambio"
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
         Index           =   1
         Left            =   7770
         TabIndex        =   29
         Top             =   420
         Width           =   3795
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Indice:"
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
            Left            =   -210
            TabIndex        =   34
            Top             =   420
            Width           =   960
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta:"
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
            Left            =   1155
            TabIndex        =   33
            Top             =   420
            Width           =   960
         End
         Begin VB.Label Label8 
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
            Height          =   300
            Index           =   1
            Left            =   840
            TabIndex        =   32
            Top             =   360
            Width           =   435
         End
         Begin VB.Label Label9 
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
            Height          =   300
            Index           =   1
            Left            =   2205
            TabIndex        =   31
            Top             =   360
            Width           =   1380
         End
         Begin VB.Label Label10 
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
            Height          =   300
            Index           =   1
            Left            =   840
            TabIndex        =   30
            Top             =   735
            Width           =   2745
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00D0FFEE&
         Caption         =   "R.E.I."
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
         Index           =   0
         Left            =   3885
         TabIndex        =   23
         Top             =   420
         Width           =   3795
         Begin VB.Label Label10 
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
            Height          =   300
            Index           =   0
            Left            =   840
            TabIndex        =   28
            Top             =   735
            Width           =   2745
         End
         Begin VB.Label Label9 
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
            Height          =   300
            Index           =   0
            Left            =   2205
            TabIndex        =   27
            Top             =   360
            Width           =   1380
         End
         Begin VB.Label Label8 
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
            Height          =   300
            Index           =   0
            Left            =   840
            TabIndex        =   26
            Top             =   360
            Width           =   435
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta:"
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
            Left            =   1155
            TabIndex        =   25
            Top             =   420
            Width           =   960
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Indice:"
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
            Left            =   -210
            TabIndex        =   24
            Top             =   420
            Width           =   960
         End
      End
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
         Height          =   2985
         Left            =   0
         TabIndex        =   6
         Top             =   2205
         Width           =   11640
      End
   End
   Begin VB.CommandButton Command3 
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
      Left            =   9345
      Picture         =   "AJINFLA.frx":145A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprime Ficha de Análisis"
      Top             =   315
      Width           =   750
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   765
      Left            =   8925
      ScaleHeight     =   705
      ScaleWidth      =   60
      TabIndex        =   3
      Top             =   315
      Width           =   120
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command24 
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
      Left            =   11025
      Picture         =   "AJINFLA.frx":1AC4
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   315
      Width           =   750
   End
   Begin VB.CommandButton Command29 
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
      Left            =   10185
      Picture         =   "AJINFLA.frx":1C0E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   315
      Width           =   750
   End
End
Attribute VB_Name = "AJINFLA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EXTE$
Dim MESAX$(13), INDIME#(13)
Dim INDI$(256)
Dim BULEIN As String * 132
Dim INDAJ#(256, 13), YACAIN%
Dim SALCUEN#(16384, 13)

Private Sub Command1_Click()
   Command1.Enabled = False
   Call ASIENREI
   Command1.Enabled = True
End Sub
'
Private Sub Command18_Click()
   '
   Command18.Enabled = False
   If ULTREG&("TAMONED") < 1 Then
      Call MENSERR(24, "Falta Definir Tabla\de Indices y Monedas de Ajuste.")
      GoTo 99
   End If
   '
10 REMO$ = REGSEL$("TAMONED")
   If Len(REMO$) <= 4 Then GoTo 99
   '
   INDIAJ% = Asc(Left$(REMO$, 1))
   DENOIN$ = TRIM$(Mid$(REMO$, 2, 20))
   '
   Call CARTAMESES
   Call LEEINDIME(INDIAJ%)
   '
   For U& = 0 To 12
      TTXX$ = String$(24, 0)
      Call REPLA(TTXX$, Chr$(U&), 1, 1)
      Call REPLA(TTXX$, MESAX$(U&), 2, 15)
      Call REPLA(TTXX$, MKD$(INDIME#(U&)), 17, 8)
      Call GRAREG("!TAINDAJ", TTXX$, U& + 1)
   Next U&
   Call SETULTREG("!TAINDAJ", 13)
   Call CIERRARCH("!TAINDAJ")
   '
   VTB% = VENTABU%("TAINDAJ/NC/TITUPAN=Indice de Ajuste: " + DENOIN$)
   If VTB% >= 0 Then
     EXTE$ = "DAT": If EJER% < 0 Then EXTE$ = "OLD"
     N1% = FILEOPEN%(LUDAT$ + "INDAJU." + EXTE$, 0, 132)
     CAREIN& = LOF(N1%) / 132
     If INDIAJ% <= CAREIN& Then
       Get #N1%, INDIAJ%, BULEIN$
       TTXX$ = BULEIN$
       For U& = 0 To 12
         XX$ = REGLEIDO$("!TAINDAJ", U& + 1)
         Call REPLA(TTXX$, Mid$(XX$, 17, 8), 21 + 8 * U&, 8)
       Next U&
       BULEIN$ = TTXX$
       Put #N1%, INDIAJ%, BULEIN$
       Close #N1%: N1% = 0
       Call CIERRARCH("*.*")
     End If
     GoTo 10
     '
   End If
   '
99 YACAIN% = 0
   Command18.Enabled = True
   '
End Sub

Private Sub Command19_Click()
   '
   Command19.Enabled = False
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAINDICE/NC/TITUPAN=Indices y Monedas")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", X$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
   Command19.Enabled = True
   '
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call LIANAREI
   Command2.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   PLAGRAL.Show 1
   Command22.Enabled = True
End Sub

Private Sub Command24_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command29_Click()
   Call HELPONLINE("AJUINFLA")
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   PRINTFORM
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call LIRESREI
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Call SELECHO("ECUECON")
   VDEVU$ = TRIM$(VALACT1$("ECUECON"))
   If VDEVU$ <> "" Then
      Label2 = VALACT2("ECUECON")
      Text1 = VDEVU$
   End If
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   INVEIMPO.Show 1
   Command6.Enabled = True
End Sub

Private Sub RESALMEN()
   '
   Screen.MousePointer = 11
   '
   FEX = COMEJ%
   MESIN% = Val(Mid$(FECHATEX$(FEX), 4, 2))
   DES% = COMEJ%
   CUEMA% = 0
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   FIN& = MAXDIARIO%
   For KK% = 1 To MAXDIARIO%
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      Call LEEDIARIO(KK%)
      CUE% = CVI(E11$)
      NUAS = CVS(E12$)
      FEAS% = CVI(E15$)
      IPAS# = CVD(E16$)
      '
      If CUE% > 0 Then
        If CUE% <= 16384 Then
          If CUE% > CUEMA% Then CUEMA% = CUE%
          FEX = FEAS%
          MESAS% = 1 + Val(Mid$(FECHATEX$(FEX), 4, 2)) - MESIN%
          If MESAS% < 1 Then MESAS% = MESAS% + 12
          If FEAS% < DES% Or (FEAS% = DES% And NUAS = 1) Then
            MESAS% = 0
          End If
          If MESAS% = 12 Then
            Call LEEPASDIA(KK%, 2)
            REFEX$ = UCase$(REFASIEN$)
            If InStr(REFEX$, "(CIERR") > 0 Or InStr(REFEX$, "(*)") > 0 Then
              IPAS# = 0
            End If
          End If
          For KKI% = MESAS% To 13
            SALCUEN#(CUE%, KKI%) = SALCUEN#(CUE%, KKI%) + (Int(100 * IPAS# + 0.5)) / 100
          Next KKI%
        End If
      End If
   Next KK%
   '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   FIN& = MAXDIA&
   For KKL& = 1 To MAXDIA&
      UU& = KKL&
      Call TRACE(20, UU&, FIN&)
      Call LEESUCAJA(KKL&)
      If Abs(CVS(D19$)) < 0.1 Then
        CUE% = CVI(D11$)
        FEAS% = CVI(D15$)
        IPAS# = CVD(D16$)
        If CUE% > 0 Then
          If CUE% <= 16384 Then
            If CUE% > CUEMA% Then CUEMA% = CUE%
            FEX = FEAS%
            MESAS% = 1 + Val(Mid$(FECHATEX$(FEX), 4, 2)) - MESIN%
            If MESAS% < 1 Then MESAS% = MESAS% + 12
            If FEAS% < DES% Or (FEAS% = DES% And NUAS = 1) Then
              MESAS% = 0
            End If
            For KKI% = MESAS% To 13
              SALCUEN#(CUE%, KKI%) = SALCUEN#(CUE%, KKI%) + (Int(100 * IPAS# + 0.5)) / 100
            Next KKI%
          End If
        End If
      End If
   Next KKL&
   '
   On Error Resume Next
   FECH1# = FileDateTime(LUDAT$ + "DIARIO0." + EXTE$)
   FECH2# = FileDateTime(LUDAT$ + "SUCAJA0." + EXTE$)
   On Error GoTo 0
   FECHAC# = FECH1#: If FECH2# > FECH1# Then FECHAC# = FECH2#
   Call GRACONTROL("RESALMEN", EXTE$, Str$(FECHAC#))
   '
   Dim BULECIE As String * 112
   NCIE% = FILEOPEN%(LUDAT$ + "CIERRE0." + EXTE$, 0, 112)
   For CUE% = 1 To CUEMA%
     TTXX$ = String$(112, 0)
     For KKI% = 0 To 13
       Call REPLA(TTXX$, MKD$(SALCUEN#(CUE%, KKI%)), 8 * KKI% + 1, 8)
     Next KKI%
     BULECIE$ = TTXX$
     Put #NCIE%, CUE%, BULECIE$
   Next CUE%
   Close #NCIE%: NCIE% = 0
   '
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call RESALMEN
   YACAIN% = 0
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   Dim BULECUE As String * 64
   MODO% = COMALTER%("Puede Actualizaren Forma Individual\o por Rango de Cuentas\\Indvidual\Rango")
   On MODO% GoTo 10, 20
   GoTo 99
   '
10 Screen.MousePointer = 11
   REBLA$ = REGBLAN$("CARAJUCU")
   N1% = FILEOPEN%(LUDAT$ + "CUENTAS." + EXTE$, 0, 64)
   CAREIN& = LOF(N1%) / 64
   '
   JJ& = 0
   For U& = 1 To ULTREG("ECUECON")
      XX$ = REGLEIDO$("ECUECON", U&)
      CUECO$ = Left$(XX$, 12)
      DENOC$ = Mid$(XX$, 13, 28)
      CUEI% = CUEINT%(CUECO$)
      If CUEI% > 0 Then
        If CUEI% <= CAREIN& Then
          Get #N1%, CUEI%, BULECUE$
          TICU$ = Mid$(BULECUE$, 13, 1)
          IAJU1% = Asc(Mid$(BULECUE$, 14, 1))
          CONTRA1% = CVI(Mid$(BULECUE$, 15, 2))
          CPAR1$ = CODCUE$(CONTRA1%)
          IAJU2% = Asc(Mid$(BULECUE$, 17, 1))
          CONTRA2% = CVI(Mid$(BULECUE$, 18, 2))
          CPAR2$ = CODCUE$(CONTRA2%)
          '
          YY$ = REBLA$
          Call REPLA(YY$, CUECO$, 1, 12)
          Call REPLA(YY$, DENOC$, 13, 19)
          Call REPLA(YY$, TICU$, 32, 1)
          Call REPLA(YY$, MKI$(CUEI%), 33, 2)
          Call REPLA(YY$, Chr$(IAJU1%), 35, 1)
          Call REPLA(YY$, CPAR1$, 36, 12)
          Call REPLA(YY$, MKI$(CONTRA1%), 48, 2)
          Call REPLA(YY$, Chr$(IAJU2%), 50, 1)
          Call REPLA(YY$, CPAR2$, 51, 12)
          Call REPLA(YY$, MKI$(CONTRA2%), 63, 2)
          '
          JJ& = JJ& + 1
          Call GRAREG("!CARAJUCU", YY$, JJ&)
          '
         End If
       End If
   Next U&
   '
   Call SETULTREG("!CARAJUCU", JJ&)
   Call CIERRARCH("!CARAJUCU")
   Screen.MousePointer = 1
   '
   VTB% = VENTABU%("CARAJUCU")
   If VTB% >= 0 Then
     Screen.MousePointer = 11
     For U& = 1 To ULTREG&("!CARAJUCU")
       YY$ = REGLEIDO$("!CARAJUCU", U&)
       CUEI% = CVI(Mid$(YY$, 33, 2))
       If CUEI% > 0 Then
         If CUEI% <= CAREIN& Then
           IAJU1% = Asc(Mid$(YY$, 35, 1))
           CONTRA1% = 0
           CPAR1$ = TRIM$(Mid$(YY$, 36, 12))
           If CPAR1$ <> "" Then
             CONTRA1% = CUEINT%(CPAR1$)
             If CONTRA1% > CAREIN& Then CONTRA1% = 0
           End If
           IAJU2% = Asc(Mid$(YY$, 50, 1))
           CONTRA2% = 0
           CPAR2$ = TRIM$(Mid$(YY$, 51, 12))
           If CPAR2$ <> "" Then
             CONTRA2% = CUEINT%(CPAR2$)
             If CONTRA2% > CAREIN& Then CONTRA2% = 0
           End If
           Get #N1%, CUEI%, BULECUE$
           TTXX$ = BULECUE$
           Call REPLA(TTXX$, Chr$(IAJU1%), 14, 1)
           Call REPLA(TTXX$, MKI$(CONTRA1%), 15, 2)
           Call REPLA(TTXX$, Chr$(IAJU2%), 17, 1)
           Call REPLA(TTXX$, MKI$(CONTRA2%), 18, 2)
           BULECUE$ = TTXX$
           Put #N1%, CUEI%, BULECUE$
         End If
       End If
     Next U&
   End If
   '
   Close #N1%: N1% = 0
   Screen.MousePointer = 1
   GoTo 99
   '
20 OBX$ = OBJPLA$("RACUEINDIAJU", VDEF$)
   VDEF$ = OBX$
   If OBX$ <> "" Then
     DECUE$ = Left$(OBX$, 12)
     If CUEINT%(DECUE$) < 1 Or CUEINT%(DECUE$) > NUCUEN% Then GoTo 20
     HACUE$ = Mid$(OBX$, 13, 12)
     If CUEINT%(HACUE$) < 1 Or CUEINT%(HACUE$) > NUCUEN% Then GoTo 20
     If HACUE$ < DECUE$ Then GoTo 20
     IAC1% = Asc(Mid$(OBX$, 25, 1))
     CAJ1% = 0
     If IAC1% > 0 Then
       CUEAJ1$ = Mid$(OBX$, 26, 12)
       CAJ1% = CUEINT%(CUEAJ1$)
       If CAJ1% < 1 Or CAJ1% > NUCUEN% Then GoTo 20
     End If
     IAC2% = Asc(Mid$(OBX$, 38, 1))
     If IAC2% > 0 Then
       CUEAJ2$ = Mid$(OBX$, 39, 12)
       CAJ2% = CUEINT%(CUEAJ2$)
       If CAJ2% < 1 Or CAJ2% > NUCUEN% Then GoTo 20
     End If
     '
     Screen.MousePointer = 11
     N1% = FILEOPEN%(LUDAT$ + "CUENTAS." + EXTE$, 0, 64)
     CAREIN& = LOF(N1%) / 64
     '
     For U& = 1 To CAREIN&
       Call TRACE(20, U&, CAREIN&)
       CUE% = U&
       CUECO$ = CODCUE$(CUE%)
       If CUECO$ >= DECUE$ Then
         If CUECO$ <= HACUE$ Then
           Get #N1%, U&, BULECUE$
           TTXX$ = BULECUE$
           Call REPLA(TTXX$, Chr$(IAC1%), 14, 1)
           Call REPLA(TTXX$, MKI$(CAJ1%), 15, 2)
           Call REPLA(TTXX$, Chr$(IAC2%), 17, 1)
           Call REPLA(TTXX$, MKI$(CAJ2%), 18, 2)
           BULECUE$ = TTXX$
           Put #N1%, U&, BULECUE$
         End If
       End If
     Next U&
     Close #N1%
     Screen.MousePointer = 1
     '
   End If
   '
99 YACAIN% = 0
   Command8.Enabled = True
   '
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + " - " + EJACT$ + " - " + EPAC$
    End If
    '
    Show
    Refresh
    EXTE$ = "DAT": If EJER% < 0 Then EXTE$ = "OLD"
    On Error Resume Next
    FECH1# = FileDateTime(LUDAT$ + "DIARIO0." + EXTE$)
    FECH2# = FileDateTime(LUDAT$ + "SUCAJA0." + EXTE$)
    On Error GoTo 0
    FECHAC# = FECH1#: If FECH2# > FECH1# Then FECHAC# = FECH2#
    If Val(CONTROL$("RESALMEN", EXTE$)) < FECHAC# - 0.000001 Then
      OPX% = COMALTER%("FLEXOFT Debe Ahora\Re-calcular los Saldos Mensuales\\Re-calcular\Abandonar")
      If OPX% <> 1 Then
        Call CIERRARCH("*.*")
        Call FINAL("")
        Exit Sub
      End If
      Call RESALMEN
    End If
    '
End Sub

Sub CARTAMESES()
   '
   MESAX$(0) = "APERTURA"
   FEX = COMEJ%
   MESAX$(1) = METEXTO$(Int(FEX))
   JJ% = 1
10 MEBA% = Val(Mid$(FECHATEX$(FEX), 4, 2))
12 FEX = DIPOST(FEX)
   If Val(Mid$(FECHATEX$(FEX), 4, 2)) <> MEBA% Then
     JJ% = JJ% + 1
     MESAX$(JJ%) = METEXTO$(Int(FEX))
     If JJ% < 12 Then GoTo 10
     GoTo 20
   End If
   GoTo 12
   '
20 JJ% = JJ% + 1
   MESAX$(JJ%) = "CIERRE"
   '
End Sub

Sub LEEINDIME(NUINDI%)
    '
    N1% = FILEOPEN%(LUDAT$ + "INDAJU." + EXTE$, 0, 132)
    CAREIN& = LOF(N1%) / 132
    For U& = 1 To ULTREG&("TAMONED")
      DENOIN$ = Mid$(REGLEIDO$("TAMONED", U&), 2, 20)
      If U& <= CAREIN& Then
           Get #N1%, U&, BULEIN$
           TTXX$ = BULEIN$
         Else
           TTXX$ = Space$(20) + String$(112, 0)
      End If
      INDI$(U&) = DENOIN$
      Call REPLA(TTXX$, DENOIN$, 1, 20)
      BULEIN$ = TTXX$
      Put #N1%, U&, BULEIN$
      '
      If U& = NUINDI% Then
        For KKI% = 0 To 13
          INDIME#(KKI%) = CVD(Mid$(BULEIN$, 21 + 8 * KKI%, 8))
        Next KKI%
      End If
      '
    Next U&
    Close #N1%
    '
End Sub

Sub CARMAINDI()
    '
    EXTE$ = "DAT": If EJER% < 0 Then EXTE$ = "OLD"
    N1% = FILEOPEN%(LUDAT$ + "INDAJU." + EXTE$, 0, 132)
    CAREIN& = LOF(N1%) / 132
    For U& = 1 To CAREIN&
      Get #N1%, U&, BULEIN$
      TTXX$ = BULEIN$
      INDA1# = 0
      For KKI% = 0 To 12
        INDA2# = CVD(Mid$(BULEIN$, 21 + 8 * KKI%, 8))
        If INDA2# > 0.0001 Then INDA1# = INDA2#
        INDAJ#(U&, KKI%) = INDA1#
      Next KKI%
    Next U&
    Close #N1%
    YACAIN% = 1
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Text1_Change()
   '
   Static BULECUE As String * 64
   '
   CUE% = CUEINT%(Text1)
   Label2 = DENOCUE$(CUE%)
   If CUE% > 0 Then
      N1% = FILEOPEN%(LUDAT$ + "CUENTAS." + EXTE$, 0, 64)
      Get #N1%, CUE%, BULECUE$
      Close #N1%
      TICU$ = Mid$(BULECUE$, 13, 1)
      IAJU1% = Asc(Mid$(BULECUE$, 14, 1))
      Label8(0) = TRIM$(Str$(IAJU1%))
      Label10(0) = ECOARCH$("TAMONED", Chr$(IAJU1%))
      CONTRA1% = CVI(Mid$(BULECUE$, 15, 2))
      CPAR1$ = CODCUE$(CONTRA1%)
      Label9(0) = CPAR1$
      IAJU2% = Asc(Mid$(BULECUE$, 17, 1))
      Label8(1) = TRIM$(Str$(IAJU2%))
      Label10(1) = ECOARCH$("TAMONED", Chr$(IAJU2%))
      CONTRA2% = CVI(Mid$(BULECUE$, 18, 2))
      CPAR2$ = CODCUE$(CONTRA2%)
      Label9(1) = CPAR2$
      '
      Call AJUINFLA(CUE%)
      '
   End If
   '
End Sub

Private Sub Text1_DblClick()
   Call SELECHO("ECUECON")
   VDEVU$ = TRIM$(VALACT1$("ECUECON"))
   If VDEVU$ <> "" Then
      Label2 = VALACT2("ECUECON")
      Text1 = VDEVU$
   End If
End Sub

Sub AJUINFLA(CUE%)
   '
   If YACAIN% < 1 Then
     Call CARTAMESES
     Call CARMAINDI
     Call CARSALMEN
     YACAIN% = 1
   End If
   '
   Dim BULECUE As String * 64
   EXTE$ = "DAT": If EJER% < 0 Then EXTE$ = "OLD"
   N1% = FILEOPEN%(LUDAT$ + "CUENTAS." + EXTE$, 0, 64)
   Get #N1%, CUE%, BULECUE$
   Close #N1%
   TICU$ = Mid$(BULECUE$, 13, 1)
   IAJU1% = Asc(Mid$(BULECUE$, 14, 1))
   CONTRA1% = CVI(Mid$(BULECUE$, 15, 2))
   CPAR1$ = CODCUE$(CONTRA1%)
   IAJU2% = Asc(Mid$(BULECUE$, 17, 1))
   CONTRA2% = CVI(Mid$(BULECUE$, 18, 2))
   CPAR2$ = CODCUE$(CONTRA1%)
   '
   SUMHIS# = 0: SUMAJU# = 0
   MASCA$ = "F12.2"
   List1.Clear
   SURESU# = 0: SUREI# = 0: SUDICA# = 0
   For MES% = 0 To 12
     TTXX$ = Space$(132)
     Call REPLA(TTXX$, MESAX$(MES%), 1, 20)
     SALANT# = 0: If MES% > 0 Then SALANT# = SALCUEN#(CUE%, MES% - 1)
     Call REPLA(TTXX$, FORMATNUM$(SALANT#, MASCA$), 15, 12)
     Call REPLA(TTXX$, FORMATNUM$(SALCUEN#(CUE%, MES%), MASCA$), 27, 12)
     VARIA# = SALCUEN#(CUE%, MES%) - SALANT#
     Call REPLA(TTXX$, FORMATNUM$(VARIA#, MASCA$), 39, 12)
     SURESU# = SURESU# + VARIA#
     '
     IAJX# = INDAJ#(IAJU1%, MES%)
     Call REPLA(TTXX$, FORMATNUM$(IAJX#, "F10.4"), 51, 10)
     RAJU1# = 0
     If INDAJ#(IAJU1%, MES%) > 0.0001 Then
       RAJU1# = VARIA# * (INDAJ#(IAJU1%, 12) / INDAJ#(IAJU1%, MES%) - 1)
     End If
     Call REPLA(TTXX$, FORMATNUM$(RAJU1#, MASCA$), 61, 12)
     SUREI# = SUREI# + RAJU1#
     '
     IAJX# = INDAJ#(IAJU2%, MES%)
     Call REPLA(TTXX$, FORMATNUM$(IAJX#, "F10.4"), 73, 10)
     RAJU2# = 0
     If INDAJ#(IAJU2%, MES%) > 0.0001 Then
       RAJU2# = VARIA# * (INDAJ#(IAJU2%, 12) / INDAJ#(IAJU2%, MES%) - 1)
       Call REPLA(TTXX$, FORMATNUM$(RAJU2# - RAJU1#, MASCA$), 83, 12)
       SUDICA# = SUDICA# + RAJU2# - RAJU1#
     End If
     '
     List1.AddItem TTXX$
   Next MES%
   '
   TTXX$ = Space$(132)
   Call REPLA(TTXX$, "Resultado Total", 1, 20)
   Call REPLA(TTXX$, FORMATNUM$(SURESU#, MASCA$), 39, 12)
   Call REPLA(TTXX$, FORMATNUM$(SUREI#, MASCA$), 61, 12)
   Call REPLA(TTXX$, FORMATNUM$(SUDICA#, MASCA$), 83, 12)
   Label13.Caption = TTXX$
   '
End Sub

Sub CARSALMEN()
   '
   Screen.MousePointer = 11
   Dim BULECIE As String * 112
   NCIE% = FILEOPEN%(LUDAT$ + "CIERRE0." + EXTE$, 0, 112)
   FIN& = LOF(NCIE%) / 112
   If FIN& > 16384 Then FIN& = 16384
   For CUE% = 1 To FIN&
     Get #NCIE%, CUE%, BULECIE$
     For KKI% = 0 To 13
       SALCUEN#(CUE%, KKI%) = CVD(Mid$(BULECIE$, 8 * KKI% + 1, 8))
     Next KKI%
   Next CUE%
   Close #NCIE%: NCIE% = 0
   Screen.MousePointer = 1
   '
End Sub

Sub LIANAREI()
   '
   TTT1$ = Text1.TEXT
   X1$ = Left$(REGLEIDO$("ECUECON", 1), 12)
   X2$ = Left$(REGLEIDO$("ECUECON", ULTREG&("ECUECON")), 12)
   VDEF$ = X1$ + X2$
   '
20 OBX$ = OBJPLA$("DESHASCUECON", VDEF$)
   VDEF$ = OBX$
   If OBX$ <> "" Then
     DECUE$ = Left$(OBX$, 12)
     If CUEINT%(DECUE$) < 1 Or CUEINT%(DECUE$) > NUCUEN% Then GoTo 20
     HACUE$ = Mid$(OBX$, 13, 12)
     If CUEINT%(HACUE$) < 1 Or CUEINT%(HACUE$) > NUCUEN% Then GoTo 20
     If HACUE$ < DECUE$ Then GoTo 20
     '
     Screen.MousePointer = 11
     If YACAIN% < 1 Then
       Call CARTAMESES
       Call CARMAINDI
       Call CARSALMEN
       YACAIN% = 1
     End If
     '
     Screen.MousePointer = 11
     JJ& = 0
     FIN& = ULTREG&("ECUECON")
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("ECUECON", U&)
       COCUE$ = Left$(X$, 12)
       If COCUE$ >= DECUE$ Then
         If COCUE$ <= HACUE$ Then
           CUEI% = CUEINT%(COCUE$)
           If CUEI% > 0 Then
             If CUEI% <= NUCUEN% Then
               '
               For KKI% = 0 To 12   ' descarta cuentas s/movimiento
                 If Abs(SALCUEN#(CUEI%, KKI%)) > 0.005 Then GoTo 30
               Next KKI%
               GoTo 80
               '
30             Text1.TEXT = COCUE$
               '
               TTXX$ = "   ": GoSub 90
               TTXX$ = "   ": GoSub 90
               TTXX$ = "Cuenta: " + TRIM$(Text1): GoSub 90
               TTXX$ = "*******": GoSub 90
               TTXX$ = "   ": GoSub 90
               TTXX$ = Label12: GoSub 90
               TTXX$ = String$(120, "-"): GoSub 90
               For KKI% = 0 To 12
                 TTXX$ = List1.List(KKI%): GoSub 90
               Next KKI%
               TTXX$ = String$(120, "-"): GoSub 90
               TTXX$ = Label13: GoSub 90
               TTXX$ = String$(120, "="): GoSub 90
             End If
           End If
         End If
       End If
80   Next U&
     '
     Call SETULTREG("LIANAREI", JJ&)
     Call CIERRARCH("LIANAREI")
     Call FINAL("*LIANAREI")
     Screen.MousePointer = 1
     '
   End If
   Text1.TEXT = TTT1$
   Exit Sub
   '
90 JJ& = JJ& + 1
   Call GRAREG("LIANAREI", TTXX$, JJ&)
   Return
   '
End Sub

Sub LIRESREI()
   '
   Dim AJUREI#(13)
   '
   MODOX% = COMALTER%("Elija Formato de Salida\\REI Acumulado\Tabulado Mensual")
   On MODOX% GoTo 10, 60
   GoTo 99
   '
10 TTT1$ = Text1.TEXT
   X1$ = Left$(REGLEIDO$("ECUECON", 1), 12)
   X2$ = Left$(REGLEIDO$("ECUECON", ULTREG&("ECUECON")), 12)
   VDEF$ = X1$ + X2$
   '
20 OBX$ = OBJPLA$("DESHASCUECON", VDEF$)
   VDEF$ = OBX$
   If OBX$ <> "" Then
     DECUE$ = Left$(OBX$, 12)
     If CUEINT%(DECUE$) < 1 Or CUEINT%(DECUE$) > NUCUEN% Then GoTo 20
     HACUE$ = Mid$(OBX$, 13, 12)
     If CUEINT%(HACUE$) < 1 Or CUEINT%(HACUE$) > NUCUEN% Then GoTo 20
     If HACUE$ < DECUE$ Then GoTo 20
     '
     Screen.MousePointer = 11
     If YACAIN% < 1 Then
       Call CARTAMESES
       Call CARMAINDI
       Call CARSALMEN
       YACAIN% = 1
     End If
     '
     Screen.MousePointer = 11
     JJ& = 0
     FIN& = ULTREG&("ECUECON")
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("ECUECON", U&)
       COCUE$ = Left$(X$, 12)
       If COCUE$ >= DECUE$ Then
         If COCUE$ <= HACUE$ Then
           CUEI% = CUEINT%(COCUE$)
           If CUEI% > 0 Then
             If CUEI% <= NUCUEN% Then
               '
               For KKI% = 0 To 12   ' descarta cuentas s/movimiento
                 If Abs(SALCUEN#(CUEI%, KKI%)) > 0.005 Then GoTo 30
               Next KKI%
               GoTo 40
               '
30             Text1.TEXT = COCUE$
               TTXX$ = Label13
               XX$ = REGBLAN$("LIRESREI")
               Call REPLA(XX$, Text1.TEXT, 1, 12)
               Call REPLA(XX$, Label2.Caption, 13, 28)
               Call REPLA(XX$, MKD$(Val(Mid$(TTXX$, 39, 12))), 41, 8)
               Call REPLA(XX$, MKD$(Val(Mid$(TTXX$, 61, 12))), 49, 8)
               Call REPLA(XX$, MKD$(Val(Mid$(TTXX$, 83, 12))), 57, 8)
               JJ& = JJ& + 1
               Call GRAREG("LIRESREI", XX$, JJ&)
               '
             End If
           End If
         End If
       End If
40   Next U&
     '
     Call SETULTREG("LIRESREI", JJ&)
     Call CIERRARCH("LIRESREI")
     Call FINAL("*LIRESREI")
     Screen.MousePointer = 1
     '
   End If
   Text1.TEXT = TTT1$
   GoTo 99
   '
   '
60 TTT1$ = Text1.TEXT
   X1$ = Left$(REGLEIDO$("ECUECON", 1), 12)
   X2$ = Left$(REGLEIDO$("ECUECON", ULTREG&("ECUECON")), 12)
   VDEF$ = X1$ + X2$
   '
70 OBX$ = OBJPLA$("DESHASCUECON", VDEF$)
   VDEF$ = OBX$
   If OBX$ <> "" Then
     DECUE$ = Left$(OBX$, 12)
     If CUEINT%(DECUE$) < 1 Or CUEINT%(DECUE$) > NUCUEN% Then GoTo 20
     HACUE$ = Mid$(OBX$, 13, 12)
     If CUEINT%(HACUE$) < 1 Or CUEINT%(HACUE$) > NUCUEN% Then GoTo 20
     If HACUE$ < DECUE$ Then GoTo 70
     '
     Screen.MousePointer = 11
     If YACAIN% < 1 Then
       Call CARTAMESES
       Call CARMAINDI
       Call CARSALMEN
       YACAIN% = 1
     End If
     '
     Screen.MousePointer = 11
     JJ& = 0
     FIN& = ULTREG&("ECUECON")
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("ECUECON", U&)
       COCUE$ = Left$(X$, 12)
       If COCUE$ >= DECUE$ Then
         If COCUE$ <= HACUE$ Then
           CUEI% = CUEINT%(COCUE$)
           If CUEI% > 0 Then
             If CUEI% <= NUCUEN% Then
               '
               For KKI% = 0 To 12   ' descarta cuentas s/movimiento
                 If Abs(SALCUEN#(CUEI%, KKI%)) > 0.005 Then GoTo 80
               Next KKI%
               GoTo 90
               '
80             Text1.TEXT = COCUE$
               XX$ = REGBLAN$("LITABREI")
               Call REPLA(XX$, Text1.TEXT, 1, 12)
               Call REPLA(XX$, Label2.Caption, 13, 28)
               '
               HAYAJU% = 0
               For KKI% = 0 To 12
                 TTXX$ = List1.List(KKI%)
                 AJUREI#(KKI%) = Val(Mid$(TTXX$, 61, 12))
                 If Abs(AJUREI#(KKI%)) > 0.005 Then HAYAJU% = 1
               Next KKI%
               '
               If HAYAJU% > 0 Then
                 SUMAREI# = 0
                 For KKI% = 0 To 12
                   Call REPLA(XX$, MKD$(AJUREI#(KKI%)), 41 + 8 * KKI%, 8)
                   SUMAREI# = SUMAREI# + AJUREI#(KKI%)
                 Next KKI%
                 Call REPLA(XX$, MKD$(SUMAREI#), 41 + 13 * 8, 8)
                 JJ& = JJ& + 1
                 Call GRAREG("LITABREI", XX$, JJ&)
               End If
               '
             End If
           End If
         End If
       End If
90   Next U&
     '
     Call SETULTREG("LITABREI", JJ&)
     Call CIERRARCH("LITABREI")
     Call FINAL("*LITABREI")
     Screen.MousePointer = 1
     '
   End If
   Text1.TEXT = TTT1$
   GoTo 99
   '
   '
99 End Sub

Sub ASIENREI()
   '
   Dim PASREI#(16384)
   '
   Screen.MousePointer = 11
   If YACAIN% < 1 Then
     Call CARTAMESES
     Call CARMAINDI
     Call CARSALMEN
     YACAIN% = 1
   End If
   '
   Screen.MousePointer = 11
   TTT1$ = Text1.TEXT
   JJ& = 0
   FIN& = ULTREG&("ECUECON")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("ECUECON", U&)
     COCUE$ = Left$(X$, 12)
     CUEI% = CUEINT%(COCUE$)
     If CUEI% > 0 Then
       If CUEI% <= NUCUEN% Then
         '
         For KKI% = 0 To 12   ' descarta cuentas s/movimiento
           If Abs(SALCUEN#(CUEI%, KKI%)) > 0.005 Then GoTo 30
         Next KKI%
         GoTo 80
         '
30       Text1.TEXT = COCUE$
         TTXX$ = Label13
         VAREI# = Val(Mid$(TTXX$, 61, 12))
         VADIC# = Val(Mid$(TTXX$, 83, 12))
         '
         If Abs(VAREI#) > 0.005 Or Abs(VADIC#) > 0.005 Then
           '
           CUECOREI$ = Label9(0)
           CUEIREI% = CUEINT%(CUECOREI$)
           If CUEIREI% > NUCUEN% Then CUEIREI% = 0
           If CUEIREI% < 1 Then
             If Abs(VAREI#) > 0.005 Then
               Call MENSERR(24, "Falta Cuenta de Contrapartida\para R.E.I en Cuenta " + COCUE$)
               GoTo 80
             End If
           End If
           '
           CUECODIC$ = Label9(1)
           CUEIDIC% = CUEINT%(CUECODIC$)
           If CUEIDIC% > NUCUEN% Then CUEIDIC% = 0
           If CUEIDIC% < 1 Then
             If Abs(VADIC#) > 0.005 Then
               Call MENSERR(24, "Falta Cuenta de Contrapartida\para Diferencia de Cambio en Cuenta " + COCUE$)
               GoTo 80
             End If
           End If
           '
           PASREI#(CUEI%) = PASREI#(CUEI%) + VAREI# + VADIC#
           PASREI#(CUEIREI%) = PASREI#(CUEIREI%) - VAREI#
           PASREI#(CUEIDIC%) = PASREI#(CUEIDIC%) - VADIC#
           '
         End If
       End If
     End If
80 Next U&
   '
   JJ& = 0
   FIN& = ULTREG&("ECUECON")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("ECUECON", U&)
     COCUE$ = Left$(X$, 12)
     CUEI% = CUEINT%(COCUE$)
     If CUEI% > 0 Then
       If CUEI% <= NUCUEN% Then
         If Abs(PASREI#(CUEI%)) > 0.005 Then
           XX$ = REGBLAN$("ASIENREI")
           Call REPLA(XX$, COCUE$, 1, 12)
           Call REPLA(XX$, DENOCUE$(CUEI%), 13, 28)
           If PASREI#(CUEI%) > 0 Then
               Call REPLA(XX$, MKD$(PASREI#(CUEI%)), 41, 8)
             Else
               Call REPLA(XX$, MKD$(Abs(PASREI#(CUEI%))), 49, 8)
           End If
           JJ& = JJ& + 1
           Call GRAREG("ASIENREI", XX$, JJ&)
         End If
       End If
     End If
   Next U&
   '
   Call SETULTREG("ASIENREI", JJ&)
   Call CIERRARCH("ASIENREI")
   Call FINAL("*ASIENREI")
   Text1.TEXT = TTT1$
   Screen.MousePointer = 1
   '
End Sub

Sub GRABAEXEL()
    ' Declara variables de objeto para Microsoft
    ' Excel, el libro de trabajo de la aplicación y
    ' los objetos de hoja de cálculo.
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet

    ' Asigna referencias de objeto a las variables.
    ' Utiliza métodos Add para crear nuevos objetos de
    ' libro de trabajo y hoja de cálculo.
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets.Add
    
    ' Asigna los valores introducidos en los cuadros
    ' de texto a las celdas de Microsoft Excel.
    xlSheet.Cells(1, 1).Value = "1"
    xlSheet.Cells(2, 1).Value = "2"
    
    ' Guarda la hoja de cálculo.
    xlSheet.SaveAs "c:\Temp.xls"
    
    ' Cierra Microsoft Excel con el método Quit.
    xlApp.Quit
    
    ' Libera los objetos.
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
End Sub
