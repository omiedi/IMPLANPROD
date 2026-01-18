VERSION 5.00
Begin VB.Form ORPAGO04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Pago (Operaciones Comerciales de Compras)"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFD0&
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
      Height          =   8500
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00404000&
         Height          =   9000
         Left            =   10950
         ScaleHeight     =   8940
         ScaleWidth      =   915
         TabIndex        =   46
         Top             =   0
         Width           =   975
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
            Height          =   690
            Left            =   140
            Picture         =   "ORPAGO04.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Terminar - Salir de la Aplicación"
            Top             =   200
            Width           =   650
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   140
            Picture         =   "ORPAGO04.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Anulación de Orden de Pago Emitida"
            Top             =   5220
            Width           =   650
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00404000&
            Height          =   8310
            Left            =   10950
            ScaleHeight     =   8250
            ScaleWidth      =   915
            TabIndex        =   57
            Top             =   0
            Width           =   975
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
               Height          =   640
               Left            =   140
               Picture         =   "ORPAGO04.frx":0454
               Style           =   1  'Graphical
               TabIndex        =   68
               ToolTipText     =   "Cambia Ordenamiento de Valores (Fecha / Importe)"
               Top             =   3360
               Width           =   650
            End
            Begin VB.PictureBox Picture5 
               BackColor       =   &H000000FF&
               Height          =   135
               Left            =   140
               ScaleHeight     =   75
               ScaleWidth      =   585
               TabIndex        =   66
               Top             =   6720
               Width           =   650
               Begin VB.Frame Frame7 
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   0  'None
                  Height          =   500
                  Left            =   0
                  TabIndex        =   67
                  Top             =   0
                  Width           =   12000
               End
            End
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
               Height          =   640
               Left            =   140
               Picture         =   "ORPAGO04.frx":075E
               Style           =   1  'Graphical
               TabIndex        =   65
               ToolTipText     =   "Configuración de Funcionamiento"
               Top             =   5775
               Width           =   650
            End
            Begin VB.CommandButton Command17 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   640
               Left            =   140
               Picture         =   "ORPAGO04.frx":0BA0
               Style           =   1  'Graphical
               TabIndex        =   64
               ToolTipText     =   "Anular, Re-Imprimir, Imputación de Créditos, Tablas"
               Top             =   5145
               Width           =   650
            End
            Begin VB.CommandButton Command16 
               Caption         =   "Imprimir"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   960
               Left            =   140
               Picture         =   "ORPAGO04.frx":0CEA
               Style           =   1  'Graphical
               TabIndex        =   63
               ToolTipText     =   "Aprueba - Graba e Imprime Orden de Pago"
               Top             =   7140
               Width           =   650
            End
            Begin VB.CommandButton Command15 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   640
               Left            =   140
               Picture         =   "ORPAGO04.frx":1354
               Style           =   1  'Graphical
               TabIndex        =   62
               ToolTipText     =   "Archivo Maestro de Proveedores"
               Top             =   1680
               Width           =   650
            End
            Begin VB.CommandButton Command14 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   640
               Left            =   140
               Picture         =   "ORPAGO04.frx":165E
               Style           =   1  'Graphical
               TabIndex        =   61
               ToolTipText     =   "Resumen de Cuenta Corriente"
               Top             =   2310
               Width           =   650
            End
            Begin VB.CommandButton Command13 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   640
               Left            =   140
               Picture         =   "ORPAGO04.frx":1968
               Style           =   1  'Graphical
               TabIndex        =   60
               ToolTipText     =   "Notas de Proveedores"
               Top             =   3990
               Width           =   650
            End
            Begin VB.CommandButton Command12 
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
               Left            =   105
               Picture         =   "ORPAGO04.frx":1C72
               TabIndex        =   59
               ToolTipText     =   "Abandona la Aplicación"
               Top             =   105
               Width           =   650
            End
            Begin VB.CommandButton Command10 
               Caption         =   "Next"
               Enabled         =   0   'False
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
               Left            =   140
               Picture         =   "ORPAGO04.frx":1F7C
               Style           =   1  'Graphical
               TabIndex        =   58
               ToolTipText     =   "Carga de Valores de Pago Propios"
               Top             =   735
               Width           =   650
            End
         End
         Begin VB.CommandButton BOTNEXT 
            Caption         =   "Next"
            Enabled         =   0   'False
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
            Left            =   140
            Picture         =   "ORPAGO04.frx":2286
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Carga de Valores de Pago Propios"
            Top             =   890
            Width           =   650
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Note"
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
            Left            =   140
            Picture         =   "ORPAGO04.frx":2590
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Notas de Proveedores"
            Top             =   3080
            Width           =   650
         End
         Begin VB.CommandButton Command4 
            Caption         =   "R-Cue"
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
            Left            =   140
            Picture         =   "ORPAGO04.frx":289A
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   1700
            Width           =   650
         End
         Begin VB.CommandButton Command11 
            Caption         =   "A-B-M"
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
            Left            =   140
            Picture         =   "ORPAGO04.frx":2BA4
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Archivo Maestro de Proveedores"
            Top             =   3770
            Width           =   650
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Print"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   140
            Picture         =   "ORPAGO04.frx":2EAE
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Aprueba - Graba e Imprime Orden de Pago"
            Top             =   7140
            Width           =   650
         End
         Begin VB.CommandButton Command6 
            Caption         =   "More"
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
            Left            =   140
            Picture         =   "ORPAGO04.frx":3518
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Anular, Re-Imprimir, Imputación de Créditos, Tablas"
            Top             =   2390
            Width           =   650
         End
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
            Left            =   140
            Picture         =   "ORPAGO04.frx":3662
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   6000
            Width           =   650
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   140
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   48
            Top             =   6825
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   49
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Sort"
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
            Left            =   140
            Picture         =   "ORPAGO04.frx":3AA4
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Cambia Ordenamiento de Valores (Fecha / Importe)"
            Top             =   4530
            Width           =   650
         End
      End
      Begin VB.ListBox LINDICE 
         Height          =   1035
         Left            =   5355
         Sorted          =   -1  'True
         TabIndex        =   45
         Top             =   2520
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   6930
         TabIndex        =   44
         Top             =   2835
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0FFD0&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   540
         Left            =   6825
         TabIndex        =   41
         Top             =   30
         Width           =   3900
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Fondo Fijo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   540
            Left            =   1155
            TabIndex        =   43
            Top             =   0
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Orden de Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   540
            Left            =   2520
            TabIndex        =   42
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   2640
         Left            =   210
         ScaleHeight     =   2580
         ScaleWidth      =   10395
         TabIndex        =   25
         Top             =   1365
         Width           =   10450
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   10515
            Begin VB.Line Line23 
               X1              =   8505
               X2              =   8505
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   " M-P    Referencia                  Numero               Emision        Vencim.           Importe            Selección"
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
               TabIndex        =   27
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line20 
               X1              =   5670
               X2              =   5670
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line13 
               X1              =   4410
               X2              =   4410
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line7 
               X1              =   2835
               X2              =   2835
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line6 
               X1              =   630
               X2              =   630
               Y1              =   0
               Y2              =   1155
            End
            Begin VB.Line Line5 
               X1              =   6930
               X2              =   6930
               Y1              =   0
               Y2              =   1050
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5580
            Left            =   -315
            TabIndex        =   28
            Top             =   -105
            Width           =   10800
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
               Height          =   2310
               Left            =   290
               TabIndex        =   29
               Top             =   440
               Width           =   10500
            End
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   3120
         Left            =   210
         ScaleHeight     =   3060
         ScaleWidth      =   10395
         TabIndex        =   15
         Top             =   4150
         Width           =   10450
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5630
            Left            =   0
            TabIndex        =   16
            Top             =   315
            Width           =   10515
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
               Height          =   2310
               Left            =   -40
               TabIndex        =   17
               Top             =   0
               Width           =   10450
            End
            Begin VB.ListBox List3 
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
               Left            =   -30
               TabIndex        =   38
               Top             =   0
               Width           =   10450
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
               TabIndex        =   24
               Top             =   2415
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
               TabIndex        =   23
               Top             =   0
               Width           =   1695
            End
            Begin VB.Line Line4 
               X1              =   8400
               X2              =   8400
               Y1              =   210
               Y2              =   2730
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
               Left            =   6510
               TabIndex        =   20
               Top             =   2415
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
               TabIndex        =   19
               Top             =   2415
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Total General"
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
               TabIndex        =   18
               Top             =   2415
               Width           =   1695
            End
            Begin VB.Line Line14 
               BorderWidth     =   2
               X1              =   0
               X2              =   10395
               Y1              =   2310
               Y2              =   2310
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   10395
               Y1              =   2730
               Y2              =   2730
            End
            Begin VB.Line Line16 
               X1              =   2520
               X2              =   2520
               Y1              =   2310
               Y2              =   2730
            End
            Begin VB.Line Line17 
               X1              =   3735
               X2              =   3735
               Y1              =   2310
               Y2              =   2730
            End
            Begin VB.Line Line18 
               X1              =   5460
               X2              =   5460
               Y1              =   2310
               Y2              =   2730
            End
            Begin VB.Line Line19 
               X1              =   6720
               X2              =   6720
               Y1              =   2310
               Y2              =   2730
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   10515
            Begin VB.Line Line1 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line2 
               X1              =   6720
               X2              =   6720
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   " Documento  Número         Emision                Importe       Vencim.                  Saldo                Acredit.       "
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
               TabIndex        =   22
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line9 
               X1              =   2520
               X2              =   2520
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line10 
               X1              =   3730
               X2              =   3730
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line11 
               X1              =   5460
               X2              =   5460
               Y1              =   0
               Y2              =   1050
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
         Left            =   6195
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   210
         Width           =   435
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
         Top             =   920
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
         Text            =   " "
         Top             =   560
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
         Top             =   560
         Width           =   180
      End
      Begin VB.TextBox Text5 
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
         Left            =   2310
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   920
         Width           =   2850
      End
      Begin VB.TextBox Text4 
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
         Left            =   1470
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   920
         Width           =   750
      End
      Begin VB.TextBox Text3 
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
         Left            =   1470
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   560
         Width           =   3690
      End
      Begin VB.TextBox Text2 
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
         Left            =   1470
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   210
         Width           =   3690
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
         Left            =   1100
         TabIndex        =   1
         Top             =   560
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
         Left            =   240
         TabIndex        =   0
         Text            =   " "
         Top             =   560
         Width           =   855
      End
      Begin VB.Label Label29 
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
         Height          =   330
         Left            =   6300
         TabIndex        =   77
         Top             =   7245
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label LabeL28 
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
         Height          =   330
         Left            =   4725
         TabIndex        =   76
         Top             =   7245
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label25 
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
         Height          =   330
         Left            =   3570
         TabIndex        =   75
         Top             =   7245
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label23 
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
         Height          =   330
         Left            =   2520
         TabIndex        =   74
         Top             =   7245
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label27 
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
         Height          =   330
         Left            =   1470
         TabIndex        =   73
         Top             =   7245
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   72
         Top             =   300
         Width           =   960
      End
      Begin VB.Label Label26 
         BackColor       =   &H00EEEEEE&
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
         Left            =   240
         TabIndex        =   71
         Top             =   915
         Width           =   1035
      End
      Begin VB.Line Line22 
         X1              =   0
         X2              =   10920
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line21 
         X1              =   210
         X2              =   10660
         Y1              =   4010
         Y2              =   4010
      End
      Begin VB.Line Line12 
         X1              =   10670
         X2              =   10665
         Y1              =   4150
         Y2              =   7250
      End
      Begin VB.Line Line8 
         X1              =   10665
         X2              =   10665
         Y1              =   1365
         Y2              =   3990
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valores:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4305
         TabIndex        =   40
         Top             =   7350
         Width           =   2220
      End
      Begin VB.Label Label21 
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
         Height          =   330
         Left            =   5985
         TabIndex        =   39
         Top             =   7590
         Width           =   1695
      End
      Begin VB.Label Label20 
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
         Height          =   330
         Left            =   8610
         TabIndex        =   37
         Top             =   7590
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6825
         TabIndex        =   36
         Top             =   7350
         Width           =   2220
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
         Height          =   330
         Left            =   630
         TabIndex        =   35
         Top             =   7590
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Retenciones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2205
         TabIndex        =   34
         Top             =   7350
         Width           =   2220
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
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   5880
         Width           =   2220
      End
      Begin VB.Label Label14 
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
         Height          =   330
         Left            =   3360
         TabIndex        =   31
         ToolTipText     =   "Doble Click para Editar Valores"
         Top             =   7590
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Creditos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -945
         TabIndex        =   30
         Top             =   7350
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Documento no Válido como Factura"
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
         Left            =   5355
         TabIndex        =   14
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   840
         Width           =   2115
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
         Height          =   330
         Left            =   7560
         TabIndex        =   12
         Top             =   920
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
         TabIndex        =   11
         Top             =   630
         Width           =   750
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
      TabIndex        =   10
      Top             =   0
      Width           =   960
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu NueOPs 
         Caption         =   "Nueva Orden de Pago"
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu MaeProves 
         Caption         =   "Maestro de Proveedores"
      End
      Begin VB.Menu NotProve 
         Caption         =   "Notas de Proveedores"
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSali 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuedi 
      Caption         =   "Edición"
      Begin VB.Menu AnuOps 
         Caption         =   "Anulación de Orden de Pago"
      End
      Begin VB.Menu veryreimp 
         Caption         =   "Ver y Re-Imprimir O/P"
      End
      Begin VB.Menu php1a 
         Caption         =   "-"
      End
      Begin VB.Menu impucredi 
         Caption         =   "Imputación de Créditos"
      End
      Begin VB.Menu AjuSaldAc 
         Caption         =   "Ajuste de Saldos Acreedores"
      End
      Begin VB.Menu php1b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTaMePa 
         Caption         =   "Tabla de Bancos y Valores de Pago"
      End
   End
   Begin VB.Menu ConsuLis 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuReCuCor 
         Caption         =   "Resumen de Cuenta Corriente"
      End
      Begin VB.Menu mnuLiCuCor 
         Caption         =   "Listado de Cuenta Corriente"
      End
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu ListOP 
         Caption         =   "Listado de Ordenes de Pago"
      End
   End
   Begin VB.Menu mnuconfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuConfiGral 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu mnuAccDir 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuMaeProv 
         Caption         =   "Maestro de Proveedores"
      End
      Begin VB.Menu php2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOtrAccDir 
         Caption         =   "Otros Accesos Directos"
      End
   End
End
Attribute VB_Name = "ORPAGO04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$
Dim MODOSORVAL%
Dim YACARFO%



Private Sub AjuSaldAc_Click()
    Call OPORPA04.Command4_Click
End Sub

Private Sub AnuOps_Click()
   Call OPORPA04.Command25_Click
End Sub

Private Sub BOTNEXT_Click()
    '
    If Val(Label18.Caption) < 0.005 Then
       Exit Sub
    End If
    '
    If Val(Label20.Caption) < 0 Then
         Call MENSERR(24, "Valores Cargados exceden Créditos")
         Exit Sub
       ElseIf Val(Label20.Caption) >= 0.005 Then
         Call CIERRARCH("!CARMEPA")
         Call BLANARCH("!CARMEPA")
         X$ = REGBLAN$("CARMEPA")
         HOI% = HOY(HO$)
         Call REPLA(X$, MKD$(Val(Label20.Caption)), 31, 8)
         Call REPLA(X$, MKI$(HOI%), 39, 2)
         Call REPLA(X$, MKI$(HOI%), 41, 2)
         Call REGAPP("!CARMEPA", X$)
         '
10       VTB% = VENTABU%("CARMEPA")
         If VTB% < 0 Then
            Call CIERRARCH("!CARMEPA")
            Call BLANARCH("!CARMEPA")
            Exit Sub
         End If
         SUVAPROP# = 0
         For U& = 1 To ULTREG&("!CARMEPA")
            X$ = REGLEIDO$("!CARMEPA", U&)
            IVAL# = CVD(Mid$(X$, 31, 8))
            If IVAL# < 0.005 Then
               Call MENSERR(24, "No se Admiten Valores Negativos")
               GoTo 10
            End If
            If CVI(Left$(X$, 2)) < 1 Or ECOARCH$("MEPAGO", Left$(X$, 2)) = "" Then
               Call MENSERR(24, "Tipo de Valor no Admitido")
               GoTo 10
            End If
            SUVAPROP# = SUVAPROP# + IVAL#
         Next U&
         If Abs(SUVAPROP# - Val(Label20.Caption)) >= 0.005 Then
            Call MENSERR(24, "Imposible Cerrar - Diferencia de Valores")
            GoTo 10
         End If
         '
         Call CIERRARCH("!CARMEPA")
         Call SIGUEORPA4
         '
       Else
         '
         Call CIERRARCH("!CARMEPA")
         Call BLANARCH("!CARMEPA")
         Call SIGUEORPA4
         '
    End If
    '
End Sub
'

Private Sub Command1_Click()
    '
    Call SELECHO("PROVED1")
    NCXI% = Val(TRIM$(VALACT1$("PROVED1")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$((-1) * NC%)
        Text3.TEXT = DOMICLI$((-1) * NC%)
        Text4.TEXT = CPOSCLI$((-1) * NC%)
        Text5.TEXT = LOCACLI$((-1) * NC%)
        Call SIGUEORPA1
    End If
    '
End Sub


Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("ARCHIG04/AMAPRO", "Archivos Maestros")
    Command11.Enabled = True
End Sub

Private Sub Command2_Click()
    Screen.MousePointer = 11
    If List1.ListCount > 0 Or List2.ListCount > 0 Then
          Call CIERRARCH("*.*")
          Call BLANARCH("!CARGAST")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label18.Caption = ""
          Label20.Caption = ""
          Label21.Caption = ""
          BOTNEXT.Enabled = False
          Command5.Enabled = False
          Screen.MousePointer = 1
          EPAC$ = UCase$(EMPREAC$)
          If InStr(EPAC$, "NEUMAN") > 0 Then
            Text1.TEXT = ""
            Text2.TEXT = ""
            Text3.TEXT = ""
            Text4.TEXT = ""
            Text5.TEXT = ""
          End If
          Text1.SetFocus
        Else
          If TRIM$(CONTROL$("ORPASTA", "AVAHOJA")) = "NEWPAGE" Then
              Call SETSPOOL("ENDDOC", "")
          End If
          Hide
          Screen.MousePointer = 1
    End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If List1.ListCount > 0 Or List2.ListCount > 0 Then
        Command2.ToolTipText = "Limpia y Refresca Pantalla"
      Else
        Command2.ToolTipText = "Cierra y Abandona la Aplicación"
   End If
End Sub

Private Sub Command25_Click()
    Call OPORPA04.Command25_Click
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Text1.TEXT)
    Call CONECRUN("*NOTAPRO/" + NCLI$, "Notas de Proveedores")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    'If EXISTE%("COGECOM.EXE") < 1 Then
    '    Call OPNOIN("")
    '    Command4.Enabled = True
    '    Exit Sub
    'End If
    RECUEP04.Text1 = Text1
    'PROLLAMA$ = Name
    'Hide
    RECUEP04.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    TOTAVAL# = CDbl(Val(Label14.Caption) + Val(Label21.Caption))
    TOTACAN# = CDbl(Val(Label18.Caption))
    If Abs(TOTACAN# - TOTAVAL#) >= 0.005 Then
        Call COMUNI("Operación Incorrecta.\  \Las Cancelaciones Difieren\de la Suma de Valores a Entregar.")
        Command5.Enabled = True
        Exit Sub
    End If
    '
    VACON% = 1       ' PARA QUE NO BLANQUEE LISTADOS
    Call GRAORPASTA(OKX%)
    If OKX% < 0 Then GoTo 10   ' abortó en observaciones
    If OKX% = 0 Then           ' RE-NUMERO ORDEN DE PAGO
       Command5.Enabled = True
       Exit Sub
    End If
    '
    If OKX% <> 7 Then
        Call COMUNI("La Registración de esta Orden de Pago\no Pudo Completarse. (" + TRIM$(Str$(OKX%)) + ")\  \Consulte a su Soporte de Sistemas FLEXOFT.")
        Call FINAL("")
    End If
    Call PRIORPASTA
    Call PRICORETIB
    VACON% = 0
    '
10  FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    NROPRO& = 0
    Call NUORPA(NROPRO&)
    Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
    '
    Call CIERRARCH("!CARMEPA")
    Call BLANARCH("!CARMEPA")
    Call CIERRARCH("!CARGAPLI")
    Call BLANARCH("!CARGAPLI")
    Call CIERRARCH("!CARGAST")
    Call BLANARCH("!CARGAST")
    List1.Clear
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label18.Caption = ""
    Label20.Caption = ""
    Label21.Caption = ""
    'BOTNEXT.Enabled = True
    Command5.Enabled = False
    EPAC$ = UCase$(EMPREAC$)
    If InStr(EPAC$, "NEUMAN") > 0 Then
      Text1.TEXT = ""
      Text2.TEXT = ""
      Text3.TEXT = ""
      Text4.TEXT = ""
      Text5.TEXT = ""
    End If
    Text1.SetFocus
    '
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    OPORPA04.Show 1
    '
    NROPRO& = 0
    Call NUORPA(NROPRO&)
    Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
    '
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    STORPA04.Show 1
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   If List1.ListCount > 0 Then
     Screen.MousePointer = 11
     MODOSORVAL% = (MODOSORVAL% + 1) Mod 2
     List4.Clear
     '
     For U& = 1 To List1.ListCount
       Z1$ = List1.List(U& - 1)
       VALI1 = FECHANUM(Mid$(Z1$, 49, 8))
       If MODOSORVAL% = 1 Then VALI1 = Val(Mid$(Z1$, 57, 12))
       For U4& = 1 To List4.ListCount
         Z4$ = List4.List(U4& - 1)
         VALI4 = FECHANUM(Mid$(Z4$, 49, 8))
         If MODOSORVAL% = 1 Then VALI4 = Val(Mid$(Z4$, 57, 12))
         If VALI4 > VALI1 Then
           List4.AddItem Z1$, U4& - 1
           GoTo 19
         End If
       Next U4&
       List4.AddItem Z1$
19   Next U&
     '
     List1.Clear
     For U& = 1 To List4.ListCount
       List1.AddItem List4.List(U& - 1)
     Next U&
     List1.Refresh
     Screen.MousePointer = 1
     '
   End If
   Command8.Enabled = True
   '
End Sub

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub Form_Activate()
   '
   If YACARFO% < 1 Then
      '
      DoEvents
      YACARFO% = 1
      '
10    FORIPRE$ = TRIM$(CONTROL$("", "FORPASTA"))
      If FORIPRE$ = "" Then
          Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Orden de Pago.")
          STORPA04.Show 1
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label18.Caption = ""
          Label20.Caption = ""
          Command5.Enabled = False
          GoTo 10
      End If
      '
      If FORIPRE$ <> "" Then
          LU$ = LUDAT$
          If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
              RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
              Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Pago Standard:\  \'" + RMCC$ + "'.")
              STORPA04.Show 1
              GoTo 10
          End If
      End If
      '
      FF% = HOY(HO$)
      Text9.TEXT = HO$
      '
      NROPRO& = 0
      Call NUORPA(NROPRO&)
      Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
      '
   End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call CENTRAFORM(Me)
    Refresh
    DoEvents
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If TRIM$(CONTROL$("ORPASTA", "AVAHOJA")) = "NEWPAGE" Then
       Call SETSPOOL("ENDDOC", "")
    End If
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    YACARFO% = 0
End Sub

Private Sub impucredi_Click()
    Call OPORPA04.Command3_Click
End Sub

Private Sub Label10_Change()
   Call CALRETIBRU
   If Val(Label10.Caption) >= 0.005 Then
        BOTNEXT.Enabled = True
      Else
        BOTNEXT.Enabled = False
   End If
End Sub

Private Sub Label14_Change()
   Label20.Caption = CSTRING$(MKD$(Val(Label18.Caption) - Val(Label14.Caption) - Val(Label21.Caption)), 4, 14, 2, 2)
   Label20.Refresh
End Sub

Private Sub Label14_DblClick()
   If Val(Label18) > 0 Then
     VDEF$ = MKD$(Val(Label23)) + MKS$(Val(Label29)) + MKS$(Val(Label27))
10   BASIMPON# = CVD(Left$(VDEF$, 8))
     IRET# = 0
     If BASIMPON# >= Val(LabeL28) Then
       IRET# = CVD(Left$(VDEF$, 8)) * CVS(Mid$(VDEF$, 9, 4)) / 100 * CVS(Mid$(VDEF$, 13, 4)) / 100
     End If
     VDEF$ = VDEF$ + MKD$(IRET#)
     OBX$ = OBJPLA$("CARIRETIB", VDEF$)
     If OBX$ = "" Then Exit Sub
     IRET1# = 0
     If CVD(Left$(OBX$, 8)) >= Val(LabeL28) Then
       IRET1# = CVD(Left$(OBX$, 8)) * CVS(Mid$(OBX$, 9, 4)) / 100 * CVS(Mid$(OBX$, 13, 4)) / 100
     End If
     If Abs(IRET# - IRET1#) > 0.0005 Then
       VDEF$ = Left$(OBX$, 16)
       GoTo 10
     End If
     '
     Label23 = TRIM$(CSTRING$(Left$(OBX$, 8), 3, 12, 2, 2))
     Label29 = TRIM$(CSTRING$(Mid$(OBX$, 9, 4), 3, 12, 2, 2))
     Label27 = TRIM$(CSTRING$(Mid$(OBX$, 13, 4), 3, 12, 2, 2))
     Label14 = TRIM$(CSTRING$(Mid$(OBX$, 17, 8), 3, 12, 2, 2))
   End If
End Sub

Private Sub Label18_Change()
   If Val(Label18.Caption) > 0 Then
        BOTNEXT.Enabled = True
      Else
        BOTNEXT.Enabled = False
   End If
End Sub

Private Sub Label2_DblClick()
    STORPA04.Show 1
End Sub

Private Sub Label20_Change()
   If Val(Label20.Caption) > 0 Then
        Command5.Enabled = False
      Else
        Command5.Enabled = True
        BOTNEXT.Enabled = False
   End If

End Sub

Private Sub Label26_Click()
   Call SELECHO("INDIACRE")
   CPROX$ = TRIM$(VALACT1$("INDIACRE"))
   If CPROX$ <> "" Then
     NCXI% = NUMEPRO%(CPROX$)
     If NCXI% > 0 Then
       If REGICLI&((-1) * NCXI%) > 0 Then
         Label26 = CPROX$
         NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NC%))
         Text2.TEXT = RASOCLI$((-1) * NC%)
         Text3.TEXT = DOMICLI$((-1) * NC%)
         Text4.TEXT = CPOSCLI$((-1) * NC%)
         Text5.TEXT = LOCACLI$((-1) * NC%)
         Call SIGUEORPA1
       End If
     End If
   End If
End Sub

Private Sub List1_Click()
    '
    If VACON% <> 0 Then
        Exit Sub
    End If
    '
    Z$ = List1.TEXT
    LIST1ANT$ = Z$
    ILISTA& = List1.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = "  " + Mid$(Z$, 57, 12)
    SUMAVAL# = Val(Label21.Caption)
    MAXAPLI# = Val(Label18.Caption) - Val(Label21.Caption) - Val(Label14.Caption)
    If IACRE$ = "" Then
         If Val(IPEN$) <= MAXAPLI# + 0.005 Then
             Call REPLA(Z$, IPEN$, 70, 14)
             SUMAVAL# = SUMAVAL# + CDbl(Val(IPEN$))
           Else
             Call MENSERR(24, "Importe Valor Excede Saldo a Cancelar")
         End If
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMAVAL# = SUMAVAL# - CDbl(Val(IACRE$))
    End If
    VACON% = 1
    List1.AddItem Z$, ILISTA&
    List1.RemoveItem ILISTA& + 1
    List1.Refresh
    List1.ListIndex = ILISTA&
    VACON% = 0
    '
    Label21.Caption = CSTRING$(MKD$(SUMAVAL#), 4, 14, 2, 2)
    Label21.Refresh
    Label20.Caption = CSTRING$(MKD$(Val(Label18.Caption) - Val(Label14.Caption) - SUMAVAL#), 4, 14, 2, 2)
    Label20.Refresh
    '
End Sub

Private Sub List1_DblClick()
    '
    'Call SIGUEORPA2
    '
End Sub

Private Sub List2_Click()
    '
    If VACON% <> 0 Then
        Exit Sub
    End If
    '
    If TICUEPRO%((-1) * NC%) = 1 Then   ' cuenta de gastos
       Call SIGUEORPA5
       Exit Sub
    End If
    '
    Z$ = List2.TEXT
    LIST2ANT$ = Z$
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = " " + Mid$(Z$, 56, 13)
    SUMACRE# = Val(Label18.Caption)
    If IACRE$ = "" Then
         Call REPLA(Z$, IPEN$, 70, 14)
         SUMACRE# = SUMACRE# + CDbl(Val(IPEN$))
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMACRE# = SUMACRE# - CDbl(Val(IACRE$))
         If Left$(Z$, 20) = "Anticipo no Aplicado" Then
            Call REPLA(Z$, Space$(13), 32, 13)
            Call REPLA(Z$, Space$(13), 56, 13)
         End If
    End If
    VACON% = 1
    List2.AddItem Z$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    List2.Refresh
    List2.ListIndex = ILISTA&
    VACON% = 0
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label10.Refresh
    Label18.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label18.Refresh
    Label20.Caption = CSTRING$(MKD$(Val(Label18.Caption) - Val(Label14.Caption) - Val(Label21.Caption)), 4, 14, 2, 2)
    Label20.Refresh
    '
End Sub
'

Private Sub List2_DblClick()
    '
    If TICUEPRO%((-1) * NC%) = 1 Then   ' cuenta de gastos
       Call SIGUEORPA5
       Exit Sub
    End If
    '
    VACON% = 1
    ILIST2& = List2.ListIndex
    TOLIST2& = List2.TopIndex
    '
    Z$ = List2.TEXT
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = Mid$(Z$, 56, 14)
    SUMACRE# = Val(Label18.Caption)
    SUMACRE# = SUMACRE# - CDbl(Val(IACRE$))
    SUMACRE# = SUMACRE# + CDbl(Val(Mid$(LIST2ANT$, 70, 14)))
    List2.Visible = False
    List2.AddItem LIST2ANT$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    '
    Call SIGUEORPA3
    '
    VACON% = 1
    List2.TopIndex = TOLIST2&
    List2.Visible = True
    If List2.ListCount > ILIST2& Then
          List2.ListIndex = ILIST2&
        Else
          Call CIERRARCH("!CARMEPA")
          Call BLANARCH("!CARMEPA")
          Call CIERRARCH("!CARGAPLI")
          Call BLANARCH("!CARGAPLI")
          Call CIERRARCH("!CARGAST")
          Call BLANARCH("!CARGAST")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label18.Caption = ""
          Label20.Caption = ""
          Label21.Caption = ""
          'BOTNEXT.Enabled = True
          Command5.Enabled = False
          EPAC$ = UCase$(EMPREAC$)
          If InStr(EPAC$, "NEUMAN") > 0 Then
            Text1.TEXT = ""
            Text2.TEXT = ""
            Text3.TEXT = ""
            Text4.TEXT = ""
            Text5.TEXT = ""
          End If
          Text1.SetFocus
    End If
    VACON% = 0
    '
End Sub

Private Sub ListOP_Click()
    Call OPORPA04.Command5_Click
End Sub

Private Sub MaeProves_Click()
    Call Command11_Click
End Sub

Private Sub mnuConfiGral_Click()
    Call Command7_Click
End Sub

Private Sub mnuLiCuCor_Click()
    Call RECUEP04.Command3_Click
End Sub

Private Sub mnuMaeProv_Click()
    Call Command11_Click
End Sub

Private Sub mnuOtrAccDir_Click()
    Call COMUNI("No Existen Accesos Directos \ para esta Aplicación")
End Sub

Private Sub mnuReCuCor_Click()
    Call Command4_Click
End Sub

Private Sub mnuSali_Click()
    Call Command2_Click
End Sub

Private Sub mnuTaMePa_Click()
    Call OPORPA04.Command1_Click
End Sub

Private Sub NotProve_Click()
    Call Command3_Click
End Sub

Private Sub NueOPs_Click()
    Call Command1_Click
End Sub

Private Sub Option2_Click()
    Hide
    FORLIFOF.Show
    Call FRESHALL
    Stop 'Call LIFOFIJ1
End Sub

Private Sub Text1_Change()
    Call CIERRARCH("!CARGAST")
    Call BLANARCH("!CARGAST")
    List1.Clear
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label18.Caption = ""
    Label20.Caption = ""
    Label21.Caption = ""
    'BOTNEXT.Enabled = True
    Command5.Enabled = False
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("PROVED1")
    NCXI% = Val(TRIM$(VALACT1$("PROVED1")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$((-1) * NC%)
        Text3.TEXT = DOMICLI$((-1) * NC%)
        Text4.TEXT = CPOSCLI$((-1) * NC%)
        Text5.TEXT = LOCACLI$((-1) * NC%)
        Call SIGUEORPA1
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
       If Val(Text1.TEXT) > 0 Then
          If Val(Text1.TEXT) <= 32767 Then
             NC% = Val(Text1.TEXT)
             If ECOARCH$("PROVED1", MKI$(NC%)) <> "" Then
                Text2.TEXT = RASOCLI$((-1) * NC%)
                Text3.TEXT = DOMICLI$((-1) * NC%)
                Text4.TEXT = CPOSCLI$((-1) * NC%)
                Text5.TEXT = LOCACLI$((-1) * NC%)
                Call SIGUEORPA1
                Exit Sub
             End If
          End If
       End If
       '
       Beep
       Text1.TEXT = ""
       '
    End If
    '
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("ECODEP")
    DPX1% = Val(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        Text10.TEXT = TRIM$(Str$(DP1%))
        Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    End If
End Sub

Private Sub Text11_LostFocus()
    '
    VDEFF$ = TRIM$(Text11.TEXT)
    Text11.TEXT = VDEFF$
    If Left$(VDEFF$, 3) <> "OP." Then GoTo 90
    NRONUE& = Val(Mid$(VDEFF$, 4))
    NROPRO& = NRONUE&
    Call NUORPA(NROPRO&)
    If NROPRO& <> NRONUE& Then
        Beep
        Call MENSERR(24, "Número no Válido o Pre-existente")
        GoTo 95
    End If
    Exit Sub
    '
90  Beep
    Call MENSERR(24, "Número no Válido")
    NROPRO& = 0
    Call NUORPA(NROPRO&)
95  Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
    Text11.SetFocus
    '
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text3_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text4_GOTFOCUS()
    Text1.SetFocus
End Sub

Private Sub Text5_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text6_Click()
    Text1.SetFocus
End Sub

Private Sub Text6_DblClick()
    STORPA04.Show 1
End Sub

Private Sub Text7_GotFocus()
    Text10.SetFocus
End Sub

Private Sub Text9_DblClick()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub Text9_GotFocus()
   Screen.MousePointer = 1
End Sub

Private Sub Text9_LostFocus()
    VDEVV$ = TRIM$(Text9.TEXT)
    If FECHANUM(VDEVV$) > 0 Then
          FF% = FECHANUM(VDEVV$)
        Else
          Call MENSERR(24, "Fecha no Válida")
          FEX = HOY(HO$)
          Text9.TEXT = HO$
          Text9.SetFocus
    End If
End Sub

Sub NUORPA(NROPRO&)
    '
    Screen.MousePointer = 11
    '
    FENUE$ = CONTROL$("", "FENUORPA")
    FLIMI = CVDAT(FECHANUM(FENUE$))
    If FLIMI > Now Then FLIMI = CVDAT(HOY(HO$))
    '
    NROMA& = 0
    Call ABRECABAN
    SQLX$ = "SELECT NumeComp FROM CajaBanc "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE 'OP*' "
    'SQLX$ = SQLX$ + "AND Status < 9 "
    'SQLX$ = SQLX$ + "AND FechEmisi >= " & FLIMI & " "
    SQLX$ = SQLX$ + "ORDER BY NumeComp DESC;"
    Set OrpaTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    On Error Resume Next
    With OrpaTemp
      .MoveFirst
      If Err < 1 Then
        NROMA& = !NumeComp
      End If
      '
      .FindFirst ("NumeComp = " & NROPRO&)
      If .NoMatch = False Then
        NROPRO& = 0
      End If
    End With
    On Error GoTo 0
    '
90  If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    Screen.MousePointer = 1
    '
End Sub

Sub SIGUEORPA1()
    '
    Screen.MousePointer = 11
    If MODORPA% = 0 Then
      MODORPA% = 1       ' por emision
      If CONTROL$("ORPASTA", "MODORDE") = "VENCIM" Then
        MODORPA% = 2
      End If
    End If
    '
    NCXI = Val(Text1.TEXT)
    If NCXI < 1 Or NCXI > 32767 Then
      Call MENSERR(24, "Numero de Cuenta no Válido")
      Text1.SetFocus
      Exit Sub
    End If
    '
    NC% = NCXI
    If ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
      Call MENSERR(24, "Cuenta no Figura\en Base de Datos de Proveedores")
      Text1.SetFocus
      Exit Sub
    End If
    '
    FE$ = TRIM$(Text9.TEXT)
    FF% = FECHANUM(FE$)
    HOI% = HOY(HO$)
    If FF% < 1 Or FF% > HOI% + 1 Then
      Call MENSERR(24, "Fecha no Válida")
      Text9.SetFocus
      Exit Sub
    End If
    '
    If DIENTRE%(FF%, HOI%) > 55 Then
      Call MENSERR(24, "Fecha no Válida")
      Text9.SetFocus
      Exit Sub
    End If
    '
    FECHALIM% = Val(CONTROL$("", "CIEMECOM"))
    If FF% <= FECHALIM% Then
      Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
      Text9.SetFocus
      Exit Sub
    End If
     '
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 11
    NC% = Val(Text1.TEXT)
    If NC% > 0 Then
      Text2.TEXT = RASOCLI$((-1) * NC%)
      Text3.TEXT = DOMICLI$((-1) * NC%)
      Text4.TEXT = CPOSCLI$((-1) * NC%)
      Text5.TEXT = LOCACLI$((-1) * NC%)
      'Text7.TEXT = TEL1CLI$(NC%)
      'Text10.TEXT = CSTRING$(MKD$(SALDCLI#(NC%)), 4, 13, 2, 2)
      Refresh
    End If
    '
100 List1.Clear
    List1.Refresh
    List2.Clear
    List2.Refresh
    '
    Z$ = Space$(128)
    Call REPLA(Z$, "Anticipo no Aplicado", 1, 20)
    Call REPLA(Z$, HO$, 23, 8)
    Call REPLA(Z$, "       0", 119, 8)
    Call REPLA(Z$, " 0", 127, 2)
    List2.AddItem Z$
    '
    SUMADEU# = 0: SUMAFAC# = 0
    If TICUEPRO%((-1) * NC%) = 1 Then        ' GASTOS
      Call BLOQARCH("!CARGAST")
      Call BLANARCH("!CARGAST")
      GoTo 200
    End If
    '
    Call ABRECABAN
    SQLX$ = "SELECT * FROM SACREPEN "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With FacPenTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          TIX$ = UCase$(Left$(!CodiCom_Lar, 2))
          X$ = !CodiCom_Lar
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "TF" Then TIPOS$ = "Ticket-Factura"
                              If TIX$ = "RH" Then TIPOS$ = "Rec.Honorarios"
                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                              If TIX$ = "CO" Then TIPOS$ = "Comprobante"
                              If TIX$ = "OP" Then TIPOS$ = "Orden de Pago": SG% = 1
                              If TIX$ = "FF" Then TIPOS$ = "O/P Fondo Fijo": SG% = 1
                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
                              If TIX$ = "DP" Then TIPOS$ = "Debito Propio": SG% = 1
                              '
          If TIX$ = "FC" Or TIX$ = "ND" Or TIX$ = "TF" Or TIX$ = "RH" Or TIX$ = "CO" Or TIX$ = "AS" Or (TIX$ = "OP" And Mid$(X$, 6, 4) = "anul") Then
            SALDO# = !SaldAcre - !SaldDebe
            If SALDO# >= 0.005 Then
              IMPO# = (-1) * !IBruComp * SG%
              SUMAFAC# = SUMAFAC# + IMPO#
              IMPO# = Abs(IMPO#)
              '
              SUMADEU# = SUMADEU# + SALDO#
              '
              NROX$ = !CodiCom_Lar
              FEMI = CVINT(!FechEmisi)
              FEVE = CVINT(!FEVENCIM)
              '
              IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 13, 2, 2)
              IVATEX$ = CSTRING$(MKD$(IVAO#), 3, 13, 2, 2)
              SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 13, 2, 2)
              '
                    REGTEX$ = CSTRING$(MKS$(REG&), 3, 8, 0, 2)
                    IRETEX$ = CSTRING$(MKI$(KKI%), 1, 2, 0, 2)
              '
              Z$ = Space$(128)
              Call REPLA(Z$, NROX$, 1, 18)
              Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
              Call REPLA(Z$, IMPTEX$, 32, 13)
              Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
              Call REPLA(Z$, SALTEX$, 56, 13)
              Call REPLA(Z$, REGTEX$, 119, 8)
              Call REPLA(Z$, IRETEX$, 127, 2)
              '
              List2.AddItem Z$
              '
            End If
          End If
        .MoveNext
        Loop
      End If
    End With
    '
200 List2.ListIndex = -1    ' CARGA DE GASTOS
    List2.Refresh
    Label8.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label8.Refresh
    Label10.Caption = ""
    Label10.Refresh
    List2.Visible = True
    List2.SetFocus
    If List2.ListCount > 0 Then
        VACON% = 1
        List2.ListIndex = 0
        VACON% = 0
    End If
    '
    Screen.MousePointer = 11
    List1.Clear
    '
    LINDICE.Clear
    If TRIM$(CONTROL$("ORPASTA", "UTICHECA")) = "SI" Then
      SQLX$ = "SELECT * FROM CAJABANC "
      SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
      SQLX$ = SQLX$ + "AND Status = 1 "
      If MODOSORVAL% = 1 Then
          SQLX$ = SQLX$ + "ORDER BY FechAcred ASC"
        Else
          SQLX$ = SQLX$ + "ORDER BY IdebeComp ASC"
      End If
      Set CHETEMP = CajaBanc.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      With CHETEMP
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          On Error GoTo 0
          Do While Not .EOF
            '
            If CVINT%(!FechEntre) < 1 Then
              FF% = CVINT(!FechAcred)
              ICHE# = !IDebeComp
              TIVA$ = TRIM$(Left$(!DescriMov, 18))
              NUCHE$ = TRIM$(Mid$(!DescriMov, 19, 14))
              '
              Z$ = Space$(128)
              Call REPLA(Z$, "Ch.3", 1, 4)
                If Left$(UCase$(TIVA$), 3) = "DOC" Then
                  Call REPLA(Z$, "Doc3", 1, 4)
                End If
              Call REPLA(Z$, TIVA$, 7, 18)
              Call REPLA(Z$, NUCHE$, 25, 14)
                FEMI = CVINT(!FechEmisi)
              Call REPLA(Z$, FECHATEX$(FEMI), 39, 8)
                FEVE = CVINT(!FechAcred)
              Call REPLA(Z$, FECHATEX$(FEVE), 49, 8)
              Call REPLA(Z$, CSTRING$(MKD$(ICHE#), 3, 12, 2, 2), 57, 12)
              Call REPLA(Z$, CSTRING$(MKS$(!NumeComp), 3, 8, 0, 2), 121, 8)
              List1.AddItem Z$
              '
            End If
          .MoveNext
          Loop
        End If
        On Error GoTo 0
      End With
      '
    End If
    '
    List1.Refresh
    Screen.MousePointer = 1
    '
    If TICUEPRO%((-1) * NC%) = 1 Then        ' GASTOS
       Call SIGUEORPA5
    End If
    '
End Sub
'
Sub SIGUEORPA3()
    '
    If List2.ListCount < 1 Then
        Exit Sub
    End If
    '
    Call BLANARCH("!CARGAPLI")
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        NROX$ = Mid$(Z$, 1, 20)
        FEMII% = FECHANUM(Mid$(Z$, 23, 8))
        IPED# = CDbl(Val(Mid$(Z$, 32, 13)))
        FEVEI% = FECHANUM(Mid$(Z$, 47, 8))
        SALD# = CDbl(Val(Mid$(Z$, 56, 13)))
        ACRE# = CDbl(Val(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 1, 18)
        REG2& = Val(Mid$(Z$, 119, 8))
        IRE% = Val(Mid$(Z$, 127, 2))
        '
        Y$ = REGBLAN$("!CARGAPLI")
        Call REPLA(Y$, NROX$, 1, 20)
        Call REPLA(Y$, MKI$(FEMII%), 21, 2)
        Call REPLA(Y$, MKD$(IPED#), 23, 8)
        Call REPLA(Y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(Y$, MKD$(SALD#), 33, 8)
        Call REPLA(Y$, MKD$(ACRE#), 41, 8)
        Call REPLA(Y$, DOCU$, 49, 18)
        Call REPLA(Y$, MKS$(REG2&), 67, 4)
        Call REPLA(Y$, MKI$(IRE%), 71, 2)
        '
        Call REGAPP("!CARGAPLI", Y$)
    Next U&
    Call CIERRARCH("!CARGAPLI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Aplicación de Pagos"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture2.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture2.Left))
    '
    VTB% = VENTABU%("CARGAPLI" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        List2.Clear
        For U& = 1 To ULTREG&("!CARGAPLI")
            Y$ = REGLEIDO$("!CARGAPLI", U&)
            NROX$ = Mid$(Y$, 1, 20)
            If TRIM$(NROX$) = "Anticipo no Aplicado" Then
               HOI% = HOY(HO$)
               ACRE# = CVD(Mid$(Y$, 41, 8))
               Call REPLA(Y$, MKI$(HOI%), 31, 2)
               Call REPLA(Y$, MKD$(ACRE#), 23, 8)
               Call REPLA(Y$, MKD$(ACRE#), 33, 8)
               Call GRAREG("!CARGAPLI", Y$, U&)
            End If
            FEMI = CVI(Mid$(Y$, 21, 2))
            IPED# = CVD(Mid$(Y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 13, 2, 2)
            FEVE = CVI(Mid$(Y$, 31, 2))
            SALD# = CVD(Mid$(Y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 13, 2, 2)
            ACRE# = CVD(Mid$(Y$, 41, 8))
            '
            If ACRE# < 0 Then
                Call MENSERR(24, "No se Permite Aplicación Negativa")
                GoTo 200
            End If
            If U& > 1 Then
              If ACRE# > SALD# Then
                Call MENSERR(24, "El Importe Acreditado Debe Ser\Inferior al Saldo del Documento")
                 GoTo 200
              End If
            End If
            ACRTEX$ = CSTRING$(MKD$(ACRE#), 4, 14, 2, 2)
            DOCU$ = Mid$(Y$, 49, 18)
            REG2& = CVS(Mid$(Y$, 67, 4))
            IRE% = CVI(Mid$(Y$, 71, 2))
            REGTEX$ = CSTRING$(MKS$(REG2&), 3, 8, 0, 2)
            IRETEX$ = CSTRING$(MKI$(IRE%), 1, 2, 0, 2)
            '
            Z$ = Space$(128)
            Call REPLA(Z$, NROX$, 1, 20)
            Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(Z$, IMPTEX$, 32, 13)
            Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
            Call REPLA(Z$, SALTEX$, 56, 13)
            Call REPLA(Z$, ACRTEX$, 70, 14)
            Call REPLA(Z$, REGTEX$, 119, 8)
            Call REPLA(Z$, IRETEX$, 127, 2)
            List2.AddItem Z$
            '
        Next U&
    End If
    '
    SUMACRE# = 0
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        SUMACRE# = SUMACRE# + CDbl(Val(Mid$(Z$, 70, 14)))
    Next U&
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label10.Refresh
    Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label18.Refresh
    Label20.Caption = CSTRING$(MKD$(Val(Label18.Caption) - Val(Label14.Caption) - Val(Label21.Caption)), 4, 14, 2, 2)
    Label20.Refresh
    '
    List2.Refresh
    VACON% = 1
    '
End Sub
'

Sub SIGUEORPA4()
   '
   VACON% = 1
   List4.Clear
10 For U& = 1 To List1.ListCount
     Z$ = List1.List(U& - 1)
     If Val(Mid$(Z$, 70, 14)) >= 0.005 Then
       List4.AddItem Z$
     End If
   Next U&
   List1.Clear
   For U& = 1 To List4.ListCount
     List1.AddItem List4.List(U& - 1)
   Next U&
   List1.Refresh
   '
   SUMAVAL# = Val(Label21.Caption)
   For U& = 1 To ULTREG&("!CARMEPA")
      X$ = REGLEIDO$("!CARMEPA", U&)
      ICHE# = CVD(Mid$(X$, 31, 8))
      COVA% = CVI(Left$(X$, 2))
      TIVA$ = ECOARCH$("MEPAGO", MKI$(COVA%))
      NUCHE$ = TRIM$(Mid$(X$, 19, 12))
      '
      Z$ = Space$(128)
      Call REPLA(Z$, "Prop", 1, 4)
      Call REPLA(Z$, TIVA$, 7, 18)
      Call REPLA(Z$, NUCHE$, 25, 14)
      FEMI = CVI(Mid$(X$, 39, 2))
      Call REPLA(Z$, FECHATEX$(FEMI), 39, 8)
      FEVE = CVI(Mid$(X$, 41, 2))
      Call REPLA(Z$, FECHATEX$(FEVE), 49, 8)
      Call REPLA(Z$, CSTRING$(MKD$(ICHE#), 3, 12, 2, 2), 57, 12)
      Call REPLA(Z$, CSTRING$(MKD$(ICHE#), 3, 14, 2, 2), 70, 14)
      Call REPLA(Z$, CSTRING$(MKI$(COVA%), 1, 6, 0, 2), 115, 6)
      List1.AddItem Z$
      SUMAVAL# = SUMAVAL# + ICHE#
      Label21.Caption = CSTRING$(MKD$(SUMAVAL#), 4, 14, 2, 2)
    '
   Next U&
   '
   List1.Refresh
   Label21.Refresh
   Label20.Caption = CSTRING$(MKD$(Val(Label18.Caption) - Val(Label14.Caption) - SUMAVAL#), 4, 14, 2, 2)
   Label20.Refresh
   '
   U& = 0
20 If U& < List2.ListCount Then
      List2.ListIndex = U&
      Z$ = List2.TEXT
      If Val(Mid$(Z$, 70, 14)) < 0.005 Then
           List2.RemoveItem U&
         Else
           U& = U& + 1
      End If
      GoTo 20
   End If
   '
   VACON% = 0
   '
   Command5.Enabled = True
   '
End Sub
'

'
Sub SIGUEORPA5()
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Distribución de Pagos"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture2.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture2.Left))
    '
305 VTB% = VENTABU%("CARGAST" + ATRI$)
    If VTB% < 0 Then
       If COMALTER%("Realmente Desea Cancelar\la Distribución Ingresada ?\\No, Continuar\Si, Cancelar") <> 2 Then
           GoTo 305
       End If
       Call CIERRARCH("!CARGAST")
       Call BLANARCH("!CARGAST")
    End If
    '
    SUMACRE# = 0
    V& = 0: UF& = ULTREG&("!CARGAST")
    For U& = 1 To UF&
       X$ = REGLEIDO$("!CARGAST", U&)
       If ECOARCH$("ECUECON", Left$(X$, 12)) <> "" Then
       CUE% = CVI(Mid$(REGACT$("ECUECON"), 41, 2))
       Call REPLA(X$, MKI$(CUE%), 21, 2)
       SUMACRE# = SUMACRE# + CVD(Mid$(X$, 13, 8))
       V& = V& + 1
       Call GRAREG("!CARGAST", X$, V&)
       End If
    Next U&
    '
    Call SETULTREG("!CARGAST", V&)
    If V& < UF& Then
       Call COMUNI("Se han Detectado Importes Mal Aplicados.\  \Revise y Confirme la Distribución de Pagos.")
       GoTo 305
    End If
    '
    Label18.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label18.Refresh
    Label20.Caption = CSTRING$(MKD$(Val(Label18.Caption) - Val(Label14.Caption) - Val(Label21.Caption)), 4, 14, 2, 2)
    Label20.Refresh
    '
    List2.ListIndex = 0
    Z$ = List2.TEXT
    IACRE$ = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Call REPLA(Z$, IACRE$, 70, 14)
    VACON% = 1
    List2.AddItem Z$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    List2.Refresh
    List2.ListIndex = 0
    VACON% = 0
    '
End Sub
'
Sub GRAORPASTA(OKX%)
    '
    Call CIERRARCH("!OBSERVOF")
    Call BLANARCH("!OBSERVOF")
    VALIDA% = VENTABU%("OBSERVOF/TITUPAN=Orden de Pagos Numero" + Str$(CVS(Left$(X3$, 4))))
    If VALIDA% < 0 Then
       OKX% = (-1): Exit Sub
    End If
    '
    HP% = HOY(HOS$)
    OKX% = 1                    ' bloqueó archivos
    '
    NC% = Val(Text1.TEXT)
    NC1% = (-1) * NC%
    NRO& = Val(Mid$(Text11.TEXT, 4))
    NROPRO& = NRO&
    Call NUORPA(NROPRO&)
    If NROPRO& <> NRO& Then
       Call MENSERR(24, "Al Grabar esta Orden de Pago\el Número Propuesto Estaba Ocupado.\  \Verifique el Nuevo Número Propuesto\y Vuelva a Grabar e Imprimir.")
       Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
       OKX% = 0
       Exit Sub
    End If
    '
    TIPO% = 4: VA% = 1: MO$ = "$"
    TOFAC#(0) = Val(Label14.Caption) + Val(Label21.Caption)
    TOFAC#(1) = 0
    TOFAC#(2) = TOFAC#(0)
    '
    FC% = FECHANUM(Text9.TEXT)
    FVII%(1) = FC%
    IMV#(1) = 0
    List2.ListIndex = 0
    X$ = List2.TEXT
    DOCU$ = Mid$(X$, 1, 20)
    If DOCU$ = "Anticipo no Aplicado" Then
      CANCE# = CDbl(Val(Mid$(X$, 70, 14)))
      If Abs(CANCE#) >= 0.005 Then
        IMV#(1) = (-1) * CANCE#
      End If
    End If
    '
' grabar orden de pago en CAJABANC
    COBRACT% = 0
    NRECIB$ = "0" ' TALONARIO
    While Len(NRECIB$) < 2: NRECIB$ = "0" + NRECIB$: Wend
    NRECIB$ = TRIM$(Str$(NRO&)) + "-" + NRECIB$
    While Len(NRECIB$) < 9: NRECIB$ = "0" + NRECIB$: Wend
    NRECIB$ = "OP-" + NRECIB$
    '
    RECILAR$ = TRIM$(Str$(NRO&))
    While Len(RECILAR$) < 8: RECILAR$ = "0" + RECILAR$: Wend
    RECILAR$ = TRIM$(Str$(COBRACT%)) + "-" + RECILAR$
    While Len(RECILAR$) < 13: RECILAR = "0" + RECILAR$: Wend
    RECILAR$ = "OP-X-" + RECILAR$
    '
    CoComCor$ = NRECIB$
    CoComLar$ = RECILAR$
    CodiEmp% = 0
    CodiSuc% = SUC%
    NumeClie% = 0
    NumeProv% = NC%
    FechEmi = CVDAT(FC%)
    Referen$ = RASOCLI$(NC1%)
    FeContab = CVDAT(FC%)
    FeAcredi = CVDAT(FC%)
    CuentaCo% = CUECOCLI%(NC1%)
    TiMovi$ = "OPAG"
    OpMovi$ = "OP"
    VaMovi$ = "CCORR"
    Descrip$ = RECILAR$ + " - " + RASOCLI$(NC1%)
    NumeTal = COBRACT%
    NumeCom = NRO&
    StatInit = 0
    ImpoDeb# = Abs(TOFAC#(0))
    Observa$ = RASOCLI$(NC1%)
    Call GRACABAN
    OKX% = 2         ' grabó Orden de Pago
    '
' grabar los valores entregados
    For U& = 1 To List1.ListCount
       Z$ = List1.List(U& - 1)
       TIDOS$ = Left$(Z$, 4)  ' Prop
       
       
       TIPO% = Val(Mid$(Z$, 115, 6))
       BANCO$ = Mid$(Z$, 7, 18)
       NUCHE$ = Mid$(Z$, 25, 14)
       IMCHE# = Val(Mid$(Z$, 70, 14))
       FECHE% = FECHANUM(Mid$(Z$, 39, 8)): If FECHE% < 1 Then FECHE% = HO%
       FEVEN% = FECHANUM(Mid$(Z$, 49, 8)): If FEVEN% < FECHE% Then FEVEN% = FECHE%
         FEX = FEVEN%: FEVECHE$ = FECHATEX$(FEX)
       IDEVA$ = AJUSTI$(BANCO$, 18) + AJUSTI$(NUCHE$, 14) + FECHATEX$(FEX)
       '
       HOII% = HOY(HOS$)
       If IMCHE# > 0 Then
          REGBA& = Val(TRIM$(Mid$(Z$, 121, 8)))
          If REGBA& > 0 Then ' es un cheque de cartera
             'ESCRIBIR ESTA RUTINA PARA DARLE SALIDA DE CARTERA
'               TTS$ = REGLEIDO$("CAJABANC", REGBA&)
'               Call REPLA(TTS$, MKI$(FC%), 275, 2)
'               Call REPLA(TTS$, NORPAG$, 277, 16)
'               Call REPLA(TTS$, RASOCLI$(NC1%), 293, 30)
'               Call REPLA(TTS$, "Orden de Pago Nro." + Str$(NRO&), 323, 40)
'               Call REPLA(TTS$, HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 363, 40)
'               Call REPLA(TTS$, "Usuario: " + TRIM$(USERNAME$), 403, 40)
'               Call REPLA(TTS$, Chr$(2), 121, 1)  ' Marca Salida de Cartera
'               Call REPLA(TTS$, IDACTU$, 263, 6)  ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
'               Call GRAREG("CAJABANC", TTS$, REGBA&)
               '
          End If
          '
          CoComCor$ = NRECIB$
          CoComLar$ = RECILAR$
          CodiEmp% = 0
          CodiSuc% = SUC%
          NumeClie% = 0
          NumeProv% = NC%
          FechEmi = CVDAT(FECHE%)
            FEX = FC%
          Referen$ = RASOCLI$(NC1%) + " - " + FECHATEX$(FEX)
          FeContab = CVDAT(FC%)
          '
          FeAcredi = CVDAT(FEVEN%)
          CuentaCo% = CUEMEPAG%(TIPO%)
          TiMovi$ = "CHENTR"
          OpMovi$ = "OP"
          VaMovi$ = ""
          StatInit = 0
          Descrip$ = IDEVA$
          NumeTal = COBRACT%
          NumeCom = NRO&
          ImpoDeb# = (-1) * Abs(IMCHE#)
          Observa$ = "O.Pago " + TRIM$(Str$(NRO&)) + " - " + RASOCLI$(NC1%)
          Z$ = TTS$
          Call REPLA(Z$, Mid$(REBLA$, 209), 209, 304)
          '
          Call GRACABAN
       End If
    Next U&
    '
    ' Ret I.Brutos Buenos Aires
    If Val(Label14) >= 0.005 Then
       NURETBA& = 1
       If ULTREG&("SURETIB") > 0 Then
         Z1$ = REGLEIDO$("SURETIB", ULTREG&("SURETIB"))
         NURETBA& = 1 + CVS(Mid$(Z1$, 3, 4))
       End If
       If NURETBA& < 10 Then
         NURETBA& = Val(InputBox$("Numero de Comprobante I.B.", "Retencion Ingresos Brutos", TRIM$(Str$(NURETBA&))))
       End If
       Label25 = TRIM$(Str$(NURETBA&))
       '
       Z1$ = REGBLAN$("SURETIB")
       Call REPLA(Z1$, MKI$(FC%), 1, 2)
       Call REPLA(Z1$, MKS$(NURETBA&), 3, 4)
       Call REPLA(Z1$, MKS$(NRO&), 7, 4)
       Call REPLA(Z1$, RASOCLI$(NC1%), 21, 30)
       Call REPLA(Z1$, CUITCLI$(NC1%), 51, 14)
       Call REPLA(Z1$, NIBRCLI$(NC1%), 65, 14)
       Call REPLA(Z1$, MKD$(Val(Label23)), 79, 8)
       Call REPLA(Z1$, MKD$(Val(Label14)), 87, 8)
       Call REGAPP("SURETIB", Z1$)
       Call CIERRARCH("SURETIB")
       '
       CoComCor$ = NRECIB$
       CoComLar$ = RECILAR$
       CodiEmp% = 0
       CodiSuc% = SUC%
       NumeClie% = 0
       NumeProv% = NC%
       FechEmi = CVDAT(FC%)
         FEX = FC%
       Referen$ = RASOCLI$(NC1%) + " - " + FECHATEX$(FEX)
       FeContab = CVDAT(FC%)
       '
       FeAcredi = CVDAT(FC%)
         TIPO% = Val(Label24)
       CuentaCo% = CUEMEPAG%(TIPO%)
       TiMovi$ = "RTIB"
       OpMovi$ = "Bs-As"
       VaMovi$ = ""
       StatInit = 0
       Descrip$ = "Ret.IB.Bs.As." + TRIM$(Str$(NURETBA&))
       NumeTal = COBRACT%
       NumeCom = NRO&
       ImpoDeb# = (-1) * Abs(Val(Label14))
       Observa$ = "O.Pago " + TRIM$(Str$(NRO&)) + " - " + RASOCLI$(NC1%)
       Call GRACABAN
    End If
    '
    OKX% = 3          ' grabo cheques de terceros
    OKX% = 4          ' grabo cheques emitidos
    '
800 ' ******************** aqui blanquear facturas canceladas
    ' ******************** tambien grabar archivo APLIPAG.DAT
    '
    If TICUEPRO%(NC1%) = 0 Then
       '
       CoComCor$ = NRECIB$
       CoComLar$ = RECILAR$
       CodiEmp% = 0
       CodiSuc% = SUC%
       NumeClie% = 0
       NumeProv% = NC%
       FechEmi = CVDAT(FC%)
       RefApli$ = "Orden de Pago " + NRECIB$
       '
       XX% = 0
       For U& = 1 To List2.ListCount
         List2.ListIndex = U& - 1
         Z$ = List2.List(U& - 1)
         DOCU$ = Left$(Z$, 20)
         If DOCU$ <> "Anticipo no Aplicado" Then
              CANCE# = CDbl(Val(Mid$(Z$, 70, 14)))
              If CANCE# > 0 Then
                DOCUCA$ = TRIM$(Left$(Z$, 18))
                ICAN# = CANCE#
                '
                DocOrCor$ = ""
                TotOrig# = 0
                IvaOrig# = 0
                '
                Call ABRECABAN
                With SacrePen
                  .FindFirst "NuProve = " & NumeProv% & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDACRE > 0"
                  If .NoMatch = False Then
850                 PARCAN# = ICAN#: If !SaldAcre < ICAN# Then PARCAN# = !SaldAcre
                    .Edit
                    !SaldAcre = !SaldAcre - PARCAN#
                    .Update
                    '
                    DocOrCor$ = !CodiCom_Cor
                    TotOrig# = !IBruComp
                    IIVAORIG# = 0
                    On Error Resume Next
                    IvaOrig# = !IIvaComp
                    On Error GoTo 0
                    '
                    ICAN# = ICAN# - PARCAN#
                    If ICAN# >= 0.005 Then
                      .FindNext "NuProve = " & NumeProv% & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDACRE > 0"
                      GoTo 850
                    End If
                  End If
                End With
                '
                DocOrLar$ = TRIM$(Left$(Z$, 18))
                ImpApli# = CANCE#
                Call GRAPLIPAG
              End If
            Else
              CANCE# = CDbl(Val(Mid$(Z$, 70, 14)))
              If CANCE# > 0 Then
                IBruTot# = Val(Label18)
                NuorVen% = 0
                FeVenc = CVDAT(FC%)
                ImVenc# = (-1) * CANCE#
                Call GRACREPEN
              End If
           End If
       Next U&
       '
    End If
    OKX% = 5         ' depuró facturas pendientes
    OKX% = 6         ' grabó aplicaciones
    '
820 If TICUEPRO%(NC1%) = 1 Then
MsgBox "FALTA GRABAR CASO CUENTAS DE GASTOS"
'
        RC& = REGICLI&(NC1%)
        DX$ = REGLEIDO$("PROVED2", RC&)
        DX$ = REPLACE$(DX$, MKD$(0), 109, 8)
        Call GRAREG("PROVED2", DX$, RC&)
        '
        For K1& = 1 To ULTREG&("!CARGAST")
           Z$ = REGLEIDO$("!CARGAST", K1&)
           ZX$ = REGBLAN$("DETPAGAS")
           ZX$ = REPLACE$(ZX$, MKI$(NC%), 1, 2)
           ZX$ = REPLACE$(ZX$, MKS$(NRO&), 3, 4)
           ZX$ = REPLACE$(ZX$, MKI$(FC%), 7, 2)
           ZX$ = REPLACE$(ZX$, Mid$(Z$, 13, 10), 9, 10)
           AA = CVD(Mid$(Z$, 3))
           Call REGAPP("DETPAGAS", ZX$)
        Next K1&
    End If
    '
    OKX% = 7          ' LIBERO ARCHIVOS
    '
End Sub
'
Sub PRIORPASTA()
    '
    HOI% = HOY(HO$)
    NC% = Val(Text1)
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORPASTA"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Orden de Pago.")
        Exit Sub
    End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Recibo de Cobranza:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    CODFOR$ = FORIPRE$
    '
    CONTA% = List1.ListCount
    If CONTA% < 1 Then
        Call COMUNI("No se ha Ingresado Detalle de Valores")
        Exit Sub
    End If
    '
    If Val(Label14.Caption) + Val(Label21.Caption) < 0.005 Then
        Call COMUNI("No se ha Ingresado Detalle de Valores")
        Exit Sub
    End If
    '
    AVAHOJA$ = TRIM$(CONTROL$("ORPASTA", "AVAHOJA"))
    Screen.MousePointer = 11
    Call BLANARCH("!PRINRECI")
    '
    ILISTA& = List1.ListIndex
    TOLISTA& = List1.TopIndex
    '
    For U& = 1 To CONTA%
        List1.ListIndex = U& - 1
        X$ = List1.TEXT
        If TRIM$(X$) = "" Then GoTo 19
        TIPOS$ = Left$(X$, 4)
        BANCO$ = Mid$(X$, 7, 18)
        NUVA$ = Mid$(X$, 25, 14)
        IMVA# = Val(Mid$(X$, 70, 14))
        If IMVA# < 0.005 Then GoTo 19
        '
        FEMII% = FECHANUM(Mid$(X$, 39, 8)): If FEMII% < 1 Then FEMII% = HOI%
        FEVEI% = FECHANUM(Mid$(X$, 49, 8)): If FEVEI% < 1 Then FEMII% = HOI%
        PROTER$ = ""
        CLEA% = 48
        '
10      Z$ = REGBLAN$("!PRINRECI")
        Call REPLA(Z$, TIPOS$, 1, 4)
        Call REPLA(Z$, NUVA$, 5, 12)
        Call REPLA(Z$, BANCO$, 17, 16)
        Call REPLA(Z$, MKD$(IMVA#), 33, 8)
        FEX = FEVEI%
        If FEX > 0 Then
            Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
        End If
        Call REGAPP("!PRINRECI", Z$)
        '
19  Next U&
    '
    If Val(Label14) >= 0.005 Then
       Z$ = REGBLAN$("!PRINRECI")
       Call REPLA(Z$, "R.IB", 1, 4)
       Call REPLA(Z$, TRIM$(Label25), 5, 12)
       Call REPLA(Z$, "Ret.IB.Bs.As.", 17, 16)
       Call REPLA(Z$, MKD$(Val(Label14)), 33, 8)
       FEX = HOY(HOS$)
       Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
       Call REGAPP("!PRINRECI", Z$)
    End If
    '
    V0& = ULTREG&("!PRINRECI")
    '
    V0& = ULTREG&("!PRINRECI")
    '
    ILISTA& = List2.ListIndex
    TOLISTA& = List2.TopIndex
    '
    V& = 0
    If TICUEPRO%((-1) * NC%) = 1 Then
       For U& = 1 To ULTREG&("!CARGAST")
          X$ = REGLEIDO$("!CARGAST", U&)
          DOCU$ = Left$(X$, 12)
          SALPE# = 0
          CANCE# = CVD(Mid$(X$, 13, 8))
          If Abs(CANCE#) >= 0.005 Then
             V& = V& + 1
             If V& <= V0& Then
                  Z$ = REGLEIDO$("!PRINRECI", V&)
                Else
                  Z$ = REGBLAN$("!PRINRECI")
             End If
             Call REPLA(Z$, DOCU$, 53, 18)
             Call REPLA(Z$, MKD$(CANCE#), 71, 8)
             Call REPLA(Z$, MKD$(SALPE#), 79, 8)
             Call GRAREG("!PRINRECI", Z$, V&)
          End If
       Next U&
       Call CIERRARCH("!CARGAST")
       GoTo 25
    End If
    '
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        X$ = List2.TEXT
        DOCU$ = Mid$(X$, 1, 20)
        SALPE# = CDbl(Val(Mid$(X$, 56, 13)))
        CANCE# = CDbl(Val(Mid$(X$, 70, 14)))
        If Abs(CANCE#) >= 0.005 Then
            V& = V& + 1
            If V& <= V0& Then
                  Z$ = REGLEIDO$("!PRINRECI", V&)
               Else
                  Z$ = REGBLAN$("!PRINRECI")
            End If
            Call REPLA(Z$, DOCU$, 53, 18)
            Call REPLA(Z$, MKD$(CANCE#), 71, 8)
            Call REPLA(Z$, MKD$(SALPE#), 79, 8)
            Call GRAREG("!PRINRECI", Z$, V&)
        End If
    Next U&
    '
25  V0& = ULTREG&("!PRINRECI")
    '
    TOTAVAL# = CDbl(Val(Label14.Caption) + Val(Label21.Caption))
    TOTACAN# = CDbl(Val(Label18.Caption)) + CDbl(Val(Label20.Caption))
33  ITEXTO$ = NUMTEX$(TOTAVAL#) + " ctvs"
    LOTEX% = Val(ATRIFORM$(CODFOR$, "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = "": IMPLET2$ = "": IMPLET3$ = ""
    If TRIM$(NRX1$) <> "" Then
       IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    End If
    If TRIM$(NRX2$) <> "" Then
       IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
    End If
    If TRIM$(NRX3$) <> "" Then
       IMPLET3$ = TRIM$(NRX3$) + " " + "***************************************************************************************************************"
    End If
    '
    Call CIERRARCH("!PRINRECI")
    '
30  If TRIM$(FORIPRE$) <> "" Then
        '
        CANCELPRINT% = 0
        If CONTROL$("ORPASTA", "SUPRINT") = "SI" Then CANCELPRINT% = 1
        '
        CODFOR$ = TRIM$(UCase$(FORIPRE$))
        MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
        MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
        '
        INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
        If INTERLI < 1 Then
            Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
            INTERLI = 1
        End If
        '
        DESMM = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
        HASMM = Val(ATRIFORM$(CODFOR$, "DETHASMM"))
        '
        CACOPI% = Val(ATRIFORM$(CODFOR$, "CAN-COP"))
        If CACOPI% < 1 Or CACOPI% > 6 Then
            Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
            CACOPI% = 1
        End If
        '
        ORIENTA% = 1
        PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
        If PORI$ = "PAISAJE" Or PORI$ = "2" Then
            ORIENTA% = 2
        End If
        '
        PORIAN = Printer.Orientation  ' retiene orientacion
        '
        NRORPA$ = TRIM$(Mid$(Text11.TEXT, 4))
        '
        Screen.MousePointer = 11
        If AVAHOJA$ <> "NEWPAGE" Then
            'Printer.Orientation = 1
            Call SETSPOOL("ORIENTATION", "1")
            If ORIENTA% = 2 Then
                'Printer.Orientation = 2
                Call SETSPOOL("ORIENTATION", "2")
            End If
        End If
        '
        NC1% = (-1) * NC%
        For COP% = 1 To CACOPI%
            '
            If AVAHOJA$ = "NEWPAGE" Then
               If PRINTERCONTROL% = 1 Then
                  'Printer.NewPage
                  Call SETSPOOL("NEWPAGE", "")
               End If
            End If
            PRINTERCONTROL% = 1
            '
            SUMAVA# = 0: SUMACAN# = 0
            Call CARFORWIN
            '
            Call PRINTFORWIN(TRIM$(Str$(NC%)), "CUEN-CLI", 0, 0)
            Call PRINTFORWIN(TRIM$(RASOCLI$(NC1%)), "RASO-CLI", 0, 0)
            Call PRINTFORWIN(DOMICLI$(NC1%), "DOMI-CLI", 0, 0)
            Call PRINTFORWIN(CPOSCLI$(NC1%), "CPOS-CLI", 0, 0)
            Call PRINTFORWIN(LOCACLI$(NC1%), "LOCA-CLI", 0, 0)
            Call PRINTFORWIN(CUITCLI$(NC1%), "CUIT-CLI", 0, 0)
            Call PRINTFORWIN(POSIVAL$(NC1%), "PIVA-CLI", 0, 0)
            FF% = FECHANUM(TRIM$(Text9.TEXT))
            Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
            Call PRINTFORWIN(NRORPA$, "NUME-COM", 0, 0)
            '
            YAC = 0
            For U& = 1 To ULTREG&("!OBSERVOF")
               TXOB$ = TRIM$(REGLEIDO$("!OBSERVOF", U&))
               Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
               YAC = YAC + INTERLI
            Next U&
            Call CIERRARCH("!OBSERVOF")
            '
            YAC = DESMM
            For YX& = 1 To ULTREG&("!PRINRECI")
               '
               YAC = YAC + INTERLI
               '
               If YAC > HASMM - 2 * INTERLI Then
                   YAC = YAC + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   '
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-VAL", 0, 0)
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-CAN", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET1", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET2", 0, 0)
                   '
                   If AVAHOJA$ = "NEWPAGE" Then
                        'Printer.NewPage
                        Call SETSPOOL("NEWPAGE", "")
                      Else
                       'Printer.EndDoc
                       Call SETSPOOL("ENDDOC", "")
                   End If
                   '
                   Call CARFORWIN
                   Call PRINTFORWIN(TRIM$(Str$(NC1%)), "CUEN-CLI", 0, 0)
                   Call PRINTFORWIN(TRIM$(RASOCLI$(NC1%)), "RASO-CLI", 0, 0)
                   Call PRINTFORWIN(DOMICLI$(NC1%), "DOMI-CLI", 0, 0)
                   Call PRINTFORWIN(CPOSCLI$(NC1%), "CPOS-CLI", 0, 0)
                   Call PRINTFORWIN(LOCACLI$(NC1%), "LOCA-CLI", 0, 0)
                   Call PRINTFORWIN(CUITCLI$(NC1%), "CUIT-CLI", 0, 0)
                   Call PRINTFORWIN(POSIVAL$(NC1%), "PIVA-CLI", 0, 0)
                   FEX = FF%
                   Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
                   Call PRINTFORWIN(NRORPA$, "NUME-COM", 0, 0)
                   '
                   YAC = DESMM + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   YAC = YAC + 2 * INTERLI
                   '
                End If
                '
                YY$ = REGLEIDO$("!PRINRECI", YX&)
                Call PRINTFORWIN(Left$(YY$, 4), "TIPO-VAL", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 5, 12), "NUME-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 17, 16), "BANC-CHE", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 33, 8), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                SUMAVA# = SUMAVA# + CVD(Mid$(YY$, 33, 8))
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 41, 8), 6, 8, 0, 1), "FECH-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 53, 18), "DOCU-APL", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 71, 8), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                SUMACAN# = SUMACAN# + CVD(Mid$(YY$, 71, 8))
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 79, 8), 4, 12, 2, 2), "IMPO-PEN", 0, YAC)
                '
            Next YX&
            '
            Call PRINTFORWIN(CSTRING$(MKD$(TOTAVAL#), 4, 12, 2, 2), "TOTA-VAL", 0, 0)
            Call PRINTFORWIN(CSTRING$(MKD$(TOTACAN#), 4, 12, 2, 2), "TOTA-CAN", 0, 0)
            Call PRINTFORWIN(IMPLET1$, "IMP-LET1", 0, 0)
            Call PRINTFORWIN(IMPLET2$, "IMP-LET2", 0, 0)
                   '
            If AVAHOJA$ <> "NEWPAGE" Then
                 'Printer.EndDoc
                 Call SETSPOOL("ENDDOC", "")
            End If
            '
        Next COP%
        '
        Call CIERRARCH("*.*")
        DESCRIDOC$ = "Orden de Pago Nro. " + TRIM$(NRORPA$) + " - " + TRIM$(RASOCLI$(NC1%))
        FEX = FF%
        DESCRIDOC$ = DESCRIDOC$ + " - " + FECHATEX$(FEX)
        '
        Call SAVESPOOL(DESCRIDOC$, OKX%)
        Screen.MousePointer = 1
        If COMALTER%("Orden de Pago Numero\'" + NRORPA$ + "'\Emitida y Registrada\\Repetir Impresion\Otra Orden de Pago") = 1 Then GoTo 30
        '
        If AVAHOJA$ <> "NEWPAGE" Then
            Printer.Orientation = PORIAN
        End If
        '
    End If
    Call CIERRARCH("*.*")
    '
End Sub

Private Sub veryreimp_Click()
    Call OPORPA04.Command23_Click
End Sub

Sub CALRETIBRU()
   '
   Static CTCALCU%, ALIBRU(8), MINIMPO#(8), MEPAGX%(8)
   If CTCALCU% = 0 Then
     CTCALCU% = (-1)
     'EPAC$ = TRIM$(UCase$(EMPREAC$))
     'If InStr(EPAC$, "SABO") > 0 Or InStr(EPAC$, "DOSIVA") > 0 Then
       CTCALCU% = 1
       For U& = 1 To ULTREG&("TACATIBP")
         TCX$ = REGLEIDO$("TACATIBP", U&)
         TCY% = Asc(Mid$(TCX$, 1, 1))
         If TCY% > 0 Then
           If TCY% <= 8 Then
             ALIBRU(TCY%) = CVS(Mid$(TCX$, 33, 4))
             MINIMPO#(TCY%) = CVS(Mid$(TCX$, 37, 4))
             MEPAGX%(TCY%) = CVI(Mid$(TCX$, 41, 2))
           End If
         End If
       Next U&
       Call CIERRARCH("TACATIBP")
     'End If
   End If
   '
   If CTCALCU% < 0 Then Exit Sub
   '
   'If FECHANUM(Text9) < FECHANUM("01/08/03") Then Exit Sub
   '
   ' determinar categoria de I.Brutos
   NPROVEI% = Val(Text1)
   RERET$ = FILTERGETRECORD$("DARIBPRO", 1, MKI$(NPROVEI%))
   CATIB1% = Asc(Mid$(RERET$, 3, 1))
   PORCONVE = CVS(Mid$(RERET$, 5, 4))
   If CATIB1% < 1 Then Exit Sub  ' exento
   '
   ' determinar monto imponible
   BASEIMTO# = 0
   DIVIAN = 1
   If PIVACLI%((-1) * NPROVEI%) = 1 Then
     DIVIAN = 1 + ALIVA / 100
   End If
   For UKI% = 1 To List2.ListCount
     Z$ = List2.List(UKI% - 1)
     IPARCI# = Val(Mid$(Z$, 70, 14))
     If IPARCI# >= 0.005 Then
       BASEIMPAR# = IPARCI# / DIVIAN
       RECUECO& = Val(Mid$(Z$, 119, 8))
       If RECUECO& > 0 Then
         X$ = REGLEIDO$("CUECORR", RECUECO&)
         IGRAEX# = Abs(CVD(Mid$(X$, 51, 8)) + CVD(Mid$(X$, 59, 8)))
         ITOTAX# = Abs(CVD(Mid$(X$, 115, 8)) - CVD(Mid$(X$, 123, 8)))
         COEFI = 1
         If ITOTAX# > 0.005 And IGRAEX# <= ITOTAX# Then
           COEFI = IGRAEX# / ITOTAX#
         End If
         BASEIMPAR# = IPARCI# * COEFI
       End If
       BASEIMTO# = BASEIMTO# + BASEIMPAR#
     End If
   Next UKI%
   '
   Label29 = "": Label27 = ""
   IRETIB# = 0
   If CATIB1% <= 8 Then
     LabeL28 = FORMATNUM$(MINIMPO#(CATIB1%), "F12.2")
     If BASEIMTO# >= MINIMPO#(CATIB1%) Then
       IRETIB# = BASEIMTO# * PORCONVE / 100 * ALIBRU(CATIB1%) / 100
       Label24 = Str$(MEPAGX%(CATIB1%))
       Label29 = FORMATNUM$(ALIBRU(CATIB1%), "F8.2")
       Label27 = FORMATNUM$(PORCONVE, "F8.2")
     End If
   End If
   '
   Label14 = FORMATNUM$(IRETIB#, "F12.2")
   Label23 = FORMATNUM$(BASEIMTO#, "F12.2")
   '
End Sub

Sub PRICORETIB()
   '
   If Val(Label14) >= 0.005 Then
      NURETBA& = Val(Label25)
      Screen.MousePointer = 11
      FECHDOC$ = FETEXTO$(Int(FECHANUM(Text9.TEXT)))
      NUMEDOC$ = TRIM$(Str$(NURETBA&))
      While Len(NUMEDOC$) < 8
         NUMEDOC$ = "0" + NUMEDOC$
      Wend
      NUMEDOC$ = "0001-" + NUMEDOC$
      TIPODOC$ = "Retencion I.Brutos"
      '
      CODPARAM$(1) = "IMPO-CHE"
      CODPARAM$(2) = "IMPO-PEN"
      CODPARAM$(3) = "IMPO-CAN"
      CODPARAM$(4) = "POR-DESC"
      CAPARDOC% = 4
      '
      DOCPARAM$(1) = TRIM$(Label23)
      DOCPARAM$(2) = TRIM$(Label29)
      DOCPARAM$(3) = TRIM$(Label14)
      DOCPARAM$(4) = TRIM$(Label27)
      '
      CACOLTAB1% = 0
      '
      DESTIDOC% = (-1) * Val(Text1)
      CAITAB1% = 0
      '
      Call PRINTDOC("FORRETIB")   ' Retencion I.Brutos
      '
   End If
   '
End Sub

