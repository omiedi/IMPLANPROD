VERSION 5.00
Begin VB.Form STORPA04 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formularios de Comprobantes de Pago Standard"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   11550
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
   ScaleHeight     =   4620
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00404000&
      Height          =   7470
      Left            =   10605
      ScaleHeight     =   7410
      ScaleWidth      =   1230
      TabIndex        =   35
      Top             =   0
      Width           =   1290
      Begin VB.PictureBox Picture16 
         Height          =   4590
         Left            =   0
         ScaleHeight     =   4530
         ScaleWidth      =   1005
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
         Begin VB.Image Image3 
            Height          =   4550
            Left            =   0
            Picture         =   "STORPA04.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1065
         End
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
         Height          =   690
         Left            =   140
         Picture         =   "STORPA04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Cotización Dolar y Tipo de Cambio"
         Top             =   3250
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Tbls"
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
         Picture         =   "STORPA04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Acceso Directo a Tabla de Bancos y Valores de Pago"
         Top             =   2560
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Picture         =   "STORPA04.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   240
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
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
         Left            =   140
         MaskColor       =   &H00404000&
         Picture         =   "STORPA04.frx":2936
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   940
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   37
         Top             =   4410
         Visible         =   0   'False
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command2 
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
         Left            =   140
         Picture         =   "STORPA04.frx":2C40
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Acceso Directo a Tabla de Bancos y Valores de Pago"
         Top             =   1750
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "STORPA04.frx":2F4A
         Top             =   4020
         Width           =   1515
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Formularios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   160
      TabIndex        =   22
      Top             =   210
      Width           =   4950
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cálculo Retención I.Brutos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   240
         TabIndex        =   56
         Top             =   2660
         Width           =   2895
         Begin VB.CheckBox Check3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Acumulación Mensual"
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
            Left            =   420
            TabIndex        =   57
            Top             =   300
            Width           =   2205
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   3255
         TabIndex        =   52
         Top             =   2660
         Width           =   1485
         Begin VB.OptionButton Option10 
            BackColor       =   &H00FFFFC0&
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
            Left            =   420
            TabIndex        =   54
            Top             =   420
            Width           =   780
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H00FFFFC0&
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
            Left            =   420
            TabIndex        =   53
            Top             =   210
            Value           =   -1  'True
            Width           =   750
         End
      End
      Begin VB.CommandButton Command12 
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
         Height          =   615
         Left            =   4095
         Picture         =   "STORPA04.frx":4E6C
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3495
         Width           =   645
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2415
         TabIndex        =   49
         Text            =   " "
         Top             =   3495
         Width           =   1380
      End
      Begin VB.CommandButton Command11 
         Caption         =   "."
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
         Left            =   3780
         TabIndex        =   48
         Top             =   3495
         Width           =   175
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
         Height          =   285
         Left            =   210
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3825
         Width           =   3750
      End
      Begin VB.TextBox Text8 
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
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2270
         Width           =   3750
      End
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
         Height          =   285
         Left            =   3780
         TabIndex        =   44
         Top             =   1935
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
         Height          =   285
         Left            =   2415
         TabIndex        =   43
         Text            =   " "
         Top             =   1935
         Width           =   1380
      End
      Begin VB.CommandButton Command9 
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
         Height          =   615
         Left            =   4095
         Picture         =   "STORPA04.frx":5176
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1935
         Width           =   645
      End
      Begin VB.CommandButton Command8 
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
         Height          =   615
         Left            =   4095
         Picture         =   "STORPA04.frx":5480
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1035
         Width           =   645
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
         Height          =   285
         Left            =   210
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1370
         Width           =   3750
      End
      Begin VB.CommandButton Command6 
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
         Height          =   285
         Left            =   3780
         TabIndex        =   29
         Top             =   1035
         Width           =   175
      End
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
         Height          =   285
         Left            =   2415
         TabIndex        =   28
         Text            =   " "
         Top             =   1035
         Width           =   1380
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
         Height          =   615
         Left            =   4095
         Picture         =   "STORPA04.frx":578A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   315
         Width           =   645
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
         Height          =   285
         Left            =   210
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   " "
         Top             =   650
         Width           =   3750
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
         Height          =   285
         Left            =   3780
         TabIndex        =   24
         Top             =   315
         Width           =   175
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
         Height          =   285
         Left            =   2415
         TabIndex        =   23
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Retención de Ganancias:"
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
         Left            =   0
         TabIndex        =   51
         Top             =   3570
         Width           =   2325
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Retención Ing.Brutos:"
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
         Left            =   0
         TabIndex        =   42
         Top             =   2040
         Width           =   2325
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   5145
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Liquidación Fondo Fijo:"
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
         Left            =   0
         TabIndex        =   32
         Top             =   1140
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de Pago:"
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
         Left            =   0
         TabIndex        =   27
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Configuración de Ordenes de Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   5250
      TabIndex        =   0
      Top             =   210
      Width           =   5160
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   " Suprime  Impresión  al  Generar  el  Documento"
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
         Left            =   240
         TabIndex        =   46
         Top             =   3900
         Width           =   4650
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cartera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   210
         TabIndex        =   33
         Top             =   420
         Width           =   1650
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Utiliza Carte- ra de Valores"
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
            Left            =   105
            TabIndex        =   34
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ordenamiento"
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
         Left            =   210
         TabIndex        =   19
         Top             =   2860
         Width           =   1650
         Begin VB.OptionButton Option8 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Vencim."
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
            TabIndex        =   21
            Top             =   525
            Width           =   1275
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Emision"
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
            TabIndex        =   20
            Top             =   315
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Prefijo"
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
         Left            =   3675
         TabIndex        =   11
         Top             =   1600
         Width           =   1275
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
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "0001"
            Top             =   525
            Width           =   870
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Papel"
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
         Left            =   210
         TabIndex        =   10
         Top             =   1600
         Width           =   3330
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   380
            Left            =   315
            TabIndex        =   16
            Top             =   735
            Width           =   2850
            Begin VB.OptionButton Option6 
               BackColor       =   &H00FFFFC0&
               Caption         =   "NewPage"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1575
               TabIndex        =   18
               Top             =   0
               Width           =   1275
            End
            Begin VB.OptionButton Option5 
               BackColor       =   &H00FFFFC0&
               Caption         =   "EndDoc"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   0
               TabIndex        =   17
               Top             =   0
               Value           =   -1  'True
               Width           =   1275
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   435
            Left            =   315
            TabIndex        =   13
            Top             =   315
            Width           =   2850
            Begin VB.OptionButton Option3 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Formulario Continuo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton Option4 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Hoja Suelta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1575
               TabIndex        =   14
               Top             =   0
               Width           =   1065
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Talonario en Uso"
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
         Left            =   1995
         TabIndex        =   6
         Top             =   2860
         Width           =   2955
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
            Left            =   2580
            TabIndex        =   8
            Top             =   315
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
            Left            =   1425
            TabIndex        =   7
            Text            =   " "
            Top             =   315
            Width           =   1185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nueva Numeración"
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
            Left            =   165
            TabIndex        =   9
            Top             =   270
            Width           =   1170
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   1995
         TabIndex        =   3
         Top             =   420
         Width           =   1965
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pre-Impresa"
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
            TabIndex        =   5
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Auto-Numera"
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
            TabIndex        =   4
            Top             =   630
            Width           =   1485
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
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
         Height          =   1065
         Left            =   4095
         TabIndex        =   1
         Top             =   420
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
            TabIndex        =   2
            TabStop         =   0   'False
            Text            =   "OP"
            Top             =   525
            Width           =   450
         End
      End
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu tablas 
      Caption         =   "Tablas"
      Begin VB.Menu tablademonedas 
         Caption         =   "Cotización Dolar y Tipo de Cambio"
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu tabladebancosyvaloresdepago 
         Caption         =   "Tabla de Bancos y Valores de Pago"
      End
   End
End
Attribute VB_Name = "STORPA04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Command11_Click()
  Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text11.TEXT = VDEVU$
        Text10.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call CONECRUN("FLFORMS/" + TRIM$(Text11.TEXT), "Formularios de Impresion")
   Command12.Enabled = True
End Sub

Sub Command13_Click()
    '
    Command13.Enabled = False
    '
    If DERACCE%("CARMEPAG", "Actualizar Tabla de Medios de Pago") < 2 Then
       Command13.Enabled = True
       Exit Sub
    End If
    '
    LU$ = LUDAT$
    Call CIERRARCH("MEPAGO")
    Call CIERRARCH("!MEPAGO")
    If COPYFILE%(LU$ + "MEPAGO.DAT", "C:\FLEXOFT\MEPAGO.WKF") <> 1 Then
     Command13.Enabled = True
        Exit Sub
    End If
    '
    Dim REMEPA&(1000)
    '
360 ATY% = VENTABU%("COMEPAG/NC")
    If ATY% < 0 Then GoTo 390
    '
361 ATX% = 0: J& = 0
    Screen.MousePointer = 11
    For U& = 1 To ULTREG&("!MEPAGO")
       X$ = REGLEIDO$("!MEPAGO", U&)
       CDX% = CVI(Left$(X$, 2))
362    If CDX% > 0 Then
          If CDX% <= 999 Then
             '
             If UCase$(Mid$(X$, 57, 4)) = "CH.3" Then
                 Call MENSERR(24, "Valor '" + Mid$(X$, 57, 4) + "' no Permitido")
                 ATX% = 1
             End If
             If UCase$(Mid$(X$, 57, 4)) = "PROP" Then
                 Call MENSERR(24, "Valor '" + Mid$(X$, 57, 4) + "' no Permitido")
                 ATX% = 1
             End If
             '
             J& = J& + 1
363          Call GRAREG("!MEPAGO", X$, J&)
             If TRIM$(Mid$(X$, 3, 12)) = "" Then
                 Call MENSERR(24, "Registro " + TRIM$(Str$(U&)) + " - Falta Descripción Abreviada.")
                 ATX% = 1
             End If
             If ECOARCH$("ECUECON", TRIM$(Mid$(X$, 45, 12))) = "" Then
                 Call MENSERR(24, "Registro " + TRIM$(Str$(U&)) + " - Cuenta Contable Inexistente o no Válida.")
                 ATX% = 1
             End If
             CCII% = CUEINT%(TRIM$(Mid$(X$, 45, 12)))
             Call REPLA(X$, MKI$(CCII%), 61, 2)
             If CCII% < 1 Or (NUCUEN% > 0 And CCII% > NUCUEN%) Then
                 Call MENSERR(24, "Registro " + TRIM$(Str$(U&)) + " - Cuenta Contable Inexistente o no Válida.")
                 ATX% = 1
             End If
364          If TRIM$(Mid$(X$, 57, 4)) = "" Then
                 Call MENSERR(24, "Registro " + TRIM$(Str$(U&)) + " - Falta Código de M.Pago.")
                 ATX% = 1
             End If
             '
             If REMEPA&(CDX%) < 1 Then REMEPA&(CDX%) = J&
             '
          End If
       End If
365 Next U&
    '
    Call SETULTREG("!MEPAGO", J&)
    '
366 Screen.MousePointer = 1
    If ATX% > 0 Then
        For UII% = 0 To 1000: REMEPA&(UII%) = 0: Next UII%
        Call COMUNI("Datos Incompletos o no Válidos.")
        GoTo 360
    End If
    '
    JJ& = 0
    For UII% = 1 To 1000
      If REMEPA&(UII%) > 0 Then
        If REMEPA&(UII%) <= J& Then
          X$ = REGLEIDO$("!MEPAGO", REMEPA&(UII%))
          JJ& = JJ& + 1
          Call GRAREG("MEPAGO", X$, JJ&)
        End If
      End If
    Next UII%
    Call SETULTREG("MEPAGO", JJ&)
    Call CIERRARCH("MEPAGO")
    Call CIERRARCH("!MEPAGO")
    Call FRESHECHO("MEPAGO")
    '
390 Command13.Enabled = True
    '
End Sub

Private Sub Command14_Click()
   '\\\OT-06-0023-FG-26/01/06
   Command14.Enabled = False
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     GoTo 99
   End If
   '
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
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
99 Command14.Enabled = True
   '\\\OT-06-0023-FG-FIN
End Sub

Private Sub Command2_Click()
    '
    FENUE$ = TRIM$(Text3.TEXT)
    If FENUE$ = "" Then
        FENUE$ = "01/01/80"
    End If
    If FECHANUM(FENUE$) < 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text3.SetFocus
        Exit Sub
    End If
    '
    Call GRACONTROL("", "FORPASTA", Text1.TEXT)
    Call GRACONTROL("", "FORLIFOF", Text6.TEXT)
    Call GRACONTROL("", "FORRETIB", Text9.TEXT)
    '****////******
    'ACA AGREGO EL CONTROL DEL FORMULARIO DE RET DE GANANCIAS
    Call GRACONTROL("", "FORREGAN", Text11.TEXT)
    '****////*******
    Call GRACONTROL("", "CODMORPA", "OP")    ' Text5.TEXT)
    '
    MONUME$ = "AUTONUME": If Option1.Value = True Then MONUME$ = "PREIMPRE"
    Call GRACONTROL("", "NUORPAST", MONUME$)
    '
    MONUME$ = "AUTONUME": If Option10.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NURETGAN", MONUME$)
    '
    TIPAPEL$ = "FORMUCON": If Option4.Value = True Then TIPAPEL$ = "HOJASUEL"
    Call GRACONTROL("", "TIPAORPA", TIPAPEL$)
    '
    AVAHOJA$ = "ENDDOC": If Option6.Value = True Then AVAHOJA$ = "NEWPAGE"
    Call GRACONTROL("ORPASTA", "AVAHOJA", AVAHOJA$)
    '
    UTICHECA$ = "NO": If Check1.Value = 1 Then UTICHECA$ = "SI"
    Call GRACONTROL("ORPASTA", "UTICHECA", UTICHECA$)
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check2.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("ORPASTA", "SUPRINT", SUPRINT$)
    '
    MODORDE$ = "EMISION": MODORECI% = 1
    If Option8.Value = True Then
        MODORDE$ = "VENCIM"
        MODORECI% = 2
    End If
    Call GRACONTROL("ORPASTA", "MODORDE", MODORDE$)
    '
    Call GRACONTROL("", "FENUORPA", FENUE$)
    '
    PREFIREM$ = TRIM$(Text4.TEXT): If PREFIREM$ = "" Then PREFIREM$ = "0001"
    If PREFIREM$ <> "" Then
       If Len(PREFIREM$) <> 4 Then
          Call COMUNI("Prefijo no Válido")
          Text4.SetFocus
          Exit Sub
       End If
       For U& = 1 To 4
          UI% = Asc(Mid$(PREFIREM$, U&))
          If UI% < 48 Or UI% > 57 Then
             Call COMUNI("Prefijo no Válido")
             Text4.SetFocus
             Exit Sub
          End If
       Next U&
    End If
    Call GRACONTROL("", "PREFORPA", PREFIREM$)
    '\\\OT-05-0319-WAR-23/05/05///
    ACUMENS$ = "NO": ACUMERIB% = 0
    If Check3.Value = 1 Then ACUMENS$ = "SI": ACUMERIB% = 1
    Call GRACONTROL("ORPASTA", "ACUMENS", ACUMENS$)
    '\\\OT-05-0319-WAR-FIN///
    Call CIERRARCH("*.*")
    Call BAJALDISCO
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
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command7_Click()
   Call HELPONLINE("COGECOM")
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CONECRUN("FLFORMS/" + TRIM$(Text6.TEXT), "Formularios de Impresion")
    Command8.Enabled = True
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.Picture16.Visible = True
      Me.BackColor = &H8000000F
      Me.Frame1.BackColor = &H8000000F
      Me.Frame2.BackColor = &H8000000F
      Me.Frame3.BackColor = &H8000000F
      Me.Frame4.BackColor = &H8000000F
      Me.Frame5.BackColor = &H8000000F
      Me.Frame6.BackColor = &H8000000F
      Me.Frame7.BackColor = &H8000000F
      Me.Frame8.BackColor = &H8000000F
      Me.Frame9.BackColor = &H8000000F
      Me.Frame10.BackColor = &H8000000F
      Me.Frame11.BackColor = &H8000000F
      Me.Frame12.BackColor = &H8000000F
      Me.Frame13.BackColor = &H8000000F
      
      Me.Check1.BackColor = &H8000000F
      Me.Check2.BackColor = &H8000000F
      Me.Check3.BackColor = &H8000000F
      
      Me.Option1.BackColor = &H8000000F
      Me.Option2.BackColor = &H8000000F
      Me.Option3.BackColor = &H8000000F
      Me.Option4.BackColor = &H8000000F
      Me.Option5.BackColor = &H8000000F
      Me.Option6.BackColor = &H8000000F
      Me.Option7.BackColor = &H8000000F
      Me.Option8.BackColor = &H8000000F
      Me.Option9.BackColor = &H8000000F
      Me.Option10.BackColor = &H8000000F
      MARCOBARRA.Top = 4170
      MARCOBARRA.Visible = True
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If
'&H8000000F
    Text1.TEXT = CONTROL$("", "FORPASTA")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text6.TEXT = CONTROL$("", "FORLIFOF")
    Text7.TEXT = ECOARCH$("INDIFRM", Text6.TEXT)
    Text9.TEXT = CONTROL$("", "FORRETIB")
    Text8.TEXT = ECOARCH$("INDIFRM", Text9.TEXT)
    Text3.TEXT = CONTROL$("", "FENUORPA")
    'agregue estas dos lineas para que cargue
    'el formulario de retencion de ganacias
    Text11.TEXT = CONTROL$("", "FORREGAN")
    Text10.TEXT = ECOARCH$("INDIFRM", Text11.TEXT)
    '
    If CONTROL$("", "NUORPAST") = "AUTONUME" Then Option2.Value = True
    If CONTROL$("", "NURETGAN") = "MANUAL" Then Option10.Value = True
    If CONTROL$("", "TIPAORPA") = "HOJASUEL" Then Option4.Value = True
    If CONTROL$("ORPASTA", "AVAHOJA") = "NEWPAGE" Then Option6.Value = True
    If CONTROL$("ORPASTA", "UTICHECA") = "SI" Then Check1.Value = 1
    If CONTROL$("ORPASTA", "SUPRINT") = "SI" Then Check2.Value = 1
    '\\\OT-05-0319-WAR-23/05/05///
    If CONTROL$("ORPASTA", "ACUMENS") = "SI" Then Check3.Value = 1
    '\\\OT-05-0319-WAR-FIN///
    MODORECI% = 1: If CONTROL$("ORPASTA", "MODORDE") = "VENCIM" Then Option8.Value = True: MODORECI% = 2
    '
    PREFIREM$ = TRIM$(CONTROL$("", "PREFORPA"))
    Text4.TEXT = PREFIREM$
    Text5.TEXT = "OP"
    '
End Sub

Private Sub tablasss_Click()
       
End Sub

Private Sub grabar_Click()
   Call Command2_Click
End Sub

Private Sub salir_Click()
    Call Command1_Click
End Sub

Private Sub tabladebancosyvaloresdepago_Click()
   Call Command13_Click
End Sub

Private Sub tablademonedas_Click()
   Call Command14_Click
End Sub

Private Sub Text1_Change()
    '\\\OT-06-0570-RG-28/12/06///
    Text2 = ECOARCH$("INDIFRM", Text1)
    '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text11_Change()
    '\\\OT-06-0570-RG-28/12/06///
    Text10 = ECOARCH$("INDIFRM", Text11)
    '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text11_DblClick()
    '\\\OT-06-0570-RG-27/12/06///
    Call Command11_Click
    '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub


Private Sub Text5_Click()
    Command2.SetFocus
End Sub

Private Sub Text6_Change()
    '\\\OT-06-0570-RG-28/12/06///
    Text7 = ECOARCH$("INDIFRM", Text6)
    '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text7_CLICK()
    Text6.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Command10_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text9.TEXT = VDEVU$
        Text8.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text8_Click()
    Text9.SetFocus
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call CONECRUN("FLFORMS/" + TRIM$(Text9.TEXT), "Formularios de Impresion")
   Command9.Enabled = True
End Sub

Private Sub Text9_Change()
   '\\\OT-06-0570-RG-28/12/06///
   Text8 = ECOARCH$("INDIFRM", Text9)
   '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text9_DblClick()
   '\\\OT-06-0570-RG-27/12/06///
   Call Command10_Click
   '\\\OT-06-0570-RG-FIN///
End Sub
