VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form IREACOMP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0FFEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Compras"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11850
   Icon            =   "IREACOMP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00D0FFEE&
      Caption         =   "A-B-M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   10760
      TabIndex        =   43
      Top             =   5580
      Width           =   1000
      Begin VB.CommandButton Command10 
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
         Height          =   740
         Left            =   105
         Picture         =   "IREACOMP.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Registra Información de Compra del Item"
         Top             =   380
         Width           =   750
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "IREACOMP.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Alta de Nuevo Item de Stock"
         Top             =   1200
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Entregas Programadas"
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
      Height          =   2175
      Left            =   6510
      TabIndex        =   33
      Top             =   5565
      Width           =   4110
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
         Height          =   285
         Index           =   3
         Left            =   2940
         TabIndex        =   12
         Top             =   1680
         Width           =   960
      End
      Begin VB.CommandButton Command2 
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
         Index           =   3
         Left            =   1890
         TabIndex        =   11
         Top             =   1680
         Width           =   180
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
         Height          =   285
         Index           =   2
         Left            =   2940
         TabIndex        =   10
         Top             =   1260
         Width           =   960
      End
      Begin VB.CommandButton Command2 
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
         Index           =   2
         Left            =   1890
         TabIndex        =   9
         Top             =   1260
         Width           =   180
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
         Height          =   285
         Index           =   1
         Left            =   2940
         TabIndex        =   8
         Top             =   840
         Width           =   960
      End
      Begin VB.CommandButton Command2 
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
         Index           =   1
         Left            =   1890
         TabIndex        =   7
         Top             =   840
         Width           =   180
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
         Height          =   285
         Index           =   0
         Left            =   2940
         TabIndex        =   6
         Top             =   420
         Width           =   960
      End
      Begin VB.CommandButton Command2 
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
         Index           =   0
         Left            =   1890
         TabIndex        =   5
         Top             =   420
         Width           =   180
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
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
         Height          =   285
         Index           =   3
         Left            =   840
         TabIndex        =   42
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
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
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   41
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
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
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   40
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant:"
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
         TabIndex        =   36
         Top             =   480
         Width           =   960
      End
      Begin VB.Label Label5 
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
         Left            =   -525
         TabIndex        =   35
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
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
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   34
         Top             =   420
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Datos de Orden de Compra"
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
      Height          =   2175
      Left            =   210
      TabIndex        =   24
      Top             =   5565
      Width           =   6105
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
         Left            =   4700
         TabIndex        =   4
         Top             =   1335
         Width           =   1275
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
         Left            =   1785
         TabIndex        =   3
         Top             =   720
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
         Height          =   285
         Left            =   1155
         TabIndex        =   1
         Top             =   315
         Width           =   1065
      End
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
         Left            =   5775
         TabIndex        =   2
         Top             =   315
         Width           =   180
      End
      Begin VB.Label Label10 
         Height          =   225
         Left            =   315
         TabIndex        =   46
         Top             =   1050
         Width           =   645
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Left            =   315
         TabIndex        =   45
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label Label7 
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
         TabIndex        =   39
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pr.Unit:"
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
         Left            =   3780
         TabIndex        =   38
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label Label6 
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
         Height          =   285
         Left            =   1155
         TabIndex        =   37
         Top             =   720
         Width           =   645
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
         TabIndex        =   32
         Top             =   380
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
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
         TabIndex        =   31
         Top             =   1380
         Width           =   1905
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
         Left            =   4830
         TabIndex        =   30
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Emisión:"
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
         Left            =   2625
         TabIndex        =   29
         Top             =   375
         Width           =   2115
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
         Left            =   2055
         TabIndex        =   28
         Top             =   720
         Width           =   3915
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
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
         Left            =   105
         TabIndex        =   27
         Top             =   765
         Width           =   960
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
         Top             =   1335
         Width           =   1605
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "U/M:"
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
         TabIndex        =   25
         Top             =   1380
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
         TabIndex        =   14
         Top             =   1335
         Width           =   435
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   630
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
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
      Height          =   4785
      Left            =   210
      TabIndex        =   0
      Top             =   630
      Width           =   10410
   End
   Begin VB.PictureBox Picture1 
      Height          =   435
      Left            =   210
      ScaleHeight     =   375
      ScaleWidth      =   10365
      TabIndex        =   21
      Top             =   210
      Width           =   10430
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   $"IREACOMP.frx":091E
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
         Left            =   0
         TabIndex        =   22
         Top             =   60
         Width           =   10410
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10760
      ScaleHeight     =   75
      ScaleWidth      =   945
      TabIndex        =   19
      Top             =   4725
      Width           =   1000
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2620
      Left            =   10760
      TabIndex        =   16
      Top             =   105
      Width           =   1000
      Begin VB.CommandButton COMMAND3 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "IREACOMP.frx":09B4
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Enviar Información de Compras por Mail"
         Top             =   1735
         Width           =   750
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "IREACOMP.frx":127E
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   265
         Width           =   750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "IREACOMP.frx":13C8
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1000
         Width           =   750
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00D0FFEE&
      Caption         =   "A.Maes."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1880
      Left            =   10760
      TabIndex        =   15
      ToolTipText     =   "Maestro de Items de Stock"
      Top             =   2780
      Width           =   1000
      Begin VB.CommandButton Command9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "IREACOMP.frx":16D2
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Maestro de Proveedores"
         Top             =   280
         Width           =   750
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
         Height          =   740
         Left            =   105
         Picture         =   "IREACOMP.frx":19DC
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Alta de Nuevo Item de Stock"
         Top             =   1015
         Width           =   750
      End
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   1470
      TabIndex        =   47
      Top             =   105
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Line Line3 
      X1              =   210
      X2              =   11865
      Y1              =   5410
      Y2              =   5410
   End
   Begin VB.Line Line1 
      X1              =   10635
      X2              =   10635
      Y1              =   -105
      Y2              =   5460
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10280
      Picture         =   "IREACOMP.frx":1B26
      Top             =   4900
      Width           =   2010
   End
End
Attribute VB_Name = "IREACOMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LISTACTIVA&

Private Sub Command1_Click()
   If List1.ListCount > 0 Then
      Call SENDCOMPRA
   End If
   '
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command10_Click()
   '
   Command10.Enabled = False
   HOII% = HOY(HOS$)
   '
   If Val(Text1) < 1 Or Val(Text1) > 99999 Then
     Call MENSERR(24, "Numero de O/C no Válido")
     GoTo 99
   End If
   '
   If Val(Label6) < 1 Then
     Call MENSERR(24, "Falta Proveedor")
     GoTo 99
   End If
   '
   If ECOARCH$("PROVED1", MKI$(Val(Label6))) = "" Then
     Call MENSERR(24, "Falta Proveedor")
     GoTo 99
   End If
   '
   If FECHANUM(Label8) < 1 Then
     Call MENSERR(24, "Fecha de Emisión no Válida")
     GoTo 99
   End If
   '
   For KKI% = 0 To 3
     If Val(Text3(KKI%)) > 0 Then
       If FECHANUM(Label4(KKI%)) < HOII% Then
         Call MENSERR(24, "Fecha Entrega Solicitada no Válida")
         GoTo 99
       End If
     End If
   Next KKI%
   '
   Call AGRERECOM
   '
99 Command10.Enabled = True
   '
End Sub

Sub AGRERECOM()
   '
   TTXX$ = "OCOMPRA "
   TTXX$ = TTXX$ + AJUSTD$(Text1, 8)  ' NUMERO DE ORDEN DE COMPRA
   TTXX$ = TTXX$ + AJUSTI$(Label8, 8) ' FECHA DE EMISION
   TTXX$ = TTXX$ + AJUSTD$(Label6, 8) ' NUMERO DE PROVEEDOR
   TTXX$ = TTXX$ + AJUSTD$(Label11, 48) ' R.SOCIAL PROVEEDOR
   TTXX$ = TTXX$ + AJUSTI$(Label14, 16) ' codigo articulo
   TTXX$ = TTXX$ + AJUSTD$(Label2, 8) ' CODIGO INTERNO MATERIAL
     Text2 = REPLACAR$(Text2, ",", ".")
   TTXX$ = TTXX$ + AJUSTD$(Text2, 16) ' PRECIO UNITARIO
   '
   For KKI% = 0 To 3
     If Val(Text3(KKI%)) > 0 Then
       TTXX$ = TTXX$ + AJUSTI$(Label4(KKI%), 8) + AJUSTD$(Text3(KKI%), 12)
     End If
   Next KKI%
   '
   ILISTA& = List2.ListIndex
   List1.AddItem TTXX$
   List2.RemoveItem (Val(Label10))
   Label10 = ""
   '
   Label14 = ""
   Label7 = ""
   Label16 = ""
   Label2 = ""
   Label8 = ""
   Text2 = ""
   '
   For U& = 0 To 3
     Label4(U&) = ""
     Text3(U&) = ""
   Next U&
   '
   On Error Resume Next
   List2.ListIndex = ILISTA&
   List2.SetFocus
   On Error GoTo 0
   '
End Sub

Private Sub Command2_Click(Index As Integer)
   Call SELECHO("SELFECHA")
   If VALACT1$("SELFECHA") <> "" Then
       Label4(Index) = VALACT1$("SELFECHA")
       On Error Resume Next
       Text3(Index).SetFocus
       On Error GoTo 0
     Else
       For KKI% = Index To 3
         Label4(Index) = ""
         Text3(Index) = ""
       Next KKI%
       On Error Resume Next
       Command10.SetFocus
       On Error GoTo 0
   End If
End Sub

Private Sub Command21_Click()
   Command1.Enabled = False
   '
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
   Call REPLA(VDEF$, "U.", 109, 2)
   Call REPLA(VDEF$, Chr$(1), 93, 1)
   '
   OKX% = 0
   Call ALTACODIGO(VDEF$, OKX%)
   '
   Command1.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call SELECHO("SELFECHA")
   If VALACT1$("SELFECHA") <> "" Then
     Label8 = VALACT1$("SELFECHA")
     On Error Resume Next
     Command6.SetFocus
     On Error GoTo 0
   End If
   Command4.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call SELECHO("PROVED1")
   If VALACT1$("PROVED1") <> "" Then
     Label6 = VALACT1$("PROVED1")
     On Error Resume Next
     Text2.SetFocus
     On Error GoTo 0
   End If
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
        '
   Command7.Enabled = False
   Label14 = ""
   Label7 = ""
   Label16 = ""
   Label2 = ""
   Label8 = ""
   Text2 = ""
   '
   For U& = 0 To 3
     Label4(U&) = ""
     Text3(U&) = ""
   Next U&
   '
   Label10 = "0"
   Command7.Enabled = True
        '
End Sub


Private Sub Command9_Click()
   Command9.Enabled = False
   Call CONECRUN("AMAPRO", "Maestro de Proveedores")
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
   '
   Call AUTOEXEC
   '
   Show
   Refresh
   '
   OPX% = COMALTER%("Puede Refrescar los Mails Recibidos\\Refrescar Recepción\Continuar")
   If OPX% = 1 Then
       MAPISession1.DownLoadMail = True
     Else
       MAPISession1.DownLoadMail = False
   End If
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   DoEvents
   '
   MAPIMessages1.Fetch
   CAMENSA& = MAPIMessages1.MsgCount
   '
   List1.Clear
   For U& = CAMENSA& To 1 Step -1
     Call TRACE(20, U&, CAMENSA&)
     MAPIMessages1.MsgIndex = U& - 1
     ASUNTO$ = MAPIMessages1.MsgSubject
     If InStr(ASUNTO$, "Necesidades de Compras (") > 0 Then
       CLA1$ = Mid$(ASUNTO$, 25, 6)
       CLA2$ = Mid$(ASUNTO$, 32, 6)
       If CLA2$ = CONTRACLA$(CLA1$) Then
         TEXTO$ = MAPIMessages1.MsgNoteText
         List2.Clear
         While Len(TEXTO$) > 0
           PPXX% = InStr(TEXTO$, Chr$(13) + Chr$(10))
           If PPXX% < 1 Then PPXX% = 1 + Len(TEXTO$)
           TTXX$ = Left$(TEXTO$, PPXX% - 1)
           TEXTO$ = Mid$(TEXTO$, PPXX% + 2)
           If TRIM$(TTXX$) <> "" Then List2.AddItem TTXX$
         Wend
         GoTo 99
       End If
     End If
     '
89 Next U&
   '
   Call COMUNI("No se Encontraron\Necesidades de Compras")
   '
99 MAPISession1.Action = 2  ' !!!
   Command5.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If List1.ListCount > 0 Then
      Call SENDCOMPRA
   End If
   '
   Cancel = 0
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label4_DblClick(Index As Integer)
   Call SELECHO("SELFECHA")
   If VALACT1$("SELFECHA") <> "" Then
     Label4(Index) = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Label6_Change()
   Label11 = ""
   If Val(Label6) > 0 Then
     Label11 = RASOCLI$((-1) * Val(Label6))
   End If
End Sub

Private Sub Label7_Change()
   Label9 = ECOARCH$("VALIMONE", Chr$(Val(Label7)))
End Sub

Private Sub Label7_DblClick()
   Call SELECHO("VALIMONE")
   If VALACT1$("VALIMONE") <> "" Then
     Label7 = VALACT1$("VALIMONE")
   End If
End Sub

Sub ALTACODIGO(VDEF$, OKX%)
   '
   Dim BUFERCOD As String * 2048
   '
10 OBX$ = OBJPLA$("ALTACODIGO", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   CODINUE$ = TRIM$(Mid$(OBX$, 77, 16))
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 10
   End If
   '
   If Len(CODINUE$) > 15 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que 15 Caracteres.")
      GoTo 10
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 10
   End If
   '
   If TRIM$(Mid$(OBX$, 5, 64)) = "" Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Descripcion")
      GoTo 10
   End If
   '
   TIMA% = Asc(Mid$(OBX$, 93, 1))
   If ECOARCH$("TATIMAT", Chr$(TIMA%)) = "" Then
      Call MENSERR(24, "Tipo de Material no Válido")
      GoTo 10
   End If
   '
   UNID$ = Mid$(OBX$, 109, 2)
   If ECOARCH$("UNIMED", UNID$) = "" Then
      Call MENSERR(24, "Unidad de Medida no Válida")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   PUNCODIG$ = PUNCODIG$ + AJUSTI$(CODINUE$, 16)
   PPCC$ = PUNCODIG$
   JI% = 1
   While Len(PPCC$) > 2048
      PPCC$ = Mid$(PPCC$, 2049)
      JI% = JI% + 1
   Wend
   While Len(PPCC$) < 2048
      PPCC$ = PPCC$ + Chr$(0)
   Wend
   '
   NX% = FILEOPEN%(LUDAT$ + "PUNCODIG.DAT", 0, 2048)
   BUFERCOD$ = PPCC$
   Put #NX%, JI%, BUFERCOD$
   Close #NX%
   '
   REMAS& = NUMAS% + 2
   Call GRAREG("MASTER", OBX$, REMAS&)
   NMSI% = NUMAS% + 1
   X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(NMSI%), 1, 2)
   Call GRAREG("MASTER", X$, 1)
   Call CIERRARCH("MASTER")
   '
   EL1$ = REGBLAN$("ECOMAST")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, Mid$(OBX$, 5, 64), 17, 46)
   Call REPLA(EL1$, MKI$(NMSI%), 63, 2)
   '
   URECO& = ULTREG&("ECOMAST")
   While URECO& > 0
     XXX$ = REGLEIDO$("ECOMAST", URECO&)
     If TRIM$(XXX$) <> "" Then GoTo 25
     URECO& = URECO& - 1
   Wend
   '
25 LOTOTE& = 64 * URECO&
   Call CIERRARCH("ECOMAST")
   FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 2048)
   Dim EL As String * 2048
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   '
   For U& = 1 To Len(TTXX$) Step 64
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 60
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
60 JJ& = 0: LOTEGRA& = Len(TTXX$)
   TTXX$ = TTXX$ + String$(2048, 0)
   For U& = 1 To LOTEGRA& Step 2048
      EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   NVII& = 1 + NUINV%
   EL1$ = REGBLAN$("INVERT")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, MKI$(NMSI%), 17, 2)
   LOTOTE& = 18 * ULTREG&("INVERT")
   If LOTOTE& < 18 * NVII& Then LOTOTE& = 18 * NVII&
   FIECO% = FILEOPEN%(LUDAT$ + "INVERT.DAT", 0, 2048)
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   While Len(TTXX$) < LOTOTE&: TTXX$ = TTXX$ + Chr$(0): Wend
   '
   NVIJ% = NVII&
   Mid$(TTXX$, 17, 2) = MKI$(Int(NVIJ%))
   For U& = 19 To Len(TTXX$) Step 18
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 70
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
70 TTXX$ = TTXX$
   JJ& = 0
   For U& = 1 To Len(TTXX$) Step 2048
      EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   Call CIERRARCH("ECOMAST")
   Call FRESHECHO("MASTER")
   Call BAJALDISCO
   OKX% = 1
   '
99 Screen.MousePointer = 1

End Sub
'
Private Sub Label8_Click()

End Sub

Private Sub Label8_DblClick()
   Call SELECHO("SELFECHA")
   If VALACT1$("SELFECHA") <> "" Then
     Label8 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub List2_DblClick()
   '
   HOII% = HOY(HOS$)
   
   TTXX$ = List2.TEXT
   TTYY$ = Mid$(List2.TEXT, 101)
   Label14 = TRIM$(Left$(TTXX$, 16))
   Label7 = TRIM$(Mid$(TTXX$, 17, 64))
   Label16 = TRIM$(Mid$(TTXX$, 81, 4))
   Label2 = TRIM$(Mid$(TTXX$, 95, 6))
   Label8 = HOS$
   '
   For U& = 1 To Len(TTYY$) Step 20
     Index = (U& - 1) / 20
     If Index <= 3 Then
       Label4(Index) = Mid$(TTYY$, U& + 1, 8)
       Text3(Index) = TRIM$(Mid$(TTYY$, U& + 9, 10))
     End If
   Next U&
   '
   Label10 = List2.ListIndex
   Label10.Refresh
   '
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
   '
End Sub

Sub SENDCOMPRA()
   '
   DESTIMAI$ = LCase$(CONTROL$("INCOMPRA", "DESTINCO"))
10 DESTIMAI$ = TRIM$(LCase$(InputBox("Direccion Electronica Destino", "Transferir Información de Compras", DESTIMAI$)))
   If DESTIMAI$ = "" Then Exit Sub
   '
   If InStr(DESTIMAI$, "@") < 2 Then
     Call MENSERR(24, "Destino no Válido")
     GoTo 10
   End If
   '
   Call GRACONTROL("INCOMPRA", "DESTINCO", DESTIMAI$)
   '
   Screen.MousePointer = 11
   TTYY$ = ""
   For U& = 1 To List1.ListCount
     TTYY$ = TTYY$ + List1.List(U& - 1)
     If U& < List1.ListCount Then
       TTYY$ = TTYY$ + Chr$(13) + Chr$(10)
     End If
   Next U&
   '
   IDXS$ = RANDSTRING$(-6) ' NEGATIVO PARA FORZAR VALORES NUMERICOS
   IDENMESS$ = IDXS$ + "." + CONTRACLA$(IDXS$)
   '
   MAPISession1.DownLoadMail = False
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   MAPIMessages1.MsgIndex = -1
   MAPIMessages1.MsgSubject = "Informacion de Compras (" + IDENMESS$ + ")"
   MAPIMessages1.MsgNoteText = TTYY$
   MAPIMessages1.RecipIndex = 0
   MAPIMessages1.RecipType = 1
   MAPIMessages1.RecipAddress = DESTIMAI$
   '
   On Error Resume Next
   MAPIMessages1.Action = 3   ' vbMessageSend
   If Err Then
     MsgBox "Ocurrio Error " & Err
     Resume 90
   End If
   '
   Call COMUNI("Informacion de Compras Transferida.")
   '
90 On Error GoTo 0
   Screen.MousePointer = 1
   MAPISession1.Action = 2  ' !!!
   DoEvents
   '

End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
     Call List2_DblClick
   End If
End Sub

Private Sub Text1_GotFocus()
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
     On Error Resume Next
     Command4.SetFocus
     On Error GoTo 0
   End If
End Sub

Private Sub Text2_GotFocus()
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2.TEXT)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
     On Error Resume Next
     Command2(0).SetFocus
     On Error GoTo 0
   End If
End Sub

Private Sub Text3_GotFocus(Index As Integer)
   Text3(Index).SelStart = 0
   Text3(Index).SelLength = Len(Text3(Index).TEXT)
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii% = 13 Then
       On Error Resume Next
       If Index >= 3 Then
           Command10.SetFocus
         Else
           Command2(Index + 1).SetFocus
       End If
       On Error GoTo 0
       Exit Sub
     ElseIf KeyAscii% = 27 Then
       On Error Resume Next
       Command10.SetFocus
       On Error GoTo 0
   End If
End Sub
