VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CONCITC07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depósito Cupones de Crédito"
   ClientHeight    =   8475
   ClientLeft      =   5070
   ClientTop       =   300
   ClientWidth     =   9090
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
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
      Left            =   4755
      TabIndex        =   87
      Top             =   4725
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
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
      Left            =   4755
      TabIndex        =   86
      Top             =   4410
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
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
      Left            =   4755
      TabIndex        =   85
      Top             =   5040
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
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
      Left            =   4755
      TabIndex        =   84
      Top             =   5670
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
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
      Index           =   6
      Left            =   4755
      TabIndex        =   83
      Top             =   6300
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
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
      Index           =   7
      Left            =   4755
      TabIndex        =   82
      Top             =   6615
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
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
      Left            =   4755
      TabIndex        =   81
      Top             =   5355
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
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
      Left            =   4755
      TabIndex        =   80
      Top             =   5985
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
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
      Index           =   8
      Left            =   4755
      TabIndex        =   79
      Top             =   6930
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
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
      Index           =   9
      Left            =   4755
      TabIndex        =   78
      Top             =   7245
      Visible         =   0   'False
      Width           =   1560
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
      Index           =   9
      Left            =   1545
      TabIndex        =   67
      Top             =   7240
      Width           =   200
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
      Index           =   8
      Left            =   1545
      TabIndex        =   63
      Top             =   6935
      Width           =   200
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   6330
      TabIndex        =   69
      Top             =   7240
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
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
      Index           =   9
      Left            =   75
      TabIndex        =   68
      Top             =   7240
      Width           =   1515
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   6330
      TabIndex        =   65
      Top             =   6935
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
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
      Index           =   8
      Left            =   75
      TabIndex        =   64
      Top             =   6935
      Width           =   1515
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1800
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   61
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6480
      TabIndex        =   57
      Text            =   " "
      Top             =   480
      Width           =   1530
   End
   Begin VB.TextBox Text22 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4320
      MaxLength       =   6
      TabIndex        =   56
      Top             =   120
      Width           =   930
   End
   Begin VB.CommandButton BOTREC 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   8280
      Picture         =   "CONCITC07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Grabar - Guardar Modificaciones"
      Top             =   7200
      Width           =   780
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
      Index           =   5
      Left            =   1545
      TabIndex        =   20
      Top             =   5985
      Width           =   200
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
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
      Left            =   75
      TabIndex        =   19
      Top             =   5985
      Width           =   1515
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   6330
      TabIndex        =   21
      Top             =   5985
      Width           =   1695
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
      Index           =   3
      Left            =   1545
      TabIndex        =   14
      Top             =   5355
      Width           =   200
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
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
      Left            =   75
      TabIndex        =   13
      Top             =   5355
      Width           =   1515
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   6330
      TabIndex        =   15
      Top             =   5355
      Width           =   1695
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
      Index           =   7
      Left            =   1545
      TabIndex        =   26
      Top             =   6615
      Width           =   200
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
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
      Index           =   7
      Left            =   75
      TabIndex        =   25
      Top             =   6615
      Width           =   1515
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   6330
      TabIndex        =   27
      Top             =   6615
      Width           =   1695
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
      Index           =   6
      Left            =   1545
      TabIndex        =   23
      Top             =   6300
      Width           =   200
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
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
      Index           =   6
      Left            =   75
      TabIndex        =   22
      Top             =   6300
      Width           =   1515
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   6330
      TabIndex        =   24
      Top             =   6300
      Width           =   1695
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
      Index           =   4
      Left            =   1545
      TabIndex        =   17
      Top             =   5670
      Width           =   200
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
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
      Left            =   75
      TabIndex        =   16
      Top             =   5670
      Width           =   1515
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   6330
      TabIndex        =   18
      Top             =   5670
      Width           =   1695
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
      Index           =   2
      Left            =   1545
      TabIndex        =   11
      Top             =   5040
      Width           =   200
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
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
      Left            =   75
      TabIndex        =   10
      Top             =   5040
      Width           =   1515
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   6330
      TabIndex        =   12
      Top             =   5040
      Width           =   1695
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
      Index           =   0
      Left            =   1545
      TabIndex        =   5
      Top             =   4410
      Width           =   200
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
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
      Left            =   75
      TabIndex        =   4
      Top             =   4410
      Width           =   1515
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   6330
      TabIndex        =   6
      Top             =   4410
      Width           =   1695
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
      Index           =   1
      Left            =   1545
      TabIndex        =   8
      Top             =   4725
      Width           =   200
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
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
      Left            =   75
      TabIndex        =   7
      Top             =   4725
      Width           =   1515
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   6330
      TabIndex        =   9
      Top             =   4725
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   1080
      TabIndex        =   43
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
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
      Left            =   2300
      TabIndex        =   2
      Top             =   510
      Width           =   200
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
      Left            =   7830
      TabIndex        =   1
      Top             =   135
      Width           =   200
   End
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
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   135
      Width           =   200
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H002C3D46&
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   8160
      ScaleHeight     =   7065
      ScaleWidth      =   1170
      TabIndex        =   32
      Top             =   0
      Width           =   1200
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H002B9973&
         Height          =   195
         Left            =   120
         ScaleHeight     =   135
         ScaleWidth      =   585
         TabIndex        =   76
         Top             =   6360
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   77
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command13 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   120
         Picture         =   "CONCITC07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Detalle de Liquidacion"
         Top             =   3480
         Width           =   680
      End
      Begin VB.CommandButton BOTEXIT 
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
         Height          =   700
         Left            =   100
         Picture         =   "CONCITC07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   240
         Width           =   680
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Picture         =   "CONCITC07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Búsqueda por Número, Banco, Importe, etc."
         Top             =   1800
         Width           =   680
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Sort"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Picture         =   "CONCITC07.frx":1AC4
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Conmuta Ordenamiento por Fecha / Importe"
         Top             =   960
         Width           =   680
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Anul."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Picture         =   "CONCITC07.frx":1C0E
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Anula Liquidacion Procesada"
         Top             =   2640
         Width           =   680
      End
   End
   Begin VB.ListBox LICHESEL 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   945
      Width           =   7905
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "CONCITC07.frx":1F18
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CONCITC07.frx":214C
      TabIndex        =   62
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0095C5E1&
      Height          =   855
      Left            =   240
      TabIndex        =   71
      Top             =   7520
      Width           =   7905
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Neto a Acreditar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4080
         TabIndex        =   75
         Top             =   465
         Width           =   1860
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6090
         TabIndex        =   74
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   73
         Top             =   460
         Width           =   1005
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   72
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Importes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7380
      TabIndex        =   89
      Top             =   4200
      Width           =   885
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "N° Comprobante"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4785
      TabIndex        =   88
      Top             =   4200
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   9
      Left            =   1755
      TabIndex        =   70
      Top             =   7245
      Width           =   3405
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   1755
      TabIndex        =   66
      Top             =   6930
      Width           =   3405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "$ Liq.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5685
      TabIndex        =   58
      Top             =   615
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Liquid."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3825
      TabIndex        =   55
      Top             =   255
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   1755
      TabIndex        =   54
      Top             =   5985
      Width           =   3405
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   1755
      TabIndex        =   53
      Top             =   5355
      Width           =   3405
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   7
      Left            =   1755
      TabIndex        =   52
      Top             =   6615
      Width           =   3405
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   1755
      TabIndex        =   51
      Top             =   6300
      Width           =   3405
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   1755
      TabIndex        =   50
      Top             =   5670
      Width           =   3405
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   1755
      TabIndex        =   49
      Top             =   5040
      Width           =   3405
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   1755
      TabIndex        =   48
      Top             =   4410
      Width           =   3405
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   1755
      TabIndex        =   47
      Top             =   4725
      Width           =   3405
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   225
      Left            =   5355
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   225
      Left            =   3675
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6960
      TabIndex        =   44
      Top             =   135
      Width           =   975
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   840
      TabIndex        =   42
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   840
      TabIndex        =   41
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6315
      TabIndex        =   40
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2520
      TabIndex        =   39
      Top             =   510
      Width           =   2175
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1560
      TabIndex        =   38
      Top             =   135
      Width           =   2175
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Cupones Seleccionados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3435
      TabIndex        =   37
      Top             =   3945
      Width           =   2820
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuentas de Gastos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      TabIndex        =   36
      Top             =   4215
      Width           =   1725
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Banco"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -240
      TabIndex        =   35
      Top             =   630
      Width           =   1005
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6240
      TabIndex        =   34
      Top             =   240
      Width           =   645
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tarjeta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -240
      TabIndex        =   33
      Top             =   225
      Width           =   1005
   End
End
Attribute VB_Name = "CONCITC07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODORDE%
Dim CtaRetIva$
Dim CtaRetIbrt$

Private Sub BOTEXIT_Click()
    Call LIMPCONCITC    'Limpia el formulario
    Label19 = ""
    Hide
End Sub

Private Sub BOTREC_Click()
    
    If Label17 = "" Then
        Call MENSERR(24, "Codigo de Tarjeta Invalido")
        Exit Sub
    End If
    If FECHANUM(Label10.Caption) > HOY(HO$) Then
        Call MENSERR(24, "Fecha Invalida")
        Exit Sub
    End If
    FEDEPO% = FECHANUM(Label10.Caption)
    FEACR% = FECHANUM(Label10.Caption)
    '
    If Label18.Caption = "" Then
        Call MENSERR(24, "Falta Seleccionar Banco")
        Exit Sub
    End If
    If XVALO(Label12) <= 0 Then
        Call MENSERR(24, "Falta Seleccionar Cupones")
        Exit Sub
    End If
    '
    For KKI% = 0 To 9
       If XVALO(Text2(KKI%)) > 0 Then
         If CUECOINT%(Text6(KKI%)) = 0 Then
            Call MENSERR(24, "Cuenta Contable de Gasto Invalida")
            Exit Sub
         End If
       End If
    Next KKI%
    '
    If XVALO(Label5) <= 0 Then
        Call MENSERR(24, "Importe Neto a Acreditar debe ser Mayor a Cero")
        Exit Sub
    End If
    '
    NROLIQ$ = TRIM$(Text22)
    If NROLIQ$ = "" Then
       Call MENSERR(24, "Falta Número de Liquidación ")
       Exit Sub
    End If
    '
    If XVALO(CantRegistros&("Cajabanc", "codicom_lar LIKE 'TC%' and Codicom_cor = '" & NROLIQ$ & "'")) > 0 Then
       Call MENSERR(24, "Número de Liquidación yá Ingresado")
       Exit Sub
    End If
    '
    If XVALO(Text1) < 0.005 Then
       Call MENSERR(24, "Falta Ingreso de Importe de Liquidacion")
       Exit Sub
    End If
    '
    VALOR# = XVALO(Label16)
    If Abs(VALOR#) > 0.005 Then
       Call MENSERR(24, "No Coinciden Valores\de Liquidación y Neto")
       Exit Sub
    End If
        '
    DIAX$ = DSEM$(FECHANUM(Label10.Caption))
    If COMALTER%("Fecha de Presentación de Cupones\Correspondiente a " & DIAX$ & " " & Label10.Caption & "'\\Modificar\Continuar") <> 2 Then: Exit Sub
    ' rutina de grabacion
    HOII% = HOY(HOS$)
    NRECIB0$ = "TC-" + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2) + "-"
    For IKI% = 0 To 99
       IXI$ = TRIM$(Str$(IKI%))
       While Len(IXI$) < 2: IXI$ = "0" + IXI$: Wend
       NRECIB$ = NRECIB0$ + IXI$
       YAESTAX$ = VALOBADA("CAJABANC", "CodiCom_Lar", "TipoMovim = 'CONCITC' and CodiCom_Lar = '" & NRECIB$ & "'")
       If YAESTAX$ = "" Then GoTo 25
    Next IKI%
    Call MENSERR(24, "Imposible Registrar")
    Exit Sub
25  CoComCor$ = NROLIQ$
    CoComLar$ = NRECIB$
    CodiEmp% = 0
    CodiSuc% = SUC%
    NumeClie = 0
    NumeProv% = 0
    FechEmi = CVDAT(FEDEPO%)
    '
    Referen$ = "Conciliacion T.Cred. " + Label17
    FeContab = CVDAT(FEDEPO%)
    FeAcredi = CVDAT(FEACR%)
    NumeTal = 0
    NumeCom = Val(Mid$(NRECIB$, 4, 6))
    Observa$ = "Conciliacion T.Cred. " + Label17
    '
    SUMACHESEL# = 0
    FEX = FEDEPO%: HOS$ = FECHATEX$(FEX)
    Call APERBASDAT("CABAN")
    '
    Dim CajaBanc2 As ADODB.Recordset

    For I& = 1 To LICHESEL.ListCount
      If LICHESEL.Selected(I& - 1) = True Then
         X$ = LICHESEL.LIST(I& - 1)
         BANUFE$ = Left$(X$, 40)
         BACH$ = TRIM$(Left$(X$, 16))
         NCHE$ = Mid$(X$, 19, 12)
         IPASEx$ = TRIM$(Mid$(X$, 41, 12))
         REGBA& = Val(Mid$(X$, 65, 8))
         CODTAR$ = Label14
         '
         SQLX$ = "SELECT * FROM CAJABANC "
         SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'CHRECI' "
         SQLX$ = SQLX$ + " AND (OPCIMOVIM = 'TC.' OR OPCIMOVIM = 'TC')"  'Selecciona los pagos con T/C
         SQLX$ = SQLX$ + " AND VARIMOVIM = '" & CODTAR$ & "'"      'codigo de la T/C
         SQLX$ = SQLX$ + " AND (MARCONSI IS NULL OR MARCONSI = '')"                   'que no se ha ingresado
         SQLX$ = SQLX$ + "AND DescriMov = '" & BANUFE$ & "' "
         SQLX$ = SQLX$ + "AND ValAbsComp = " & IPASEx$ & " "
         SQLX$ = SQLX$ + "AND NumeComp = " & REGBA& & " "
         Set CajaBanc2 = New ADODB.Recordset
         CajaBanc2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
         On Error Resume Next
         With CajaBanc2
           .MoveFirst
           If Err < 1 Then
              On Error GoTo 0
              SUMACHESEL# = SUMACHESEL# + XVALO(IPASEx$)
              !FEDEPOSENT = CVDAT(FEDEPO%)
              !Document = NRECIB$
              !DESTICHEQ = TRIM$(Label18)
              !REFESALCHE = Left$(Referen$, 48)
              !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
              !UsuaSalChe = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
              !MARCONSI = "*"                 ' Marca DE CONCILIACION
              !IdenBaja = IDACTU$             ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
              .Update
           End If
         End With
         '
      End If
    Next I&
    NOCLOSE% = 0
    Set CajaBanc2 = Nothing
    '
    CUETA$ = Label11
    CuentaCo% = CUEINT%(CUETA$)
    TiMovi = "CONCITC"            ' DEPOSITO BANCARIO
    OpMovi = TRIM$(Label14)       ' MECOBRA DE LA TARJETA ELEGIDA
    VaMovi = ""                   '
    Descrip = "Conciliacion T.Cred. " + Label17
    ImpoDeb# = (-1) * SUMACHESEL#
    StatInit = 0
    Call GRACABAN
    
    Call REPLA(TX$, "Conc. T.Cred. ", 1, 14)
    Call REPLA(TX$, Label17, 15, 12)
    For KKI% = 0 To 9
      If XVALO(Text2(KKI%)) >= 0.005 Then
        CUEBAX$ = Text6(KKI%)
        CuentaCo% = CUEINT%(CUEBAX$)
        TiMovi = "GASTC"         ' GASTO TARJETA DE CREDIRO
        OpMovi = "1"             ' MECOBRA DE LA TARJETA ELEGIDA
        VaMovi = ""
        '
        On Error Resume Next
        Erase CADENATEX$
        On Error GoTo 0
        Call TEXTOLOTES(TRIM$(Text16(KKI%)), ";")
        
        On Error Resume Next
        For I& = 1 To UBound(CADENATEX$)
          Select Case I&
             Case 1
               TIPOCOMPRO$ = Left$(CADENATEX(I&), 1)
             Case 2
               COMPROB$ = CADENATEX(I&)
             Case 3
               LETR$ = CADENATEX(I&)
             End Select
        Next I&
        On Error GoTo 0
        DATOS1$ = TIPOCOMPRO$ + ";" + COMPROB$ + ";" + LETR$
        Descrip = TX$ + DATOS1$  'AGREGA EL NUMERO DE COMPROBANTE (SE AGREGA PARA GENERAR EL AR
        ImpoDeb# = XVALO(Text2(KKI%))
        SUMACHESEL# = SUMACHESEL# - ROUND#(ImpoDeb#)
        StatInit = 0
        Call GRACABAN
      End If
    Next KKI%
    '
    CUEBAX$ = Label20
    CuentaCo% = CUEINT%(CUEBAX$)
    TiMovi = "CONCITC"            ' DEPOSITO BANCARIO
    OpMovi = TRIM$(Label14)       ' MECOBRA DE LA TARJETA ELEGIDA
    VaMovi = ""                   '
    Descrip = "Conciliacion T.Cred. " + Label17
    ImpoDeb# = SUMACHESEL#
    StatInit = 0
    Call GRACABAN
    '
GoTo 299
    '\\\OT-05-0775-WAR-07/07/05///
    'IMPRESION DE DEPOSITO BANCARIO
      FORIPRE$ = TRIM$(CONTROL$("", "FORDEPBA"))
237   If TRIM$(FORIPRE$) <> "" Then
         LU$ = LUDAT$
         CODFOR$ = TRIM$(UCase$(FORIPRE$))
         '
         YAC = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
         MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
         MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
         '
         INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
         If INTERLI < 1 Then
            Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
            INTERLI = 1
         End If
         '
         CACOPI% = Val(ATRIFORM$(CODFOR$, "CAN-COP"))
         If CACOPI% < 1 Or CACOPI% > 6 Then
            Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
            CACOPI% = 1
         End If
         '
         Screen.MousePointer = 11
         FEX = FEDEPO%
         FECHDOC$ = FECHATEX(FEX)
         NUMEDOC$ = NRECIB$
         TIPODOC$ = "Deposito Bancario"
         FEHORE$ = Format(Now, " dd/mm/yy  hh:mm")
         '
         CODPARAM$(1) = "NUME-COM" 'NUMERO DE DEPOSITO BANCARIO
         CODPARAM$(2) = "FECH-EMI" 'FECHA DE EMISION
         CODPARAM$(3) = "COD-CUEN" 'CUENTA CONTABLE-BANCO
         CODPARAM$(4) = "DEN-CUEN" 'DESCRIPCION CUENTA
         CODPARAM$(5) = "REF-ASIE" 'REFERENCIA
         CODPARAM$(6) = "TIE-TOTA" 'PLAZO DE DEPOSITO /CLEARING
         CODPARAM(7) = "F-PLANO" 'FECHA Y HORA REAL
         CODPARAM(8) = "IMP-TOTA" 'SUMA TOTAL DE IMPORTES CHEQUES
    
         CAPARDOC% = 8
    
         DOCPARAM(1) = NUME1
         DOCPARAM(2) = FECHDOC
         DOCPARAM(3) = CUEBA$
         DOCPARAM(4) = ECOARCH("ECUECON", CUEBA$) 'DESTICHE$
         DOCPARAM(5) = REFE$
         DOCPARAM(6) = PLADEPO$
         DOCPARAM(7) = FEHORE$
         '
         CAITAB1% = 0
         CAITCRU% = 0
         '
         'AQUI PONER LOS CODIGOS DE COLUMNAS A LISTAR
         CODTABU1$(1) = "REF-MAT1" 'BANCO
         CODTABU1$(2) = "REF-MAT2" 'NUMERO
         CODTABU1$(3) = "DES-MAT" 'OBSERVACIONES
         CODTABU1$(4) = "SUB-TOTA" 'IMPORTE
         CODTABU1$(5) = "FECH-REC" 'FECHA DE ACREDITACION
         CACOLTAB1% = 5
         '
         'ACA PONGO EL DETALLE
         SUMACHESEL# = 0
         FEX = FEDEPO%: HOS$ = FECHATEX$(FEX)
         Call APERBASDAT("CABAN")
         For I& = 1 To LICHESEL.ListCount
           If LICHESEL.Selected(I& - 1) = True Then
             X$ = LICHESEL.LIST(I& - 1)
             BANUFE$ = Left$(X$, 40)
             BAN1$ = TRIM(Mid(BANUFE, 1, 19))
             NUMCHE1$ = TRIM(Mid(BANUFE, 19, 14))
             FECHE1$ = TRIM(Mid(BANUFE, 31, 12))
             BACH$ = TRIM$(Left$(X$, 16))
             NCHE$ = Mid$(X$, 19, 12)
             IPASE# = Val(Mid$(X$, 41, 12))
             SUMACHESEL# = SUMACHESEL# + IPASE#
             '
             CAITAB1% = CAITAB1% + 1
             TETABU1$(1, CAITAB1%) = BAN1$
             TETABU1$(2, CAITAB1%) = NUMCHE1$
             TETABU1$(3, CAITAB1%) = OBSERVA1$ 'FALTA
             ' TETABU1$(4, CAITAB1%) = IPASE#   ' Walter, aqui hay un error - ver conmigo
             TETABU1$(4, CAITAB1%) = FORMATNUM$(IPASE#, "F12.2")
             TETABU1$(5, CAITAB1%) = FECHE1$
           End If
         Next I&
         'DOCPARAM(8) = SUMACHESEL#  ' WALTER, AQUI HAY UN ERROR
         DOCPARAM(8) = FORMATNUM$(SUMACHESEL#, "F14.2")
         '
         Call PRINTDOC("FORDEPBA") ' DEPOSITO BANCARIO
         '
         'LIMPIA LA LISTA
         For KK1% = 1 To CACOLTAB1%
           For KK2% = 1 To CAITAB1%
             TETABU1$(KK1%, KK2%) = ""
           Next KK2%
         Next KK1%
         '
199      Call CIERRARCH("*.*")
         Call BAJALDISCO
      End If
      '\\\OT-05-0475-WAR-FIN///
299   Screen.MousePointer = 1
      Call COMUNI("Liquidación Procesada")
      Call LIMPCONCITC
      Label19 = ""
    '
End Sub

Private Sub Command1_Click()

    If MODORDE% = 1 Then
      MODORDE% = 2
        For KI% = 1 To LICHESEL.ListCount
10        IMPOI# = XVALO(Mid$(LICHESEL.LIST(KI% - 1), 41, 12))
          For KJ% = KI% + 1 To LICHESEL.ListCount
            IMPOJ# = XVALO(Mid$(LICHESEL.LIST(KJ% - 1), 41, 12))
            If IMPOJ# < IMPOI# Then
              TTXX$ = LICHESEL.LIST(KJ% - 1)
              MARCA% = 0: If LICHESEL.Selected(KJ% - 1) = True Then MARCA% = 1
              LICHESEL.RemoveItem KJ% - 1
              LICHESEL.AddItem TTXX$, KI% - 1
              If MARCA% = 1 Then LICHESEL.Selected(KI% - 1) = True
              GoTo 10
            End If
          Next KJ%
        Next KI%
      ElseIf MODORDE% = 2 Then
        MODORDE% = 1
        For KI% = 1 To LICHESEL.ListCount
20        FECHI% = FECHANUM(Mid$(LICHESEL.LIST(KI% - 1), 33, 8))
          For KJ% = KI% + 1 To LICHESEL.ListCount
            FECHJ% = FECHANUM(Mid$(LICHESEL.LIST(KJ% - 1), 33, 8))
            If FECHJ% < FECHI% Then
              TTXX$ = LICHESEL.LIST(KJ% - 1)
              MARCA% = 0: If LICHESEL.Selected(KJ% - 1) = True Then MARCA% = 1
              LICHESEL.RemoveItem KJ% - 1
              LICHESEL.AddItem TTXX$, KI% - 1
              If MARCA% = 1 Then LICHESEL.Selected(KI% - 1) = True
              GoTo 20
            End If
          Next KJ%
        Next KI%
      
    End If
    '
End Sub

Private Sub Command13_Click()
  '
  Command13.Enabled = False
   Call ABRECABAN
  '
   Screen.MousePointer = 11
   Call CARGALISEL("SELDEPO", "CAJABANC", "CODICOM_LAR/A15; CODICOM_COR/A15; FECHCONTA/D10; OPCIMOVIM/A6", "TIPOMOVIM = 'CONCITC' ORDER BY CODICOM_LAR", "DISTINCT")
   Screen.MousePointer = 1
   Call SELECHO("SELDEPO")
   liqui$ = TRIM$(VALACT1$("SELDEPO"))
  
   If liqui$ = "" Then GoTo 99
  '
   Call SETULTREG("CUPONES", 0)
  '
   SQLX$ = "SELECT * FROM CAJABANC"
   SQLX$ = SQLX$ + " WHERE  DOCUMENT = '" & liqui$ & "'"
   Dim CAJATMP As New ADODB.Recordset
   Set CAJATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With CAJATMP
    Do While Not .EOF
       U& = U& + 1
       REFERENCIA$ = SAFETEXT$(!DESCRIMOV)
       FECHA% = CVINT(!FECHCONTA)
       COMPROBANTE$ = SAFETEXT$(!CODICOM_LAR)
       IMPORDEBE# = XVALO(!Idebecomp)

       NCLIE = XVALO(!NUCLIEN)
       CUENTACONT% = XVALO(!CUECONTAI)
       CUENEXT$ = CODCUE$(CUENTACONT%)

       Call REPLA(TX$, REFERENCIA$, 1, 64)
       Call REPLA(TX$, MKI$(FECHA%), 65, 2)
       Call REPLA(TX$, COMPROBANTE$, 67, 18)
       Call REPLA(TX$, MKS$(NCLIE), 85, 4)
       Call REPLA(TX$, MKD$(IMPORDEBE#), 97, 8)
       Call REPLA(TX$, CUENEXT$, 105, 12)

       Call REGAPP("CUPONES", TX$)

       .MoveNext
    Loop
    
   End With
   CAJATMP.Close
   Set CAJATMP = Nothing
   '
   SQLX1$ = "SELECT * FROM CAJABANC"
   SQLX1$ = SQLX1$ + " WHERE  Codicom_lar = '" & liqui$ & "'"
   Dim CAJATMP1 As New ADODB.Recordset
   Set CAJATMP1 = FLEXCONN.Execute(FILTSQL$(SQLX1$))
   With CAJATMP1
    Do While Not .EOF
       U& = U& + 1
       IMPORTHABE# = XVALO(!Idebecomp)
       If IMPORTHABE# > 0 Then
           CUENTACONT% = XVALO(!CUECONTAI)
           REFERENCIA$ = DENOCUE$(CUENTACONT%)
           FECHA% = CVINT(!FECHCONTA)
           COMPROBANTE$ = SAFETEXT$(!CODICOM_LAR)
           CUENEXT$ = CODCUE$(CUENTACONT%)
           Call REPLA(TX$, REFERENCIA$, 1, 64)
           Call REPLA(TX$, MKI$(FECHA%), 65, 2)
           Call REPLA(TX$, COMPROBANTE$, 67, 18)
           Call REPLA(TX$, MKI$(0), 85, 4)
           Call REPLA(TX$, MKD$(IMPORTHABE#), 89, 8)
           Call REPLA(TX$, MKD$(0), 97, 8)
           Call REPLA(TX$, CUENEXT$, 105, 12)
           Call REGAPP("CUPONES", TX$)
       End If
       .MoveNext
    Loop
    
   End With
   CAJATMP1.Close
   Set CAJATMP1 = Nothing

  
   Call FINAL("*CUPONES")
  'Call FINAL("?LISTACUPONES(" + liqui$ + ")")

99 Command13.Enabled = True
End Sub

Private Sub Command2_Click()
   Static ABUSQUE$
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", "Introduzca Argumento de Búsqueda", ABUSQUE$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     INI& = LICHESEL.ListIndex
10   For U& = INI& + 2 To LICHESEL.ListCount
       If InStr(UCase$(LICHESEL.LIST(U& - 1)), ABUSQUE$) > 0 Then
         LICHESEL.ListIndex = U& - 1
         Exit Sub
       End If
     Next U&
     If INI& >= 0 Then
       INI& = -1
       GoTo 10
     End If
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If
End Sub

Private Sub Command3_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Label10 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Command4_Click()

    Call SELECHO("MECOBTC")    'Medios de cobranza Tarjetas de Credito
    VDEVU$ = TRIM$(VALACT1$("MECOBTC"))
    If VDEVU$ <> "" Then
        Label19.Caption = VALACT1$("MECOBTC")
        Label17.Caption = VALACT2$("MECOBTC")
    End If
    
End Sub

Private Sub Command5_Click()
  '
  If DERACCE%("ANULIQUI", "Anulación de Liquidación") < 2 Then GoTo 99
  ' LISTA DE LIQUIDACIONES
  Screen.MousePointer = 11
  Call CARGALISEL("SELDEPO", "CAJABANC", "CODICOM_LAR/A15; CODICOM_COR/A15; FECHCONTA/D10; OPCIMOVIM/A6", "TIPOMOVIM = 'CONCITC' ORDER BY CODICOM_LAR", "DISTINCT")
  Screen.MousePointer = 1
  Call SELECHO("SELDEPO")
  liqui$ = TRIM$(VALACT1$("SELDEPO"))
  '
  ' PROCESO DE ANULACION DE LIQUIDACION
  If liqui$ <> "" Then
       CONDI1$ = "CODICOM_LAR = '" & liqui$ & "'"
       SUMALIQ1 = XVALO(VALOBADA("CAJABANC", "SUM(IdebeComp)", CONDI1$))
       CONDI2$ = "TipoMovim = 'CHRECI' AND DocumEnt = '" & liqui$ & "'"
       SUMALIQ2 = XVALO(VALOBADA("CAJABANC", "SUM(IDEBEComp)", CONDI2$))
       If Abs(SUMALIQ1 - SUMALIQ2) >= 0.005 Then
          Call MENSERR(24, "Liquidación No Balanceada.\Imposible Anular.\   \Consulte !")
          GoTo 99
       End If
       '
       If COMALTER%("¿ Confirma la Anulación de la Liquidación '" & liqui$ & "' ?\\No, Cancelar\Si, Anular") <> 2 Then GoTo 99
       '
       Screen.MousePointer = 11
       ' ELIMINA MOVIMIENTO DE LIQUIDACION
       SQLD$ = "DELETE FROM CAJABANC"
       SQLD$ = SQLD$ + " WHERE " & CONDI1$
        FLEXCONN.Execute (FILTSQL$(SQLD$))
       '
       ' BLANQUEA LOS CUPONES QUE FORMAN LA LIQUIDACION
       Call ACTUREGISTRO("CAJABANC", "STATUS", CONDI2$, 1, REG&)
       Call ACTUREGISTRO("CAJABANC", "DESTICHEQ", CONDI2$, "", REG&)
       Call ACTUREGISTRO("CAJABANC", "REFESALCHE", CONDI2$, "", REG&)
       Call ACTUREGISTRO("CAJABANC", "IDENSALCHE", CONDI2$, "", REG&)
       Call ACTUREGISTRO("CAJABANC", "FEDEPOSENT", CONDI2$, CVDAT(0), REG&)
       Call ACTUREGISTRO("CAJABANC", "MARCONSI", CONDI2$, "", REG&)
       Call ACTUREGISTRO("CAJABANC", "DOCUMENT", CONDI2$, "", REG&)
       '
       Call COMUNI("Se ha Anulado la Liquidación: '" & liqui$ & "'.")
  End If
  '
99 Screen.MousePointer = 1
  '
End Sub



Private Sub Command6_Click(Index As Integer)
   Call SELECHO("ECUECON")
   VDEVU$ = TRIM$(VALACT1$("ECUECON"))
   If VDEVU$ <> "" Then
      Text6(Index).TEXT = VALACT1$("ECUECON")
      Label3(Index).Caption = ECOARCH$("ECUECON", Text6(Index))
   End If
   Text6(Index).SetFocus
End Sub

Private Sub Command9_Click()
   Call SELECHO("CUEBANC")
   VDEVU$ = TRIM$(VALACT1$("CUEBANC"))
   If VDEVU$ <> "" Then
      Label20.Caption = VALACT1$("CUEBANC")
      Label18.Caption = ECOARCH("CUEBANC", Label20)
   End If
End Sub

Private Sub Form_Activate()
    MODORDE% = 1
    FF% = HOY(HO$)
    Label10 = HO$
End Sub

Private Sub Form_Load()
    Call APLICACIONSKINS(Me, Picture2)
    CtaRetIva$ = CONTROL$("OPRECIBO", "CURETIVA")
End Sub

Private Sub Label12_Change()
    NETOACRE# = XVALO(Label12) - XVALO(Text2) - XVALO(Text4) - XVALO(Text3)
    Label5 = FORMATNUM$(NETOACRE#, "F11.2")
End Sub

Private Sub Label19_Change()
    '
    If Label19 = "" Then Exit Sub
    Call LIMPCONCITC
    Call APERBASDAT("CABAN")
    
    MEDCOB% = XVALO(Label19)
    If MEDCOB% = 0 Then Exit Sub
    XX$ = FILTERGETRECORD$("MECOBTC", 1, MKI$(MEDCOB%))
    Label11 = TRIM$(Mid$(XX$, 45, 12))
    CODTAR$ = TRIM$(Mid$(XX$, 57, 4))    'Codigo de la tarjeta de credito
    Label14 = CODTAR$
    '
    SQLX$ = "SELECT * FROM CAJABANC with(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'CHRECI' "
    SQLX$ = SQLX$ + " AND (OPCIMOVIM = 'TC.' OR OPCIMOVIM = 'TC')"  'Selecciona los pagos con T/C
    SQLX$ = SQLX$ + " AND VARIMOVIM = '" & CODTAR$ & "'"      'codigo de la T/C
    SQLX$ = SQLX$ + " AND (MARCONSI IS NULL OR MARCONSI = '')"                   'que no se ha ingresado
    '
    Set BUSCATCTMP = READSET(SQLX$)
'    Dim BUSCATCTMP As ADODB.Recordset
'    Set BUSCATCTMP = New ADODB.Recordset
'25  On Error Resume Next
'    BUSCATCTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'       On Error GoTo 0
'       Call CapturaErrorOpen
'       GoTo 25
'    End If
'    On Error GoTo 0
    
    JJ& = 0: TRANSPOR# = 0
    FORMALTER.LINOSORT.Clear
    CONCITC07.List1.Clear
    CONCITC07.LICHESEL.Clear
    With BUSCATCTMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         '
         FEACRED% = CVINT(!fechacred)
         BANUFE$ = SAFETEXT$(!DESCRIMOV)
         ICHEQ# = XVALO(!ValAbsComp)
         '
         If TRIM$(Left$(BANUFE$, 12)) = "" Then
           Call REPLA(BANUFE$, "Cupon Nro:", 1, 11)
         End If
         RECHE& = XVALO(!NumeComp)
         '
         Z$ = "A" + Space$(80)
         Call REPLA(Z$, CSTRING$(MKI$(FEACRED%), 1, 5, 0, 2), 2, 5)
         Call REPLA(Z$, TRIM$(Left$(BANUFE$, 40)), 9, 40)
         Call REPLA(Z$, CSTRING$(MKD$(ICHEQ#), 4, 12, 2, 2), 49, 12)
         Call REPLA(Z$, CSTRING$(MKS$(RECHE&), 3, 8, 0, 2), 73, 8)
         CONCITC07.List1.AddItem Z$
         '
19     .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Set BUSCATCTMP = Nothing
   
   SUCHETO# = 0
   For I& = 1 To CONCITC07.List1.ListCount
      Z$ = Mid$(CONCITC07.List1.LIST(I& - 1), 9)
      CONCITC07.LICHESEL.AddItem Z$
      SUCHETO# = SUCHETO# + XVALO(Mid$(Z$, 41, 12))
   Next I&
   '
   Screen.MousePointer = 1
   DoEvents
   '
End Sub

Private Sub Label19_DblClick()
    Call Command4_Click
End Sub

Private Sub Label20_DblClick()
    Call Command9_Click
End Sub

Private Sub Label5_Change()
  Call Text1_Change
End Sub

Private Sub LICHESEL_Click()
    SUCHESE# = 0
    For I& = 1 To LICHESEL.ListCount  'Carga el importe "Neto a Acreditar:" segun los cupones seleccionados
       If LICHESEL.Selected(I& - 1) = True Then
         SUCHESE# = SUCHESE# + XVALO(Mid$(LICHESEL.LIST(I& - 1), 41, 12))
       End If
    Next I&
    Label12.Caption = CSTRING$(MKD$(SUCHESE#), 4, 11, 2, 2)
    Label12.Refresh
End Sub

Private Sub Text1_Change()
   VALOR# = XVALO(Text1) - XVALO(Label5)
   Label16 = FORMATNUM$(VALOR#, "F12.2")
   
End Sub

Private Sub Text1_LostFocus()
    VALOR# = XVALO(Text1)
    Text1 = FORMATNUM$(VALOR#, "F12.2")
End Sub

Private Sub Text2_Change(Index As Integer)
    RESTA# = 0
    For KKI% = 0 To 9: RESTA# = RESTA# + XVALO(Text2(KKI%)): Next KKI%
    NETOACRE# = XVALO(Label12) - RESTA#
    Label5 = FORMATNUM$(NETOACRE#, "F11.2")
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Index < 7 Then
           Text6(Index + 1).SetFocus
         Else
           BOTREC.SetFocus
       End If
    End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    VALOR# = XVALO(Text2(Index))
    Text2(Index) = FORMATNUM$(VALOR#, "F11.2")
End Sub


Private Sub Text6_Change(Index As Integer)
    DECUE$ = ECOARCH("ECUECON", TRIM$(Text6(Index)))
    If DECUE$ <> "" Then
        Label3(Index) = DECUE$
      Else
        Label3(Index) = ""
    End If
    
    If TRIM$(Text6(Index)) <> TRIM$(Text16(Index).ToolTipText) Then
       Text16(Index).ToolTipText = ""
       Text16(Index) = ""
    End If
End Sub

Private Sub Text6_DblClick(Index As Integer)
    Call Command6_Click(Index)
End Sub

Private Sub Text6_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
         Call Command6_Click(Index)
      ElseIf KeyCode = 27 Then
         BOTREC.SetFocus
    End If
End Sub

Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2(Index).SetFocus
    End If
End Sub
Private Sub Command6_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2(Index).SetFocus
    End If
End Sub
Sub LIMPCONCITC()
    LICHESEL.Clear
    Label17 = ""
    Label20 = ""
    Label18 = ""
    Label12 = ""
    Label5 = ""
    Text1 = ""
    Label16 = ""
    Text22 = ""
    For KKI% = 0 To 9
       Text2(KKI%).TEXT = ""
       Text6(KKI%).TEXT = ""
       Label3(KKI%) = ""
       Text16(KKI%).TEXT = ""
    Next KKI%
    '
End Sub


Sub CONTROL_CUENTA_RETENCIONES(Index)
    CUENSELEC% = CUEINT%(Text6(Index))
    If TRIM$(Text16(Index).ToolTipText) <> TRIM$(Text6(Index)) Then
       Text16(Index).ToolTipText = ""
       Text16(Index) = ""
    End If
    
    If CUENSELEC% > 0 Then
      CtaRetIva1% = CUEINT%(CtaRetIva$)
      If CtaRetIva1% = CUENSELEC% Then ' si es la cuenta contable de retencion de iva
30      RETEDAT.Label6.Visible = False
        RETEDAT.Label7.Visible = False
        RETEDAT.CMBCOMPROBANTE.Visible = False
        RETEDAT.TXTLETRA.Visible = False
        RETEDAT.TXTCERTIFICADO = Text16(Index)
        RETEDAT.Show 1
        If TRIM$(RETEDAT.TXTCERTIFICADO) = "" Then
          Unload RETEDAT
          Call COMUNI("Debe Ingresar El Certificado - Código de Operación ")
          GoTo 30
        End If
        Text16(Index) = RETEDAT.TXTCERTIFICADO
        Text16(Index).ToolTipText = Text6(Index) 'GUARDO PARA COMPARAR SI CAMBIA EN ESE CASO LIMPIO EL CASILLERO
        Unload RETEDAT
        GoTo 99
      End If
      
      CuentaRetIB% = CtaRetIIBBSegunProvyProvincia%(TRIM$(Label17))
      If CuentaRetIB% = CUENSELEC% Then ' si es la cuenta contable de retencion de iva
          Call LIMPIARTEXT(RETEDAT, 1)
          Erase CADENATEX$
          Call TEXTOLOTES$(TRIM$(Text16(Index)), ";")
          On Error Resume Next
          
          For I& = 1 To UBound(CADENATEX$)
            Select Case I&
              Case 1: RETEDAT.CMBCOMPROBANTE = CADENATEX$(I&)
              Case 2: RETEDAT.TXTCERTIFICADO = CADENATEX$(I&)
              Case 3: RETEDAT.TXTLETRA = CADENATEX$(I&)
            End Select
          Next I&
          On Error GoTo 0
          If IsFormLoaded(RETEDAT) = True Then
            RETEDAT.Show 1
          End If
          TIPOCOX$ = TRIM$(RETEDAT.CMBCOMPROBANTE.TEXT)
          CERTIFICAX$ = TRIM$(RETEDAT.TXTCERTIFICADO)
          LETRAX$ = TRIM$(RETEDAT.TXTLETRA)
          Text16(Index) = TIPOCOX$ + ";" + CERTIFICAX$ + ";" + LETRAX$
          Text16(Index).ToolTipText = Text6(Index)
          Unload RETEDAT
          GoTo 99
       End If
      End If
      
    Unload RETEDAT
    Text16(Index) = ""
99

End Sub


Private Sub Text6_LostFocus(Index As Integer)
   Call CONTROL_CUENTA_RETENCIONES(Index)
End Sub
