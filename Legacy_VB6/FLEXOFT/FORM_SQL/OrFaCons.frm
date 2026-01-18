VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form OrFaCons 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONTROL TOTAL DE PEDIDOS GENERADORES DE ORDEN DE FABRICACION"
   ClientHeight    =   8850
   ClientLeft      =   -3210
   ClientTop       =   765
   ClientWidth     =   15270
   DrawWidth       =   10
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   8805
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15135
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "Componentes de Artículo  Activo"
         Top             =   4560
         Width           =   7155
      End
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   12255
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "Stock de Depósito"
         Top             =   2040
         Width           =   1650
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13905
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   2040
         Width           =   450
      End
      Begin VB.TextBox TXTNROCOTIZA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   13920
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   1080
         Width           =   1170
      End
      Begin VB.TextBox TXTFTECNICA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   13920
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   720
         Width           =   1170
      End
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   12840
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "F.Técnica:"
         Top             =   720
         Width           =   1055
      End
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   12840
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Cotización:"
         Top             =   1080
         Width           =   1055
      End
      Begin VB.TextBox TXTCODIN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   14280
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1440
         Width           =   810
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "Códigos Posibles según Rango "
         Top             =   2040
         Width           =   7170
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Componentes de Artículo  Activo"
         Top             =   2040
         Width           =   7170
      End
      Begin VB.TextBox TXTNUMCLIE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11880
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
         Width           =   930
      End
      Begin VB.TextBox TXTREFEPEDI 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8040
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   4770
      End
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   12840
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Entrega:"
         Top             =   360
         Width           =   1055
      End
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   12840
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Pedido Nro:"
         Top             =   0
         Width           =   1055
      End
      Begin VB.CommandButton CMDOT 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2840
         TabIndex        =   25
         ToolTipText     =   "Ordenes de Fabricación Abiertas"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox TXTOT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   0
         Width           =   2805
      End
      Begin VB.TextBox TXTCANTPED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   360
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   4
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Cant. Pedida:"
         ToolTipText     =   " "
         Top             =   360
         Width           =   1290
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2840
         TabIndex        =   21
         ToolTipText     =   "Ordenes de Fabricación Todas (Abiertas y Cerradas)"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox TXTCLIENTE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   7320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   0
         Width           =   5490
      End
      Begin VB.TextBox TXTCODIGO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Primera vez que se Fabrica Fondo Rojo "
         Top             =   0
         Width           =   4050
      End
      Begin VB.TextBox TXTNROPEDIDO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   13920
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   0
         Width           =   1170
      End
      Begin VB.TextBox TXTFENTREGA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   13920
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   360
         Width           =   1170
      End
      Begin VB.TextBox TXTDESCRI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Primera vez que se Fabrica Fondo Rojo "
         Top             =   360
         Width           =   4050
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "Situación del Cliente y Pediddo Activo"
         Top             =   840
         Width           =   12810
      End
      Begin VB.TextBox TXTNROPED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6765
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   1070
      End
      Begin VB.TextBox TXTNUMCLI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton Command2 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7845
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   175
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   5640
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Nro. Pedido:"
         Top             =   720
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1725
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   175
      End
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   29
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Cuenta:"
         Top             =   720
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox TXTRASOCLI 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   3645
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Situación General de los Pedidos del Cliente Activo"
         Top             =   7320
         Width           =   14475
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   720
         Left            =   0
         TabIndex        =   5
         Top             =   1215
         Width           =   12795
         _ExtentX        =   22569
         _ExtentY        =   1270
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   16501405
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   1035
         Left            =   0
         TabIndex        =   6
         Top             =   7680
         Width           =   14475
         _ExtentX        =   25532
         _ExtentY        =   1826
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   8421504
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   2055
         Left            =   0
         TabIndex        =   29
         Top             =   2400
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   3625
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   16501405
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSF4 
         Height          =   2055
         Left            =   7200
         TabIndex        =   31
         ToolTipText     =   "Doble Clic Ficha Kardex"
         Top             =   2400
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   3625
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   16501405
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSF5 
         Height          =   2250
         Left            =   0
         TabIndex        =   40
         ToolTipText     =   "Doble Clic Para Guardar o Editar Ubicación"
         Top             =   4920
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   3969
         _Version        =   393216
         Rows            =   26
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stock depo1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   735
         Left            =   12840
         TabIndex        =   37
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         Caption         =   "TEPRUEBA"
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   870
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   3240
      Top             =   5040
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.Menu MNUARCHIVO 
      Caption         =   "ARCHIVO"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar 
         Caption         =   " Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu InformaCantidadEnProceso 
         Caption         =   "Informar Cantidad en Proceso"
         Visible         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnudesap 
         Caption         =   "Des-Aplicar Informe de Inicio de Proceso"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MNUARCHIVO2 
      Caption         =   "ARCHIVO2"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar2 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu modfinproc 
         Caption         =   "Modificar Informe de Fin de Proceso"
      End
      Begin VB.Menu S2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnureporte 
      Caption         =   "Reportes"
      Begin VB.Menu sininformar 
         Caption         =   "Ordenes de Fabricación Sin Informar "
      End
      Begin VB.Menu inicioinfo 
         Caption         =   "Ordenes de Fabricación Informado  su inicio"
      End
      Begin VB.Menu cantacerrar 
         Caption         =   "Ordenes de Fabricación Informada Cantidad lista para Cerrar"
      End
   End
   Begin VB.Menu pp 
      Caption         =   "Planilla de Producción"
      Visible         =   0   'False
      Begin VB.Menu geplapro 
         Caption         =   "Generar Planilla de Producción"
      End
      Begin VB.Menu viplage 
         Caption         =   "Visualizar Planillas Generadas"
      End
   End
End
Attribute VB_Name = "OrFaCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GLOBALFRM_strContenidoCelda$
Dim GLOBALFRM_FILA&
Public NORFAEXTERNA$

Sub Carga_Sit_General()
     MSF3.Rows = 1
     HOX% = HOY(HO$)
     SQLX$ = "SELECT PEDDETA.CodiExt, PEDDETA.Descrip, PEDDETA.CanPed, PEDDETA.CantEnt, PEDDETA.NROPED, PEDDETA.FeSolEnt,"
     SQLX$ = SQLX$ + " ORDEFABR.IdSubOr, ORDEFABR.CanProy, ORDEFABR.CanApro, PEDDETA.NuOrItem "
     SQLX$ = SQLX$ + " FROM PEDDETA WITH(NOLOCK)  INNER JOIN ORDEFABR ON PEDDETA.NroPed = ORDEFABR.NroPed"
     SQLX$ = SQLX$ + " WHERE ORDEFABR.Cod_Inte = PEDDETA.CODIINT "
     SQLX$ = SQLX$ + " AND NROCLIE = " & XVALO(TXTNUMCLI)
     SQLX$ = SQLX$ + "  ORDER BY PEDDETA.NROPED "
     Set RS1 = READSET(SQLX$)
     With RS1
       Do While Not .EOF
         CODI$ = SAFETEXT$(!CODIEXT)
         DESCRIX$ = SAFETEXT$(!Descrip)
         CANP = XVALO(!CanPed)
         CANEN = XVALO(!CANTENT)
         CAPRO = XVALO(!CanApro)
         NOR% = XVALO(!NuOrItem)
         FESOL% = CVINT(!FeSolEnt)
         NP& = XVALO(!NROPED)
         J& = J& + 1
         MSF3.Rows = J& + 1
         MSF3.TextMatrix(J&, 1) = NP&
         MSF3.TextMatrix(J&, 2) = CODI$
         MSF3.TextMatrix(J&, 3) = DESCRIX$
         MSF3.TextMatrix(J&, 4) = CANP
         MSF3.TextMatrix(J&, 5) = CANEN
         MSF3.TextMatrix(J&, 6) = FECHATEX$(FESOL%)
         MSF3.TextMatrix(J&, 7) = CAPRO
         DISPONIBLE = 0
         If (CAPRO - CANEN) > 0 Then DISPONIBLE = (CAPRO - CANEN)
         MSF3.TextMatrix(J&, 8) = DISPONIBLE
         ATRAZO = 0
         If HOX% > FESOL% Then ATRAZO = DIENTRE%(FESOL%, HOX%)
         MSF3.TextMatrix(J&, 9) = ATRAZO
         .MoveNext
       Loop
     End With
   RS1.Close
   Set RS1 = Nothing
     
End Sub

Private Sub CMDOT_Click()

   CMDOT.Enabled = False
   NROOF$ = ORFASEL_ENVAPLAST$
   Me.TXTCODIGO.BackColor = &HE0E0E0: Me.TXTCODIGO.ForeColor = &H80000008: Me.TXTCODIGO.ToolTipText = ""
   Me.TXTDESCRI.BackColor = &HE0E0E0: Me.TXTDESCRI.ForeColor = &H80000008: Me.TXTDESCRI.ToolTipText = ""
   If NROOF$ <> "" Then
     SQLX$ = "SELECT REFEREN,NROPED, Cod_Inte   FROM ORDEFABR WITH(NOLOCK) WHERE IDSUBOR = '" & NROOF$ & "'"
     Set RS0 = READSET(SQLX$)
     CI1% = XVALO(RS0!COD_INTE)
     Me.TXTCODIN = CI1%
     Me.TXTCODIGO = CODEXT$(CI1%)
     TXTOT.TEXT = NROOF$
     NROPED& = XVALO(RS0!NROPED)
     TXTDESCRI = DESCRIT0$(CI1%)
     Me.TXTNROPEDIDO = NROPED&
     TXTECOOF = SAFETEXT$(RS0!Referen)
     NUOR% = XVALO(RESP_ZELE_VECT(9))
     'STAT% = XVALO(RS0!StatuOrf)
     RS0.Close
     Set RS0 = Nothing
     
     'CONTROLO SI EL CODIGO SELECCIONADO YA TUVO ALGUNA VEZ ORDENES DE FABRICACION CERRADAS, SI NO ES ASI
     'LO CONSIDERO QUE ES UN CODIGO O UNA VERSION NUEVA Y RESALTO EL NUMERO DE OF. PARA QUE PRESTEN ATENCION
     NOHAYOF% = CantRegistros&("ORDEFABR", "StatuOrf = 3 AND COD_INTE = " & CI1%, "NOMESS")
     If NOHAYOF% < 1 Then
       Me.TXTCODIGO.BackColor = vbRed: Me.TXTCODIGO.ForeColor = vbYellow
       Me.TXTDESCRI.BackColor = vbRed: Me.TXTDESCRI.ForeColor = vbYellow
       Me.TXTCODIGO.ToolTipText = "Código o Versión Nueva"
       Me.TXTDESCRI.ToolTipText = "Código o Versión Nueva"
     End If
     '
   End If

   If NROPED& > 0 Then
       SQLX$ = "SELECT NroClie, RasoClie FROM PEDENCA WITH(NOLOCK) WHERE NROPED = " & NROPED&
       
       Set RS = READSET(SQLX$)
       TXTCLIENTE = SAFETEXT$(RS!RaSoClie)
       TXTNUMCLIE = XVALO(RS!NROCLIE)
       RS.Close
       Set RS = Nothing
       SQLX$ = "SELECT NUORITEM,FeSolEnt, CanPed FROM PEDDETA WITH(NOLOCK) WHERE NROPED = " & NROPED& & " AND CODIINT = " & CI1%
       If NUOR% > 0 Then SQLX$ = SQLX$ + " AND NUORITEM = " & NUOR%
       Set RS = READSET(SQLX$)
       With RS
         FESOL% = CVINT(!FeSolEnt)
         CANTI = XVALO(!CanPed)
       End With
       RS.Close
       Set RS = Nothing
       Me.TXTCANTPED = FORMATNUM$(CANTI, "G12.1")
'       FESOL% = CVINT(VALOBADA("PEDDETA", "FeSolEnt", "NROPED = " & NROPED& & " AND CODIINT = " & CI1%, "NOMESS"))
'       CANTI = XVALO(VALOBADA("PEDDETA", "CanPed", "NROPED = " & NROPED& & " AND CODIINT = " & CI1%, "NOMESS"))
'       Me.TXTCANTPED = FORMATNUM$(CANTI, "G12.1")
   End If
   TXTFENTREGA = FECHATEX$(FESOL%)
   'Me.TXTCANT = FORMATNUM$(CANTI, "F10.1")
   CMDOT.Enabled = True
   
End Sub



Private Sub Command1_Click()

  Command1.Enabled = False
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "Cuenta/A10;Razon Social/a48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cuenta/A10;Razon Social/a48", FRMZELECHO)
  FRMZELECHO.Caption = "Clientes Activos"

  Dim obj As New RECORXET
  Set RS1 = obj.RS_DEUDOR12()
  Call Carga_MSF_Recordset(RS1, "Cuenta/A10;Razon Social/a48", FRMZELECHO.MSF2)
  Set obj = Nothing
  RS1.Close
  Set RS1 = Nothing

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  
  NCLI = XVALO(RESP_ZELE_VECT(1))
  If NCLI > 0 Then Me.TXTNUMCLI = NCLI
99 Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "Nro.Pedido/f9;Referencia/a48", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Nro.Pedido/f9;Referencia/a48", FRMZELECHO)
    FRMZELECHO.Caption = "Pedidos Pendientes de Entrega del Cliente:"
    '****************************************************************
    
    Dim obj As New RECORXET
    Set RS1 = obj.RS_PedidoPendxCte(XVALO(Me.TXTNUMCLI))
    Call Carga_MSF_Recordset(RS1, "Nro.Pedido/f9;Referencia/a48", FRMZELECHO.MSF2)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    
    FRMZELECHO.Width = 8500
    FRMZELECHO.Show 1
    NRO = XVALO(RESP_ZELE_VECT(1))
    If NRO > 0 Then TXTNROPED.TEXT = NRO
99  Command2.Enabled = True

End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If DERACCE%("PDRITDAS", "SELECCIÓN DE ORDENES DE FABRICACIÓN ABIERTAS Y CERRADAS") < 2 Then GoTo 99
'    SQLX$ = "SELECT GETDATE() AS FECHA"
'    Set RS = READSET(SQLX$, , 1)
'    FFF = RS!FECHA
'    FE% = CVINT(RS!FECHA)
   NROOF$ = ORFASEL_ENVAPLAST$(1)
   If NROOF$ <> "" Then
     SQLX$ = "SELECT REFEREN,NROPED, Cod_Inte   FROM ORDEFABR WITH(NOLOCK) WHERE IDSUBOR = '" & NROOF$ & "'"
     Set RS0 = READSET(SQLX$)
     CI1% = XVALO(RS0!COD_INTE)
     Me.TXTCODIN = CI1%
     Me.TXTCODIGO = CODEXT$(CI1%)
     TXTOT.TEXT = NROOF$
     NROPED& = XVALO(RS0!NROPED)
     TXTDESCRI = DESCRIT0$(CI1%)
     Me.TXTNROPEDIDO = NROPED&
     TXTECOOF = SAFETEXT$(RS0!Referen)
     NUOR% = XVALO(RESP_ZELE_VECT(9))
     'VUELVO LOS CASILLEROS AL COLOR ORIGINAL (LAS ORDENES CERRADAS NUNCA DEBERIAN ESTAR CON ROJO.)
     Me.TXTCODIGO.BackColor = &HE0E0E0: Me.TXTCODIGO.ForeColor = &H80000008: Me.TXTCODIGO.ToolTipText = ""
     Me.TXTDESCRI.BackColor = &HE0E0E0: Me.TXTDESCRI.ForeColor = &H80000008: Me.TXTDESCRI.ToolTipText = ""
   End If
   If NROPED& > 0 Then
       SQLX$ = "SELECT NroClie, RasoClie FROM PEDENCA WITH(NOLOCK) WHERE NROPED = " & NROPED&
       Set RS = READSET(SQLX$)
       TXTCLIENTE = SAFETEXT$(RS!RaSoClie)
       TXTNUMCLIE = XVALO(RS!NROCLIE)
       RS.Close
       Set RS = Nothing
       '-------------------
       SQLX$ = "SELECT NUORITEM,FeSolEnt, CanPed FROM PEDDETA WITH(NOLOCK) WHERE NROPED = " & NROPED& & " AND CODIINT = " & CI1%
       If NUOR% > 0 Then SQLX$ = SQLX$ + " AND NUORITEM = " & NUOR%
       Set RS = READSET(SQLX$)
       With RS
         FESOL% = CVINT(!FeSolEnt)
         CANTI = XVALO(!CanPed)
       End With
       RS.Close
       Set RS = Nothing
       Me.TXTCANTPED = FORMATNUM$(CANTI, "G12.1")
       '-----------------------
'       FESOL% = CVINT(VALOBADA("PEDDETA", "FeSolEnt", "NROPED = " & NROPED& & " AND CODIINT = " & CI1%, "NOMESS"))
'       CANTI = XVALO(VALOBADA("PEDDETA", "CanPed", "NROPED = " & NROPED& & " AND CODIINT = " & CI1%, "NOMESS"))
'       Me.TXTCANTPED = FORMATNUM$(CANTI, "G12.1")
   End If
   TXTFENTREGA = FECHATEX$(FESOL%)
   'Me.TXTCANT = FORMATNUM$(CANTI, "F10.1")
99 Command3.Enabled = True

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Activate()
   On Error Resume Next
   Me.TXTNUMCLI.SetFocus
   On Error GoTo 0
End Sub

Private Sub Form_Load()
   Campos$ = "N.Ped./F8;Código/a16;Descripción/A24;Q.Ped/F9;Q.Entr./F9;F.Sol./D8;Q.Fabricada/F12;Disp.Ent/F10;Atrazo/F5"
   Call CARGA_ENCABEZADO(Me.MSF2, "Código/A12;Descripción/A28;U./A2;Cantidad/F12", Me)
   Call CARGA_ENCABEZADO(MSF3, Campos$, Me)
   Call CARGA_ENCABEZADO(MSF5, "Código/A16;Descripción/A28;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/a0;IDUbic/a0;N°Operario/F9;Nombre/A16;Peso/f8.2;IdReg/a0", Me)

   
   Campos$ = "Código R./A0;"                '1
   Campos$ = Campos$ + "Descripción R./A0;" '2
   Campos$ = Campos$ + "Código/A16;"        '3
   Campos$ = Campos$ + "Descripción/A24;"   '4
   Campos$ = Campos$ + "CIR/A0;"            '5
   Campos$ = Campos$ + "CIMATP/A0;"         '6
   Campos$ = Campos$ + "ID/A0;"             '7
   Campos$ = Campos$ + "Stock/f10.1"             '8

   Call CARGA_ENCABEZADO(MSF4, Campos$, Me)
   
   Text2 = Control$("", "DEPOCEN")

End Sub

Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub

Private Sub MSF1_DblClick()
    Exit Sub
    VENTANA.Inicializar = 0
    VENTANA.CampEditables = "3"
    VENTANA.AgregaRegistro = 0
    VENTANA.UtilizaFormula = 0
    VENTANA.HabilitaBorrar = 0
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Situación de Pedidos:"
    VENTANA.Inicializar = 1
    VENTANA.Campox = "Código/a16;Descripción/A32;Q.Pedida/f9;Q.Entreg./F9;O.Fabric./a12;Q.Proy./f9;Q.Aprob./f9;Orden/F0"
    
    '*** SE DEBERA CARGAR PARA EL PARAMETRO  VENTANA.Campox = LOS CAMPOS DEL MSF BASE  ****
    '****************************************************************
    '
    '*** SE DEBERA CARGAR PARA LOS PARAMETROS CopiarMsfAOtroMsf(GRILLABASE, VENTADISE.MSF)   ****
    '*** REALIZAR LA LLAMADA
    Call CopiarMsfAOtroMsf(MSF1, VENTABNEW.MSF, 1, MSF1.Width)
    '****************************************************************
    
30  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
       If VALIDAR_CAMBIOS(MSF1, VENTABNEW.MSF) >= 0 Then
         Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF1, 1)
       Else: GoTo 30
       End If
    End If
    '
    Set VENTANA = Nothing
End Sub
Function VALIDAR_CAMBIOS(MSFBASE As MSFlexGrid, MSFMODI As MSFlexGrid)
'
'   ' "Código/a16;Descripción/A32;Q.Pedida/f9;Q.Entreg./F9;O.Fabric./a12;Q.Proy./f9;Q.Aprob./f9;Orden/F0"
'   'SI LA CANTIDAD PEDIDA ES < A LA ENTREGADA NO SE PUEDE
'   With MSFMODI
'     For i& = 1 To .Rows - 1
'       CANTPED = XVALO(.TextMatrix(i&, 3))
'       CANTENT = XVALO(.TextMatrix(i&, 4))
'       If CANTPED < CANTENT Then
'          Call COMUNI("En Fila: " & i& & " Imposible Cambiar la Cantidad Pedida a un Valor Menor que la entregada")
'          VALIDAR_CAMBIOS = -1
'          Exit Function
'       End If
'     Next i&
'   End With
'   '
'    With MSFMODI
'     For i& = 1 To .Rows - 1
'       CANTPED = XVALO(.TextMatrix(i&, 3))
'       CANPROY = XVALO(.TextMatrix(i&, 6))
'       If CANTPED > CANPROY Then
'          OPX% = COMALTER("En Fila: " & i& & " Se Han Detectado Aumentos en Cantidad Pedida\\Mantener Cantidad a Fabricar\Aumentar en Mismo Modo La Cantidad a Fabricar")
'          If OPX% = 2 Then
'           ' VER QUE HACER POR QUE AL ACTUALIZAR PUEDE QUE YA TENGA PARCIALES APROBADOS, Y QUE PASARIA CON LAS RESERVAS.
'          End If
'          Exit Function
'       End If
'     Next i&
'   End With
'
'   'SI LA CANTIDAD PEDIDA ES > A LA ENTREGADA OK
'      ' CONSULTAR SI ACTUALIZA  LA CANTIDAD A FABRICAR
'
'   'SI LA CANTIDAD PEDIDA ES < A LA CANTIDAD FABRICADA AVISAR QUE LA MISMA ES PARA STOCK Y DAR POR CUMPLIDA LA ORDEN
End Function
Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     REG& = MSF1.MouseRow
     COL% = MSF1.MouseCol
     Call MENU_CELDA(MSF1, REG&, COL%)
  End If
End Sub

Sub INVICTRL()
  '
  mnucopiar.Visible = Not (mnucopiar.Visible)
  mnudesap.Visible = Not (mnudesap.Visible)
'  InformaCantidadEnProceso.Visible = Not InformaCantidadEnProceso.Visible
  '
End Sub

Sub MENU_CELDA(MSF As MSFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub

        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&
        PopupMenu MNUARCHIVO
        Call INVICTRL
        On Error Resume Next
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
        On Error GoTo 0
End Sub

''****************************************************************
Private Sub mnucopiar2_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub


Private Sub MSF2_Click()
    REG& = MSF2.MouseRow
    COL& = MSF2.MouseCol
    If REG& = 0 Or REG& >= MSF2.Rows Then Exit Sub
    CODI$ = MSF2.TextMatrix(REG&, 1)
    CI1% = CODINT%(CODI$)
    If CI1% < 1 Then Exit Sub
    MSF4.Rows = 1
    If ESCODIGORANGO&(CI1%) > 0 Then
      Call CARGARASOCIADOARANGO(CI1%)
    Else
      MSF4.Rows = 2
      MSF4.TextMatrix(1, 3) = MSF2.TextMatrix(REG&, 1)
      MSF4.TextMatrix(1, 4) = MSF2.TextMatrix(REG&, 2)
      MSF4.TextMatrix(1, 6) = CI1%
      
    End If
    
    Condi$ = "Cod_inte = " & CI1%
    CANTRES = TRIM$(FORMATNUM$(XVALO(VALOBADA("RESERVA", "SUM(CANTRES-CANTCONS)", Condi$, "NOMESS")), "f10.1"))
    Text5(3) = "Códigos Posibles según Rango (Reservas Totales: " & CANTRES & ")"
    
    'le agrego el stock
    DEPOCEN% = XVALO(Text2)
    For i% = 1 To MSF4.Rows - 1
      CIXI% = XVALO(MSF4.TextMatrix(i%, 6))
      MSF4.TextMatrix(i%, 8) = STODEPOS(CIXI%, DEPOCEN%)
    Next i%

End Sub
Sub CARGARASOCIADOARANGO(CIRANGO%)
    MSF4.Rows = 1
    Campos$ = "Código R./A0;"                '1
    Campos$ = Campos$ + "Descripción R./A0;" '2
    Campos$ = Campos$ + "Código/A16;"        '3
    Campos$ = Campos$ + "Descripción/A24;"   '4
    Campos$ = Campos$ + "CIR/A0;"            '5
    Campos$ = Campos$ + "CIMATP/A0;"         '6
    Campos$ = Campos$ + "ID/A0;"             '7
    Campos$ = Campos$ + "Stock/f10.1"        '8
   
    Dim obj As New RECORXET
    Set RS1 = obj.RS_MATERIASPRIMASDECODIGORANGO(CIRANGO%)
    Call Carga_MSF_Recordset(RS1, Campos$, MSF4)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
End Sub

Private Sub MSF3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     REG& = MSF3.MouseRow
     COL% = MSF3.MouseCol
     Call MENU_CELDA2(MSF3, REG&, COL%)
  End If
End Sub

Sub INVICTRL2()
  '
  mnucopiar2.Visible = Not (mnucopiar2.Visible)
  '
End Sub
Sub MENU_CELDA2(MSF As MSFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL2
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&
        PopupMenu MNUARCHIVO2
        Call INVICTRL2
        On Error Resume Next
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
        On Error GoTo 0
End Sub


Private Sub MSF4_DblClick()
     REG& = MSF4.MouseRow
     COL% = MSF4.MouseCol
     If REG& = 0 Or REG& > MSF4.Rows Then Exit Sub
     FIKARDEX09.Text1 = MSF4.TextMatrix(REG&, 3)
     FIKARDEX09.Show 1

End Sub

Private Sub TXTCODIN_Change()
   Dim CTP As New Conteplast
   TXTFTECNICA = CTP.NroFTecnicaXCodigo(XVALO(TXTCODIN))
End Sub


Private Sub TXTNROPED_Change()
     Me.TXTREFEPEDI = VALOBADA("PEDENCA", "REFERENCIA", "NROPED = " & XVALO(Me.TXTNROPED), "NOMESS")
     Call CARGA_ENCABEZADO(MSF1, "Código/a16;Descripción/A32;Q.Pedida/f9;Q.Entreg./F9;O.Fabric./a12;Q.Proy./f9;Q.Aprob./f9;Orden/F0", Me)
     Call CARGAGRILLADETALLE(XVALO(TXTNROPED))
End Sub
Sub CARGAGRILLADETALLE(NROPED&)
   Call CARGA_ENCABEZADO(MSF1, "Código/a16;Descripción/A32;Q.Pedida/f9;Q.Entreg./F9;O.Fabric./a12;Q.Proy./f9;Q.Aprob./f9;Orden/F0", Me)

   Dim obj As New RECORXET
   Set RS1 = obj.RS_PedidosVsOF(NROPED&)
   Call Carga_MSF_Recordset(RS1, "Código/a16;Descripción/A32;Q.Pedida/f9;Q.Entreg./F9;O.Fabric./a12;Q.Proy./f9;Q.Aprob./f9;Orden/F0", MSF1, 1)
   Set obj = Nothing
   RS1.Close
   Set RS1 = Nothing
End Sub

Private Sub TXTNROPEDIDO_Change()
   TXTNROPED = XVALO(TXTNROPEDIDO)
   TXTNROCOTIZA = NROCOTIZA(XVALO(TXTNROPEDIDO))

End Sub

Private Sub TXTNUMCLI_Change()
   TXTRASOCLI = rasocli$(XVALO(Me.TXTNUMCLI))
   Call Carga_Sit_General
End Sub
Private Sub TXTNUMCLI_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.TXTNROPED.SetFocus
   End If
End Sub


Private Sub TXTNUMCLIE_Change()
    TXTNUMCLI = XVALO(TXTNUMCLIE)
End Sub

Private Sub TXTOT_Change()
    NORFA$ = TRIM$(TXTOT)
    CIFAB% = CODIFAB%(NORFA$)
    If NORFA$ = "" Then Exit Sub
    CANTI = CantiProyectadaOF(NORFA$)
    Call CARGA_ESTRUCTURA_EN_GRILLA(MSF2, XVALO(TXTCODIN), CANTI)

    Dim CTP As New Conteplast
    Call CTP.CargaOperacionesInformadas(NORFA$, CIFAB%, MSF5)
    Set CTP = Nothing
End Sub
