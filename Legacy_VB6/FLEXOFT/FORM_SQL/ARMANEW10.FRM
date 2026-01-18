VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ARMANEW10 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0F0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Armar Mail Circular"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0F0C0&
      Caption         =   "Destinos Seleccionados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   7200
      TabIndex        =   55
      Top             =   3360
      Width           =   3690
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   120
         TabIndex        =   56
         Top             =   300
         Width           =   3400
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   4200
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0F0C0&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   7200
      TabIndex        =   37
      Top             =   1850
      Width           =   3615
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E0F0C0&
         Caption         =   "Minim."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2640
         TabIndex        =   60
         ToolTipText     =   "Ejecuta Minimizado"
         Top             =   1200
         Width           =   915
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0F0C0&
         Caption         =   "Test"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   1800
         TabIndex        =   58
         ToolTipText     =   "Envío a Asistec@flexoft.com.ar"
         Top             =   1200
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0F0C0&
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
         Height          =   250
         Left            =   960
         TabIndex        =   57
         ToolTipText     =   "Envío Automatico sin Mostrar Formulario"
         Top             =   1200
         Width           =   915
      End
      Begin VB.OptionButton optEncodeType 
         BackColor       =   &H00E0F0C0&
         Caption         =   "UUEncode"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   44
         ToolTipText     =   "Use UU Encoding for Attachments."
         Top             =   400
         Width           =   1095
      End
      Begin VB.OptionButton optEncodeType 
         BackColor       =   &H00E0F0C0&
         Caption         =   "MIME"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   43
         ToolTipText     =   "Use MIME encoding for Mail & Attachments."
         Top             =   400
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.CheckBox ckLogin 
         BackColor       =   &H00E0F0C0&
         Caption         =   "Login"
         Height          =   250
         Left            =   2340
         TabIndex        =   42
         ToolTipText     =   "Use Login Authorization When Connecting to a Host"
         Top             =   620
         Width           =   915
      End
      Begin VB.CheckBox ckHtml 
         BackColor       =   &H00E0F0C0&
         Caption         =   "Html"
         Height          =   250
         Left            =   2340
         TabIndex        =   41
         ToolTipText     =   "Mail Body is HTML / Plain Text"
         Top             =   180
         Width           =   1035
      End
      Begin VB.ComboBox cboPriority 
         Height          =   315
         ItemData        =   "ARMANEW10.frx":0000
         Left            =   360
         List            =   "ARMANEW10.frx":0002
         TabIndex        =   40
         Text            =   "cboPriority"
         ToolTipText     =   "Sets the Prioirty of the Mail Message"
         Top             =   750
         Width           =   1530
      End
      Begin VB.CheckBox ckReceipt 
         BackColor       =   &H00E0F0C0&
         Caption         =   "Receipt"
         Height          =   250
         Left            =   2340
         TabIndex        =   39
         ToolTipText     =   "Request a Return Receipt"
         Top             =   400
         Width           =   1035
      End
      Begin VB.CheckBox ckPopLogin 
         BackColor       =   &H00E0F0C0&
         Caption         =   "POP Login"
         Height          =   250
         Left            =   2340
         TabIndex        =   38
         ToolTipText     =   "Use Login Authorization When Connecting to a Host"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   59
         Top             =   1230
         Width           =   540
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   3600
         Y1              =   1150
         Y2              =   1150
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0F0C0&
      Caption         =   "Envio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   30
      Top             =   1850
      Width           =   6855
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar Archivo"
         Height          =   495
         Left            =   5640
         TabIndex        =   53
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtSubject 
         Height          =   285
         Left            =   1005
         TabIndex        =   33
         Top             =   240
         Width           =   5640
      End
      Begin VB.TextBox txtMsg 
         Height          =   1500
         Left            =   1005
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   600
         Width           =   5640
      End
      Begin VB.TextBox txtAttach 
         Height          =   525
         Left            =   1005
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   2160
         Width           =   4560
      End
      Begin VB.Label lblSubject 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   36
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   35
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblAttach 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attachment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   34
         Top             =   2220
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0F0C0&
      Caption         =   "Datos Remitente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   21
      Top             =   120
      Width           =   10575
      Begin VB.CommandButton Command10 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   9800
         Picture         =   "ARMANEW10.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   600
         Width           =   480
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Envíos Pendientes"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   7380
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Re-Enviar Pendientes"
         Top             =   600
         Width           =   980
      End
      Begin VB.CommandButton Command8 
         Caption         =   "X"
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
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   960
         Width           =   300
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   7365
         PasswordChar    =   "*"
         TabIndex        =   50
         Top             =   1320
         Width           =   2970
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   1935
         TabIndex        =   49
         Top             =   1320
         Width           =   4170
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Ch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   8370
         Picture         =   "ARMANEW10.frx":0446
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Cambiar Cuenta de Envío"
         Top             =   600
         Width           =   465
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Test"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   9310
         Picture         =   "ARMANEW10.frx":0590
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   600
         Width           =   480
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   8840
         Picture         =   "ARMANEW10.frx":089A
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   600
         Width           =   480
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   1935
         TabIndex        =   25
         Top             =   240
         Width           =   3465
      End
      Begin VB.TextBox txtFromName 
         Height          =   285
         Left            =   1935
         TabIndex        =   24
         Top             =   600
         Width           =   4850
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   1935
         TabIndex        =   23
         Top             =   960
         Width           =   4605
      End
      Begin VB.TextBox txtPopServer 
         Height          =   285
         Left            =   6840
         TabIndex        =   22
         Top             =   240
         Width           =   3480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6360
         TabIndex        =   52
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   585
         TabIndex        =   51
         Top             =   1380
         Width           =   1305
      End
      Begin VB.Label lblServer 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP Server:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   585
         TabIndex        =   29
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label lblFromName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sender Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   585
         TabIndex        =   28
         Top             =   645
         Width           =   1305
      End
      Begin VB.Label lblFrom 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Email:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   585
         TabIndex        =   27
         Top             =   1020
         Width           =   1305
      End
      Begin VB.Label lblPopServer 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "POP3 Server:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   26
         Top             =   300
         Width           =   1170
      End
   End
   Begin VB.CommandButton Command3 
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
      Height          =   810
      Left            =   10020
      Picture         =   "ARMANEW10.frx":0BA4
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Limpiar Marcas de Envío"
      Top             =   5580
      Width           =   750
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0F0C0&
      Caption         =   "Envios Anteriores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1400
      Left            =   5055
      TabIndex        =   15
      Top             =   6525
      Width           =   4750
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   480
         TabIndex        =   16
         Top             =   525
         Width           =   3700
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   2295
         TabIndex        =   18
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label14 
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
         Height          =   255
         Left            =   3450
         TabIndex        =   17
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0F0C0&
      Caption         =   "Suprimidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1400
      Left            =   210
      TabIndex        =   11
      Top             =   6525
      Width           =   4700
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   480
         TabIndex        =   12
         Top             =   525
         Width           =   3795
      End
      Begin VB.Label Label9 
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
         Height          =   255
         Left            =   3315
         TabIndex        =   14
         Top             =   210
         Width           =   750
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0F0C0&
      Caption         =   "Pendientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1780
      Left            =   210
      TabIndex        =   3
      Top             =   4675
      Width           =   4695
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   465
         TabIndex        =   4
         Top             =   525
         Width           =   3795
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   3315
         TabIndex        =   7
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   10020
      Picture         =   "ARMANEW10.frx":0CEE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Interrumpe Transmisión"
      Top             =   6400
      Width           =   750
   End
   Begin VB.CommandButton Command4 
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
      Height          =   810
      Left            =   10020
      Picture         =   "ARMANEW10.frx":1130
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Iniciar / Continuar Transmisión"
      Top             =   4760
      Width           =   750
   End
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0F0C0&
      Caption         =   "Enviados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1780
      Left            =   5055
      TabIndex        =   5
      Top             =   4675
      Width           =   4750
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   480
         TabIndex        =   6
         Top             =   525
         Width           =   3700
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   63
         Top             =   210
         Width           =   255
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Progress  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   46
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   2295
         TabIndex        =   10
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   3450
         TabIndex        =   9
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.Label Label15 
      Height          =   225
      Left            =   7560
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   9420
      Picture         =   "ARMANEW10.frx":19FA
      Top             =   7395
      Width           =   2010
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
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
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
End
Attribute VB_Name = "ARMANEW10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    ckHtml.Value = 1
    ckLogin.Value = 1
    Command1.Enabled = False
    txtMsg = ""
    txtAttach = ""
    
    CommonDialog1.FileName = Control("", "RUTA1")
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.Filter = "All Files (*.*)|*.*|HTML Files (*.htm;*.html;*.shtml)|*.htm;*.html;*.shtml"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    '
        
    DoEvents
    Screen.MousePointer = 11
    nombarch$ = UCase$(CommonDialog1.FileName)

    For i& = Len(nombarch$) To 1 Step -1 ' ME QUEDO CON LA RUTA SIN EL ARCHIVO
       CARACTER$ = Mid$(nombarch$, i&)
       If InStr(CARACTER$, "\") > 0 Then
         RUTA$ = Left$(nombarch$, i&)
         Exit For
       End If
    Next i&
    '
    
    Call GRACONTROL("", "RUTA1", nombarch$) ' GUARDA LA ULTIMA RUTA

    TEXTOX$ = LOADFILE$(nombarch$)
    txtMsg = TEXTOX$
    POS% = 0
    While InStr(TEXTOX$, "src=" & Chr$(34)) > 0
       POS% = InStr(TEXTOX$, "src=" & Chr$(34))
       TEXTO2$ = Mid$(TEXTOX$, 5 + POS%)
       POS2% = InStr(TEXTO2$, Chr$(34))
       ARCHIVO$ = Mid$(TEXTOX$, 5 + POS%, POS2% - 1)
       txtAttach = txtAttach + RUTA$ + ARCHIVO$ + ";"
       TEXTOX$ = TRIM$(TEXTO2$)
    Wend
    txtAttach = REPLACAR$(txtAttach, "%20", " ")
    If POS% > 0 Then txtAttach = Left$(txtAttach.TEXT, Len(txtAttach) - 1)
    
    Screen.MousePointer = 1

99 Command1.Enabled = True
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Call TRALOCAL("LISMAIL", "")
10  VTB% = VENTABU%("LISMAIL")
    If VTB% >= 0 Then
      FIN& = ULTREG&("!LISMAIL")
      For U& = 1 To FIN&
        XX$ = REGLEIDO$("!LISMAIL", U&)
        MAIL$ = TRIM$(Mid$(XX$, 1, 64))
        USSUU$ = TRIM$(Mid$(XX$, 65, 32))
        PASSW$ = TRIM$(Mid$(XX$, 97, 16))
        SMTPP$ = TRIM$(Mid$(XX$, 113, 32))
        If MAIL$ = "" Or USSUU$ = "" Or SMTPP = "" And U& <= FIN& Then
          Call COMUNI("Datos Incompletos. Verificar")
          GoTo 10
        End If
        '
        UKI% = U&
        Call REPLA(XX$, MAIL$, 1, 64)
        Call REPLA(XX$, USSUU$, 65, 32)
        Call REPLA(XX$, PASSW$, 97, 16)
        Call REPLA(XX$, SMTPP$, 113, 32)
        Call GRAREG("!LISMAIL", XX$, U&)
      Next U&
      Call TRACENTRAL("LISMAIL", "")
   End If
   Command10.Enabled = True
End Sub

 Sub Command2_Click()
   Label15.Caption = "STOP"
End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   If COMALTER%("Está Seguro de Blanquear Lista de Envíos ?\\No, Cancelar\Si, Refrescar") = 2 Then
     Screen.MousePointer = 11
     Call CREACAMPO("", "BSEGCOM", "FULTIEN", 3, 2, 1)
     Call ACTUREGISTRO("BSEGCOM", "FULTIEN", "CODI_CLI <> '' AND CODI_CLI IS NOT NULL", 0, RAFE&)
     Screen.MousePointer = 1
   End If
   Command3.Enabled = True
   Unload Me
   '
End Sub

 Sub Command4_Click()
   ' VALIDACIONES
   Command4.Enabled = False
   HOII% = HOY(HOS$)
   Label15.Caption = ""
   If UCase$(Left$(TRIM(txtServer), 4)) <> "SMTP" Then
      Call COMUNI("Servidor Smtp No Válido"): GoTo 99
   End If
   '
   If TRIM$(txtFromName) = "" Then
      Call COMUNI("Falta Ingresar Nombre"): GoTo 99
   End If
   '
   If ckPopLogin.Value > 0 And TRIM$(txtPopServer) = "" Then
      Call COMUNI("Falta Ingresar Pop3"): GoTo 99
   End If
   '
   If ckHtml.Value > 0 And TRIM$(txtMsg) = "" Then
      Call COMUNI("Falta Ingresar Texto HTML"): GoTo 99
   End If
   '
   If TRIM$(txtUserName) = "" Or TRIM$(txtPassword) = "" Then
      Call COMUNI("Falta Ingresar Usuario y/o Password"): GoTo 99
   End If
   '
   If ckLogin.Value <> 1 Then
      Call COMUNI("Falta Tildar Casilla de Login"): GoTo 99
   End If
   '
   If TRIM$(txtSubject) = "" Then
      Call COMUNI("Debe Ingresar el Asunto (Subject)"): GoTo 99
   End If
   '
   Timer1.Enabled = True
   Exit Sub
   '
99 Command4.Enabled = True
   '
End Sub
'
Sub COPIAR_DATOS()
    
    MAI_SSMTP$ = txtServer                       ' SMTP
    MAI_SNAME$ = txtFromName                     ' SENDER NAME
    MAI_SMAIL$ = txtFrom                         ' SENDER EMAIL
    ' SENDMAIL09.txtPopServer = txtPopServer ' POP 3
    ' MAI_TEXTO$ = txtMsg    ' CUERPO DE MAIL lo carga en el for por que varia
    MAI_ADJUN$(0) = txtAttach                    ' ADJUNTOS
    'SENDMAIL09.ckHtml.Value = ckHtml.Value ' OPTION HTML
    'SENDMAIL09.ckLogin.Value = ckLogin.Value ' OPTION LOGIN
    'SENDMAIL09.ckPopLogin.Value = ckPopLogin.Value ' OPTION POP3
    MAI_UNAME$ = txtUserName                     ' CUENTA DE USUARIO
    MAI_PWORD$ = txtPassword                     ' CLAVE

End Sub

Private Sub Dir1_Change()
   Label11.Caption = Dir1.Path
   File1.Path = Dir1.Path
   Label1 = ""
End Sub

Private Sub Drive1_Change()
   On Error GoTo 20
   Dir1.Path = Drive1.Drive
   GoTo 99
   '
20 Resume 21
21 On Error GoTo 0
   Call MENSERR(24, "Unidad no Disponible.")
   Drive1.Drive = Left$(CurDir$, 2)
   '
99 On Error GoTo 0
End Sub

Private Sub file1_click()
  Label1 = File1.FileName
End Sub

Private Sub Command5_Click()
    '
    Command5.Enabled = False
    '
    If UCase$(Left$(TRIM(txtServer), 4)) <> "SMTP" Then
      Call COMUNI("Servidor Smtp No Válido")
      GoTo 99
    End If
    '
    If TRIM$(txtFrom.TEXT) = "" Then
      Call COMUNI("Falta Ingresar Cuenta")
      GoTo 99
    End If
    '
    If TRIM$(txtUserName) = "" Or TRIM$(txtPassword) = "" Then
      Call COMUNI("Falta Ingresar Usuario o Password")
      GoTo 99
    End If
    '
    Call GRACONTROL("", "SMTPSRVR", txtServer.TEXT) 'smtpserver
    Call GRACONTROL("", "FROMNAME", txtFromName.TEXT)
    Call GRACONTROL("", "FROMMAIL", txtFrom.TEXT)
    Call GRACONTROL("", "POPSERV", txtPopServer.TEXT)
    Call GRACONTROL("", "USUARIO", txtUserName.TEXT)
    Call GRACONTROL("", "KLAVE", txtPassword.TEXT)
    '
    URE& = ULTREG&("LISMAIL")
    CUENTACT& = URE& + 1
    For KKI& = 1 To URE&
       XX$ = REGLEIDO$("LISMAIL", KKI&)
       If TRIM$(Mid$(XX$, 1, 64)) = txtFrom.TEXT Then
          CUENTACT& = KKI&
          Exit For
       End If
    Next KKI&
    '
    XX$ = REGBLAN$("LISMAIL")
    Mid$(XX$, 1, 64) = txtFrom.TEXT           'MAIL
    Mid$(XX$, 65, 32) = txtUserName.TEXT
    Mid$(XX$, 97, 16) = txtPassword.TEXT
    Mid$(XX$, 113, 32) = txtServer.TEXT
    Call GRAREG("LISMAIL", XX$, CUENTACT&)
    Call CIERRARCH("LISMAIL")
    '
    Call COMUNI("Guardada")
    '
99  Command5.Enabled = True

End Sub

Sub Command6_Click()
' VALIDACIONES
    Command4.Enabled = False
    HOII% = HOY(HOS$)
    Label15.Caption = ""
    If UCase$(Left$(TRIM(txtServer), 4)) <> "SMTP" Then
       Call COMUNI("Servidor Smtp No Válido")
       GoTo 99
    End If
    '
    If TRIM$(txtFromName) = "" Then
       Call COMUNI("Falta Ingresar Nombre")
       GoTo 99
    End If
    '
    If ckPopLogin.Value > 0 And TRIM$(txtPopServer) = "" Then
       Call COMUNI("Falta Ingresar Pop3")
       GoTo 99
    End If
    '
    If ckHtml.Value > 0 And TRIM$(txtMsg) = "" Then
       Call COMUNI("Falta Ingresar Texto HTML")
       GoTo 99
    End If
    '
    If TRIM$(txtUserName) = "" Or TRIM$(txtPassword) = "" Then
       Call COMUNI("Falta Ingresar Usuario o Password")
       GoTo 99
    End If
    '
    If ckLogin.Value <> 1 Then
       Call COMUNI("Falta Tildar Casilla de Login")
       GoTo 99
    End If
    '
    MAI_SSMTP$ = txtServer                       ' SMTP
    MAI_SNAME$ = txtFromName                     ' SENDER NAME
    MAI_SMAIL$ = txtFrom                         ' SENDER EMAIL
    ' SENDMAIL09.txtPopServer = txtPopServer ' POP 3
    MAI_TEXTO$ = txtMsg
    MAI_ADJUN$(0) = txtAttach                    ' ADJUNTOS
    'SENDMAIL09.ckHtml.Value = ckHtml.Value ' OPTION HTML
    'SENDMAIL09.ckLogin.Value = ckLogin.Value ' OPTION LOGIN
    'SENDMAIL09.ckPopLogin.Value = ckPopLogin.Value ' OPTION POP3
    MAI_UNAME$ = txtUserName                     ' CUENTA DE USUARIO
    MAI_PWORD$ = txtPassword                     ' CLAVE
    MAI_DESTI$ = "FLEXOFT"
         If txtsubjet = "" Then txtsubjet = "TEST"
    MAI_ASUNT$ = txtsubjet
    MAI_DIREL$ = "asistec@flexoft.com.ar"
    Call GENERAMAIL
    '
99  Command4.Enabled = True
End Sub

Sub Command7_Click()
    '
    CUENTACT& = 0
    For KKI& = 1 To ULTREG&("LISMAIL")
       XX$ = REGLEIDO$("LISMAIL", KKI&)
       If TRIM$(Mid$(XX$, 1, 64)) = txtFrom.TEXT Then
          CUENTACT& = KKI&
          Exit For
       End If
    Next KKI&
    '
    CUENTA& = CUENTACT& + 1
    CANVUEL& = 0
10  If CUENTA& > ULTREG&("LISMAIL") Then CUENTA& = 1
    '
    XX$ = REGLEIDO$("LISMAIL", CUENTA&)
    txtFrom.TEXT = TRIM$(Mid$(XX$, 1, 64)) 'MAIL
    txtUserName.TEXT = TRIM$(Mid$(XX$, 65, 32))
    txtPassword.TEXT = TRIM$(Mid$(XX$, 97, 16))
    txtServer.TEXT = TRIM$(Mid$(XX$, 113, 32))
    STATCUE% = CVI(Mid$(XX$, 145, 2))
    If STATCUE% < 1 Then
       CUENTA& = CUENTA& + 1
       CANVUEL& = CANVUEL& + 1
       If CANVUEL& < 64 Then GoTo 10
    End If
    '
End Sub

Private Sub Command8_Click()
    '
    Command8.Enabled = False
    If COMALTER%("Confirma DesHabilitar Cuenta '" + txtFrom + "' ?\\No, Mantenerla\Si, Des-Habilitarla") = 2 Then
        CUENTACT& = 0
        URE& = ULTREG&("LISMAIL")
        For KKI& = 1 To URE&
           XX$ = REGLEIDO$("LISMAIL", KKI&)
           If TRIM$(Mid$(XX$, 1, 64)) = txtFrom.TEXT Then
              CUENTACT& = KKI&
              GoTo 20
           End If
        Next KKI&
        Call COMUNI("Cuenta no Encontrada en la Base de Datos")
        GoTo 99
        '
20      XX$ = REGLEIDO$("LISMAIL", CUENTACT&)
        Call REPLA(XX$, MKI$(0), 145, 2)
        Call GRAREG("LISMAIL", XX$, CUENTACT&)
        '
        CUENTA& = CUENTACT& + 1
        If CUENTA& > ULTREG&("LISMAIL") Then CUENTA& = 1
        '
        XX$ = REGLEIDO$("LISMAIL", CUENTA&)
        txtFrom.TEXT = TRIM$(Mid$(XX$, 1, 64)) 'MAIL
        txtUserName.TEXT = TRIM$(Mid$(XX$, 65, 32))
        txtPassword.TEXT = TRIM$(Mid$(XX$, 97, 16))
        txtServer.TEXT = TRIM$(Mid$(XX$, 113, 32))
    End If
    '
99  Command8.Enabled = True
    '
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    RESENDMA09.Show 1
    Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    EPAC$ = EMPREAC$
    If EPAC$ <> "" Then
      Caption = Caption + " - " + EMPREAC$
    End If
    '
    Call CREACAMPO("", "BSEGCOM", "FULTIEN", 3, 2, 1)
    Call PoneEnCeroNull("BSEGCOM", "FULTIEN", RAFE&)
    'Visible = True
    Refresh
    DoEvents
    '
    Screen.MousePointer = 11
    ENVIOMASIVO% = 1
    TODOS% = 0: SELSTRING$ = String$(SELDESTI09.List5.ListCount, "0")
    List5.Clear
    If SELDESTI09.Check1.Value = 1 Then
        SELSTRING$ = String$(SELDESTI09.List5.ListCount, "1")
        TODOS% = 1
      Else
        For KKI% = 1 To SELDESTI09.List5.ListCount
          If SELDESTI09.List5.Selected(KKI% - 1) = True Then Mid$(SELSTRING$, KKI%) = "1"
        Next KKI%
    End If
    '
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    '
    Dim CTACSEL As ADODB.Recordset
    Set CTASEL = New ADODB.Recordset
    SQLX$ = "SELECT * FROM BSEGCOM WITH(NOLOCK) ORDER BY CODI_CLI ASC"
    Set CTASEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    YAENVIOELPRIMER% = 0
    With CTASEL
      Do While Not .EOF
        OPENFORMS = DoEvents
        TTXX$ = AJUSTI$(SAFETEXT$(!CODI_CLI), 15) + SAFETEXT$(!Raso_Cli)
        PERLIST$ = SAFETEXT$(!GRUPOS)
        FULEN% = XVALO(!FULTIEN)
        If TODOS% = 1 Then GoTo 15
        '
        For KKI% = 1 To Len(PERLIST$)
           PERLI% = Asc(Mid$(PERLIST$, KKI%, 1))
           If PERLI% > 0 Then
             If Mid$(SELSTRING$, PERLI%, 1) = "1" Then
               CODCLISE$ = SAFETEXT$(!CODI_CLI)
               GoTo 15
             End If
           End If
        Next KKI%
        GoTo 35
        '
15      If FULEN% > 0 Then
           ' YA FUE ENVIADO
           List4.AddItem TTXX$
           List4.ListIndex = List4.ListCount - 1
           Label14 = TRIM$(Str$(List4.ListCount))
           GoTo 39
        End If
        '
25      ' agregar a lists de envios
        GoSub AGRELISTEN
        GoTo 39
        '
35      ' NO ESTA SELECCIONADO
        List3.AddItem TTXX$
        List3.ListIndex = List3.ListCount - 1
        Label9 = TRIM$(Str$(List3.ListCount))
        GoTo 39
        '
39    .MoveNext
      Loop
    End With
    CTASEL.Close
    Set CTASEL = Nothing
    '
    On Error Resume Next
    List1.TopIndex = 0
    List1.ListIndex = -1
    List2.TopIndex = 0
    List2.ListIndex = -1
    List3.TopIndex = 0
    List3.ListIndex = -1
    List4.TopIndex = 0
    List4.ListIndex = -1
    On Error GoTo 0
    '
    ' PRESENTA LOS DATOS
    txtServer.TEXT = Control("", "SMTPSRVR") 'smtpserver
    txtFromName.TEXT = Control("", "FROMNAME")
    txtFrom.TEXT = Control("", "FROMMAIL")
    txtPopServer.TEXT = Control("", "POPSERV")
    txtUserName.TEXT = Control("", "USUARIO")
    txtPassword.TEXT = Control("", "KLAVE")
    ckLogin.Value = 1
    '
    Screen.MousePointer = 1
Exit Sub

AGRELISTEN:
    '
    KOKLI$ = TRIM$(Left$(TTXX$, 12))
    FISECOM09.Label26 = KOKLI$
    '
    EMPRA$ = REPLACEN$(TRIM$(FISECOM09.Text4))
    YENVI% = 0
    For KI% = 1 To FISECOM09.Text6.Count
       APENO$ = REPLACEN$(TRIM$(FISECOM09.Text6(KI% - 1)))
       DIREL$ = REPLACEN$(TRIM$(FISECOM09.Text8(KI% - 1)))
       PUENVI% = 1
       '
       ' CAPTURA MAIL NO VÁLIDO
       If InStr(DIREL$, "@") < 2 Then PUENVI% = 0
       If InStr(DIREL$, "@") > Len(DIREL$) - 1 Then PUENVI% = 0
       '
       HC% = 0
       If Right$(DIREL$, 1) = "." Or Right$(DIREL$, 1) = ";" Then
           DIREL$ = Left$(DIREL$, Len(DIREL$) - 1)
           HC% = 1
       End If
       If Left$(UCase$(DIREL$), 5) = "EMAIL" Then
           DIREL$ = Mid$(DIREL$, 6)
           HC% = 1
       End If
       If Left$(DIREL$, 1) = ";" Then
           DIREL$ = Mid$(DIREL$, 2)
           HC% = 1
       End If
          '
       If PUENVI% > 0 Then
          TTXX$ = AJUSTI$(KOKLI$, 15) + AJUSTI$(APENO$, 60) + AJUSTI$(EMPRA$, 60) + DIREL$
          List1.AddItem TTXX$
          List1.ListIndex = List1.ListCount - 1
          Label4 = TRIM$(Str$(List1.ListCount))
          YENVI% = 1
       End If
    Next KI%
    '
    If YENVI% < 1 Then
       PUENVI% = 1
       APENO$ = ""
       DIREL$ = FISECOM09.Text5
       If InStr(DIREL$, "@") < 2 Then PUENVI% = 0
       If InStr(DIREL$, "@") > Len(DIREL$) - 1 Then PUENVI% = 0
       '
       HC% = 0
       If Right$(DIREL$, 1) = "." Or Right$(DIREL$, 1) = ";" Then
           DIREL$ = Left$(DIREL$, Len(DIREL$) - 1)
           HC% = 1
       End If
       If Left$(UCase$(DIREL$), 5) = "EMAIL" Then
           DIREL$ = Mid$(DIREL$, 6)
           HC% = 1
       End If
       If Left$(DIREL$, 1) = ";" Then
           DIREL$ = Mid$(DIREL$, 2)
           HC% = 1
       End If
          '
       If PUENVI% > 0 Then
          TTXX$ = AJUSTI$(KOKLI$, 15) + AJUSTI$(APENO$, 60) + AJUSTI$(EMPRA$, 60) + DIREL$
          List1.AddItem TTXX$
          List1.ListIndex = List1.ListCount - 1
          Label4 = TRIM$(Str$(List1.ListCount))
          YENVI% = 1
       End If
    End If
Return

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Hide
End Sub


Private Sub Label11_Change()
  If Right$(Label11, 1) = "\" Then
    Label11 = Left$(Label11, Len(Label11) - 1)
  End If
End Sub

Private Sub Label6_Change()
   Label4 = XVALO(Label4) - 1
End Sub

Private Sub List1_DblClick()
   '
   KOKLI$ = TRIM$(Left$(List1.TEXT, 12))
   FISECOM09.Label26 = KOCLI$
   Hide
   '
End Sub

Private Sub List4_DblClick()
    '
    KU& = List4.ListIndex
    TTXX$ = List4.TEXT
    List4.RemoveItem KU&
    List1.AddItem TTXX$, 0
    '
End Sub

Private Sub Text1_Change()
   '
   File1.Pattern = Text1
   '
End Sub

Private Sub Timer1_Timer()
   '
   Static CTX%, CUENTA&, SEGDESULTEN&, U&, STRIX$, CONTA%, HOII%
   Static NOENVIADOS&
   Static LASTSEN$
   '
   If CTX% < 1 Then
      Call CREACAMPO("", "BSEGCOM", "FULTIEN", 3, 2, 1)
      U& = 0: CONTA% = 0
      SEGDESULTEN& = 10                           ' para que lo envie directamente
      HOII% = HOY(HOS$)
      STRIX$ = "mailto:flexoft@yahoo.com.ar?subject=Eliminacion de Lista de Envios&body=Eliminar Contacto: "
      CUENTA& = 1
      CANVUEL& = 0
5     CANVUEL& = CANVUEL& + 1
      If CANVUEL& > ULTREG&("LISMAIL") Then GoTo 9
      GoSub LEEDATCUENTA
      If STATCUE% < 1 Then
         CUENTA& = CUENTA& + 1
         If CUENTA& > ULTREG&("LISMAIL") Then CUENTA& = 1
         GoTo 5
      End If
9     CTX% = 1
   End If
   '
   SEGDESULTEN& = SEGDESULTEN& + 1
   Label10 = SEGDESULTEN&: Label10.Refresh
' 10 If SEGDESULTEN& < 10 Then Exit Sub     ' DEPUES AJUSTAR AL INTERVALO
   '
' MsgBox SEGDESULTEN& & "..........." & EXISTE%(ARENVI$) & "············" & ARENVI$
10 If SEGDESULTEN& < 10 Then
     If ARENVI$ <> "" Then
       If EXISTE%(ARENVI$) > 0 Then Exit Sub      ' DEPUES AJUSTAR AL INTERVALO
     End If
   End If
   SEGDESULTEN& = 0
   Label10 = SEGDESULTEN&: Label10.Refresh
   
   ' SI SE ACUMULARON MÁS DE 3 NO ENVIADOS, CAMBIA LA CUENTA
   If XVALO(RESENDMA09.Label4) - NOENVIADOS& > 3 Then
      If NOENVIADOS& > 0 Then
         NOENVIADOS& = XVALO(RESENDMA09.Label4)
         CANVUEL& = 0
15       CANVUEL& = CANVUEL& + 1
         If CANVUEL& > ULTREG&("LISMAIL") Then GoTo 9
         CUENTA& = CUENTA& + 1
         If CUENTA& > ULTREG&("LISMAIL") Then CUENTA& = 1
         GoSub LEEDATCUENTA
         If STATCUE% < 1 Then
            CUENTA& = CUENTA& + 1
            GoTo 15
         End If
      End If
      NOENVIADOS& = XVALO(RESENDMA09.Label4)
   End If
   '
   If Check3.Value = 1 Then
      WindowState = 1
   End If
   '
20 MODENVI$ = ""
   If Check1.Value = 1 Then MODENVI$ = MODENVI$ + "/AUTO"
   If Check2.Value = 1 Then MODENVI$ = MODENVI$ + "/TEST"
   '
   If U& < List1.ListCount Then
        U& = U& + 1
        If Label15.Caption = "STOP" Then GoTo 99
        If U& > List1.ListCount Then GoTo 95 ' MENSAJE DE FIN DE PROCESO
        '
        List1.ListIndex = U& - 1
        '
        TTXX$ = List1.List(U& - 1)
'''        If TRIM$(Left$(TTXX$, 15)) <> KOKLI$ Then     ' cambio la cuenta principal
'''            KOKLI1$ = TRIM$(Left$(TTXX$, 15))
'''            If XVALO(VALOBADA("BSEGCOM", "FULTIEN", "CODI_CLI = '" & KOKLI1$ & "'")) > 0 Then
'''               ' ESTO ATAJA EL ENVIO DESDE OTRA ESTACION DE TRABAJO
'''               List2.AddItem List1.TEXT
'''               List2.ListIndex = List2.ListCount - 1
'''               On Error Resume Next
'''               List2.TopIndex = List2.ListCount - 6
'''               On Error GoTo 0
'''               Label6 = List2.ListCount
'''               List1.RemoveItem (List1.ListIndex)
'''               Label4 = List1.ListCount
'''               U& = U& - 1
'''               GoTo 20
'''            End If
'''        End If
'''        '
        KOKLI$ = TRIM$(Left$(TTXX$, 15))
        APENO$ = TRIM$(Mid$(TTXX$, 16, 60))
        EMPRA$ = TRIM$(Mid$(TTXX$, 76, 60))
        DIREL$ = TRIM$(Mid$(TTXX$, 136))
        '
        MUESTRA$ = ""
        If TRIM$(APENO$) <> "" Then MUESTRA$ = MUESTRA$ + " " + TRIM$(APENO$)
        If TRIM$(EMPRA$) <> "" Then
          If TRIM$(MUESTRA$) <> "" Then
            MUESTRA$ = MUESTRA$ + " -"
          End If
          MUESTRA$ = MUESTRA$ + " " + EMPRA$
        End If
        MUESTRA$ = REPLACEN$(MUESTRA$)
        MUESTRA$ = REPLACAR$(MUESTRA$, "(", " ")
        MUESTRA$ = REPLACAR$(MUESTRA$, ")", " ")
        '
        ' Ajuste del codigo html para la devolucion del mail con remover
        MAI_SSMTP$ = txtServer                       ' SMTP
        MAI_SNAME$ = txtFromName                     ' SENDER NAME
        MAI_SMAIL$ = txtFrom                         ' SENDER EMAIL
        MAI_UNAME$ = txtUserName                     ' CUENTA DE USUARIO
        MAI_PWORD$ = txtPassword                     ' CLAVE
        '
        MAI_DESTI$ = TRIM$(APENO$)
           If MAI_DESTI$ <> "" Then MAI_DESTI$ = MAI_DESTI$ + " - "
        MAI_DESTI$ = MAI_DESTI$ + TRIM$(EMPRA$)
        MAI_DIREL$ = DIREL$
        MAI_ASUNT$ = "Att: " + MUESTRA$
           TXTO$ = txtMsg
           TXTO1$ = REPLACAR(TXTO$, "mailto:remove", STRIX$ & " " & EMPRA$ & "(" & DIREL$ & ")")
        MAI_TEXTO$ = TXTO1$ '
        MAI_FTEXT$ = "PLAIN"
           If ckHtml.Value = 1 Then MAI_FTEXT$ = "HTML"
        MAI_ADJUN$(0) = txtAttach                    ' ADJUNTOS
        '
        Call GENERAMAIL(MODENVI$)
        YENVI% = 1
        SEGDESULTEN& = 0
        '
        CONTA% = CONTA% + 1 ' CONTADOR ENVIA 40 MENSAJES
        If CONTA% >= 32 Then      ' ENVÍA 32 POR BLOQUE
            ' ENVÍA MAIL DE CONTROL
            MAI_ASUNT$ = "Control " + txtFrom.TEXT + " " + Time$ + " - Att: " + MUESTRA$
            MAI_DIREL$ = "asistec@flexoft.com.ar"
            Call GENERAMAIL(MODENVI$)
            SEGDESULTEN& = 0
            '
            CONTA% = 0: CANVUEL& = 0
            FIN& = ULTREG&("LISMAIL")
84          CANVUEL& = CANVUEL& + 1
            CUENTA& = CUENTA& + 1        ' pasa a la cuenta siguiente
            If CUENTA& > FIN& Then CUENTA& = 1
            GoSub LEEDATCUENTA   ' DE LA CUENTA CUENTA&
            If CANVUEL& < 64 Then
               If STATCUE% < 1 Then GoTo 84
            End If
            '
        End If
        '
95      ' MARCAR COMO PROCESADO
        HOII% = HOY(HO$)
        Call ACTUREGISTRO("BSEGCOM", "FULTIEN", "CODI_CLI = '" & KOKLI$ & "'", HOII%, RAFE&)
        '
        List2.AddItem List1.TEXT
        List2.ListIndex = List2.ListCount - 1
        On Error Resume Next
        List2.TopIndex = List2.ListCount - 6
        On Error GoTo 0
        Label6 = List2.ListCount
        List1.RemoveItem (List1.ListIndex)
        Label4 = List1.ListCount
        U& = U& - 1
        GoTo 10                                ' YA FUE ENVIADO
        '
   End If
   '
   WindowState = 0
   OPENFORMS = DoEvents
   Call COMUNI("Fin de Proceso")
   '
99 Command4.Enabled = True
   CTX% = 1
   Timer1.Enabled = False
Exit Sub
   '
   '
LEEDATCUENTA:
    XX$ = REGLEIDO$("LISMAIL", CUENTA&)
    txtFrom.TEXT = TRIM$(Mid$(XX$, 1, 64)) 'MAIL
    txtUserName.TEXT = TRIM$(Mid$(XX$, 65, 32))
    txtPassword.TEXT = TRIM$(Mid$(XX$, 97, 16))
    txtServer.TEXT = TRIM$(Mid$(XX$, 113, 32))
    STATCUE% = CVI(Mid$(XX$, 145, 2))
Return
'
'
End Sub

Sub RECUERROR(OKX%)
   '
   ' trata de apagar Outlook Express
   OKX% = 0
   TINI = Timer
3   On Error Resume Next
   AppActivate "Outlook Express", False
   If Err Then GoTo 5
   On Error GoTo 0
   SK$ = "%" + "C"
   On Error Resume Next
   SendKeys Chr$(13), True
   On Error GoTo 0
   On Error Resume Next
   SendKeys SK$, True
   On Error GoTo 0
   If Timer < TINI + 10 Then GoTo 3
   '
   ' CIERRA LA BANDEJA DE ENTRADA
5  OKX% = 0
   TINI = Timer
10 On Error Resume Next
   AppActivate "Bandeja de Entrada", False
   If Err Then
      If Timer < TINI + 60 Then GoTo 10
      GoTo 90
   End If
   '
   On Error GoTo 0
   SK$ = "%" + "{F4}"
   On Error Resume Next
   SendKeys SK$, True
   If Err Then
      On Error GoTo 0
      GoTo 90
   End If
   '
   On Error Resume Next
   AppActivate "Bandeja de Entrada", False
   If Err Then
      On Error GoTo 0
      OKX% = 1
      GoTo 90
   End If
   If Timer < TINI + 60 Then GoTo 10
   '
90 End Sub

Private Sub Timer2_Timer()
     '
     Static U&, RETVAL, VUELTA%
     '
     If (IsNull(RETVAL) = True Or IsEmpty(RETVAL) = True) Then GoTo 10 ' ES LA PRIMERA PASADA
     '
     ' VERIFICA SI TODAVÍA SE ESTÁ EJECUTANDO OUTLOOK EXPRESS
     On Error Resume Next
     AppActivate RETVAL
     If Err Then
        ' YA SE CERRO EL OUTLOOK EXPRESS ANTERIOR
        On Error GoTo 0
        GoTo 10
     End If
     If VUELTA% < 3 Then
        SK$ = "%{F4}{ENTER}{ENTER}"     ' para terminar de enviar
        SendKeys SK$, False            ' con false para que lo haga inmediatamente
        VUELTA% = VUELTA% + 1
        Exit Sub
     End If
     '
10   VUELTA% = 0
     '
11   On Error Resume Next
     AppActivate "Mensaje nuevo", False
     If Err Then
        On Error GoTo 0
        GoTo 12
     End If
     On Error GoTo 0
     SK$ = "Att: " + MUESTRA$ + "%AE" + "%" + "{F4}"  ' para terminar de enviar el anterior
     SendKeys SK$, True             ' con false para que lo haga inmediatamente
     If VUEL2% < 3 Then
        VUEL2% = VUEL2% + 1
        GoTo 11
     End If
     '
12   If U& < 1 Then U& = 0
     HOII% = HOY(HOS$)
     U& = U& + 1
     If U& > List1.ListCount Then
        Call COMUNI("Fin de Transmisión")
        Timer2.Enabled = False
        Exit Sub
     End If
     '
     GoSub LEEDATOS
     '
20   TINI = Timer
     RETVAL = Shell("C:\Archivos de Programa\Outlook Express\msimn.exe", 1)
     On Error Resume Next
     AppActivate RETVAL
     If Err > 0 Then
        NERRO = Err.Number
        DERRO$ = Err.Description
        On Error GoTo 0
        If COMALTER%("Se ha Producido Error " + TRIM$(Str$(NERRO)) + "\'" + DERRO$ + "'\\Continuar\Terminar") = 1 Then
          GoTo 20
        End If
        Call FINAL("")
     End If
     On Error GoTo 0
     '
     While Timer < TINI + 2: OPENFORMS = DoEvents: Wend
     SendKeys "{ESC}", True
     SendKeys "{ESC}", True
     SendKeys "{ESC}", True
     '
     DIREL2$ = ""
     PPXX% = InStr(DIREL$, "/")
     If PPXX% > 0 Then
        DIREL2$ = TRIM$(Mid$(DIREL$, PPXX% + 1))
        DIREL$ = TRIM$(Left$(DIREL$, PPXX% - 1))
     End If
     '
     ' 1.- OUTLOOK EXPRESS - HERRAMIENTAS - OPCIONES - REDACTAR
     ' 2.- TILDAR CORREO - SELECCIONAR ARCHIVO DEL NEWSLETTER
     ' 3.- CLICK SOBRE ACEPTAR
     MUESTRA$ = REPLACAR$(MUESTRA$, "(", " ")
     MUESTRA$ = REPLACAR$(MUESTRA$, ")", " ")
     SK$ = "%ANM" + DIREL$ + "{TAB}{TAB}+Att: " + MUESTRA$ + "%AE" + "%" + "{F4}"
     SendKeys SK$, True
     While Timer < TINI + 1: OPENFORMS = DoEvents: Wend
     '
30   XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
     Call REPLA(XX$, MKI$(HOII%), 511, 2)
     Call GRAREG("LIDIRMAI", XX$, REMAI&)
     '
     Label4 = TRIM$(Str$(Val(Label4) - 1))
     List2.AddItem List1.TEXT
     List2.ListIndex = List2.ListCount - 1
     Label6.Caption = TRIM$(Str$(List2.ListCount))
     '
     If Val(Text2) < 99 Then
       If List2.ListCount Mod Val(Text2) = 0 Then
         OPX% = COMALTER%("Opciones de Procesamiento:\\Continuar\Interrumpir")
         If OPX% = 2 Then Label15 = "STOP"
       End If
     End If
     '
     If Label15.Caption = "STOP" Then GoTo 98
     '
   Exit Sub
   '
98 Call CIERRARCH("LIDIRMAI")
   Call BAJALDISCO
   Timer2.Enabled = False
   Exit Sub
   '
   '
LEEDATOS:
    List1.ListIndex = U& - 1
    REMAI& = Val(Mid$(List1.List(U& - 1), 81))
    XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
    APENO$ = TRIM$(Mid$(XX$, 65, 40))
    EMPRA$ = TRIM$(Mid$(XX$, 173, 64))
    If InStr(EMPRA$, "!") > 0 Then
      EMPRA$ = REPLACAR$(EMPRA$, "!", " ")
      Call REPLA(XX$, EMPRA$, 173, 64)
      Call GRAREG("LIDIRMAI", XX$, REMAI&)
    End If
    DIREL$ = REPLACEN$(TRIM$(Mid$(XX$, 105, 60)))
    '
    HC% = 0
    If Right$(DIREL$, 1) = "." Then
      DIREL$ = Left$(DIREL$, Len(DIREL$) - 1)
      HC% = 1
    End If
    If Left$(UCase$(DIREL$), 5) = "EMAIL" Then
      DIREL$ = Mid$(DIREL$, 6)
      HC% = 1
    End If
    If Left$(DIREL$, 1) = ";" Then
      DIREL$ = Mid$(DIREL$, 2)
      HC% = 1
    End If
    If HC% = 1 Then
      Call REPLA(XX$, TRIM$(DIREL$), 105, 60)
      Call GRAREG("LIDIRMAI", XX$, REMAI&)
    End If
    '
    MUESTRA$ = ""
    If TRIM$(APENO$) <> "" Then MUESTRA$ = MUESTRA$ + " " + TRIM$(APENO$)
    If TRIM$(EMPRA$) <> "" Then
      If MUESTRA$ <> "" Then
        MUESTRA$ = MUESTRA$ + " -"
      End If
      MUESTRA$ = MUESTRA$ + " " + EMPRA$
    End If
    MUESTRA$ = REPLACEN$(MUESTRA$)
Return
'
End Sub

Private Sub Timer4_Timer()
   Static i&
   '
   If i& < 1 Then
     i& = 0
   End If
   '
   'TIMER INTERVAL = 30000
   i& = i& + 1
   If i& < 300 Then ' PARA QUE PASE CUANDO SEAN LOS 5 MINUTOS
     lblProgress = "Pausa - " & 300 - i&
     Exit Sub
   End If
   Timer4.Enabled = False
   i& = 0
   Command4.Enabled = True
   Command4_Click
End Sub


