VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FSCASHFLOW 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   16455
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
   ScaleHeight     =   8055
   ScaleWidth      =   16455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asociación a Bancos de Rubros Fijos (fact.x Cobrar, Cheq.,Pagos  etc)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2500
      Left            =   120
      TabIndex        =   47
      Top             =   5520
      Width           =   6615
      Begin VB.TextBox Text14 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton Command47 
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
         Height          =   300
         Left            =   600
         TabIndex        =   50
         Top             =   2040
         Width           =   175
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   49
         ToolTipText     =   "listabconta"
         Top             =   2040
         Width           =   4605
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5760
         TabIndex        =   48
         ToolTipText     =   "Edita la Lista"
         Top             =   2040
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid MSF6 
         Height          =   1515
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   2672
         _Version        =   393216
         Cols            =   4
         RowHeightMin    =   300
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Roboto Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "N° de Lista"
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
         Index           =   9
         Left            =   120
         TabIndex        =   53
         Top             =   1800
         Width           =   2925
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tiempos y Comisiones de T.Crédito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   2
      Left            =   6840
      TabIndex        =   36
      Top             =   5520
      Width           =   8535
      Begin VB.TextBox Text12 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton Command10 
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
         Height          =   300
         Left            =   600
         TabIndex        =   39
         Top             =   1680
         Width           =   175
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   38
         ToolTipText     =   "confimeco"
         Top             =   1680
         Width           =   4000
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         TabIndex        =   37
         ToolTipText     =   "Edita la Lista"
         Top             =   1680
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid MSF5 
         Height          =   1125
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   1984
         _Version        =   393216
         Cols            =   4
         RowHeightMin    =   300
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Roboto Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "N° de Lista "
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
         Index           =   6
         Left            =   120
         TabIndex        =   42
         Top             =   1440
         Width           =   2925
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Saldos a ingresar de Cuentas Bancarias"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   1
      Left            =   6840
      TabIndex        =   29
      Top             =   3000
      Width           =   8535
      Begin VB.CommandButton Command5 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         TabIndex        =   33
         ToolTipText     =   "Edita la Lista"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   32
         ToolTipText     =   "salcuban"
         Top             =   2040
         Width           =   4000
      End
      Begin VB.CommandButton Command3 
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
         Height          =   300
         Left            =   600
         TabIndex        =   31
         Top             =   2040
         Width           =   175
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2040
         Width           =   495
      End
      Begin MSFlexGridLib.MSFlexGrid MSF4 
         Height          =   1485
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   2619
         _Version        =   393216
         Cols            =   4
         RowHeightMin    =   300
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Roboto Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "N° de Lista "
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
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   1800
         Width           =   2925
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fechas Configurables de Pagos de I.V.A. - IIBB, etc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   6615
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   45
         Text            =   "Combo1"
         Top             =   1800
         Width           =   3315
      End
      Begin VB.TextBox TXTPORCIIBB 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   5880
         TabIndex        =   43
         Text            =   " "
         Top             =   1800
         Width           =   540
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Command45 
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
         Height          =   300
         Left            =   600
         TabIndex        =   25
         Top             =   2400
         Width           =   175
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   24
         ToolTipText     =   "vencimiva"
         Top             =   2400
         Width           =   4725
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5760
         TabIndex        =   23
         ToolTipText     =   "Edita la Lista"
         Top             =   2400
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   1485
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   2619
         _Version        =   393216
         Cols            =   4
         RowHeightMin    =   300
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Roboto Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tasa (%)"
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
         Index           =   8
         Left            =   5040
         TabIndex        =   46
         Top             =   1860
         Width           =   765
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cta.de IIBB:"
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
         Index           =   7
         Left            =   120
         TabIndex        =   44
         Top             =   1860
         Width           =   1605
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "N° de Lista"
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
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   2160
         Width           =   2925
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ingresos y Egresos Fijos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Index           =   0
      Left            =   6840
      TabIndex        =   15
      Top             =   50
      Width           =   8535
      Begin VB.TextBox Text9 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton Command44 
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
         Height          =   300
         Left            =   600
         TabIndex        =   18
         Top             =   2280
         Width           =   175
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "ingrsali"
         Top             =   2280
         Width           =   4000
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         TabIndex        =   16
         ToolTipText     =   "Edita la Lista"
         Top             =   2280
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   1725
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   3043
         _Version        =   393216
         Cols            =   4
         RowHeightMin    =   300
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Roboto Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "N° de Lista "
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
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   2925
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ctas. Contables que Reciben Percepciones o Retenciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   120
      TabIndex        =   7
      Top             =   50
      Width           =   6615
      Begin VB.CommandButton Command6 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5760
         TabIndex        =   14
         ToolTipText     =   "Edita la Lista"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "listabconta"
         Top             =   2160
         Width           =   4725
      End
      Begin VB.CommandButton Command4 
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
         Height          =   300
         Left            =   600
         TabIndex        =   11
         Top             =   2160
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2160
         Width           =   495
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   1605
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   2831
         _Version        =   393216
         Cols            =   4
         RowHeightMin    =   300
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Roboto Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "N° de Lista"
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
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   2925
      End
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   11040
      TabIndex        =   5
      Text            =   " "
      Top             =   7580
      Width           =   420
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   15480
      TabIndex        =   3
      Text            =   " "
      Top             =   7580
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   7530
      Left            =   15540
      ScaleHeight     =   7470
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   0
      Width           =   1170
      Begin VB.CommandButton command1 
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
         Height          =   500
         Left            =   105
         Picture         =   "FSCASHFLOW.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   90
         Width           =   645
      End
      Begin VB.CommandButton Command2 
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
         Height          =   500
         Left            =   120
         Picture         =   "FSCASHFLOW.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5160
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "FSCASHFLOW.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5760
         Width           =   1515
      End
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo de Efectivización de Facturas a Cobrar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   6840
      TabIndex        =   6
      Top             =   7680
      Width           =   4125
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo de Acreditación de Cheques a Cobrar:"
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
      Index           =   1
      Left            =   11520
      TabIndex        =   4
      Top             =   7680
      Width           =   4125
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuConsultasListados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuReportesVarios 
         Caption         =   "Reportes Varios"
      End
   End
   Begin VB.Menu mnuconfig 
      Caption         =   "Configuraciones"
      Begin VB.Menu mnuTabaux 
         Caption         =   "Tablas Auxiliares"
      End
      Begin VB.Menu mnuAsociarReportes 
         Caption         =   "Vincular Reportes al Formulario"
      End
   End
End
Attribute VB_Name = "FSCASHFLOW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CARGARDATOSFIJOSCASHFLOW(MSF As MSFlexGrid, IDL, DESC$)
   MSF.Rows = 1
   
   TITULOS$ = "Fact.x Cobrar;Compras;Cheq x Cobrar"
   Call TEXTOLOTES$(TITULOS$, ";")
   On Error GoTo ERROR_GUARDAR
   
   FLEXCONN.BeginTrans
   MBORRA% = 0
   For i& = 1 To UBound(CADENATEX$)
       MSF.Rows = i& + 1
       MSF.TextMatrix(i&, 1) = CADENATEX$(i&)
       SQLX$ = "INSERT INTO DATASQL "
       SQLX$ = SQLX$ + " (CAMPO1,ID_LISTA,LISTA,CODIEMPR,MARBORRA)"
       SQLX$ = SQLX$ + " VALUES('" & CADENATEX$(i&) & "'"
       SQLX$ = SQLX$ + " ," & IDL
       SQLX$ = SQLX$ + " ,'" & DESC$ & "'"
       SQLX$ = SQLX$ + " ," & CodiEmp%
       SQLX$ = SQLX$ + " ," & MBORRA%
       
       SQLX$ = SQLX$ + ")"
       FLEXCONN.Execute (SQLX$)
   Next i&
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
    Else
      FLEXCONN.CommitTrans
    End If
    

   
   
   

   
End Sub

Sub CargaTablaMediosCobranza()
   Condi$ = " LISTA =   'CTASCBMC'"
   IDLI& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", Condi$, "NOMESS"))
   If IDLI& > 0 Then
    For U& = 1 To ULTREG&("MECOBRA")
       X$ = REGLEIDO$("MECOBRA", U&)
       CDX% = CVI(Left$(X$, 2))
       If CDX% > 0 Then
             J& = J& + 1
             DAT1$ = TRIM$(Mid$(X$, 3, 12))
             DAT2$ = TRIM$(Mid$(X$, 45, 12))
             DAT3$ = TRIM$(Mid$(X$, 57, 4))
             SQLX$ = "SELECT * FROM DATASQL WHERE ID_LISTA = " & IDLI& & " AND  CAST(CAMPO1 AS INT) = " & CDX%
             
             Dim RS As ADODB.Recordset
             Set RS = New ADODB.Recordset
             RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
             With RS
               If .EOF Then 'cambie para que actulice siempre aunque exista
                 .AddNew
               End If
                 !CAMPO1 = TRIM$(Str(CDX%))
                 !CAMPO2 = DAT1$
                 !CAMPO3 = DAT2$
                 !CAMPO4 = DAT3$
                 !ID_Lista = IDLI&
                 !LISTA = "CTASCBMC"
                 !CODIEMPR = CodiEmp%
                 .Update
               
             End With
             RS.Close
             Set RS = Nothing
           End If
        Next U&
   End If



End Sub

Sub CargaTablasVencImp()
  If XVALO(Text10) < 1 Then Exit Sub
  If TRIM$(Text2) = "" Then Exit Sub
  '

  IDL = XVALO(Text10)
  DESC$ = TRIM$(Text2)
  J& = 0
  Combo1.Clear
  Call CARGA_ENCABEZADO(FSCASHFLOW.MSF3, "Día V./A8;Imp./F6;Denominación/A32", FSCASHFLOW)
  SQLX$ = "SELECT CAMPO1, CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK) WHERE ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'"
  SQLX$ = SQLX$ + " order by campo1"
  
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
      J& = J& + 1
      MSF3.Rows = J& + 1
      MSF3.TextMatrix(J&, 1) = SAFETEXT$(!CAMPO1)
      MSF3.TextMatrix(J&, 2) = SAFETEXT$(!CAMPO2)
      MSF3.TextMatrix(J&, 3) = SAFETEXT$(!CAMPO3)
      'CARGA EL COMBO
      Call REPLA(TX$, MSF3.TextMatrix(J&, 2), 1, 8)
      Call REPLA(TX$, MSF3.TextMatrix(J&, 3), 10, 48)
      Combo1.AddItem TX$
      .MoveNext
    Loop
  End With
  Combo1.TEXT = CONTROL("CASHFLOW", "CTACASIB")
  

End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()

  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(rs1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.MSF2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1

  Text12 = XVALO(RESP_ZELE_VECT(1))
  Text11 = RESP_ZELE_VECT(2)


End Sub

Private Sub Command11_Click()
  Call ABM_TABLA(XVALO(Text14))
  CargaTablaAsocBancosARubrosFijos


End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "CLEABANC", Text6.TEXT)
    Call GRACONTROL("", "EFECFACT", Text7.TEXT)

    Call GRACONTROL("", "CTAREPER", Text1)
    Call GRACONTROL("", "DENREPER", Text3)
    
    Call GRACONTROL("", "CTAINGGAS", Text9)
    Call GRACONTROL("", "DENINGGAS", Text8)
    
    Call GRACONTROL("", "CTAFVENC", Text10)
    Call GRACONTROL("", "DENFVENC", Text2)
    '
    Call GRACONTROL("", "CTALCUBA", Text4)
    Call GRACONTROL("", "DENOCUBA", Text5)

    Call GRACONTROL("", "CTAACCOM", Text12)
    Call GRACONTROL("", "DENACCOM", Text11)
    
    Call GRACONTROL("", "CTAASBAN", Text14)
    Call GRACONTROL("", "DENASBAN", Text13)

    Call GRACONTROL("CASHFLOW", "CTACASIB", Combo1.TEXT)
    Call GRACONTROL("CASHFLOW", "TASCASIB", TRIM$(FORMATNUM$(XVALO(TXTPORCIIBB), "F10.2")))
    

    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completada.\Re-Inicie Ahora la Aplicación '" + UCase$(App.EXEName) + "'.")
    Call FINAL("")
    '
End Sub

Private Sub Command3_Click()
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(rs1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.MSF2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1

  Text4 = XVALO(RESP_ZELE_VECT(1))
  Text5 = RESP_ZELE_VECT(2)

End Sub
'

Private Sub Command4_Click()
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(rs1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.MSF2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1

  Text1 = XVALO(RESP_ZELE_VECT(1))
  Text3 = RESP_ZELE_VECT(2)
End Sub

Private Sub Command44_Click()
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(rs1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.MSF2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1

  Text9 = XVALO(RESP_ZELE_VECT(1))
  Text8 = RESP_ZELE_VECT(2)
  
End Sub

Private Sub Command45_Click()
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(rs1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.MSF2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1

  Text10 = XVALO(RESP_ZELE_VECT(1))
  Text2 = RESP_ZELE_VECT(2)

End Sub


Private Sub Command47_Click()
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(rs1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.MSF2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1

  Text14 = XVALO(RESP_ZELE_VECT(1))
  Text13 = RESP_ZELE_VECT(2)

End Sub

Private Sub Command5_Click()
   VENTANA.ValidarDuplicadoFila1 = "N"
   Call ABM_TABLA(XVALO(Text4))
  
   Call CargaTablaSaldoBancarias

End Sub

Private Sub Command6_Click()
   VENTANA.ValidarDuplicadoFila1 = "N"
   Call ABM_TABLA(XVALO(Text1))
   Call cargaTablasRetPerc

End Sub

Private Sub Command7_Click()
   VENTANA.ValidarDuplicadoFila1 = "N"
   Call ABM_TABLA(XVALO(Text12))
   Call CargaTablaAcredComisiones
End Sub

Private Sub Command8_Click()
   VENTANA.ValidarDuplicadoFila1 = "N"
   Call ABM_TABLA(XVALO(Text9))
   Call cargaTablasINGGASTOS

End Sub

Private Sub Command9_Click()
10  VENTANA.ValidarDuplicadoFila1 = "N"
  Call ABM_TABLA(XVALO(Text10))
  IDL = XVALO(Text10)
  DESC$ = TRIM$(Text2)
  SQLX$ = "SELECT CAMPO1, CAMPO2,CAMPO3,CAMPO4,CAMPO5 FROM DATASQL WITH(NOLOCK) WHERE ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'"
  SQLX$ = SQLX$ + " order by campo1"
  J& = 0
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
      J& = J& + 1
      DIAX = XVALO(SAFETEXT$(!CAMPO1))
      If DIAX > 28 Then
         Call COMUNI("Atención El Día de Vencimiento Debe ser Mayor a 0 y Menor a 29 en Fila: " & J&)
         GoTo 10
      End If
      If TRIM$(SAFETEXT$(!CAMPO2)) = "" Then
         Call COMUNI("Atención Tipo de Impuesto No Válido en Fila: " & J&)
         GoTo 10
      End If
      
      .MoveNext
    Loop
  End With
  
  Call CargaTablasVencImp
End Sub

Private Sub Form_Load()

    Call TRANSLABELS(Name)
    '
    Text6 = XVALO(CONTROL("", "CLEABANC"))
    Text7 = XVALO(CONTROL("", "EFECFACT"))
    
    Text1 = CONTROL("", "CTAREPER")
    Text3 = CONTROL("", "DENREPER")

    Text9 = CONTROL("", "CTAINGGAS")
    Text8 = CONTROL("", "DENINGGAS")
    
    Text10 = CONTROL("", "CTAFVENC")
    Text2 = CONTROL("", "DENFVENC")
    
    Text4 = CONTROL("", "CTALCUBA")
    Text5 = CONTROL("", "DENOCUBA")

    Text12 = CONTROL("", "CTAACCOM")
    Text11 = CONTROL("", "DENACCOM")
    
    Text14 = CONTROL("", "CTAASBAN")
    Text13 = CONTROL("", "DENASBAN")

    TXTPORCIIBB = TRIM$(FORMATNUM$(XVALO(CONTROL("CASHFLOW", "TASCASIB")), "F10.2"))
    Call GENERAR_ESTRUCTURAS_SQL
    Call cargaTablasRetPerc
    Call cargaTablasINGGASTOS
    Call CargaTablasVencImp
    Call CargaTablasBancarias
    Call CargaTablaSaldoBancarias
    Call CargaTablaMediosCobranza
    Call CargaTablaAcredComisiones
    Call CargaTablaAsocBancosARubrosFijos
    
End Sub
'
Sub CargaTablaAsocBancosARubrosFijos()
  If XVALO(Text14) < 1 Then Exit Sub
  If TRIM$(Text13) = "" Then Exit Sub
  '
  IDL = XVALO(Text14)
  DESC$ = TRIM$(Text13)
  Condi$ = "ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'"
  If CantRegistros&("DATASQL", Condi$, "NOMESS") < 1 Then
     Call CARGARDATOSFIJOSCASHFLOW(MSF6, IDL, DESC$)
  End If
  J& = 0
  Call CARGA_ENCABEZADO(FSCASHFLOW.MSF6, "Rubro/A20;Cuenta/A18;Referencia/A32", FSCASHFLOW)
  SQLX$ = "SELECT CAMPO1, CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK) WHERE ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'"
  SQLX$ = SQLX$ + " order by campo2"
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
      J& = J& + 1
      MSF6.Rows = J& + 1
      MSF6.TextMatrix(J&, 1) = SAFETEXT$(!CAMPO1)
      MSF6.TextMatrix(J&, 2) = SAFETEXT$(!CAMPO2)
      MSF6.TextMatrix(J&, 3) = SAFETEXT$(!CAMPO3)
      .MoveNext
    Loop
  End With
  

End Sub
Sub CargaTablaSaldoBancarias()
  If XVALO(Text4) < 1 Then Exit Sub
  If TRIM$(Text5) = "" Then Exit Sub
  '
  IDL = XVALO(Text4)
  DESC$ = TRIM$(Text5)
  J& = 0
  Call CARGA_ENCABEZADO(FSCASHFLOW.MSF4, "Cta/A14;Denominación/A32;Saldo/F14.2", FSCASHFLOW)
  SQLX$ = "SELECT CAMPO1, CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK) WHERE ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'"
  SQLX$ = SQLX$ + " order by campo2"
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
      J& = J& + 1
      MSF4.Rows = J& + 1
      MSF4.TextMatrix(J&, 1) = SAFETEXT$(!CAMPO1)
      MSF4.TextMatrix(J&, 2) = SAFETEXT$(!CAMPO2)
      MSF4.TextMatrix(J&, 3) = SAFETEXT$(!CAMPO3)
      .MoveNext
    Loop
  End With
  

End Sub
Sub CargaTablaAcredComisiones()
  If XVALO(Text12) < 1 Then Exit Sub
  If TRIM$(Text11) = "" Then Exit Sub
  '
  IDL = XVALO(Text12)
  DESC$ = TRIM$(Text11)
  J& = 0
  Call CARGA_ENCABEZADO(FSCASHFLOW.MSF5, "Cta/A14;Denominación/A32;D.H.Acred/F14.2;% Comis./F5.2", FSCASHFLOW)
  SQLX$ = "SELECT CAMPO1, CAMPO2,CAMPO3,CAMPO4 FROM DATASQL WITH(NOLOCK) WHERE ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'"
  SQLX$ = SQLX$ + " order by campo2"
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
      J& = J& + 1
      MSF5.Rows = J& + 1
      MSF5.TextMatrix(J&, 1) = SAFETEXT$(!CAMPO1)
      MSF5.TextMatrix(J&, 2) = SAFETEXT$(!CAMPO2)
      MSF5.TextMatrix(J&, 3) = SAFETEXT$(!CAMPO3)
      MSF5.TextMatrix(J&, 4) = SAFETEXT$(!CAMPO4)
      .MoveNext
    Loop
  End With
  

End Sub

Sub CargaTablasBancarias()
   Condi$ = " LISTA =   'CUENBANTA'"
   IDLI& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", Condi$, "NOMESS"))
   If IDLI& > 0 Then
        For U& = 1 To ULTREG&("CUEBANC")
           XX$ = REGLEIDO$("CUEBANC", U&)
           CXI% = CUEINT%(TRIM$(Left$(XX$, 12)))
           CODE$ = TRIM$(Left$(XX$, 12))
           DESCR$ = TRIM$(Mid$(XX$, 13))
           If CXI% > 0 Then
             JJ& = JJ& + 1
             SQLX$ = "SELECT * FROM DATASQL WHERE CAMPO1 = '" & CODE$ & "'"
             Dim RS As ADODB.Recordset
             Set RS = New ADODB.Recordset
             RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
             With RS
               If .EOF Then 'SOLO AGREGA SI NO EXISTIA
                 .AddNew
                 !CAMPO1 = CODE$
                 !CAMPO2 = DESCR$
                 !ID_Lista = IDLI&
                 !LISTA = "CUENBANTA"
                 !CODIEMPR = CodiEmp%
                 .Update
               End If
             End With
             RS.Close
             Set RS = Nothing
           End If
        Next U&
   End If
End Sub
Sub cargaTablasRetPerc()
   
  If XVALO(Text1) < 1 Then Exit Sub
  If TRIM$(Text3) = "" Then Exit Sub
  '
  IDL = XVALO(Text1)
  DESC$ = TRIM$(Text3)
  J& = 0
  Call CARGA_ENCABEZADO(FSCASHFLOW.MSF, "Imp./F6;Denominación/A32;Cuenta Cble./A16;Denominación/A32;Cuenta Cble./A16;Denominación/A32", FSCASHFLOW)
  SQLX$ = "SELECT CAMPO2, CAMPO3,CAMPO4, CAMPO5,CAMPO6, CAMPO7 FROM DATASQL WITH(NOLOCK) WHERE ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'"
  SQLX$ = SQLX$ + " order by campo2"
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
      J& = J& + 1
      MSF.Rows = J& + 1
      MSF.TextMatrix(J&, 1) = SAFETEXT$(!CAMPO2)
      MSF.TextMatrix(J&, 2) = SAFETEXT$(!CAMPO3)
      MSF.TextMatrix(J&, 3) = SAFETEXT$(!CAMPO4)
      MSF.TextMatrix(J&, 4) = SAFETEXT$(!CAMPO5)
      MSF.TextMatrix(J&, 5) = SAFETEXT$(!CAMPO6)
      MSF.TextMatrix(J&, 6) = SAFETEXT$(!CAMPO7)
      
      .MoveNext
    Loop
  End With
  
End Sub

Sub cargaTablasINGGASTOS()

  If XVALO(Text9) < 1 Then Exit Sub
  If TRIM$(Text8) = "" Then Exit Sub
  '
  IDL = XVALO(Text9)
  DESC$ = TRIM$(Text8)
  J& = 0
  Call CARGA_ENCABEZADO(FSCASHFLOW.MSF2, "Fecha/d8;Cod./F4;Referencia/A26;Ingreso/F14.2;Gasto/F14.2;Observación/A48", FSCASHFLOW)
  SQLX$ = "SELECT CAMPOFE1, CAMPO2,CAMPO3,CAMPO4,CAMPO5,CAMPO6 FROM DATASQL WITH(NOLOCK) WHERE ID_LISTA = " & IDL & " And LISTA = '" & DESC$ & "'"
  SQLX$ = SQLX$ + " ORDER BY CAMPO1"
  
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
      J& = J& + 1
      MSF2.Rows = J& + 1
      MSF2.TextMatrix(J&, 1) = FECHATEX$(CVINT(!CAMPOFE1))
      MSF2.TextMatrix(J&, 2) = SAFETEXT$(!CAMPO2)
      MSF2.TextMatrix(J&, 3) = SAFETEXT$(!CAMPO3)
      MSF2.TextMatrix(J&, 4) = SAFETEXT$(!CAMPO4)
      MSF2.TextMatrix(J&, 5) = SAFETEXT$(!CAMPO5)
      MSF2.TextMatrix(J&, 6) = SAFETEXT$(!CAMPO6)
      
      .MoveNext
    Loop
  End With
  RS.Close
  Set RS = Nothing
End Sub

Private Sub mnuAsociarReportes_Click()
    Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub mnuReportesVarios_Click()
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub mnuTabaux_Click()
  Call GENERAR_ESTRUCTURAS_SQL
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(rs1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.MSF2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  IDLIST& = XVALO(RESP_ZELE_VECT(1))
  If InStr(EPAC$, "ENVAPLAST") > 0 Then
    LIXTA$ = TRIM$(RESP_ZELE_VECT(2))
    If LIXTA$ = "MATERIALES ENVAPLAST" Then
       Call COMUNI("Atención !!! Editar Tabla de Familias Desde Maestro de Artículos")
       Exit Sub
    End If
  End If

  Call ABM_TABLA(IDLIST&)

End Sub

Private Sub TXTPORCIIBB_LostFocus()
    TXTPORCIIBB = TRIM$(FORMATNUM$(XVALO(TXTPORCIIBB), "F10.2"))
End Sub
