VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form EDITAGENDA 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar o Quitar de la Lista"
   ClientHeight    =   6615
   ClientLeft      =   1905
   ClientTop       =   3300
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   26638.78
   ScaleMode       =   0  'User
   ScaleWidth      =   5864.796
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox TXTC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   330
      Index           =   14
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Text            =   "Nombre Realizó:"
      Top             =   4920
      Width           =   1500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "."
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
      Left            =   1800
      TabIndex        =   39
      ToolTipText     =   "Click Para Seleccionar Sección"
      Top             =   4200
      Width           =   175
   End
   Begin VB.CommandButton cmdelinminar 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2135
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1300
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1300
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5805
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5385
      Begin VB.TextBox TXTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   15
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "ID.:"
         Top             =   240
         Width           =   540
      End
      Begin VB.TextBox TXTID 
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
         Left            =   4600
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "   "
         Top             =   240
         Width           =   525
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
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
         Left            =   1800
         TabIndex        =   38
         ToolTipText     =   "Click Para Seleccionar Sección"
         Top             =   3480
         Width           =   175
      End
      Begin VB.TextBox txtcodigo2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2040
         TabIndex        =   37
         Top             =   960
         Width           =   1290
      End
      Begin VB.TextBox txtcodusuario 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2040
         TabIndex        =   36
         Top             =   4560
         Width           =   1290
      End
      Begin VB.TextBox TXTCUENTA 
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
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "   "
         Top             =   240
         Width           =   1605
      End
      Begin VB.TextBox TXTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   13
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Cuenta:"
         Top             =   240
         Width           =   1500
      End
      Begin VB.TextBox TXTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   12
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "R.Social:"
         Top             =   600
         Width           =   1500
      End
      Begin VB.TextBox TXTRSOCIAL 
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
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   600
         Width           =   3400
      End
      Begin VB.TextBox TXTFACTURADO 
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
         Height          =   330
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   5280
         Width           =   1605
      End
      Begin VB.TextBox TXTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   11
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "Facturar ? "
         Top             =   5280
         Width           =   1500
      End
      Begin VB.CommandButton Command3 
         Caption         =   "."
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
         Left            =   1800
         TabIndex        =   29
         ToolTipText     =   "Click Para Seleccionar Sección"
         Top             =   4560
         Width           =   175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "."
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
         Left            =   1800
         TabIndex        =   6
         ToolTipText     =   "Click Para Seleccionar Sección"
         Top             =   960
         Width           =   175
      End
      Begin VB.TextBox TXTHORORGANIZADO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1725
         TabIndex        =   25
         Text            =   "   "
         Top             =   3840
         Width           =   1605
      End
      Begin VB.TextBox TXTDESCRIPCION 
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
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1320
         Width           =   3400
      End
      Begin VB.TextBox TXTSEMANA 
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
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "   "
         Top             =   2400
         Width           =   1605
      End
      Begin VB.TextBox TXTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   10
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Cód.Realiza:"
         Top             =   4560
         Width           =   1500
      End
      Begin VB.TextBox TXTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   9
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Realizado:"
         Top             =   4200
         Width           =   1500
      End
      Begin VB.TextBox TXTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   8
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Hor.Original:"
         Top             =   2040
         Width           =   1500
      End
      Begin VB.TextBox TXTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   7
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Hor.Organizado"
         Top             =   3840
         Width           =   1500
      End
      Begin VB.TextBox TXTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   6
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "F.Organizada:"
         Top             =   3480
         Width           =   1500
      End
      Begin VB.TextBox TXTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   5
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "Cuotas:"
         Top             =   3120
         Width           =   1500
      End
      Begin VB.TextBox TXTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   4
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Continuidad:"
         Top             =   2760
         Width           =   1500
      End
      Begin VB.TextBox TXTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   3
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "N° de Sem.Orig."
         Top             =   2400
         Width           =   1500
      End
      Begin VB.TextBox TXTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   2
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "Fecha Original:"
         Top             =   1680
         Width           =   1500
      End
      Begin VB.TextBox TXTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Descripción:"
         Top             =   1320
         Width           =   1500
      End
      Begin VB.TextBox TXTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Cód.Trabajo:"
         Top             =   960
         Width           =   1500
      End
      Begin VB.TextBox TXTREALIZO 
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
         Height          =   330
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   4920
         Width           =   3400
      End
      Begin VB.TextBox TXTREALIZADO 
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
         Height          =   330
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   4200
         Width           =   1290
      End
      Begin VB.TextBox TXTFECHAORIGINAL 
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
         Left            =   1725
         TabIndex        =   9
         Text            =   "   "
         Top             =   1680
         Width           =   1605
      End
      Begin VB.TextBox TXTCONTINUIDAD 
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
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "Doble Click Para Seleccionar Sección"
         Top             =   2760
         Width           =   1605
      End
      Begin VB.TextBox TXTCUOTAS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1725
         TabIndex        =   7
         Text            =   "   "
         Top             =   3120
         Width           =   1605
      End
      Begin VB.TextBox TXTHorOrig 
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
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "   "
         Top             =   2040
         Width           =   1605
      End
      Begin VB.TextBox TXTFORGANIZADA 
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
         Height          =   330
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   3480
         Width           =   1290
      End
   End
   Begin VB.PictureBox Picture999 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   8400
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1305
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "EDITAGENDA.frx":0000
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "EDITAGENDA.frx":007A
      Top             =   1800
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "EDITAGENDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ENTER EN CUALQUIER LUGAR DEL FORMULARIO Y REALIZA CLICK EN EL BOTON
'SE DEBE MARCAR EN TRUE LA PROPIEDAD DEFAULT DEL BOTON


Private Sub cmdcancelar_Click()
   
   Unload Me

End Sub




Sub ACTUALIZA_FLEX(MSF As MSFlexGrid, RENGLON%, FILA%, VALOR1$)
    '
    
End Sub

Private Sub cmdelinminar_Click()
    cmdelinminar.Enabled = False
    OPX% = COMALTER%("Realmente Desea Eliminar Este Trabajo de la Agenda\\Cancelar\Sí, Eliminar")
    'BORRO
    If OPX% = 2 Then
        SQLX$ = "DELETE * FROM AGENDA  WHERE ID = " & XVALO(TXTID)
        Badaclie.Execute (SQLX$)
    End If
    Call QuitarFiladeMSF2(XVALO(TXTID))
    'REFRESCO LA GRILLA
    'Agenda.CARGAR_GRILLA_TRABAJO
'    Agenda.FILTRAR_GRILLA_TRABAJO
    '
    Unload Me
    cmdelinminar.Enabled = True
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   SQLX$ = "SELECT * FROM  AGENDA  "
   SQLX$ = SQLX$ + " WHERE ID = " & XVALO(TXTID)
   Set RS = Badaclie.OpenRecordset(SQLX$, dbOpenDynaset)
   With RS
     .Edit
        !NUME_CLI = XVALO(Me.TXTCUENTA)
        !CODTRABAJO = XVALO(Me.txtcodigo2)
        !Fecha_Original = CVDAT(CVINT(Me.TXTFECHAORIGINAL))
        !OrdenSemana = XVALO(Me.TXTSEMANA)
        !Horario_Orig = Me.TXTHorOrig
        !CONTINUIDAD = XVALO(Me.TXTCONTINUIDAD)
        !CUOTAS = Me.TXTCUOTAS
        !MARBORRA = 0
        !Dia_Organizado = CVDAT(CVINT(Me.TXTFORGANIZADA))
        !Horario_Organizado = Me.TXTHORORGANIZADO
        !Dia_Realizado = CVDAT(CVINT(TXTREALIZADO))
        !OPER_REALIZO = XVALO(Me.txtcodusuario)
        
        FACTU$ = TRIM$(TXTFACTURADO)
        If FACTU$ = "" Then
          FACTUNU% = 0
        Else
          FACTUNU% = 1
        End If
        !FACTURADO = FACTUNU%
        !REALIZADO = CVINT(TXTREALIZADO)
        
        
'           Call CREACAMPO("BADACLIEN", "Agenda", "FACTURADO", 10, 8)
'   Call CREACAMPO("BADACLIEN", "Agenda", "MESANO_ORIG", 10, 32)
'   Call CREACAMPO("BADACLIEN", "Agenda", "ID_ORIGINAL", 4, 0)
'   Call CREACAMPO("BADACLIEN", "Agenda", "REALIZADO", 3, 0)
'   Call CREACAMPO("BADACLIEN", "Agenda", "OPER_REALIZO", 3, 0)

     .Update
   End With
   RS.Close
   Set RS = Nothing
   Call ACTUALIZAR_FILA(XVALO(TXTID))
   'Agenda.CARGAR_GRILLA_TRABAJO
   Unload Me
   Command1.Enabled = True
End Sub
Sub ACTUALIZAR_FILA(ID)
   With Agenda
        For REG& = 1 To .MSF2.Rows - 1
          If XVALO(.MSF2.TextMatrix(REG&, 16)) = XVALO(ID) Then
            .MSF2.TextMatrix(REG&, 1) = TXTCUENTA
            .MSF2.TextMatrix(REG&, 2) = TXTRSOCIAL
            .MSF2.TextMatrix(REG&, 3) = txtcodigo2
            .MSF2.TextMatrix(REG&, 4) = TXTDESCRIPCION
            .MSF2.TextMatrix(REG&, 5) = TXTFECHAORIGINAL
            .MSF2.TextMatrix(REG&, 7) = TXTHorOrig
            .MSF2.TextMatrix(REG&, 6) = TXTSEMANA
            .MSF2.TextMatrix(REG&, 8) = TXTCONTINUIDAD
            .MSF2.TextMatrix(REG&, 9) = TXTCUOTAS
            .MSF2.TextMatrix(REG&, 10) = TXTFORGANIZADA
            .MSF2.TextMatrix(REG&, 11) = TXTHORORGANIZADO
            .MSF2.TextMatrix(REG&, 12) = TXTREALIZADO
            .MSF2.TextMatrix(REG&, 13) = TXTREALIZO
            .MSF2.TextMatrix(REG&, 14) = TXTFACTURADO
            .MSF2.TextMatrix(REG&, 17) = txtcodusuario
            Exit For
          End If
        Next
   End With
End Sub
Sub QuitarFiladeMSF2(ID)
        For REG& = 1 To Agenda.MSF2.Rows - 1
          If XVALO(Agenda.MSF2.TextMatrix(REG&, 16)) = XVALO(ID) Then
             Call cmdEliminaItem(Agenda.MSF2, REG&)
             Exit For
          End If
        Next
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
'   Call SELECHO("TIPOTRAB")
'   CODTRA% = XVALO(VALACT1$("TIPOTRAB"))
'   If CODTRA% < 1 Then GoTo 99
   Call SELECHO("MASTER")
   CODTRA% = CODINT%(VALACT1$("MASTER"))
   If CODTRA% < 1 Then GoTo 99
   '
   txtcodigo2 = CODTRA%
   '
99 Command2.Enabled = True

End Sub

Function borrafila_CON_ALTURA(MSF As MSFlexGrid, REG&)

        
End Function

Private Sub Command3_Click()
   Command3.Enabled = False
   Call SELECHO("PERSONAL")
   CODPER% = XVALO(VALACT1$("PERSONAL"))
   If CODPER% < 1 Then GoTo 99
   '
   Me.txtcodusuario = CODPER%
   
99 Command3.Enabled = True

End Sub

Private Sub Command4_Click()
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
    Me.TXTFORGANIZADA = VDEV$
  End If

End Sub

Private Sub Command5_Click()
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
    Me.TXTREALIZADO = VDEV$
  End If

End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_DblClick()

End Sub

Private Sub txtcodigo2_Change()
   TXTDESCRIPCION = ""
   TXTDESCRIPCION = DESCRIT0$(CODINT%(XVALO(txtcodigo2)))
   CONTINU% = Agenda.CONTINUIDAD%(txtcodigo2)
   'ESCRIBE MES SI ES 1 , MESES SI SON MAS DE 1 O NADA SI ES < 1
   If CONTINU% = 1 Then
     MESX$ = "Mes"
   ElseIf CONTINU% > 1 Then
     MESX$ = "Meses"
   Else
     MESX$ = ""
   End If
   
   If MESX$ <> "" Then
     Me.TXTCONTINUIDAD = CONTINU% & " " & MESX$
   Else
     Me.TXTCONTINUIDAD = ""
   End If

End Sub

Private Sub txtcodusuario_Change()
   TXTREALIZO = ""
   TXTREALIZO = ECOARCH$("PERSONAL", MKS$(XVALO(txtcodusuario)))

End Sub

Private Sub TXTNOPER_Change()

End Sub

Private Sub TXTFACTURADO_DblClick()

  If TRIM$(TXTFACTURADO) = "" Then
     TXTFACTURADO = "SÍ"
  Else
     TXTFACTURADO = ""
  End If

End Sub

Private Sub TXTHORORGANIZADO_Change()
    X1 = Agenda.ERRTIEMPO%(TXTHORORGANIZADO)
    
End Sub

Private Sub TXTHORORGANIZADO_GotFocus()
  TXTHORORGANIZADO.SelStart = 0
  TXTHORORGANIZADO.SelLength = Len(TXTHORORGANIZADO)
End Sub

Private Sub TXTHORORGANIZADO_LostFocus()
    TXTHORORGANIZADO = HORATEX$(HORANUM(TXTHORORGANIZADO))
End Sub

Private Sub TXTHorOrig_Change()
'    X1 = ERRTIEMPO%(TXTHorOrig)
End Sub

Private Sub TXTHorOrig_GotFocus()
  TXTHorOrig.SelStart = 0
  TXTHorOrig.SelLength = Len(TXTHorOrig)
End Sub

Private Sub TXTHorOrig_LostFocus()
   TXTHorOrig = HORATEX$(HORANUM(TXTHorOrig))
End Sub
