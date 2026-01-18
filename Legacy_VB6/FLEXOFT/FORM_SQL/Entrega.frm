VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Entrega 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instrucciones de Despacho"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9200
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   10220
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Datos de Transporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   120
         TabIndex        =   22
         Top             =   3600
         Width           =   9900
         Begin VB.CommandButton Command6 
            Caption         =   "Grabar e Imprmir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5160
            TabIndex        =   55
            Top             =   5040
            Visible         =   0   'False
            Width           =   2100
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
            Height          =   360
            Left            =   1500
            TabIndex        =   54
            Top             =   1440
            Width           =   195
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Retiran de nuestro depósito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   51
            Top             =   4440
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1720
            MaxLength       =   48
            TabIndex        =   48
            Top             =   1440
            Width           =   3345
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Grabar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   7560
            TabIndex        =   42
            Top             =   5040
            Width           =   2100
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Continuar sin Instrucciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   240
            TabIndex        =   41
            Top             =   5040
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3240
            TabIndex        =   40
            Top             =   5040
            Width           =   1500
         End
         Begin VB.TextBox HoraTransporte 
            Appearance      =   0  'Flat
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
            Left            =   5400
            MaxLength       =   48
            TabIndex        =   38
            Top             =   360
            Width           =   4275
         End
         Begin VB.TextBox TeleTransporte 
            Appearance      =   0  'Flat
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
            Left            =   5880
            MaxLength       =   21
            TabIndex        =   36
            Top             =   1440
            Width           =   3800
         End
         Begin VB.TextBox CodPostTransporte 
            Appearance      =   0  'Flat
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
            Left            =   1005
            MaxLength       =   21
            TabIndex        =   35
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox ProvTransporte 
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
            Left            =   1005
            MaxLength       =   32
            TabIndex        =   33
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox LocTransporte 
            Appearance      =   0  'Flat
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
            Left            =   4560
            MaxLength       =   48
            TabIndex        =   29
            Top             =   1095
            Width           =   5115
         End
         Begin VB.TextBox DomiTransporte 
            Appearance      =   0  'Flat
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
            Left            =   1005
            MaxLength       =   48
            TabIndex        =   28
            Top             =   720
            Width           =   8670
         End
         Begin VB.TextBox InstrucTransporte 
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
            Height          =   1215
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Top             =   3720
            Width           =   9500
         End
         Begin VB.TextBox DenoTransporte 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1770
            MaxLength       =   48
            TabIndex        =   24
            Top             =   360
            Width           =   3585
         End
         Begin VB.TextBox NumeTransporte 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1005
            TabIndex        =   23
            Top             =   375
            Width           =   645
         End
         Begin MSFlexGridLib.MSFlexGrid MSFINSTRUC 
            Height          =   1755
            Left            =   120
            TabIndex        =   26
            Top             =   1920
            Width           =   9525
            _ExtentX        =   16801
            _ExtentY        =   3096
            _Version        =   393216
            Cols            =   5
            ForeColor       =   0
            WordWrap        =   -1  'True
            AllowBigSelection=   -1  'True
            Enabled         =   -1  'True
            FocusRect       =   2
            ScrollBars      =   2
            AllowUserResizing=   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Horario"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   11
            Left            =   5040
            TabIndex        =   39
            Top             =   160
            Width           =   915
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   10
            Left            =   4920
            TabIndex        =   37
            Top             =   1455
            Width           =   915
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.P."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   0
            TabIndex        =   34
            Top             =   1095
            Width           =   915
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Provincia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   0
            TabIndex        =   32
            Top             =   1440
            Width           =   915
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Domicilio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   0
            TabIndex        =   31
            Top             =   735
            Width           =   915
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Localidad"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   3480
            TabIndex        =   30
            Top             =   1095
            Width           =   915
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Transporte"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   -120
            TabIndex        =   27
            Top             =   360
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Datos de Envío de Mercadería"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   9900
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
            Left            =   1500
            TabIndex        =   53
            Top             =   1440
            Width           =   195
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Cambiar Sucursal y traer datos del Maestro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   5295
            TabIndex        =   52
            ToolTipText     =   "Al Seleccionar una Sucursal Trae los Datos Originales Configurados en el Maestro de Clientes."
            Top             =   240
            Visible         =   0   'False
            Width           =   2820
         End
         Begin VB.TextBox HorarioSucu 
            Appearance      =   0  'Flat
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
            Left            =   7920
            MaxLength       =   32
            TabIndex        =   49
            Top             =   320
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox NroCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   8520
            TabIndex        =   45
            Top             =   720
            Width           =   1160
         End
         Begin VB.TextBox NroPedido 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   8520
            TabIndex        =   43
            Top             =   320
            Width           =   1160
         End
         Begin VB.TextBox InstrucEntrega 
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
            Height          =   1215
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   1845
            Width           =   9500
         End
         Begin VB.TextBox TelEntrega 
            Appearance      =   0  'Flat
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
            Left            =   5880
            MaxLength       =   32
            TabIndex        =   19
            Top             =   1440
            Width           =   3800
         End
         Begin VB.TextBox CodPostalEntrega 
            Appearance      =   0  'Flat
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
            Left            =   1005
            MaxLength       =   10
            TabIndex        =   18
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox ProvEntrega 
            Appearance      =   0  'Flat
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
            Left            =   1005
            MaxLength       =   32
            TabIndex        =   16
            Top             =   1440
            Width           =   495
         End
         Begin VB.CommandButton Command25 
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
            Left            =   1650
            TabIndex        =   10
            Top             =   320
            Width           =   195
         End
         Begin VB.TextBox SucuEntrega 
            Appearance      =   0  'Flat
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
            Left            =   1005
            TabIndex        =   0
            Top             =   320
            Width           =   645
         End
         Begin VB.TextBox LocEntrega 
            Appearance      =   0  'Flat
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
            Left            =   5400
            MaxLength       =   32
            TabIndex        =   9
            Top             =   1080
            Width           =   4275
         End
         Begin VB.TextBox DomiEntrega 
            Appearance      =   0  'Flat
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
            Left            =   1005
            MaxLength       =   32
            TabIndex        =   8
            Top             =   720
            Width           =   7110
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Horario"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   12
            Left            =   4800
            TabIndex        =   50
            Top             =   360
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1720
            TabIndex        =   47
            Top             =   1440
            Width           =   3315
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "N°C."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   8160
            TabIndex        =   46
            Top             =   720
            Width           =   315
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "N°P."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   8160
            TabIndex        =   44
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   4920
            TabIndex        =   20
            Top             =   1440
            Width           =   915
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.P."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   0
            TabIndex        =   17
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Provincia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   0
            TabIndex        =   15
            Top             =   1440
            Width           =   915
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   13
            Left            =   0
            TabIndex        =   14
            Top             =   300
            Width           =   915
         End
         Begin VB.Label DenoSucursal 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1920
            TabIndex        =   13
            Top             =   320
            Width           =   3315
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Domicilio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   0
            TabIndex        =   12
            Top             =   720
            Width           =   915
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Localidad"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   4320
            TabIndex        =   11
            Top             =   1080
            Width           =   915
         End
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   0
         X2              =   10200
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   0
         X2              =   10200
         Y1              =   9160
         Y2              =   9160
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   10
         X2              =   10
         Y1              =   3480
         Y2              =   9120
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   0
         X2              =   10200
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   10
         X2              =   10
         Y1              =   120
         Y2              =   3400
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   10200
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         Caption         =   "TEPRUEBA"
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
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   4
         Top             =   -120
         Visible         =   0   'False
         Width           =   9495
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   7320
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "Entrega.frx":0000
      Top             =   4800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4800
      OleObjectBlob   =   "Entrega.frx":0234
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   10350
      X2              =   10350
      Y1              =   120
      Y2              =   3400
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   10350
      X2              =   10350
      Y1              =   3480
      Y2              =   9120
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   135
      Left            =   6480
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Entrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub TraeDatosDeEntregadeDatosEntrega(NPED&, NC)
    
    CONDI$ = " NROPED = " & NPED&
    SQLX$ = "SELECT * FROM DatoEntrega"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With rs
    
        NroPedido = NPED&
        NroCliente = XVALO(NC)
        SucuEntrega = XVALO(!NUME_SUC)
        DenoSucursal = SAFETEXT$(!Deno_Suc)
        DomiEntrega = SAFETEXT$(!DOMI_SUC)
        CodPostalEntrega = SAFETEXT$(!CodPostal_Suc) '
        LocEntrega = SAFETEXT$(!LOCA_SUC)
        TelEntrega = SAFETEXT$(!Tele_Suc)
        NumeTransporte = XVALO(!CODI_TRA)
        'HorarioSucu = SAFETEXT$(!Hora_Ent_Suc)
        InstrucEntrega = SAFETEXT$(!InstruccionSuc)
        ProvEntrega = SAFETEXT$(!ProvEntre_ClI)
        DenoTransporte = SAFETEXT$(!DENO_TRA)
        'TXTCODACTIVO = SAFETEXT$(!Raso_Tra) '
        DomiTransporte = SAFETEXT$(!DOMI_TRA)
        LocTransporte = SAFETEXT$(!LOCA_TRA)
        CodPostTransporte = SAFETEXT$(!CodPost_Tra)
        ProvTransporte = SAFETEXT$(!PROV_TRA)
        TeleTransporte = SAFETEXT$(!Tele_Tra)
        HoraTransporte = SAFETEXT$(!HoraTransporte)
        InstrucTransporte = SAFETEXT$(!InstrucTransporte)
    End With
    rs.Close
    Set rs = Nothing
    Call CargarInstruccionesTransporte(XVALO(Me.NroCliente), XVALO(SucuEntrega), XVALO(NumeTransporte))
End Sub

Sub Command14_Click()
    Command14.Enabled = False
    
    'si retira en domicilio tildado
'    If Check1.Value = 1 Then
'        Label1.Caption = "OK"
'        Hide
'        GoTo 99
'    End If
    
    Call ValidarDatosEntrega
    If Label1.Caption = "OK" Then
      Hide
    End If
99  Command14.Enabled = True
End Sub
Sub ValidarDatosEntrega()
    If TRIM$(SucuEntrega) = "" Then
       Call COMUNI("Atención Debe Seleccionar una Sucursal")
       Call PINTATEXT(SucuEntrega)
       GoTo 99
    End If
'    If TRIM$(HorarioSucu) = "" Then
'       Call COMUNI("Atención Debe Seleccionar Horario de Sucursal")
'       Call PINTATEXT(HorarioSucu)
'       GoTo 99
'    End If

    If TRIM$(DomiEntrega) = "" Then
       Call COMUNI("Atención Debe Seleccionar el Domicilio de la Sucursal")
       Call PINTATEXT(DomiEntrega)
       GoTo 99
    End If

    If TRIM$(LocEntrega) = "" Then
       Call COMUNI("Atención Debe Seleccionar la Localidad de la Sucursal")
       Call PINTATEXT(LocEntrega)
       GoTo 99
    End If

    If TRIM$(CodPostalEntrega) = "" Then
       Call COMUNI("Atención Debe Ingresar el Código Postal de la Sucursal")
       Call PINTATEXT(CodPostalEntrega)
       GoTo 99
    End If
    If TRIM$(ProvEntrega) = "" Then
       Call COMUNI("Atención Debe Seleccionar Provincia de Sucursal")
       Call PINTATEXT(ProvEntrega)
       GoTo 99
    End If

    If TRIM$(TelEntrega) = "" Then
       Call COMUNI("Atención Debe Seleccionar Ingresar el Teléfono de la Sucursal")
       Call PINTATEXT(TelEntrega)
       GoTo 99
    End If
'    If TRIM$(InstrucEntrega) = "" Then
'       Call COMUNI("Atención Debe Ingresar la Instrucción para la Sucursal")
'       Call PINTATEXT(InstrucEntrega)
'       GoTo 99
'    End If
    If TRIM$(NumeTransporte) = "" Then
       Call COMUNI("Atención! Falta Código de Transporte")
       Call PINTATEXT(NumeTransporte)
       GoTo 99
    End If

    If TRIM$(HoraTransporte) = "" Then
       Call COMUNI("Atención Falta Horario Transporte")
       Call PINTATEXT(HoraTransporte)
       GoTo 99
    End If

    If TRIM$(DomiTransporte) = "" Then
       Call COMUNI("Atención! Falta Domicilio de Transporte")
       Call PINTATEXT(DomiTransporte)
       GoTo 99
    End If

    If TRIM$(CodPostTransporte) = "" Then
       Call COMUNI("Atención Código de Postal de Transporte")
       Call PINTATEXT(CodPostTransporte)
       GoTo 99
    End If
    
    If TRIM$(LocTransporte) = "" Then
       Call COMUNI("Atención! Falta Localidad de Transporte")
       Call PINTATEXT(LocTransporte)
       GoTo 99
    End If

    If TRIM$(TeleTransporte) = "" Then
       Call COMUNI("Atención Falta Teléfono de Transporte")
       Call PINTATEXT(TeleTransporte)
       GoTo 99
    End If
    
     If TRIM$(ProvTransporte) = "" Then
       Call COMUNI("Atención! Falta Provincia de Transporte")
       Call PINTATEXT(ProvTransporte)
       GoTo 99
    End If

'    If TRIM$(InstrucTransporte) = "" Then
'       Call COMUNI("Atención Falta Instrucción de Transporte")
'       Call PINTATEXT(InstrucTransporte)
'       GoTo 99
'    End If
   
    Label1.Caption = "OK"
99
End Sub
Private Sub Command1_Click()
   Call Command25_Click
End Sub

Private Sub Command22_Click()

End Sub

Private Sub Command2_Click()
   OPX% = COMALTER%("Opciones de Información\\Cancelar\Continuar con Datos Incompletos")
   If OPX% < 2 Then Exit Sub
   Label1.Caption = "OK"
   Hide
End Sub

Private Sub Command24_Click()

End Sub

Private Sub Command25_Click()
    CONDI$ = "NUME_CLI = " & XVALO(NroCliente)
    CONDI$ = CONDI$ + " AND STATUS >=0"
    
    Call CARGALISEL("SUCENTRE", "SUCENTRE", "NUME_SUC/f10;DENO_SUC/A24", CONDI$, "NOMESS")
    Call SELECHO("SUCENTRE")
    Call T_Espera(1 / 2)
    VDEVU$ = TRIM$(VALACT1$("SUCENTRE"))
    If VDEVU$ <> "" Then
       If XVALO(VDEVU$) = XVALO(SucuEntrega) And XVALO(Me.NroPedido) > 0 And Me.command1.Visible = True Then SucuEntrega = ""
       SucuEntrega = VDEVU$
    End If
End Sub

Private Sub Command3_Click()
       Label1.Caption = "Faltan Datos"
       Unload Me
End Sub

Private Sub Command4_Click()


Call SELECHO("PROVINC")
        VDEVU$ = VALACT1$("PROVINC")
        If TRIM$(VDEVU$) <> "" Then
           Me.ProvEntrega.TEXT = VDEVU$
        End If
End Sub

Private Sub Command5_Click()

Call SELECHO("PROVINC")
        VDEVU$ = VALACT1$("PROVINC")
        If TRIM$(VDEVU$) <> "" Then
           Me.ProvTransporte.TEXT = VDEVU$
        End If
End Sub



Private Sub Command6_Click()
'    Command6.Enabled = False
'    Call Command14_Click
'    NROPED& = XVALO(NroPedido)
'    NC = XVALO(NroCliente)
'    If NC < 1 Then GoTo 99
'        Call ValidarDatosEntrega
'    If Label1.Caption = "OK" Then
'      Call grabarDatosEntregaAhp(NROPED&, NC)
'    End If
'    PrintEtiDes.NroPedido = NROPED&
'
'    PrintEtiDes.Show 1
'99  Command6.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Campos$ = "id/A0;Instrucción/A80;NC/A0;NSUCURSAL/A0;CODITRANSP/A0"
    Call CARGA_ENCABEZADO(Me.MSFINSTRUC, Campos$, Me)
'    FORPED07.Text19 = Text19
'    FORPED07.Text20 = Text20
'    FORPED07.LENTREGA(0) = LENTREGA(0)
'    FORPED07.LENTREGA(1) = LENTREGA(1)
    Me.MSFINSTRUC.Rows = 1
    Me.InstrucEntrega = ""
    Me.InstrucTransporte = ""
    Call CREATABLA_BADACLIEN
End Sub

Private Sub Label2_DblClick()
    IMPREDESPA.Show 1
End Sub
Private Sub Label3_DblClick()
    IMPREDESPA.Show 1
End Sub

Private Sub MSFINSTRUC_Click()
    If MSFINSTRUC.Row >= 1 Then
     With Entrega
          .InstrucTransporte.TEXT = MSFINSTRUC.TextMatrix(MSFINSTRUC.Row, 2)
     End With
    End If

End Sub



Sub CargarInstrucciones()

    Me.MSFINSTRUC.Rows = 1
    
    NC = XVALO(FORPED07.Text1)
    nTransporte% = XVALO(Text19) 'NRO TRANSPORTE
    NSUCURSAL% = XVALO(Text20) 'NRO SUCURSAL
    
    If NC < 1 Then Exit Sub
    
    If NSUCURSAL% < 1 Then Exit Sub
    
    If nTransporte% < 1 Then Exit Sub

    CONDI$ = "Nume_Cli = " & NC
    CONDI$ = CONDI$ + " AND Nume_Suc = " & NSUCURSAL%
    CONDI$ = CONDI$ + " AND Codi_Tra = " & nTransporte%
    CONDI$ = CONDI$ + " AND MARBORRA < 1"
    
    Dim MSF As msFlexGrid
    Set MSF = MSFINSTRUC
    MSF.Rows = 1
    If MSF.Rows < 2 Then MSF.Rows = 2
    '
    SQLX$ = "SELECT * FROM INSTRUC_ENTRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    SQLX$ = SQLX$ + " ORDER BY INSTRUC_ENTRE_ID  "
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        REG& = REG& + 1
        MSF.Rows = REG& + 1
        MSF.TextMatrix(REG&, 1) = XVALO(!INSTRUC_ENTRE_ID)
            INSTRU$ = SAFETEXT$(!INSTRUCCION)
            Label22 = INSTRU$
            MSF.RowHeight(REG&) = Label22.Height + 100
        MSF.TextMatrix(REG&, 2) = SAFETEXT$(!INSTRUCCION)
        MSF.TextMatrix(REG&, 3) = XVALO(!nume_cli)
        MSF.TextMatrix(REG&, 4) = XVALO(!NUME_SUC)
        MSF.TextMatrix(REG&, 5) = XVALO(!CODI_TRA)
       .MoveNext
      Loop
    End With

    rs.Close
    Set rs = Nothing
    Label22 = ""
End Sub

Sub grabarDatosEntregaAhp(NROEPED&, NC)
    Nuevo% = 0
    CONDI$ = "NROPED = " & NROEPED&
    If CantRegistros("DatoEntrega", CONDI$, "NOMESS") < 1 Then
        Nuevo% = 1
    End If
    SQLX$ = "SELECT * FROM DatoEntrega"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With rs
        If Nuevo% = 1 Then
         .AddNew
        End If
        !NROPED = XVALO(NROEPED&)
        !NROCLIE = XVALO(NC)
        !MarBorra = 0
        !NUME_SUC = XVALO(SucuEntrega)
        !Deno_Suc = Left$(DenoSucursal, 128)
        !DOMI_SUC = Left$(DomiEntrega, 128)
        !CodPostal_Suc = Left$(CodPostalEntrega, 16)
        !LOCA_SUC = Left$(LocEntrega, 128)
        !Tele_Suc = Left$(TelEntrega, 64)
        !CODI_TRA = XVALO(NumeTransporte)
        '!Hora_Ent_Suc = Left$(HorarioSucu, 64)
        !InstruccionSuc = Left$(InstrucEntrega, 512)
        !ProvEntre_ClI = Left$(ProvEntrega, 4)
        !DENO_TRA = Left$(DenoTransporte, 128)
         !DOMI_TRA = Left$(DomiTransporte, 128)
        !LOCA_TRA = Left$(LocTransporte, 128)
        !CodPost_Tra = Left$(CodPostTransporte, 16)
        !PROV_TRA = Left$(ProvTransporte, 4)
        !Tele_Tra = Left$(TeleTransporte, 64)
        !HoraTransporte = Left$(HoraTransporte, 64)
        !InstrucTransporte = Left$(InstrucTransporte, 512)
        !EntEnDomi = 0
        !CODIEMPR = CodiEmp%
        .Update
    End With
    
   End Sub

Private Sub NroPedido_Change()
   NPED& = NroPedido
   NC = XVALO(NroCliente)
   Call TraeDatosDeEntregadeDatosEntrega(NPED&, NC) 'carga datos de la tabla DatoEntrega
End Sub

Private Sub ProvEntrega_Change()
  ProvEntrega = UCase$(ProvEntrega)
   Label2 = TRIM$(ECOARCH$("PROVINC", TRIM$(ProvEntrega)))
End Sub

Private Sub ProvTransporte_Change()
   ProvTransporte = UCase$(ProvTransporte)
   Text1 = TRIM$(ECOARCH$("PROVINC", TRIM$(ProvTransporte)))
End Sub


Private Sub SucuEntrega_Change()
  Call BLANFORMU
  NSUC% = XVALO(SucuEntrega)
  NC = XVALO(NroCliente)
  NPED& = XVALO(NroPedido)
  CONDI$ = "NUME_SUC = " & NSUC%
  CONDI$ = CONDI$ + " AND NUME_CLI = " & NC
  CONDI$ = CONDI$ + " AND STATUS >=0"
  If CantRegistros&("SUCENTRE", CONDI$, "NOMESS") < 1 Then Exit Sub 'SI NO EXISTE O ESTA DE BAJA NO HACE NADA
  
  DenoSucursal = VALOBADA("SUCENTRE", "DENO_SUC", CONDI$, "NOMESS")
  If NPED& > 0 And Me.command1.Visible = False Then
    Exit Sub
  Else
    Call TraerDatosDeEntregaFichaCliente(NC, NSUC%) 'carga datos de la tabla sucentre y transportes
  End If
End Sub
Sub TraerDatosDeEntregaFichaCliente(NC, NSUC%)
     'ojo que si solo tiene sucursal sin transporte no muestra nada.

    SQLX$ = "SELECT S.Nume_Cli,S.CodiEmpr,S.Nume_Suc,S.Deno_Suc,S.Domi_Suc,S.Loca_Suc,S.Tele_Suc"
    SQLX$ = SQLX$ + " ,S.Codi_Tra,S.CodPostal_Suc,S.Prov_Cli,S.Hora_Ent,S.InstruccionSuc"
    SQLX$ = SQLX$ + " ,T.Deno_Tra ,T.Raso_Tra,T.Domi_Tra,T.Loca_Tra,T.Cuit_Tra,T.Piva_Tra"
    SQLX$ = SQLX$ + " ,T.Tele_Tra , T.Faxi_Tra, T.Mail_Tra, T.Ctac_Tra, Prov_Tra,T.CodPostal_Tra"
    SQLX$ = SQLX$ + "   From SUCENTRE S INNER JOIN TRANSPORTES T  WITH(NOLOCK)"
    SQLX$ = SQLX$ + "   on S.codi_tra = T.codi_tra "
    SQLX$ = SQLX$ + "   Where S.NUME_CLI = " & NC
    SQLX$ = SQLX$ + " And S.NUME_SUC = " & NSUC%
    Set rs = READSET(SQLX$)
    With rs
        'SucuEntrega = XVALO(!Nume_Suc)
        'DenoSucursal = SAFETEXT$(!Deno_Suc)
        DomiEntrega = SAFETEXT$(!DOMI_SUC)
        CodPostalEntrega = SAFETEXT$(!CodPostal_Suc)
        LocEntrega = SAFETEXT$(!LOCA_SUC)
        TelEntrega = SAFETEXT$(!Tele_Suc)
        NumeTransporte = SAFETEXT$(!CODI_TRA)
        'HorarioSucu = SAFETEXT$(!Hora_Ent)
        InstrucEntrega = SAFETEXT$(!InstruccionSuc)
        ProvEntrega = SAFETEXT$(!Prov_Cli)
        DenoTransporte = SAFETEXT$(!DENO_TRA)
        TXTCODACTIVO = SAFETEXT$(!Raso_Tra)
        DomiTransporte = SAFETEXT$(!DOMI_TRA)
        LocTransporte = SAFETEXT$(!LOCA_TRA)
        CodPostTransporte = SAFETEXT$(!CodPostal_Tra)
        ProvTransporte = SAFETEXT$(!PROV_TRA)
        TeleTransporte = SAFETEXT$(!Tele_Tra)
        HoraTransporte = SAFETEXT$(!Faxi_Tra)
        InstrucTransporte = ""
    End With
    rs.Close
    Set rs = Nothing
    Call CargarInstruccionesTransporte(XVALO(Me.NroCliente), XVALO(SucuEntrega), XVALO(NumeTransporte))
End Sub

Sub CargarInstruccionesTransporte(NC, nSucEntrega, numTransporte)

    Me.MSFINSTRUC.Rows = 1
    
    nTransporte% = XVALO(numTransporte) 'NRO TRANSPORTE
    NSUCURSAL% = XVALO(nSucEntrega) 'NRO SUCURSAL
    NC = XVALO(NC)
    If NC < 1 Then Exit Sub
    
    If NSUCURSAL% < 1 Then Exit Sub
    
    If nTransporte% < 1 Then Exit Sub

    CONDI$ = "Nume_Cli = " & NC
    CONDI$ = CONDI$ + " AND Nume_Suc = " & NSUCURSAL%
    CONDI$ = CONDI$ + " AND Codi_Tra = " & nTransporte%
    CONDI$ = CONDI$ + " AND MARBORRA < 1"
    
    Dim MSF As msFlexGrid
    Set MSF = MSFINSTRUC
    MSF.Rows = 1
    If MSF.Rows < 2 Then MSF.Rows = 2
    '
    SQLX$ = "SELECT * FROM INSTRUC_ENTRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    SQLX$ = SQLX$ + " ORDER BY INSTRUC_ENTRE_ID  "
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        REG& = REG& + 1
        MSF.Rows = REG& + 1
        MSF.TextMatrix(REG&, 1) = XVALO(!INSTRUC_ENTRE_ID)
            INSTRU$ = SAFETEXT$(!INSTRUCCION)
            Label22 = INSTRU$
            MSF.RowHeight(REG&) = Label22.Height + 100
            MSF.TextMatrix(REG&, 2) = SAFETEXT$(!INSTRUCCION)
            MSF.TextMatrix(REG&, 3) = XVALO(!nume_cli)
            MSF.TextMatrix(REG&, 4) = XVALO(!NUME_SUC)
            MSF.TextMatrix(REG&, 5) = XVALO(!CODI_TRA)
        
       .MoveNext
      Loop
    End With

    rs.Close
    Set rs = Nothing

End Sub
Sub BLANFORMU()
        DenoSucursal = ""
        DomiEntrega = ""
        CodPostalEntrega = ""
        LocEntrega = ""
        TelEntrega = ""
        NumeTransporte = ""
        'HorarioSucu = ""
        InstrucEntrega = ""
        ProvEntrega = ""
        DenoTransporte = ""
        TXTCODACTIVO = ""
        DomiTransporte = ""
        LocTransporte = ""
        CodPostTransporte = ""
        ProvTransporte = ""
        TeleTransporte = ""
        HoraTransporte = ""
        InstrucTransporte = ""
        MSFINSTRUC.Rows = 1
End Sub
