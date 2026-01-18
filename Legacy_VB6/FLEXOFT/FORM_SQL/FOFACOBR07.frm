VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FOFACOBR07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación de Proveedores por Boleta de Recepción"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   6240
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C7D9CC&
      BorderStyle     =   0  'None
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
      TabIndex        =   4
      Top             =   0
      Width           =   12000
      Begin VB.Frame Frame3 
         BackColor       =   &H00C7D9CC&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1360
         Left            =   6600
         TabIndex        =   36
         Top             =   4200
         Width           =   4575
         Begin VB.TextBox Text12 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            TabIndex        =   37
            Text            =   " "
            ToolTipText     =   "Ingrese la Bonificación"
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mon.:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   480
            TabIndex        =   46
            Top             =   330
            Width           =   480
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   1050
            TabIndex        =   45
            Top             =   240
            Width           =   405
         End
         Begin VB.Label LBLIMPO 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   2880
            TabIndex        =   43
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Importe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1820
            TabIndex        =   42
            Top             =   330
            Width           =   960
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   2880
            TabIndex        =   41
            Top             =   600
            Width           =   1485
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bonificación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   45
            TabIndex        =   40
            Top             =   720
            Width           =   1440
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   2880
            TabIndex        =   39
            Top             =   960
            Width           =   1485
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Importe Neto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1710
            TabIndex        =   38
            Top             =   1070
            Width           =   1080
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C7D9CC&
         Caption         =   "IMPORTE SEGÚN ORDEN DE COMPRA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1360
         Left            =   1200
         TabIndex        =   28
         Top             =   4200
         Width           =   5295
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
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
            Left            =   2280
            TabIndex        =   29
            Text            =   " "
            ToolTipText     =   "Ingrese la Bonificación"
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   300
            Left            =   3600
            TabIndex        =   35
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Importe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2550
            TabIndex        =   34
            Top             =   330
            Width           =   960
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   300
            Left            =   3600
            TabIndex        =   33
            Top             =   600
            Width           =   1485
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bonificación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   765
            TabIndex        =   32
            Top             =   720
            Width           =   1440
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   300
            Left            =   3600
            TabIndex        =   31
            Top             =   975
            Width           =   1485
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Importe Neto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2550
            TabIndex        =   30
            Top             =   1070
            Width           =   960
         End
      End
      Begin VB.TextBox TXTVALCOMPRA 
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
         Height          =   300
         Left            =   120
         TabIndex        =   27
         Text            =   " "
         ToolTipText     =   "Fecha Correspondiente al Mes de Contabilización del I.V.A."
         Top             =   0
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton Command12 
         Caption         =   "O/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   120
         Picture         =   "FOFACOBR07.frx":0000
         TabIndex        =   23
         ToolTipText     =   "Visualizar Ordenes de Compra"
         Top             =   4320
         Width           =   855
      End
      Begin VB.CommandButton BOTNEXT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   11190
         Picture         =   "FOFACOBR07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Grabar - Registrar Documento en la Base de Datos"
         Top             =   4300
         Width           =   615
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   120
         TabIndex        =   21
         Top             =   1740
         Width           =   11685
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   0
            ScaleHeight     =   270
            ScaleWidth      =   11670
            TabIndex        =   26
            Top             =   0
            Width           =   11700
            Begin VB.Label Label18 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Código             Descripción    B/R     O/C     Pendiente     Facturado   P. Unitario       Importe  I/C   "
               BeginProperty Font 
                  Name            =   "Roboto Mono"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   0
               TabIndex        =   44
               Top             =   0
               Width           =   11670
            End
         End
         Begin VB.Line Line2 
            X1              =   6320
            X2              =   6320
            Y1              =   0
            Y2              =   1155
         End
         Begin VB.Line Line6 
            Index           =   0
            X1              =   1980
            X2              =   1980
            Y1              =   0
            Y2              =   1155
         End
         Begin VB.Line Line20 
            X1              =   4935
            X2              =   4935
            Y1              =   0
            Y2              =   1155
         End
         Begin VB.Line Line23 
            X1              =   7800
            X2              =   7800
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line6 
            Index           =   1
            X1              =   9300
            X2              =   9300
            Y1              =   0
            Y2              =   1155
         End
         Begin VB.Line Line6 
            Index           =   2
            X1              =   10640
            X2              =   10640
            Y1              =   0
            Y2              =   1155
         End
         Begin VB.Line Line6 
            Index           =   3
            X1              =   3330
            X2              =   3330
            Y1              =   0
            Y2              =   1155
         End
         Begin VB.Line Line6 
            Index           =   4
            X1              =   4080
            X2              =   4080
            Y1              =   0
            Y2              =   1155
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         ItemData        =   "FOFACOBR07.frx":0614
         Left            =   120
         List            =   "FOFACOBR07.frx":0616
         TabIndex        =   20
         Top             =   2040
         Width           =   11700
      End
      Begin VB.TextBox Text13 
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
         Height          =   300
         Left            =   1365
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   3480
      End
      Begin VB.TextBox Text6 
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
         Height          =   300
         Left            =   9825
         TabIndex        =   1
         Text            =   " "
         ToolTipText     =   "fecha REAL del Comprobante Recibido"
         Top             =   650
         Width           =   960
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1190
         Left            =   11160
         Picture         =   "FOFACOBR07.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   210
         Width           =   650
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7995
         TabIndex        =   3
         ToolTipText     =   "Número de Comprobante"
         Top             =   1050
         Width           =   2805
      End
      Begin VB.TextBox Text9 
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
         Height          =   300
         Left            =   9825
         TabIndex        =   2
         Text            =   " "
         ToolTipText     =   "Fecha Correspondiente al Mes de Contabilización del I.V.A."
         Top             =   210
         Width           =   960
      End
      Begin VB.TextBox Text5 
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
         Height          =   300
         Left            =   2205
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   2640
      End
      Begin VB.TextBox Text4 
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
         Height          =   300
         Left            =   1365
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   750
      End
      Begin VB.TextBox Text3 
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
         Height          =   300
         Left            =   1365
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   525
         Width           =   3480
      End
      Begin VB.TextBox Text2 
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
         Height          =   300
         Left            =   1365
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   210
         Width           =   3480
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
         Height          =   300
         Left            =   210
         MaxLength       =   5
         TabIndex        =   0
         Text            =   " "
         ToolTipText     =   "Número de Proveedor"
         Top             =   840
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C000&
         X1              =   240
         X2              =   10740
         Y1              =   1970
         Y2              =   1970
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5085
         TabIndex        =   18
         ToolTipText     =   "Categoria IVA"
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Emision"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7680
         TabIndex        =   14
         Top             =   720
         Width           =   2040
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Factura Nro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   5790
         TabIndex        =   12
         Top             =   1050
         Width           =   2115
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Contable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8025
         TabIndex        =   11
         Top             =   300
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   960
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FOFACOBR07.frx":0922
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "FOFACOBR07.frx":099C
      Top             =   1200
   End
   Begin VB.Label ONEGRA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   840
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label IMPIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   2415
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label IMPIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   2415
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1230
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
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "FOFACOBR07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOTACU#(32767)
Private Sub BOTNEXT_Click()
  '
  For J& = 0 To 32767
    TOTACU#(J&) = 0
  Next J&
  '
  TXTVALCOMPRA = ""
  For J& = 1 To ULTREG&("!DETFACOM")
    XY$ = REGLEIDO$("!DETFACOM", J&)
    '
    QFAC1# = CVD(Mid$(XY$, 95, 8)) 'Q. Cantidad p/ Facturar
    If Abs(QFAC1#) >= 0.05 Then
       INPU1% = CVI(Mid$(XY$, 143, 2))
       If INPU1% <= 0 Then
          Call MENSERR(24, "Todos los Items deben tener Imputación Contable.")
          Call Label12_DblClick
          Exit Sub
       End If
    End If
    '
    QPEFAC1 = CVD(Mid$(XY$, 87, 8)) 'Q. Ya Facturada
    QFAC1# = CVD(Mid$(XY$, 95, 8)) 'Q. Cantidad p/ Facturar
    PUFAC1# = CVD(Mid$(XY$, 111, 8)) 'Pr. Unit p/ Facturar
    CIXI% = CVI(Mid$(XY$, 25, 2))
    If CIXI% < 1 Then
       Call MENSERR(24, "Esta Modalidad no Aplicable\a Materiales sin Código.")
       Exit Sub
    End If
    NUOC& = CVS(Mid$(XY$, 67, 4))
    NPROX% = CVI(Mid$(XY$, 1, 2))
    '
    'NUEVO PARA GRABAR EL PRECIO UNITARIO BONIFICADO
    BN1 = COEFDESCU(TRIM$(Text12)) 'BONIFICACION
    PUNIBON# = PUFAC1# - (PUFAC1# * BN1 / 100)
    Call REPLA(XY$, MKD$(PUNIBON#), 119, 8)
    Call GRAREG("!DETFACOM", XY$, J&)
    'HASTA ACA NUEVO PARA GRABAR PRECIO UNITARIO BONIFICADO
    '
    If NUOC& > 0 Then
       REGOC$ = ULTIPRE$(NPROX%, CIXI%, NUOC&)
       CAPEDOC1# = CVS(Mid$(REGOC$, 55, 4)) 'Q. Pedida O/C
       PREUNOC1# = CVD(Mid$(REGOC$, 67, 8)) 'Pr. Unit. O/C
       COEFU = CVS(Mid$(REGOC$, 63, 4))
       '
       'AGREGADO PARA QUE COMPARE CON EL PRECIO EN PESOS SI NO PRESENTA MENSAJE DIFIERE PRECIO
       MONEX% = Asc(Mid$(REGOC$, 122, 1))
       If MONEX% < 1 Then MONEX% = 1
       PREUNOC1# = CVD(Mid$(REGOC$, 67, 8)) * TICAMONE(MONEX%) / TICAMONE(MONEMI%)
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       If InStr(EPAC$, "AHP") > 0 Then
         If (CostoRecpMovisto(CIXI%, NBORE&, MONEMI%) - PREUNOC1#) > 0.005 Then
          PREUNOC1TEX$ = FORMATNUM$(PREUNOC1#, "F12.2")
          OPX1% = COMALTER%("Atención Código: " & TRIM$(CODEXT$(CIXI%)) & " Precio Unitario : " & PREUNOC1TEX$ & "\Difiere del que Tenía en la Orden de Compra al Momento de la Recepción\\Cancelar\Continuar")
          If OPX1% >= 2 Then
             Call ActualizarCostoMovisto(CIXI%, NBORE&, PREUNOC1#, MONEMI%)
          End If
         End If
       End If
       'FIN MODIFICACION
       '
       If COEFU > 1 Then PREUNOC1# = PREUNOC1# / COEFU
       TXTVALCOMPRA = XVALO(TXTVALCOMPRA) + (PREUNOC1# * QPEFAC1)
       '
       If PREUNOC1# > 0.005 Then
         If Abs(PUFAC1# - PREUNOC1#) > 0.00005 Then
            ATX% = COMALTER%(TRIM$(CODEXT$(CIXI%)) & ": Precio Unitario Difiere de la Orden de Compra.\\Continuar\Editar")
            If ATX% = 2 Then
              Call Label12_DblClick
              Exit Sub
            End If
         End If
       End If
    End If
    '
    If QFAC1 > 0 Then
      If QFAC1 > QPEFAC1 Then
        ATX% = COMALTER%(TRIM$(CODEXT$(CIXI%)) & ": Cantidad Facturada Supera a la Cantidad Pendiente.\\Continuar\Editar")
        If ATX% = 2 Then
          Call Label12_DblClick
          Exit Sub
        End If
      End If
      '
      CADEC% = XVALO(CNTDC$)        ' cantidad de decimales
      PREUNOC1# = CVD(Mid$(REGOC$, 67, 8)) 'Pr. Unit. O/C
      If NUOC& < 1 Or PREUNOC1# < 0.00005 Then
         ' COMPARA PRECIO CON COSUNI
         COUN# = COSPESOS(CIXI%)
         If COUN# >= 0.00005 Then
           If PUNIBON# > COUN# + 0.00005 Then
              COSTUNIT$ = TRIM$(FORMATNUM$(PUNIBON#, "F12.4"))
              COSTREPP$ = TRIM$(FORMATNUM$(COSUNI(CIXI%) * TICAMONE(MONECOSTO%(CIXI%)), "F12.4"))
              If DERACCE%("AUTPRECO/NOMESS", "Autorizar Precios Sobre Costo Reposicion") < 2 Then
                Call MENSERR(24, "Codigo '" + TRIM$(CODEXT$(CIXI%)) + "' - Precio Unitario $ " + COSTUNIT$ + "\Superior Costo Autorizado $ " + COSTREPP$ + ".\   \Usuario sin Privilegios para Registrar con Precio Excedido.\Solicite Asignación de Derecho con Clave AUTPRECO.")
                Call Label12_DblClick
                Exit Sub
              Else
                TTXX$ = "Codigo '" & TRIM$(CODEXT$(CIXI%)) & "' - Precio Unitario: $ " & COSTUNIT$ & "\Superior a Costo de Reposicion: $ " & COSTREPP$
                If COMALTER%(TTXX$ + "\\Continuar\Editar") = 2 Then
                   Call Label12_DblClick
                   Exit Sub
                End If
              End If
           End If
         End If
      End If
    End If
    '
    IMPOFAC1# = (Int(100 * PUFAC1# * QFAC1 + 0.5)) / 100
    TOTACU#(INPU1%) = IMPOFAC1 + TOTACU#(INPU1%)
    '
  Next J&
  '
  On Error Resume Next
  RECUEP07.LINDICE.Clear
  FORFAC07.List1.Clear
  Call BLANARCH("!CARICOM")
  ITOTA1# = 0
  For JJ% = 1 To 32767 - 1
    If TOTACU#(JJ%) >= 0.005 Then
      'LO AGREGO AL LIST DEL FORM FACTURACION PPAL.
      '
      BON1 = XVALO(Text12)
      '
      CUEB$ = FORMATNUM(JJ%, "F5")
      Z$ = Space$(128)
      Call REPLA(Z$, CUEB$, 1, 5)
      DECUE$ = CODCUE$(JJ%) + " " + DENOCUE$(JJ%)
      Call REPLA(Z$, DECUE$, 8, 30)
      'NUEVO POR BONIFICACION
      'TOTACU#(JJ%) = TOTACU#(JJ%) - (TOTACU#(JJ%) * BON1 / 100)
      '
      IMPO1$ = FORMATNUM(TOTACU#(JJ%), "F13.2")
      Call REPLA(Z$, IMPO1$, 38, 13)
      RECUEP07.LINDICE.AddItem Z$
      '
      CUE% = JJ%
      IPA# = TOTACU#(JJ%)
      ITOTA1# = ITOTA1# + IPA#
      '
      y$ = REGBLAN$("!CARICOM")
      Call REPLA(y$, MKI$(CUE%), 1, 2)
      Call REPLA(y$, MKD$(IPA#), 3, 8)
      Call REGAPP("!CARICOM", y$)
      '
    End If
  Next JJ%
  '
  For U& = 1 To RECUEP07.LINDICE.ListCount
    RECUEP07.LINDICE.ListIndex = U& - 1
    Z$ = RECUEP07.LINDICE.TEXT
    FORFAC07.List1.AddItem Z$
  Next U&
  '
  FORFAC07.List1.Refresh
  FORFAC07.ITOGRA = FORMATNUM(ITOTA1#, "F10.2")
  FORFAC07.Label36 = FORMATNUM(ITOTA1#, "F10.2")
  TOTIMPU# = Val(Label9)
  FORFAC07.Text19 = Text12
  FORFAC07.Text19.BackColor = &HE0E0E0
  FORFAC07.Label35 = Label7
  '
  'limpia los casilleros de importes ingresados a diferentes tipos de iva
  For KKI% = 1 To 6
    FORFAC07.IMPOGRA(KKI%) = Space$(11)
    FORFAC07.IMPOIVA(KKI%) = Space$(8)
  Next KKI%
  '
  On Error GoTo 0
  Call CIERRARCH("!DETFACOM")
  Call CIERRARCH("!CARICOM")
  Hide
  '
End Sub

Private Sub Command12_Click()
    Command12.Enabled = False
    '
    Call APERBASDAT("COMPRAS")
    SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM OCOMENCA"
    SQLX$ = SQLX$ + " ORDER BY Nume_Ocom "
    'Set VRIMPTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim VRIMPTEMP As ADODB.Recordset
    Set VRIMPTEMP = New ADODB.Recordset
25  On Error Resume Next

    VRIMPTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0

    With VRIMPTEMP
        On Error Resume Next ' si hay algun error msgbox err
        .MoveFirst
        If Err Then
          On Error GoTo 0
          Call MENSERR(24, "No Aparecen Ordenes de Compra Emitidas")
          GoTo 99
        End If
        On Error GoTo 0
        JJ& = 0
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDIOCOM")
            Call REPLA(XXY$, MKS(!nume_ocom), 1, 4)
            Call REPLA(XXY$, !Refe_OcoM, 5, 44)
            JJ& = JJ& + 1
            Call GRAREG("!INDIOCOM", XXY$, JJ&)
            .MoveNext
        Loop
    End With
    Set VRIMPTEMP = Nothing
    Call SETULTREG("!INDIOCOM", JJ&)
    Call CIERRARCH("!INDIOCOM")
    Call FRESHECHO("!INDIOCOM")
    '
3   Call SELECHO("!INDIOCOM/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NPX& = XVALO(VALACT1$("!INDIOCOM")) ' ACA LEO DEL INDIOCOM
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Nota de Pedido"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU1$ = " " + NUDOCU$ + " "
    NUDOCU2$ = " " + NUDOCU$ + "-"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 99
    '
5   Call MUESTRADOC(DOCUNU&)
    '
    GoTo 3
    '
99  Command12.Enabled = True
    '
End Sub

Private Sub Command2_Click()
    '
    Call FORFAC07.BLAFACOM
    List1.Clear
    Unload Me
    '
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    Me.Label20.Caption = FORFAC07.Label20
    '
    Call AplicaFuenteAnchoFijo(List1, , 10)
    Call AplicaFuenteAnchoFijo(Text7, , 10)
    Call AplicaFuenteAnchoFijo(Label17, , 10)
    Call AplicaFuenteAnchoFijo(Label15, , 10)
    Call AplicaFuenteAnchoFijo(Label13, , 10)
    Call AplicaFuenteAnchoFijo(Text12, , 10)
    Call AplicaFuenteAnchoFijo(LBLIMPO, , 10)
    Call AplicaFuenteAnchoFijo(Label7, , 10)
    Call AplicaFuenteAnchoFijo(Label9, , 10)
    
    '
End Sub


Private Sub Label12_DblClick()
  '
  IMPOTOT# = 0
  '
'  FOFACOBR07.Label12 = MONEMI%
  ATRI$ = "/NC"
  ATRI$ = ATRI$ + "/TITUPAN=Imputacion Contable"
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(FORFAC07.Top + FORFAC07.Picture1.Top))
  ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(FORFAC07.Left + FORFAC07.Picture1.Left))
  '
  VTB% = VENTABU%("FACBORE" + ATRI$)
  '
  If VTB% >= 0 Then
    '
    List1.Clear
     For J& = 1 To ULTREG&("!DETFACOM")
        XY$ = REGLEIDO$("!DETFACOM", J&)
        NPRO1 = CVI(Mid$(XY$, 1, 2))      ' numero de proveedor
        CIXI1% = CVI(Mid$(XY$, 25, 2))    ' codigo
        '
        NBORE1 = CVS(Mid$(XY$, 27, 4))    ' n.br formato numero
        NUOC1& = CVS(Mid$(XY$, 67, 4))    ' n.o.compra formato numerico
        QPENFAC1 = CVD(Mid$(XY$, 87, 8))  ' cantidad pendiente de fabricacion
        QFAC1 = CVD(Mid$(XY$, 95, 8))     ' cantidad ingresada a facturar
        PUFAC1 = CVD(Mid$(XY$, 111, 8))   ' precio unitario
        INPU1% = CVI(Mid$(XY$, 143, 2))   ' cuenta contable numero
        IMPOR1# = (Int(100 * PUFAC1 * QFAC1 + 0.5)) / 100 ' importe total
        IMPOTOT# = IMPOTOT# + IMPOR1#
        '
        NPRO2$ = FORMATNUM(NPRO1, "F5")
        CIXI2$ = CODEXT(CIXI1%)
        DESCRI2$ = Left$(DESCRIT0(CIXI1%), 15)
        NBORE2$ = FORMATNUM(NBORE1, "F5")
        NUOC2$ = FORMATNUM(NUOC1&, "F5")
        QPENFAC2$ = FORMATNUM(QPENFAC1, "F10.1")
        QFAC2$ = FORMATNUM(QFAC1, "F10.1")
        PUFAC2$ = FORMATNUM(PUFAC1, "F10.4")
        INPUCO$ = FORMATNUM(INPU1%, "F5")
        IMPOR2$ = FORMATNUM(IMPOR1#, "F10.2")
        'TRAIGO EL PRECIO DE LA O. COMPRA CON MISMA FUNCION QUE LO MUESTRA AL PRINCIPIO, PERO VUELVO A LEER X Q SE PUEDE CAMBIAR.
        If NUOC1& > 0 Then
            REGOC$ = ULTIPRE$(XVALO(Text1), CIXI1%, NUOC1&)
            COEFU = CVS(Mid$(REGOC$, 63, 4))
            MONEX% = Asc(Mid$(REGOC$, 122, 1))
            If MONEX% < 1 Then MONEX% = 1
            '
            PREUN1# = CVD(Mid$(REGOC$, 67, 8)) * TICAMONE(MONEX%) / TICAMONE(MONEMI%)
            If COEFU > 1 Then PREUN1# = PREUN1# / COEFU
        Else
            PREUN1# = 0
        End If
                         
        ACU_IMPOCOMPRA# = ACU_IMPOCOMPRA# + (QFAC1 * PREUN1#) 'ACUMULO VALOR POR LA ORDEN DE COMPRA
        Label17 = FORMATNUM$(ACU_IMPOCOMPRA#, "F12.2")
        '
        LI1$ = Space(128)
        '
        Call REPLA(LI1$, CIXI2$, 1, 16)
        Call REPLA(LI1$, DESCRI2$, 17, 10)
        Call REPLA(LI1$, NBORE2$, 29, 5)
        Call REPLA(LI1$, NUOC2$, 36, 5)
        Call REPLA(LI1$, QPENFAC2$, 43, 10)
        Call REPLA(LI1$, QFAC2$, 55, 10)
        Call REPLA(LI1$, PUFAC2$, 67, 10)
        '
        Call REPLA(LI1$, IMPOR2$, 79, 11)
        Call REPLA(LI1$, INPUCO$, 89, 5) 'LE SAQUE UNO AL 90. ERA 91
        '
        List1.AddItem LI1$
      Next J&
    '
    LBLIMPO = FORMATNUM(IMPOTOT#, "F11.2")

    '
  End If
End Sub

Private Sub Label17_Change()
    Call Text7_Change
    TN1# = XVALO(Label17) - XVALO(Label15)
    Label13 = FORMATNUM(TN1#, "F11.2")

End Sub

Private Sub Label7_Change()
  '
  TT1# = Val(LBLIMPO) - Val(Label7)
  Label9 = FORMATNUM(TT1#, "F11.2")
  '
End Sub

Private Sub LBLIMPO_CHANGE()
    '
    Call Text12_Change
    TN1# = Val(LBLIMPO) - Val(Label7)
    Label9 = FORMATNUM(TN1#, "F11.2")
    '
End Sub

Sub List1_DblClick()
  '
  Call Label12_DblClick
  '
End Sub

Private Sub Text1_GotFocus()
  On Error Resume Next
    Command2.SetFocus
  On Error GoTo 0
End Sub

Private Sub Text11_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text12_Change()
    '
    TN1# = Val(LBLIMPO)
    BN1 = COEFDESCU(TRIM$(Text12))
    IMTO1# = (Int(100 * TN1# * (1 - BN1 / 100) + 0.5)) / 100
    BN2# = TN1# - IMTO1#
    Label7 = FORMATNUM(BN2#, "F11.2")
    Label9 = FORMATNUM(IMTO1#, "F11.2")
    '
End Sub
Private Sub Text13_GotFocus()
   Text1.SetFocus
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text3_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text4_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text5_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text6_GotFocus()
  Text1.SetFocus
End Sub

Private Sub Text7_Change()
    TN1# = XVALO(Label17)
    BN1 = COEFDESCU(TRIM$(Text7))
    IMTO1# = (Int(100 * TN1# * (1 - BN1 / 100) + 0.5)) / 100
    BN2# = TN1# - IMTO1#
    Label15 = FORMATNUM(BN2#, "F11.2")
    Label13 = FORMATNUM(IMTO1#, "F11.2")

End Sub

Private Sub Text9_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text1.SetFocus
   End If
End Sub
'
Function ULTIPRE$(NPROX%, CIXI%, NUOC&)
   '
   ' trae el registro de OCOMDETA correspondiente a la
   ' última compra no anulada del artículo y proveedor
   ' con precio registrado mayor que cero
   '
   Call APERBASDAT("COMPRAS")
   VUELTA% = 0
   '
20 UCO$ = REGBLAN$("OCOMDETA")
   Call REPLA(UCO$, MKI$(NPROX%), 7, 2)
   Call REPLA(UCO$, MKI$(CIXI%), 9, 2)
   '
   SQLX$ = "SELECT * FROM OCOMDETA "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
   SQLX$ = SQLX$ + "AND Npro_Ocom = " & NPROX% & " "
   SQLX$ = SQLX$ + "AND Nume_Ocom = " & NUOC& & " "
   If VUELTA% = 0 Then SQLX$ = SQLX$ + "AND Pre_Unit >= 0.00005 "
   SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
   'Set OcomdetaTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim OcomdetaTemp As ADODB.Recordset
   Set OcomdetaTemp = New ADODB.Recordset
25 On Error Resume Next
   OcomdetaTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
 
   On Error Resume Next
   With OcomdetaTemp
     .MoveLast
     If Err < 1 Then
         On Error GoTo 0
         Call REPLA(UCO$, MKS$(!nume_ocom), 1, 4)
         Call REPLA(UCO$, MKI$(FECHANUM(Format(!Femi_Ocom, "dd/mm/yy"))), 5, 2)
         Call REPLA(UCO$, !Uni_Mas, 11, 4)
         Call REPLA(UCO$, !Uni_Com, 59, 4)
         Call REPLA(UCO$, MKS$(!Coef_Unids), 63, 4)
         Call REPLA(UCO$, Chr$(!Mone_Emis), 122, 1)
         Call REPLA(UCO$, MKD$(!Pre_Unit), 67, 8)
         Call REPLA(UCO$, MKS$(!Desc_Item), 75, 4)
       Else
         On Error GoTo 0
         If VUELTA% < 1 Then
           VUELTA% = VUELTA% + 1
           GoTo 20
         End If
     End If
   End With
   OcomdetaTemp.Close
   Set OcomdetaTemp = Nothing
   '
   ULTIPRE$ = UCO$
   '
End Function



