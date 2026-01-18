VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form RECIIM14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Recibos por impuestos"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   12240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   12015
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "Carga de Recibos por Impuestos"
         Top             =   240
         Width           =   11760
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "Observaciones:"
         Top             =   6840
         Width           =   1335
      End
      Begin VB.TextBox Text28 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   7200
         Width           =   11750
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Eliminar Recibo Activo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5040
         TabIndex        =   34
         Top             =   8340
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir de Modo de Edición"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   8340
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Frame Frame5 
         Height          =   1575
         Left            =   9000
         TabIndex        =   28
         Top             =   2160
         Width           =   2895
         Begin VB.TextBox TXTIMPORTE 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   3
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Ingresar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   1150
            Width           =   2655
         End
         Begin VB.TextBox TXTTOTAL 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   "Total:"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Text            =   "Importe:"
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1815
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   8775
         Begin VB.CommandButton Command5 
            Caption         =   "!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   360
            Width           =   195
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   "Cod.Imp.:"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1395
            MaxLength       =   24
            TabIndex        =   1
            Top             =   360
            Width           =   975
         End
         Begin VB.Frame Frame6 
            Height          =   735
            Left            =   2520
            TabIndex        =   37
            Top             =   120
            Visible         =   0   'False
            Width           =   6015
            Begin VB.TextBox TXTID 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox TXTNUMLISTA 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   3240
               Locked          =   -1  'True
               TabIndex        =   39
               TabStop         =   0   'False
               Text            =   "Lista de Pago:"
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   38
               TabStop         =   0   'False
               Text            =   "Identificador:"
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   885
            Width           =   7335
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   "Descripción"
            Top             =   885
            Width           =   1095
         End
         Begin VB.TextBox TXTCONCEPTO 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   2
            Top             =   1320
            Width           =   7335
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "Concepto:"
            Top             =   1320
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   9000
         TabIndex        =   15
         Top             =   600
         Width           =   2895
         Begin VB.TextBox TXTFECHA 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1395
            TabIndex        =   22
            Top             =   240
            Width           =   1400
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "Fecha:"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TXTNURECIBO 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox TXTUSUARIO 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "Usuario"
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "N° Recibo:"
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton Command6 
            Caption         =   "!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   1200
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            Width           =   195
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   8775
         Begin VB.TextBox TXTSALDO 
            Alignment       =   1  'Right Justify
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
            Height          =   435
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            ToolTipText     =   "Doble Click Para Visualizar Cta.Cte."
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
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
            Height          =   435
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Text            =   "Saldo:"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "Nombre:"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   840
            Width           =   7335
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1395
            TabIndex        =   0
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "Cliente:"
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            Caption         =   "!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   1200
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   360
            Width           =   195
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar Recibo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9480
         TabIndex        =   6
         Top             =   8340
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   225
         Index           =   0
         Left            =   0
         OleObjectBlob   =   "RECIIM14.frx":0000
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2835
         Left            =   120
         TabIndex        =   10
         Top             =   3840
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   5001
         _Version        =   393216
         BackColorFixed  =   16501405
         SelectionMode   =   1
         AllowUserResizing=   1
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
   End
   Begin VB.PictureBox Picture999 
      Height          =   375
      Index           =   0
      Left            =   2640
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "RECIIM14.frx":005E
      Top             =   120
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Menu mnuver 
      Caption         =   "Ver"
      Begin VB.Menu mnuverimag 
         Caption         =   "Ver Imagen de Recibo"
      End
      Begin VB.Menu mnusp 
         Caption         =   "Situación de Pagos"
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "Edición"
      Begin VB.Menu Mnumod 
         Caption         =   "Modificar Recibo"
      End
   End
   Begin VB.Menu mnulp 
      Caption         =   "Pagos"
      Begin VB.Menu mnusalimp 
         Caption         =   "Generar Listado de Pagos de Impuestos de Clientes"
      End
   End
   Begin VB.Menu mnuRep 
      Caption         =   "Reportes"
      Begin VB.Menu mnurecitodo 
         Caption         =   "Recibos todos"
      End
      Begin VB.Menu mnurfecha 
         Caption         =   "Recibos por Rango de Fecha"
      End
      Begin VB.Menu mnudetareci 
         Caption         =   "Detalle de Recibos"
      End
      Begin VB.Menu mnureciusu 
         Caption         =   "Recibos por Usuario"
      End
      Begin VB.Menu Mnupt 
         Caption         =   "Pagos Realizados (Todos)"
      End
      Begin VB.Menu mnupl 
         Caption         =   "Pagos Realizados por N° de Listado"
      End
      Begin VB.Menu mnuop 
         Caption         =   "Pagos Realizados por N° Orden de Pago"
      End
      Begin VB.Menu mnupf 
         Caption         =   "Pagos Realizados por N° de Usuario"
      End
      Begin VB.Menu mnupu 
         Caption         =   "Pagos Realizados por Rango de Fechas"
      End
      Begin VB.Menu mnuppct 
         Caption         =   "Pagos Pendientes de Confirmación (ticket)"
      End
      Begin VB.Menu mnudpur 
         Caption         =   "Detalle de Pagos por Usuario Responsable"
      End
   End
   Begin VB.Menu mnuccte 
      Caption         =   "Cta.Ctes"
      Begin VB.Menu mnuctexnc 
         Caption         =   "Cta.Cte de Pagos x Cliente"
      End
      Begin VB.Menu MNUCNEXA 
         Caption         =   "Cta.Cte de Pagos x Cliente - CAJA ANEXA - "
      End
   End
   Begin VB.Menu mnusc 
      Caption         =   "Saldo de Caja"
      Begin VB.Menu mnusaldocaja 
         Caption         =   "Saldo de Caja"
      End
   End
   Begin VB.Menu mnl 
      Caption         =   "Listado de Pagos"
      Begin VB.Menu mnusalimpu 
         Caption         =   "Salida por Impuesto"
      End
      Begin VB.Menu mnuqpp 
         Caption         =   "Que hay Para Pagar?"
      End
   End
End
Attribute VB_Name = "RECIIM14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CONMUTAR(MODO$)
    If MODO$ = "EDICION" Then
      If Command1.Caption = "Edición de Recibo" Then Exit Sub
    Else
      If Command1.Caption = "Generar Recibo" Then Exit Sub
    End If
    '
    Text3.Enabled = Not Text3.Enabled
    Command3.Enabled = Not Command3.Enabled
    Command4.Visible = Not Command4.Visible
    TXTFECHA.Enabled = Not TXTFECHA.Enabled
    Command6.Enabled = Not Command6.Enabled
    Command2.Visible = Not Command2.Visible
    Frame6.Visible = Not Frame6.Visible
    If Command1.Caption = "Edición de Recibo" Then
       Command1.Caption = "Generar Recibo"
    Else
       Command1.Caption = "Edición de Recibo"
    End If
  

End Sub

Sub GRABA_EDIT()
   NC% = XVALO(Text3)
   FE% = CVINT(TXTFECHA)
   FECH = CVDAT(FE%)
   NURECI& = XVALO(TXTNURECIBO)
   SALIDA% = 0
   '
   If COMALTER("Confirmación de Modificación de Recibo\\Cancelar\Si Modificar") <> 2 Then GoTo 99
   COCOR$ = "00000000" & TRIM$(TXTNURECIBO): COCOR$ = Right$(COCOR$, 8)
   COCOR$ = "RI-" & COCOR$
   COLAR$ = COCOR$
   '
   CUECONTABLE% = CUECONTABLECAJA%
   IDX$ = "ID-" + TRIM$(TXTNURECIBO) + "-"

   On Error GoTo ERROR_GRABAR
   Dim RS1 As ADODB.Recordset
   Set RS1 = New ADODB.Recordset
25 On Error Resume Next
   CONDI$ = "nuclien =   " & NC% & " AND NUMECOMP = " & NURECI& & " AND (TipoMovim = 'CHRCIM' OR TipoMovim = 'RCIM')"
   Call ACTUREGISTRO("CAJABANC", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
   
   SQLX$ = "SELECT * FROM CAJABANC WHERE nuclien < 1"
   RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
   FLEXCONN.BeginTrans
   With RS1
        For I& = 1 To MSF1.Rows - 1
          IMPUESTO$ = MSF1.TextMatrix(I&, 1)
          CONCEPTO$ = MSF1.TextMatrix(I&, 2)
          IMPOX# = XVALO(MSF1.TextMatrix(I&, 3))
          CODIMPU% = MSF1.TextMatrix(I&, 4)
          NUMELISTA& = XVALO(MSF1.TextMatrix(I&, 6))
          '

40        .AddNew
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !CODICOM_COR = Left$(COCOR$, 12)
          !CODICOM_LAR = Left$(COLAR$, 24)
          !RefeCom = Left$("Carga de Impuesto de : " & IMPUESTO$, 48)
          !DESCRI_IMPU = Left$(IMPUESTO$, 48)
          !FECHCONTA = SafeDate(FECH)
          !fechacred = SafeDate(FECH)
          !CUECONTAI = 0
          !NUMEIMPU = CODIMPU%
          !ItemConc = Left$("", 8) 'nuevo p/ concepto
          !NUCLIEN = NC%
          !NuProve = 0
          !OpciMovim = "CCORR"
          !VARIMOVIM = ""
          !DESCRIMOV = Left$(CONCEPTO$, 48)
          !FECHEMISI = FECH
          !CodiEmpr = CodiEmp%
          !CodiSucu = 0
          !NumeTalo = 0
          !NUMECOMP = NURECI&
          !Status = 0
          On Error Resume Next
          !Mone_Emis = 1
          !Tipo_Cam = 1
          !ValAbsComp = Abs(IMPOX#)
          If IMPOX# > 0 Then
            !IDEBECOMP = IMPOX#
            !IHABECOMP = 0
            !TipoMovim = "CHRCIM"
            !ID = Left$(IDX$ + SAFETEXT$(I&), 24) ' GRABA EL IDENTIFICADOR 'NUMERO DE RECIBO + INDICE
            !OpciMovim = "CH"
          Else
            !IDEBECOMP = 0
            !IHABECOMP = Abs(IMPOTOTAL#)
            !TipoMovim = "RCIM"
            !OpciMovim = "CCORR"
            !ID = ""
          End If
          !Impo_Dola = 0
          !BASEIMPO = 0
          !NROCAI = 0
          On Error Resume Next
          !IncSubIva = ""
          On Error GoTo 0
         !FEVTOCAI = SafeDate(FECH)
         !ObserGen = ""
         !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
         !IdenBaja = String$(6, 0)
         !FeHoReal = Now()
         !NuAsiDiar = ""
         !Niv_Clasi = 0
         !CAE = ""
         !IDCAE = ""
         !IDCAENUM = 0
         !CODCARPIMP = ""
         !NOTAS = ""
         !INCOTERMS = ""
         !CondPag = 0
         !PEDIDOS = ""
         !OCompra = ""
         !Cli_Real = NC%
         !SimMone = "$" ' GRABA EL SIMBOLO DE LA MONEDA
         !AtraCance = 0
         !AtraCheque = 0
         !MARBORRA = 0
         !REFERENCIA_LISTA = ""
         '
         'DEscuentos general y descuento especial (Gabal) para presentar en F.Electronica.
         !Descu1 = ""
         !Descu2 = ""
         !Impodesc1 = 0
         !Impodesc2 = 0
           NUSER1% = USNBR% Mod 100
         !NUMEUSER = NUSER1%
         !NUMLISTA = NUMELISTA& ' NUMERO DE LISTA
         !NOTAS = Text28
         .Update
        Next I&
        If SALIDA% = 0 Then
          'CARGO INFORMACION PARA QUE GRABE EL MOVIMIENTO POR EL TOTAL
           IMPOTOTAL# = XVALO(TXTTOTAL)
           IMPOX# = 0
           CUECONTABLE% = CUECONTABLECLI%
           SALIDA% = 1
           GoTo 40
        End If
        '
    End With
ERROR_GRABAR:
    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      'SI HAY ERROR VUELVO PARA ATRAS MARCA MARBORRA.
      Call ACTUREGISTRO("CAJABANC", "MARBORRA", CONDI$, 0, REGAF&, "NOMESS")
      MsgBox "SE HA DETECTADO UN FALLO AL GRABAR,  NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      
      GoTo 99
    Else
      FLEXCONN.CommitTrans
      'SI SE HIZO LA OPERACION BORRO LOS REGISTROS
      CONDI$ = CONDI$ + " AND MARBORRA = 1 "
      Call BORRAREGISTROS("CAJABANC", CONDI$, REGAF&, "NOMESS")
      Call COMUNI("Grabación Realizada Exitosamente")
    End If
    RS1.Close
    Set RS1 = Nothing
    USUARIONOMBRE$ = USERNAM$(NUSER1%)
    FILTX$ = COCOR$ & ";" & TRIM$(Str$(NC%)) & ";" & USUARIONOMBRE$
    Call STDFORM("RCOB-GSM", FILTX$)

    Call BLANFORMU
    TXTNURECIBO = PRONURECIBO&
    Call CONMUTAR("NORMAL")
99

End Sub




Function PROXIMO_ID_GRILLA$()
  For I& = 1 To MSF1.Rows - 1
    
    POS1% = InStrRev(MSF1.TextMatrix(I&, 5), "-")
    If POS1% > 0 Then
       LEIDO = XVALO(Mid$(MSF1.TextMatrix(I&, 5), POS1% + 1))
       If LEIDO > MAYOR Then MAYOR = LEIDO
    End If
  Next I&
  MAYOR = MAYOR + 1
  IDX$ = "ID-" + TRIM$(TXTNURECIBO) + "-"
  PROXIMO_ID_GRILLA$ = IDX$ + SAFETEXT$(MAYOR)
End Function

Private Sub Command1_Click()

   Command1.Enabled = False
   If XVALO(Text3) < 1 Then
     Call COMUNI("Falta Código de Cliente")
     Call PINTATEXT(Text3)
     On Error Resume Next
     Text3.SetFocus
     On Error GoTo 0
     GoTo 99
   End If
   '
   If TRIM$(TXTFECHA) = "" Then
     Call COMUNI("Falta Fecha")
     Call PINTATEXT(TXTFECHA)
     TXTFECHA.SetFocus
     GoTo 99
   End If
   '
   FE% = HOY(HO$)
   FEPANT% = FECHANUM(TXTFECHA)
   If DIENTRE%(FE%, FEPANT%) > 30 Then
      Call COMUNI("La Fecha Seleccionada Difiere en más de 30 Días con la Fecha Actual")
      Call PINTATEXT(TXTFECHA)
      TXTFECHA.SetFocus
      GoTo 99
   End If
   '
   If DIENTRE%(FEPANT%, FE%) > 30 Then
      Call COMUNI("La Fecha Seleccionada Difiere en más de 30 Días con la Fecha Actual")
      Call PINTATEXT(TXTFECHA)
      TXTFECHA.SetFocus
      GoTo 99
   End If
   '
   If MSF1.Rows < 2 Then
     Call COMUNI("Falta Detalle de Recibo")
     Call PINTATEXT(Text7)
     Text7.SetFocus
     GoTo 99
   End If
   '
   CUECONTABLECLI% = CUECOINT(Control$("AMACLI", "ECUECON"))
   If CUECONTABLECLI% < 1 Then
      Call COMUNI("Falta Cuenta Madre de Clientes")
      GoTo 99
   End If
   '
   CUECONTABLECAJA% = CUECOINT(XVALO(Control$("", "CUENCAJA")))
   If CUECONTABLECAJA% < 1 Then
      Call COMUNI("Falta Cuenta Contable de Caja")
      GoTo 99
   End If

   'SI ES EDICION SE VA A LA RUTINA DE GRABA_EDIT
   If Command1.Caption = "Edición de Recibo" Then
      Call GRABA_EDIT
      GoTo 99
   End If
   '
   
   NC% = XVALO(Text3)
   FE% = CVINT(TXTFECHA)
   FECH = CVDAT(FE%)
   NURECI& = XVALO(TXTNURECIBO)
   SALIDA% = 0
   '
   If PRONURECIBO& <> NURECI& Then
       Call COMUNI("Se va a Actualizar el Número de Recibo")
       TXTNURECIBO = PRONURECIBO&
       GoTo 99
   End If
   '
   
   COCOR$ = "00000000" & TRIM$(TXTNURECIBO): COCOR$ = Right$(COCOR$, 8)
   COCOR$ = "RI-" & COCOR$
   COLAR$ = COCOR$
   '
   'ASIGNA PRIMERO LA CUENTA CONTABLE DE CAJA LUEGO CUANDO GRABA EL MOVIMIENTO POR EL IMPORTE TOTAL SE CAMBIA A LA DE D.X.VENTA
   CUECONTABLE% = CUECONTABLECAJA%
   IDX$ = "ID-" + TRIM$(TXTNURECIBO) + "-"
   '
   Call ABRECABAN
   On Error GoTo ERROR_GRABAR
   Dim RS1 As ADODB.Recordset
   Set RS1 = New ADODB.Recordset
   
   SQLX$ = "SELECT * FROM CAJABANC  WHERE nuclien < 1  "
   RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
   FLEXCONN.BeginTrans
   With RS1
        For I& = 1 To MSF1.Rows - 1
          IMPUESTO$ = MSF1.TextMatrix(I&, 1)
          CONCEPTO$ = MSF1.TextMatrix(I&, 2)
          IMPOX# = XVALO(MSF1.TextMatrix(I&, 3))
          CODIMPU% = MSF1.TextMatrix(I&, 4)
          '

40        .AddNew
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !CODICOM_COR = Left$(COCOR$, 12)
          !CODICOM_LAR = Left$(COLAR$, 24)
          !RefeCom = Left$("Carga de Impuesto de : " & IMPUESTO$, 48)
          !DESCRI_IMPU = Left$(IMPUESTO$, 48)
          !FECHCONTA = SafeDate(FECH)
          !fechacred = SafeDate(FECH)
          !CUECONTAI = CUECONTABLE%
          
          !ItemConc = Left$("", 8) 'nuevo p/ concepto
          !NUCLIEN = NC%
          !NuProve = 0
          !OpciMovim = "CCORR"
          !VARIMOVIM = ""
          !DESCRIMOV = Left$(CONCEPTO$, 48)
          !FECHEMISI = FECH
          !CodiEmpr = CodiEmp%
          !CodiSucu = 0
          !NumeTalo = 0
          !NUMECOMP = XVALO(TXTNURECIBO)
          !Status = 0
          On Error Resume Next
          !Mone_Emis = 1
          !Tipo_Cam = 1
          !ValAbsComp = Abs(IMPOX#)
          '
          If IMPOX# > 0 Then
            !IDEBECOMP = IMPOX#
            !IHABECOMP = 0
            !TipoMovim = "CHRCIM"
            !ID = Left$(IDX$ + SAFETEXT$(I&), 24) ' GRABA EL IDENTIFICADOR 'NUMERO DE RECIBO + INDICE
            !NUMEIMPU = CODIMPU%
          Else
            !IDEBECOMP = 0
            !IHABECOMP = Abs(IMPOTOTAL#)
            !TipoMovim = "RCIM"
            !ID = Left$(IDX$ + "0", 24) ' GRABA EL IDENTIFICADOR 'NUMERO DE RECIBO + INDICE
            !NUMEIMPU = 0
          End If
          !Impo_Dola = 0
          !BASEIMPO = 0
          !NROCAI = 0
          On Error Resume Next
          !IncSubIva = ""
          On Error GoTo 0
         !FEVTOCAI = SafeDate(FECH)
         !ObserGen = ""
         !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
         !IdenBaja = String$(6, 0)
         !FeHoReal = Now()
         !NuAsiDiar = ""
         !Niv_Clasi = 0
         !CAE = ""
         !IDCAE = ""
         !IDCAENUM = 0
         !CODCARPIMP = ""
         !NOTAS = ""
         !INCOTERMS = ""
         !CondPag = 0
         !PEDIDOS = ""
         !OCompra = ""
         !Cli_Real = NC%
         !SimMone = "$" ' GRABA EL SIMBOLO DE LA MONEDA
         !AtraCance = 0
         !AtraCheque = 0
         !MARBORRA = 0
         !REFERENCIA_LISTA = ""
         !NOTAS = Text28
         '
         'DEscuentos general y descuento especial (Gabal) para presentar en F.Electronica.
         !Descu1 = ""
         !Descu2 = ""
         !Impodesc1 = 0
         !Impodesc2 = 0
           NUSER1% = USNBR% Mod 100
         !NUMEUSER = NUSER1%
         !NUMLISTA = 0 ' NUMERO DE LISTA
         .Update
        Next I&
        If SALIDA% = 0 Then
          'CARGO INFORMACION PARA QUE GRABE EL MOVIMIENTO POR EL TOTAL
           IMPOTOTAL# = XVALO(TXTTOTAL)
           IMPOX# = 0
           SALIDA% = 1
           CUECONTABLE% = CUECONTABLECLI%
           GoTo 40
        End If
        '
    End With
ERROR_GRABAR:
    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      MsgBox "SE HA DETECTADO UN FALLO AL GRABAR,  NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      GoTo 99
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizada Exitosamente")
    End If
    RS1.Close
    Set RS1 = Nothing
    '
    USUARIONOMBRE$ = USERNAM$(NUSER1%)
    FILTX$ = COCOR$ & ";" & TRIM$(Str$(NC%)) & ";" & USUARIONOMBRE$
    Call STDFORM("RCOB-GSM", FILTX$)

    Call BLANFORMU
    
    TXTNURECIBO = PRONURECIBO&
    
99  Command1.Enabled = True

End Sub

Private Sub Command2_Click()
    Call BLANFORMU
    TXTNURECIBO = PRONURECIBO&
'    Command1.Caption = "Generar Recibo"
    Call CONMUTAR("NORMAL")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call SELECHO("DEUDOR1")
    NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
    If NCLIE% > 0 Then
      Text3.TEXT = TRIM$(Str$(NCLIE%))
    End If
    Command3.Enabled = True

End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   If DERACCE%("ElIRECIB", "Eliminación de Recibo") < 2 Then
     GoTo 99
   End If
   
   NC% = XVALO(Text3)
   NURECI& = XVALO(TXTNURECIBO)
   If COMALTER%("Confirmación:\\¿Confirma Eliminar Recibo: ? " & TRIM$(NURECI&) & "\Cancelar") <> 1 Then
       GoTo 99
   End If
   
   CONDI$ = "nuclien =   " & NC% & " AND NUMECOMP = " & NURECI& & " AND (TipoMovim = 'CHRCIM' OR TipoMovim = 'RCIM')"
   CONDI$ = CONDI$ + " AND STATUS < 1 "
   Call BORRAREGISTROS("CAJABANC", CONDI$, REGAF&, "NOMESS")
   If REGAF& > 0 Then
      Call COMUNI("Eliminación Realizada Exitosamente")
   Else
      Call COMUNI("No se Pudo Realizar la Eliminación")
   End If
   '
   Call CONMUTAR("NORMAL")
99 Command4.Enabled = True
   Call BLANFORMU
   
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call SELECHO("ITEMDAT/IMPUESTOS")
    CI1% = XVALO(VALACT1$("ITEMDAT"))
    Text7 = CI1%
    Command5.Enabled = True

End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call SELECHO("SELFECHA")
   TXTFECHA.TEXT = VALACT1$("SELFECHA")
99 Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   '
   If XVALO(Text3) < 1 Then
     Call COMUNI("Falta Código de Cliente")
     Call PINTATEXT(Text3)
     Text3.SetFocus
     GoTo 99
   End If
   '
   If XVALO(Text7) < 1 Then
     Call COMUNI("Falta Código de Impuesto")
     Call PINTATEXT(Text7)
     Text7.SetFocus
     GoTo 99
   End If
   '
   If TRIM$(Text6(5)) = "" Then
     Call COMUNI("Falta Descripción de Impuesto")
     Call PINTATEXT(Text6(5))
     GoTo 99
   End If
   '
   If TRIM$(TXTCONCEPTO) = "" Then
     Call COMUNI("Falta Concepto")
     Call PINTATEXT(TXTCONCEPTO)
     TXTCONCEPTO.SetFocus
     GoTo 99
   End If
   '
   If TRIM$(TXTFECHA) = "" Then
     Call COMUNI("Falta Fecha")
     Call PINTATEXT(TXTFECHA)
     TXTFECHA.SetFocus
     GoTo 99
   End If
   '
   FE% = HOY(HO$)
   FEPANT% = FECHANUM(TXTFECHA)
   If DIENTRE%(FE%, FEPANT%) > 30 Then
      Call COMUNI("La Fecha Seleccionada Difiere en más de 30 Días con la Fecha Actual")
      Call PINTATEXT(TXTFECHA)
      TXTFECHA.SetFocus
      GoTo 99
   End If
   '
   If DIENTRE%(FEPANT%, FE%) > 30 Then
      Call COMUNI("La Fecha Seleccionada Difiere en más de 30 Días con la Fecha Actual")
      Call PINTATEXT(TXTFECHA)
      TXTFECHA.SetFocus
      GoTo 99
   End If
   '
   If XVALO(TXTIMPORTE) < 1 Then
     Call COMUNI("Falta Importe o Valor No Válido")
     Call PINTATEXT(TXTIMPORTE)
     TXTIMPORTE.SetFocus
     GoTo 99
   End If
   '
   If COMALTER%("Confirmación:\\¿Confirma Cargar? " & TRIM$(Text6(5) & " " & TRIM$(TXTIMPORTE)) & "\Cancelar") <> 1 Then
       GoTo 99
   End If
   If TRIM$(TXTID) = "" Then
      TXTID = PROXIMO_ID_GRILLA$
   End If
   
   REG& = MSF1.Rows
   MSF1.Rows = MSF1.Rows + 1
   MSF1.TextMatrix(REG&, 1) = TRIM$(Text6(5))
   MSF1.TextMatrix(REG&, 2) = TRIM$(TXTCONCEPTO)
   MSF1.TextMatrix(REG&, 3) = TRIM$(TXTIMPORTE)
   MSF1.TextMatrix(REG&, 4) = TRIM$(Text7)
   MSF1.TextMatrix(REG&, 5) = TRIM$(TXTID)
   MSF1.TextMatrix(REG&, 6) = TRIM$(TXTNUMLISTA)
   '
   Text7.SetFocus
   '
   
   Call ACTUTOTAL
   Text7 = ""
   TXTCONCEPTO = ""
   TXTIMPORTE = ""
   Text28 = ""
   TXTID = ""
   TXTNUMLISTA = ""

99 Command7.Enabled = True

End Sub

Sub ACTUTOTAL()
   ACUMU# = 0
   For I& = 1 To MSF1.Rows - 1
      IMPOX# = ROUND(XVALO(MSF1.TextMatrix(I&, 3)), 2)
      ACUMU# = ACUMU# + IMPOX#
   Next I&
   TXTTOTAL = FORMATNUM$(ACUMU#, "F12.2")
   
End Sub
Private Sub Form_Load()
    
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    Call CENTRAFORM(Me)
    
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
15  EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    '
    'VALIDACION PARA AMBEST POR SI NO PAGA
''If InStr(EPAC$, "AMBEST") > 0 Then
'     FF% = FECHANUM("20/06/14")
'     If FF% <= HOY(HO$) Then
'       Call Form_Load
'     End If
''End If

    Call ABRECONEXION
    TXTFECHA = FECHATEX$(HOY(HO$))
    TXTNURECIBO = PRONURECIBO&
    TXTUSUARIO = USERNAME$
    '
    Call CARGA_ENCABEZADO(RECIIM14.MSF1, "Impuesto/A32;Concepto/A32;Importe/F12.2;Cod_Impuesto/A0;ID/A0;NUMELISTA/F0", RECIIM14)
    Call CREACAMPO("", "CAJABANC", "MARBORRA", 3, 0)
    '
    Call CREACAMPO("", "CAJABANC", "MARBORRA", 3, 0)
    Call CREACAMPO("", "CAJABANC", "NUMEUSER", 3, 0)
    Call CREACAMPO("", "CAJABANC", "NUMEIMPU", 6, 0, 1)
    Call CREACAMPO("", "CAJABANC", "NUMLISTA", 6, 0, 1)
    Call CREACAMPO("", "CAJABANC", "DESCRI_IMPU", 10, 64)
    Call CREACAMPO("", "CAJABANC", "STATUS", 3, 0, 1)
    Call CREACAMPO("", "CAJABANC", "REFERENCIA_LISTA", 10, 64)
    Call CREACAMPO("", "CAJABANC", "ID", 10, 24)
    '
    Call PoneEnCeroNull("CAJABANC", "STATUS", REGAF&, "NOMESS")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call FINAL("")
End Sub

Private Sub MnuEditar_Click()

End Sub


Private Sub MnuModificar_Click()

End Sub


Private Sub MnuArchivo_Click()

End Sub

Private Sub mnuAjus_Click()

    Call SELECHO("DEUDOR1")
    NC% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
    If NC% < 1 Then Exit Sub

    SQLX$ = "SELECT SUM(IDEBECOMP) AS DEBEX , SUM(IHABECOMP) AS HABEX  FROM CAJABANC WITH(NOLOCK)"
    SQLX$ = SQLX$ + " where NuClien = " & NC%
    SQLX$ = SQLX$ + " AND (TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + " or TipoMovim = 'CHRCIM' "
    SQLX$ = SQLX$ + " or TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + " GROUP BY NUCLIEN "
    Set RS = READSET(SQLX$)
    With RS
     SALDOX# = XVALO(!DEBEX) - XVALO(!HABEX)
    End With
    '
    AJUSTE14.Text28 = ""
    AJUSTE14.Text2 = ""
    AJUSTE14.Text1 = FORMATNUM$(SALDOX#, "F12.2")
    AJUSTE14.Show 1
    '
    'SI SE CERRO EL FRM DE AJUSTES
    If ROUND(XVALO(AJUSTE14.Text1)) + ROUND(XVALO(AJUSTE14.Text2)) = 0 Then Exit Sub
    '
    SAJUSTADO# = XVALO(AJUSTE14.SALDO_AJUSTADO#)
    Observa$ = AJUSTE14.Text28
    CCONT% = CUECOINT%(Control$("", "CUECHECA"))
    If CCONT% < 1 Then
       Call COMUNI("Falta Cuenta Contable de Ajustes")
       Exit Sub
    End If
    DIFERENCIA# = SAJUSTADO# - SALDOX#
    If DIFERENCIA# = 0 Then
       Call COMUNI("Saldo Ajustado Igual a Saldo ")
       Exit Sub
    End If
    '
    FEX = HOY(HO$)
    FETX$ = FECHATEX$(FEX)
    NRECIB$ = "AS-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
    '
    CONDI$ = " TipoMovim = 'AJUD' AND CodiCom_Cor LIKE '" & NRECIB$ + "%' "
    ORMA% = XVALO(Mid$(VALOBADA("CAJABANC", "max(CodiCom_Cor)", CONDI$, "NOMESS"), 11, 2))
    ORMX$ = TRIM$(Str$(ORMA% + 1))
    While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
    NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
    '
    CUECONTABLECLI% = CUECOINT(Control$("AMACLI", "ECUECON"))
    CUECONTABLE% = CUECONTABLECLI%
    Call ABRECABAN
    Dim RS1 As ADODB.Recordset
    Set RS1 = New ADODB.Recordset
25  On Error Resume Next
   
    SQLX$ = "SELECT * FROM CAJABANC  WHERE nuclien < 1  "
    RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
    With RS1
          '
40        .AddNew
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !CODICOM_COR = Left$(NRECIB$, 12)
          !CODICOM_LAR = Left$(NRECIB$, 24)
          !RefeCom = Left$("Aj.Saldo - " + RASOCLI$(NC%), 48)
          !DESCRI_IMPU = "Aj.Saldo - " + RASOCLI$(NC%)
          FEX1% = FEX
          !FECHCONTA = CVDAT(FEX1%)
          !fechacred = CVDAT(FEX1%)
          !CUECONTAI = CUECONTABLE%
          !NUMEIMPU = 0
          !ItemConc = "" 'nuevo p/ concepto
          !NUCLIEN = NC%
          !NuProve = 0
          !OpciMovim = "CCORR"
          !VARIMOVIM = ""
          !DESCRIMOV = Left$("Aj.Saldo - " + RASOCLI$(NC%), 48)
          !FECHEMISI = CVDAT(FEX1%)
          !CodiEmpr = CodiEmp%
          !CodiSucu = 0
          !NumeTalo = 0
          !NUMECOMP = XVALO(Mid$(NRECIB$, 4, 6))
          !Status = 0
          On Error Resume Next
          !Mone_Emis = 1
          !Tipo_Cam = 1
          !ValAbsComp = Abs(IMPOX#)
          '
          If SALIDA% = 0 Then
            If DIFERENCIA# > 0 Then
              !IDEBECOMP = Abs(DIFERENCIA#)
              !IHABECOMP = 0
            Else
              !IHABECOMP = Abs(DIFERENCIA#)
              !IDEBECOMP = 0
            End If
            !TipoMovim = "AJUD"
            !ID = 0

          Else
           If DIFERENCIA# > 0 Then
            !IDEBECOMP = 0
            !IHABECOMP = Abs(DIFERENCIA#)
           Else
            !IDEBECOMP = Abs(DIFERENCIA#)
            !IHABECOMP = 0
           End If
            !TipoMovim = "AJUSALD"
            !ID = 0
          End If
          !Impo_Dola = 0
          !BASEIMPO = 0
          !NROCAI = 0
          On Error Resume Next
          !IncSubIva = ""
          On Error GoTo 0
         !FEVTOCAI = CVDAT(FEX1%)
         !ObserGen = ""
         !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
         !IdenBaja = String$(6, 0)
         !FeHoReal = Now()
         !NuAsiDiar = ""
         !Niv_Clasi = 0
         !CAE = ""
         !IDCAE = ""
         !IDCAENUM = 0
         !CODCARPIMP = ""
         !NOTAS = ""
         !INCOTERMS = ""
         !CondPag = 0
         !PEDIDOS = ""
         !OCompra = ""
         !Cli_Real = NC%
         !SimMone = "$" ' GRABA EL SIMBOLO DE LA MONEDA
         !AtraCance = 0
         !AtraCheque = 0
         !MARBORRA = 0
         !REFERENCIA_LISTA = ""
         !NOTAS = Observa$
         '
         'DEscuentos general y descuento especial (Gabal) para presentar en F.Electronica.
         !Descu1 = ""
         !Descu2 = ""
         !Impodesc1 = 0
         !Impodesc2 = 0
           NUSER1% = USNBR% Mod 100
         !NUMEUSER = NUSER1%
         !NUMLISTA = 0 ' NUMERO DE LISTA
         .Update
        If SALIDA% = 0 Then
          'CARGO INFORMACION PARA QUE GRABE EL MOVIMIENTO POR EL TOTAL
           SALIDA% = 1
           CUECONTABLE% = CCONT%
           GoTo 40
        End If
        '
    End With
    RS1.Close
    Set RS1 = Nothing

End Sub


Private Sub MNUCNEXA_Click()
   CCTEANEX14.Show 1
End Sub

Private Sub mnuctexnc_Click()
   CCTEIMP14.Show 1
End Sub

Private Sub mnudetareci_Click()
   Call FINAL("?RECIDETA")
End Sub

Private Sub mnudpur_Click()
   Call FINAL("?PAGOS7")
End Sub

Private Sub Mnumod_Click()
  
  CONDI$ = "TipoMovim = 'RCIM' AND STATUS = 0 ORDER BY CODICOM_COR "
  Call CARGALISEL("EDIRECI", "CAJABANC", "CODICOM_COR/A12;NUCLIEN/F6+DEUDOR12/A24;IHABECOMP/F12.2", CONDI$, "NOMESS")
  Call SELECHO("EDIRECI")
  RECI1$ = TRIM$(VALACT1$("EDIRECI"))
  If RECI1$ = "" Then Exit Sub
  '
  Call CONMUTAR("EDICION")
  MSF1.Rows = 1
  VUELTA% = 0
  CONDI$ = "CODICOM_COR = '" & RECI1$ & "' AND TipoMovim = 'CHRCIM'"
  SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE " & CONDI$
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
          I& = MSF1.Rows
          MSF1.Rows = MSF1.Rows + 1

          IMPUESTO$ = SAFETEXT$(!DESCRI_IMPU)
          IMPOX# = XVALO(!IDEBECOMP)
          CONCEPTO$ = SAFETEXT$(!DESCRIMOV)
          CODIMPU% = XVALO(!NUMEIMPU)
          If VUELTA% = 0 Then
             Text3 = XVALO(!NUCLIEN)
             USUARIOX$ = SAFETEXT$(!NUMEUSER)
             POS1% = InStr(USUARIOX$, "-")
             If POS1% > 0 Then
               NUMEUS% = Left$(USUARIOX$, POS1% - 1)
               NAMEUS$ = Mid$(USUARIOX$, POS1% + 1)
             End If
             TXTUSUARIO = NAMEUS$
             TXTFECHA = FECHATEX$(CVINT(!FECHEMISI))
             NURECI$ = Mid$(!CODICOM_COR, 4)
             TXTNURECIBO = XVALO(NURECI$)
             Text28 = SAFETEXT$(!NOTAS)
             TXTID = ""
             TXTNUMLISTA = ""

             VUELTA% = 1
          End If
          '
          MSF1.TextMatrix(I&, 1) = IMPUESTO$
          MSF1.TextMatrix(I&, 2) = CONCEPTO$
          MSF1.TextMatrix(I&, 3) = FORMATNUM$(IMPOX#, "F12.2")
          MSF1.TextMatrix(I&, 4) = SAFETEXT$(CODIMPU%)
          MSF1.TextMatrix(I&, 5) = SAFETEXT$(!ID)
          MSF1.TextMatrix(I&, 6) = SAFETEXT$(!NUMLISTA)
          .MoveNext
    Loop
  End With
  '
  'ESCRIBO EL IMPORTE TOTAL DEL RECIBO ACTIVO
  CONDI$ = "CODICOM_LAR = '" & RECI1$ & "' AND TipoMovim = 'RCIM'"
  TXTTOTAL = TRIM$(FORMATNUM$(XVALO(VALOBADA("CAJABANC", "IHABECOMP", CONDI$, "")), "F12.2"))
  
  
  End Sub

Private Sub mnurp_Click()
   RENDPAG14.Show 1
End Sub

Private Sub mnuop_Click()
Call FINAL("?PAGOS3")
End Sub

Private Sub mnupf_Click()
Call FINAL("?PAGOS4")
End Sub

Private Sub mnupl_Click()
Call FINAL("?PAGOS2")
End Sub

Private Sub mnuppct_Click()
   Call FINAL("?PAGOS6")
End Sub

Private Sub Mnupt_Click()
   Call FINAL("?PAGOS")
End Sub

Private Sub mnupu_Click()
Call FINAL("?PAGOS5")
End Sub

Private Sub mnuqpp_Click()
    Call FINAL("?QUEPAGAR")
End Sub

Private Sub mnurecitodo_Click()
   Call FINAL("?RECITODO")
End Sub

Private Sub mnureciusu_Click()
   Call FINAL("?RECIUSER")
End Sub

Private Sub mnurfecha_Click()
  Call FINAL("?RECIRFEC")
End Sub


Private Sub mnusaldocaja_Click()
    Call FINAL("?SALDCAJA")
End Sub

Private Sub MnuSalimp_Click()
   LISTAP14.Show 1
End Sub

Private Sub mnusalimpu_Click()
   Call FINAL("?SALIIMPU")
End Sub

Private Sub mnusp_Click()
   SITPAGO.Show 1
End Sub

Private Sub mnuverimag_Click()
  CONDI$ = "TipoMovim = 'RCIM' AND STATUS = 0 ORDER BY CODICOM_COR "
  Call CARGALISEL("EDIRECI", "CAJABANC", "CODICOM_COR/A12;NUCLIEN/F6+DEUDOR12/A24;IHABECOMP/F12.2", CONDI$, "NOMESS")
  Call SELECHO("EDIRECI")
  RECI1$ = TRIM$(VALACT1$("EDIRECI"))
  If RECI1$ = "" Then Exit Sub
  '
  COCOR$ = TRIM$(VALACT1$("EDIRECI"))
  NC% = XVALO(Left$(VALACT2$("EDIRECI"), 6))
  'TRAIGO EL NUMERO DE USUARIO DEL MOVMIEMNTO
  CONDI$ = "TipoMovim = 'RCIM' AND NUCLIEN = " & NC%
  NUMUSU = XVALO(VALOBADA("CAJABANC", "NUMEUSER", CONDI$, "NOMESS")) Mod 100
  USUARIONOMBRE$ = USERNAM$(NUMUSU)
  FILTX$ = COCOR$ & ";" & TRIM$(Str$(NC%)) & ";" & USUARIONOMBRE$
  Call STDFORM("RCOB-GSM", FILTX$)


End Sub

Private Sub MSF1_DblClick()
    I& = MSF1.MouseRow
    j& = MSF1.MouseCol
    If I& = 0 Or I& > MSF1.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF1.Row = I&
    MSF1.COL = j&
    On Error Resume Next

    REG& = I&
    If REG& < 1 Then
        Exit Sub
    End If
    '
    'BUSCO NUMERO DE LISTA PARA VALIDAR
    NUMELI& = XVALO(MSF1.TextMatrix(REG&, 6))
    IDENTI$ = (MSF1.TextMatrix(REG&, 5))
    TXTID = IDENTI$
    TXTNUMLISTA = NUMELI&
    '
    'SI TIENE NUMERO DE LISTA ADVIERTO QUE DEBE QUITAR DE LA LISTA PRIMERO DE ESTA MANERA ANULA EL PAGO DE LA MISMA.
    If NUMELI& > 0 Then
       Call COMUNI("El Impuesto Seleccionado Está Incluido Dentro de La Lista de Pago N°: " & NUMELI& & "\Deberá Primero Quitarla de La Lista de Pago")
       Exit Sub
    End If
       
    Text6(5) = MSF1.TextMatrix(REG&, 1)
    TXTCONCEPTO = MSF1.TextMatrix(REG&, 2)
    TXTIMPORTE = MSF1.TextMatrix(REG&, 3)
    Text7 = MSF1.TextMatrix(REG&, 4)
    '
    Text7.SetFocus
    '
    REGI% = REG&
    Call borrafila(REGI%, RECIIM14, MSF1)
    '
    Call ACTUTOTAL

End Sub


Private Sub Text3_Change()
   '
   NC% = XVALO(Text3)
   CONDI$ = "Nume_Cli = " & NC% & " And Stat_Cli >= 0"
   If CantRegistros("DEUDOR12", CONDI$, "NOMESS") > 0 Then
     Text6(3) = VALOBADA("DEUDOR12", "RASO_CLI", CONDI$, "NOMESS")
   Else
     Text6(3) = ""
   End If
   '
   Call VERISALDO(NC%, RECIIM14.TXTSALDO)
End Sub

Private Sub TXTIMPUESTO_Change()
End Sub
Function PRONURECIBO&()
    CONDI$ = "TipoMovim = 'RCIM' "
    CONDI$ = CONDI$ + " AND CodiCom_Cor LIKE 'RI%' "
    PRONURECIBO& = 1 + XVALO(VALOBADA("CAJABANC", "MAX(NUMECOMP)", CONDI$, "NOMESS"))
End Function

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub Text6_Change(Index As Integer)
'   If Index <> 3 Then Exit Sub
'   Static NOMBREX$, NC%
'   If NOMBREX$ <> "" Then
'      If TRIM$(NOMBREX$) <> TRIM$(Text6) Then
'         If MSF1.Rows > 1 Then
'           OPX% = COMALTER("No se Ha Grabado el Movimiento del Cliente: " & NOMBREX & "\\Continuar\Cancelar")
'           If OPX% <> 2 Then
'              Text6(3) = NOMBREX$
'              Text3 = NC%
'           End If
'
'         End If
'      End If
'   Else
'      NOMBREX$ = Text6(3)
'      NC% = XVALO(Text3)
'   End If
   
End Sub

Private Sub Text7_Change()
   Text6(5) = ""
   CI% = XVALO(Text7)
   If CI% < 1 Then Exit Sub
   '
   CONDI$ = "CODINT = " & CI% & " AND STATUS >= 0 "
   If CantRegistros("ITEMS", CONDI$, "NOMESS") > 0 Then
      SQLX$ = "SELECT * FROM ITEMS WITH(NOLOCK) WHERE CODINT = " & CI%
      Set RS = READSET(SQLX$)
      With RS
        Text6(5) = SAFETEXT$(!Descripcion)
      End With

   End If



End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TXTCONCEPTO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Sub BLANFORMU()
   Text3 = ""
   Text7 = ""
   TXTCONCEPTO = ""
   Text28 = ""
   'TXTFECHA = ""
   TXTNURECIBO = ""
   'TXTUSUARIO = ""
   TXTIMPORTE = ""
   TXTTOTAL = ""
   MSF1.Rows = 1
   TXTID = ""
   TXTNUMLISTA = ""
End Sub

Private Sub TXTIMPORTE_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Command7_Click
  End If

End Sub

Private Sub TXTIMPORTE_LostFocus()
   TXTIMPORTE = FORMATNUM$(XVALO(TXTIMPORTE), "F12.2")

End Sub


Private Sub TXTSALDO_DblClick()
  CCTEIMP14.TXTID = XVALO(TXTID)
  CCTEIMP14.Show 1

End Sub

