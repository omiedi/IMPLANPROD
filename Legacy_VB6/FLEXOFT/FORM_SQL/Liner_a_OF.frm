VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Liner_a_OF 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asociar Liner a Orden de Fabricación"
   ClientHeight    =   5850
   ClientLeft      =   1050
   ClientTop       =   1155
   ClientWidth     =   15495
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   15495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   38
      Top             =   3900
      Width           =   15495
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "<<<<<<    Liner Asociados a la Orden de Fabricación Activa    >>>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   1
         Left            =   4000
         TabIndex        =   39
         Top             =   120
         Width           =   8970
      End
   End
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   15495
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "<<<<<<     Asignación Liner a Orden de Fabricación     >>>>>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   0
         Left            =   5040
         TabIndex        =   37
         Top             =   120
         Width           =   7290
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Liner Ingresados a Depósito General (Solo Con Saldo)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   15495
      Begin VB.TextBox TXTSALDOLOTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   12480
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   " "
         Top             =   3000
         Width           =   780
      End
      Begin VB.CommandButton Command40 
         Appearance      =   0  'Flat
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
         Left            =   12390
         TabIndex        =   41
         Top             =   1320
         Width           =   200
      End
      Begin VB.TextBox TXTNUMOPERARIO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11760
         TabIndex        =   40
         Top             =   1320
         Width           =   600
      End
      Begin VB.TextBox TXTELEMSELEC 
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
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   960
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox TXTKGS 
         Alignment       =   1  'Right Justify
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
         Left            =   14400
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   " "
         Top             =   2520
         Width           =   780
      End
      Begin VB.TextBox TXTCORTESINFO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11205
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   " "
         Top             =   2160
         Width           =   780
      End
      Begin VB.TextBox TXTCORTXBSON 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14400
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   " "
         Top             =   2160
         Width           =   780
      End
      Begin VB.TextBox TXTCODINT 
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
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1320
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox TXTLOTE2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   12600
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   " "
         Top             =   1800
         Width           =   1020
      End
      Begin VB.CommandButton Command2 
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
         Height          =   315
         Left            =   13920
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "O.K."
         Top             =   3000
         Width           =   1245
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Registrar Peso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12240
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Lee el Peso de La Balanza"
         Top             =   2520
         Width           =   1485
      End
      Begin VB.TextBox TXTCORTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14400
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   " "
         Top             =   1800
         Width           =   780
      End
      Begin VB.TextBox TXTCANTCORTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11205
         TabIndex        =   20
         Text            =   " "
         Top             =   3000
         Width           =   780
      End
      Begin VB.TextBox TXTNECESIDADCORTES 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11205
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   " "
         Top             =   1800
         Width           =   780
      End
      Begin VB.TextBox TXTRESTAN 
         Alignment       =   1  'Right Justify
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
         Left            =   11205
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   " "
         Top             =   2520
         Width           =   780
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
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
         Left            =   12390
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.TextBox TXTNOP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox TXTCANPROY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14235
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   " "
         Top             =   240
         Width           =   955
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   7.5
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13080
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   175
      End
      Begin MSFlexGridLib.MSFlexGrid MSF0 
         Height          =   3015
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Doble Clic Para Asignar a Orden de Fabricación/Clic Botón Derecho Eliminar Registro"
         Top             =   240
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   5318
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operario:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   10800
         TabIndex        =   43
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label LBLNOMBREOPERARIO 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   12600
         TabIndex        =   42
         Top             =   1320
         Width           =   2590
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kgs:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   13880
         TabIndex        =   34
         Top             =   2520
         Width           =   405
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Informados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   10200
         TabIndex        =   32
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Unidades x Bolsón: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   12360
         TabIndex        =   29
         Top             =   2160
         Width           =   1965
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   11660
         TabIndex        =   27
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Restan:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   10200
         TabIndex        =   23
         Top             =   2640
         Width           =   885
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese Cantidad de Liner:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   9840
         TabIndex        =   21
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Corte m.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   13440
         TabIndex        =   19
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Necesidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   5
         Left            =   10200
         TabIndex        =   16
         Top             =   1755
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10845
         TabIndex        =   15
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label DESCRIOP 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   12600
         TabIndex        =   14
         Top             =   960
         Width           =   2590
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.Proy.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   13260
         TabIndex        =   11
         Top             =   315
         Width           =   885
      End
      Begin VB.Label LBLDESCRIOF 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   10305
         TabIndex        =   9
         Top             =   600
         Width           =   4885
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de O/F: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10200
         TabIndex        =   8
         Top             =   315
         Width           =   855
      End
      Begin VB.Label LBLOF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   11070
         TabIndex        =   7
         ToolTipText     =   "Seleccione Orden de Fabricación Asociada"
         Top             =   240
         Width           =   1890
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1440
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "Liner_a_OF.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   3240
      OleObjectBlob   =   "Liner_a_OF.frx":0234
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   1410
      Left            =   0
      TabIndex        =   3
      Top             =   4395
      Width           =   15525
      _ExtentX        =   27384
      _ExtentY        =   2487
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   1650
      Left            =   120
      TabIndex        =   44
      Top             =   6000
      Width           =   15465
      _ExtentX        =   27279
      _ExtentY        =   2910
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   10468997
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edición"
   End
   Begin VB.Menu mnuCons 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu mnuut 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuaccesodirecto 
      Caption         =   "Accesos Directos"
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuDepoCinESl 
         Caption         =   "Depósito de Cintilla y Eslingas"
      End
   End
End
Attribute VB_Name = "Liner_a_OF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub blanformu1()
        TXTLECTURA = ""
        TXTCODIGOELEM = ""
        TXTPESOORIG = ""
        TXTDESCRIELEM = ""
        TXTPESO = ""
        TXTMETROS = ""
        TXTLOTE = ""
        txtcodintelem = ""
        TXTKGS = ""
End Sub

Private Sub Command1_Click()
    LBLOF = OrdenFabricacionAbiertas$
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    CI1% = XVALO(TXTELEMSELEC)
    If CI1% < 1 Then
      Call COMUNI("Código No Válido")
      GoTo 99
    End If
    
    Dim CTP As New Conteplast
    PESOX# = CTP.PesoLeidoBalanza(CI1%, 2)
    If PESOX# < 0 Then PESOX# = 0
    TXTKGS = FORMATNUM$(PESOX#, "f12.2")
99  Command11.Enabled = True

End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    FECHA% = HOY(HO$)
    CMOV$ = "TR"
    DEPOGEN% = XVALO(CONTROL$("", "DEPOCEN"))
    DEPOLINER% = XVALO(CONTROL("CONTEPLA", "DEPLINER"))
    DEPOCORTE% = XVALO(CONTROL("", "DEPODEST"))
    NUMEOPERACION% = XVALO(TXTNOP)
    If DEPOLINER% < 1 Then
       Call MENSERR(24, "Falta Depósito de Liner")
       GoTo 99
    End If
    If DEPOGEN% < 1 Then
       Call MENSERR(24, "Falta Depósito General")
       GoTo 99
    End If
    If XVALO(TXTELEMSELEC) < 1 Then
       Call MENSERR(24, "Código Elemento No Válido")
       GoTo 99
    End If
    If TRIM$(Me.TXTLOTE2) = "" Then
       Call MENSERR(24, "Lote No Válido")
       GoTo 99
    End If
    If XVALO(TXTCANTCORTE) <= 0 Then
       Call MENSERR(24, "Cantidad No Válida")
       GoTo 99
    End If
    If NUMEOPERACION% < 1 Then
       Call MENSERR(24, "Operación No Válida")
       GoTo 99
    End If
    '
    If XVALO(TXTKGS) <= 0 Then
       Call MENSERR(24, "Pesaje Informado No Válido")
       GoTo 99
    End If
    
    LOTE$ = UCase$(TXTLOTE2)
    CI1% = TXTELEMSELEC
    SL = SaldoLote(CI1%, LOTE$, DEPOGEN%)
    If SL < XVALO(TXTCANTCORTE) Then
       Call COMUNI("La Cantidad a Transferir (" & TXTCANTCORTE & ") Supera El Saldo del Lote (" & SL & ") ")
       GoTo 99
    End If
    coefi# = CoeficientePesoPorMetroSegunLote#(LOTE$)
    If coefi# = 0 Then
       Call MENSERR(24, "Falta Registrar Peso para Este Lote")
       GoTo 99
    End If

    CantdeCortes = XVALO(TXTCANTCORTE)
    NCOMP& = ProNroComprobante(CMOV$)
    DOPRI$ = "TR." + TRIM$(str$(NCOMP&))
    DOSEC$ = TRIM$(LBLOF)
    NUMEOPERACION% = XVALO(TXTNOP)
    DESCRIDEPO$ = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPOLINER%)))
    REFE$ = "Tr.de " & TRIM$(DESCRIDEPO$)
    CANTI = XVALO(CantdeCortes)
    Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOGEN%, (-1) * CANTI, , , , , , , , , , NUMEOPERACION%)   'GUARDO EL NUMERO DE OPERACION DE CADA ITEMS TRANSFERIDO
    
    DESCRIDEPO$ = ECOARCH$("TADEPOSI", Chr$(DEPOGEN%))
    REFE$ = "Tr.A " & TRIM$(DESCRIDEPO$)
    NCOMP& = ProNroComprobante(CMOV$)

    Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOLINER%, CANTI, , , , , , , , , , NUMEOPERACION%)   'GUARDO EL NUMERO DE OPERACION DE CADA ITEMS TRANSFERIDO
    Call ACSALOTE(CI1%, LOTE$, "NOMESS")

    Dim CTP As New Conteplast
    '
    NORFA$ = TRIM$(LBLOF)
    CIFAB% = CODIFAB%(NORFA$)
    NUMEOPERARIO% = XVALO(TXTNUMOPERARIO)
    CORTEENMETROS = 1
    PESOKG = XVALO(TXTKGS)
    Call CTP.GrabarOperacion(NORFA$, CODIEMPR%, CIFAB%, NUMEOPERACION%, CantdeCortes, LOTE$, CI1%, XVALO(TXTCORTXBSON), NUMEOPERARIO%, CORTEENMETROS, PESOKG, RET_NEWORDEN%, Ret_NuevoIDGenerado&)
    
    
'    Call CTP.CargaMatTranfXOrden(MSF1, LBLOF, DEPOGEN%, NUMEOPERACION%)
'    Call CTP.TransfIngresada(MSF1, DEPOLINER%)
    
    familia$ = CONTROL("CONTEPLA", "FAMILINE")
    Call CTP.LinerIngresadaEnDeposito(MSF0, DEPOGEN%, familia$)
    Call CTP.CargaOperacionesInformadas(NORFA$, CIFAB%, MSF2, 0, "", 99)
    LBLOF = ""
    Set CTP = Nothing
    Call BLANFORMU2
    Screen.MousePointer = 1
    Command2.Enabled = True
    
   

    CIELEM% = CI1% 'material
    CANTI = CantdeCortes 'cantidad de cortes)
    NOP% = NUMEOPERACION%
    NROOF$ = NORFA$
    NROPED& = NroPedidoSegunOf&(NORFA$)
    NC = cliente_SEGUN_PEDIDO(NROPED&)
    ORDENINFO = RET_NEWORDEN%
    UBICACION$ = TRIM$(MSF2.TextMatrix(REG&, 10))
    IDUBICACION = XVALO(MSF2.TextMatrix(REG&, 12))
     DESCRIOPERACION$ = MSF2.TextMatrix(REG&, 7)
    IDINFOFAB& = XVALO(MSF2.TextMatrix(REG&, 16))
    CODIGOINTFAB = CIFAB%
     Call INFOCORTE.PRINTETIQPALLET(CODEXT$(CIELEM%), LOTE$, CANTI, NOP%, NC, NROPED&, NROOF$, ORDENINFO, UBICACION$, IDUBICACION, DESCRIOPERACION$, CODIGOINTFAB, IDINFOFAB&)
     '+++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++
     'BUSCO EL REGISTRO QUE SE CREO segun CAINFANT%
     CODIGOEXT$ = CODEXT$(CI1%)
     For i& = 1 To MSF2.Rows - 1
        If XVALO(MSF2.TextMatrix(i&, 11)) = RET_NEWORDEN% Then
          If MSF2.TextMatrix(i&, 9) = TRIM$(CODIGOEXT$) Then
            If TRIM$(MSF2.TextMatrix(i&, 7)) = TRIM$(DESCRIOP) Then
               REG& = i&
               Exit For
            End If
          End If
        End If
     Next i&
     
     CIELEM% = CODINT(MSF2.TextMatrix(REG&, 9)) 'material
     LOTE$ = TRIM$(MSF2.TextMatrix(REG&, 3))
     CANTI = XVALO(MSF2.TextMatrix(REG&, 4)) 'cantidad de cortes)
     NOP% = XVALO(MSF2.TextMatrix(REG&, 6))
     NROPED& = NroPedidoSegunOf&(NORFA$)
     NC = cliente_SEGUN_PEDIDO(NROPED&)
     NROOF$ = NORFA$
     ORDENINFO = XVALO(MSF2.TextMatrix(REG&, 11))
     UBICACION$ = TRIM$(MSF2.TextMatrix(REG&, 10))
     IDUBICACION = XVALO(MSF2.TextMatrix(REG&, 12))
     DESCRIOPERACION$ = MSF2.TextMatrix(REG&, 7)
     IDINFOFAB& = XVALO(MSF2.TextMatrix(REG&, 16))
     CODIGOINTFAB = CIFAB%
     Call INFOCORTE.PRINTETIQPALLET(CODEXT$(CIELEM%), LOTE$, CANTI, NOP%, NC, NROPED&, NROOF$, ORDENINFO, UBICACION$, IDUBICACION, DESCRIOPERACION$, CODIGOINTFAB, IDINFOFAB&)
     '++++++++++++++++++++++++++++++++++++++++++++++
    Call BLANFORMU2
99 Screen.MousePointer = 1
   Command2.Enabled = True
   
End Sub
Sub BLANFORMU2()
    'LBLOF = ""
    TXTCANPROY = ""
    LBLDESCRIOF = ""
    TXTELEMSELEC = ""
    TXTNOP = ""
    DESCRIOP = ""
    LBLNOMBREOPERARIO = ""
    TXTNUMOPERARIO = ""
    TXTNECESIDADCORTES = ""
    TXTCORTESINFO = ""
    TXTRESTAN = ""
    TXTCANTCORTE = ""
    TXTCODINT = ""
    TXTKGS = ""
    TXTCORTXBSON = ""
    TXTCORTE = ""
End Sub
Private Sub Command4_Click()
   IDLOT$ = TRIM$(TXTLOTE.TEXT)
   If IDLOT$ <> "" Then
     On Error Resume Next
     CONSUMLOT07.Label4 = TXTCODIGOELEM
     If Err Then Exit Sub
     On Error GoTo 0
     CONSUMLOT07.Label1 = IDLOT$
     CONSUMLOT07.Show 1
   End If
End Sub

Private Sub Command40_Click()
  Call SELECHO("PADRON")
  VDEV$ = VALACT1$("PADRON")
  If VDEV$ <> "" Then
    Me.TXTNUMOPERARIO = TRIM$(VDEV$)
    Me.LBLNOMBREOPERARIO = TRIM$(VALACT2$("PADRON"))
  End If
End Sub


Private Sub Form_Load()
    Campos$ = "Código/A14;Descripción/A26;Lote/A12;Saldo./f10.2;CI/a0;DOPRILAR/A0"
    Call CARGA_ENCABEZADO(Me.MSF0, Campos$, Me)
    DEPOGEN% = XVALO(CONTROL$("", "DEPOCEN"))
    DEPOLINER% = XVALO(CONTROL("CONTEPLA", "DEPLINER"))
    If DEPOLINER% < 1 Then
       Call COMUNI("Debe Configurar Depósito de Reservas Asignadas a O.F")
       Exit Sub
    End If
    
    FAMI$ = TRIM$(CONTROL("CONTEPLA", "FAMILINE"))
    If FAMI$ = "" Then
       Call COMUNI("Falta Configurar Familia Asignada a Liner")
       Exit Sub
    End If
    
    Dim CTP As New Conteplast
    familia$ = CONTROL("CONTEPLA", "FAMILINE")
    Call CTP.LinerIngresadaEnDeposito(MSF0, DEPOGEN%, familia$)
    Call CARGA_ENCABEZADO(MSF1, "CODIGO/A20;Descripción/A40;LOTE/A16;Cons.Mts./f10.1;CI/f6;Op.N°/F6;Kg.Transf/F9.2", Me)
    Call CARGA_ENCABEZADO(MSF2, "Código/A16;Descripción/A28;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/a0;IDUbic/a0;N°Operario/F9;Nombre/A16;Peso/f8.2;IdReg/a0", Me)

End Sub

Private Sub Label13_Change()
    Label6 = TRIM$(ECOARCH$("TADEPOSI", Chr$(XVALO(Label13))))
End Sub

Private Sub Label5_DblClick(Index As Integer)
   If Me.Height >= 9600 Then
      Me.Height = 7860
   ElseIf Me.Height <= 8500 Then
      Me.Height = 9915
   End If
End Sub


Private Sub lbldepogeneral_Change()
   DEPO% = XVALO(lbldepogeneral)
   Label3 = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%)))
End Sub

Private Sub LBLOF_Change()
   NORFA$ = TRIM$(Me.LBLOF)
   If NORFA$ = "" Then MSF1.Rows = 1: Exit Sub
   LBLDESCRIOF = RESP_ZELE_VECT(2)
   CIFAB% = CODIFAB%(NORFA$)
   NombreOperacion$ = "LINER"
     
   Dim CTP As New Conteplast
   NOPER% = CTP.NumeOperXNombre%(CIFAB%, NombreOperacion$)
   Set CTP = Nothing
   If NOPER% < 1 Then
     Call COMUNI("Atención !! La Orden de Fabricación Seleccionada No Registra Operacion de Liner")
     Exit Sub
   End If
   
   TXTCANPROY = CantiProyectadaOF(NORFA$)
   TXTNOP = NOPER%
   DESCRIOP = "Liner"
   
'   DEPOCORTE% = XVALO(Control("", "DEPODEST"))
   DEPOLINER% = XVALO(CONTROL("CONTEPLA", "DEPLINER"))

   Call CTP.CargaMatTranfXOrden(Me.MSF1, NORFA$, DEPOLINER%, NOPER%)
   Call CTP.CargaOperacionesInformadas(NORFA$, CIFAB%, MSF2, 0, "", 99)

   Set CTP = Nothing
   



End Sub

Private Sub mnuaccesodirecto_Click()
   Liner_a_OF.Show 1
End Sub

Private Sub mnuDepoCinESl_Click()
    ConfigDepo.Show 1
End Sub

Private Sub MSF0_DblClick()
    REG& = MSF0.MouseRow
    COL& = MSF0.MouseCol
    If REG& = 0 Or REG& >= MSF0.Rows Then Exit Sub
    
    Call SELECCIONARFILA(MSF0, REG&)
    SaldoLoteSeleccionado = XVALO(MSF0.TextMatrix(REG&, 4))
    Me.TXTLOTE2 = MSF0.TextMatrix(REG&, 3)
    NORFA$ = OrdenFabricacionAbiertas$
    SALDO_OR_FAB = XVALO(RESP_ZELE_VECT(3))
    CANPROY = CantiProyectadaOF(NORFA$)
    If NORFA$ = "" Then Exit Sub
    CODIGOFAB% = CODIFAB%(NORFA$)
    LBLOF = NORFA$ 'ACA EN EL EVENTO CHANGE CARGA DATOS DE LA OF
    NOPER% = XVALO(TXTNOP)
    CIELEM% = XVALO(MSF0.TextMatrix(REG&, 5))
    TXTCODINT = CIELEM%
    Dim CTP As New Conteplast
    LINER_SI_NO$ = CTP.FuncionLINER_SI_NO(CODIGOFAB%)
    If LINER_SI_NO$ <> "SI" Then
        Call COMUNI("Atención Ficha Técnica del Código Seleccionado de la Orden de Fabriación Seleccionada\No Tiene Informado Que Utiliza Liner")
        Call BLANFORMU2
        LBLOF = ""
        Exit Sub
    End If
    '
    TXTELEMSELEC = CIELEM%
    CantidadATransferir = CANPROY
    If CantidadATransferir > SaldoLoteSeleccionado Then CantidadATransferir = SaldoLoteSeleccionado
    TXTCANTCORTE = CantidadATransferir
    

    Me.TXTCORTXBSON = 1
    LOTE$ = TRIM$(TXTLOTE2)
    
    TXTNECESIDADCORTES = ROUND(CANPROY, 0)
    'CANTIDAD DE CORTES DE LA OF X LOTE Y OPERACION
   'CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND IDENLOTE = '" & LOTE$ & "' AND NUMEOPER = " & NOPER%
   CONDI$ = "IDSUBOR = '" & NORFA$ & "'  AND NUMEOPER = " & NOPER% ' LW QUITE CONDI DEL LOTE PARA QUE MUESTRE _
   SI LO TRANSFERDIO ASI SEA DE OTRO LOTE.
   Me.TXTCORTESINFO = XVALO(VALOBADA("INFOFAB", "Sum(CANTREAL)", CONDI$, "NOMESS"))
   'CANTIDAD DE CORTES DE LA OF X  OPERACION
   CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND NUMEOPER = " & NOPER%
   TXTRESTAN = FORMATNUM$(XVALO(TXTNECESIDADCORTES) - XVALO(TXTCORTESINFO), "F8")
   CANTIPROPUESTA = SALDO_OR_FAB
'   If CANTIPROPUESTA > XVALO(TXTRESTAN) Then CANTIPROPUESTA = ROUND(XVALO(TXTRESTAN), 0)
'   Me.TXTCANTCORTE = CANTIPROPUESTA
End Sub


Private Sub MSF0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    REG& = MSF0.MouseRow
    COL& = MSF0.MouseCol
    If Button = 2 Then
        Dim CTP As New Conteplast
    
        If REG& = 0 Or REG& >= MSF0.Rows Then Exit Sub
        OPX% = COMALTER%("Confirmar ELIMINACIÓN  de Registro de Modo Permanente\\Cancelar\Eliminar Registro Seleccionado")
        If OPX% < 2 Then Exit Sub
        If OPX% = 2 Then
           DEPOLINER% = XVALO(CONTROL("CONTEPLA", "DEPLINER"))
           LOTE$ = MSF0.TextMatrix(REG&, 3)
           PesoOriginal# = PesoOriginalSegúnLote#(LOTE$)
           CI1% = XVALO(MSF0.TextMatrix(REG&, 6))
           NROTRANSF$ = MSF0.TextMatrix(REG&, 7)
           CONDI$ = " IDENLOTE = '" & LOTE$ & "'   AND CODIMOVI = 'TR'  AND DOPRILAR <> '" & NROTRANSF$ & "'"
           If CantRegistros&("MOVISTO", CONDI$, "NOMESS") < 1 Then
             CONDI$ = "CODIMOVI = 'TR' AND DOPRILAR = '" & NROTRANSF$ & "' AND IDENLOTE = '" & LOTE$ & "' AND COD_INTE = " & CI1%
             Call BORRAREGISTROS("MOVISTO", CONDI$, REGAF&, "NOMESS")
             Call CTP.TransfIngresada(MSF0, DEPOLINER%)
           Else
             Call COMUNI("Imosible Eliminar Este Registro\ El mismo Posee Transferencias e Infomes de Asignación a Orden de Fabricación")
             Exit Sub
           End If
        End If

    End If
End Sub

Private Sub TXTLECTURA_DblClick()
'   TXTLECTURA = IngresoManualCodigoLote$
End Sub

Private Sub TXTCANTCORTE_Change()
  TXTCORTE = TXTCANTCORTE
End Sub

Private Sub TXTNUMOPERARIO_Change()
   Me.LBLNOMBREOPERARIO = ECOARCH$("PADRON", MKI$(XVALO(TXTNUMOPERARIO)))
End Sub


