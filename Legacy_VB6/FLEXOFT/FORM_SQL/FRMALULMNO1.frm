VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FRMALUMNO1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de Alumnos"
   ClientHeight    =   6630
   ClientLeft      =   -15
   ClientTop       =   675
   ClientWidth     =   16545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   16545
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   7200
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   6570
      Left            =   8265
      TabIndex        =   29
      Top             =   0
      Width           =   8235
      Begin VB.CommandButton Command30 
         Caption         =   ".."
         Height          =   300
         Left            =   0
         TabIndex        =   47
         ToolTipText     =   "CLICK - Para Editar Observación por Item"
         Top             =   0
         Width           =   300
      End
      Begin VB.TextBox TXTNUMCELU 
         Appearance      =   0  'Flat
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
         Left            =   2760
         TabIndex        =   12
         Top             =   1080
         Width           =   5170
      End
      Begin VB.TextBox Text16 
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "Cursos del Alumno:"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.PictureBox Picture2 
         Height          =   3000
         Left            =   5200
         ScaleHeight     =   2940
         ScaleWidth      =   2640
         TabIndex        =   41
         Top             =   2640
         Width           =   2700
         Begin VB.Image Image1 
            Height          =   3000
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2700
         End
      End
      Begin VB.CommandButton Command9 
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   360
         Picture         =   "FRMALULMNO1.frx":0000
         TabIndex        =   17
         Top             =   5880
         Width           =   540
      End
      Begin VB.TextBox TXTRUTAFIRMA 
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
         Height          =   600
         IMEMode         =   3  'DISABLE
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   5880
         Width           =   6855
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   780
         Sorted          =   -1  'True
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   360
         Width           =   1320
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
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
         Left            =   2130
         TabIndex        =   9
         Top             =   360
         Width           =   5800
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
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
         Left            =   2130
         TabIndex        =   10
         Top             =   720
         Width           =   5800
      End
      Begin VB.TextBox TXTPREFIJO 
         Appearance      =   0  'Flat
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
         Left            =   2130
         TabIndex        =   11
         Top             =   1080
         Width           =   525
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
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
         Left            =   2130
         TabIndex        =   14
         Top             =   1800
         Width           =   5800
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
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
         Left            =   2130
         TabIndex        =   13
         Top             =   1440
         Width           =   5800
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "Télefono:"
         Top             =   720
         Width           =   1815
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
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Doc.:"
         Top             =   360
         Width           =   555
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Celular:"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text16 
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "Contacto:"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox Text17 
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "E-Mail:"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
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
         Left            =   2130
         TabIndex        =   15
         Top             =   2160
         Width           =   5800
      End
      Begin VB.TextBox Text16 
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "Celular de Contacto:"
         Top             =   2160
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   2550
         Left            =   240
         TabIndex        =   42
         Top             =   3165
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   4498
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         GridLines       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
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
      Begin ACTIVESKINLibCtl.SkinLabel SKCOMENT 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FRMALULMNO1.frx":030A
         TabIndex        =   46
         Top             =   2920
         Visible         =   0   'False
         Width           =   4635
      End
      Begin VB.Label Label1 
         Caption         =   "Texto en Rojo Curso Des-Aplicado "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   5680
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6600
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   8000
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "Guardar "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3720
         TabIndex        =   38
         Top             =   5880
         Width           =   3210
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   240
         TabIndex        =   37
         Top             =   5880
         Width           =   1050
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Caption         =   "Baja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2400
         TabIndex        =   36
         Top             =   5880
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox TEXT13 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2850
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   2880
         Width           =   7530
      End
      Begin VB.TextBox TXTDESCRIPROOVINCIA 
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1800
         Width           =   4445
      End
      Begin VB.CommandButton Command5 
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
         Height          =   315
         Left            =   3070
         TabIndex        =   5
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox TXTCODPROVINCIA 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   4
         Top             =   1800
         Width           =   795
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
         Index           =   2
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Provinicia"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   4920
         TabIndex        =   0
         Top             =   240
         Width           =   2475
      End
      Begin VB.CommandButton Command1 
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
         Height          =   315
         Left            =   7480
         TabIndex        =   26
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text17 
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
         Index           =   6
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Legajo:"
         Top             =   240
         Width           =   1605
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
         Index           =   6
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "Sexo:"
         Top             =   2160
         Width           =   795
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
         Index           =   5
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Fecha de Nacimiento:"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text17 
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
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Nombres y Apellido:"
         Top             =   690
         Width           =   1815
      End
      Begin VB.TextBox Text16 
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
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Domicilio:"
         Top             =   1050
         Width           =   1815
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
         Index           =   0
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Localidad:"
         Top             =   1410
         Width           =   1815
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   1
         Top             =   690
         Width           =   5500
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   2
         Top             =   1050
         Width           =   5500
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   3
         Top             =   1410
         Width           =   5500
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   5865
         Sorted          =   -1  'True
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   2160
         Width           =   1960
      End
      Begin ACTIVESKINLibCtl.SkinLabel SKALTA 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FRMALULMNO1.frx":03CA
         TabIndex        =   24
         Top             =   300
         Visible         =   0   'False
         Width           =   1995
      End
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   330
         Left            =   2220
         TabIndex        =   6
         Top             =   2160
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   131596289
         CurrentDate     =   41123
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "FRMALULMNO1.frx":0444
         TabIndex        =   39
         Top             =   2520
         Width           =   1935
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   10815
      OleObjectBlob   =   "FRMALULMNO1.frx":04BC
      Top             =   105
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Menu ABM 
      Caption         =   "ABM"
      Begin VB.Menu MNUALTA 
         Caption         =   "ALTA"
         Shortcut        =   ^A
      End
      Begin VB.Menu MNUBAJA 
         Caption         =   "BAJA"
         Shortcut        =   ^B
      End
      Begin VB.Menu MNUREST 
         Caption         =   "RESTAURAR BAJA"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuAccesos 
      Caption         =   "Accesos "
      Begin VB.Menu mnuAsignarCurso 
         Caption         =   "Asignar Curso a Alumno Activo"
      End
      Begin VB.Menu mnuTomarLista 
         Caption         =   "Tomar lista"
      End
   End
End
Attribute VB_Name = "FRMALUMNO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CARGA_DATOS_ALUMNMOS()

   LEGAJO& = XVALO(Text1)
   If LEGAJO& <= 0 Then Exit Sub
   CONDI$ = "NUME_CLI = " & LEGAJO&
   If CantRegistros&("DEUDOR12", "NUME_CLI = " & LEGAJO&, "NOMESS") < 1 Then Exit Sub
   '
   SQLX$ = "SELECT * FROM DEUDOR12 WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   Set ALUMTMP = FLEXCONN.Execute(SQLX$)
   With ALUMTMP
        Text14 = SAFETEXT$(!RASO_CLI)
        Text2 = SAFETEXT$(!DOMI_CLI)
        Text3 = SAFETEXT$(!LOCA_CLI)
'        Combo1(0).Text = SAFETEXT$(!PROV_CLI) VER COMO ASIGNARLE LOS DATOS A LOS COMBOS
        Combo1(1).Text = SAFETEXT$(!TIPODOCU_ALUM)
        Combo1(2).Text = SAFETEXT$(!SEXO_CLI)
        Text12.Text = SAFETEXT$(!DOCU_CLI)
        Text11.Text = SAFETEXT$(!Tele_Cli)
        TXTPREFIJO.Text = SAFETEXT$(!PrefCel_CLI)
        TXTNUMCELU.Text = SAFETEXT$(!NumCel_CLI)
'        Text10.Text = SAFETEXT$(!CELU_CLI)
        Text8.Text = SAFETEXT$(!MAIL_CLI)
        Text9.Text = SAFETEXT$(!CTAC_CLI)
        Text18.Text = SAFETEXT$(!CELCONTAC_CLI)
        TEXT13 = SAFETEXT$(!nota_Cli)
        Me.TXTRUTAFIRMA = SAFETEXT$(!IMAG_CLI)
        Me.TXTCODPROVINCIA = SAFETEXT$(!PROV_CLI)
        Me.DTP1.Value = CVDAT(CVINT(!FECHNAC))
      
    End With
    ALUMTMP.Close
    Set ALUMTMP = Nothing
    '
    'CARGO LOS CURSOS ASOCIADOS DEL ALUMNO
    Campox$ = "Id/F5;Fecha/D8;Descripción/A48;CI/A0"
    Call CARGA_ENCABEZADO(Me.MSF, Campox$, Me)

    SQLX$ = "SELECT CURSOS_ID,FECHA,DESCRIPCION ,STATUS"
    SQLX$ = SQLX$ + " FROM ALUMNO_CURSO  WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE NUME_CLI = " & LEGAJO&
    
    
    SQLX$ = SQLX$ + " ORDER BY FECHA DESC"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        ST% = XVALO(!Status)
        MSF.Font.Underline = False
        j& = j& + 1
        MSF.Rows = j& + 1
        MSF.TextMatrix(j&, 1) = XVALO(!CURSOS_ID)
        MSF.TextMatrix(j&, 2) = FECHATEX$(CVINT(!FECHA))
        MSF.TextMatrix(j&, 3) = SAFETEXT$(!Descripcion)
        
        If ST% < 0 Then
            MSF.COL = 3
            MSF.Row = j&
            MSF.TextMatrix(j&, 0) = "*"
            Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbRed, j&)
        End If
        .MoveNext
      Loop
    End With



End Sub




Private Sub Combo2_Change()
   TXCOMBO2 = Combo2.Text
End Sub

Private Sub Combo2_Click()
   TXCOMBO2 = Combo2.Text
End Sub

Private Sub Combo3_Change()
   TXCOMBO3 = Combo3.Text
End Sub

Private Sub Combo3_Click()
   TXCOMBO3 = Combo3.Text
End Sub

Private Sub Combo5_Change()
  TXCOMBO5 = Combo5.Text
End Sub

Private Sub Combo5_Click()
  TXCOMBO5 = Combo5.Text
End Sub


Private Sub Combo6_Change()
   TXCOMBO6 = Combo6.Text
End Sub

Private Sub Combo6_Click()
   TXCOMBO6 = Combo6.Text
End Sub

Private Sub Combo7_Change()
   TXCOMBO7 = Combo7.Text
End Sub

Private Sub Combo7_Click()
   TXCOMBO7 = Combo7.Text
End Sub

Private Sub Combo8_Change()
   TXCOMBO8 = Combo8.Text
End Sub

Private Sub Combo8_Click()
   TXCOMBO8 = Combo8.Text
End Sub


Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If
End Sub


Private Sub Command1_Click()
    Command1.Enabled = False
    '
    If SKALTA.Visible = True Then
      TX$ = "PARA SELECCIONAR ALUMNO EXISTENTE" + vbCrLf
      TX$ = TX$ + "DEBE VOLVER A MODO EDICIÓN (Presionando el Botón 'Edición')"
      MsgBox TX$
      GoTo 99
    End If
    '
    Call ABRECONEXION
    '
    If ALUMONOSEL > 0 Then
      Text1 = XVALO(RESP_ZELE_VECT(1))
    End If
    '
99  Command1.Enabled = True
    '
    
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   'ALTAX
   
   If Command2.Caption = "Nuevo" Then
     Command2.Caption = "Edición"
     SKALTA.Visible = True
     NUMPRO& = 1 + MAXNUMCAMPO&("DEUDOR12", "NUME_CLI")
     Text1 = NUMPRO&
     Me.TXTPREFIJO = "011"
   Else
     Command2.Caption = "Nuevo"
     SKALTA.Visible = False
     Text1 = ""
   End If

99 Command2.Enabled = True

End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    'VALIDAR SI HAY UN SOLO GRADO O AÑO O SALA DE
    'VALIDAR LA CUOTA.
    Call ABRECONEXION
    JJ& = 0
    '
    'TOMO EL NUMERO DEL LABEL
    LEGAJO& = XVALO(Text1)
    '
    'VALIDACION QUE SE HAYA SELECCIONADO UN CODIGO
    If LEGAJO& < 1 Then
       MsgBox "LEGAJO NO VALIDO"
       GoTo 99
    End If
    '
    If TRIM(Text14) = "" Then
       MsgBox "Debe Ingresar Nombre"
       GoTo 99
    End If
    '
    If TRIM(Me.Combo1(2)) = "" Then
       MsgBox "Falta el Sexo del Alumno"
       GoTo 99
    Else
      If Combo1(2).Text <> "Femenino" And Combo1(2) <> "Masculino" Then
         MsgBox "Sexo Seleccionado No Válido"
         GoTo 99
      End If
    End If
    '
    If TRIM(Me.Combo1(1)) = "" Then
       MsgBox "Tipo de Documento No Válido"
       GoTo 99
    Else
      If TRIM$(Combo1(1).Text) <> "DNI" And TRIM$(Combo1(1)) <> "CÉDULA" And TRIM$(Combo1(1)) <> "Pasaporte" And TRIM$(Combo1(1)) <> "Otro" Then
         MsgBox "Tipo de Documento Seleccionado No Válido"
         GoTo 99
      End If
    End If
    
    If TRIM$(TXTDESCRIPROOVINCIA) = "" Then
         MsgBox "Ingreso de Provincia Seleccionada No Válido"
         GoTo 99
    End If
    '
    'VALIDAR SI EL DOCUMENTO ES EXISTENTE CON MISMO NUMERO EN OTRO ALUMNO
    If TRIM$(Text12) <> "" Then
        CONDI$ = "DOCU_CLI =  '" & TRIM$(Text12) & "' AND TIPODOCU_ALUM = '" & UCase$(TRIM$(Combo1(1).Text)) & "' AND NUME_CLI <> " & XVALO(Text1)
        VALIDA_LEGAJO& = XVALO(CantRegistros("DEUDOR12", CONDI$, "NOMESS"))
        If VALIDA_LEGAJO& > 0 Then
           NOMYAPE$ = TRIM$(VALOBADA("DEUDOR12", "RASO_CLI", CONDI$, "NOMESS"))
           MsgBox "DOCUMENTO YA EXISTENTE " + vbCrLf + " ALUMNO: " & NOMAPE$ & " LEGAJO: " & VALIDA_LEGAJO& & vbCrLf + "FAVOR DE CAMBIAR Y VOLVER A GRABAR"
           GoTo 99
        End If
        If Len(TRIM$(Text12)) < 7 Then
           OPX% = COMALTER%("El Número Ingresado Tiene Menos de 7 Dígitos\\Cancelar\Grabar Igual")
           If OPX% < 2 Then GoTo 99
        End If
    Else
        MsgBox ("Falta Nro. De Documento")
        GoTo 99
    End If
    '
    If TRIM$(Text8) <> "" Then
       If InStr(Text8, "@") < 1 Then
          Call COMUNI("Atención !!! Mail Inválido")
          GoTo 99
       End If
    End If
    '
    FF% = HOY(HO$)
    FECHA1 = FETEXCOR1$(FF%)
    '
    'ALTAX
    If SKALTA.Visible = True Then
      LEGAJO& = 1 + MAXNUMCAMPO&("DEUDOR12", "NUME_CLI")
    End If
    '
    GRABO% = 0
    '
    'GRABACION , SI NO EXISTE LO GENERA SI NO LO EDITA
    SQLX$ = "SELECT * FROM DEUDOR12"
    SQLX$ = SQLX$ + " WHERE NUME_CLI = " & LEGAJO&
    '
    On Error GoTo ERROR_GUARDAR
    Set AUX = New ADODB.Recordset
    AUX.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    FLEXCONN.BeginTrans
    With AUX
        If (.BOF And .EOF) Then
          .AddNew
          !NUME_CLI = LEGAJO&
          !ALTA_CLI = FECHA1
        Else
          !NUME_CLI = LEGAJO&
        End If
        
        !RASO_CLI = Left$(Text14, 64)
        !DOMI_CLI = Left$(Text2, 64)
        !LOCA_CLI = Left$(Text3, 64)
        '
        !TIPODOCU_ALUM = Left$(Combo1(1).Text, 16)
        !DOCU_CLI = Left$(Text12.Text, 32) 'TEXTO
        !Tele_Cli = Left$(Text11.Text, 64) 'TEXTO
        !PrefCel_CLI = Left(TRIM$(Me.TXTPREFIJO.Text), 24)
        !NumCel_CLI = Left(TRIM$(Me.TXTNUMCELU.Text), 24)
        !CELU_CLI = Left$(TRIM$(Me.TXTPREFIJO.Text & Me.TXTNUMCELU.Text), 128)
        
        !MAIL_CLI = Left$(Text8.Text, 64) 'TEXTO
        '
        !CTAC_CLI = Left$(Text9.Text, 64) 'TEX
        !CELCONTAC_CLI = Left$(Text18.Text, 64) 'TX
        !PROV_CLI = Me.TXTCODPROVINCIA
        !SEXO_CLI = Left$(Combo1(2).Text, 18)
        !IMAG_CLI = Left$(Me.TXTRUTAFIRMA.Text, 512)
        '
        !STAT_CLI = 1
        !FECHNAC = CVDAT(CVINT(Me.DTP1.Value))
        '
        !nota_Cli = TEXT13
      .Update
      '
    End With
    
ERROR_GUARDAR:
    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      GoTo 99
    Else
      FLEXCONN.CommitTrans
      AUX.Close
      Set AUX = Nothing
      GRABO% = 1
    End If
    '
    If SKALTA.Visible = True Then
       TXT$ = "Alta de Alumno"
       Call ANOTAX_ALUMNO(LEGAJO&, TXT$)
    End If

    SKALTA.Visible = False
    'REFRESCO PARA EL CASO QUE SE DIO ALTA Y EL NUMERO QUE GRABO ES <> AL QUE QUEDO EN LABEL
    Text1 = LEGAJO&
    Command2.Caption = "Nuevo"
    
    MsgBox "HECHO"
99  Command3.Enabled = True

End Sub

Private Sub Command4_Click()
    Command4.Enabled = False

    Call ABRECONEXION
    JJ& = 0
    '
    'TOMO EL NUMERO DEL LABEL
    LEGAJO& = XVALO(Text1)
    '
    'VALIDACION QUE SE HAYA SELECCIONADO UN CODIGO
    If LEGAJO& < 1 Then
       MsgBox "LEGAJO No Valido"
       GoTo 99
    End If
    '
    CONDI$ = "NUME_CLI = " & LEGAJO&
    '
    'VALIDO SI EXISTE
    CNRG& = CantRegistros&("DEUDOR12", CONDI$)
    If CNRG& < 1 Then
      MsgBox "Legajo No Valido"
      GoTo 99
    End If
    '
    TTXX$ = "Realmente desea dar de Baja el Alumno Activo"
    If MsgBox(TTXX$, vbYesNo, "Baja de Alumno") <> vbYes Then GoTo 99
    'GRABACION , SI NO EXISTE LO GENERA SI NO LO EDITA
    SQLX$ = "SELECT * FROM deudor12 "
    SQLX$ = SQLX$ + " WHERE nume_cli = " & LEGAJO&
    '
    Dim RECTMP As ADODB.Recordset
    Set RECTMP = New ADODB.Recordset
    RECTMP.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

    With RECTMP
        !STAT_CLI = -1
        .Update
    End With
    '
    RECTMP.Close
    Set RECTMP = Nothing
    '
    TXT$ = "Baja de Alumno"
    Call ANOTAX_ALUMNO(LEGAJO&, TXT$)
    '
    Call LIMPIAR
    Text1 = ""
    
99  Command4.Enabled = True

End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
    Call ABRECONEXION
    '
    If PROVINSEL% > 0 Then
      Me.TXTCODPROVINCIA = RESP_ZELE_VECT(2)
    End If
   
   
99 Command5.Enabled = True

End Sub

Private Sub Command9_Click()
   RUTAX$ = SELECCION_Archivo$("*", CMD1)
   If RUTAX$ <> "" Then
        On Error Resume Next
        Image1.Picture = LoadPicture(RUTAX$)
        If Err > 0 Then MsgBox Err.Description
   Else
      GoTo 99
   End If
   TXTRUTAFIRMA = RUTAX$
99 Screen.MousePointer = 1

End Sub

Private Sub DTP1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If
End Sub


Private Sub Form_Load()
    ''
    Call ABRECONEXION
    Call ABRECLIEN
    Call CREACAMPO("BADACLIEN", "Deudor12", "TIPODOCU_ALUM", 10, 24)
    Call CREACAMPO("BADACLIEN", "Deudor12", "DOCU_CLI", 10, 24)
    Call CREACAMPO("BADACLIEN", "Deudor12", "CELCONTAC_CLI", 10, 48)
    Call CREACAMPO("BADACLIEN", "Deudor12", "ALTA_CLI", 8, 0)
    Call CREACAMPO("BADACLIEN", "Deudor12", "SEXO_CLI", 10, 18)
    Call CREACAMPO("BADACLIEN", "Deudor12", "IMAG_CLI", 10, 512)
    Call CREACAMPO("BADACLIEN", "Deudor12", "FECHNAC", 8, 0)
    Call CREACAMPO("BADACLIEN", "Deudor12", "PrefCel_CLI", 10, 24)
    Call CREACAMPO("BADACLIEN", "Deudor12", "NumCel_CLI", 10, 24)
    
    
    
    Campox$ = "Id/F5;Fecha/D8;Descripción/A48"
    Call CARGA_ENCABEZADO(Me.MSF, Campox$, Me)
    Call APLICAR_SKIN1(Me, App.Path & "\SKINS\dogmax.skn")
    
''  'CARGA COMBO DE TIPO DE DOCUMENTO
    Combo1(1).AddItem ""
    Combo1(1).AddItem "DNI"
    Combo1(1).AddItem "CÉDULA"
    Combo1(1).AddItem "Pasaporte"
    Combo1(1).AddItem "Otro"
    Combo1(1).Text = "DNI"
'   '
    Combo1(2).Text = "Masculino"
    Combo1(2).AddItem "Femenino"
    Combo1(2).AddItem "Masculino"

    DTP1.Value = Date

    
End Sub


Sub LIMPIAR()
   Text14 = ""
   Text2 = ""
   Text3 = ""
   Text8 = ""
   Text9 = ""
   Text10 = ""
   Text11 = ""
   Text12 = ""
   TEXT13 = ""
   Text18 = ""
   Text19 = ""
   Text20 = ""
   Text21 = ""
   Text22 = ""
   Text23 = ""
   Text24 = ""
   Text29 = ""
   Text28 = ""
   Text26 = ""
   Text29 = ""
   Text27 = ""
   Text38 = ""
   Text31 = ""
   Text32 = ""
   Text33 = ""
   Text25 = ""
   Me.TXTRUTAFIRMA = ""
   TXTCODPROVINCIA = ""
   Me.TXTDESCRIPROOVINCIA = ""
   Command30.Visible = False 'BOTON DE OBSERVACIONES POR ITEM


End Sub
Private Sub Command30_Click()
    FilaActiva = MSF.Row
    MSF.COL = j&
    '
    CI1 = XVALO(MSF.TextMatrix(FilaActiva, 1))
    LGJO& = XVALO(Me.Text1)
    If LGJO& < 1 Or CI1 < 1 Then Exit Sub
    CURSOX$ = MSF.TextMatrix(FilaActiva, 3)
    '
    'TRAE TODOS LOS CURSOS QUE ESTEN VIGENTES. (QUE LA FECHA DE FINALIZACION NO HAYA SUPERADO LA FECHA DE HOY)
    SQLX$ = " select MESTEX,FECHLIS, DIAATEX,PRESTE from presentes WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE NUME_CLI = " & LGJO&
    SQLX$ = SQLX$ + " AND CODINT = " & CI1
    SQLX$ = SQLX$ + " AND PRESTE > 0"
    SQLX$ = SQLX$ + " order by FECHLIS"
    CAMPOS$ = "MES/A20;FECHA/D8;DÍA/A16;PRESENTE/A12"
    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, CAMPOS$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, CAMPOS$, GRILLAGRAL, 1)
    
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    GRILLAGRAL.Height = Screen.Height - 2000
    GRILLAGRAL.Width = Screen.Width - 1000
    GRILLAGRAL.NOPINTAR = 1
    '
    'CARGO EL RECORDSET
    j& = 0
    Set RS1 = READSET(SQLX$)
    With RS1
      Do While Not .EOF
         j& = j& + 1
         GRILLAGRAL.GRID.Rows = j& + 1
         GRILLAGRAL.GRID.TextMatrix(j&, 1) = SAFETEXT(!MESTEX)
         GRILLAGRAL.GRID.TextMatrix(j&, 2) = CVDAT(CVINT(!FECHLIS))
         GRILLAGRAL.GRID.TextMatrix(j&, 3) = SAFETEXT(!DIAATEX)
         GRILLAGRAL.GRID.TextMatrix(j&, 4) = "PRESENTE"
      .MoveNext
      Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    '*******************************************
    Screen.MousePointer = 1
    If GRILLAGRAL.GRID.Rows < 2 Then
      Call COMUNI("NO SE DETECTO ASISTENCIA PARA EL ALUMNO ACTIVO Y CURSO SELECCIONADO")
      Exit Sub
    End If
    
    TX$ = "LISTADO DE ASISTENCIA (ALUMNO: " & Text14.Text & " CURSO: " & CURSOX$
    GRILLAGRAL.Caption = TX$
    '*******************************************
30  GRILLAGRAL.Show 1
    If GRILLAGRAL.MENU1 > 0 Then
        Dim MSFX As MSFlexGrid
        Set MSFX = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSFX = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSFX, CAMPOS$, "PRESENTES", "ASISTENCIA", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
    End If
    
End Sub
Private Sub MNUALTA_Click()
   Call Command2_Click
End Sub

Private Sub mnuAsignarCurso_Click()
   NC = XVALO(Text1)
   If NC < 1 Then
     Call COMUNI("Falta Seleccionar El Alumno")
     Exit Sub
   End If
   ASIGCURSO.Command7.Visible = False
   ASIGCURSO.Text2.Locked = True
   ASIGCURSO.Text2.BackColor = &HE0E0E0
   ASIGCURSO.Text2 = Text1
   ASIGCURSO.Height = 8280
   ASIGCURSO.Show 1
   
   Text1 = ""
   Text1 = NC
End Sub

Private Sub MNUBAJA_Click()
   Call Command4_Click
End Sub

Private Sub MNUREST_Click()
    If ALUMONOSEL(1) > 0 Then
      LEGAJO& = XVALO(RESP_ZELE_VECT(1))
    Else
      GoTo 99
    End If
    '
    
    SQLX$ = "SELECT * FROM DEUDOR12 "
    SQLX$ = SQLX$ + " WHERE NUME_CLI = " & LEGAJO&
    '
    Dim RECTMP As ADODB.Recordset
    Set RECTMP = New ADODB.Recordset
    RECTMP.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RECTMP
        !STAT_CLI = 0
        NOMBREYAPE$ = SAFETEXT$(!RASO_CLI)
        .Update
    End With
    '
    RECTMP.Close
    Set RECTMP = Nothing
    '
    MsgBox "ALUMNO: " & NOMBREYAPE$ & " RESTAURADO"
    TXT$ = "Restauración de Alumno"
    Call ANOTAX_ALUMNO(LEGAJO&, TXT$)
    '
    Call LIMPIAR
    Text1 = SAFETEXT$(LEGAJO&)
99
End Sub

Private Sub mnuTomarLista_Click()
   FRMPRESENTE.Show 1
End Sub

Private Sub MSF_Click()
   COL& = MSF.MouseCol
   FilaActiva = MSF.Row
   If COL& = 0 Then
        With MSF
            .COL = COL&
            Command30.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
            Command30.Visible = True
            Command30.ZOrder 0
        End With
   End If

End Sub
Private Sub MSF_RowColChange()
    FILA& = MSF.Row
    On Error Resume Next
    Text28 = MSF.TextMatrix(FILA&, 9)
    Command30.Visible = False
    On Error GoTo 0
End Sub

Private Sub MSF_DblClick()
    If MSF.Rows <= 1 Then Exit Sub
    REG& = MSF.MouseRow
    If REG& < 1 Then Exit Sub
    IDCUR& = MSF.TextMatrix(REG&, 1)
    NC = XVALO(Text1)
    If IDCUR& > 0 Then
      If MSF.TextMatrix(REG&, 0) = "*" Then
        If DERACCE%("ALTACURS", "ACTIVAR ESTE ALUMNO A CURSO") < 2 Then Exit Sub
        OPX% = COMALTER%("Desea Dar de Alta a Este Alumno al Curso Seleccionado\\Cancelar\Dar de Alta del Curso")
        If OPX% > 1 Then
           CONDI$ = "NUME_CLI = " & XVALO(Text1) & " and CURSOS_ID = " & IDCUR&
           Call ACTUREGISTRO("ALUMNO_CURSO", "STATUS", CONDI$, 1, REGAF&, "NOMES")
           Text1 = "": Text1 = NC ' PARA QUE REFRESQUE
        End If
      Else
        If DERACCE%("BAJACURS", "DAR DE BAJA UN ALUMNO DE UN CURSO") < 2 Then Exit Sub
        OPX% = COMALTER%("Desea Dar de Baja a Este Alumno al Curso Seleccionado\\Cancelar\Dar de baja del Curso")
        If OPX% > 1 Then
           CONDI$ = "NUME_CLI = " & XVALO(Text1) & " and CURSOS_ID = " & IDCUR&
           Call ACTUREGISTRO("ALUMNO_CURSO", "STATUS", CONDI$, -1, REGAF&, "NOMES")
           Text1 = "": Text1 = NC ' PARA QUE REFRESQUE
        End If
      End If
    End If

End Sub


Private Sub MSF_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, Y As Single)
  If MSF.Rows <= 1 Then Exit Sub
  
  With MSF
    ' verifica que el MsFlexgrid contiene filas
    If .MouseRow > 0 And .MouseCol >= 0 Then
        ' le asigna el dato con TextMatrix
        .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
    Else
        ' No muestra ningún ToolTip
        .ToolTipText = vbNullString
          
    End If
  End With
    
End Sub
Private Sub Text1_Change()
     Call LIMPIAR
     Call CARGA_DATOS_ALUMNMOS
End Sub
Private Sub Text1_DblClick()
   Call Command1_Click
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If
End Sub

Private Sub Text11_Change()
   Text11 = UCase(Text11)
   Text11.SelStart = Len(Text11)

End Sub

Private Sub Text13_Change()
   TEXT13 = UCase(TEXT13)
   TEXT13.SelStart = Len(TEXT13)

End Sub

Private Sub Text14_Change()
   Text14 = UCase(Text14)
   Text14.SelStart = Len(Text14)
   
End Sub

Private Sub TEXT14_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text17_DblClick(Index As Integer)
    If Index = 6 Then Command1_Click
    
End Sub


Private Sub Text18_Change()
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text2_Change()
   Text2 = UCase(Text2)
   Text2.SelStart = Len(Text2)

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text25_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text26_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text27_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text28_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text29_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text3_Change()
   Text3 = UCase(Text3)
   Text3.SelStart = Len(Text3)

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub


Private Sub Text32_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text8_Change()
   Text8 = UCase(Text8)
   Text8.SelStart = Len(Text8)

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text9_Change()
  Text9 = UCase(Text9)
  Text9.SelStart = Len(Text9)

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

'Private Sub TXCOMBO1_Change(Index As Integer)
'   TXCOMBO1(Index) = UCase(TXCOMBO1(Index))
'End Sub


Private Sub TXTCODPROVINCIA_Change()
     If TRIM$(Me.TXTCODPROVINCIA) <> "" Then
         Me.TXTDESCRIPROOVINCIA = VALOBADA("TAPROVINCIA", "DESCRI_PROV", "CODITEX= '" & Me.TXTCODPROVINCIA & "'", "NOMESS")
     Else
         Me.TXTDESCRIPROOVINCIA = ""
     End If
     TXTCODPROVINCIA = UCase(TXTCODPROVINCIA)
     TXTCODPROVINCIA.SelStart = Len(TXTCODPROVINCIA)

End Sub


Private Sub TXTRUTAFIRMA_Change()
   Image1.Picture = LoadPicture("")
   If EXISTE%(Me.TXTRUTAFIRMA) > 0 Then
     On Error Resume Next
     Image1.Picture = LoadPicture(TXTRUTAFIRMA)
     On Error GoTo 0
   Else
     
   End If

End Sub


