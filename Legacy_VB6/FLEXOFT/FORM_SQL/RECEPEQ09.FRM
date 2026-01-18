VERSION 5.00
Begin VB.Form RECEPEQ09 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepcion de Equipos para Reparación"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   495
      Left            =   8880
      TabIndex        =   60
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Siguiente >>>>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   8640
      Picture         =   "RECEPEQ09.frx":0000
      TabIndex        =   40
      ToolTipText     =   "Pantalla Siguiente"
      Top             =   5640
      Width           =   915
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "ENCABEZADO"
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
      Left            =   120
      TabIndex        =   33
      Top             =   240
      Width           =   8415
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
         Height          =   300
         Left            =   8085
         TabIndex        =   47
         Top             =   325
         Width           =   180
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Solic.Entrega"
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
         Index           =   4
         Left            =   5535
         TabIndex        =   46
         Top             =   420
         Width           =   1440
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
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
         Left            =   7050
         TabIndex        =   45
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label Label17 
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
         Left            =   1260
         TabIndex        =   39
         Top             =   735
         Width           =   3675
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
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
         Left            =   3255
         TabIndex        =   38
         Top             =   315
         Width           =   2235
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
         Left            =   7410
         TabIndex        =   37
         Top             =   720
         Width           =   825
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Recepción: Fecha y Hora"
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
         Left            =   630
         TabIndex        =   36
         Top             =   420
         Width           =   2520
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Recibió"
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
         Left            =   600
         TabIndex        =   35
         ToolTipText     =   "Doble Click para Refrescar"
         Top             =   840
         Width           =   960
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Talón"
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
         Left            =   4515
         TabIndex        =   34
         Top             =   840
         Width           =   2760
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "DATOS DEL CLIENTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   8415
      Begin VB.TextBox Text14 
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
         Height          =   300
         Left            =   6960
         TabIndex        =   59
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text13 
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
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   " "
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton Command8 
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
         Height          =   300
         Left            =   1350
         TabIndex        =   55
         ToolTipText     =   "Desplegar Lista de Selección"
         Top             =   2040
         Width           =   180
      End
      Begin VB.TextBox Text12 
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
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   54
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   2040
         Width           =   3525
      End
      Begin VB.TextBox Text10 
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
         Height          =   300
         Left            =   5520
         TabIndex        =   31
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1020
         Width           =   2775
      End
      Begin VB.TextBox Text7 
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
         Height          =   300
         Left            =   6120
         TabIndex        =   25
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   330
         Width           =   2175
      End
      Begin VB.TextBox Text8 
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
         Height          =   300
         Left            =   6120
         TabIndex        =   24
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   675
         Width           =   2175
      End
      Begin VB.TextBox Text6 
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
         Left            =   2205
         TabIndex        =   23
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1020
         Width           =   2850
      End
      Begin VB.TextBox Text5 
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
         Height          =   300
         Left            =   1365
         TabIndex        =   22
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1020
         Width           =   750
      End
      Begin VB.TextBox Text4 
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
         Height          =   300
         Left            =   1365
         TabIndex        =   21
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   675
         Width           =   3690
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
         TabIndex        =   20
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   330
         Width           =   3690
      End
      Begin VB.CommandButton Command2 
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
         Height          =   300
         Left            =   1080
         TabIndex        =   19
         Top             =   675
         Width           =   180
      End
      Begin VB.TextBox Text1 
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
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Text            =   " "
         Top             =   675
         Width           =   855
      End
      Begin VB.TextBox Text9 
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
         Height          =   300
         Left            =   1365
         TabIndex        =   17
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1380
         Width           =   3690
      End
      Begin VB.Label label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N°Talón:"
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
         Left            =   6000
         TabIndex        =   58
         Top             =   2100
         Width           =   825
      End
      Begin VB.Label label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Técnico Externo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   1920
         Width           =   795
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
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
         Left            =   5040
         TabIndex        =   32
         Top             =   1100
         Width           =   480
      End
      Begin VB.Label label2 
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
         Index           =   2
         Left            =   5100
         TabIndex        =   30
         Top             =   405
         Width           =   960
      End
      Begin VB.Label Label12 
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
         TabIndex        =   29
         ToolTipText     =   "Doble Click para Refrescar"
         Top             =   435
         Width           =   960
      End
      Begin VB.Label label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto"
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
         Left            =   240
         TabIndex        =   28
         Top             =   1455
         Width           =   1065
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
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
         Left            =   5100
         TabIndex        =   27
         Top             =   795
         Width           =   960
      End
      Begin VB.Label Label26 
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
         Left            =   210
         TabIndex        =   26
         Top             =   1020
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "IDENTIFICACIONES DEL EQUIPO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   1550
      Width           =   8415
      Begin VB.TextBox Text11 
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
         Height          =   285
         Left            =   5760
         TabIndex        =   50
         Top             =   780
         Width           =   2490
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Historial"
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
         Left            =   3840
         Picture         =   "RECEPEQ09.frx":030A
         TabIndex        =   44
         ToolTipText     =   "Muestra el Historial del Equipo Elegido"
         Top             =   780
         Width           =   960
      End
      Begin VB.TextBox Text2 
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
         Height          =   285
         Left            =   4560
         TabIndex        =   0
         Top             =   360
         Width           =   3690
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
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
         ForeColor       =   &H000000FF&
         Height          =   230
         Left            =   7800
         TabIndex        =   62
         Top             =   1560
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sector"
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
         Left            =   4680
         TabIndex        =   49
         Top             =   890
         Width           =   960
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Garantia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6850
         TabIndex        =   15
         Top             =   1900
         Width           =   870
      End
      Begin VB.Label Label7 
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
         Height          =   285
         Left            =   7800
         TabIndex        =   14
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecnico -   Ultimo Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Left            =   2805
         TabIndex        =   13
         Top             =   1700
         Width           =   1215
      End
      Begin VB.Label Label5 
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
         Height          =   285
         Left            =   4080
         TabIndex        =   12
         Top             =   1800
         Width           =   2745
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ultima Reparacion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Left            =   -70
         TabIndex        =   11
         Top             =   1700
         Width           =   1455
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   1800
         Width           =   1185
      End
      Begin VB.Label label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   980
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Serie >>>>>>>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   470
         Width           =   3960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   840
         TabIndex        =   7
         Top             =   560
         Width           =   870
      End
      Begin VB.Label Label9 
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
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   780
         Width           =   2145
      End
      Begin VB.Label Label10 
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
         Height          =   285
         Left            =   800
         TabIndex        =   5
         Top             =   1200
         Width           =   7470
      End
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   5150
      Left            =   8610
      ScaleHeight     =   5115
      ScaleWidth      =   1350
      TabIndex        =   1
      Top             =   -20
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   52
         Top             =   4500
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   53
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "RECEPEQ09.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Re-Imprimir Talóm"
         Top             =   3480
         Width           =   690
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "RECEPEQ09.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Generacion Directa de Número de Serie"
         Top             =   2570
         Width           =   690
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "RECEPEQ09.frx":0C08
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Acceso a Base de Datos de Clientes"
         Top             =   1850
         Width           =   690
      End
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "RECEPEQ09.frx":104A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   225
         Width           =   690
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "RECEPEQ09.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   950
         Width           =   690
      End
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
      TabIndex        =   61
      Top             =   0
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LBLSTATGARAN_ORIGINAL 
      Height          =   330
      Left            =   5520
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label APRO 
      Height          =   330
      Left            =   7770
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "RECEPEQ09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG07/AMACLI"
   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
   Command1.Enabled = True
End Sub

Private Sub Command11_Click()
    '
    Unload Me
    '
End Sub

Private Sub Command2_Click()
    '
    If TRIM$(Text2) = "" Then Text2.SetFocus
    '
    NCLIE = ListaSeleccionClientes("", True)
    If NCLIE < 1 Then Exit Sub
    '
    NC = NCLIE
    Text1.TEXT = TRIM$(Str$(NC))

''    Call SELECHO("DEUDOR12")
''    VSELE$ = TRIM$(VALACT1$("DEUDOR12"))
''    If VSELE$ <> "" Then
''       NCLIE = XVALO(VSELE$)
''       Text1.TEXT = TRIM$(Str$(NCLIE))
''    End If
    '
End Sub

Private Sub Command2_GotFocus()
    If TRIM$(Text2) = "" Then Text2.SetFocus
End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   If DERACCE%("CADINUSE", "Carga Directa de Numeros de Serie") >= 2 Then
       '
       MODOEYM% = 0
       EPAC$ = TRIM$(EMPREAC$)
       If InStr(EPAC$, "EYM") > 0 Then
          ALTAMANUAL_NSERIE_EYM
          GoTo 99
       End If
       
1      Opcion% = COMALTER%("Ingreso de Equipo al Sistema\\Generación Automática de Número de Serie\Re-impresión de Etiqueta")
       If Opcion% < 1 Or Opcion% > 2 Then GoTo 99
       If Opcion% = 1 Then GoTo 3
       
       Call SELECHO("INDINUSE")
       NUMEROSERIE$ = TRIM$(VALACT1$("INDINUSE"))
       If NUMEROSERIE$ = "" Then GoTo 1
       CodigoInterno% = XVALO(VALOBADA("TANUMSE", "COD_INTE", "NUMEROSERIE='" & NUMEROSERIE$ & "'"))
       FechaAlta% = CVINT(VALOBADA("TANUMSE", "FECHALTA", "NUMEROSERIE='" & NUMEROSERIE$ & "'"))
       Call PRINTALBANUMSE(NUMEROSERIE$, CodigoInterno%, FECHATEX$(FechaAlta%), 1)
       GoTo 99
       
       
3      Call SELECHO("MASTER")
       Codigo$ = VALACT1("MASTER")
       CI% = CODINT(Codigo$)
       If CI% < 1 Then GoTo 1
       NUSERIE$ = NUEVONUMEROSERIE$(Codigo$)
5      VDEF$ = String$(48, 0)

       Call REPLA$(VDEF$, NUSERIE$, 1, 32)
       Call REPLA$(VDEF$, MKI$(CI%), 33, 2)
       Call REPLA$(VDEF$, MKI$(HOY(HO$)), 35, 2)
       '
10     OBX$ = OBJPLA$("CARNUMSE", VDEF$) ' OBJETO TANUMSE PARA INGRESAR LOS NUMEROS DE SERIE
       VDEF$ = OBX$
       If OBX$ = "" Then GoTo 1
       '
       NUSERIEE$ = TRIM$(Mid$(VDEF$, 1, 32))
       CI% = CVI(Mid$(VDEF$, 33, 2))
       fech% = CVI(Mid$(VDEF$, 35, 2))
       FECC = FECHATEX(fech%)
       TENSION$ = Mid$(VDEF$, 37, 6)
       NUDEPO% = CVI(Mid$(VDEF$, 43, 2))
       '
       If NUDEPO% > 0 Then
          Call MENSERR(24, "Está Ingresando un Equipo para Reparación.\   \El Depósito debe Quedar en Blanco.")
          Call REPLA(VDEF$, MKI$(0), 43, 2)
          GoTo 10
       End If
       '
       CANTIRE& = CantRegistros&(" TANUMSE", "NumeroSerie = '" & NUSERIEE$ & "'") 'VERIFICA SI EL NUMERO DE SERIE EXISTE
       If CANTIRE& > 0 Then
           Call MENSERR(24, "Número de Serie Existente")
           GoTo 10 ' VALIDACIONES
       End If
       '
       If CI% < 1 Then
           Call MENSERR(24, "Código no Válido")
           GoTo 10
       End If
       '
       If InStr(EMPREAC$, "MEDITEA") > 0 Then
         If Len(NUSERIEE$) > 4 Then
           Ulti4Cod$ = Right(CODEXT$(CI%), 4) ' Valida que los ultimos cuatro numeros del codigo coincidan con los 4 primeros del numero de serie
           Prim4NuSe1$ = Left$(NUSERIEE$, 4)
           '
           If Ulti4Cod$ <> Prim4NuSe1$ Then
             Call COMUNI("Número de Serie Difiere de Código")
             GoTo 10
           End If
         Else
           Call MENSERR(24, "Número de Serie Cantidad de Caracteres Menor a 5")
           GoTo 10
         End If
       End If
       '
       MESYANIO$ = Mid$(NUSERIEE$, 5, 4)
       NORDEN% = XVALO(Mid$(NUSERIEE$, 9, 4))
       '
       FECHAA$ = FECHATEX$(fech%)
       SQLX$ = "SELECT * FROM TANUMSE" ' SI PASO LAS VALIDACIONES AGREGA EL NUEVO NUMERO DE SERIE
       Dim TANUTMP As ADODB.Recordset
       Set TANUTMP = New ADODB.Recordset
       TANUTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
         With TANUTMP
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CodiEmpr = CodiEmp%
           On Error GoTo 0
           !NUMEROSERIE = NUSERIEE$
           !MESANIO = MESYANIO$
           !NUMEORD = NORDEN%
           !COD_INTE = CI%
           !COD_EXTE = CODEXT(CI%)
           Descripcion$ = Left$(DESCRIT$(CI%), 64)    'modificado por que devolvia un error de datos
           !Descrip = Descripcion$
           !FECHALTA = FECC
           !Ten_Ensa = XVALO(TENSION$)
           !DEPONUMSE = NUDEPO%
           .Update
         End With
       TANUTMP.Close
       Set TANUTMP = Nothing
       '
       Call ACTLISELNUMSER   'Actualiza lista de selecciona de numeros de serie
       '
       'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
       CUERPO$ = "Alta N/Serie en Forma Manual"
       Call ANOTANUSE(NUSERIEE$, CUERPO$)
       '
       Call PRINTALBANUMSE(NUSERIEE$, CI%, FECC, 1)
       '
       Call MENSERR(24, "Alta Manual Realizada")
       VDEF$ = ""
       Screen.MousePointer = 1
       Me.Text2 = NUSERIEE$
       '
       GoTo 1
       '
   End If
   '
99 Command3.Enabled = True
   '
End Sub
Sub ALTAMANUAL_NSERIE_EYM()
3      Call SELECHO("MASTER")
       Codigo$ = VALACT1("MASTER")
       CI1% = CODINT(Codigo$)
       If CI1% < 1 Then Exit Sub
       
5      VDEF$ = String$(48, 0)

       Call REPLA$(VDEF$, MKI$(CI1%), 1, 2)
       Call REPLA$(VDEF$, NUSERIE$, 3, 32)
       Call REPLA$(VDEF$, MKI$(HOY(HO$)), 35, 2)
       Call REPLA$(VDEF$, "220", 37, 12)
       '
10     OBX$ = OBJPLA$("ALSERIE2", VDEF$) ' OBJETO TANUMSE PARA INGRESAR LOS NUMEROS DE SERIE
       VDEF$ = OBX$
       If OBX$ = "" Then GoTo 3
       '
       CI1% = CVI(Mid$(VDEF$, 1, 2))
       NUSERIEE$ = TRIM$(Mid$(VDEF$, 3, 32))
       fech% = CVI(Mid$(VDEF$, 35, 2))
       FECC = FETEXCOR1$(fech%)
       TENSION$ = TRIM$(Mid$(VDEF$, 37, 12))
       '
       CANTIRE& = CantRegistros&(" TANUMSE", "NumeroSerie = '" & NUSERIEE$ & "'") 'VERIFICA SI EL NUMERO DE SERIE EXISTE
       If CANTIRE& > 0 Then
           Call MENSERR(24, "Número de Serie Existente")
           GoTo 10 ' VALIDACIONES
       End If
       '
       If CI1% < 1 Then
           Call MENSERR(24, "Código no Válido")
           GoTo 10
       End If
       '
       MESYANIO$ = Left(NUSERIEE$, 2)
       NORDEN% = XVALO(Right(NUSERIEE$, 3))
       
       '
       FECHAA$ = FECHATEX$(fech%)
       SQLX$ = "SELECT * FROM TANUMSE" ' SI PASO LAS VALIDACIONES AGREGA EL NUEVO NUMERO DE SERIE
       Dim TANUTMP As ADODB.Recordset
       Set TANUTMP = New ADODB.Recordset
       TANUTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
         With TANUTMP
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CodiEmpr = CodiEmp%
           On Error GoTo 0
           !NUMEROSERIE = NUSERIEE$
           !MESANIO = MESYANIO$
           !NUMEORD = NORDEN%
           !COD_INTE = CI1%
           !COD_EXTE = CODEXT(CI1%)
           Descripcion$ = Left$(DESCRIT$(CI1%), 64)    'modificado por que devolvia un error de datos
           !Descrip = Descripcion$
           !FECHALTA = FECC
           !Ten_Ensa = XVALO(TENSION$)
           !DEPONUMSE = 0
           FEGARANTIA% = DIASPOST(fech%, 730)
           !FEVENGAR = FETEXCOR1$(FEGARANTIA%)
           .Update
         End With
       TANUTMP.Close
       Set TANUTMP = Nothing
       '
       Call ACTLISELNUMSER   'Actualiza lista de selecciona de numeros de serie
       '
       'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
       CUERPO$ = "Alta N/Serie en Forma Manual"
       Call ANOTANUSE(NUSERIEE$, CUERPO$)

End Sub
 Sub ACTLISELNUMSER() ' Actualiza listado de seleccion de numeros de serie
    '
    Call CARGALISEL("INDINUSE", "TANUMSE", "NUMEROSERIE/A16;COD_EXTE/A18;DESCRIP/A24", "NUMEROSERIE <> '' ORDER BY FECHALTA ASC, NUMEORD ASC")
    '
 End Sub
     
 Sub PRINTALBANUMSE(NSERIE$, CI%, FECHA, Optional TIP%)
 
    'FORIPRE$ = TRIM$(CONTROL$("", "ALBANSE"))
    FORIPRE$ = TRIM$(CONTROL$("FOREPAR", "DOCUREP6"))
    If FORIPRE$ = "" Then Exit Sub
    '
    Codigo$ = CODEXT$(CI%)
    DESXRIP$ = DESCRIT0$(CI%)
    NUSERIE$ = NSERIE$
    FechEmix$ = FECHA
    '
    If TIP% = 1 Then
        NUMEDOC$ = NUSERIE$
        TIPODOC$ = "Alta de Numero de Serie"
    ElseIf TIP% = 2 Then
        NUMEDOC$ = NUSERIE$
        TIPODOC$ = "Baja de Numero de Serie"
    End If
    CODPARAM$(1) = "COD-MAT": DOCPARAM$(1) = Codigo$
    CODPARAM$(2) = "DES-MAT": DOCPARAM$(2) = DESXRIP$
    CODPARAM$(3) = "REF-MAT1": DOCPARAM$(3) = NUSERIE$
    CODPARAM$(4) = "FECH-EMI": DOCPARAM$(4) = FechEmix$
    CODPARAM$(5) = "N-PLANO": DOCPARAM$(5) = NUSERIE$
    CODPARAM$(6) = "R-PLANO": DOCPARAM$(6) = CODIBARRA$(NUSERIE$, 1)
    CODPARAM$(7) = "F-PLANO": DOCPARAM$(7) = CONVIERTOCODIGOBARRAS$(NUSERIE$)
    
    CAPARDOC = 7
    '
    PRINTERPORT$ = ""
    PRINTERPORT$ = CONTROL$(IDENTER$, "PORETIRE") ' IMPRESORA CONFIGURADA
    If PRINTERPORT$ = "" Then PRINTERPORT$ = CONTROL$("*", "PORETIRE")
    Call SETPRINPORT(PRINTERPORT$)
    '
    Call PRINTDOC("@" + FORIPRE$)
    '
    Call SETPRINPORT("RESTORE")
    '
 End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   If VALIDATOS% < 1 Then GoTo 99
   '
   APRO = "NEXT"
   Command4.Enabled = True
   Hide
   '
99 Command4.Enabled = True

End Sub

Function VALIDATOS%()
   '
   VALIDATOS% = 0
   QTR& = CantRegistros("TANUMSE", "NUMEROSERIE = '" & Text2 & "'", "NOMESS")
   If QTR& < 1 Or TRIM$(Text2) = "" Then
      Call MENSERR(24, "Numero de Serie Inexistente o No Válido")
      Call PINTATEXT(Text2)
      Exit Function
   End If
   '
   CIXI% = CODINT(Label9)
   If CIXI% < 1 Then
      Call MENSERR(24, "Codigo No Válido")
      Exit Function
   End If
   '
   If TRIM$(Text11) = "" Then
      Call MENSERR(24, "Falta Declarar Sector (Obligatorio)")
      Call PINTATEXT(Text11)
      Exit Function
   End If
   '
   If XVALO(Text1) < 1 Or rasocli$(XVALO(Text1)) = "" Then
      Call MENSERR(24, "Cliente Inexistente o No Válido")
      Call PINTATEXT(Text1)
      Exit Function
   End If
   '
   If TRIM$(Text7) = "" Then
      Call MENSERR(24, "Falta Telefono de Contacto")
      Call PINTATEXT(Text7)
      Exit Function
   End If
   '
   If TRIM$(Text10) = "" Or InStr(Text10, "@") < 2 Or LCase(Text10) <> Text10 Then
      Call MENSERR(24, "Falta Direccion de Mail de Contacto o No Válida")
      Call PINTATEXT(Text10)
      Exit Function
   End If
   '
   If TRIM$(Text4) = "" Or TRIM$(Text6) = "" Then
      Call MENSERR(24, "Falta Datos Domicilio de Entrega")
      Call PINTATEXT(Text4)
      Exit Function
   End If
   '
   If TRIM$(Text9) = "" Then
      Call MENSERR(24, "Falta Nombre de la Persona de Contacto")
      Call PINTATEXT(Text9)
      Exit Function
   End If
   '
   If TRIM$(Text14) = "" And XVALO(Text13) > 0 Then
      Call MENSERR(24, "Falta N° de Talón")
      Call PINTATEXT(Text14)
      Exit Function
   End If
   
   VALIDATOS% = 1
   '
End Function

Private Sub Command5_Click()
    Command5.Enabled = False
    CONUSE07.Label7 = Text2
    CONUSE07.Show 1
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Call SELECHO("FECHA")
   If VALACT1$("SELFECHA") <> "" Then
      Label14 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   If DERACCE%("REIMPTAL", "Re-Impresión de Talon") > 0 Then
       Call COPYSTRU("INDIPRES", "INDERPCI")
15     Call CARGALISEL("INDERPCI", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0  AND STATUS < 4", "DISTINCT")
       Call SELECHO("INDERPCI/Indice General de Reparación")
       NroRep& = XVALO(TRIM$(VALACT1("INDERPCI")))
       If NroRep& > 0 Then
          FR% = CVINT(VALOBADA("REPARA", "FechRecep", "NUMREPA = " & NroRep& & "", "NOMESS"))
          FREP$ = FECHATEX$(FR%)
          NC = XVALO(VALOBADA("REPARA", "NUME_CLI", "NUMREPA = " & NroRep& & "", "NOMESS"))
          TEX& = XVALO(VALOBADA("REPARA", "TEC_EXT", "NUMREPA = " & NroRep& & "", "NOMESS"))
          FILTX$ = TRIM$(Str$(NroRep&)) & ";" & TRIM$(FREP$) & ";" & TRIM$(Str$(NC))
          ARCSQR$ = TRIM$(CONTROL$("FOREPAR", "DOCUREP" + "5"))
          If ARCSQR$ <> "" Then
               'PARA EYM
               Call STDFORM(ARCSQR$, FILTX$)
               'PARA EL CLIENTE
               FILTX$ = TRIM$(Str$(NroRep&)) & ";" & TRIM$(FREP$) & ";" & TRIM$(Str$(NC))
               ARCSQR$ = TRIM$(CONTROL$("FOREPAR", "DOCUREP" + "7"))
               Call STDFORM(ARCSQR$, FILTX$)
               Exit Sub
          End If
          
          If TEX& > 0 Then
            FILTX$ = FILTX$ & ";" & "-TECNICO EXTERNO-"
            Call STDFORM("NREP-MDT", FILTX$)
          Else
            FILTX$ = FILTX$ & ";" & "===================="
            Call STDFORM("NREP-MDT", FILTX$, "PRINT")
          End If
       End If
   End If
99 Command7.Enabled = True
End Sub

Private Sub Command8_Click()
  Text13 = MUESTRATECNICO
  Text12 = RESP_ZELE_VECT(2) 'NOMBRE DEL TECNICO
End Sub

'
Private Sub Command9_Click()
   NSERIE$ = Text2
   CONDI$ = "NUMEROSERIE = '" & NSERIE$ & " '"
   NORFA$ = VALOBADA("TANUMSE", "IDSUBOR", CONDI$, "NOMESS")
   FRMREGARMADO.Label13(2) = NSERIE$
   FRMREGARMADO.Show 1
   
End Sub

Private Sub Form_Load()
   '
   ' ACTUALIZO LOS CAMPOS NULL
   ' Call PoneEnCeroNull("MOVISTO", "VERIFICADOPOR", REGAF&, "NOMESS")
   ' Call ACTUREGISTRO("MOVISTO", "FHOR_VERIF", "FHOR_VERIF Is Null", NUNCA$, REGAF&, "NOMESS")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If InStr(UCase$(EPAC$), "EYM") > 0 Then
      Label18.Visible = True
   End If
End Sub

Private Sub Label16_DblClick()
   Call SELECHO("FECHA")
   If VALACT1$("SELFECHA") <> "" Then
      ferece% = FECHANUM(VALACT1$("SELFECHA"))
      If ferece% < 30 Then ferece% = Now
      FEXA$ = FECHATEX$(ferece%)
      Label16 = Format(FEXA$, "dd/mm/yy - hh:mm")
   End If
End Sub

Private Sub Label7_DblClick()
   If DERACCE%("STATGARA", "CAMBIAR ESTADO DE GARANTÍA") < 2 Then Exit Sub
   If TRIM$(Label7.Caption) = "" Then Exit Sub
   '
   SIONO$ = TRIM$(Label7.Caption)
   If SIONO$ = "SI" Then
      Label7.Caption = "NO"
'      If Label7.BackColor = &HEEEEEE Then no le cambio el color por que si no hay que verificar en el cambio de datos.
'         Label7.BackColor = &H808080
'      Else
'         Label7.BackColor = &HEEEEEE
'      End If
   ElseIf SIONO$ = "NO" Then
      Label7.Caption = "SI"
'      If Label7.BackColor = &HEEEEEE Then
'         Label7.BackColor = &H808080
'      Else
'         Label7.BackColor = &HEEEEEE
'      End If

   End If

End Sub


Private Sub Text1_Change()
    '
    Call BLANFORMU
    NCX7& = XVALO(Text1.TEXT)
    If NCX7& > 0 Then
        Text3.TEXT = rasocli$(NCX7&)
        Text4.TEXT = DOMICLI$(NCX7&)
        Text5.TEXT = CPOSCLI$(NCX7&)
        Text6.TEXT = LOCACLI$(NCX7&)
        Text7.TEXT = TEL1CLI$(NCX7&)
        Text8 = FAXCLI$(NCX7&)
        Text9 = CONTACTCLI$(NCX7&)
        Text10 = EMAILCLI$(NCX7&)
        Label26 = CODICLI$(NCX7&)
    End If
    '
    NREPA& = XVALO(Label15)
    CONDI$ = "NUMREPA = " & NREPA&
    If CantRegistros("REPARA", CONDI$) > 0 Then
        'Text3.TEXT = VALOBADA("REPARA", "Raso_Cli", CONDI$, "NOMESS")
        Text3.TEXT = rasocli$(NCX7&) 'ESTO LO TRAIGO POR ECO POR QUE NO ES EDITABLE EL RESTO SE PUEDE MODIFICAR
           If TRIM$(Text3) = "" Then Text3.TEXT = rasocli$(NCX7&)
        Text4.TEXT = VALOBADA("REPARA", "Domi_Cli", CONDI$, "NOMESS")
           If TRIM$(Text4) = "" Then Text4 = DOMICLI$(NCX7&)
        Text6.TEXT = VALOBADA("REPARA", "Loca_Cli", CONDI$, "NOMESS")
           If TRIM$(Text6) = "" Then Text6 = LOCACLI$(NCX7&)
        Text7.TEXT = VALOBADA("REPARA", "Tele_Cli", CONDI$, "NOMESS")
           If TRIM$(Text7) = "" Then Text7.TEXT = TEL1CLI$(NCX7&)
        Text10 = VALOBADA("REPARA", "Mail_Cli", CONDI$, "NOMESS")
           If TRIM$(Text10) = "" Then Text10 = EMAILCLI$(NCX7&)
        Text9 = VALOBADA("REPARA", "Ctac_Cli", CONDI$, "NOMESS")
           If TRIM$(Text9) = "" Then Text9 = CONTACTCLI$(NCX7&)
    End If
End Sub

Private Sub Text1_GotFocus()
    If TRIM$(Text2) = "" Then Text2.SetFocus
End Sub


Private Sub Text10_GotFocus()
    If TRIM$(Text2) = "" Then Text2.SetFocus
End Sub

Private Sub Text2_Change()
      Label9 = ""
      Label10 = ""
      Label31 = ""
      Label5 = ""
      Label7 = ""
      Text1 = ""
      Text11 = ""
      Text12 = ""
      Text13 = ""
      Text14 = ""
      Label18 = ""
      '
      EPAC$ = TRIM$(EMPREAC$)

      CONDI$ = "NUMEROSERIE = '" & TRIM$(Text2) & "'"
      QTR& = CantRegistros("TANUMSE", CONDI$, "NOMESS")
      If QTR& > 0 Then
         CIXI% = XVALO(VALOBADA("TANUMSE", "COD_INTE", CONDI$, "NOMESS"))
         If CIXI% > 0 Then
            Label9 = CODEXT$(CIXI%)
            Label10 = DESCRIT0$(CIXI%)
         End If
         '
         ' TECNICO ULTIMO CONTROL
         Label5 = SAFETEXT$(VALOBADA("TANUMSE", "Tec_Ensa", CONDI$, "NOMESS"))
         '
         ' FECHA DE ULTIMA REPARACION
         Condi1$ = "NUMEROSERIE = '" & TRIM$(Text2) & "' ORDER BY FECHENT DESC"
         FUREP% = CVINT(VALOBADA("REPARA", "FECHENT", Condi1$, "NOMESS"))
         Label31 = FECHATEX$(FUREP%)
         '
         ' FALTA GARANTIA SI / NO
         ENGARAN$ = "NO"
         FEVENTA% = CVINT(VALOBADA("TANUMSE", "FechEntrega", CONDI$, "NOMESS"))
         If FEVENTA% < 1 Then
             FEVENTA% = CVINT(VALOBADA("MOVISTO", "MAX(FECHMOV)", "DOSECLAR LIKE 'FC%' AND IDENLOTE = '" & Text2 & "'"))
         End If
         
         If DIENTRE%(FEVENTA%, HOY(HOS$)) <= 730 Then ENGARAN$ = "SI"    ' 2 AÑOS DE LA VENTA

         If DIENTRE%(FUREP%, HOY(HOS$)) <= 90 Then ENGARAN$ = "SI"
         
         Label7 = ENGARAN$
         
         If InStr(UCase$(EPAC$), "EYM") > 0 Then
            'modificacion 14/09/21
            NUSERIE$ = TRIM$(Text2)
            'PARA EYM SOLO TOMO FECHA DE SALIDA DEL TANUMSE NO DEL MOVISTO POR QUE SI NO TRAE LA DE LAS REPARACIONES
            CONDI$ = "NUMEROSERIE = '" & NUSERIE$ & "'"
            FEVENTA% = CVINT(VALOBADA("TANUMSE", "FechEntrega", CONDI$, "NOMESS"))
'Call mostrarDatosx(CONDI$)
'Call mostrarDatosx("FECHA DE VENTA " & FECHATEX(FEVENTA%))

            Label7.Caption = EstadoGarantiayLetra$(FEVENTA%, FUREP%, RetLetra$, NUSERIE$)
            Label18.Caption = RetLetra$
            
         End If
         
         'MANTENGO EL VALOR DE LA GARANTIA ORIGINAL (X Q SE PUEDE CAMBIAR EN PANTALLA)
         ENGARANTX% = 0: If TRIM$(Label7.Caption) = "SI" Then ENGARANTX% = 1
         LBLSTATGARAN_ORIGINAL = ENGARANTX%
         '
         FUREPAR = VALOBADA("REPARA", "MAX(FechEnt)", CONDI$, "NOMESS")
         If FUREPAR = "" Then
            FUREPAR = VALOBADA("REPARA", "MAX(FechRECEP)", CONDI$, "NOMESS")
         End If
         If FUREPAR <> "" Then
            Label31 = FECHATEX$(CVINT(FUREPAR))
         End If
         '
         Text1 = XVALO(VALOBADA("TANUMSE", "NUMECLI", CONDI$, "NOMESS"))
         '
      End If
      '
End Sub

Function EstadoGarantiayLetra$(FEVENTA%, FUREP%, RetLetra$, NSERIE$)
'              POR DEFECTO NO
               EstadoGarantiayLetra$ = "NO"
'Call mostrarDatosx("FECHA ULTIMA REPARCION " & FECHATEX(FUREP%))

               If FUREP% < 1 Then
                  'SI NO TIENE FECHA DE ENTREGA SE FIJA SI TIENE FECHA DE CIERRE DE REPARACION
                  CONDI$ = "NUMEROSERIE = '" & NUSERIE$ & "'"
                  FUREP% = CVINT(VALOBADA("REPARA", "MAX(FECHACIERREREPA)", CONDI$, "NOMESS"))
'Call mostrarDatosx("FECHA ULTIMA REPARCION DE REPARA" & FECHATEX(FUREP%))
               End If

               If FEVENTA% > 60 Or FUREP% > 60 Then
'Call mostrarDatosx("ENTRO EN CONDICION 1")
                    'SI ESTA ENTRE 2 Y 5 AÑOS ES B
                    If DIENTRE%(FEVENTA%, HOY(HOS$)) <= 730 Then EstadoGarantiayLetra$ = "SI": RetLetra$ = "A": Exit Function
                    If DIENTRE%(FUREP%, HOY(HOS$)) <= 180 Then EstadoGarantiayLetra$ = "SI": RetLetra$ = "A": Exit Function
                    If DIENTRE%(FEVENTA%, HOY(HOS$)) > 730 And DIENTRE%(FEVENTA%, HOY(HOS$)) <= 1825 Then RetLetra$ = "B": Exit Function 'entre 2 y 5 años
                    If DIENTRE%(FEVENTA%, HOY(HOS$)) > 1825 Then RetLetra$ = "C": Exit Function 'mayor a 5 años
               Else
'Call mostrarDatosx("ENTRO EN CONDICION 2")
                     'SACAR FECHA DE VENTA A PARTIR DEL NRO DE SERIE
                     'ESTO APLICARIA SOLO PARA LOS PRODUCTOS QUE NO TUVIERON REPARACION
                     SINGUION$ = TRIM$(REPLACAR(NSERIE$, "-", ""))
'Call mostrarDatosx("SINGUION$ " & SINGUION$)
                     ANO1$ = Left(SINGUION$, 2)
'Call mostrarDatosx("ANO1$ " & ANO1$)
                     SEMANAX$ = Mid$(SINGUION$, 3, 2)
'Call mostrarDatosx("SEMANAX$ " & SEMANAX$)
                     ANOSEMANAX$ = SEMANAX$ & ANO1$
'Call mostrarDatosx("ANOSEMANAX$ " & ANOSEMANAX$)
                     FeSegunSerie% = FECHSEMAN(ANOSEMANAX$)
'Call mostrarDatosx("FeSegunSerie% " & FECHATEX(FeSegunSerie%))
                     If DIENTRE%(FeSegunSerie%, HOY(HOS$)) <= 730 Then
'Call mostrarDatosx("DIENTRE%(DIENTRE% " & DIENTRE%(FeSegunSerie%, HOY(HOS$)))
                        EstadoGarantiayLetra$ = "SI": RetLetra$ = "A"
                     ElseIf DIENTRE%(FeSegunSerie%, HOY(HOS$)) > 730 And DIENTRE%(FeSegunSerie%, HOY(HOS$)) <= 1825 Then
                        RetLetra$ = "B"
                    Else
                        'SI NO ES NINGUNO DE LOS DOS ANTERIORES TIENE MAS DE 5 ES C
                        RetLetra$ = "C"
                    End If
              End If

End Function

Sub BLANFORMU()
    '
    Text3.TEXT = ""
    Text4.TEXT = ""
    Text5.TEXT = ""
    Text7 = ""
    Text6 = ""
    Text8 = ""
    Text9 = ""
    Text10 = ""
    Label26 = ""
End Sub

Private Sub Text3_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text4_GotFocus()
    If TRIM$(Text2) = "" Then Text2.SetFocus
End Sub

Private Sub Text5_GotFocus()
    If TRIM$(Text2) = "" Then Text2.SetFocus
End Sub

Private Sub Text6_GotFocus()
    If TRIM$(Text2) = "" Then Text2.SetFocus
End Sub

Private Sub Text7_GotFocus()
    If TRIM$(Text2) = "" Then Text2.SetFocus
End Sub

Private Sub Text8_GotFocus()
    If TRIM$(Text2) = "" Then Text2.SetFocus
End Sub

Private Sub Text9_GotFocus()
    If TRIM$(Text2) = "" Then Text2.SetFocus
End Sub
