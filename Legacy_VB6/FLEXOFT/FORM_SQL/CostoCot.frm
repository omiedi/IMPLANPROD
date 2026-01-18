VERSION 5.00
Begin VB.Form CostoCot 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Rutas de Base Waldbott  de Origen y Lectura"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   4815
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
   ScaleHeight     =   3015
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   6500
      Left            =   3780
      ScaleHeight     =   6435
      ScaleWidth      =   945
      TabIndex        =   4
      Top             =   0
      Width           =   1000
      Begin VB.CommandButton Command2 
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
         Picture         =   "CostoCot.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1710
         Width           =   645
      End
      Begin VB.CommandButton command1 
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
         Picture         =   "CostoCot.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   105
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "CostoCot.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   2520
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Presentación de Costo y Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   0
      Top             =   9000
      Width           =   5160
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Calculado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Precio Base (C.Rep. + 15 %)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Informado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3615
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
         Index           =   20
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "GRAMAJE PAPEL"
         Top             =   480
         Width           =   1800
      End
      Begin VB.TextBox TXTGRAMAJEPAPEL 
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
         Left            =   2040
         TabIndex        =   16
         Top             =   480
         Width           =   1000
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
         Index           =   27
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "PRECIO PP"
         Top             =   960
         Width           =   1800
      End
      Begin VB.TextBox TXTPRECIOPP 
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
         Left            =   2040
         TabIndex        =   14
         Top             =   960
         Width           =   1000
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
         Index           =   28
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "PRECIO PAPEL"
         Top             =   1320
         Width           =   1800
      End
      Begin VB.TextBox TXTPRECIOPAPEL 
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
         Left            =   2040
         TabIndex        =   12
         Top             =   1320
         Width           =   1000
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
         Index           =   30
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "COSTO IMPRESIÓN"
         Top             =   1800
         Width           =   1800
      End
      Begin VB.TextBox TXTCOSTOIMPRESION 
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
         Left            =   2040
         TabIndex        =   10
         Top             =   1800
         Width           =   1000
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
         Index           =   32
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "COSTO CONFECCIÓN"
         Top             =   2160
         Width           =   1800
      End
      Begin VB.TextBox TXTCOSTOCONFECCION 
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
         Left            =   2040
         TabIndex        =   8
         Top             =   2160
         Width           =   1000
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuConfRutas 
         Caption         =   "Configuracion de rutas"
      End
   End
End
Attribute VB_Name = "CostoCot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()


    Call GRACONTROL("COTICONT", "GRAMAPEL", TXTGRAMAJEPAPEL.TEXT)
    Call GRACONTROL("COTICONT", "PRECIOPP", TXTPRECIOPP.TEXT)
    Call GRACONTROL("COTICONT", "PRECIPAP", TXTPRECIOPAPEL.TEXT)
    Call GRACONTROL("COTICONT", "COSTIMPR", TXTCOSTOIMPRESION.TEXT)
    Call GRACONTROL("COTICONT", "COSTCONF", TXTCOSTOCONFECCION.TEXT)
   
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completada.\Re-Inicie Ahora la Aplicación '" + UCase$(App.EXEName) + "'.")
    Call FINAL("")
    '
End Sub


Private Sub Form_Load()
    '
    Me.TXTGRAMAJEPAPEL = CONTROL("COTICONT", "GRAMAPEL")
    Me.TXTPRECIOPP = CONTROL("COTICONT", "PRECIOPP")
    Me.TXTPRECIOPAPEL = CONTROL("COTICONT", "PRECIPAP")
    Me.TXTCOSTOIMPRESION = CONTROL("COTICONT", "COSTIMPR")
    Me.TXTCOSTOCONFECCION = CONTROL("COTICONT", "COSTCONF")
    '
End Sub

Private Sub mnuConfRutas_Click()
   If DERACCE%("RutaExce", "Configuraciòn de Ruta Excel") >= 2 Then
      BkpWaldb.Show 1
   End If
End Sub

Private Sub TXTGRAMAJEPAPEL_LostFocus()
   TXTGRAMAJEPAPEL = FORMATNUM$(XVALO(TXTGRAMAJEPAPEL), "F12.2")
End Sub
Private Sub TXTPRECIOPP_LostFocus()
   TXTPRECIOPP = FORMATNUM$(XVALO(TXTPRECIOPP), "F12.2")
End Sub
Private Sub TXTPRECIOPAPEL_LostFocus()
   TXTPRECIOPAPEL = FORMATNUM$(XVALO(TXTPRECIOPAPEL), "F12.2")
End Sub
Private Sub TXTCOSTOIMPRESION_LostFocus()
   TXTCOSTOIMPRESION = FORMATNUM$(XVALO(TXTCOSTOIMPRESION), "F12.2")
End Sub
Private Sub TXTCOSTOCONFECCION_LostFocus()
   TXTCOSTOCONFECCION = FORMATNUM$(XVALO(TXTCOSTOCONFECCION), "F12.2")
End Sub


