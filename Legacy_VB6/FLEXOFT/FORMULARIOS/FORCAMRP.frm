VERSION 5.00
Begin VB.Form FORCAMRP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cálculo Necesidad de Componentes según Programa de Entregas"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4935
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   10
      Top             =   2205
      Width           =   870
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   855
      Left            =   4935
      Picture         =   "FORCAMRP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Abandona Proceso de Cálculo"
      Top             =   320
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Avance de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   210
      TabIndex        =   8
      Top             =   210
      Width           =   4515
      Begin VB.CheckBox Check7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenes de Compra y Fabricación a Emitir"
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
         Left            =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2415
         Width           =   4250
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3960
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   525
         Width           =   270
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Generación de Necesidades Secundarias"
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
         Left            =   240
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2100
         Width           =   4250
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cálculo de Necesidades Netas"
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
         Left            =   240
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1785
         Width           =   4250
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Apertura de Archivos Maestros"
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
         Left            =   240
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   525
         Width           =   4250
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Existencia Actual de Componentes"
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
         Left            =   240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   840
         Width           =   4250
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Inicialización Necesidades Secundarias"
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
         Left            =   240
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1155
         Width           =   4250
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Necesidades según Entregas Pendientes"
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
         Left            =   240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1470
         Width           =   4250
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5120
      Picture         =   "FORCAMRP.frx":030A
      Top             =   1470
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4360
      Picture         =   "FORCAMRP.frx":0614
      Top             =   2550
      Width           =   2010
   End
End
Attribute VB_Name = "FORCAMRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LOST$(10)
Dim LOSTN(10)
Dim CII1%(1024), CANT1(1024), PAD1%(1024)
Dim INIPER%(52), CAPER(52), DEPER$(52)
'
Dim LOREC As Integer
Dim CACOL As Integer
Dim COSEMACT As Integer
Dim HOO As Integer
Dim HO As String
Dim DMX%(12)
'
'''' Dim DISPACT(16384)
Dim DISPACT(16384), PUNOCPE$(16384)
Dim LOREPOS(16384)
Dim DPI%(16384)
Dim PROVARTI%(16384)
'
Dim SST As String * 40
Dim M0 As String * 128
Dim RA As String * 128
Dim RB As String * 256
Dim NESTRU As Long
Dim PUNTCI As String
Dim PUNTRG As String
'
Dim SEMANAL As Integer
Dim PERIO As String
Dim PLADISTA As Integer
Dim REPOMIN As Single
'
Dim CTX As Integer
Dim CACO As Long
Dim PUNTE$(16384)
'
Dim PROVINT%(32767)

' realizar calculo bajo formulario FORCAMRP
' verificar lectura de compras pendientes segun OCOMDETA
' netear programa segun entregas realizadas en el mes
' mal fecha programada - ver limite de fecha considerada
' considerar pedidos pendientes de entrega - ver como

Private Sub Form_Activate()
    Refresh
    Call CANECOM
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Timer1_Timer()
   If Image1.Visible = True Then
       Image1.Visible = False
     Else
       Image1.Visible = True
   End If
   DoEvents

End Sub

