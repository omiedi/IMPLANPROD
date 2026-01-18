VERSION 5.00
Begin VB.Form SELITFAC 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion de Items a Facturar"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command18 
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
      Height          =   570
      Left            =   9795
      Picture         =   "SELITFAC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancela la Factura Activa"
      Top             =   900
      Width           =   570
   End
   Begin VB.CommandButton Command19 
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
      Height          =   570
      Left            =   9795
      Picture         =   "SELITFAC.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Continúa Generación de la Factura"
      Top             =   240
      Width           =   570
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00E0E0E0&
      Height          =   2865
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   0
      Width           =   9645
   End
   Begin VB.Line Line5 
      X1              =   9640
      X2              =   9640
      Y1              =   0
      Y2              =   3200
   End
End
Attribute VB_Name = "SELITFAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command18_Click()
    List2.Clear
    Unload Me
End Sub

Private Sub Command19_Click()
    Hide
End Sub
