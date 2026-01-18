VERSION 5.00
Begin VB.Form OPFACTURA 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Facturación"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6660
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
   ScaleHeight     =   6105
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5600
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   13
      Top             =   5040
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Opciones de Trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5730
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   5160
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Despacho de Materiales con Remito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   17
         Top             =   1050
         Width           =   4110
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Texto Libre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3465
         TabIndex        =   15
         Top             =   2100
         Width           =   1590
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Anulación de Comprobantes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   11
         Top             =   5250
         Width           =   3900
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Devoluciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2940
         TabIndex        =   10
         Top             =   4200
         Width           =   1800
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Texto Libre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   8
         Top             =   4200
         Width           =   1800
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Formato de Texto Libre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   6
         Top             =   3150
         Width           =   2745
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Formato Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   4
         Top             =   2100
         Value           =   -1  'True
         Width           =   2325
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Remitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   420
         TabIndex        =   16
         Top             =   525
         Width           =   2850
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Anulaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   420
         TabIndex        =   12
         Top             =   4725
         Width           =   2850
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas de Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   420
         TabIndex        =   9
         Top             =   3675
         Width           =   2850
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas de Débito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   420
         TabIndex        =   7
         Top             =   2625
         Width           =   2850
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Facturación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   420
         TabIndex        =   5
         Top             =   1575
         Width           =   2850
      End
   End
   Begin VB.CommandButton Command7 
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
      Height          =   855
      Left            =   5600
      Picture         =   "OPFACTURA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2205
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "O.K."
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
      Left            =   5600
      Picture         =   "OPFACTURA.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1260
      Width           =   855
   End
   Begin VB.CommandButton command1 
      Caption         =   "Exit"
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
      Left            =   5600
      Picture         =   "OPFACTURA.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   335
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   5340
      Picture         =   "OPFACTURA.frx":075E
      Stretch         =   -1  'True
      Top             =   5365
      Width           =   5250
   End
End
Attribute VB_Name = "OPFACTURA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_Click()
   Hide
   If Option1.Value = True Then       ' Facturación Itemizada
       FORFACTURA.Show
     ElseIf Option2.Value = True Then ' Notas de Débito de Texto Libre
        FORFACTURA.Hide
        FORDELIDIP.Show
     ElseIf Option3.Value = True Then ' Notas de Crédito de Texto Libre
        FORFACTURA.Hide
        FORCRELIDIP.Show
     ElseIf Option4.Value = True Then ' Notas de Crédito por Devoluciones
        FORFACTURA.Hide
        FORCREDEVDIP.Show
     ElseIf Option5.Value = True Then ' Anulacion de Comprobantes
       Call ANUFAC
     ElseIf Option6.Value = True Then ' Facturas de Acopio
        FORFACTURA.Hide
        FORFACABO.Show
     ElseIf Option7.Value = True Then ' Remito
        Call CONECRUN("REMSTAN", "Despacho de Materiales con Remito")
     Else
       FORFACTURA.Show
       Option1.Value = True
       Call MENSERR(24, "La Opción Elegida no se Encuentra Disponible por el Momento.")
   End If
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Hide
   FORFACTURA.Hide
   FORCONFABO.Show
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()   ' estadistica de entregas a abonados
    '
    ESTADIPRO.Show 1
    '
End Sub

Private Sub Command7_Click()
   Call HELPONLINE("FACTUDIP")
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Option1_dblClick() ' Facturacion Itemizada
   Hide
   FORFACDIP.Show
End Sub

Private Sub Option2_dblClick() ' Notas de Débito de Texto Libre
   Hide
   FORFACDIP.Hide
   FORDELIDIP.Show
End Sub

Private Sub Option3_DBLClick() ' Notas de Crédito de Texto Libre
   Hide
   FORFACDIP.Hide
   FORCRELIDIP.Show
End Sub

Private Sub Option4_DBLClick() ' Notas de Crédito por Devoluciones
   Hide
   FORFACDIP.Hide
   FORCREDEVDIP.Show
End Sub

Private Sub Option5_DblClick() ' Anulacion de Comprobantes
   Hide
   Call ANUFAC
End Sub

Private Sub Option6_DblClick()    ' Factura de Acopio
   Hide
   FORFACDIP.Hide
   FORFACABO.Show
End Sub

Private Sub Option7_DblClick()
   Call CONECRUN("REMSTAN", "Despacho de Materiales con Remito")
End Sub
