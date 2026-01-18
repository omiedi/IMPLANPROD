VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FORCONSHK 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición Accesos Directos"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Icono"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   6615
      Begin VB.TextBox Text4 
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
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   5895
      End
      Begin VB.TextBox Text3 
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
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   4695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Examinar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Archivo:"
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
         TabIndex        =   14
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   "Ayuda Emergente:"
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
         TabIndex        =   13
         Top             =   960
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox Text2 
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
         Left            =   360
         TabIndex        =   7
         Top             =   1260
         Width           =   5850
      End
      Begin VB.TextBox Text1 
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
         Left            =   315
         TabIndex        =   6
         Top             =   630
         Width           =   4740
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Examinar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   5
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Linea de Comando:"
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
         TabIndex        =   12
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Aplicación:"
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
         TabIndex        =   11
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         Top             =   600
         Width           =   15
      End
   End
   Begin VB.Label Label7 
      Height          =   330
      Left            =   840
      TabIndex        =   16
      Top             =   3990
      Width           =   435
   End
   Begin VB.Label Label6 
      Height          =   330
      Left            =   105
      TabIndex        =   15
      Top             =   3885
      Visible         =   0   'False
      Width           =   435
   End
End
Attribute VB_Name = "FORCONSHK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
   Label7 = "OK"
   Hide
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   CommonDialog1.CancelError = True
   '
   If Label5.Caption = "Ruta - Carpeta" Then
       CommonDialog1.DialogTitle = "Carpeta Destino Imagen Capturada"
       CommonDialog1.Filter = " (*\.)|*\.|"
     Else
       CommonDialog1.DialogTitle = "Archivo Origen de Datos"
          ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
       CommonDialog1.Filter = " (*.bmp)|*.bmp| (*.ico)|*.ico| (*.jpg)|*.jpg|"
   End If
   CommonDialog1.FilterIndex = 0
   
   On Error Resume Next
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Text3 = UCase$(CommonDialog1.FileName)
   Command3.Enabled = True
   '
End Sub

Private Sub Command4_Click()

   Command4.Enabled = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Aplicacion(*.EXE)|*.EXE|"
   CommonDialog1.FilterIndex = 1
   On Error Resume Next
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Text1 = UCase$(CommonDialog1.FileName)

   On Error Resume Next
  
   Command4.Enabled = True

End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me)

End Sub
