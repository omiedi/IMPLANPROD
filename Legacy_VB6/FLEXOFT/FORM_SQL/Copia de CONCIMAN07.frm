VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form CONCIMAN07 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Situación de Conciliación Automática"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   105
      TabIndex        =   7
      Top             =   6300
      Width           =   10725
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "MOVIMIENTOS DEL BANCO (EN EL SISTEMA) PENDIENTES DE CONCILIACION"
         Top             =   210
         Width           =   10455
      End
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   2430
         Left            =   105
         TabIndex        =   9
         Top             =   525
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   4286
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   105
      TabIndex        =   3
      Top             =   3150
      Width           =   10725
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "LISTA DE MOVIMIENTOS NO  CONCILIADOS AUTOMATICAMENTE"
         Top             =   210
         Width           =   10455
      End
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   2430
         Left            =   105
         TabIndex        =   6
         Top             =   525
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   4286
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   10725
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "LISTA DE MOVIMIENTOS CONCILIADOS AUTOMATICAMENTE"
         Top             =   210
         Width           =   10455
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2430
         Left            =   105
         TabIndex        =   5
         Top             =   525
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   4286
         _Version        =   393216
      End
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   0
      X2              =   11040
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "CONCIMAN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   '
   '
End Sub


Private Sub MSF1_Click()
   If MSF1.Row = 0 Then Exit Sub
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   '
   I& = MSF6.MouseRow
   j& = MSF6.MouseCol
   '
   Call COLOREAR_GRILLASQL(CONCIMAN07.MSF1, &HC0FFFF, I&)
   '
End Sub

Sub COLOREAR_GRILLASQL(MSF As MSFlexGrid, COLOR As Variant, Optional REG&)
   With MSF
        .Row = REG&
        For ICOL% = 1 To .Cols - 1 ' DENTRO DEL RENGLON RECORRE CADA COLUMNA
          .Col = ICOL%
          .CellBackColor = COLOR
        Next ICOL%
   End With
End Sub

