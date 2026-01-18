VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FMOTRETR 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Motivos y Retrabajo"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   21
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Motivo Y Retrabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   5775
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
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
         Left            =   2310
         MaxLength       =   1
         TabIndex        =   11
         Text            =   " "
         Top             =   1335
         Width           =   240
      End
      Begin VB.TextBox Text19 
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
         Left            =   1200
         TabIndex        =   10
         Top             =   1935
         Width           =   480
      End
      Begin VB.CommandButton Command14 
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
         Height          =   285
         Left            =   1650
         TabIndex        =   9
         Top             =   1935
         Width           =   175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   20
         Top             =   255
         Width           =   1200
      End
      Begin VB.Label LABEL10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Top             =   240
         Width           =   1710
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1050
         TabIndex        =   17
         Top             =   660
         Width           =   4620
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   -360
         TabIndex        =   16
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Retrabajo (s/n)"
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
         Left            =   630
         TabIndex        =   15
         Top             =   1335
         Width           =   1665
      End
      Begin VB.Label Label55 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1890
         TabIndex        =   14
         ToolTipText     =   "Haga Doble-Click para Informe por Célula"
         Top             =   1935
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   3000
         TabIndex        =   13
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3990
         TabIndex        =   12
         Top             =   1335
         Width           =   1665
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   -1800
         TabIndex        =   8
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de O/F: "
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
         Left            =   2400
         TabIndex        =   7
         Top             =   1395
         Width           =   1335
      End
   End
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
      Height          =   615
      Left            =   6120
      Picture         =   "FMOTRETR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "O.K."
      Top             =   2400
      Width           =   650
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   2340
      Left            =   6000
      ScaleHeight     =   2280
      ScaleWidth      =   1005
      TabIndex        =   0
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FMOTRETR.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   845
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "FMOTRETR.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   1560
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "FMOTRETR.frx":075E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1800
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "FMOTRETR.frx":2680
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FMOTRETR.frx":28B4
      TabIndex        =   22
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "FMOTRETR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Command1.Enabled = False
    RES% = COMALTER%("Realmente Desea Salir Sin Guardar?\\No,Continuar \Si,Salir")
    If RES% < 2 Then GoTo 99
    Label2 = ""
    
    Hide
99  Command1.Enabled = True
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    Call SELECHO("TAMOTIV")
    If XVALO(VALACT1("TAMOTIV")) < 1 Then GoTo 99
    Text19 = XVALO(VALACT1("TAMOTIV"))
99  Command14.Enabled = True

End Sub

Private Sub Command2_Click()
     Command2.Enabled = False
     RETRAB$ = TRIM$(Text6)
     If UCase$(RETRAB$) <> "N" And UCase$(RETRAB$) <> "S" Then
        Call MENSERR(24, "Solo Valido 'S'(Re-Trabajo) o 'N'(No Re-Trabajo)")
        GoTo 99
     End If
     '
     NUMMOTIV% = XVALO(Text19)
     If UCase$(RETRAB$) = "S" And (NUMMOTIV% < 1 Or ECOARCH$("TAMOTIV", Chr$(NUMMOTIV%)) = "") Then
       Call MENSERR(24, "Ingreso de Motivo de Rechazo No Válido")
       GoTo 99
     End If
     '
     
     Stat% = 0
     TX$ = FILTERGETRECORD$("TAMOTIV", 1, Chr$(NUMMOTIV%))
     If Asc(Mid$(TX$, 32, 1)) > 0 Then ' SI ESTA MARCADO ESTE MOTIVO EN LA COLUMNA DE SCRAP CON 1
        Stat% = 8
     End If
     '
     If UCase$(RETRAB$) = "N" And Stat% = 8 Then
        Call MENSERR(24, "Ingreso de Motivo de Scrap\Solo si es Retrabajo")
        GoTo 99
     End If
     '
     If Stat% = 8 Then
     TTXX$ = "Ha Ingresado un Informe Con Motivo Scrap\Esto 'Elimina' Todos los Informes Ingresados\Hasta el Momento,para la Presente O/F.\¿Esta Seguro de  Continuar?\\No, Cancelar\Sí, Continuar"
        If COMALTER%(TTXX$) < 2 Then
          GoTo 99
        End If
        If DERACCE%("GENESCRAP", "Informe de Produccion con Motivo Scrap") < 2 Then
           Label2 = ""
           Call Command1_Click
        End If

     End If
     Label2 = "OK"
     Hide
99   Command2.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub Text19_Change()
    On Error Resume Next
    TXT$ = TRIM$(ECOARCH$("TAMOTIV", Chr$(Text19)))
    On Error GoTo 0
    If TXT$ <> "" Then
       Label55 = TXT$
    Else
       Label55 = ""
    End If

End Sub
