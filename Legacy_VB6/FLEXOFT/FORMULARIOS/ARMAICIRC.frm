VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form ARMAICIRC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Armar Mail Circular"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   330
      Left            =   3675
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7980
      ScaleHeight     =   75
      ScaleWidth      =   675
      TabIndex        =   36
      Top             =   7665
      Width           =   740
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   37
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7980
      Picture         =   "ARMAICIRC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Depurar La Base de Datos de Direcciones Electronicas"
      Top             =   7000
      Width           =   750
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Texto de Mensaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4320
      Left            =   210
      TabIndex        =   29
      Top             =   105
      Width           =   8550
      Begin VB.CommandButton Command7 
         Height          =   525
         Left            =   7665
         Picture         =   "ARMAICIRC.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Guardar"
         Top             =   3650
         Width           =   750
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Archivo Adjunto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         TabIndex        =   31
         Top             =   3570
         Width           =   7470
         Begin VB.CommandButton Command6 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7270
            TabIndex        =   32
            Top             =   315
            Width           =   180
         End
         Begin VB.Label Label16 
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
            Height          =   255
            Left            =   0
            TabIndex        =   33
            Top             =   315
            Width           =   7280
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Text            =   "ARMAICIRC.frx":0614
         Top             =   315
         Width           =   8310
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   435
      Left            =   5565
      TabIndex        =   28
      Top             =   10200
      Width           =   2010
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pausa"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   640
      Left            =   8000
      TabIndex        =   26
      Top             =   4515
      Width           =   750
      Begin VB.TextBox Text2 
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
         Left            =   150
         TabIndex        =   27
         Text            =   "10"
         Top             =   250
         Width           =   435
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   330
      Left            =   210
      TabIndex        =   22
      Top             =   4620
      Width           =   6105
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ag.Telef."
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
         Left            =   2730
         TabIndex        =   25
         Top             =   105
         Width           =   1170
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Seg.Com."
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
         Left            =   1365
         TabIndex        =   24
         Top             =   105
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Li-Dir-Mai"
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
         Left            =   0
         TabIndex        =   23
         Top             =   105
         Width           =   1170
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7980
      Picture         =   "ARMAICIRC.frx":061C
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Limpiar Marcas de Envío"
      Top             =   5800
      Width           =   750
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Envios Anteriores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1400
      Left            =   4095
      TabIndex        =   16
      Top             =   6930
      Width           =   3690
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   -20
         TabIndex        =   17
         Top             =   525
         Width           =   3700
      End
      Begin VB.Label Label13 
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
         Height          =   330
         Left            =   1575
         TabIndex        =   19
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   2730
         TabIndex        =   18
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Suprimidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1400
      Left            =   210
      TabIndex        =   12
      Top             =   6930
      Width           =   3795
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   0
         TabIndex        =   13
         Top             =   525
         Width           =   3795
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   2835
         TabIndex        =   15
         Top             =   210
         Width           =   750
      End
      Begin VB.Label Label8 
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
         Height          =   330
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enviados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1780
      Left            =   4095
      TabIndex        =   6
      Top             =   5040
      Width           =   3690
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   0
         TabIndex        =   7
         Top             =   525
         Width           =   3700
      End
      Begin VB.Label Label7 
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
         Height          =   330
         Left            =   1575
         TabIndex        =   11
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   2730
         TabIndex        =   10
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pendientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1780
      Left            =   210
      TabIndex        =   4
      Top             =   5040
      Width           =   3795
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   -20
         TabIndex        =   5
         Top             =   525
         Width           =   3795
      End
      Begin VB.Label Label5 
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
         Height          =   330
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   2835
         TabIndex        =   8
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.CommandButton Command2 
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
      Height          =   600
      Left            =   7980
      Picture         =   "ARMAICIRC.frx":0766
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Interrumpe Transmisión"
      Top             =   6400
      Width           =   750
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7980
      Picture         =   "ARMAICIRC.frx":0BA8
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Iniciar / Continuar Transmisión"
      Top             =   5200
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Selecc.Dest."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6405
      TabIndex        =   1
      Top             =   4620
      Width           =   1380
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   630
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label Label15 
      Height          =   225
      Left            =   7560
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7380
      Picture         =   "ARMAICIRC.frx":1472
      Top             =   7830
      Width           =   2010
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
End
Attribute VB_Name = "ARMAICIRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
   If Check1.Value = 1 Then
     List1.Clear
   End If
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   Command1.Enabled = True
End Sub

Private Sub Command2_Click()
   Label15.Caption = "STOP"
End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   If COMALTER%("Está Seguro de Blanquear Lista de Envíos ?\\No, Cancelar\Si, Refrescar") = 2 Then
     Screen.MousePointer = 11
     If Option1.Value = True Then
         ' LIMPIA LIDIRMAI
         For REMAI& = 1 To ULTREG&("LIDIRMAI")
           XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
           Call REPLA(XX$, MKI$(0), 511, 2)
           Call GRAREG("LIDIRMAI", XX$, REMAI&)
         Next REMAI&
         Call CIERRARCH("LIDIRMAI")
         Call Option1_Click
         GoTo 80
       ElseIf Option2.Value = True Then
         ' LIMPIA BSEGCOM
         For REMAI& = 1 To ULTREG&("BSEGCOM")
           XX$ = REGLEIDO$("BSEGCOM", REMAI&)
           Call REPLA(XX$, MKI$(0), 991, 2)
           Call GRAREG("BSEGCOM", XX$, REMAI&)
         Next REMAI&
         Call CIERRARCH("BSEGCOM")
         Call Option2_Click
         GoTo 80
     End If
80   Screen.MousePointer = 1
   End If
   '
   Command3.Enabled = True
   Unload Me
   '
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
Call POROUTLOOK
GoTo 99

   HOII% = HOY(HOS$)
   Label15 = ""
   '
   'FILTROX$ = ""
   'If Option2.Value = True Then FILTROX$ = "BSEGCOM"
   'If Option5.Value = True Then FILTROX$ = "FLEXOFT"
   '
   MAPISession1.DownLoadMail = False
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   Dim DIREDEST$(10)
   For U& = 2 + List1.ListIndex To List1.ListCount
     List1.ListIndex = U& - 1
     REMAI& = Val(Mid$(List1.List(U& - 1), 81))
     CADESTI% = 0
     If Option1.Value = True Then
          XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
          APELA$ = "Sr / Sra"
          APENO$ = TRIM$(Mid$(XX$, 65, 40))
          EMPRA$ = TRIM$(Mid$(XX$, 173, 64))
          If InStr(EMPRA$, "!") > 0 Then
            EMPRA$ = REPLACAR$(EMPRA$, "!", " ")
          End If
          CANUM$ = ""
          CPOST$ = ""
          LOCA$ = ""
          DIREL$ = REPLACEN$(TRIM$(Mid$(XX$, 105, 60))) + "/"
          '
       ElseIf Option2.Value = True Then
          XX$ = REGLEIDO$("BSEGCOM", REMAI&)
          APELA$ = "Sr / Sra"
          APENO$ = TRIM$(Mid$(XX$, 449, 32))
          EMPRA$ = TRIM$(Mid$(XX$, 3, 62))
          If InStr(EMPRA$, "!") > 0 Then
            EMPRA$ = REPLACAR$(EMPRA$, "!", " ")
          End If
          CANUM$ = TRIM$(Mid$(XX$, 129, 48))
          CPOST$ = TRIM$(Mid$(XX$, 177, 8))
          LOCA$ = TRIM$(Mid$(XX$, 185, 32))
          If CPOST$ <> "" Then LOCA$ = CPOST$ + " - " + LOCA$
          DIREL$ = REPLACEN$(TRIM$(Mid$(XX$, 65, 64))) + "/"
     End If
     '
     DIREL$ = REPLACAR$(DIREL$, ";", "/")
     DIREL$ = REPLACAR$(DIREL$, "\", "/")
     While InStr(DIREL$, "/") > 0
       PPXX% = InStr(DIREL$, "/")
       If PPXX% > 0 Then
         DIREDE$ = TRIM$(Left$(DIREL$, PPXX% - 1))
         DIREL$ = TRIM$(Mid$(DIREL$, PPXX% + 1))
       End If
       If DIREDE$ <> "" Then
         If InStr(DIREDE$, "@") > 0 Then
           CADESTI% = CADESTI% + 1
           DIREDEST$(CADESTI%) = DIREDE$
         End If
       End If
     Wend
     '
     For KKI% = 1 To CADESTI%
       DIREL$ = DIREDEST$(KKI%)
       If InStr(DIREL$, "@") < 2 Then GoTo 54
       '
       MUESTRA$ = ""
       If TRIM$(APENO$) <> "" Then MUESTRA$ = MUESTRA$ + " " + TRIM$(APENO$)
       If TRIM$(EMPRA$) <> "" Then
         If MUESTRA$ <> "" Then
           MUESTRA$ = MUESTRA$ + " -"
         End If
         MUESTRA$ = MUESTRA$ + " " + EMPRA$
       End If
       MUESTRA$ = REPLACEN$(MUESTRA$)
       MUESTRA$ = REPLACAR$(MUESTRA$, ",", " ")
       '
       HOII% = HOY(HOS$)
       TTYY$ = Space$(40) + "Buenos Aires, " + FETEXTO$(HOII%) + ".-" + Chr$(13) + Chr$(10)
       If EMPRA$ <> "" Then
            TTYY$ = TTYY$ + "Sres." + Chr$(13) + Chr$(10)
            TTYY$ = TTYY$ + EMPRA$ + Chr$(13) + Chr$(10)
            If CANUM$ <> "" Then TTYY$ = TTYY$ + CANUM$ + Chr$(13) + Chr$(10)
            If LOCA$ <> "" Then TTYY$ = TTYY$ + LOCA$ + Chr$(13) + Chr$(10)
            TTYY$ = TTYY$ + " " + CRLF$
            If APENO$ <> "" Then
                TTYY$ = TTYY$ + "At Sr./Sra. " + APENO$ + Chr$(13) + Chr$(10)
              Else
                TTYY$ = TTYY$ + "At Sr.Gerente de Planta" + Chr$(13) + Chr$(10)
            End If
         Else
            TTYY$ = TTYY$ + "Sr./Sra." + Chr$(13) + Chr$(10)
            TTYY$ = TTYY$ + APENO$ + Chr$(13) + Chr$(10)
            If CANUM$ <> "" Then TTYY$ = TTYY$ + CANUM$ + Chr$(13) + Chr$(10)
            If LOCA$ <> "" Then TTYY$ = TTYY$ + LOCA$ + Chr$(13) + Chr$(10)
       End If
       '
       For KKII% = 1 To 2: TTYY$ = TTYY$ + CRLF$: Next KKII%
       '
       TTXX$ = Text1 + CRLF$
       TTXX$ = TTXX$ + "Un cordial saludo," + CRLF$
       '
       TTYY$ = TTYY$ + TTXX$
       ' For KKII% = 1 To 2: TTYY$ = TTYY$ + Chr$(13) + Chr$(10): Next KKII%
       ' TTYY$ = TTYY$ + "                                        Ing. Ernesto P. Bauer" + Chr$(13) + Chr$(10)
Stop   ' ESTO SE PUSO PARA EL MAIL CIRCULAR POR LA VENTA DE LA EMPRESA
       TTYY$ = TTYY$ + "===================================================================" + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + "FLEXOFT Soluciones Informáticas Integrales" + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + "Miembro de la Camara de Empresas de Software y Servicios Informáticos" + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + "Bauness 2170 - 5to. 'D' - 1431 Buenos Aires - Argentina" + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + "Telefax (00-54-11) 4524-1284 / 4541-1515 / 4522-2048 " + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + "e-mail: info@flexoft.com.ar - www.flexoft.com.ar" + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + "===================================================================" + Chr$(13) + Chr$(10)
       ' TTYY$ = TTYY$ + "Si quiere ser retirado de lista, favor enviar con 'Asunto: REMOVER'" + Chr$(13) + Chr$(10)
       ' TTYY$ = TTYY$ + "       Le pedimos disculpas por los inconvenientes causados.       " + Chr$(13) + Chr$(10)
       '
       '
              'On Error Resume Next
              'AppActivate "BANDEJA DE ENTRADA"
              'If Err Then
              '  On Error GoTo 0
              '  MsgBox "ERROR"
              '  Stop
              'End If
              'SendKeys "%ANM" + MUESTRA$ + "{TAB}{TAB}"
              'SendKeys "Att: " + MUESTRA$ + "{TAB}"
              ''GoTo 55
              'SendKeys Text1
              'SendKeys "%N"
              'SendKeys "%" + "{F4}"
              'GoTo 55
              'Stop
50     MAPIMessages1.MsgIndex = -1
       MAPIMessages1.MsgSubject = "Att: " + MUESTRA$ ' + " - Software PCP"   ' " - X-News Setiembre 2003"
       MAPIMessages1.MsgNoteText = TTYY$
       MAPIMessages1.RecipIndex = 0
       MAPIMessages1.RecipType = 1
       MAPIMessages1.RecipDisplayName = MUESTRA$
       MAPIMessages1.RecipAddress = DIREL$
       MAPIMessages1.MsgReceiptRequested = True
       '
       If TRIM$(Label16) <> "" Then
         MAPIMessages1.AttachmentIndex = 0
         MAPIMessages1.AttachmentPathName = Label16
         AADJUN$ = Label16
52       PPXX% = InStr(AADJUN$, "\")
         If PPXX% > 0 Then
           AADJUN$ = Mid$(AADJUN$, PPXX% + 1)
           GoTo 52
         End If
         MAPIMessages1.AttachmentName = AADJUN$
       End If
       '
       On Error Resume Next
       MAPIMessages1.Action = 3   ' vbMessageSend
       If Err Then
         MsgBox "Ocurrio Error " & Err
         Resume 90
       End If
54   Next KKI%
     '
     If Option1.Value = True Then
         XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
         Call REPLA(XX$, MKI$(HOII%), 511, 2)
         Call GRAREG("LIDIRMAI", XX$, REMAI&)
       ElseIf Option2.Value = True Then
         XX$ = REGLEIDO$("BSEGCOM", REMAI&)
         Call REPLA(XX$, MKI$(HOII%), 991, 2)
         Call GRAREG("BSEGCOM", XX$, REMAI&)
     End If
     
     Label4 = TRIM$(Str$(Val(Label4) - 1))
     List2.AddItem List1.TEXT
     List2.ListIndex = List2.ListCount - 1
     Label6.Caption = TRIM$(Str$(List2.ListCount))
     '
55   If Val(Text2) < 99 Then
       If List2.ListCount Mod Val(Text2) = 0 Then
         OPX% = COMALTER%("Opciones de Procesamiento:\\Continuar\Interrumpir")
         If OPX% = 2 Then Label15 = "STOP"
       End If
     End If
     '
90   On Error GoTo 0
     MAPISession1.Action = 2  ' !!!
     DoEvents
     '
     If Label15.Caption = "STOP" Then GoTo 98
     '
   MAPISession1.Action = 1   ' !!!
   MAPIMessages1.SessionID = MAPISession1.SessionID  '!!!
95 Next U&
   '
98 'MAPISession1.Action = 2
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
99 Command4.Enabled = True
   '
End Sub

Sub POROUTLOOK()
   HOII% = HOY(HOS$)
   Label15 = ""
   '
   Dim DIREDEST$(10)
   For U& = 2 + List1.ListIndex To List1.ListCount
     List1.ListIndex = U& - 1
     REMAI& = Val(Mid$(List1.List(U& - 1), 81))
     CADESTI% = 0
     If Option1.Value = True Then
          XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
          APELA$ = "Sr / Sra"
          APENO$ = TRIM$(Mid$(XX$, 65, 40))
          EMPRA$ = TRIM$(Mid$(XX$, 173, 64))
          If InStr(EMPRA$, "!") > 0 Then
            EMPRA$ = REPLACAR$(EMPRA$, "!", " ")
          End If
          CANUM$ = ""
          CPOST$ = ""
          LOCA$ = ""
          DIREL$ = REPLACEN$(TRIM$(Mid$(XX$, 105, 60))) + "/"
          '
       ElseIf Option2.Value = True Then
          XX$ = REGLEIDO$("BSEGCOM", REMAI&)
          APELA$ = "Sr / Sra"
          APENO$ = TRIM$(Mid$(XX$, 449, 32))
          EMPRA$ = TRIM$(Mid$(XX$, 3, 62))
          If InStr(EMPRA$, "!") > 0 Then
            EMPRA$ = REPLACAR$(EMPRA$, "!", " ")
          End If
          CANUM$ = TRIM$(Mid$(XX$, 129, 48))
          CPOST$ = TRIM$(Mid$(XX$, 177, 8))
          LOCA$ = TRIM$(Mid$(XX$, 185, 32))
          If CPOST$ <> "" Then LOCA$ = CPOST$ + " - " + LOCA$
          DIREL$ = REPLACEN$(TRIM$(Mid$(XX$, 65, 64))) + "/"
     End If
     '
     DIREL$ = REPLACAR$(DIREL$, ";", "/")
     DIREL$ = REPLACAR$(DIREL$, "\", "/")
     While InStr(DIREL$, "/") > 0
       PPXX% = InStr(DIREL$, "/")
       If PPXX% > 0 Then
         DIREDE$ = TRIM$(Left$(DIREL$, PPXX% - 1))
         DIREL$ = TRIM$(Mid$(DIREL$, PPXX% + 1))
       End If
       If DIREDE$ <> "" Then
         If InStr(DIREDE$, "@") > 0 Then
           CADESTI% = CADESTI% + 1
           DIREDEST$(CADESTI%) = DIREDE$
         End If
       End If
     Wend
     '
     For KKI% = 1 To CADESTI%
       DIREL$ = DIREDEST$(KKI%)
       If InStr(DIREL$, "@") < 2 Then GoTo 54
       '
       MUESTRA$ = ""
       If TRIM$(APENO$) <> "" Then MUESTRA$ = MUESTRA$ + " " + TRIM$(APENO$)
       If TRIM$(EMPRA$) <> "" Then
         If MUESTRA$ <> "" Then
           MUESTRA$ = MUESTRA$ + " -"
         End If
         MUESTRA$ = MUESTRA$ + " " + EMPRA$
       End If
       MUESTRA$ = REPLACEN$(MUESTRA$)
       MUESTRA$ = REPLACAR$(MUESTRA$, ",", " ")
       '
       HOII% = HOY(HOS$)
       '
       'If abc1% = 0 Then
       abc1% = Shell("C:\Archivos de Programa\Outlook Express\msimn.exe", 1)
       'On Error Resume Next
       AppActivate abc1%
       'If Err Then
         On Error GoTo 0
         'TINI = Timer
         'While Timer < TINI + 1: Wend
         'GoTo 10
        ' MsgBox "ERROR"
         'AA1 = Err.Number
         'Stop
       'End If
       'On Error GoTo 0
       'SendKeys "%A"
       'End If
 DIREL$ = "1@2.COM.AR"
       SendKeys "{ESC}", True
       SendKeys "{ESC}", True
       SendKeys "{ESC}", True
       SendKeys "%MUD\\FLEXSERVER\DISCOD\DOCU_COM\NEWSLETTER\FLEXOFT1.HTM" + Chr$(13), True ' "{TAB}{TAB}"
              'SendKeys "%ANM" + MUESTRA$ + "{TAB}{TAB}"
       SendKeys DIREL$ + "{TAB}{TAB}", True
       SendKeys "Att: " + MUESTRA$ + Chr$(13) + Chr$(13) + Chr$(13), True
              ''GoTo 55
       SendKeys "%AE", True   ' Text1
       'SendKeys "%N"
       SendKeys "%" + "{F4}", True
       'SendKeys "N"
              'GoTo 55
              'Stop
'TINI = Timer
'While Timer < TINI + 3
'DoEvents
'Wend
54   Next KKI%
     '
     If Option1.Value = True Then
         XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
         Call REPLA(XX$, MKI$(HOII%), 511, 2)
         Call GRAREG("LIDIRMAI", XX$, REMAI&)
       ElseIf Option2.Value = True Then
         XX$ = REGLEIDO$("BSEGCOM", REMAI&)
         Call REPLA(XX$, MKI$(HOII%), 991, 2)
         Call GRAREG("BSEGCOM", XX$, REMAI&)
     End If
     
     Label4 = TRIM$(Str$(Val(Label4) - 1))
     List2.AddItem List1.TEXT
     List2.ListIndex = List2.ListCount - 1
     Label6.Caption = TRIM$(Str$(List2.ListCount))
     '
55   If Val(Text2) < 99 Then
       If List2.ListCount Mod Val(Text2) = 0 Then
         OPX% = COMALTER%("Opciones de Procesamiento:\\Continuar\Interrumpir")
         If OPX% = 2 Then Label15 = "STOP"
       End If
     End If
     '
     If Label15.Caption = "STOP" Then GoTo 98
     '
95 Next U&
   '
98 Call CIERRARCH("*.*")

End Sub
'
Private Sub Command5_Click()
   Dim FILTX$(128), CAFIX%(128)
   For U& = 1 To List1.ListCount
     List1.ListIndex = U& - 1
     REMAI& = Val(Mid$(List1.List(U& - 1), 81))
     XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
     FILTROX$ = TRIM$(Mid$(XX$, 237, 16))
       For JI% = 1 To CAFIL%
         If FILTX$(JI%) = FILTROX$ Then
           CAFIX%(JI%) = CAFIX%(JI%) + 1
           GoTo 90
         End If
       Next JI%
       CAFIL% = CAFIL% + 1
       FILTX$(CAFIL%) = FILTROX$
       CAFIX%(CAFIL%) = 1
90 Next U&
   '
   For JI% = 1 To CAFIL%
     AA1 = FILTX$(JI%)
     aa2 = CAFIX%(JI%)
   Next JI%
   '
End Sub

Private Sub Command6_Click()
   SELADJUN.Show 1
   Label16 = SELADJUN.Label4
End Sub

Private Sub Dir1_Change()
   Label11.Caption = Dir1.Path
   File1.Path = Dir1.Path
   Label1 = ""
End Sub

Private Sub Drive1_Change()
   On Error GoTo 20
   Dir1.Path = Drive1.Drive
   GoTo 99
   '
20 Resume 21
21 On Error GoTo 0
   Call MENSERR(24, "Unidad no Disponible.")
   Drive1.Drive = Left$(CurDir$, 2)
   '
99 On Error GoTo 0
End Sub

Private Sub file1_click()
  Label1 = File1.FileName
End Sub

Private Sub Command7_Click()
    Call SAVEFILE(LUDAT$ + "TEXTMAIL.TXT", Text1)
    Call SAVEFILE(LUDAT$ + "DOCADJUN.TXT", TRIM$(Label16))
End Sub

Private Sub Command8_Click()
   Command5.Enabled = False
   '
   Screen.MousePointer = 11
   Dim ABUSQUE$(32)
   ABUSQUE$(1) = "MAIL DELIVERY SERVICE"
   ABUSQUE$(2) = "MAIL DELIVERY SUBSYSTEM"
   ABUSQUE$(3) = "MAIL DELIVERY SYSTEM"
   ABUSQUE$(4) = "MAILER-DAEMON"
   ABUSQUE$(5) = "SYSTEM ADMINIS"
   ABUSQUE$(6) = "MAIL ADMINIS"
   ABUSQUE$(7) = "POSTMASTER"
   ABUSQUE$(8) = "THE POST OFFICE"
   ABUSQUE$(9) = "AVG FOR EMAIL"
   CABUS% = 9
   'ABUS$ = TRIM$(UCase$(InputBox$("Buscar Remitentes", , "MAIL")))
   'If ABUS$ = "" Then GoTo 99
   '
   Screen.MousePointer = 11
   MAPISession1.DownLoadMail = False
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   MAPIMessages1.Fetch
   CAMENSA& = MAPIMessages1.MsgCount
   '
   For U& = 1 To CAMENSA&
10   If U& > MAPIMessages1.MsgCount Then
        Call COMUNI("Fin del Control")
        GoTo 99
     End If
     '
     Screen.MousePointer = 11
     Call TRACE(20, U&, CAMENSA&)
     MAPIMessages1.MsgIndex = U& - 1
     ORENVI$ = TRIM$(MAPIMessages1.MsgOrigDisplayName)
     MODOBO% = 0
     '
     For KKI% = 1 To CABUS%
       If Left$(UCase$(ORENVI$), Len(ABUSQUE$(KKI%))) = ABUSQUE$(KKI%) Then
         ATENTI% = 0
         MODOBO% = 0
         FORLIMPIA.Text1 = ORENVI$
         'FORLIMPIA.Text2 = Left$(MAPIMessages1.MsgNoteText, 1024)
         FORLIMPIA.Text2 = MAPIMessages1.MsgNoteText
         '
         ' busca por RFC:
         PINI% = 1
15       PPXX% = InStr(PINI%, UCase$(FORLIMPIA.Text2), "RFC822;")
         If PPXX% > 0 Then
           '
           FORLIMPIA.Text2.SelStart = PPXX% + 6
           PPYY% = InStr(PPXX%, FORLIMPIA.Text2, Chr$(13) + Chr$(10))
           If PPYY% <= PPXX% Then GoTo 18
           If PPYY% > PPXX% Then
             FORLIMPIA.Text2.SelLength = PPYY% - PPXX% - 7
           End If
           TEBUX$ = Mid$(FORLIMPIA.Text2, PPXX% + 7, FORLIMPIA.Text2.SelLength)
           If InStr(UCase$(TEBUX$), "FLEXOFT") > 0 Then
             PINI% = PPYY%
             GoTo 15
           End If
           If InStr(TEBUX$, "@") < 1 Then
             PINI% = PPYY%
             GoTo 15
           End If
           FORLIMPIA.Text3 = TEBUX$
           'MODOBO% = 1
           GoTo 50
         End If
         '
18       PINI% = 1
20       PPXX% = InStr(PINI%, FORLIMPIA.Text2, "<")
         If PPXX% > 0 Then
           FORLIMPIA.Text2.SelStart = PPXX%
           PPYY% = InStr(PPXX%, FORLIMPIA.Text2, ">")
           If PPYY% <= PPXX% Then GoTo 30
           If PPYY% > PPXX% Then
             FORLIMPIA.Text2.SelLength = PPYY% - PPXX% - 1
           End If
           TEBUX$ = Mid$(FORLIMPIA.Text2, PPXX% + 1, FORLIMPIA.Text2.SelLength)
           If InStr(UCase$(TEBUX$), "FLEXOFT") > 0 Then
             PINI% = PPYY%
             GoTo 20
           End If
           If InStr(TEBUX$, "@") < 1 Then
             PINI% = PPYY%
             GoTo 20
           End If
           FORLIMPIA.Text3 = TEBUX$
           'MODOBO% = 1
           GoTo 50
         End If
         '
30       PPYY% = InStr(FORLIMPIA.Text2, "; Failed")
         If PPYY% > 0 Then
           PPYY% = PPYY% - 1
           For PPXX% = PPYY% To 1 Step -1
             If Asc(Mid$(FORLIMPIA.Text2, PPXX%, 1)) < 45 Then
               FORLIMPIA.Text2.SelStart = PPXX%
               FORLIMPIA.Text2.SelLength = PPYY% - PPXX%
               TEBUX$ = Mid$(FORLIMPIA.Text2, PPXX% + 1, PPYY% - PPXX%)
               FORLIMPIA.Text3 = TEBUX$
               'MODOBO% = 1
               GoTo 50
             End If
           Next PPXX%
         End If
         '
50       'If MODOBO% = 1 Then
'           HOII% = HOY(HOS$)
'           If TRIM$(TEBUX$) <> "" Then
'             JJ& = 0
'             For UKK& = 1 To ULTREG&("LIDIRMAI")
'               XXX$ = REGLEIDO$("LIDIRMAI", UKK&)
'               XX1$ = TRIM$(Mid$(XXX$, 105, 60))
'               If XX1$ = TRIM$(TEBUX$) Then
'                 Call REPLA(XXX$, MKI$(HOII%), 255, 2)
'                 Call REPLA(XXX$, "MAL MAIL", 257, 8)
'                 Call GRAREG("LIDIRMAI", XXX$, UKK&)
'               End If
'             Next UKK&
'             ATENTI% = 1
'           End If
'           GoTo 60
'         End If
'         '
         FORLIMPIA.Check1.Value = 1
         FORLIMPIA.Check2.Value = 1
         FORLIMPIA.Text4 = "MAL MAIL"
         FORLIMPIA.Label7 = U&
         Screen.MousePointer = 1
         '
         If Option1.Value = True Then
           If InStr(FORLIMPIA.Text3, "@") > 0 Then
             Call FORLIMPIA.Command1_Click
             GoTo 60
           End If
         End If
         '
         FORLIMPIA.Show 1
         '
60       If ATENTI% = 1 Then
            MAPIMessages1.Delete
            If U& <= MAPIMessages1.MsgCount Then
              GoTo 10
            End If
          ElseIf ATENTI% = 2 Then
            GoTo 99
         End If
         GoTo 89
         '
       End If
     Next KKI%
     '
89 Next U&
   '
99 MAPISession1.Action = 2  ' !!!
   Screen.MousePointer = 1
   Command5.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   HOII% = HOY(HOS$)
   Label15 = ""
   '
   'FILTROX$ = ""
   'If Option2.Value = True Then FILTROX$ = "BSEGCOM"
   'If Option5.Value = True Then FILTROX$ = "FLEXOFT"
   '
   'MAPISession1.DownLoadMail = False
   'MAPISession1.LogonUI = False
   'MAPISession1.Action = 1
   'MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   Dim DIREDEST$(10)
   For U& = 2 + List1.ListIndex To List1.ListCount
     List1.ListIndex = U& - 1
     REMAI& = Val(Mid$(List1.List(U& - 1), 81))
     CADESTI% = 0
     If Option1.Value = True Then
          XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
          APELA$ = "Sr / Sra"
          APENO$ = TRIM$(Mid$(XX$, 65, 40))
          EMPRA$ = TRIM$(Mid$(XX$, 173, 64))
          If InStr(EMPRA$, "!") > 0 Then
            EMPRA$ = REPLACAR$(EMPRA$, "!", " ")
          End If
          CANUM$ = ""
          CPOST$ = ""
          LOCA$ = ""
          DIREL$ = REPLACEN$(TRIM$(Mid$(XX$, 105, 60))) + "/"
          '
       ElseIf Option2.Value = True Then
          XX$ = REGLEIDO$("BSEGCOM", REMAI&)
          APELA$ = "Sr / Sra"
          APENO$ = TRIM$(Mid$(XX$, 449, 32))
          EMPRA$ = TRIM$(Mid$(XX$, 3, 62))
          If InStr(EMPRA$, "!") > 0 Then
            EMPRA$ = REPLACAR$(EMPRA$, "!", " ")
          End If
          CANUM$ = TRIM$(Mid$(XX$, 129, 48))
          CPOST$ = TRIM$(Mid$(XX$, 177, 8))
          LOCA$ = TRIM$(Mid$(XX$, 185, 32))
          If CPOST$ <> "" Then LOCA$ = CPOST$ + " - " + LOCA$
          DIREL$ = REPLACEN$(TRIM$(Mid$(XX$, 65, 64))) + "/"
     End If
     '
     DIREL$ = REPLACAR$(DIREL$, ";", "/")
     DIREL$ = REPLACAR$(DIREL$, "\", "/")
     While InStr(DIREL$, "/") > 0
       PPXX% = InStr(DIREL$, "/")
       If PPXX% > 0 Then
         DIREDE$ = TRIM$(Left$(DIREL$, PPXX% - 1))
         DIREL$ = TRIM$(Mid$(DIREL$, PPXX% + 1))
       End If
       If DIREDE$ <> "" Then
         If InStr(DIREDE$, "@") > 0 Then
           CADESTI% = CADESTI% + 1
           DIREDEST$(CADESTI%) = DIREDE$
         End If
       End If
     Wend
     '
     For KKI% = 1 To CADESTI%
       DIREL$ = DIREDEST$(KKI%)
       If InStr(DIREL$, "@") < 2 Then GoTo 54
       '
       MUESTRA$ = ""
       If TRIM$(APENO$) <> "" Then MUESTRA$ = MUESTRA$ + " " + TRIM$(APENO$)
       If TRIM$(EMPRA$) <> "" Then
         If MUESTRA$ <> "" Then
           MUESTRA$ = MUESTRA$ + " -"
         End If
         MUESTRA$ = MUESTRA$ + " " + EMPRA$
       End If
       MUESTRA$ = REPLACEN$(MUESTRA$)
       MUESTRA$ = REPLACAR$(MUESTRA$, ",", " ")
       '
       HOII% = HOY(HOS$)
       TTYY$ = Space$(40) + "Buenos Aires, " + FETEXTO$(HOII%) + ".-" + Chr$(13) + Chr$(10)
       If EMPRA$ <> "" Then
            TTYY$ = TTYY$ + "Sres." + Chr$(13) + Chr$(10)
            TTYY$ = TTYY$ + EMPRA$ + Chr$(13) + Chr$(10)
            If CANUM$ <> "" Then TTYY$ = TTYY$ + CANUM$ + Chr$(13) + Chr$(10)
            If LOCA$ <> "" Then TTYY$ = TTYY$ + LOCA$ + Chr$(13) + Chr$(10)
            TTYY$ = TTYY$ + " " + CRLF$
            If APENO$ <> "" Then
                TTYY$ = TTYY$ + "At Sr./Sra. " + APENO$ + Chr$(13) + Chr$(10)
              Else
                TTYY$ = TTYY$ + "At Sr.Gerente de Planta" + Chr$(13) + Chr$(10)
            End If
         Else
            TTYY$ = TTYY$ + "Sr./Sra." + Chr$(13) + Chr$(10)
            TTYY$ = TTYY$ + APENO$ + Chr$(13) + Chr$(10)
            If CANUM$ <> "" Then TTYY$ = TTYY$ + CANUM$ + Chr$(13) + Chr$(10)
            If LOCA$ <> "" Then TTYY$ = TTYY$ + LOCA$ + Chr$(13) + Chr$(10)
       End If
       '
       For KKII% = 1 To 2: TTYY$ = TTYY$ + CRLF$: Next KKII%
       '
       TTXX$ = Text1 + CRLF$
       TTXX$ = TTXX$ + "Un cordial saludo," + CRLF$
       '
       TTYY$ = TTYY$ + TTXX$
       'For KKII% = 1 To 2: TTYY$ = TTYY$ + Chr$(13) + Chr$(10): Next KKII%
       'TTYY$ = TTYY$ + "                                        Ing. Ernesto P. Bauer" + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + "===================================================================" + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + "FLEXOFT Soluciones Informáticas Integrales" + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + "Miembro de la Camara de Empresas de Software y Servicios Informáticos" + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + "Bauness 2170 - 5to. 'D' - 1431 Buenos Aires - Argentina" + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + "Telefax (00-54-11) 4524-1284 / 4541-1515 / 4522-2048 " + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + "e-mail: info@flexoft.com.ar - www.flexoft.com.ar" + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + "===================================================================" + Chr$(13) + Chr$(10)
       'TTYY$ = TTYY$ + "Si quiere ser retirado de lista, favor enviar con 'Asunto: REMOVER'" + Chr$(13) + Chr$(10)
       'TTYY$ = TTYY$ + "       Le pedimos disculpas por los inconvenientes causados.       " + Chr$(13) + Chr$(10)
       '
       '
       abcd = Shell("C:\Archivos de Programa\Outlook Express\msimn.exe", 1)
       AppActivate abcd
       SendKeys "%mu1"
 Exit Sub
       '+ MUESTRA$ + "{TAB}{TAB}"
       
 Stop
       
              'On Error Resume Next
              'AppActivate "BANDEJA DE ENTRADA"
              'If Err Then
              '  On Error GoTo 0
              '  MsgBox "ERROR"
              '  Stop
              'End If
              'SendKeys "%ANM" + MUESTRA$ + "{TAB}{TAB}"
              'SendKeys "Att: " + MUESTRA$ + "{TAB}"
              ''GoTo 55
              'SendKeys Text1
              'SendKeys "%N"
              'SendKeys "%" + "{F4}"
              'GoTo 55
              'Stop
50     MAPIMessages1.MsgIndex = -1
       MAPIMessages1.MsgSubject = "Att: " + MUESTRA$ ' + " - Software PCP"   ' " - X-News Setiembre 2003"
       MAPIMessages1.MsgNoteText = TTYY$
       MAPIMessages1.RecipIndex = 0
       MAPIMessages1.RecipType = 1
       MAPIMessages1.RecipDisplayName = MUESTRA$
       MAPIMessages1.RecipAddress = DIREL$
       MAPIMessages1.MsgReceiptRequested = True
       '
       If TRIM$(Label16) <> "" Then
         MAPIMessages1.AttachmentIndex = 0
         MAPIMessages1.AttachmentPathName = Label16
         AADJUN$ = Label16
52       PPXX% = InStr(AADJUN$, "\")
         If PPXX% > 0 Then
           AADJUN$ = Mid$(AADJUN$, PPXX% + 1)
           GoTo 52
         End If
         MAPIMessages1.AttachmentName = AADJUN$
       End If
       '
       On Error Resume Next
       MAPIMessages1.Action = 3   ' vbMessageSend
       If Err Then
         MsgBox "Ocurrio Error " & Err
         Resume 90
       End If
54   Next KKI%
     '
     If Option1.Value = True Then
         XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
         Call REPLA(XX$, MKI$(HOII%), 511, 2)
         Call GRAREG("LIDIRMAI", XX$, REMAI&)
       ElseIf Option2.Value = True Then
         XX$ = REGLEIDO$("BSEGCOM", REMAI&)
         Call REPLA(XX$, MKI$(HOII%), 991, 2)
         Call GRAREG("BSEGCOM", XX$, REMAI&)
     End If
     
     Label4 = TRIM$(Str$(Val(Label4) - 1))
     List2.AddItem List1.TEXT
     List2.ListIndex = List2.ListCount - 1
     Label6.Caption = TRIM$(Str$(List2.ListCount))
     '
55   If Val(Text2) < 99 Then
       If List2.ListCount Mod Val(Text2) = 0 Then
         OPX% = COMALTER%("Opciones de Procesamiento:\\Continuar\Interrumpir")
         If OPX% = 2 Then Label15 = "STOP"
       End If
     End If
     '
90   On Error GoTo 0
     MAPISession1.Action = 2  ' !!!
     DoEvents
     '
     If Label15.Caption = "STOP" Then GoTo 98
     '
   MAPISession1.Action = 1   ' !!!
   MAPIMessages1.SessionID = MAPISession1.SessionID  '!!!
95 Next U&
   '
98 'MAPISession1.Action = 2
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
99 Command19.Enabled = True
   '

End Sub

Private Sub Form_Load()
    '
    EPAC$ = EMPREAC$
    If EPAC$ <> "" Then
      Caption = Caption + " - " + EMPREAC$
    End If
    '
    Refresh
    DoEvents
    '
    Text1 = LOADFILE(LUDAT$ + "TEXTMAIL.TXT")
    Label16 = LOADFILE(LUDAT$ + "DOCADJUN.TXT")
    'Label11.Caption = Dir1.Path
    'File1.Path = Dir1.Path
    '
    Screen.MousePointer = 11
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    '
'    For U& = 1 To ULTREG&("LIDIRMAI")
'      XX$ = REGLEIDO$("LIDIRMAI", U&)
'      TTXX$ = AJUSTI$(Mid$(XX$, 5, 60), 80) + Str$(U&)
'      If CVI(Mid$(XX$, 255, 2)) > 0 Or TRIM$(Mid$(XX$, 257, 8)) <> "" Then
'         ' esta marcado de baja
'         List3.AddItem TTXX$
'         List3.ListIndex = List3.ListCount - 1
'         Label9 = TRIM$(Str$(List3.ListCount))
'       ElseIf CVI(Mid$(XX$, 511, 2)) < 0 Then
'         ' esta marcado suprimido
'         List3.AddItem TTXX$
'         List3.ListIndex = List3.ListCount - 1
'         Label9 = TRIM$(Str$(List3.ListCount))
'       ElseIf CVI(Mid$(XX$, 511, 2)) > 0 Then
'         ' ya fue enviado
'         List4.AddItem TTXX$
'         List4.ListIndex = List4.ListCount - 1
'         Label14 = TRIM$(Str$(List4.ListCount))
'       Else
'         ' agregar a lists de envios
'         List1.AddItem TTXX$
'         List1.ListIndex = List1.ListCount - 1
'         Label4 = TRIM$(Str$(List1.ListCount))
'      End If
'      DoEvents
'    Next U&
'    '
    On Error Resume Next
    List1.TopIndex = 0
    List1.ListIndex = -1
    List2.TopIndex = 0
    List2.ListIndex = -1
    List3.TopIndex = 0
    List3.ListIndex = -1
    List4.TopIndex = 0
    List4.ListIndex = -1
    On Error GoTo 0
    '
    Call Option2_Click
    '
    Screen.MousePointer = 1

End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Hide
End Sub


Private Sub Label11_Change()
  If Right$(Label11, 1) = "\" Then
    Label11 = Left$(Label11, Len(Label11) - 1)
  End If
End Sub

Private Sub Label16_DblClick()
   Call Command6_Click
End Sub

Private Sub List1_DblClick()
   '
   U& = List1.ListIndex
   REMAI& = Val(Mid$(List1.List(U&), 81))
   '
   XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
   Call REPLA(XX$, MKI$(-1), 511, 2)
   Call GRAREG("LIDIRMAI", XX$, REMAI&)
   Call CIERRARCH("LIDIRMAI")
   '
   List3.AddItem List1.List(U&)
   List1.RemoveItem U&
   Label9 = TRIM$(Str$(List3.ListCount))
   '
End Sub

Private Sub Option1_Click()
   '
   If Option1.Value = True Then
     FIDAT1 = FileDateTime("F:\LIDIRMAI\LIDIRMAI.DAT")
     FIDAT2 = FileDateTime("F:\SEGUICOM\LIDIRMAI.DAT")
     If FIDAT1 > FIDAT2 Then
       RESPX% = COPYFILE("F:\LIDIRMAI\LIDIRMAI.DAT", "F:\SEGUICOM\LIDIRMAI.DAT")
     End If
     ' carga la lista de envios de BSEGCOM
     Screen.MousePointer = 11
     List1.Clear
     List2.Clear
     List3.Clear
     List4.Clear
     '
     FIN& = ULTREG&("LIDIRMAI")
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("LIDIRMAI", U&)
       TTXX$ = AJUSTI$(Mid$(XX$, 5, 60), 80) + Str$(U&)
       If CVI(Mid$(XX$, 255, 2)) > 0 Or TRIM$(Mid$(XX$, 257, 8)) <> "" Then
          ' esta marcado de baja
          List3.AddItem TTXX$
          List3.ListIndex = List3.ListCount - 1
          Label9 = TRIM$(Str$(List3.ListCount))
        ElseIf CVI(Mid$(XX$, 511, 2)) < 0 Then
          ' esta marcado suprimido
          List3.AddItem TTXX$
          List3.ListIndex = List3.ListCount - 1
          Label9 = TRIM$(Str$(List3.ListCount))
        ElseIf CVI(Mid$(XX$, 511, 2)) > 0 Then
          ' ya fue enviado
AA1 = CVI(Mid$(XX$, 511, 2))
          List4.AddItem TTXX$
          List4.ListIndex = List4.ListCount - 1
          Label14 = TRIM$(Str$(List4.ListCount))
        Else
          ' agregar a lists de envios
          List1.AddItem TTXX$
          List1.ListIndex = List1.ListCount - 1
          Label4 = TRIM$(Str$(List1.ListCount))
       End If
       DoEvents
     Next U&
     '
     On Error Resume Next
     List1.TopIndex = 0
     List1.ListIndex = -1
     List2.TopIndex = 0
     List2.ListIndex = -1
     List3.TopIndex = 0
     List3.ListIndex = -1
     List4.TopIndex = 0
     List4.ListIndex = -1
     On Error GoTo 0
     '
   End If
   Screen.MousePointer = 1

End Sub

Private Sub Option2_Click()
   If Option2.Value = True Then
     ' carga la lista de envios de BSEGCOM
     Screen.MousePointer = 11
     List1.Clear
     List2.Clear
     List3.Clear
     List4.Clear
     '
     FIN& = ULTREG&("BSEGCOM")
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("BSEGCOM", U&)
       TTXX$ = AJUSTI$(Mid$(XX$, 3, 62), 80) + FORMATNUM$(U&, "F6.0")
       FEBAJ% = CVI(Mid$(XX$, 973, 2))
       FULEN% = CVI(Mid$(XX$, 991, 2))
       '
       If FULEN% > 0 Then
           List4.AddItem TTXX$
         ElseIf FEBAJ% > 0 And FEBAJ% <> 8224 Then
           List3.AddItem TTXX$
         Else
           List1.AddItem TTXX$
       End If
     Next U&
     '
     Label4 = List1.ListCount
     Label9 = List3.ListCount
     Label6 = List2.ListCount
     Label14 = List4.ListCount
     Screen.MousePointer = 1
   End If
End Sub
