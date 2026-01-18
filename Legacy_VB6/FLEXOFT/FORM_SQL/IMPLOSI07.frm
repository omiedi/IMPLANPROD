VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form IMPLOSI07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   495
   ClientTop       =   3105
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2520
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   0
      TabIndex        =   1
      Top             =   450
      Width           =   8100
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   8070
      TabIndex        =   0
      Top             =   0
      Width           =   8100
      Begin VB.Label Label76 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   30
         TabIndex        =   7
         Top             =   100
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1950
         TabIndex        =   6
         Top             =   100
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "U-M"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   5
         Top             =   100
         Width           =   495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Uso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7350
         TabIndex        =   4
         Top             =   100
         Width           =   495
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "IMPLOSI07.frx":0000
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "IMPLOSI07.frx":0234
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "IMPLOSI07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
End Sub

Private Sub List1_DblClick()
   '
   APLICA$ = TRIM$(UCase$(App.EXEName))
   If Left$(APLICA$, 8) = "ANACOS07" Then
       If IUCOD% < 128 Then
         If TRIM$(ANACOSTO07.Text1) <> "" Then
           If ANACOSTO07.Text1 <> UCOD$(IUCOD%) Then
             IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = ANACOSTO07.Text1.TEXT
                                  ICOMPO%(IUCOD%) = ANACOSTO07.List1.ListIndex
           End If
         End If
       End If
       ANACOSTO07.Text1 = Left$(List1.TEXT, 16)
     ElseIf Left$(APLICA$, 6) = "ANACOS" Then
       If IUCOD% < 128 Then
         If TRIM$(ANACOSTO.Text1) <> "" Then
           If ANACOSTO.Text1 <> UCOD$(IUCOD%) Then
             IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = ANACOSTO.Text1.TEXT
                                  ICOMPO%(IUCOD%) = ANACOSTO.List1.ListIndex
           End If
         End If
       End If
       ANACOSTO.Text1 = Left$(List1.TEXT, 16)
     ElseIf Left$(APLICA$, 8) = "AMASTO07" Then
       If IUCOD% < 128 Then
         If TRIM$(AMASTOK07.Text1) <> "" Then
           If AMASTOK07.Text1 <> UCOD$(IUCOD%) Then
             IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = AMASTOK07.Text1.TEXT
                                  ICOMPO%(IUCOD%) = AMASTOK07.List1.ListIndex
           End If
         End If
       End If
       AMASTOK07.Text1 = Left$(List1.TEXT, 16)
     ElseIf Left$(APLICA$, 8) = "AMASTO07" Then
       If IUCOD% < 128 Then
         If TRIM$(AMASTOK07.Text1) <> "" Then
           If AMASTOK07.Text1 <> UCOD$(IUCOD%) Then
             IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = AMASTOK07.Text1.TEXT
                                  ICOMPO%(IUCOD%) = AMASTOK07.List1.ListIndex
           End If
         End If
       End If
       AMASTOK07.Text1 = Left$(List1.TEXT, 16)
     ElseIf Left$(APLICA$, 7) = "AMASTO0" Then
       If IUCOD% < 128 Then
         If TRIM$(AMASTOCK.Text1) <> "" Then
           If AMASTOCK.Text1 <> UCOD$(IUCOD%) Then
             IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = AMASTOCK.Text1.TEXT
                                  ICOMPO%(IUCOD%) = AMASTOCK.List1.ListIndex
           End If
         End If
       End If
       AMASTOCK.Text1 = Left$(List1.TEXT, 16)
     ElseIf Left$(APLICA$, 7) = "FLEXO00" Then
       If IUCOD% < 128 Then
         If TRIM$(AMASTO00.Text1) <> "" Then
           If AMASTO00.Text1 <> UCOD$(IUCOD%) Then
             IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = AMASTO00.Text1.TEXT
                                  ICOMPO%(IUCOD%) = AMASTO00.List1.ListIndex
           End If
         End If
       End If
       AMASTO00.Text1 = Left$(List1.TEXT, 16)
   End If
   Hide
End Sub
