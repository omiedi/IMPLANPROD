VERSION 5.00
Begin VB.Form IMPLOSI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   495
   ClientTop       =   3105
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   2
      Top             =   450
      Width           =   8100
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C0C0&
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   8025
      TabIndex        =   0
      Top             =   0
      Width           =   8085
      Begin VB.Label Label27 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " Codigo                   Descripcion                                             U-M                Uso     "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   105
         Width           =   7995
      End
   End
End
Attribute VB_Name = "IMPLOSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub List1_DblClick()
   '
   APLICA$ = TRIM$(UCase$(App.EXEName))
   If Left$(APLICA$, 8) = "ANACOS04" Then
       If IUCOD% < 128 Then
         If TRIM$(ANACOSTO04.Text1) <> "" Then
           If ANACOSTO04.Text1 <> UCOD$(IUCOD%) Then
             IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = ANACOSTO04.Text1.TEXT
                                  ICOMPO%(IUCOD%) = ANACOSTO04.List1.ListIndex
           End If
         End If
       End If
       ANACOSTO04.Text1 = Left$(List1.TEXT, 16)
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
     ElseIf Left$(APLICA$, 8) = "AMASTO04" Then
       If IUCOD% < 128 Then
         If TRIM$(AMASTOK04.Text1) <> "" Then
           If AMASTOK04.Text1 <> UCOD$(IUCOD%) Then
             IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = AMASTOK04.Text1.TEXT
                                  ICOMPO%(IUCOD%) = AMASTOK04.List1.ListIndex
           End If
         End If
       End If
       AMASTOK04.Text1 = Left$(List1.TEXT, 16)
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
