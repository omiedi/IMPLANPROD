VERSION 5.00
Begin VB.Form DIMPOSTO 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importación de Stocks"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   5460
      Picture         =   "DIMPOSTO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1575
      Width           =   750
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4410
      TabIndex        =   8
      Top             =   1890
      Width           =   750
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1995
      TabIndex        =   6
      Top             =   1890
      Width           =   750
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4410
      TabIndex        =   4
      Top             =   1155
      Width           =   750
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1995
      TabIndex        =   2
      Top             =   1155
      Width           =   750
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1995
      TabIndex        =   0
      Top             =   315
      Width           =   4215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1995
      TabIndex        =   9
      Top             =   1995
      Width           =   2325
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock - Desde:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -420
      TabIndex        =   7
      Top             =   1995
      Width           =   2325
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1995
      TabIndex        =   5
      Top             =   1260
      Width           =   2325
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo - Desde:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -420
      TabIndex        =   3
      Top             =   1260
      Width           =   2325
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta / Archivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -420
      TabIndex        =   1
      Top             =   420
      Width           =   2325
   End
End
Attribute VB_Name = "DIMPOSTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
   '
   ARIMPO$ = TRIM$(Text1.TEXT)
   DESCOD% = Val(Text2.TEXT)
   HASCOD% = Val(Text3.TEXT)
   DESTOK% = Val(Text4.TEXT)
   HASTOK% = Val(Text5.TEXT)
   '
   If ARIMPO$ = "" Then Exit Sub
   If EXISTE%(ARIMPO$) < 1 Then
      Call MENSERR(24, "No Existe " + ARIMPO$)
      Exit Sub
   End If
   If DESCOD% < 1 Then Exit Sub
   If HASCOD% < DESCOD% Then Exit Sub
   If DESTOK% < 1 Then Exit Sub
   If HASTOK% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   J& = 0: KK% = 0
   Dim SALSTO(16384)
   For U& = 1 To ULTREG&("MOVISTO")
      X$ = REGLEIDO$("MOVISTO", U&)
      CI% = CVI(Mid$(X$, 3, 2))
      CAN = CVD(Mid$(X$, 88, 8)) - CVD(Mid$(X$, 96, 8))
      If CI% > 0 Then
         If CI% <= 16384 Then
            SALSTO(CI%) = SALSTO(CI%) + CAN
         End If
      End If
19 Next U&
   '
   HOII% = HOY(HO$)
   NRO& = 0
   CMOV$ = TRIM$(Control$("", "CODAJUST"))
   If CMOV$ = "" Then CMOV$ = "AJ"
   N1% = FILEOPEN%(ARIMPO$, 1, 0)
   While Not EOF(N1%)
      Line Input #N1%, AA$
      COD$ = TRIM$(Mid$(AA$, DESCOD%, HASCOD% - DESCOD% + 1))
      CI% = CODINT%(COD$)
      If CI% > 0 Then
         CANTI = Val(Mid$(AA$, DESTOK%, HASTOK% - DESTOK% + 1))
         If Abs(CANTI) > 0.05 Then
            If CI% <= 16384 Then
               SALSTO(CI%) = SALSTO(CI%) + CANTI
            End If
            NRO& = NRO& + 1
            DOSEC$ = "AJ." + Mid$(HO$, 7, 2) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
            DOPRI$ = "AS." + TRIM$(Str$(NRO&))
            VUNI# = 0
            MONE$ = ""
            NC% = 0
            REFE$ = "ST.IMPORT.AL " + HO$
            LTX$ = "ST.INICIAL"
            DEPO% = 1
            FF% = HOII%
            '
            Call MOVISTO(FF%, CI%, LTX$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI)
            '
         End If
      End If
   Wend
   '
   For CI% = 1 To NUMAS%
      U& = CI% + 1
      REBLA$ = String$(36, 0) + MKS$(0.00001)
      Call REPLA(REBLA$, MKS$(SALSTO(CI%)), 1, 4)
      Call GRAREG("SALSTOCK", REBLA$, U&)
   Next CI%
   '
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   Unload DIMPOSTO
   '
End Sub
