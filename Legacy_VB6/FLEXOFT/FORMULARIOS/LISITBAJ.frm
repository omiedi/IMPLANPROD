VERSION 5.00
Begin VB.Form LISITBAJ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Items de Stock Dados de Baja"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Haga un Doble Click para Re-Vivir el Item"
      Top             =   0
      Width           =   7785
   End
End
Attribute VB_Name = "LISITBAJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub List1_DblClick()
   '
   Screen.MousePointer = 11
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call BLOQARCH("MASTER")
   Call BLOQARCH("INVERT")
   Call BLOQARCH("ECOMAST")
   '
   TTXX$ = List1.TEXT
   CODINUE$ = Left$(TTXX$, 15)
   DEX$ = Mid$(TTXX$, 19, 46)
   REGINV& = 1 + XVALO(Mid$(TTXX$, 65, 8))
   CII% = XVALO(Mid$(TTXX$, 73, 8))
   REGMA& = CII% + 1
   LINDEX& = List1.ListIndex
   '
   If PUNCODIG$ = "" Then
      PUNCODIG$ = LOADFILE$(LUDAT$ + "PUNCODIG.DAT")
   End If
   '
   If CII% < 1 Or CII% > NUMAS% Then GoTo 12
   X$ = REGLEIDO$("MASTER", REGMA&)
   If Mid$(X$, 77, 15) <> AJUSTI$(CODINUE$, 15) Then
12    Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
      LISITBAJ.Hide
      GoTo 99
   End If
   Call REPLA(X$, " ", 92, 1)
   Call GRAREG("MASTER", X$, REGMA&)
   Call CIERRARCH("MASTER")
   '
   If REGINV& < 2 Or REGINV& > NUINV% + 1 Then GoTo 22
   X$ = REGLEIDO$("INVERT", REGINV&)
   If Left$(X$, 15) <> AJUSTI$(CODINUE$, 15) Then
22    Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
      LISITBAJ.Hide
      GoTo 99
   End If
   Call REPLA(X$, " ", 16, 1)
   Call GRAREG("INVERT", X$, REGINV&)
   Call CIERRARCH("INVERT")
   '
   For U& = 1 To Len(PUNCODIG$) Step 16
      If Mid$(PUNCODIG$, U&, 15) = CODINUE$ Then
         Mid$(PUNCODIG$, U&, 16) = CODINUE$ + " "
         GoTo 25
      End If
   Next U&
   Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
   LISITBAJ.Hide
   GoTo 99
   '
25 PPCC$ = PUNCODIG$
   NX% = FILEOPEN%(LUDAT$ + "PUNCODIG.DAT", 0, 2048)
   JI% = 0
   While Len(PPCC$) > 0
      BUFERCOD$ = Left$(PPCC$ + String$(2048, 0), 2048)
      JI% = JI% + 1
      Put #NX%, JI%, BUFERCOD$
      PPCC$ = Mid$(PPCC$, 2049)
   Wend
   Close #NX%
   '
   EL1$ = REGBLAN$("ECOMAST")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, DEX$, 17, 46)
   Call REPLA(EL1$, MKI$(CII%), 63, 2)
   '
   URECO& = ULTREG&("ECOMAST")
   While URECO& > 0
     XXX$ = REGLEIDO$("ECOMAST", URECO&)
     If TRIM$(XXX$) <> "" Then GoTo 35
     URECO& = URECO& - 1
   Wend
   '
35 LOTOTE& = 64 * URECO&
   Call CIERRARCH("ECOMAST")
   FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 2048)
   Dim EL As String * 2048
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   '
   For U& = 1 To Len(TTXX$) Step 64
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 60
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
60 JJ& = 0: LOTEGRA& = Len(TTXX$)
   TTXX$ = TTXX$ + String$(2048, 0)
   For U& = 1 To LOTEGRA& Step 2048
      EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   List1.RemoveItem LINDEX
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
99 Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub
