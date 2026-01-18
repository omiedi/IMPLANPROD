VERSION 5.00
Begin VB.Form INPRIFOR09 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresor de Formularios"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Height          =   645
      Left            =   525
      TabIndex        =   0
      Top             =   945
      Width           =   3690
   End
End
Attribute VB_Name = "INPRIFOR09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   For I% = 1 To CANOBJ%
      Call CARGATEXTO(I%)
      Call PRINOBJ(I%)
   Next I%
   Printer.EndDoc
   '
End Sub
'Public FORPA$(512), FOREPLA$(512), TASEL$(512), CASEL$(512), CONDSEL$(512)

Function TEXREPLA$(TEXORI$, Optional ENTRECOMI%)
     '
     TREP$ = TEXORI$
10   PPXX1% = InStr(TREP$, "{")
     If PPXX1% > 0 Then      ' BUSCA ENCIERRO ENTRE LLAVES
        PPXX2% = InStr(TREP$, "}"): If PPXX2% <= PPXX1% Then PPXX2% = 1 + Len(TREP$)
        IZQUI$ = Left$(TREP$, PPXX1% - 1)
        DEREX$ = Mid$(TREP$, PPXX2% + 1)
        VAREPLA$ = Mid$(TREP, PPXX1% + 1, PPXX2% - PPXX1% - 1)
          If UCase$(Left(VAREPLA$, 7)) = "COMMPAR" Then VAREPLA$ = "COMPAR" + Mid$(VAREPLA$, 8)
          If UCase$(Left$(VAREPLA$, 6)) = "PARCOM" Then VAREPLA$ = "COMPAR" + Mid$(VAREPLA$, 7)
          If VAREPLA$ = "COMPAR" Then VAREPLA$ = "COMPAR-0"
        For JJ& = 1 To CANOBJ%
           If UCase$(ONAME$(JJ&)) = UCase$(VAREPLA$) Then
              VAREPLA = TEXTO(JJ&)
              GoTo 29
           End If
        Next JJ&
        '
29      COMILLA$ = "": If ENTRECOMI% = 1 Then COMILLA$ = "'"
        TREP$ = IZQUI$ + COMILLA$ + VAREPLA$ + COMILLA$ + DEREX$
        GoTo 10
     End If
     '
     TEXREPLA = TREP$
     '
End Function
'
Sub CARGATEXTO(I%)
   '
   If TIOBJ$(I%) = "PARAM" Then
      FOREP$ = LTrim$(RTrim$(FOREPLA$(I%)))
      If FOREP$ <> "" Then
         ' TRABAJA CON TEXTO REEMPLAZADO
         FOREP$ = TEXREPLA$(FOREP$)
         TEXTO(I%) = FOREP$     ' ASIGNA COMO TEXTO EL PARAMETRO
         Exit Sub
      End If
      '
      If TASEL$(I%) <> "" Then
         ' TRABAJA CON VALOBADA
         CONDIX$ = TEXREPLA$(CONDSEL$(I%), 1)
         VAREPLA = VALOBADA(TASEL$(I%), CASEL$(I%), CONDIX$)
         If VarType(VAREPLA) = 7 Then VAREPLA = CVINT(VAREPLA)         ' CASO FECHA
         TEXTO(I%) = VAREPLA
      End If
   End If
   '
End Sub
'
Sub PRINOBJ(I%)
   '
   MODOX$ = TIOBJ$(I%)
   On Error Resume Next
   Select Case MODOX$
      Case "LABEL"
         Printer.CurrentY = 56.7 * V1(I%)
         Printer.CurrentX = 56.7 * H1(I%)
         Printer.Font = FUENTE$(I%)
         Printer.FontSize = XVALO(TAFUEN$(I%))
         Printer.FontBold = NEGRI(I%)
         Printer.FontItalic = ITALI(I%)
         Printer.FontUnderline = SUBRA(I%)
         Printer.Print TEXTO(I%)
      Case "PARAM"
         Printer.CurrentY = 56.7 * V1(I%)
         Printer.CurrentX = 56.7 * H1(I%)
         Printer.Font = FUENTE$(I%)
         Printer.FontSize = XVALO(TAFUEN(I%))
         Printer.FontBold = NEGRI(I%)
         Printer.FontItalic = ITALI(I%)
         Printer.FontUnderline = SUBRA(I%)
         Printer.Print FORMATNUM$(TEXTO(I%), FORPA$(I%))
      Case "LINE"
         Printer.DrawStyle = 0
         Printer.DrawWidth = XVALO(TAFUEN$(I%))
         Printer.Line (56.7 * H1(I%), 56.7 * V1(I%))-(56.7 * H2(I%), 56.7 * V2(I%))
      Case "BOX"
         Printer.DrawStyle = 0
         Printer.DrawWidth = XVALO(TAFUEN$(I%))
         Printer.Line (56.7 * H1(I%), 56.7 * V1(I%))-(56.7 * H2(I%), 56.7 * V2(I%)), , B
   End Select
   Call FRESHALL
   On Error GoTo 0
   '

End Sub

Private Sub Form_Load()
    Call CENTRAFORM(Me)
    Show
    Refresh
    DoEvents
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
    End If
    DoEvents
    '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   If APLINVO$ <> "" Then Height = 3570
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   Call ABRECONEXION
   KLIK% = 0
   '
   APLINVO$ = "AJST-STD(AS.71)"
   If APLINVO$ <> "" Then
        FILTX$ = "": CAPARINV% = 0
        PPXX% = InStr(APLINVO$, "(")
        If PPXX% > 0 Then
           PPYY% = InStr(PPXX% + 1, APLINVO$, ")"): If PPYY% < PPXX% Then PPYY% = 1 + Len(APLINVO$)
           FILTX$ = Mid$(APLINVO$, PPXX% + 1, PPYY% - PPXX% - 1)
           APLINVO$ = Left$(APLINVO$, PPXX% - 1)
        End If
        '
        Screen.MousePointer = 11
        CANFI% = 0
        Do While Len(FILTX$) > 0
           PPXX% = InStr(FILTX$, ";"): If PPXX% < 1 Then PPXX% = 1 + Len(FILTX$)
           TSEL$ = TRIM$(Left$(FILTX$, PPXX% - 1))
           FILTX$ = Mid$(FILTX$, PPXX% + 1)
           If TSEL$ <> "" Then
              CANFI% = CANFI% + 1
              ONAME$(CANFI%) = "COMPAR-" & CStr(CANFI% - 1)
              TEXTO(CANFI%) = TSEL$
              '
              TIOBJ$(CANFI%) = "COMPAR"
              H1(CANFI%) = 0
              V1(CANFI%) = 0
              H2(CANFI%) = 0
              V2(CANFI%) = 0
              '
              FUENTE$(CANFI%) = ""
              TAFUEN$(CANFI%) = ""
              FORPA$(CANFI%) = ""
              '
              NEGRI(CANFI%) = 0
              ITALI(CANFI%) = 0
              SUBRA(CANFI%) = 0
              '
              FOREPLA$(CANFI%) = ""
              TASEL$(CANFI%) = ""
              CASEL$(CANFI%) = ""
              CONDSEL$(CANFI%) = ""
           End If
        Loop
        FORNA$ = APLINVO$
        Call CARGADATFOR(FORNA$, CANFI%)
        Screen.MousePointer = 1
        '
     Else
        Call FINAL("")
   End If
   '
End Sub

