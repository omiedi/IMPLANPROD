VERSION 5.00
Begin VB.Form IMPOPECO 
   Caption         =   "Procesos de Importacion de Datos"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "IMPORTAR"
      Height          =   435
      Left            =   2625
      TabIndex        =   6
      Top             =   1785
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Origen de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   315
      TabIndex        =   0
      Top             =   0
      Width           =   6090
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   4815
         TabIndex        =   1
         Top             =   1110
         Width           =   4875
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "PDTES-COBRO.TXT"
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
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   6975
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1110
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   4
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D:\Mis documentos\"
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
         Left            =   945
         TabIndex        =   3
         Top             =   420
         Width           =   4845
      End
   End
End
Attribute VB_Name = "IMPOPECO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'
Call ABRECABAN
On Error Resume Next
  SaDeuPen.MoveLast
While SaDeuPen.RecordCount > 0 'limpiar la TABLA
  SaDeuPen.MoveLast
  SaDeuPen.Delete
Wend
'
On Error Resume Next
 CajaBanc.MoveLast
While CajaBanc.RecordCount > 0 'LIMPIAR TABLA
 CajaBanc.MoveLast
 CajaBanc.Delete
Wend
'
CLIAC$ = ""
RUTA1$ = Label9 + Label1
NX% = FILEOPEN%(RUTA1$, 1, 128)
Do While Not EOF(NX%)
  Line Input #NX%, A$
  A1$ = Left(A$, 2)
  A2$ = Left(A$, 8)
  A3$ = Mid$(A$, 9, 1)
  '
  If A2$ = "        " And A3$ <> "" Then
    CLIAC$ = TRIM$(Mid$(A$, 9, 8))
  End If
  '
  If Val(A1$) > 0 Then
    If Val(A1$) <= 31 Then
       NPREFIJ$ = TRIM$(Mid$(A$, 17, 4))
       NCOMP$ = TRIM$(Mid$(A$, 25, 8))
       NOTIMOV$ = TRIM$(Mid$(A$, 10, 6))
       '
       Select Case NOTIMOV$
         Case "DEART"
           TMOV$ = "AS"
         Case "FAGUI"
           TMOV$ = "FC"
         Case "CRINT"
           TMOV$ = "AS"
         Case "DEINT"
           TMOV$ = "AS"
         Case "FA"
           TMOV$ = "FC"
         Case "RECIB"
           TMOV$ = "RB"
         Case "NDC"
           TMOV$ = "ND"
         Case "RFB"
           TMOV$ = "RB"
         Case "FAEXP"
           TMOV$ = "FC"
         Case "REGCO"
           TMOV$ = "RB"
         Case "NODEB"
           TMOV$ = "ND"
         Case "CINTD"
           TMOV$ = "AS"
         Case "NC"
           TMOV$ = "NC"
         Case "ND"
           TMOV = "ND"
       End Select
       'CODICOMCOR
       If TMOV = "FC" Then
        NCOMPCOR$ = TMOV$ & "-" & NCOMP$ & "-" & "A"
       Else
        NCOMPCOR$ = TMOV$ & "-" & NCOMP$ & "-" & TMOV$
       End If
       'CODICOMLAR
       'COMPLETO PREFIJO
       If Len(NPREFIJ$) < 4 Then
          NJJ% = Len(NPREFIJ$)
          For I = NJJ% To 3
            NPREFIJ$ = "0" + NPREFIJ$
          Next I
       End If
       'COMPLETO NUMERO
       NUMECOMP1 = NCOMP$
       If Len(NCOMP$) < 8 Then
          NJJ1% = Len(NCOMP$)
          For I = NJJ1% To 7
             NCOMP$ = "0" + NCOMP$
          Next I
       End If
       '
       Select Case TMOV
        Case "FC"
            NCOMPLAR$ = TMOV$ & "-" & NPREFIJ$ & "-" & NCOMP$ & "-" & "A"
        Case "RB"
            NCOMPLAR$ = TMOV$ & "-" & "X" & "-" & NPREFIJ$ & "-" & NCOMP$
        Case Else
            NCOMPLAR$ = TMOV$ & "-" & NPREFIJ$ & "-" & NCOMP$ & "-" & TMOV$
       End Select
       'FECHA EMISION
       FEC1 = TRIM$(Mid$(A$, 1, 8))
       'IMPORTE BRUTO
       IMPOBRU# = CDbl(TRIM$(Mid$(A$, 65, 14)))
       'FECHA VTO
       FEC2 = TRIM$(Mid$(A$, 85, 8))
       'ACA GRABO EL REGISTRO EN SaDeuPen.
       Call ABRECABAN
       With SaDeuPen
        .AddNew
         !CodiEmpr = 0
         !CodiSucu = 0
         NCLIENT% = NUMECLI(CLIAC$)
         !NuClien = NCLIENT%
         !NuProve = 0
         !CodCliPro = CLIAC$
         !RaSoCliPro = RASOCLI(NCLIENT%)
         !CodiCom_Cor = NCOMPCOR$
         !CodiCom_Lar = NCOMPLAR$
         !FechEmisi = FEC1
         !IbruComp = IMPOBRU#
         !IIvaComp = 0
         !OrVencim = 0
         !FeVencim = FEC2
         If IMPOBRU# < 0 Then
           !SALDACRE = IMPOBRU#
           !SALDDEBE = 0
         Else
           !SALDDEBE = IMPOBRU#
           !SALDACRE = 0
         End If
         !STATUS = 0
        .Update
       End With
       'GRABO EN CAJABANC
        CoComCor$ = NCOMPCOR$
        CoComLar$ = NCOMPLAR$
        CodiEmp% = 0
        CodiSuc% = 0
        NumeClie% = NCLIENT%
        NumeProv% = 0
        FechEmi = CVDAT(CVINT(FEC1))
        Referen$ = RASOCLI$(NCLIENT%)
        FeContab = CVDAT(CVINT(FEC1))
        FeAcredi = CVDAT(CVINT(FEC1))
        CuentaCo% = CUECOCLI%(NCLIENT%)
        If TMOV = "FC" Then
            TiMovi$ = "FVEN"
        End If
        If TMOV = "RB" Then
            TiMovi$ = "RCOB"
        End If
        If TMOV = "ND" Then
            TiMovi$ = "VENTA"
        End If
        If TMOV = "NC" Then
            TiMovi$ = "VENTA"
        End If
        If TMOV = "AS" Then
            TiMovi$ = "VENTA"
        End If
        '
        OpMovi$ = TMOV
        '
        If TMOV = "RB" Then
            VaMovi$ = "X"
            Descrip$ = "Cobranza Cta.Corriente" & " " & NCOMPCOR$
        End If
        If TMOV = "FC" Then
            VaMovi$ = "A"
            Descrip$ = "Factura: A (Mat. Codif)" & NPREFIJ$ & "-" & NCOMP$
        End If
        If TMOV = "NC" Then
            VaMovi$ = TMOV
            Descrip$ = "VENTA" & " " & NCOMPCOR$
        End If
        If TMOV = "ND" Then
            VaMovi$ = TMOV
            Descrip$ = "VENTA" & " " & NCOMPCOR$
        End If
        If TMOV = "AS" Then
            VaMovi$ = TMOV
            Descrip$ = "VENTA" & " " & NCOMPCOR$
        End If
        '
        NumeTal = NPREFIJ$
        NumeCom = CDbl(NUMECOMP1)
        StatInit = 0
        '//////***HASTA ACA OK*****///////
        'FALTA ASIGNAR EN ImpoDeb
        'EL IMPORTE CORRESPONDIENTE AL DEBE
        'Y EN OTRO GRACABAN EL IMPORTE CORRESPONDIENTE
        'AL HABER
            'ImpoDeb =
        '
        Observa$ = RASOCLI$(NCLIENT%)
        '
       Call GRACABAN
    End If
  End If
Loop
'
End Sub

