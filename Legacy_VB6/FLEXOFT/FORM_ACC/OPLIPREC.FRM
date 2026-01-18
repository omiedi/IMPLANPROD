VERSION 5.00
Begin VB.Form OPLIPREC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listas de Precios - Complementos"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4020
   Icon            =   "OPLIPREC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   1155
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   6
      Top             =   1365
      Width           =   2325
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
         Height          =   600
         Left            =   0
         Picture         =   "OPLIPREC.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Unir Listas de Precios "
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
         Left            =   840
         TabIndex        =   8
         Top             =   105
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   840
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   0
      Top             =   840
      Width           =   2325
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
         Height          =   600
         Left            =   -20
         Picture         =   "OPLIPREC.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Datos Adicionales de Facturacion"
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
         Left            =   630
         TabIndex        =   2
         Top             =   50
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   525
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   3
      Top             =   315
      Width           =   2325
      Begin VB.CommandButton Command6 
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
         Left            =   0
         Picture         =   "OPLIPREC.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenes de Com- pra Terminales"
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
         Left            =   735
         TabIndex        =   5
         Top             =   50
         Width           =   1590
      End
   End
End
Attribute VB_Name = "OPLIPREC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Command1.Enabled = False
   If DERACCE%("UNILIPRE", "Re-Unir Listas de Precios") = 2 Then
     Call UNIPRECIO
   End If
   Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call FINAL("*DATADI")
    Command2.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call FINAL("*OCOTERM")
    Command6.Enabled = True
End Sub

Sub UNIPRECIO()
   '
   Dim LISTAX$(3), FINLI&(3)
   '
   VDEF$ = String$(3, 0)
10 OBX$ = OBJPLA$("UNILISPRE", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   VDEF$ = OBX$
   For kki% = 1 To 3
     ILI% = Asc(Mid$(OBX$, kki%, 1))
     If ILI% < 1 Then
       Call MENSERR(24, "Lista Inexistente")
       GoTo 10
     End If
     If ECOARCH$("LISTAS", Chr$(ILI%)) = "" Then
       Call MENSERR(24, "Lista Inexistente")
       GoTo 10
     End If
   Next kki%
   '
   Screen.MousePointer = 11
   'KKI% identifica el contenido del ingreso
   'de las listas (es una matriz de 3)
   For kki% = 1 To 3
     ILI% = Asc(Mid$(OBX$, kki%, 1))
     LISTAX$(kki%) = TRIM$(Str$(ILI%))
     'ARMA EL NUMERO DE LISTA DE TRES CIFRAS
     While Len(LISTAX$(kki%)) < 3: LISTAX$(kki%) = "0" + LISTAX$(kki%): Wend
     LISTAX$(kki%) = "LIPRE" + LISTAX$(kki%)
     'COPIA LA ESTRUCTURA DE LIPREVEN A LAS LISTAS SELECCIONADAS
     Call COPYSTRU("LIPREVEN", LISTAX$(kki%))
     FINLI&(kki%) = ULTREG&(LISTAX$(kki%))
   Next kki%
   '
   If ULTREG&(LISTAX$(3)) > 0 Then
     Screen.MousePointer = 1
     If COMALTER%("Lista Destino con Información !!!\\Cancelar\Sobre-Escribir") <> 2 Then
       Exit Sub
     End If
   End If
   '
   JJ& = 0
   For kki% = 1 To 2
     FIN& = FINLI&(kki%)
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$(LISTAX$(kki%), U&)
       JJ& = JJ& + 1
       Call GRAREG(LISTAX$(3), X$, JJ&)
     Next U&
     Call CIERRARCH(LISTAX$(kki%))
   Next kki%
   Call SETULTREG(LISTAX$(3), JJ&)
   Call CIERRARCH(LISTAX$(3))
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
   Call COMUNI("Fusión de Listas Completa.\Puede Hacer un Ordenamiento\de la Nueva Lista de Precios.")
   '
End Sub

