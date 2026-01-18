VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   16170
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   8040
      ScaleHeight     =   16.669
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   10.319
      TabIndex        =   0
      Top             =   3000
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://www.recursosvisualbasic.com.ar/htm/tutoriales/metodos-graficos.htm#line
'HAY VARIOS EJEMPLOS MAS
  Option Explicit

Sub Dibujar_cuadricula( _
    Objeto As Object, _
    CountX As Single, _
    CountY As Single, _
    Optional x_Color As Long = vbBlack, _
    Optional y_Color As Long = vbBlack)
    
    Dim I As Integer
    Dim mx As Long
    Dim my As Long
    
    ' limpiar objeto
    Objeto.Cls
    
    mx = CLng(Objeto.ScaleWidth / CountX)
    my = CLng(Objeto.ScaleHeight / CountY)
    
    ' cambiar color del lápiz
    Objeto.ForeColor = x_Color
    
    ' dibujar las líneas verticales de la cuadricula
    For I = 0 To CountX
        Objeto.Line (I * mx, 0)-(I * mx, Objeto.ScaleHeight)
'        Printer.Line (i * mx, 0)-(i * mx, Objeto.ScaleHeight)
    Next I
    
    Objeto.ForeColor = y_Color
    
    ' dibujar las líneas horizontales de la cuadricula
    For I = 0 To CountY
        Objeto.Line (0, I * my)-(Objeto.ScaleWidth, I * my)
'        Printer.Line (i * mx, 0)-(i * mx, Objeto.ScaleHeight)
    Next I
    
End Sub


Private Sub Form_Load()
    
    With Me
        .BackColor = vbWhite
        .ForeColor = vbBlue
        .FontSize = 12
        
        .AutoRedraw = True
        
    End With
        
   P.Cls
    P.ScaleMode = vbPixels
Dim I As Integer
        For I = 0 To ScaleWidth Step 25

            P.Line (I, 10)-(I, P.ScaleHeight - 10), vbGreen

            P.DrawWidth = P.DrawWidth + 10

        Next
    
    P.DrawWidth = 1
    'Call Dibujar_cuadricula(Me, 25, 25, vbBlue, vbRed)
    Call Dibujar_cuadricula(Me, 20, 20, vbBlue, vbRed)
    'Call Dibujar_cuadricula(Me, 25, 50, vbBlue, vbRed)
    
End Sub


