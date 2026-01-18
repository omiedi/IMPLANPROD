VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form TRACEFORM 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   170
      Left            =   525
      TabIndex        =   0
      Top             =   420
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   291
      _Version        =   327680
      Appearance      =   1
   End
End
Attribute VB_Name = "TRACEFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
